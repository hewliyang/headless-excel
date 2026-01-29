"""Shared daemon constants and utilities."""

import os
import shutil
import socket
from functools import cache
from pathlib import Path
from sys import platform

from headless_excel.daemon.config import PID_FILE, PORT_FILE, Config, get_daemon_port
from headless_excel.errors import LibreOfficeNotFoundError, RecalcError


def cleanup_daemon_files() -> None:
    """Clean up daemon PID and port files."""
    PID_FILE.unlink(missing_ok=True)
    PORT_FILE.unlink(missing_ok=True)


def _get_config() -> Config:
    """Get daemon config (internal helper)."""
    return Config.from_env()


def get_unobridge_macro(config: Config | None = None) -> str:
    """
    Generate the Python macro that runs inside LibreOffice.

    Creates a TCP server and handles commands for session-based document operations.
    Values for port and idle timeout are injected from current config.

    Session-based commands (for keeping documents in memory):
    - OPEN:filepath -> SESSION:id or ERROR:...
    - CLOSE:session_id -> OK or ERROR:...
    - GET:session_id:Sheet!A1 -> VALUE:type:data or ERROR:...
    - GET_RANGE:session_id:Sheet!A1:B2 -> DATA:json or ERROR:...
    - SET:session_id:Sheet!A1=value -> OK or ERROR:...
    - SET_RANGE:session_id:Sheet!A1:B2:json -> OK or ERROR:...
    - CALC:session_id -> OK (in-memory recalc)
    - SAVE:session_id -> OK or ERROR:...
    - SHEETS:session_id -> SHEETS:json_list
    - CREATE_SHEET:session_id:name[:index] -> OK or ERROR:...
    - DELETE_SHEET:session_id:name -> OK or ERROR:...
    - GET_ACTIVE:session_id -> SHEET:name
    - SET_ACTIVE:session_id:name -> OK or ERROR:...
    - NEW:filepath -> SESSION:id (create new workbook)

    Legacy commands (for backward compatibility):
    - PING -> PONG
    - QUIT -> OK (shuts down daemon)
    - RECALC:filepath -> OK (load, recalc, save, close)
    """
    if config is None:
        config = _get_config()
    return f'''\
"""TCP bridge for headless-excel recalculation daemon with session support."""
import socket
import time
import json
import uuid
import base64
import traceback
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table.CellContentType import VALUE, TEXT, FORMULA, EMPTY

DAEMON_PORT = {config.daemon_port}
IDLE_TIMEOUT = {config.idle_timeout}

# Session storage: session_id -> (doc, filepath)
sessions = {{}}

def recv_all(conn, initial_data=""):
    """Receive all data until newline or connection close."""
    data = initial_data
    while True:
        if "\\n" in data:
            return data.split("\\n")[0]
        try:
            chunk = conn.recv(65536).decode('utf-8')
            if not chunk:
                break
            data += chunk
        except socket.timeout:
            break
    return data.strip()

def get_cell_value(cell):
    """Get cell value with type information."""
    cell_type = cell.getType()
    if cell_type == EMPTY:
        return ("empty", None)
    elif cell_type == VALUE:
        return ("number", cell.getValue())
    elif cell_type == TEXT:
        return ("string", cell.getString())
    elif cell_type == FORMULA:
        # Return computed value, not formula
        # Check if result is numeric or string
        if cell.getValue() != 0 or cell.getString() == "0":
            val = cell.getValue()
            # Check for error
            err = cell.getError()
            if err != 0:
                return ("error", cell.getString())
            return ("number", val)
        s = cell.getString()
        if s:
            return ("string", s)
        return ("number", cell.getValue())
    return ("unknown", str(cell.getValue()))

def get_cell_formula(cell):
    """Get cell formula if it has one."""
    f = cell.getFormula()
    if f and f.startswith("="):
        return f
    return None

def set_cell_value(cell, value):
    """Set cell value based on type."""
    if value is None:
        cell.setString("")
    elif isinstance(value, bool):
        cell.setValue(1 if value else 0)
    elif isinstance(value, (int, float)):
        cell.setValue(value)
    elif isinstance(value, str):
        if value.startswith("="):
            cell.setFormula(value)
        else:
            cell.setString(value)
    else:
        cell.setString(str(value))

def parse_cell_ref(ref):
    """Parse cell reference like 'Sheet1!A1' or 'A1' into (sheet_name, col, row)."""
    if "!" in ref:
        sheet_name, cell_ref = ref.split("!", 1)
    else:
        sheet_name = None
        cell_ref = ref
    
    # Parse column letters and row number
    col_str = ""
    row_str = ""
    for c in cell_ref:
        if c.isalpha():
            col_str += c.upper()
        else:
            row_str += c
    
    # Convert column letters to index (A=0, B=1, etc.)
    col = 0
    for c in col_str:
        col = col * 26 + (ord(c) - ord('A') + 1)
    col -= 1  # 0-indexed
    
    row = int(row_str) - 1  # 0-indexed
    return sheet_name, col, row

def parse_range_ref(ref):
    """Parse range reference like 'Sheet1!A1:B2' into (sheet_name, start_col, start_row, end_col, end_row)."""
    if "!" in ref:
        sheet_name, range_ref = ref.split("!", 1)
    else:
        sheet_name = None
        range_ref = ref
    
    if ":" in range_ref:
        start_ref, end_ref = range_ref.split(":")
    else:
        start_ref = end_ref = range_ref
    
    _, start_col, start_row = parse_cell_ref(start_ref)
    _, end_col, end_row = parse_cell_ref(end_ref)
    
    return sheet_name, start_col, start_row, end_col, end_row

def get_sheet(doc, sheet_name=None):
    """Get sheet by name or active sheet."""
    sheets = doc.getSheets()
    if sheet_name:
        if sheets.hasByName(sheet_name):
            return sheets.getByName(sheet_name)
        return None
    # Return first sheet as default
    return sheets.getByIndex(0)

def handle_command(cmd, desktop):
    """Handle a single command and return response."""
    global sessions
    
    try:
        if cmd == "PING":
            return "PONG"
        
        elif cmd == "QUIT":
            # Close all sessions
            for sid, (doc, _) in list(sessions.items()):
                try:
                    doc.close(False)
                except:
                    pass
            sessions.clear()
            return "QUIT_OK"
        
        elif cmd.startswith("OPEN:"):
            filepath = cmd[5:]
            try:
                url = uno.systemPathToFileUrl(filepath)
                doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
                if doc is None:
                    return "ERROR:Failed to open document"
                session_id = str(uuid.uuid4())[:8]
                sessions[session_id] = (doc, filepath)
                return f"SESSION:{{session_id}}"
            except Exception as e:
                return f"ERROR:{{e}}"
        
        elif cmd.startswith("NEW:"):
            filepath = cmd[4:]
            try:
                doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
                if doc is None:
                    return "ERROR:Failed to create document"
                session_id = str(uuid.uuid4())[:8]
                sessions[session_id] = (doc, filepath)
                return f"SESSION:{{session_id}}"
            except Exception as e:
                return f"ERROR:{{e}}"
        
        elif cmd.startswith("CLOSE:"):
            session_id = cmd[6:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions.pop(session_id)
            try:
                doc.close(False)
            except:
                pass
            return "OK"
        
        elif cmd.startswith("SAVE:"):
            session_id = cmd[5:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, filepath = sessions[session_id]
            try:
                url = uno.systemPathToFileUrl(filepath)
                props = (PropertyValue("FilterName", 0, "Calc MS Excel 2007 XML", 0),)
                doc.storeToURL(url, props)
                return "OK"
            except Exception as e:
                return f"ERROR:{{e}}"
        
        elif cmd.startswith("CALC:"):
            session_id = cmd[5:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            doc.calculateAll()
            return "OK"
        
        elif cmd.startswith("RELOAD:"):
            # Reload document from disk (picks up external changes)
            session_id = cmd[7:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, filepath = sessions[session_id]
            try:
                # Close current document
                doc.close(False)
                # Reopen from disk
                url = uno.systemPathToFileUrl(filepath)
                doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
                if doc is None:
                    del sessions[session_id]
                    return "ERROR:Failed to reload document"
                sessions[session_id] = (doc, filepath)
                return "OK"
            except Exception as e:
                # Clean up on error
                if session_id in sessions:
                    del sessions[session_id]
                return f"ERROR:{{e}}"
        
        elif cmd.startswith("SHEETS:"):
            session_id = cmd[7:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            sheets = doc.getSheets()
            names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            return f"SHEETS:{{json.dumps(names)}}"
        
        elif cmd.startswith("GET_ACTIVE:"):
            session_id = cmd[11:]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            # Get active sheet (first sheet for calc docs)
            controller = doc.getCurrentController()
            if controller:
                sheet = controller.getActiveSheet()
                return f"SHEET:{{sheet.getName()}}"
            return f"SHEET:{{doc.getSheets().getByIndex(0).getName()}}"
        
        elif cmd.startswith("SET_ACTIVE:"):
            parts = cmd[11:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid SET_ACTIVE format"
            session_id, sheet_name = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            sheets = doc.getSheets()
            if not sheets.hasByName(sheet_name):
                return f"ERROR:Sheet not found: {{sheet_name}}"
            controller = doc.getCurrentController()
            if controller:
                controller.setActiveSheet(sheets.getByName(sheet_name))
            return "OK"
        
        elif cmd.startswith("CREATE_SHEET:"):
            parts = cmd[13:].split(":", 2)
            if len(parts) < 2:
                return "ERROR:Invalid CREATE_SHEET format"
            session_id = parts[0]
            sheet_name = parts[1]
            index = int(parts[2]) if len(parts) > 2 else -1
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            sheets = doc.getSheets()
            if index < 0:
                index = sheets.getCount()
            sheets.insertNewByName(sheet_name, index)
            return "OK"
        
        elif cmd.startswith("DELETE_SHEET:"):
            parts = cmd[13:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid DELETE_SHEET format"
            session_id, sheet_name = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            sheets = doc.getSheets()
            if sheets.getCount() <= 1:
                return "ERROR:Cannot delete the only sheet"
            if not sheets.hasByName(sheet_name):
                return f"ERROR:Sheet not found: {{sheet_name}}"
            sheets.removeByName(sheet_name)
            return "OK"
        
        elif cmd.startswith("RENAME_SHEET:"):
            parts = cmd[13:].split(":", 2)
            if len(parts) != 3:
                return "ERROR:Invalid RENAME_SHEET format"
            session_id, old_name, new_name = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            sheets = doc.getSheets()
            if not sheets.hasByName(old_name):
                return f"ERROR:Sheet not found: {{old_name}}"
            sheet = sheets.getByName(old_name)
            sheet.setName(new_name)
            return "OK"
        
        elif cmd.startswith("GET:"):
            # GET:session_id:Sheet!A1 or GET:session_id:A1
            parts = cmd[4:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid GET format"
            session_id, cell_ref = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            
            sheet_name, col, row = parse_cell_ref(cell_ref)
            sheet = get_sheet(doc, sheet_name)
            if sheet is None:
                return f"ERROR:Sheet not found: {{sheet_name}}"
            
            cell = sheet.getCellByPosition(col, row)
            val_type, val = get_cell_value(cell)
            formula = get_cell_formula(cell)
            
            result = {{"type": val_type, "value": val}}
            if formula:
                result["formula"] = formula
            return f"VALUE:{{json.dumps(result)}}"
        
        elif cmd.startswith("GET_RANGE:"):
            # GET_RANGE:session_id:Sheet!A1:B2
            parts = cmd[10:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid GET_RANGE format"
            session_id, range_ref = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            
            sheet_name, start_col, start_row, end_col, end_row = parse_range_ref(range_ref)
            sheet = get_sheet(doc, sheet_name)
            if sheet is None:
                return f"ERROR:Sheet not found: {{sheet_name}}"
            
            # Get data array first (fast for values)
            cell_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
            data = cell_range.getDataArray()
            
            # Convert to list of lists
            result = [list(row) for row in data]
            
            # getDataArray returns None for error cells - replace with error strings
            num_rows = end_row - start_row + 1
            num_cols = end_col - start_col + 1
            for row_idx in range(num_rows):
                for col_idx in range(num_cols):
                    if result[row_idx][col_idx] is None:
                        cell = sheet.getCellByPosition(start_col + col_idx, start_row + row_idx)
                        if cell.getError() != 0:
                            # Cell has an error - get the error string
                            result[row_idx][col_idx] = cell.getString()
            
            return f"DATA:{{json.dumps(result)}}"
        
        elif cmd.startswith("GET_FORMULAS:"):
            # GET_FORMULAS:session_id:Sheet!A1:B2
            parts = cmd[13:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid GET_FORMULAS format"
            session_id, range_ref = parts
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            
            sheet_name, start_col, start_row, end_col, end_row = parse_range_ref(range_ref)
            sheet = get_sheet(doc, sheet_name)
            if sheet is None:
                return f"ERROR:Sheet not found: {{sheet_name}}"
            
            cell_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
            formulas = cell_range.getFormulaArray()
            result = [list(row) for row in formulas]
            return f"FORMULAS:{{json.dumps(result)}}"
        
        elif cmd.startswith("SET:"):
            # SET:session_id:Sheet!A1=value
            parts = cmd[4:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Invalid SET format"
            session_id, rest = parts
            if "=" not in rest:
                return "ERROR:Invalid SET format (missing =)"
            cell_ref, value_str = rest.split("=", 1)
            
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            
            sheet_name, col, row = parse_cell_ref(cell_ref)
            sheet = get_sheet(doc, sheet_name)
            if sheet is None:
                return f"ERROR:Sheet not found: {{sheet_name}}"
            
            cell = sheet.getCellByPosition(col, row)
            
            # Try to parse value as JSON, fallback to string
            try:
                value = json.loads(value_str)
            except:
                value = value_str
            
            set_cell_value(cell, value)
            return "OK"
        
        elif cmd.startswith("SET_RANGE:"):
            # SET_RANGE:session_id:Sheet!A1:B2:[[1,2],[3,4]]
            parts = cmd[10:].split(":", 2)
            if len(parts) != 3:
                return "ERROR:Invalid SET_RANGE format"
            session_id, range_ref, data_json = parts
            
            if session_id not in sessions:
                return "ERROR:Invalid session"
            doc, _ = sessions[session_id]
            
            sheet_name, start_col, start_row, end_col, end_row = parse_range_ref(range_ref)
            sheet = get_sheet(doc, sheet_name)
            if sheet is None:
                return f"ERROR:Sheet not found: {{sheet_name}}"
            
            try:
                data = json.loads(data_json)
            except Exception as e:
                return f"ERROR:Invalid JSON data: {{e}}"
            
            # Determine actual range from data dimensions
            num_rows = len(data)
            num_cols = max(len(row) for row in data) if data else 0
            actual_end_row = start_row + num_rows - 1
            actual_end_col = start_col + num_cols - 1
            
            # Check if data contains formulas (strings starting with =)
            has_formulas = any(
                isinstance(cell, str) and cell.startswith("=")
                for row in data for cell in row
            )
            
            if has_formulas:
                # Set cell by cell for formulas
                for row_idx, row_data in enumerate(data):
                    for col_idx, value in enumerate(row_data):
                        cell = sheet.getCellByPosition(start_col + col_idx, start_row + row_idx)
                        set_cell_value(cell, value)
            else:
                # Use bulk setDataArray for pure values (faster)
                cell_range = sheet.getCellRangeByPosition(start_col, start_row, actual_end_col, actual_end_row)
                # Pad rows to same length
                padded_data = [row + [None] * (num_cols - len(row)) for row in data]
                # Convert to tuple of tuples, replacing None with empty string
                tuple_data = tuple(
                    tuple("" if v is None else v for v in row)
                    for row in padded_data
                )
                cell_range.setDataArray(tuple_data)
            
            return "OK"
        
        elif cmd.startswith("RECALC:"):
            # Legacy: load file, recalc, save, close
            filepath = cmd[7:]
            try:
                url = uno.systemPathToFileUrl(filepath)
                doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
                doc.calculateAll()
                doc.store()
                doc.close(True)
                return "OK"
            except Exception as e:
                return f"ERROR:{{e}}"
        
        elif cmd.startswith("EXEC:"):
            # EXEC:session_id:base64_encoded_code
            parts = cmd.split(":", 2)
            if len(parts) != 3:
                return "ERROR:Invalid EXEC format (expected EXEC:session_id:base64_code)"
            session_id, code_b64 = parts[1], parts[2]
            if session_id not in sessions:
                return "ERROR:Invalid session"
            
            try:
                code = base64.b64decode(code_b64).decode('utf-8')
            except Exception as e:
                return f"ERROR:Failed to decode code: {{e}}"
            
            doc, filepath = sessions[session_id]
            
            # Minimal globals - just the essentials
            # Use same dict for globals/locals to avoid comprehension scope issues
            exec_ns = {{
                'doc': doc,
                'desktop': desktop,
                'uno': uno,
                'json': json,
            }}
            
            try:
                exec(code, exec_ns)
                result = exec_ns.get('result', None)
                return f"RESULT:{{json.dumps(result)}}"
            except Exception as e:
                return f"EXEC_ERROR:{{traceback.format_exc()}}"
        
        else:
            return "ERROR:Unknown command"
    
    except Exception as e:
        return f"ERROR:{{e}}"

def start_server(*args):
    """Start TCP server for recalc commands."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(('127.0.0.1', DAEMON_PORT))
    server.listen(5)
    server.settimeout(60)  # Check idle every 60s

    last_activity = time.time()

    while True:
        try:
            conn, addr = server.accept()
            last_activity = time.time()
            conn.settimeout(30)
            
            # Read initial data
            initial = conn.recv(65536).decode('utf-8')
            data = recv_all(conn, initial).strip()

            response = handle_command(data, desktop)
            
            if response == "QUIT_OK":
                conn.send(b"OK")
                conn.close()
                server.close()
                desktop.terminate()
                break
            
            conn.sendall(response.encode('utf-8'))
            conn.close()
        except socket.timeout:
            # Check if idle too long
            if time.time() - last_activity > IDLE_TIMEOUT:
                # Close all sessions
                for sid, (doc, _) in list(sessions.items()):
                    try:
                        doc.close(False)
                    except:
                        pass
                sessions.clear()
                server.close()
                desktop.terminate()
                break
            continue
        except Exception:
            break

g_exportedScripts = (start_server,)
'''


# Install instructions per platform
_INSTALL_INSTRUCTIONS = {
    "darwin": "brew install --cask libreoffice",
    "linux": "sudo apt install libreoffice libreoffice-calc  # or dnf/pacman equivalent",
    "win32": "Download from https://www.libreoffice.org/download/",
}

# Common Windows installation paths for LibreOffice
# Use soffice.com (console wrapper) instead of soffice.exe for proper CLI behavior
_WINDOWS_SOFFICE_PATHS = [
    Path(os.environ.get("PROGRAMFILES", "C:\\Program Files"))
    / "LibreOffice"
    / "program"
    / "soffice.com",
    Path(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"))
    / "LibreOffice"
    / "program"
    / "soffice.com",
    Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local")))
    / "Programs"
    / "LibreOffice"
    / "program"
    / "soffice.com",
]


@cache
def get_soffice_path() -> str | None:
    """Get the path to soffice executable, checking common locations on Windows."""
    # On Windows, prefer soffice.com (console wrapper) for proper CLI behavior
    if platform == "win32":
        # Check common installation directories first
        for candidate in _WINDOWS_SOFFICE_PATHS:
            if candidate.exists():
                return str(candidate)
        # Try PATH but look for .com version
        path = shutil.which("soffice.com")
        if path:
            return path

    # Try PATH (works for macOS/Linux)
    path = shutil.which("soffice")
    if path:
        return path

    return None


@cache
def _libreoffice_available() -> bool:
    """Check if LibreOffice is available (cached)."""
    return get_soffice_path() is not None


def ensure_libreoffice_installed() -> None:
    """Check if LibreOffice is available, raise with install instructions if not."""
    if not _libreoffice_available():
        instructions = _INSTALL_INSTRUCTIONS.get(
            platform, "Install LibreOffice from https://www.libreoffice.org/download/"
        )
        raise LibreOfficeNotFoundError(
            f"LibreOffice is required for formula recalculation but 'soffice' was not found.\n\n"
            f"Install it:\n  {instructions}"
        )


def send_daemon_command(cmd: str, timeout: float | None = None) -> str:
    """Send a command to the daemon and return the response.

    Commands are terminated with newline. Responses are read until
    connection closes or a complete response is received.
    """
    config = _get_config()
    if timeout is None:
        timeout = config.socket_timeout

    # Read port from file (written by daemon on startup)
    port = get_daemon_port()
    if port is None:
        raise RecalcError("Daemon port file not found - daemon may not be running")

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(timeout)
            sock.connect((config.daemon_host, port))
            # Send command with newline terminator
            sock.send((cmd + "\n").encode("utf-8"))
            # Read response (may be large for bulk data)
            chunks = []
            while True:
                try:
                    chunk = sock.recv(65536)
                    if not chunk:
                        break
                    chunks.append(chunk)
                except TimeoutError:
                    break
            return b"".join(chunks).decode("utf-8")
    except (TimeoutError, ConnectionRefusedError, OSError) as e:
        raise RecalcError(f"Failed to connect to daemon: {e}") from e


def is_daemon_running(cleanup_if_stale: bool = True) -> bool:
    """
    Check if the daemon is running and responsive.

    Args:
        cleanup_if_stale: If True, clean up PID/port files when daemon
            is not responding. Set to False during startup polling.
    """
    try:
        response = send_daemon_command("PING", timeout=2)
        return response == "PONG"
    except RecalcError:
        if cleanup_if_stale:
            cleanup_daemon_files()
        return False
