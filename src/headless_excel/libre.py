"""LibreOffice integration for formula recalculation."""

import os
import shutil
import signal
import socket
import subprocess
import time
from functools import cache
from pathlib import Path
from sys import platform

from headless_excel.errors import LibreOfficeNotFoundError, RecalcError

# Install instructions per platform
_INSTALL_INSTRUCTIONS = {
    "darwin": "brew install --cask libreoffice",
    "linux": "sudo apt install libreoffice libreoffice-calc  # or dnf/pacman equivalent",
}


@cache
def _libreoffice_available() -> bool:
    """Check if LibreOffice is available (cached)."""
    return shutil.which("soffice") is not None


def ensure_libreoffice_installed() -> None:
    """Check if LibreOffice is available, raise with install instructions if not."""
    if not _libreoffice_available():
        instructions = _INSTALL_INSTRUCTIONS.get(
            platform, "Install LibreOffice from https://www.libreoffice.org/download/"
        )
        raise LibreOfficeNotFoundError(
            f"LibreOffice is required for formula recalculation but 'soffice' was not found in PATH.\n\n"
            f"Install it:\n  {instructions}"
        )


# =============================================================================
# Cold-start recalc via Basic macro
# =============================================================================

MACRO_MODULE_NAME = "HeadlessExcel"
MACRO_SUB_NAME = "RecalculateAndSave"

MACRO_TEMPLATE = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{MACRO_MODULE_NAME}" script:language="StarBasic">
    Sub {MACRO_SUB_NAME}()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def _get_libreoffice_user_dir() -> Path:
    """Get the LibreOffice user directory for the current platform."""
    if platform == "darwin":
        return Path.home() / "Library/Application Support/LibreOffice/4/user"
    return Path.home() / ".config/libreoffice/4/user"


def _get_basic_macro_dir() -> Path:
    """Get the full path to the Standard Basic macro directory."""
    return _get_libreoffice_user_dir() / "basic/Standard"


def _get_basic_macro_file() -> Path:
    """Get the full path to the Basic macro file."""
    return _get_basic_macro_dir() / f"{MACRO_MODULE_NAME}.xba"


def _get_macro_uri() -> str:
    """Get the URI to call the Basic macro."""
    return f"macro:///Standard.{MACRO_MODULE_NAME}.{MACRO_SUB_NAME}"


def _run_soffice(cmd: list[str], timeout: int) -> tuple[int, str]:
    """Run soffice command with timeout and process group termination."""
    proc = subprocess.Popen(
        cmd,
        start_new_session=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    try:
        _, stderr = proc.communicate(timeout=timeout)
        return proc.returncode, stderr or ""
    except subprocess.TimeoutExpired:
        os.killpg(proc.pid, signal.SIGTERM)
        proc.wait()
        raise RecalcError(f"LibreOffice command timed out after {timeout} seconds")


def _register_module_in_script_xlb(macro_dir: Path) -> bool:
    """Register the module in script.xlb so LibreOffice recognizes it."""
    script_xlb = macro_dir / "script.xlb"

    if not script_xlb.exists():
        return False

    content = script_xlb.read_text()

    if f'library:name="{MACRO_MODULE_NAME}"' in content:
        return True

    new_element = f' <library:element library:name="{MACRO_MODULE_NAME}"/>\n'
    content = content.replace("</library:library>", new_element + "</library:library>")
    script_xlb.write_text(content)
    return True


def setup_libreoffice_macro() -> bool:
    """Setup LibreOffice Basic macro for recalculation if not already configured."""
    macro_dir = _get_basic_macro_dir()
    macro_file = _get_basic_macro_file()

    if macro_file.exists():
        content = macro_file.read_text()
        if MACRO_SUB_NAME in content:
            _register_module_in_script_xlb(macro_dir)
            return True

    if not macro_dir.exists():
        try:
            _run_soffice(
                ["soffice", "--headless", "--terminate_after_init"], timeout=10
            )
        except RecalcError:
            return False
        macro_dir.mkdir(parents=True, exist_ok=True)

    try:
        macro_file.write_text(MACRO_TEMPLATE)
        _register_module_in_script_xlb(macro_dir)
        return True
    except Exception:
        return False


def _cold_recalc(filename: str | Path, timeout: int = 30) -> None:
    """Recalculate via cold-start (spawns new soffice process)."""
    ensure_libreoffice_installed()
    filepath = Path(filename)

    if not filepath.exists():
        raise FileNotFoundError(f"File {filename} does not exist")

    if not setup_libreoffice_macro():
        raise RecalcError("Failed to setup LibreOffice macro")

    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        _get_macro_uri(),
        str(filepath.absolute()),
    ]

    returncode, stderr = _run_soffice(cmd, timeout=timeout)

    if returncode != 0:
        if MACRO_MODULE_NAME in stderr and MACRO_SUB_NAME not in stderr:
            raise RecalcError("LibreOffice macro not configured properly")
        else:
            raise RecalcError(stderr or "Unknown error during recalculation")


# =============================================================================
# Daemon mode via Python macro
# =============================================================================

DAEMON_HOST = "127.0.0.1"
DAEMON_PORT = 2023
SOCKET_TIMEOUT = 30
PID_FILE = Path.home() / ".headless-excel" / "daemon.pid"

UNOBRIDGE_MACRO = '''\
"""TCP bridge for headless-excel recalculation daemon."""
import socket
import uno

def start_server(*args):
    """Start TCP server for recalc commands."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(('127.0.0.1', 2023))
    server.listen(5)

    while True:
        try:
            conn, addr = server.accept()
            conn.settimeout(30)
            data = conn.recv(4096).decode('utf-8').strip()

            if data == "PING":
                conn.send(b"PONG")
            elif data == "QUIT":
                conn.send(b"OK")
                conn.close()
                server.close()
                desktop.terminate()
                break
            elif data.startswith("RECALC:"):
                filepath = data[7:]
                try:
                    url = uno.systemPathToFileUrl(filepath)
                    doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
                    doc.calculateAll()
                    doc.store()
                    doc.close(True)
                    conn.send(b"OK")
                except Exception as e:
                    conn.send(f"ERROR:{e}".encode())
            else:
                conn.send(b"ERROR:Unknown command")
            conn.close()
        except socket.timeout:
            continue
        except Exception:
            break

g_exportedScripts = (start_server,)
'''


def _get_python_macro_dir() -> Path:
    """Get the LibreOffice Python macro directory."""
    return _get_libreoffice_user_dir() / "Scripts/python"


def _install_daemon_macro() -> bool:
    """Install the unobridge Python macro if not present."""
    macro_dir = _get_python_macro_dir()
    macro_file = macro_dir / "unobridge.py"

    if macro_file.exists():
        if macro_file.read_text() == UNOBRIDGE_MACRO:
            return True

    macro_dir.mkdir(parents=True, exist_ok=True)
    macro_file.write_text(UNOBRIDGE_MACRO)
    return True


def _send_daemon_command(cmd: str, timeout: float = SOCKET_TIMEOUT) -> str:
    """Send a command to the daemon and return the response."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(timeout)
            sock.connect((DAEMON_HOST, DAEMON_PORT))
            sock.send(cmd.encode("utf-8"))
            return sock.recv(4096).decode("utf-8")
    except (TimeoutError, ConnectionRefusedError, OSError) as e:
        raise RecalcError(f"Failed to connect to daemon: {e}") from e


def is_daemon_running() -> bool:
    """Check if the daemon is running and responsive."""
    try:
        response = _send_daemon_command("PING", timeout=2)
        return response == "PONG"
    except RecalcError:
        return False


def start_daemon(wait: bool = True, timeout: float = 15) -> int:
    """
    Start the LibreOffice daemon.

    Args:
        wait: If True, wait for daemon to be ready before returning
        timeout: Maximum time to wait for daemon to start

    Returns:
        PID of the daemon process

    Raises:
        LibreOfficeNotFoundError: If LibreOffice is not installed
        RecalcError: If daemon fails to start
    """
    ensure_libreoffice_installed()

    if is_daemon_running():
        if PID_FILE.exists():
            return int(PID_FILE.read_text().strip())
        return -1

    if not _install_daemon_macro():
        raise RecalcError("Failed to install LibreOffice macro")

    cmd = [
        "soffice",
        "--headless",
        "--invisible",
        "--nologo",
        "--norestore",
        "vnd.sun.star.script:unobridge.py$start_server?language=Python&location=user",
    ]

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )

    PID_FILE.parent.mkdir(parents=True, exist_ok=True)
    PID_FILE.write_text(str(proc.pid))

    if wait:
        start_time = time.time()
        while time.time() - start_time < timeout:
            if is_daemon_running():
                return proc.pid
            time.sleep(0.2)

        stop_daemon()
        raise RecalcError(f"Daemon failed to start within {timeout} seconds")

    return proc.pid


def stop_daemon() -> bool:
    """
    Stop the LibreOffice daemon.

    Returns:
        True if daemon was stopped, False if it wasn't running
    """
    stopped = False

    if is_daemon_running():
        try:
            _send_daemon_command("QUIT", timeout=5)
            stopped = True
            time.sleep(0.5)
        except RecalcError:
            pass

    if PID_FILE.exists():
        try:
            pid = int(PID_FILE.read_text().strip())
            os.killpg(pid, signal.SIGTERM)
            stopped = True
        except (ProcessLookupError, ValueError, PermissionError):
            pass
        PID_FILE.unlink(missing_ok=True)

    try:
        subprocess.run(
            ["pkill", "-f", "unobridge.py"],
            capture_output=True,
            timeout=5,
        )
    except Exception:
        pass

    return stopped


def daemon_recalc(filename: str | Path, timeout: float = SOCKET_TIMEOUT) -> None:
    """
    Recalculate formulas using the daemon.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation

    Raises:
        RecalcError: If recalculation fails or daemon not running
        FileNotFoundError: If file doesn't exist
    """
    filepath = Path(filename)
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filename}")

    if not is_daemon_running():
        raise RecalcError(
            "LibreOffice daemon is not running. "
            "Start it with: headless-excel libreoffice start"
        )

    abs_path = str(filepath.absolute())
    response = _send_daemon_command(f"RECALC:{abs_path}", timeout=timeout)

    if response == "OK":
        return
    elif response.startswith("ERROR:"):
        raise RecalcError(response[6:])
    else:
        raise RecalcError(f"Unexpected response: {response}")


# =============================================================================
# Public API
# =============================================================================


def recalc(filename: str | Path, timeout: int = 30) -> None:
    """
    Recalculate formulas in Excel file via LibreOffice.

    If the LibreOffice daemon is running, uses it for faster recalculation.
    Otherwise, falls back to cold-start mode.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation (seconds)

    Raises:
        LibreOfficeNotFoundError: If LibreOffice is not installed
        RecalcError: If recalculation fails
        FileNotFoundError: If the file does not exist
    """
    if is_daemon_running():
        daemon_recalc(filename, timeout=timeout)
    else:
        _cold_recalc(filename, timeout=timeout)
