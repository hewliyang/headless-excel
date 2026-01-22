"""macOS daemon implementation using Python macro inside LibreOffice.

On macOS, LibreOffice bundles its own Python framework and uses @loader_path
relative linking, making it completely isolated from system/venv Python.
This allows us to run Python macros directly inside LibreOffice.
"""

import os
import signal
import subprocess
import time
from pathlib import Path

from headless_excel.daemon.base import (
    PID_FILE,
    is_daemon_running,
    send_daemon_command,
)
from headless_excel.errors import RecalcError


def _get_libreoffice_user_dir() -> Path:
    """Get the LibreOffice user directory for macOS."""
    return Path.home() / "Library/Application Support/LibreOffice/4/user"


def _get_python_macro_dir() -> Path:
    """Get the LibreOffice Python macro directory."""
    return _get_libreoffice_user_dir() / "Scripts/python"


# Python macro that runs inside LibreOffice
# Creates a TCP server and handles PING/RECALC/QUIT commands
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


def start_daemon_macos(wait: bool, timeout: float) -> int:
    """Start the daemon on macOS using Python macro inside LibreOffice."""
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

        stop_daemon_macos()
        raise RecalcError(f"Daemon failed to start within {timeout} seconds")

    return proc.pid


def stop_daemon_macos() -> bool:
    """Stop the macOS daemon."""
    stopped = False

    if is_daemon_running():
        try:
            send_daemon_command("QUIT", timeout=5)
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
