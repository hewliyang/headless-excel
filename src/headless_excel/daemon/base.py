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

    Creates a TCP server and handles PING/RECALC/QUIT commands.
    Values for port and idle timeout are injected from current config.
    """
    if config is None:
        config = _get_config()
    return f'''\
"""TCP bridge for headless-excel recalculation daemon."""
import socket
import time
import uno

DAEMON_PORT = {config.daemon_port}
IDLE_TIMEOUT = {config.idle_timeout}

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
                    conn.send(f"ERROR:{{e}}".encode())
            else:
                conn.send(b"ERROR:Unknown command")
            conn.close()
        except socket.timeout:
            # Check if idle too long
            if time.time() - last_activity > IDLE_TIMEOUT:
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
    """Send a command to the daemon and return the response."""
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
            sock.send(cmd.encode("utf-8"))
            return sock.recv(4096).decode("utf-8")
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
