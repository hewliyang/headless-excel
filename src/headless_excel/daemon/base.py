"""Shared daemon constants and utilities."""

import os
import shutil
import socket
from functools import cache
from pathlib import Path
from sys import platform

from headless_excel.errors import LibreOfficeNotFoundError, RecalcError

# Network configuration
DAEMON_HOST = "127.0.0.1"
DAEMON_PORT = 2023
UNO_PORT = 2002
SOCKET_TIMEOUT = 30

# File paths
PID_FILE = Path.home() / ".headless-excel" / "daemon.pid"

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


def send_daemon_command(cmd: str, timeout: float = SOCKET_TIMEOUT) -> str:
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
        response = send_daemon_command("PING", timeout=2)
        return response == "PONG"
    except RecalcError:
        return False
