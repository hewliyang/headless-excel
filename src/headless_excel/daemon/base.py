"""Shared daemon constants and utilities."""

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
