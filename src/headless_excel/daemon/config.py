from __future__ import annotations

import os
import random
import socket
from dataclasses import dataclass

from headless_excel.constants import HEADLESS_EXCEL_DIR

# Private/dynamic port range (IANA)
PRIVATE_PORT_START = 49152
PRIVATE_PORT_END = 65535

# Daemon state files
PID_FILE = HEADLESS_EXCEL_DIR / "daemon.pid"
PORT_FILE = HEADLESS_EXCEL_DIR / "daemon.port"


@dataclass(frozen=True)
class Config:
    """Daemon configuration settings."""

    daemon_host: str = "127.0.0.1"
    daemon_port: int = PRIVATE_PORT_START
    uno_port: int = 2002
    socket_timeout: int = 30
    idle_timeout: int = 300  # seconds before daemon auto-exits

    @classmethod
    def from_env(cls) -> Config:
        """Create config from environment variables."""
        return cls(
            daemon_port=int(
                os.environ.get("HEADLESS_EXCEL_PORT", str(PRIVATE_PORT_START))
            ),
            idle_timeout=int(os.environ.get("HEADLESS_EXCEL_IDLE_TIMEOUT", "300")),
        )


def find_free_port(preferred: int | None = None, max_attempts: int = 100) -> int:
    """
    Find a free port in the private range.

    If preferred is specified, tries that first. Otherwise picks randomly
    from the private range (49152-65535).

    Args:
        preferred: Specific port to try first (optional)
        max_attempts: Maximum random attempts before giving up

    Returns:
        A free port number

    Raises:
        RuntimeError: If no free port found after max_attempts
    """

    def try_bind(port: int) -> bool:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
                return True
        except OSError:
            return False

    # Try preferred port first if specified
    if preferred is not None and PRIVATE_PORT_START <= preferred <= PRIVATE_PORT_END:
        if try_bind(preferred):
            return preferred

    # Pick random ports in the private range
    for _ in range(max_attempts):
        port = random.randint(PRIVATE_PORT_START, PRIVATE_PORT_END)
        if try_bind(port):
            return port

    raise RuntimeError(
        f"No free port found after {max_attempts} attempts "
        f"in range {PRIVATE_PORT_START}-{PRIVATE_PORT_END}"
    )


def get_daemon_port() -> int | None:
    """
    Get the port the daemon is running on by reading the port file.

    Returns:
        Port number if file exists and is valid, None otherwise
    """
    if not PORT_FILE.exists():
        return None
    try:
        return int(PORT_FILE.read_text().strip())
    except (ValueError, OSError):
        return None


def write_daemon_port(port: int) -> None:
    """Write the daemon port to the port file."""
    PORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    PORT_FILE.write_text(str(port))
