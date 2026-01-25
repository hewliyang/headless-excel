"""LibreOffice daemon management for fast recalculation."""

from headless_excel.daemon.api import start_daemon, stop_daemon
from headless_excel.daemon.base import (
    DAEMON_HOST,
    DAEMON_PORT,
    ensure_libreoffice_installed,
    get_soffice_path,
    is_daemon_running,
)

__all__ = [
    "DAEMON_HOST",
    "DAEMON_PORT",
    "ensure_libreoffice_installed",
    "get_soffice_path",
    "is_daemon_running",
    "start_daemon",
    "stop_daemon",
]
