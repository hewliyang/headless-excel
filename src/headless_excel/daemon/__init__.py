"""LibreOffice daemon management for fast recalculation."""

from headless_excel.daemon.api import start_daemon, stop_daemon
from headless_excel.daemon.base import (
    ensure_libreoffice_installed,
    get_soffice_path,
    is_daemon_running,
)
from headless_excel.daemon.config import Config

__all__ = [
    "Config",
    "ensure_libreoffice_installed",
    "get_soffice_path",
    "is_daemon_running",
    "start_daemon",
    "stop_daemon",
]
