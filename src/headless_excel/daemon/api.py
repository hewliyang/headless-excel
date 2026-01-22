"""Cross-platform daemon API."""

from sys import platform

from headless_excel.daemon.base import (
    PID_FILE,
    ensure_libreoffice_installed,
    is_daemon_running,
)


def start_daemon(wait: bool = True, timeout: float = 15) -> int:
    """
    Start the LibreOffice daemon.

    On macOS, uses Python macro embedded in LibreOffice.
    On Linux, uses a helper process with system Python + UNO socket.

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

    if platform.startswith("linux"):
        from headless_excel.daemon.linux import start_daemon_linux

        return start_daemon_linux(wait, timeout)
    else:
        from headless_excel.daemon.macos import start_daemon_macos

        return start_daemon_macos(wait, timeout)


def stop_daemon() -> bool:
    """
    Stop the LibreOffice daemon.

    Returns:
        True if daemon was stopped, False if it wasn't running
    """
    if platform.startswith("linux"):
        from headless_excel.daemon.linux import stop_daemon_linux

        return stop_daemon_linux()
    else:
        from headless_excel.daemon.macos import stop_daemon_macos

        return stop_daemon_macos()
