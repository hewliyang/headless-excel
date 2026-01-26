"""Linux daemon implementation using helper process + UNO socket.

On Linux, distro-packaged LibreOffice links libpyuno.so against system Python.
When a venv with a different Python version is active, this causes ABI mismatch
and crashes. We avoid this by:

1. Starting soffice in socket mode (no Python macro execution inside LO)
2. Running a helper script with system Python (/usr/bin/python3) that has uno
3. The helper connects to LO via UNO and exposes the same TCP protocol
"""

import os
import signal
import subprocess
import time
from pathlib import Path

from headless_excel.daemon.base import (
    PID_FILE,
    cleanup_daemon_files,
    is_daemon_running,
    send_daemon_command,
)
from headless_excel.daemon.config import Config, find_free_port, write_daemon_port
from headless_excel.errors import RecalcError


def _get_helper_script_path() -> Path:
    """Get the path to the Linux helper script."""
    return Path(__file__).parent / "linux_helper.py"


def _check_uno_available() -> bool:
    """Check if the uno module is available in system Python."""
    try:
        result = subprocess.run(
            ["/usr/bin/python3", "-c", "import uno"],
            capture_output=True,
            timeout=5,
        )
        return result.returncode == 0
    except Exception:
        return False


def start_daemon_linux(wait: bool, timeout: float) -> int:
    """Start the daemon on Linux using helper process + UNO socket."""
    if not _check_uno_available():
        raise RecalcError(
            "python3-uno is required for daemon mode on Linux.\n"
            "Install it with: sudo apt install python3-uno  # or equivalent"
        )

    helper_path = _get_helper_script_path()
    if not helper_path.exists():
        raise RecalcError(f"Helper script not found: {helper_path}")

    # Find a free port in the private range
    config = Config.from_env()
    port = find_free_port(config.daemon_port)

    # Pass config via environment to the helper process
    env = os.environ.copy()
    env["HEADLESS_EXCEL_PORT"] = str(port)
    env["HEADLESS_EXCEL_IDLE_TIMEOUT"] = str(config.idle_timeout)

    proc = subprocess.Popen(
        ["/usr/bin/python3", str(helper_path)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        env=env,
    )

    PID_FILE.parent.mkdir(parents=True, exist_ok=True)
    PID_FILE.write_text(str(proc.pid))
    write_daemon_port(port)

    if wait:
        start_time = time.time()
        while time.time() - start_time < timeout:
            if is_daemon_running(cleanup_if_stale=False):
                return proc.pid
            time.sleep(0.2)

        stop_daemon_linux()
        raise RecalcError(f"Daemon failed to start within {timeout} seconds")

    return proc.pid


def stop_daemon_linux() -> bool:
    """Stop the Linux daemon (helper process + soffice)."""
    stopped = False

    # Send QUIT command if daemon is responsive
    if is_daemon_running():
        try:
            send_daemon_command("QUIT", timeout=5)
            stopped = True
            time.sleep(0.5)
        except RecalcError:
            pass

    # Kill helper process via PID file
    if PID_FILE.exists():
        try:
            pid = int(PID_FILE.read_text().strip())
            os.killpg(pid, signal.SIGTERM)
            stopped = True
        except (ProcessLookupError, ValueError, PermissionError):
            pass

    cleanup_daemon_files()

    # Kill any orphaned soffice processes started by helper
    try:
        subprocess.run(
            ["pkill", "-f", "accept=socket.*2002"],
            capture_output=True,
            timeout=5,
        )
    except Exception:
        pass

    return stopped
