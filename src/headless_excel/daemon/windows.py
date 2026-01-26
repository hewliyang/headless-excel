"""Windows daemon implementation using Python macro inside LibreOffice.

On Windows, LibreOffice bundles its own Python interpreter, similar to macOS.
This allows us to run Python macros directly inside LibreOffice.
"""

import os
import subprocess
import time
from pathlib import Path

from headless_excel.daemon.base import (
    PID_FILE,
    cleanup_daemon_files,
    get_soffice_path,
    get_unobridge_macro,
    is_daemon_running,
    send_daemon_command,
)
from headless_excel.daemon.config import Config, find_free_port, write_daemon_port
from headless_excel.errors import RecalcError


def _get_libreoffice_user_dir() -> Path:
    """Get the LibreOffice user directory for Windows."""
    appdata = os.environ.get("APPDATA", str(Path.home() / "AppData/Roaming"))
    return Path(appdata) / "LibreOffice/4/user"


def _get_python_macro_dir() -> Path:
    """Get the LibreOffice Python macro directory."""
    return _get_libreoffice_user_dir() / "Scripts/python"


def _install_daemon_macro(config: Config) -> None:
    """Install the unobridge Python macro with the given config."""
    macro_dir = _get_python_macro_dir()
    macro_file = macro_dir / "unobridge.py"
    macro_content = get_unobridge_macro(config)

    macro_dir.mkdir(parents=True, exist_ok=True)
    macro_file.write_text(macro_content)


def start_daemon_windows(wait: bool, timeout: float) -> int:
    """Start the daemon on Windows using Python macro inside LibreOffice."""
    soffice = get_soffice_path()
    if not soffice:
        raise RecalcError("LibreOffice soffice executable not found")

    # Find a free port in the private range
    config = Config.from_env()
    port = find_free_port(config.daemon_port)

    # Create config with the actual port we'll use
    config = Config(
        daemon_port=port,
        idle_timeout=config.idle_timeout,
    )

    # Install macro with this port
    _install_daemon_macro(config)

    cmd = [
        soffice,
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
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,  # type: ignore[attr-defined]
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

        stop_daemon_windows()
        raise RecalcError(f"Daemon failed to start within {timeout} seconds")

    return proc.pid


def stop_daemon_windows() -> bool:
    """Stop the Windows daemon."""
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
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
            )
            stopped = True
        except (ValueError, subprocess.SubprocessError):
            pass

    cleanup_daemon_files()

    # Kill any orphaned soffice processes
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "soffice.bin", "/T"],
            capture_output=True,
        )
    except subprocess.SubprocessError:
        pass

    return stopped
