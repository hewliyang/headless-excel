"""LibreOffice integration for formula recalculation."""

import os
import signal
import subprocess
from pathlib import Path
from sys import platform

from headless_excel.daemon import is_daemon_running
from headless_excel.daemon.base import (
    ensure_libreoffice_installed,
    get_soffice_path,
    send_daemon_command,
)
from headless_excel.errors import RecalcError

# =============================================================================
# Cold-start recalc via Basic macro
# =============================================================================

MACRO_MODULE_NAME = "HeadlessExcel"
MACRO_SUB_NAME = "RecalculateAndSave"

MACRO_TEMPLATE = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{MACRO_MODULE_NAME}" script:language="StarBasic">
    Sub {MACRO_SUB_NAME}()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def _get_libreoffice_user_dir() -> Path:
    """Get the LibreOffice user directory for the current platform."""
    if platform == "darwin":
        return Path.home() / "Library/Application Support/LibreOffice/4/user"
    elif platform == "win32":
        appdata = os.environ.get("APPDATA", str(Path.home() / "AppData/Roaming"))
        return Path(appdata) / "LibreOffice/4/user"
    return Path.home() / ".config/libreoffice/4/user"


def _get_basic_macro_dir() -> Path:
    """Get the full path to the Standard Basic macro directory."""
    return _get_libreoffice_user_dir() / "basic/Standard"


def _get_basic_macro_file() -> Path:
    """Get the full path to the Basic macro file."""
    return _get_basic_macro_dir() / f"{MACRO_MODULE_NAME}.xba"


def _get_macro_uri() -> str:
    """Get the URI to call the Basic macro."""
    return f"macro:///Standard.{MACRO_MODULE_NAME}.{MACRO_SUB_NAME}"


def _run_soffice(cmd: list[str], timeout: int) -> tuple[int, str]:
    """Run soffice command with timeout and process group termination."""
    if platform == "win32":
        proc = subprocess.Popen(
            cmd,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    else:
        proc = subprocess.Popen(
            cmd,
            start_new_session=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    try:
        _, stderr = proc.communicate(timeout=timeout)
        return proc.returncode, stderr or ""
    except subprocess.TimeoutExpired:
        if platform == "win32":
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True,
            )
        else:
            os.killpg(proc.pid, signal.SIGTERM)
        proc.wait()
        raise RecalcError(f"LibreOffice command timed out after {timeout} seconds")


def _register_module_in_script_xlb(macro_dir: Path) -> bool:
    """Register the module in script.xlb so LibreOffice recognizes it."""
    script_xlb = macro_dir / "script.xlb"

    if not script_xlb.exists():
        return False

    content = script_xlb.read_text()

    if f'library:name="{MACRO_MODULE_NAME}"' in content:
        return True

    new_element = f' <library:element library:name="{MACRO_MODULE_NAME}"/>\n'
    content = content.replace("</library:library>", new_element + "</library:library>")
    script_xlb.write_text(content)
    return True


def setup_libreoffice_macro() -> bool:
    """Setup LibreOffice Basic macro for recalculation if not already configured."""
    macro_dir = _get_basic_macro_dir()
    macro_file = _get_basic_macro_file()

    if macro_file.exists():
        content = macro_file.read_text()
        if MACRO_SUB_NAME in content:
            _register_module_in_script_xlb(macro_dir)
            return True

    if not macro_dir.exists():
        soffice = get_soffice_path()
        if not soffice:
            return False
        try:
            _run_soffice([soffice, "--headless", "--terminate_after_init"], timeout=10)
        except RecalcError:
            return False
        macro_dir.mkdir(parents=True, exist_ok=True)

    try:
        macro_file.write_text(MACRO_TEMPLATE)
        _register_module_in_script_xlb(macro_dir)
        return True
    except Exception:
        return False


def _cold_recalc(filename: str | Path, timeout: int = 30) -> None:
    """Recalculate via cold-start (spawns new soffice process)."""
    ensure_libreoffice_installed()
    filepath = Path(filename)

    if not filepath.exists():
        raise FileNotFoundError(f"File {filename} does not exist")

    if not setup_libreoffice_macro():
        raise RecalcError("Failed to setup LibreOffice macro")

    soffice = get_soffice_path()
    if soffice is None:
        raise RecalcError("LibreOffice soffice executable not found")

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        _get_macro_uri(),
        str(filepath.absolute()),
    ]

    returncode, stderr = _run_soffice(cmd, timeout=timeout)

    if returncode != 0:
        if MACRO_MODULE_NAME in stderr and MACRO_SUB_NAME not in stderr:
            raise RecalcError("LibreOffice macro not configured properly")
        else:
            raise RecalcError(stderr or "Unknown error during recalculation")


def daemon_recalc(filename: str | Path, timeout: float = 30) -> None:
    """
    Recalculate formulas using the daemon.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation

    Raises:
        RecalcError: If recalculation fails or daemon not running
        FileNotFoundError: If file doesn't exist
    """
    filepath = Path(filename)
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filename}")

    if not is_daemon_running():
        raise RecalcError(
            "LibreOffice daemon is not running. "
            "Start it with: headless-excel libreoffice start"
        )

    abs_path = str(filepath.absolute())
    response = send_daemon_command(f"RECALC:{abs_path}", timeout=timeout)

    if response == "OK":
        return
    elif response.startswith("ERROR:"):
        raise RecalcError(response[6:])
    else:
        raise RecalcError(f"Unexpected response: {response}")


def recalc(filename: str | Path, timeout: int = 30) -> None:
    """
    Recalculate formulas in Excel file via LibreOffice.

    If the LibreOffice daemon is running, uses it for faster recalculation.
    Otherwise, falls back to cold-start mode.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation (seconds)

    Raises:
        LibreOfficeNotFoundError: If LibreOffice is not installed
        RecalcError: If recalculation fails
        FileNotFoundError: If the file does not exist
    """
    if is_daemon_running():
        daemon_recalc(filename, timeout=timeout)
    else:
        _cold_recalc(filename, timeout=timeout)
