import os
import signal
import subprocess
from pathlib import Path
from sys import platform

from headless_excel.errors import RecalcError

# LibreOffice macro configuration
MACRO_MODULE_NAME = "HeadlessExcel"
MACRO_SUB_NAME = "RecalculateAndSave"

# Platform-specific LibreOffice user directories
LIBREOFFICE_USER_DIR_MACOS = (
    Path.home() / "Library/Application Support/LibreOffice/4/user"
)
LIBREOFFICE_USER_DIR_LINUX = Path.home() / ".config/libreoffice/4/user"

# Macro paths (relative to user dir)
BASIC_STANDARD_DIR = Path("basic/Standard")
MACRO_FILENAME = f"{MACRO_MODULE_NAME}.xba"
SCRIPT_XLB_FILENAME = "script.xlb"

# Macro content template
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
        return LIBREOFFICE_USER_DIR_MACOS
    return LIBREOFFICE_USER_DIR_LINUX


def _get_macro_dir() -> Path:
    """Get the full path to the Standard macro directory."""
    return _get_libreoffice_user_dir() / BASIC_STANDARD_DIR


def _get_macro_file() -> Path:
    """Get the full path to the macro file."""
    return _get_macro_dir() / MACRO_FILENAME


def _get_macro_uri() -> str:
    """Get the URI to call the macro."""
    return f"vnd.sun.star.script:Standard.{MACRO_MODULE_NAME}.{MACRO_SUB_NAME}?language=Basic&location=application"


def _run_soffice(cmd: list[str], timeout: int) -> tuple[int, str]:
    """Run soffice command with timeout and process group termination."""
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
        os.killpg(proc.pid, signal.SIGTERM)
        proc.wait()
        raise RecalcError(f"LibreOffice command timed out after {timeout} seconds")


def _register_module_in_script_xlb(macro_dir: Path) -> bool:
    """Register the module in script.xlb so LibreOffice recognizes it.

    Only appends to existing script.xlb - never creates or overwrites.

    Returns:
        True if registered successfully or already registered, False if script.xlb doesn't exist.
    """
    script_xlb = macro_dir / SCRIPT_XLB_FILENAME

    if not script_xlb.exists():
        # Don't create - let LibreOffice manage this file
        return False

    content = script_xlb.read_text()

    # Check if already registered
    if f'library:name="{MACRO_MODULE_NAME}"' in content:
        return True

    # Append our module element before </library:library>
    new_element = f' <library:element library:name="{MACRO_MODULE_NAME}"/>\n'
    content = content.replace("</library:library>", new_element + "</library:library>")
    script_xlb.write_text(content)
    return True


def setup_libreoffice_macro() -> bool:
    """Setup LibreOffice macro for recalculation if not already configured."""
    macro_dir = _get_macro_dir()
    macro_file = _get_macro_file()

    # Check if macro already exists and is correctly configured
    if macro_file.exists():
        content = macro_file.read_text()
        if MACRO_SUB_NAME in content:
            # Ensure it's registered in script.xlb
            _register_module_in_script_xlb(macro_dir)
            return True

    # Initialize LibreOffice to create user directory if needed
    if not macro_dir.exists():
        try:
            _run_soffice(
                ["soffice", "--headless", "--terminate_after_init"], timeout=10
            )
        except RecalcError:
            return False
        macro_dir.mkdir(parents=True, exist_ok=True)

    # Write macro file
    try:
        macro_file.write_text(MACRO_TEMPLATE)
        _register_module_in_script_xlb(macro_dir)
        return True
    except Exception:
        return False


def recalc(filename: str | Path, timeout: int = 30) -> None:
    """
    Recalculate formulas in Excel file via LibreOffice.

    Args:
        filename: Path to Excel file
        timeout: Maximum time to wait for recalculation (seconds)

    Raises:
        RecalcError: If recalculation fails
        FileNotFoundError: If the file does not exist
    """
    filepath = Path(filename)

    if not filepath.exists():
        raise FileNotFoundError(f"File {filename} does not exist")

    abs_path = str(filepath.absolute())

    if not setup_libreoffice_macro():
        raise RecalcError("Failed to setup LibreOffice macro")

    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        _get_macro_uri(),
        abs_path,
    ]

    returncode, stderr = _run_soffice(cmd, timeout=timeout)

    if returncode != 0:
        if MACRO_MODULE_NAME in stderr or MACRO_SUB_NAME not in stderr:
            raise RecalcError("LibreOffice macro not configured properly")
        else:
            raise RecalcError(stderr or "Unknown error during recalculation")
