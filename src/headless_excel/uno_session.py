"""UNO session API for direct LibreOffice code execution.

This provides a fast, agent-friendly interface for Excel/Calc operations
by executing Python code directly inside the LibreOffice process.
"""

from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any

from headless_excel.daemon.api import start_daemon
from headless_excel.daemon.base import is_daemon_running, send_daemon_command
from headless_excel.errors import RecalcError


class UnoExecError(RecalcError):
    """Error during UNO code execution."""

    pass


class UnoSession:
    """Direct UNO code execution session.

    Provides these objects in exec() scope:
    - doc: The SpreadsheetDocument
    - desktop: The LibreOffice Desktop
    - uno: The uno module
    - json: The json module

    Set `result` variable to return a value from exec().

    Example:
        with UnoSession("model.xlsx") as s:
            s.exec('''
                sheet = doc.getSheets().getByIndex(0)
                sheet.getCellByPosition(0, 0).setValue(100)
                sheet.getCellByPosition(1, 0).setFormula("=A1*2")
                doc.calculateAll()
                result = sheet.getCellByPosition(1, 0).getValue()
            ''')
    """

    def __init__(self, path: str | Path, ensure_daemon: bool = True):
        """Create a UNO session for a workbook.

        Args:
            path: Path to the workbook (will be created if doesn't exist)
            ensure_daemon: If True, start daemon if not running
        """
        self.path = Path(path).absolute()
        self._session_id: str | None = None
        self._ensure_daemon = ensure_daemon

    def __enter__(self) -> UnoSession:
        """Open the session."""
        if self._ensure_daemon and not is_daemon_running():
            start_daemon(wait=True)

        # Open or create file in LibreOffice
        if self.path.exists():
            resp = send_daemon_command(f"OPEN:{self.path}")
        else:
            resp = send_daemon_command(f"NEW:{self.path}")

        if resp.startswith("ERROR:"):
            raise RecalcError(f"Failed to open session: {resp}")

        self._session_id = resp.split(":")[1]
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Save and close the session."""
        if self._session_id:
            send_daemon_command(f"SAVE:{self._session_id}")
            send_daemon_command(f"CLOSE:{self._session_id}")
            self._session_id = None

    def exec(self, code: str) -> Any:
        """Execute Python code inside LibreOffice.

        Available in scope:
        - doc: SpreadsheetDocument
        - desktop: Desktop service
        - uno: uno module
        - json: json module

        Set `result` variable in your code to return a value.

        Returns:
            The value of `result` variable if set, else None
        """
        if not self._session_id:
            raise RecalcError("Session not open")

        code_b64 = base64.b64encode(code.encode()).decode()
        resp = send_daemon_command(f"EXEC:{self._session_id}:{code_b64}")

        if resp.startswith("EXEC_ERROR:"):
            raise UnoExecError(resp[11:])
        if resp.startswith("ERROR:"):
            raise RecalcError(resp[6:])
        if resp.startswith("RESULT:"):
            return json.loads(resp[7:])
        return None

    def save(self) -> None:
        """Save the document to disk."""
        if not self._session_id:
            raise RecalcError("Session not open")
        resp = send_daemon_command(f"SAVE:{self._session_id}")
        if resp.startswith("ERROR:"):
            raise RecalcError(f"Failed to save: {resp}")


def uno_session(path: str | Path) -> UnoSession:
    """Create a UNO session context manager.

    Example:
        with uno_session("model.xlsx") as s:
            s.exec('''
                sheet = doc.getSheets().getByIndex(0)
                sheet.getCellByPosition(0, 0).setValue(42)
                result = sheet.getCellByPosition(0, 0).getValue()
            ''')
    """
    return UnoSession(path)
