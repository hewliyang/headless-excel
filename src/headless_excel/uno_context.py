"""UNO-based context with hooks support.

Provides a context manager similar to ExcelContext but backed entirely by UNO.
Agents write UNO code directly, with thin helpers for common pain points.
"""

from __future__ import annotations

import base64
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from headless_excel.daemon.api import start_daemon
from headless_excel.daemon.base import is_daemon_running, send_daemon_command
from headless_excel.errors import RecalcError
from headless_excel.hooks import (
    run_on_exit_hooks,
    run_on_open_hooks,
    run_post_sync_hooks,
    run_pre_sync_hooks,
)


class UnoExecError(RecalcError):
    """Error during UNO code execution."""

    pass


# Excel error patterns
EXCEL_ERRORS = (
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
)


@dataclass
class SyncResult:
    """Result of a sync operation."""

    success: bool
    total_errors: int
    errors: dict[str, list[str]] = field(default_factory=dict)

    def raise_on_errors(self) -> None:
        if not self.success:
            from headless_excel.errors import FormulaError

            raise FormulaError(
                f"Found {self.total_errors} formula errors: {self.errors}"
            )


# Helper code injected into every exec
HELPER_CODE = '''
def cell(ref):
    """Parse A1 notation to (col, row) 0-indexed tuple."""
    match = __import__("re").match(r"^([A-Z]+)(\\d+)$", ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    col_str, row_str = match.groups()
    col = 0
    for c in col_str:
        col = col * 26 + (ord(c) - ord("A") + 1)
    return col - 1, int(row_str) - 1

def rng(ref):
    """Parse A1:B2 notation to (c1, r1, c2, r2) 0-indexed tuple."""
    if ":" in ref:
        start, end = ref.split(":")
        c1, r1 = cell(start)
        c2, r2 = cell(end)
        return c1, r1, c2, r2
    else:
        c, r = cell(ref)
        return c, r, c, r

def sheet(name_or_idx=0):
    """Get sheet by name or index."""
    sheets = doc.getSheets()
    if isinstance(name_or_idx, int):
        return sheets.getByIndex(name_or_idx)
    return sheets.getByName(name_or_idx)

def recalc():
    """Recalculate all formulas."""
    doc.calculateAll()

def sheets():
    """List all sheet names."""
    s = doc.getSheets()
    return [s.getByIndex(i).getName() for i in range(s.getCount())]
'''


class UnoContext:
    """UNO-based Excel context with hooks support.

    Usage:
        with UnoContext("model.xlsx") as ctx:
            ctx.exec('''
                s = sheet(0)
                c, r = cell("A1")
                s.getCellByPosition(c, r).setValue(100)
                recalc()
            ''')

            result = ctx.sync()  # triggers hooks

    Helpers available in exec():
        - cell("A1") -> (col, row) 0-indexed
        - rng("A1:C3") -> (c1, r1, c2, r2) 0-indexed
        - sheet(name_or_idx) -> UNO sheet object
        - sheets() -> list of sheet names
        - recalc() -> doc.calculateAll()
        - doc -> the SpreadsheetDocument
        - uno, json -> modules
    """

    def __init__(
        self,
        path: str | Path,
        auto_sync: bool = True,
        ensure_daemon: bool = True,
        hooks: bool = True,
    ):
        self.path = Path(path).absolute()
        self._auto_sync = auto_sync
        self._ensure_daemon = ensure_daemon
        self._hooks = hooks
        self._session_id: str | None = None
        self._synced = False

    def __enter__(self) -> UnoContext:
        if self._ensure_daemon and not is_daemon_running():
            start_daemon(wait=True)

        # Open or create file
        if self.path.exists():
            resp = send_daemon_command(f"OPEN:{self.path}")
        else:
            resp = send_daemon_command(f"NEW:{self.path}")

        if resp.startswith("ERROR:"):
            raise RecalcError(f"Failed to open: {resp}")

        self._session_id = resp.split(":")[1]

        # Run on_open hooks
        if self._hooks:
            run_on_open_hooks(self)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._session_id:
            # Auto-sync if enabled and no exception
            if self._auto_sync and exc_type is None:
                self.sync()

            # Run on_exit hooks
            if self._hooks:
                run_on_exit_hooks(self)

            # Save and close
            send_daemon_command(f"SAVE:{self._session_id}")
            send_daemon_command(f"CLOSE:{self._session_id}")
            self._session_id = None

    def exec(self, code: str) -> Any:
        """Execute UNO code inside LibreOffice.

        Available in scope:
            - doc: SpreadsheetDocument
            - uno, json: modules
            - cell(), rng(), sheet(), sheets(), recalc(): helpers

        Set `result` variable to return a value.
        """
        if not self._session_id:
            raise RecalcError("Context not open")

        # Prepend helper code
        full_code = HELPER_CODE + "\n" + code

        code_b64 = base64.b64encode(full_code.encode()).decode()
        resp = send_daemon_command(f"EXEC:{self._session_id}:{code_b64}")

        if resp.startswith("EXEC_ERROR:"):
            raise UnoExecError(resp[11:])
        if resp.startswith("ERROR:"):
            raise RecalcError(resp[6:])
        if resp.startswith("RESULT:"):
            return json.loads(resp[7:])
        return None

    def sync(self, raise_on_errors: bool = False) -> SyncResult:
        """Recalculate and check for errors.

        Triggers pre_sync and post_sync hooks.
        """
        if self._hooks:
            run_pre_sync_hooks(self)

        # Recalculate
        self.exec("recalc()")

        # Scan for errors
        errors = self._scan_errors()

        result = SyncResult(
            success=len(errors) == 0,
            total_errors=sum(len(v) for v in errors.values()),
            errors=errors,
        )

        self._synced = True

        if self._hooks:
            run_post_sync_hooks(self, result)

        if raise_on_errors and not result.success:
            result.raise_on_errors()

        return result

    def _scan_errors(
        self, max_rows: int = 1000, max_cols: int = 50
    ) -> dict[str, list[str]]:
        """Scan for formula errors across all sheets."""
        code = f"""
errors = {{}}
error_types = {list(EXCEL_ERRORS)!r}

for sheet_idx in range(doc.getSheets().getCount()):
    sh = doc.getSheets().getByIndex(sheet_idx)
    sheet_name = sh.getName()
    
    for row in range({max_rows}):
        for col in range({max_cols}):
            cell = sh.getCellByPosition(col, row)
            if cell.getError() != 0:
                # Get error string
                err_str = cell.getString()
                for et in error_types:
                    if et in err_str or err_str.startswith("#"):
                        if et not in errors:
                            errors[et] = []
                        # Convert to A1 notation
                        col_str = ""
                        c = col
                        while c >= 0:
                            col_str = chr(ord("A") + c % 26) + col_str
                            c = c // 26 - 1
                        coord = f"{{sheet_name}}!{{col_str}}{{row + 1}}"
                        errors[et].append(coord)
                        break

result = errors
"""
        return self.exec(code) or {}

    def save(self) -> None:
        """Save the document to disk."""
        if not self._session_id:
            raise RecalcError("Context not open")
        resp = send_daemon_command(f"SAVE:{self._session_id}")
        if resp.startswith("ERROR:"):
            raise RecalcError(f"Failed to save: {resp}")

    # Compatibility properties for hooks that might access these
    @property
    def workbook(self):
        """Not available in UNO mode - raises helpful error."""
        raise NotImplementedError(
            "ctx.workbook is not available in UNO mode. "
            "Use ctx.exec() to run UNO code directly."
        )

    @property
    def active(self):
        """Not available in UNO mode - raises helpful error."""
        raise NotImplementedError(
            "ctx.active is not available in UNO mode. "
            "Use ctx.exec('s = sheet(0)') to get the active sheet."
        )


def uno_context(path: str | Path, **kwargs) -> UnoContext:
    """Create a UNO context manager.

    Example:
        with uno_context("model.xlsx") as ctx:
            ctx.exec('''
                s = sheet(0)
                c, r = cell("A1")
                s.getCellByPosition(c, r).setValue(42)
                recalc()
                result = s.getCellByPosition(c, r).getValue()
            ''')
    """
    return UnoContext(path, **kwargs)
