"""Excel context manager with automatic sync and recalculation."""

from __future__ import annotations

from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, TypeVar

from openpyxl import Workbook, load_workbook

from headless_excel.errors import FormulaError, RecalcError, SyncError
from headless_excel.proxy import WorkbookProxy, WorksheetProxy
from headless_excel.recalc import recalc

T = TypeVar("T")


@dataclass
class SyncResult:
    """Result of a sync operation.

    Attributes:
        success: Whether sync completed without formula errors
        total_errors: Number of formula errors found
        errors: Dict mapping error type to list of cell locations
    """

    success: bool
    total_errors: int = 0
    errors: dict[str, list[str]] = field(default_factory=dict)

    def raise_on_errors(self) -> None:
        """Raise FormulaError if any errors were found."""
        if not self.success:
            raise FormulaError(errors=self.errors, total=self.total_errors)


class ExcelContext:
    """Context for Excel operations with automatic sync support.

    Similar to OfficeJS Excel.run() pattern - make changes, then sync()
    to save, recalculate, and reload materialized values.

    Example:
        with ExcelContext("model.xlsx") as ctx:
            ctx.active["A1"] = "=SUM(B1:B10)"

            result = ctx.sync()  # saves, recalcs via LibreOffice, reloads
            if not result.success:
                print(f"Errors: {result.errors}")

            # After sync, .value returns materialized results automatically
            print(ctx.active["A1"].value)  # 450 (computed value)
    """

    def __init__(
        self,
        path: str | Path,
        create: bool = False,
        recalc_timeout: int = 30,
    ) -> None:
        """Initialize Excel context.

        Args:
            path: Path to the Excel file
            create: If True, create new workbook; if False, load existing
            recalc_timeout: Timeout in seconds for LibreOffice recalculation
        """
        self.path = Path(path)
        self._recalc_timeout = recalc_timeout
        self._workbook: Workbook | None = None
        self._values_workbook: Workbook | None = None
        self._proxy: WorkbookProxy | None = None
        self._dirty = False

        if create:
            self._workbook = Workbook()
        elif self.path.exists():
            self._workbook = load_workbook(self.path)
        else:
            raise FileNotFoundError(f"File not found: {path}")

        self._proxy = WorkbookProxy(self._workbook, None)

    @property
    def workbook(self) -> WorkbookProxy:
        """The workbook proxy (returns materialized values after sync)."""
        if self._proxy is None:
            raise RuntimeError("Context not initialized")
        self._dirty = True
        return self._proxy

    @property
    def wb(self) -> WorkbookProxy:
        """Alias for workbook."""
        return self.workbook

    @property
    def values(self) -> Workbook:
        """The raw workbook with materialized values (read after sync).

        Note: Prefer using workbook/active/sheet() which return proxied
        cells that automatically show materialized values after sync.
        This property is for advanced use cases.
        """
        if self._values_workbook is None:
            raise RuntimeError("Call sync() first to materialize values")
        return self._values_workbook

    @property
    def active(self) -> WorksheetProxy:
        """The active worksheet (returns materialized values after sync)."""
        ws = self.workbook.active
        if ws is None:
            raise RuntimeError("No active worksheet")
        return ws

    def sheet(self, name: str) -> WorksheetProxy:
        """Get worksheet by name."""
        return self.workbook[name]

    def create_sheet(self, name: str, index: int | None = None) -> WorksheetProxy:
        """Create a new worksheet."""
        self._dirty = True
        return self.workbook.create_sheet(name, index)

    def sync(self, raise_on_errors: bool = False) -> SyncResult:
        """Save, recalculate via LibreOffice, and reload with materialized values.

        This is the core operation - similar to OfficeJS ctx.sync().

        Args:
            raise_on_errors: If True, raise FormulaError when errors found

        Returns:
            SyncResult with success status and any formula errors

        Raises:
            SyncError: If save or recalc fails
            FormulaError: If raise_on_errors=True and formula errors found
        """
        if self._workbook is None:
            raise RuntimeError("Context not initialized")

        # Save current workbook
        try:
            self._workbook.save(self.path)
        except Exception as e:
            raise SyncError(f"Failed to save workbook: {e}") from e

        # Recalculate via LibreOffice
        result = recalc(self.path, timeout=self._recalc_timeout)

        if "error" in result:
            raise RecalcError(result["error"])

        # Reload formula workbook (to pick up any LibreOffice changes)
        self._workbook = load_workbook(self.path)

        # Load values workbook
        self._values_workbook = load_workbook(self.path, data_only=True)

        # Update proxy with values workbook (preserves object identity)
        if self._proxy:
            self._proxy._formula_wb = self._workbook
            self._proxy._update_values_wb(self._values_workbook)
        else:
            self._proxy = WorkbookProxy(self._workbook, self._values_workbook)

        self._dirty = False

        # Build sync result
        errors: dict[str, list[str]] = {}
        for err_type, details in result.get("error_summary", {}).items():
            errors[err_type] = details.get("locations", [])

        sync_result = SyncResult(
            success=result.get("status") == "success",
            total_errors=result.get("total_errors", 0),
            errors=errors,
        )

        if raise_on_errors:
            sync_result.raise_on_errors()

        return sync_result

    def read_value(self, sheet: str, cell: str) -> Any:
        """Read a materialized value after sync.

        Args:
            sheet: Sheet name
            cell: Cell reference (e.g., "A1")

        Returns:
            The computed value of the cell
        """
        if self._values_workbook is None:
            raise RuntimeError("Call sync() first to materialize values")
        return self._values_workbook[sheet][cell].value

    def read_range(self, sheet: str, range_ref: str) -> list[list[Any]]:
        """Read a range of materialized values after sync.

        Args:
            sheet: Sheet name
            range_ref: Range reference (e.g., "A1:C3")

        Returns:
            2D list of computed values
        """
        if self._values_workbook is None:
            raise RuntimeError("Call sync() first to materialize values")

        ws = self._values_workbook[sheet]
        rows = []
        for row in ws[range_ref]:
            rows.append([cell.value for cell in row])
        return rows

    def close(self) -> None:
        """Close both workbooks."""
        if self._workbook:
            self._workbook.close()
            self._workbook = None
        if self._values_workbook:
            self._values_workbook.close()
            self._values_workbook = None

    def __enter__(self) -> ExcelContext:
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()


@contextmanager
def run(
    path: str | Path,
    create: bool = False,
    auto_sync: bool = True,
    raise_on_errors: bool = False,
    recalc_timeout: int = 30,
) -> Generator[ExcelContext, None, None]:
    """Run Excel operations with automatic context management.

    Similar to OfficeJS Excel.run() - provides a context, handles cleanup,
    and optionally auto-syncs on exit.

    Example:
        with run("model.xlsx") as ctx:
            ctx.active["A1"] = 100
            ctx.active["A2"] = "=A1*2"
            # auto-syncs on exit

        # Or with explicit sync:
        with run("model.xlsx", auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.sync()  # manual sync
            print(ctx.values.active["A1"].value)

    Args:
        path: Path to the Excel file
        create: If True, create new workbook
        auto_sync: If True, automatically sync on context exit
        raise_on_errors: If True, raise FormulaError on sync errors
        recalc_timeout: Timeout in seconds for LibreOffice recalculation

    Yields:
        ExcelContext for operations
    """
    ctx = ExcelContext(path, create=create, recalc_timeout=recalc_timeout)
    try:
        yield ctx
        if auto_sync and ctx._dirty:
            ctx.sync(raise_on_errors=raise_on_errors)
    finally:
        ctx.close()
