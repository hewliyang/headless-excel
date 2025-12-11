"""Excel context manager with automatic sync and recalculation."""

from __future__ import annotations

import re
from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, TypeVar

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection

from headless_excel.errors import FormulaError, RecalcError, SyncError
from headless_excel.proxy import WorkbookProxy, WorksheetProxy
from headless_excel.recalc import recalc

T = TypeVar("T")

CELL_REF_PATTERN = re.compile(
    r"(?<![A-Za-z_])(?:([A-Za-z_][A-Za-z0-9_]*!)?)?\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?![A-Za-z0-9_])"
)


@dataclass
class ErrorDetail:
    """Detailed information about a formula error.

    Attributes:
        location: Cell location (e.g., 'Sheet1!A1')
        error: Error type (e.g., '#DIV/0!')
        formula: The formula that caused the error
        neighbors: Values of cells referenced in the formula
    """

    location: str
    error: str
    formula: str | None = None
    neighbors: dict[str, Any] = field(default_factory=dict)


@dataclass
class SyncResult:
    """Result of a sync operation.

    Attributes:
        success: Whether sync completed without formula errors
        total_errors: Number of formula errors found
        errors: Dict mapping error type to list of cell locations
        error_details: List of detailed error information (when available)
    """

    success: bool
    total_errors: int = 0
    errors: dict[str, list[str]] = field(default_factory=dict)
    error_details: list[ErrorDetail] = field(default_factory=list)

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
        for err_type, err_info in result.get("error_summary", {}).items():
            errors[err_type] = err_info.get("locations", [])

        error_details = self._build_error_details(errors)

        sync_result = SyncResult(
            success=result.get("status") == "success",
            total_errors=result.get("total_errors", 0),
            errors=errors,
            error_details=error_details,
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

    def apply_style(
        self,
        sheet: str,
        range_ref: str,
        font: Font | None = None,
        fill: PatternFill | None = None,
        alignment: Alignment | None = None,
        border: Border | None = None,
        protection: Protection | None = None,
        number_format: str | None = None,
    ) -> None:
        """Apply styles to a range of cells.

        Args:
            sheet: Sheet name
            range_ref: Cell or range reference (e.g., "A1" or "A1:G5")
            font: Font style to apply
            fill: Fill/background style to apply
            alignment: Alignment style to apply
            border: Border style to apply
            protection: Protection settings to apply
            number_format: Number format string (e.g., '#,##0;(#,##0);-')
        """
        if self._workbook is None:
            raise RuntimeError("Context not initialized")

        self._dirty = True
        ws = self._workbook[sheet]
        cells = ws[range_ref]

        if not isinstance(cells, tuple):
            cells = ((cells,),)

        for row in cells:
            for cell in row:
                if font is not None:
                    cell.font = font
                if fill is not None:
                    cell.fill = fill
                if alignment is not None:
                    cell.alignment = alignment
                if border is not None:
                    cell.border = border
                if protection is not None:
                    cell.protection = protection
                if number_format is not None:
                    cell.number_format = number_format

    def get_formulas(self, sheet: str | None = None) -> dict[str, str]:
        """Get all formulas in a sheet or the entire workbook.

        Args:
            sheet: Sheet name, or None for all sheets

        Returns:
            Dict mapping cell location to formula string.
            Keys are 'SheetName!A1' format when querying all sheets,
            or just 'A1' format when querying a single sheet.
        """
        if self._workbook is None:
            raise RuntimeError("Context not initialized")

        formulas: dict[str, str] = {}
        sheets = [sheet] if sheet else self._workbook.sheetnames

        for sheet_name in sheets:
            ws = self._workbook[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if (
                        cell.value is not None
                        and isinstance(cell.value, str)
                        and cell.value.startswith("=")
                    ):
                        if sheet:
                            formulas[cell.coordinate] = cell.value
                        else:
                            formulas[f"{sheet_name}!{cell.coordinate}"] = cell.value

        return formulas

    def _extract_cell_refs(self, formula: str, default_sheet: str) -> list[str]:
        """Extract cell references from a formula.

        Args:
            formula: Formula string (e.g., '=B1/C1' or '=Sheet2!A1+B1')
            default_sheet: Default sheet name for unqualified references

        Returns:
            List of fully qualified cell references (e.g., ['Sheet1!B1', 'Sheet1!C1'])
        """
        refs = []
        for match in CELL_REF_PATTERN.finditer(formula):
            sheet_prefix, col, row = match.groups()
            if sheet_prefix:
                sheet_name = sheet_prefix.rstrip("!")
            else:
                sheet_name = default_sheet
            refs.append(f"{sheet_name}!{col}{row}")
        return refs

    def _build_error_details(
        self, error_locations: dict[str, list[str]]
    ) -> list[ErrorDetail]:
        """Build detailed error information for each error location.

        Args:
            error_locations: Dict mapping error type to list of locations

        Returns:
            List of ErrorDetail with formula and neighbor information
        """
        if self._workbook is None or self._values_workbook is None:
            return []

        details = []
        for error_type, locations in error_locations.items():
            for location in locations:
                if "!" in location:
                    sheet_name, cell_ref = location.split("!", 1)
                else:
                    sheet_name = (
                        self._workbook.active.title
                        if self._workbook.active
                        else "Sheet1"
                    )
                    cell_ref = location

                formula = None
                neighbors: dict[str, Any] = {}

                try:
                    formula_ws = self._workbook[sheet_name]
                    formula_cell = formula_ws[cell_ref]
                    if (
                        formula_cell.value
                        and isinstance(formula_cell.value, str)
                        and formula_cell.value.startswith("=")
                    ):
                        formula = formula_cell.value

                        cell_refs = self._extract_cell_refs(formula, sheet_name)
                        for ref in cell_refs:
                            ref_sheet, ref_cell = ref.split("!", 1)
                            try:
                                val_ws = self._values_workbook[ref_sheet]
                                neighbors[ref] = val_ws[ref_cell].value
                            except (KeyError, AttributeError):
                                neighbors[ref] = None
                except (KeyError, AttributeError):
                    pass

                details.append(
                    ErrorDetail(
                        location=location,
                        error=error_type,
                        formula=formula,
                        neighbors=neighbors,
                    )
                )

        return details

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
