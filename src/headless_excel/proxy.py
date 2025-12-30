"""Proxy wrappers for ergonomic value access."""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Any

from openpyxl.cell.cell import Cell
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Protection,
)
from openpyxl.utils import column_index_from_string
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from headless_excel.errors import ColorLintViolation, ErrorScanResult
from headless_excel.formats import infer_financial_color

# Excel error types to detect
EXCEL_ERRORS = (
    "#VALUE!",
    "#DIV/0!",
    "#REF!",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
)

# Type alias for write callback
OnWriteCallback = Callable[[], None] | None

# Pattern for parsing range references like "A1:B2" or "A1"
RANGE_PATTERN = re.compile(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", re.IGNORECASE)


def _parse_range(ref: str) -> tuple[int, int, int, int]:
    """Parse range reference into (min_row, min_col, max_row, max_col).

    Args:
        ref: Range like "A1:B2" or single cell like "A1"

    Returns:
        Tuple of (min_row, min_col, max_row, max_col), 1-indexed

    Raises:
        ValueError: If ref is not a valid range format
    """
    match = RANGE_PATTERN.match(ref.upper())
    if not match:
        raise ValueError(f"Invalid range reference: {ref!r}")

    start_col, start_row, end_col, end_row = match.groups()
    min_col = column_index_from_string(start_col)
    min_row = int(start_row)

    if end_col and end_row:
        max_col = column_index_from_string(end_col)
        max_row = int(end_row)
    else:
        max_col = min_col
        max_row = min_row

    # Normalize so min <= max
    if min_row > max_row:
        min_row, max_row = max_row, min_row
    if min_col > max_col:
        min_col, max_col = max_col, min_col

    return min_row, min_col, max_row, max_col


class RangeProxy:
    """Proxy for a cell range supporting bulk read/write operations.

    Example:
        r = sheet.range("A1:C3")
        r.values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        r.apply_style(font=Font(bold=True))
    """

    def __init__(
        self,
        ws_proxy: WorksheetProxy,
        range_ref: str,
        on_write: OnWriteCallback = None,
    ) -> None:
        """Initialize range proxy.

        Args:
            ws_proxy: Parent WorksheetProxy
            range_ref: Range reference like "A1:B2" or "A1"
            on_write: Callback to invoke when write operations occur
        """
        self._ws = ws_proxy
        self._range_ref = range_ref
        self._on_write = on_write
        self._min_row, self._min_col, self._max_row, self._max_col = _parse_range(
            range_ref
        )

    @property
    def num_rows(self) -> int:
        """Number of rows in range."""
        return self._max_row - self._min_row + 1

    @property
    def num_cols(self) -> int:
        """Number of columns in range."""
        return self._max_col - self._min_col + 1

    @property
    def shape(self) -> tuple[int, int]:
        """Shape as (rows, cols)."""
        return (self.num_rows, self.num_cols)

    @property
    def values(self) -> list[list[Any]]:
        """Get 2D array of values from the range."""
        result = []
        for row_idx in range(self._min_row, self._max_row + 1):
            row_values = []
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws.cell(row_idx, col_idx)
                row_values.append(cell.value)
            result.append(row_values)
        return result

    @values.setter
    def values(self, data: list[list[Any]]) -> None:
        """Set 2D array of values to the range.

        Args:
            data: 2D list matching range dimensions

        Raises:
            ValueError: If data dimensions don't match range dimensions
        """
        # Validate dimensions
        if not isinstance(data, list):
            raise ValueError("Data must be a 2D list")

        if len(data) != self.num_rows:
            raise ValueError(
                f"Row count mismatch: got {len(data)}, expected {self.num_rows}"
            )

        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ValueError(f"Row {i} must be a list")
            if len(row) != self.num_cols:
                raise ValueError(
                    f"Column count mismatch in row {i}: got {len(row)}, expected {self.num_cols}"
                )

        # Mark dirty before writing
        if self._on_write:
            self._on_write()

        # Write values
        for row_offset, row_data in enumerate(data):
            for col_offset, value in enumerate(row_data):
                row_idx = self._min_row + row_offset
                col_idx = self._min_col + col_offset
                self._ws._formula_ws.cell(row_idx, col_idx).value = value

    @property
    def formulas(self) -> dict[str, str]:
        """Get formulas in range as {coordinate: formula}.

        Only cells containing formulas are included.
        More token-efficient than a 2D array for sparse formulas.

        Example:
            >>> ws.range("A1:C3").formulas
            {'B2': '=A1*2', 'C3': '=SUM(A1:B2)'}
        """
        result: dict[str, str] = {}
        for row_idx in range(self._min_row, self._max_row + 1):
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws._formula_ws.cell(row_idx, col_idx)
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    result[cell.coordinate] = val
        return result

    @property
    def formula_count(self) -> int:
        """Count of formulas in this range.

        Example:
            >>> ws.range("A1:C3").formula_count
            5
        """
        return len(self.formulas)

    def find_errors(self) -> ErrorScanResult:
        """Find formula errors in this range.

        Scans the values (materialized after sync) for Excel error values
        like #DIV/0!, #REF!, #NAME?, etc.

        Returns:
            ErrorScanResult with errors organized by type and total count.
            Has a nice __repr__ with truncation for agent-friendly output.

        Example:
            >>> result = ws.range("A1:C10").find_errors()
            >>> print(result)
            Formula errors (3):
              #DIV/0!: A3, B5
              #REF!: C10
        """
        if self._ws._values_ws is None:
            return ErrorScanResult()

        error_details: dict[str, list[str]] = {err: [] for err in EXCEL_ERRORS}

        for row_idx in range(self._min_row, self._max_row + 1):
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws._values_ws.cell(row_idx, col_idx)
                if cell.value is not None and isinstance(cell.value, str):
                    for err in EXCEL_ERRORS:
                        if err in cell.value:
                            error_details[err].append(cell.coordinate)
                            break

        # Filter to only include error types that have occurrences
        errors_by_type = {k: v for k, v in error_details.items() if v}
        total = sum(len(locs) for locs in errors_by_type.values())
        return ErrorScanResult(errors_by_type=errors_by_type, total_errors=total)

    def dump(self, show_formulas: bool = False) -> str:
        """Print range contents as formatted table for debugging.

        Args:
            show_formulas: If True, show formulas instead of values

        Returns:
            Formatted string representation of the range

        Example:
            >>> print(ws.range("A1:C3").dump())
            |     A |     B |     C |
            |-------|-------|-------|
            |   100 |   200 |   300 |
            |    10 |    20 |    30 |
            |     1 |     2 |     3 |
        """
        from openpyxl.utils import get_column_letter

        # Collect data
        if show_formulas:
            data = []
            for row_idx in range(self._min_row, self._max_row + 1):
                row_data = []
                for col_idx in range(self._min_col, self._max_col + 1):
                    cell = self._ws._formula_ws.cell(row_idx, col_idx)
                    row_data.append(cell.value)
                data.append(row_data)
        else:
            data = self.values

        # Calculate column widths
        col_letters = [
            get_column_letter(c) for c in range(self._min_col, self._max_col + 1)
        ]
        col_widths = [len(letter) for letter in col_letters]

        for row in data:
            for i, val in enumerate(row):
                str_val = "" if val is None else str(val)
                col_widths[i] = max(col_widths[i], len(str_val))

        # Cap column width at 20 for readability
        col_widths = [min(w, 20) for w in col_widths]

        # Build output
        lines = []

        # Header row with column letters
        header = (
            "|"
            + "|".join(
                f" {col_letters[i]:>{col_widths[i]}} " for i in range(len(col_letters))
            )
            + "|"
        )
        lines.append(header)

        # Separator
        sep = (
            "|"
            + "|".join("-" * (col_widths[i] + 2) for i in range(len(col_widths)))
            + "|"
        )
        lines.append(sep)

        # Data rows
        for row in data:
            row_strs = []
            for i, val in enumerate(row):
                if val is None:
                    str_val = ""
                elif isinstance(val, float):
                    str_val = f"{val:.4g}"
                else:
                    str_val = str(val)
                # Truncate if too long
                if len(str_val) > col_widths[i]:
                    str_val = str_val[: col_widths[i] - 1] + "…"
                row_strs.append(f" {str_val:>{col_widths[i]}} ")
            lines.append("|" + "|".join(row_strs) + "|")

        result = "\n".join(lines)
        return result

    def apply_style(
        self,
        font: Font | None = None,
        fill: PatternFill | None = None,
        gradient_fill: GradientFill | None = None,
        alignment: Alignment | None = None,
        border: Border | None = None,
        protection: Protection | None = None,
        number_format: str | None = None,
    ) -> None:
        """Apply styles to all cells in the range.

        Args:
            font: Font style to apply
            fill: Fill/background style (PatternFill)
            gradient_fill: Gradient fill (takes precedence over fill)
            alignment: Alignment style
            border: Border style
            protection: Protection settings
            number_format: Number format string
        """
        # Mark dirty before applying styles
        if self._on_write:
            self._on_write()

        for row_idx in range(self._min_row, self._max_row + 1):
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws._formula_ws.cell(row_idx, col_idx)
                if font is not None:
                    cell.font = font
                if gradient_fill is not None:
                    cell.fill = gradient_fill
                elif fill is not None:
                    cell.fill = fill
                if alignment is not None:
                    cell.alignment = alignment
                if border is not None:
                    cell.border = border
                if protection is not None:
                    cell.protection = protection
                if number_format is not None:
                    cell.number_format = number_format

    def auto_financial_colors(self) -> None:
        """Apply conventional financial modeling colors based on cell content.

        Automatically sets font colors according to financial modeling conventions:
        - Blue (HARDCODE): Numeric input values
        - Black (FORMULA): Formulas without sheet references
        - Green (EXTERNAL_LINK): Formulas with sheet references (contain '!')
        - Text labels/strings are left unchanged

        Example:
            ws.range("A1:D10").auto_financial_colors()
        """

        if self._on_write:
            self._on_write()

        for row_idx in range(self._min_row, self._max_row + 1):
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws._formula_ws.cell(row_idx, col_idx)
                if cell.value is not None:
                    color = infer_financial_color(cell.value)
                    # Skip cells that don't need coloring (e.g., text labels)
                    if color is None:
                        continue
                    # Preserve existing font properties, just change color
                    old_font = cell.font
                    cell.font = Font(
                        name=old_font.name,
                        size=old_font.size,
                        bold=old_font.bold,
                        italic=old_font.italic,
                        underline=old_font.underline,
                        strike=old_font.strike,
                        color=color,
                    )

    def lint_financial_colors(self) -> list:
        """Check cells for financial color convention violations.

        Returns a list of ColorLintViolation for cells that don't follow conventions:
        - Blue (HARDCODE): Numeric input values
        - Black (FORMULA): Formulas without sheet references
        - Green (EXTERNAL_LINK): Formulas with sheet references
        - Text labels/strings are not checked

        Returns:
            List of ColorLintViolation objects for cells with wrong colors

        Example:
            violations = ws.range("A1:D10").lint_financial_colors()
            for v in violations:
                print(f"{v.cell}: expected {v.expected_color}, got {v.current_color}")
        """

        violations = []
        for row_idx in range(self._min_row, self._max_row + 1):
            for col_idx in range(self._min_col, self._max_col + 1):
                cell = self._ws._formula_ws.cell(row_idx, col_idx)
                if cell.value is not None:
                    expected = infer_financial_color(cell.value)
                    # Skip cells that don't need coloring (e.g., text labels)
                    if expected is None:
                        continue
                    current = None
                    if cell.font and cell.font.color:
                        # Get color as RGB string
                        color = cell.font.color
                        if color.type == "rgb" and color.rgb:
                            current = color.rgb
                        elif color.type == "theme":
                            # Theme colors are harder to resolve, treat as set
                            current = f"theme:{color.theme}"

                    if current != expected:
                        violations.append(
                            ColorLintViolation(
                                cell=cell.coordinate,
                                value=cell.value,
                                expected_color=expected,
                                current_color=current,
                            )
                        )
        return violations

    def __repr__(self) -> str:
        return f"<RangeProxy '{self._ws._formula_ws.title}'!{self._range_ref}>"


class CellProxy:
    """Proxy for openpyxl Cell that returns materialized values after sync.

    Delegates to the formula cell for writes and attributes,
    but returns values from the data-only cell when available.
    """

    def __init__(
        self,
        formula_cell: Cell,
        values_cell: Cell | None = None,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._formula_cell = formula_cell
        self._values_cell = values_cell
        self._on_write = on_write

    def _update_values_cell(self, values_cell: Cell | None) -> None:
        """Update the values cell reference (called after sync)."""
        self._values_cell = values_cell

    @property
    def value(self) -> Any:
        """Get materialized value if available, otherwise formula value."""
        if self._values_cell is not None:
            return self._values_cell.value
        return self._formula_cell.value

    @value.setter
    def value(self, val: Any) -> None:
        """Set value on formula cell."""
        if self._on_write:
            self._on_write()
        self._formula_cell.value = val

    @property
    def formula(self) -> str | None:
        """Get formula string if cell contains a formula, otherwise None.

        Example:
            >>> ws['A1'] = '=SUM(B1:B10)'
            >>> ws['A1'].formula
            '=SUM(B1:B10)'
            >>> ws['A2'] = 42
            >>> ws['A2'].formula
            None
        """
        val = self._formula_cell.value
        if isinstance(val, str) and val.startswith("="):
            return val
        return None

    # Forward all other attributes to formula cell
    def __getattr__(self, name: str) -> Any:
        return getattr(self._formula_cell, name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in ("_formula_cell", "_values_cell", "_on_write"):
            object.__setattr__(self, name, value)
        else:
            # Mark dirty for any attribute write (font, number_format, etc.)
            on_write = object.__getattribute__(self, "_on_write")
            if on_write:
                on_write()
            setattr(self._formula_cell, name, value)

    def __repr__(self) -> str:
        return f"<CellProxy {self._formula_cell.coordinate}>"


class WorksheetProxy:
    """Proxy for openpyxl Worksheet that returns CellProxy objects.

    Wraps cell access to return CellProxy instances that show
    materialized values after sync.
    """

    def __init__(
        self,
        formula_ws: Worksheet,
        values_ws: Worksheet | None = None,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._formula_ws = formula_ws
        self._values_ws = values_ws
        self._on_write = on_write
        self._cell_cache: dict[str, CellProxy] = {}  # Cache cell proxies

    def _update_values_ws(self, values_ws: Worksheet | None) -> None:
        """Update the values worksheet and refresh all cached cell proxies."""
        self._values_ws = values_ws
        # Update all cached cell proxies
        for coord, cell_proxy in self._cell_cache.items():
            values_cell = values_ws[coord] if values_ws else None
            cell_proxy._update_values_cell(values_cell)

    def __getitem__(self, key: str) -> CellProxy | tuple[tuple[CellProxy, ...], ...]:
        """Get cell or range, wrapped in CellProxy."""
        formula_item = self._formula_ws[key]

        # Single cell access like ws["A1"]
        if isinstance(formula_item, Cell):
            # Return cached proxy if it exists
            if key in self._cell_cache:
                return self._cell_cache[key]

            values_cell = None
            if self._values_ws is not None:
                values_cell = self._values_ws[key]
            proxy = CellProxy(formula_item, values_cell, self._on_write)
            self._cell_cache[key] = proxy
            return proxy

        # Range access like ws["A1:B2"] - returns tuple of tuples of cells
        if isinstance(formula_item, tuple):
            if self._values_ws is not None:
                values_item = self._values_ws[key]
                return tuple(
                    tuple(
                        CellProxy(f_cell, v_cell, self._on_write)
                        for f_cell, v_cell in zip(f_row, v_row)
                    )
                    for f_row, v_row in zip(formula_item, values_item)
                )
            else:
                return tuple(
                    tuple(CellProxy(cell, None, self._on_write) for cell in row)
                    for row in formula_item
                )

        return formula_item

    def __setitem__(self, key: str, value: Any) -> None:
        """Set cell value."""
        if self._on_write:
            self._on_write()
        self._formula_ws[key] = value

    def cell(self, row: int, column: int, value: Any = None) -> CellProxy:
        """Access cell by row/column index."""
        if value is not None and self._on_write:
            self._on_write()
        formula_cell = self._formula_ws.cell(row, column, value)
        values_cell = None
        if self._values_ws is not None:
            values_cell = self._values_ws.cell(row, column)
        return CellProxy(formula_cell, values_cell, self._on_write)

    def iter_rows(
        self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False
    ):
        """Iterate over rows, yielding CellProxy objects."""
        formula_rows = self._formula_ws.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=values_only,
        )

        if values_only:
            # If values_only, openpyxl returns raw values, not cells
            yield from formula_rows
            return

        if self._values_ws is not None:
            values_rows = self._values_ws.iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
            )
            for f_row, v_row in zip(formula_rows, values_rows):
                yield tuple(
                    CellProxy(f_cell, v_cell, self._on_write)
                    for f_cell, v_cell in zip(f_row, v_row)
                )
        else:
            for f_row in formula_rows:
                yield tuple(CellProxy(cell, None, self._on_write) for cell in f_row)

    def iter_cols(
        self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False
    ):
        """Iterate over columns, yielding CellProxy objects."""
        formula_cols = self._formula_ws.iter_cols(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=values_only,
        )

        if values_only:
            yield from formula_cols
            return

        if self._values_ws is not None:
            values_cols = self._values_ws.iter_cols(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
            )
            for f_col, v_col in zip(formula_cols, values_cols):
                yield tuple(
                    CellProxy(f_cell, v_cell, self._on_write)
                    for f_cell, v_cell in zip(f_col, v_col)
                )
        else:
            for f_col in formula_cols:
                yield tuple(CellProxy(cell, None, self._on_write) for cell in f_col)

    def range(self, ref: str) -> RangeProxy:
        """Get a range for bulk read/write operations.

        Args:
            ref: Range reference like "A1:C3" or single cell "A1"

        Returns:
            RangeProxy for bulk operations

        Example:
            r = sheet.range("A1:C3")
            r.values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
            r.apply_style(font=Font(bold=True))
        """
        return RangeProxy(self, ref, self._on_write)

    @property
    def formulas(self) -> dict[str, str]:
        """Get all formulas in this sheet as {coordinate: formula}.

        Example:
            >>> sheet.formulas
            {'A1': '=B1+C1', 'D5': '=SUM(A1:A4)'}
        """
        result: dict[str, str] = {}
        for row in self._formula_ws.iter_rows():
            for cell in row:
                if (
                    cell.value is not None
                    and isinstance(cell.value, str)
                    and cell.value.startswith("=")
                ):
                    result[cell.coordinate] = cell.value
        return result

    @property
    def formula_count(self) -> int:
        """Count of formulas in this sheet.

        Example:
            >>> sheet.formula_count
            42
        """
        return len(self.formulas)

    def find_errors(self) -> ErrorScanResult:
        """Find formula errors in this sheet.

        Scans the values worksheet (materialized values after sync) for
        Excel error values like #DIV/0!, #REF!, #NAME?, etc.

        Returns:
            ErrorScanResult with errors organized by type and total count.
            Has a nice __repr__ with truncation for agent-friendly output.

        Example:
            >>> result = sheet.find_errors()
            >>> print(result)
            Formula errors (3):
              #DIV/0!: A3, B5
              #REF!: C10
        """
        if self._values_ws is None:
            return ErrorScanResult()

        error_details: dict[str, list[str]] = {err: [] for err in EXCEL_ERRORS}

        for row in self._values_ws.iter_rows():
            for cell in row:
                if cell.value is not None and isinstance(cell.value, str):
                    for err in EXCEL_ERRORS:
                        if err in cell.value:
                            error_details[err].append(cell.coordinate)
                            break

        # Filter to only include error types that have occurrences
        errors_by_type = {k: v for k, v in error_details.items() if v}
        total = sum(len(locs) for locs in errors_by_type.values())
        return ErrorScanResult(errors_by_type=errors_by_type, total_errors=total)

    def auto_financial_colors(self) -> None:
        """Apply conventional financial modeling colors to all cells in sheet.

        See RangeProxy.auto_financial_colors() for details.
        """

        if self._on_write:
            self._on_write()

        for row in self._formula_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    color = infer_financial_color(cell.value)
                    # Skip cells that don't need coloring (e.g., text labels)
                    if color is None:
                        continue
                    old_font = cell.font
                    cell.font = Font(
                        name=old_font.name,
                        size=old_font.size,
                        bold=old_font.bold,
                        italic=old_font.italic,
                        underline=old_font.underline,
                        strike=old_font.strike,
                        color=color,
                    )

    def lint_financial_colors(self) -> list:
        """Check all cells in sheet for financial color convention violations.

        See RangeProxy.lint_financial_colors() for details.

        Returns:
            List of ColorLintViolation objects
        """

        violations = []
        for row in self._formula_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    expected = infer_financial_color(cell.value)
                    # Skip cells that don't need coloring (e.g., text labels)
                    if expected is None:
                        continue
                    current = None
                    if cell.font and cell.font.color:
                        color = cell.font.color
                        if color.type == "rgb" and color.rgb:
                            current = color.rgb
                        elif color.type == "theme":
                            current = f"theme:{color.theme}"

                    if current != expected:
                        violations.append(
                            ColorLintViolation(
                                cell=cell.coordinate,
                                value=cell.value,
                                expected_color=expected,
                                current_color=current,
                            )
                        )
        return violations

    # Forward all other attributes to formula worksheet
    def __getattr__(self, name: str) -> Any:
        return getattr(self._formula_ws, name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in ("_formula_ws", "_values_ws", "_on_write", "_cell_cache"):
            object.__setattr__(self, name, value)
        else:
            # Mark dirty for worksheet attribute writes (like title)
            on_write = object.__getattribute__(self, "_on_write")
            if on_write:
                on_write()
            setattr(self._formula_ws, name, value)

    def __repr__(self) -> str:
        return f"<WorksheetProxy '{self._formula_ws.title}'>"


class WorkbookProxy:
    """Proxy for openpyxl Workbook that returns WorksheetProxy objects.

    Wraps worksheet access to return WorksheetProxy instances.
    """

    def __init__(
        self,
        formula_wb: Workbook,
        values_wb: Workbook | None = None,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._formula_wb = formula_wb
        self._values_wb = values_wb
        self._on_write = on_write
        self._sheet_cache: dict[str, WorksheetProxy] = {}  # Cache sheet proxies

    def _update_values_wb(self, values_wb: Workbook | None) -> None:
        """Update the values workbook and refresh all cached sheet proxies."""
        self._values_wb = values_wb
        # Update all cached sheet proxies with new values worksheet
        for sheet_name, sheet_proxy in self._sheet_cache.items():
            values_ws = None
            if values_wb is not None and sheet_name in values_wb.sheetnames:
                values_ws = values_wb[sheet_name]
            sheet_proxy._update_values_ws(values_ws)

    def __getitem__(self, key: str) -> WorksheetProxy:
        """Get worksheet by name."""
        # Return cached proxy if it exists
        if key in self._sheet_cache:
            return self._sheet_cache[key]

        formula_ws = self._formula_wb[key]
        values_ws = None
        if self._values_wb is not None:
            values_ws = self._values_wb[key]
        proxy = WorksheetProxy(formula_ws, values_ws, self._on_write)
        self._sheet_cache[key] = proxy
        return proxy

    @property
    def active(self) -> WorksheetProxy:
        """Get active worksheet."""
        formula_ws = self._formula_wb.active
        if formula_ws is None:
            raise RuntimeError("No active worksheet")

        # Use sheet name for caching to maintain identity
        sheet_name = formula_ws.title
        if sheet_name in self._sheet_cache:
            return self._sheet_cache[sheet_name]

        values_ws = None
        if self._values_wb is not None:
            values_ws = self._values_wb.active
        proxy = WorksheetProxy(formula_ws, values_ws, self._on_write)
        self._sheet_cache[sheet_name] = proxy
        return proxy

    @active.setter
    def active(self, sheet: WorksheetProxy | str) -> None:
        """Set active worksheet by proxy or name."""
        if self._on_write:
            self._on_write()
        if isinstance(sheet, str):
            self._formula_wb.active = self._formula_wb[sheet]
            if self._values_wb is not None:
                self._values_wb.active = self._values_wb[sheet]
        elif isinstance(sheet, WorksheetProxy):
            self._formula_wb.active = sheet._formula_ws
            if self._values_wb is not None and sheet._values_ws is not None:
                self._values_wb.active = sheet._values_ws
        else:
            raise TypeError(f"Expected WorksheetProxy or str, got {type(sheet)}")

    def create_sheet(
        self, title: str | None = None, index: int | None = None
    ) -> WorksheetProxy:
        """Create a new worksheet."""
        if self._on_write:
            self._on_write()
        formula_ws = self._formula_wb.create_sheet(title, index)
        # New sheets don't have values yet, but cache so sync() updates them
        proxy = WorksheetProxy(formula_ws, None, self._on_write)
        if formula_ws.title:
            self._sheet_cache[formula_ws.title] = proxy
        return proxy

    @property
    def worksheets(self) -> list[WorksheetProxy]:
        """Get list of all worksheets."""
        result = []
        for f_ws in self._formula_wb.worksheets:
            values_ws = None
            if self._values_wb is not None and f_ws.title in self._values_wb.sheetnames:
                values_ws = self._values_wb[f_ws.title]
            result.append(WorksheetProxy(f_ws, values_ws, self._on_write))
        return result

    # Forward all other attributes to formula workbook
    def __getattr__(self, name: str) -> Any:
        return getattr(self._formula_wb, name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in ("_formula_wb", "_values_wb", "_sheet_cache", "_on_write"):
            object.__setattr__(self, name, value)
        elif name == "active":
            # Use our custom setter for active
            type(self).active.fset(self, value)
        else:
            # Mark dirty for any workbook attribute write
            on_write = object.__getattribute__(self, "_on_write")
            if on_write:
                on_write()
            setattr(self._formula_wb, name, value)

    def __repr__(self) -> str:
        return f"<WorkbookProxy with {len(self._formula_wb.worksheets)} sheets>"
