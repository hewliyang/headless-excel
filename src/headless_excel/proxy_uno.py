"""UNO-backed proxy wrappers that communicate with LibreOffice daemon.

These proxies implement the same interface as the openpyxl-based proxies
but communicate with an open document in LibreOffice for faster operations.
"""

from __future__ import annotations

import json
import re
import sys
from collections.abc import Callable
from typing import Any, Literal

from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Protection,
)
from openpyxl.utils import get_column_letter

from headless_excel.daemon.base import send_daemon_command
from headless_excel.errors import ErrorScanResult, RecalcError

# Type alias for write callback
OnWriteCallback = Callable[[], None] | None

# Excel error values to detect
EXCEL_ERRORS = (
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
)

# Pattern for parsing range references
RANGE_PATTERN = re.compile(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", re.IGNORECASE)


def _col_index_from_string(col: str) -> int:
    """Convert column letters to 1-indexed column number."""
    result = 0
    for c in col.upper():
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result


def _parse_range(ref: str) -> tuple[int, int, int, int]:
    """Parse range reference into (min_row, min_col, max_row, max_col), 1-indexed."""
    match = RANGE_PATTERN.match(ref.upper())
    if not match:
        raise ValueError(f"Invalid range reference: {ref!r}")

    start_col, start_row, end_col, end_row = match.groups()
    min_col = _col_index_from_string(start_col)
    min_row = int(start_row)

    if end_col and end_row:
        max_col = _col_index_from_string(end_col)
        max_row = int(end_row)
    else:
        max_col = min_col
        max_row = min_row

    if min_row > max_row:
        min_row, max_row = max_row, min_row
    if min_col > max_col:
        min_col, max_col = max_col, min_col

    return min_row, min_col, max_row, max_col


class RangeProxyUNO:
    """UNO-backed proxy for a cell range supporting bulk read/write operations."""

    def __init__(
        self,
        ws_proxy: WorksheetProxyUNO,
        range_ref: str,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._ws = ws_proxy
        self._range_ref = range_ref
        self._on_write = on_write
        self._min_row, self._min_col, self._max_row, self._max_col = _parse_range(
            range_ref
        )

    @property
    def num_rows(self) -> int:
        return self._max_row - self._min_row + 1

    @property
    def num_cols(self) -> int:
        return self._max_col - self._min_col + 1

    @property
    def shape(self) -> tuple[int, int]:
        return (self.num_rows, self.num_cols)

    @property
    def values(self) -> list[list[Any]]:
        """Get 2D array of values from the range via UNO."""
        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        start_cell = f"{get_column_letter(self._min_col)}{self._min_row}"
        end_cell = f"{get_column_letter(self._max_col)}{self._max_row}"
        range_ref = f"{sheet_name}!{start_cell}:{end_cell}"

        response = send_daemon_command(f"GET_RANGE:{session_id}:{range_ref}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("DATA:"):
            return json.loads(response[5:])
        raise RecalcError(f"Unexpected response: {response}")

    @values.setter
    def values(self, data: list[list[Any]]) -> None:
        """Set 2D array of values to the range via UNO."""
        if not isinstance(data, list):
            raise ValueError("Data must be a 2D list")

        if len(data) == 0:
            print(
                "[headless-excel] .values no data to write (empty list)",
                file=sys.stderr,
            )
            return

        max_cols = 0
        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ValueError(f"Row {i} must be a list")
            max_cols = max(max_cols, len(row))

        actual_rows = len(data)
        actual_cols = max_cols

        actual_max_row = self._min_row + actual_rows - 1
        actual_max_col = self._min_col + actual_cols - 1

        start_cell = f"{get_column_letter(self._min_col)}{self._min_row}"
        end_cell = f"{get_column_letter(actual_max_col)}{actual_max_row}"
        if start_cell == end_cell:
            actual_range = start_cell
        else:
            actual_range = f"{start_cell}:{end_cell}"

        specified_matches = (
            actual_rows == self.num_rows and actual_cols == self.num_cols
        )

        if self._on_write:
            self._on_write()

        # Send to daemon
        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        range_ref = f"{sheet_name}!{start_cell}:{end_cell}"
        data_json = json.dumps(data)

        response = send_daemon_command(
            f"SET_RANGE:{session_id}:{range_ref}:{data_json}"
        )
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

        if specified_matches:
            print(
                f"[headless-excel] .values wrote to {actual_range} ({actual_rows}×{actual_cols})",
                file=sys.stderr,
            )
        else:
            print(
                f"[headless-excel] .values wrote to {actual_range} (specified {self._range_ref}, got {actual_rows}×{actual_cols})",
                file=sys.stderr,
            )

    @property
    def formulas(self) -> dict[str, str]:
        """Get formulas in range as {coordinate: formula}."""
        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        start_cell = f"{get_column_letter(self._min_col)}{self._min_row}"
        end_cell = f"{get_column_letter(self._max_col)}{self._max_row}"
        range_ref = f"{sheet_name}!{start_cell}:{end_cell}"

        response = send_daemon_command(f"GET_FORMULAS:{session_id}:{range_ref}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("FORMULAS:"):
            formula_data = json.loads(response[9:])
            result: dict[str, str] = {}
            for row_idx, row in enumerate(formula_data):
                for col_idx, formula in enumerate(row):
                    if formula and isinstance(formula, str) and formula.startswith("="):
                        cell_ref = f"{get_column_letter(self._min_col + col_idx)}{self._min_row + row_idx}"
                        result[cell_ref] = formula
            return result
        raise RecalcError(f"Unexpected response: {response}")

    @property
    def formula_count(self) -> int:
        return len(self.formulas)

    def find_errors(self) -> ErrorScanResult:
        """Find formula errors in this range."""
        values = self.values
        error_details: dict[str, list[str]] = {err: [] for err in EXCEL_ERRORS}

        for row_idx, row in enumerate(values):
            for col_idx, value in enumerate(row):
                if value is not None and isinstance(value, str):
                    for err in EXCEL_ERRORS:
                        if err in value:
                            cell_ref = f"{get_column_letter(self._min_col + col_idx)}{self._min_row + row_idx}"
                            error_details[err].append(cell_ref)
                            break

        errors_by_type = {k: v for k, v in error_details.items() if v}
        total = sum(len(locs) for locs in errors_by_type.values())
        return ErrorScanResult(errors_by_type=errors_by_type, total_errors=total)

    def clear(self, styles: bool = True) -> None:
        """Clear cell values in the range."""
        if self._on_write:
            self._on_write()

        # Create empty data array
        empty_data = [[None] * self.num_cols for _ in range(self.num_rows)]
        self.values = empty_data

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
        """Apply styles - NOTE: styles are not fully supported in UNO mode."""
        # Mark dirty
        if self._on_write:
            self._on_write()
        # Styles are handled when saving via openpyxl on final save
        # For now, this is a no-op in UNO mode
        pass

    def auto_fill(
        self,
        direction: Literal["down", "right", "up", "left"] | None = None,
        source_rows: int = 1,
        source_cols: int = 1,
        copy_styles: bool = True,
    ) -> None:
        """Fill the range by repeating source cells."""
        if self.num_rows == 1 and self.num_cols == 1:
            raise ValueError("Cannot auto_fill a single cell range")

        if direction is None:
            direction = "down" if self.num_rows >= self.num_cols else "right"

        if direction in ("down", "up") and source_rows >= self.num_rows:
            raise ValueError(
                f"source_rows ({source_rows}) must be less than range rows ({self.num_rows})"
            )
        if direction in ("left", "right") and source_cols >= self.num_cols:
            raise ValueError(
                f"source_cols ({source_cols}) must be less than range cols ({self.num_cols})"
            )

        if self._on_write:
            self._on_write()

        # Get current formulas to translate
        formulas = self.formulas
        values = self.values

        if direction == "down":
            self._fill_down(values, formulas, source_rows)
        elif direction == "right":
            self._fill_right(values, formulas, source_cols)
        elif direction == "up":
            self._fill_up(values, formulas, source_rows)
        elif direction == "left":
            self._fill_left(values, formulas, source_cols)

    def _translate_formula(self, formula: str, row_delta: int, col_delta: int) -> str:
        """Translate formula references by row/col delta."""

        # Simple regex-based translation
        def replace_ref(match):
            prefix = match.group(1) or ""
            col_abs = match.group(2) or ""
            col = match.group(3)
            row_abs = match.group(4) or ""
            row = int(match.group(5))

            new_col = col
            new_row = row

            if not col_abs:
                col_idx = _col_index_from_string(col)
                new_col_idx = col_idx + col_delta
                if new_col_idx > 0:
                    new_col = get_column_letter(new_col_idx)

            if not row_abs:
                new_row = row + row_delta

            return f"{prefix}{col_abs}{new_col}{row_abs}{new_row}"

        pattern = r"(\w+!)?\$?([A-Z]+)\$?(\d+)"
        pattern = r"([A-Za-z_][A-Za-z0-9_]*!)?\$?([A-Z]{1,3})\$?([1-9][0-9]*)"
        pattern = r"([A-Za-z_][A-Za-z0-9_]*!)?(\$)?([A-Z]{1,3})(\$)?([1-9][0-9]*)"
        return re.sub(pattern, replace_ref, formula, flags=re.IGNORECASE)

    def _fill_down(self, values, formulas, source_rows):
        new_data = [row[:] for row in values]
        for target_row_idx in range(source_rows, self.num_rows):
            source_offset = (target_row_idx - source_rows) % source_rows
            source_row_idx = source_offset
            row_delta = target_row_idx - source_row_idx

            for col_idx in range(self.num_cols):
                cell_ref = f"{get_column_letter(self._min_col + col_idx)}{self._min_row + source_row_idx}"
                if cell_ref in formulas:
                    new_data[target_row_idx][col_idx] = self._translate_formula(
                        formulas[cell_ref], row_delta, 0
                    )
                else:
                    new_data[target_row_idx][col_idx] = values[source_row_idx][col_idx]

        self.values = new_data

    def _fill_right(self, values, formulas, source_cols):
        new_data = [row[:] for row in values]
        for target_col_idx in range(source_cols, self.num_cols):
            source_offset = (target_col_idx - source_cols) % source_cols
            source_col_idx = source_offset
            col_delta = target_col_idx - source_col_idx

            for row_idx in range(self.num_rows):
                cell_ref = f"{get_column_letter(self._min_col + source_col_idx)}{self._min_row + row_idx}"
                if cell_ref in formulas:
                    new_data[row_idx][target_col_idx] = self._translate_formula(
                        formulas[cell_ref], 0, col_delta
                    )
                else:
                    new_data[row_idx][target_col_idx] = values[row_idx][source_col_idx]

        self.values = new_data

    def _fill_up(self, values, formulas, source_rows):
        new_data = [row[:] for row in values]
        target_end = self.num_rows - source_rows - 1

        for target_row_idx in range(target_end, -1, -1):
            source_offset = (target_end - target_row_idx) % source_rows
            source_row_idx = self.num_rows - 1 - source_offset
            row_delta = target_row_idx - source_row_idx

            for col_idx in range(self.num_cols):
                cell_ref = f"{get_column_letter(self._min_col + col_idx)}{self._min_row + source_row_idx}"
                if cell_ref in formulas:
                    new_data[target_row_idx][col_idx] = self._translate_formula(
                        formulas[cell_ref], row_delta, 0
                    )
                else:
                    new_data[target_row_idx][col_idx] = values[source_row_idx][col_idx]

        self.values = new_data

    def _fill_left(self, values, formulas, source_cols):
        new_data = [row[:] for row in values]
        target_end = self.num_cols - source_cols - 1

        for target_col_idx in range(target_end, -1, -1):
            source_offset = (target_end - target_col_idx) % source_cols
            source_col_idx = self.num_cols - 1 - source_offset
            col_delta = target_col_idx - source_col_idx

            for row_idx in range(self.num_rows):
                cell_ref = f"{get_column_letter(self._min_col + source_col_idx)}{self._min_row + row_idx}"
                if cell_ref in formulas:
                    new_data[row_idx][target_col_idx] = self._translate_formula(
                        formulas[cell_ref], 0, col_delta
                    )
                else:
                    new_data[row_idx][target_col_idx] = values[row_idx][source_col_idx]

        self.values = new_data

    def dump(self, show_formulas: bool = False) -> str:
        """Print range contents as formatted table."""
        if show_formulas:
            # Get formulas and values
            formulas = self.formulas
            values = self.values
            data = []
            for row_idx, row in enumerate(values):
                row_data = []
                for col_idx, val in enumerate(row):
                    cell_ref = f"{get_column_letter(self._min_col + col_idx)}{self._min_row + row_idx}"
                    if cell_ref in formulas:
                        row_data.append(formulas[cell_ref])
                    else:
                        row_data.append(val)
                data.append(row_data)
        else:
            data = self.values

        col_letters = [
            get_column_letter(c) for c in range(self._min_col, self._max_col + 1)
        ]
        col_widths = [len(letter) for letter in col_letters]

        for row in data:
            for i, val in enumerate(row):
                str_val = "" if val is None else str(val)
                col_widths[i] = max(col_widths[i], len(str_val))

        col_widths = [min(w, 20) for w in col_widths]

        lines = []
        header = (
            "|"
            + "|".join(
                f" {col_letters[i]:>{col_widths[i]}} " for i in range(len(col_letters))
            )
            + "|"
        )
        lines.append(header)

        sep = (
            "|"
            + "|".join("-" * (col_widths[i] + 2) for i in range(len(col_widths)))
            + "|"
        )
        lines.append(sep)

        for row in data:
            row_strs = []
            for i, val in enumerate(row):
                if val is None:
                    str_val = ""
                elif isinstance(val, float):
                    str_val = f"{val:.4g}"
                else:
                    str_val = str(val)
                if len(str_val) > col_widths[i]:
                    str_val = str_val[: col_widths[i] - 1] + "…"
                row_strs.append(f" {str_val:>{col_widths[i]}} ")
            lines.append("|" + "|".join(row_strs) + "|")

        return "\n".join(lines)

    def __repr__(self) -> str:
        return f"<RangeProxyUNO '{self._ws._name}'!{self._range_ref}>"


class CellProxyUNO:
    """UNO-backed proxy for a single cell."""

    def __init__(
        self,
        ws_proxy: WorksheetProxyUNO,
        row: int,
        col: int,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._ws = ws_proxy
        self._row = row  # 1-indexed
        self._col = col  # 1-indexed
        self._on_write = on_write
        self._coordinate = f"{get_column_letter(col)}{row}"

    @property
    def coordinate(self) -> str:
        return self._coordinate

    @property
    def value(self) -> Any:
        """Get cell value from UNO."""
        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        cell_ref = f"{sheet_name}!{self._coordinate}"

        response = send_daemon_command(f"GET:{session_id}:{cell_ref}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("VALUE:"):
            data = json.loads(response[6:])
            return data.get("value")
        raise RecalcError(f"Unexpected response: {response}")

    @value.setter
    def value(self, val: Any) -> None:
        """Set cell value via UNO."""
        if self._on_write:
            self._on_write()

        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        cell_ref = f"{sheet_name}!{self._coordinate}"

        # Encode value as JSON for proper type handling
        value_json = json.dumps(val)
        response = send_daemon_command(f"SET:{session_id}:{cell_ref}={value_json}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

    @property
    def formula(self) -> str | None:
        """Get formula if cell contains one."""
        session_id = self._ws._wb._session_id
        sheet_name = self._ws._name
        cell_ref = f"{sheet_name}!{self._coordinate}"

        response = send_daemon_command(f"GET:{session_id}:{cell_ref}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("VALUE:"):
            data = json.loads(response[6:])
            return data.get("formula")
        return None

    # Style properties - stored locally until save
    # Note: In UNO mode, styles are not fully supported
    @property
    def font(self) -> Font:
        return Font()

    @font.setter
    def font(self, value: Font) -> None:
        if self._on_write:
            self._on_write()

    @property
    def fill(self) -> PatternFill:
        return PatternFill()

    @fill.setter
    def fill(self, value: PatternFill | GradientFill) -> None:
        if self._on_write:
            self._on_write()

    @property
    def border(self) -> Border:
        return Border()

    @border.setter
    def border(self, value: Border) -> None:
        if self._on_write:
            self._on_write()

    @property
    def number_format(self) -> str:
        return "General"

    @number_format.setter
    def number_format(self, value: str) -> None:
        if self._on_write:
            self._on_write()

    @property
    def alignment(self) -> Alignment:
        return Alignment()

    @alignment.setter
    def alignment(self, value: Alignment) -> None:
        if self._on_write:
            self._on_write()

    @property
    def protection(self) -> Protection:
        return Protection()

    @protection.setter
    def protection(self, value: Protection) -> None:
        if self._on_write:
            self._on_write()

    def __repr__(self) -> str:
        return f"<CellProxyUNO {self._coordinate}>"


class WorksheetProxyUNO:
    """UNO-backed proxy for a worksheet."""

    def __init__(
        self,
        wb_proxy: WorkbookProxyUNO,
        name: str,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._wb = wb_proxy
        self._name = name
        self._on_write = on_write
        self._cell_cache: dict[str, CellProxyUNO] = {}

    @property
    def title(self) -> str:
        return self._name

    @title.setter
    def title(self, value: str) -> None:
        if self._on_write:
            self._on_write()
        session_id = self._wb._session_id
        response = send_daemon_command(
            f"RENAME_SHEET:{session_id}:{self._name}:{value}"
        )
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        self._name = value

    def __getitem__(
        self, key: str
    ) -> CellProxyUNO | tuple[tuple[CellProxyUNO, ...], ...]:
        """Get cell or range."""
        if ":" in key:
            # Range access
            min_row, min_col, max_row, max_col = _parse_range(key)
            rows = []
            for row in range(min_row, max_row + 1):
                row_cells = []
                for col in range(min_col, max_col + 1):
                    row_cells.append(self.cell(row, col))
                rows.append(tuple(row_cells))
            return tuple(rows)
        else:
            # Single cell
            if key in self._cell_cache:
                return self._cell_cache[key]
            min_row, min_col, _, _ = _parse_range(key)
            proxy = CellProxyUNO(self, min_row, min_col, self._on_write)
            self._cell_cache[key] = proxy
            return proxy

    def __setitem__(self, key: str, value: Any) -> None:
        """Set cell value."""
        if self._on_write:
            self._on_write()
        cell = self[key]
        if isinstance(cell, CellProxyUNO):
            cell.value = value
        else:
            raise ValueError("Cannot set value on a range")

    def get_cell(self, ref: str) -> CellProxyUNO:
        """Get a single cell by reference."""
        if ":" in ref:
            raise ValueError(f"get_cell() only accepts single cell refs: {ref}")
        result = self[ref]
        return result  # type: ignore

    def cell(self, row: int, column: int, value: Any = None) -> CellProxyUNO:
        """Access cell by row/column index (1-indexed)."""
        coord = f"{get_column_letter(column)}{row}"
        if coord in self._cell_cache:
            proxy = self._cell_cache[coord]
        else:
            proxy = CellProxyUNO(self, row, column, self._on_write)
            self._cell_cache[coord] = proxy

        if value is not None:
            proxy.value = value
        return proxy

    def range(self, ref: str) -> RangeProxyUNO:
        """Get a range for bulk operations."""
        return RangeProxyUNO(self, ref, self._on_write)

    def write(self, cell: str, data: list[list[Any]]) -> str:
        """Write 2D data starting at anchor cell."""
        if ":" in cell:
            raise ValueError(f"write() takes a single cell anchor: {cell}")

        if not isinstance(data, list):
            raise ValueError("Data must be a 2D list")

        if len(data) == 0:
            print(
                "[headless-excel] .write() no data to write (empty list)",
                file=sys.stderr,
            )
            return cell

        max_cols = 0
        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ValueError(f"Row {i} must be a list")
            max_cols = max(max_cols, len(row))

        if max_cols == 0:
            print(
                "[headless-excel] .write() no data to write (empty rows)",
                file=sys.stderr,
            )
            return cell

        min_row, min_col, _, _ = _parse_range(cell)
        actual_rows = len(data)
        actual_cols = max_cols

        actual_max_row = min_row + actual_rows - 1
        actual_max_col = min_col + actual_cols - 1

        start_cell = f"{get_column_letter(min_col)}{min_row}"
        end_cell = f"{get_column_letter(actual_max_col)}{actual_max_row}"
        if start_cell == end_cell:
            actual_range = start_cell
        else:
            actual_range = f"{start_cell}:{end_cell}"

        if self._on_write:
            self._on_write()

        # Bypass the range proxy feedback print since we do our own
        session_id = self._wb._session_id
        range_ref = f"{self._name}!{start_cell}:{end_cell}"
        data_json = json.dumps(data)
        response = send_daemon_command(
            f"SET_RANGE:{session_id}:{range_ref}:{data_json}"
        )
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

        print(
            f"[headless-excel] .write() wrote to {actual_range} ({actual_rows}×{actual_cols})",
            file=sys.stderr,
        )
        return actual_range

    @property
    def formulas(self) -> dict[str, str]:
        """Get all formulas in the sheet."""
        # Get used range and extract formulas
        # This is expensive, so we'll limit to a reasonable area
        # For now, scan A1:ZZ1000
        session_id = self._wb._session_id
        range_ref = f"{self._name}!A1:ZZ1000"

        response = send_daemon_command(f"GET_FORMULAS:{session_id}:{range_ref}")
        if response.startswith("ERROR:"):
            # If range is too large or invalid, return empty
            return {}
        if response.startswith("FORMULAS:"):
            formula_data = json.loads(response[9:])
            result: dict[str, str] = {}
            for row_idx, row in enumerate(formula_data):
                for col_idx, formula in enumerate(row):
                    if formula and isinstance(formula, str) and formula.startswith("="):
                        cell_ref = f"{get_column_letter(col_idx + 1)}{row_idx + 1}"
                        result[cell_ref] = formula
            return result
        return {}

    @property
    def formula_count(self) -> int:
        return len(self.formulas)

    def find_errors(self) -> ErrorScanResult:
        """Find formula errors in the sheet."""
        # Scan A1:ZZ1000 for errors
        try:
            range_proxy = self.range("A1:ZZ1000")
            return range_proxy.find_errors()
        except Exception:
            return ErrorScanResult()

    def iter_rows(
        self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False
    ):
        """Iterate over rows."""
        min_row = min_row or 1
        max_row = max_row or 100  # Default limit
        min_col = min_col or 1
        max_col = max_col or 26  # Default to column Z

        if values_only:
            range_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
            range_proxy = self.range(range_ref)
            for row in range_proxy.values:
                yield tuple(row)
        else:
            for row in range(min_row, max_row + 1):
                yield tuple(self.cell(row, col) for col in range(min_col, max_col + 1))

    def iter_cols(
        self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False
    ):
        """Iterate over columns."""
        min_row = min_row or 1
        max_row = max_row or 100
        min_col = min_col or 1
        max_col = max_col or 26

        if values_only:
            range_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
            range_proxy = self.range(range_ref)
            values = range_proxy.values
            for col in range(max_col - min_col + 1):
                yield tuple(row[col] if col < len(row) else None for row in values)
        else:
            for col in range(min_col, max_col + 1):
                yield tuple(self.cell(row, col) for row in range(min_row, max_row + 1))

    def __repr__(self) -> str:
        return f"<WorksheetProxyUNO '{self._name}'>"


class WorkbookProxyUNO:
    """UNO-backed proxy for a workbook with an open session."""

    def __init__(
        self,
        session_id: str,
        filepath: str,
        on_write: OnWriteCallback = None,
    ) -> None:
        self._session_id = session_id
        self._filepath = filepath
        self._on_write = on_write
        self._sheet_cache: dict[str, WorksheetProxyUNO] = {}
        self._active_sheet: str | None = None

    @property
    def sheetnames(self) -> list[str]:
        """Get list of sheet names."""
        response = send_daemon_command(f"SHEETS:{self._session_id}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("SHEETS:"):
            return json.loads(response[7:])
        raise RecalcError(f"Unexpected response: {response}")

    def __getitem__(self, key: str) -> WorksheetProxyUNO:
        """Get worksheet by name."""
        if key in self._sheet_cache:
            return self._sheet_cache[key]

        # Verify sheet exists
        if key not in self.sheetnames:
            raise KeyError(f"Sheet '{key}' not found")

        proxy = WorksheetProxyUNO(self, key, self._on_write)
        self._sheet_cache[key] = proxy
        return proxy

    @property
    def active(self) -> WorksheetProxyUNO:
        """Get active worksheet."""
        response = send_daemon_command(f"GET_ACTIVE:{self._session_id}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])
        if response.startswith("SHEET:"):
            sheet_name = response[6:]
            return self[sheet_name]
        raise RecalcError(f"Unexpected response: {response}")

    @active.setter
    def active(self, sheet: WorksheetProxyUNO | str) -> None:
        """Set active worksheet."""
        if self._on_write:
            self._on_write()

        if isinstance(sheet, WorksheetProxyUNO):
            sheet_name = sheet._name
        else:
            sheet_name = sheet

        response = send_daemon_command(f"SET_ACTIVE:{self._session_id}:{sheet_name}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

    def create_sheet(
        self, title: str | None = None, index: int | None = None
    ) -> WorksheetProxyUNO:
        """Create a new worksheet."""
        if self._on_write:
            self._on_write()

        if title is None:
            # Generate unique name
            existing = set(self.sheetnames)
            i = 1
            while f"Sheet{i}" in existing:
                i += 1
            title = f"Sheet{i}"

        cmd = f"CREATE_SHEET:{self._session_id}:{title}"
        if index is not None:
            cmd += f":{index}"

        response = send_daemon_command(cmd)
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

        proxy = WorksheetProxyUNO(self, title, self._on_write)
        self._sheet_cache[title] = proxy
        return proxy

    @property
    def worksheets(self) -> list[WorksheetProxyUNO]:
        """Get list of all worksheets."""
        return [self[name] for name in self.sheetnames]

    def save(self, path: str | None = None) -> None:
        """Save workbook to file."""
        # Note: path parameter is ignored, always saves to original path
        response = send_daemon_command(f"SAVE:{self._session_id}")
        if response.startswith("ERROR:"):
            raise RecalcError(response[6:])

    def close(self) -> None:
        """Close the session."""
        send_daemon_command(f"CLOSE:{self._session_id}")
        # Ignore errors on close

    def __repr__(self) -> str:
        return f"<WorkbookProxyUNO session={self._session_id}>"
