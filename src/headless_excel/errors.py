"""Exception types for exwrap."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from headless_excel.context import ErrorDetail


class ExcelError(Exception):
    """Base exception for exwrap errors."""


class RecalcError(ExcelError):
    """Error during LibreOffice recalculation."""


class SyncError(ExcelError):
    """Error during sync operation."""


@dataclass
class FormulaError(ExcelError):
    """Formula errors found after recalculation.

    Attributes:
        errors: Dict mapping error type (e.g., '#REF!') to list of cell locations
        total: Total number of formula errors
        error_details: List of detailed error information (formula, neighbors)
    """

    errors: dict[str, list[str]] = field(default_factory=dict)
    total: int = 0
    error_details: list[ErrorDetail] = field(default_factory=list)

    def __str__(self) -> str:
        if not self.errors:
            return "No formula errors"

        lines = [f"Formula errors ({self.total}):"]

        if self.error_details:
            for detail in self.error_details:
                parts = [f"  {detail.location}: {detail.error}"]
                if detail.formula:
                    parts.append(f" formula={detail.formula}")
                if detail.neighbors:
                    neighbors_str = ", ".join(
                        f"{k}={_format_value(v)}"
                        for k, v in detail.neighbors.items()
                    )
                    parts.append(f" inputs={{{neighbors_str}}}")
                lines.append("".join(parts))
        else:
            for error_type, locations in self.errors.items():
                lines.append(f"  {error_type}: {locations}")

        return "\n".join(lines)


def _format_value(v: Any) -> str:
    """Format a value for display, truncating long strings."""
    if v is None:
        return "None"
    if isinstance(v, str):
        if len(v) > 20:
            return repr(v[:17] + "...")
        return repr(v)
    return str(v)


@dataclass
class ColorLintViolation:
    """Details about a single color lint violation.

    Attributes:
        cell: Cell coordinate (e.g., 'A1')
        value: The cell's value
        expected_color: The color that should be applied based on conventions
        current_color: The color currently applied (None if no color set)
    """

    cell: str
    value: Any
    expected_color: str
    current_color: str | None


@dataclass
class ColorLintError(ExcelError):
    """Financial color conventions violated.

    Raised when cells don't follow standard financial modeling color conventions:
    - Blue: Hardcoded input values
    - Black: Formulas
    - Green: External/cross-sheet links

    Attributes:
        violations: Dict mapping sheet names to list of violations
        total: Total number of violations
    """

    violations: dict[str, list[ColorLintViolation]] = field(default_factory=dict)
    total: int = 0

    def __str__(self) -> str:
        if not self.violations:
            return "No color lint violations"

        lines = [f"Financial color violations ({self.total}):"]
        for sheet_name, sheet_violations in self.violations.items():
            for v in sheet_violations:
                expected = _color_name(v.expected_color)
                current = _color_name(v.current_color) if v.current_color else "none"
                val_str = _format_value(v.value)
                lines.append(
                    f"  {sheet_name}!{v.cell}: expected {expected}, got {current} (value={val_str})"
                )
        return "\n".join(lines)


def _color_name(color_code: str | None) -> str:
    """Convert color code to human-readable name."""
    from headless_excel.formats import Colors

    if color_code is None:
        return "none"
    color_map = {
        Colors.HARDCODE: "HARDCODE (blue)",
        Colors.FORMULA: "FORMULA (black)",
        Colors.EXTERNAL_LINK: "EXTERNAL_LINK (green)",
    }
    return color_map.get(color_code, color_code)
