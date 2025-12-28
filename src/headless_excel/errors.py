"""Exception types for exwrap."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from headless_excel.context import ErrorDetail


# Global configuration for error display limits
# This helps LLM agents by preventing context pollution from walls of errors
_max_errors_displayed: int = 10


def set_max_errors_displayed(limit: int) -> None:
    """Set the maximum number of errors to display in error messages.

    This is a global setting that affects:
    - FormulaError.__str__() and __repr__()
    - format_color_violations()

    Useful for LLM agents with limited context windows.

    Args:
        limit: Maximum number of errors to display (default: 10)

    Example:
        >>> from headless_excel.errors import set_max_errors_displayed
        >>> set_max_errors_displayed(5)  # Show only first 5 errors
    """
    global _max_errors_displayed
    _max_errors_displayed = limit


def get_max_errors_displayed() -> int:
    """Get the current maximum number of errors to display.

    Returns:
        Current limit for errors displayed
    """
    return _max_errors_displayed


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

        max_display = get_max_errors_displayed()
        lines = [f"Formula errors ({self.total}):"]
        displayed = 0
        truncated = 0

        if self.error_details:
            for detail in self.error_details:
                if displayed >= max_display:
                    truncated += 1
                    continue
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
                displayed += 1
        else:
            for error_type, locations in self.errors.items():
                for loc in locations:
                    if displayed >= max_display:
                        truncated += 1
                        continue
                    lines.append(f"  {error_type}: {loc}")
                    displayed += 1

        if truncated > 0:
            lines.append(f"  ... and {truncated} more error(s) (truncated)")

        return "\n".join(lines)

    def __repr__(self) -> str:
        # Use same truncated format for repr to avoid context pollution
        return self.__str__()


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


def format_color_violations(
    violations: dict[str, list[ColorLintViolation]],
) -> str:
    """Format color lint violations as a human-readable string.

    Respects the global max_errors_displayed limit to prevent
    context pollution for LLM agents.

    Args:
        violations: Dict mapping sheet names to list of violations

    Returns:
        Formatted string describing all violations (truncated if needed)
    """
    if not violations:
        return "No color lint violations"

    total = sum(len(v) for v in violations.values())
    max_display = get_max_errors_displayed()
    lines = [f"Financial color violations ({total}):"]
    displayed = 0
    truncated = 0

    for sheet_name, sheet_violations in violations.items():
        for v in sheet_violations:
            if displayed >= max_display:
                truncated += 1
                continue
            expected = _color_name(v.expected_color)
            current = _color_name(v.current_color) if v.current_color else "none"
            val_str = _format_value(v.value)
            lines.append(
                f"  {sheet_name}!{v.cell}: expected {expected}, got {current} (value={val_str})"
            )
            displayed += 1

    if truncated > 0:
        lines.append(f"  ... and {truncated} more violation(s) (truncated)")

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
