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
                        f"{k}={_format_value(v)}" for k, v in detail.neighbors.items()
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


@dataclass
class ColorLintResult:
    """Result of linting financial colors across sheets.

    Attributes:
        violations_by_sheet: Dict mapping sheet names to list of violations
        total_violations: Total count of violations across all sheets
    """

    violations_by_sheet: dict[str, list[ColorLintViolation]] = field(
        default_factory=dict
    )
    total_violations: int = 0

    def __bool__(self) -> bool:
        """Return True if there are violations."""
        return self.total_violations > 0

    def __len__(self) -> int:
        """Return number of sheets with violations."""
        return len(self.violations_by_sheet)

    def __contains__(self, sheet_name: str) -> bool:
        """Check if a sheet has violations."""
        return sheet_name in self.violations_by_sheet

    def __getitem__(self, sheet_name: str) -> list[ColorLintViolation]:
        """Get violations for a specific sheet."""
        return self.violations_by_sheet[sheet_name]

    def __iter__(self):
        """Iterate over sheet names."""
        return iter(self.violations_by_sheet)

    def items(self):
        """Iterate over (sheet_name, violations) pairs."""
        return self.violations_by_sheet.items()

    def __repr__(self) -> str:
        """Format violations with truncation to prevent context pollution."""
        if not self.violations_by_sheet:
            return "No color lint violations"

        max_display = get_max_errors_displayed()
        lines = [f"Color violations ({self.total_violations}):"]
        displayed = 0
        truncated = 0

        for sheet_name, sheet_violations in self.violations_by_sheet.items():
            # Group by expected color for token efficiency
            by_expected: dict[str, list[str]] = {}
            for v in sheet_violations:
                if displayed >= max_display:
                    truncated += 1
                    continue
                expected = _color_name(v.expected_color)
                if expected not in by_expected:
                    by_expected[expected] = []
                by_expected[expected].append(v.cell)
                displayed += 1

            for expected, cells in by_expected.items():
                lines.append(f"  {sheet_name}! need {expected}: {', '.join(cells)}")

        if truncated > 0:
            lines.append(f"  ... +{truncated} more")

        return "\n".join(lines)

    __str__ = __repr__


@dataclass
class ErrorScanResult:
    """Result of scanning for formula errors.

    Attributes:
        errors_by_type: Dict mapping error type (e.g., '#DIV/0!') to list of locations
        total_errors: Total count of errors across all types
    """

    errors_by_type: dict[str, list[str]] = field(default_factory=dict)
    total_errors: int = 0

    def __bool__(self) -> bool:
        """Return True if there are errors."""
        return self.total_errors > 0

    def __len__(self) -> int:
        """Return number of error types found."""
        return len(self.errors_by_type)

    def __contains__(self, error_type: str) -> bool:
        """Check if an error type exists."""
        return error_type in self.errors_by_type

    def __getitem__(self, error_type: str) -> list[str]:
        """Get locations for a specific error type."""
        return self.errors_by_type[error_type]

    def __iter__(self):
        """Iterate over error types."""
        return iter(self.errors_by_type)

    def items(self):
        """Iterate over (error_type, locations) pairs."""
        return self.errors_by_type.items()

    def __eq__(self, other):
        """Allow comparison with empty dict for backwards compatibility."""
        if isinstance(other, dict):
            return self.errors_by_type == other
        return super().__eq__(other)

    def __repr__(self) -> str:
        """Format errors with truncation to prevent context pollution."""
        if not self.errors_by_type:
            return "No formula errors"

        max_display = get_max_errors_displayed()
        lines = [f"Formula errors ({self.total_errors}):"]
        displayed = 0
        truncated = 0

        for error_type, locations in self.errors_by_type.items():
            # Group locations per error type for token efficiency
            shown_locs = []
            for loc in locations:
                if displayed >= max_display:
                    truncated += 1
                else:
                    shown_locs.append(loc)
                    displayed += 1
            if shown_locs:
                lines.append(f"  {error_type}: {', '.join(shown_locs)}")

        if truncated > 0:
            lines.append(f"  ... +{truncated} more")

        return "\n".join(lines)

    __str__ = __repr__


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
