"""Exception types and result dataclasses for headless-excel."""

from dataclasses import dataclass, field
from typing import Any

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
    """Base exception for headless-excel errors."""


class RecalcError(ExcelError):
    """Error during LibreOffice recalculation."""


class LibreOfficeNotFoundError(ExcelError):
    """LibreOffice is not installed or not in PATH."""


class SyncError(ExcelError):
    """Error during sync operation."""


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
