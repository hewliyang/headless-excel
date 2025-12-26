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
