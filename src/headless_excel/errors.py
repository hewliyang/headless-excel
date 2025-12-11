"""Exception types for exwrap."""

from dataclasses import dataclass, field


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
    """

    errors: dict[str, list[str]] = field(default_factory=dict)
    total: int = 0

    def __str__(self) -> str:
        if not self.errors:
            return "No formula errors"
        lines = [f"Found {self.total} formula error(s):"]
        for error_type, locations in self.errors.items():
            lines.append(f"  {error_type}: {locations}")
        return "\n".join(lines)
