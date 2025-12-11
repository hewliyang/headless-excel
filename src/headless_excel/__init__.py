"""headless-excel - Excel wrapper with automatic recalculation via LibreOffice."""

from headless_excel.context import ErrorDetail, ExcelContext, SyncResult, run
from headless_excel.errors import ExcelError, FormulaError, RecalcError, SyncError
from headless_excel.proxy import CellProxy, WorkbookProxy, WorksheetProxy

__all__ = [
    "CellProxy",
    "ErrorDetail",
    "ExcelContext",
    "ExcelError",
    "FormulaError",
    "RecalcError",
    "SyncError",
    "SyncResult",
    "WorkbookProxy",
    "WorksheetProxy",
    "run",
]
