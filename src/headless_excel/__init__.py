"""headless-excel - Excel wrapper with automatic recalculation via LibreOffice."""

from headless_excel.context import ErrorDetail, ExcelContext, SyncResult, run
from headless_excel.errors import ExcelError, FormulaError, RecalcError, SyncError
from headless_excel.formats import Colors, NumberFormats
from headless_excel.proxy import CellProxy, RangeProxy, WorkbookProxy, WorksheetProxy

__all__ = [
    "CellProxy",
    "Colors",
    "ErrorDetail",
    "ExcelContext",
    "ExcelError",
    "FormulaError",
    "NumberFormats",
    "RangeProxy",
    "RecalcError",
    "SyncError",
    "SyncResult",
    "WorkbookProxy",
    "WorksheetProxy",
    "run",
]
