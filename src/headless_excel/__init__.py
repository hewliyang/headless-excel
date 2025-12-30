"""headless-excel - Excel wrapper with automatic recalculation via LibreOffice."""

from headless_excel.context import ErrorDetail, ExcelContext, SyncResult, create, run
from headless_excel.errors import (
    ColorLintViolation,
    ExcelError,
    FormulaError,
    RecalcError,
    SyncError,
    get_max_errors_displayed,
    set_max_errors_displayed,
)
from headless_excel.formats import Colors, NumberFormats, infer_financial_color
from headless_excel.proxy import (
    EXCEL_ERRORS,
    CellProxy,
    RangeProxy,
    WorkbookProxy,
    WorksheetProxy,
)

__all__ = [
    "CellProxy",
    "ColorLintViolation",
    "Colors",
    "EXCEL_ERRORS",
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
    "create",
    "get_max_errors_displayed",
    "infer_financial_color",
    "run",
    "set_max_errors_displayed",
]
