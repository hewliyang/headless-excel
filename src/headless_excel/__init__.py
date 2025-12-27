"""headless-excel - Excel wrapper with automatic recalculation via LibreOffice."""

from headless_excel.context import ErrorDetail, ExcelContext, SyncResult, create, run
from headless_excel.errors import (
    ColorLintError,
    ColorLintViolation,
    ExcelError,
    FormulaError,
    RecalcError,
    SyncError,
)
from headless_excel.formats import Colors, NumberFormats, infer_financial_color
from headless_excel.proxy import CellProxy, RangeProxy, WorkbookProxy, WorksheetProxy

__all__ = [
    "CellProxy",
    "ColorLintError",
    "ColorLintViolation",
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
    "create",
    "infer_financial_color",
    "run",
]
