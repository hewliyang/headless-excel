"""headless-excel - Excel wrapper with automatic recalculation via LibreOffice."""

from headless_excel.context import ExcelContext, SyncResult, create, run
from headless_excel.errors import (
    ErrorDetail,
    ExcelError,
    FormulaError,
    LibreOfficeNotFoundError,
    RecalcError,
    SyncError,
    get_max_errors_displayed,
    set_max_errors_displayed,
)
from headless_excel.formats import NumberFormats
from headless_excel.hooks import (
    ExtensionAPI,
    clear_hooks,
    extension,
    get_discovery_errors,
    on_exit,
    post_sync,
    pre_sync,
)
from headless_excel.proxy import (
    EXCEL_ERRORS,
    CellProxy,
    RangeProxy,
    WorkbookProxy,
    WorksheetProxy,
)

__all__ = [
    # Context
    "ExcelContext",
    "SyncResult",
    "create",
    "run",
    # Hooks
    "ExtensionAPI",
    "clear_hooks",
    "extension",
    "get_discovery_errors",
    "on_exit",
    "post_sync",
    "pre_sync",
    # Proxies
    "CellProxy",
    "EXCEL_ERRORS",
    "RangeProxy",
    "WorkbookProxy",
    "WorksheetProxy",
    # Errors
    "ErrorDetail",
    "ExcelError",
    "FormulaError",
    "LibreOfficeNotFoundError",
    "RecalcError",
    "SyncError",
    "get_max_errors_displayed",
    "set_max_errors_displayed",
    # Formats
    "NumberFormats",
]
