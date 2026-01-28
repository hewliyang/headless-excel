"""Hide gridlines on all sheets.

This hook automatically hides gridlines when any workbook is opened.
Gridlines are a display-only setting and don't affect data or printing.

Usage:
    # Copy to .headless-excel/hooks/no_gridlines.py
    headless-excel create file.xlsx  # gridlines hidden automatically
"""

from headless_excel import ExcelContext, on_open


@on_open
def hide_gridlines(ctx: ExcelContext) -> None:
    """Hide gridlines on all sheets when workbook opens."""
    for ws in ctx.workbook.worksheets:
        ws._formula_ws.sheet_view.showGridLines = False
    print(f"Hid gridlines on {len(ctx.workbook.worksheets)} sheet(s)")
