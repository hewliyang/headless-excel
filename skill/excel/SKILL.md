---
name: excel
description: "Creating & editing Excel workbooks via CLI"
---

# headless-excel CLI

All operations are bash calls. Code runs with `ctx` (ExcelContext) and `NumberFormats` in globals.

## Commands

```bash
headless-excel check                    # verify LibreOffice + hooks
headless-excel create output.xlsx       # create new file
headless-excel eval file.xlsx "code"    # run Python openpyxl code against file
```

## Workflow

1. `create` to scaffold
2. `eval` to mutate (auto-syncs on exit)
3. For read-after-write in same eval, call `ctx.sync()` first

## Write

```bash
headless-excel eval model.xlsx "
ws = ctx.active
ws['A1'] = 'Revenue'
ws['A2'] = 100
ws['A3'] = '=SUM(A2:A2)'
"
```

## Read

```bash
headless-excel eval model.xlsx "
ctx.sync()
print(ctx.active['A3'].value)
print(ctx.active.range('A1:D10').dump())
"
```

## Multi-Sheet Workbooks

```bash
headless-excel eval model.xlsx "
# Get sheets directly by name or create
data = ctx.sheet('Sheet')      # get existing sheet
data.title = 'Data'
data['A1'] = 100

summary = ctx.create_sheet('Summary')  # create new sheet
summary['A1'] = '=SUM(Data!A:A)'

# Verify
ctx.sync()
print(data.range('A1').dump())
print(summary.range('A1').dump())
"
```

## Styling

```bash
headless-excel eval model.xlsx "
from openpyxl.styles import Font, PatternFill

ws = ctx.active
ws['A1'].font = Font(bold=True)
ws['A1'].fill = PatternFill('solid', fgColor='4472C4')
ws.range('B2:B10').apply_style(number_format=NumberFormats.ACCOUNTING)
"
```

`NumberFormats`: ACCOUNTING, ACCOUNTING_0DP, PERCENTAGE, PERCENTAGE_1DP, PERCENTAGE_2DP, NUMBER, NUMBER_0DP, DATE, DATE_LONG

## Auto-Fill

```bash
headless-excel eval model.xlsx "
ws = ctx.active
ws['A1'] = '=B1*C1'
ws.range('A1:A10').auto_fill()  # fills A2:A10 with adjusted formulas
"
```

## Clearing Cells

```bash
headless-excel eval model.xlsx "
ws = ctx.active
ws.range('A1:D10').clear()  # clears values, formulas, and styles
"
```

## Iterative Calculations (Circular References)

For formulas with circular references that should converge (e.g., goal-seek scenarios), enable iterative calculation on the workbook:

```bash
headless-excel eval model.xlsx "
# Enable iterative calculations
ctx.wb.calculation.iterate = True
ctx.wb.calculation.iterateCount = 100   # max iterations
ctx.wb.calculation.iterateDelta = 0.001 # convergence threshold

ws = ctx.active
ws['A1'] = 1            # seed value
ws['B1'] = '=A1/2'
ws['A1'] = '=B1+0.5'    # circular: A1 -> B1 -> A1, converges to A1=1, B1=0.5
"
```

Without `iterate = True`, circular refs produce `#VALUE!` errors. With it enabled, LibreOffice iterates until convergence or max iterations.

## Pivot Tables

Declare a pivot with `ws.add_pivot(...)`. Pivots are **declared, not computed in Python** — the source cache is flagged `refreshOnLoad`, so `ctx.sync()` (LibreOffice) or Excel materialises every value and grand total. You must `ctx.sync()` to see results.

Fluent form:

```bash
headless-excel eval model.xlsx "
data = ctx.sheet('Sheet'); data.title = 'Data'
data.range('A1:D7').values = [
    ['Region', 'Product', 'Units', 'Sales'],
    ['North', 'A', 10, 100], ['North', 'B', 5, 80],
    ['South', 'A', 7, 70],  ['South', 'B', 3, 60],
    ['North', 'A', 4, 40],  ['South', 'B', 9, 90],
]

pivot = ctx.create_sheet('Pivot')
(pivot.add_pivot('Data!A1:D7', anchor='A3', name='ByRegion')
      .rows('Region')
      .columns('Product')
      .values('Sales', func='sum', num_format='\$#,##0')
      .style('PivotStyleMedium9'))

ctx.sync()  # LibreOffice computes the aggregations
print(pivot.range('A3:E8').dump())
"
```

Declarative form (equivalent):

```python
pivot.add_pivot(
    'Data!A1:D7', anchor='A3', name='ByRegion',
    rows=['Region'], columns=['Product'],
    values=[('Sales', 'sum', 'Total Sales', '$#,##0')],
    style='PivotStyleMedium9',
)
```

`filters=[...]` adds report-filter (page) fields when your source has them. All field names must match the source header row.

Aggregation functions: `sum, count, countNums, average, max, min, product, stdDev, stdDevp, var, varp` (aliases like `avg`/`mean` accepted).

LibreOffice may ignore a value field's number format. To force it onto the rendered cells, call `.format_values(ctx)` after building (it syncs and styles the materialised value region).

## Best Practices

The whole point of Excel is that values recalculate automatically when inputs change. Avoid computing values in Python and simply writing static numbers.

Bad (hardcoded):

```python
total = sum(values)  # computed in Python
ws['A10'] = total    # user edits data, total is now wrong
```

Good (formula):

```python
ws['A10'] = '=SUM(A1:A9)'  # always up to date
```

## Layout

- Prefer uniform column widths
- Use empty columns for indentation (not varying widths)
- Always specify units in headers: `Revenue ($mm)`, `Growth (%)`

## API Reference

```python
# ExcelContext (ctx)
ctx.active                      # get/set active WorksheetProxy
ctx.sheet(name)                 # get existing sheet by name
ctx.create_sheet(title, index=None)  # create new sheet
ctx.delete_sheet(name)          # delete sheet
ctx.sync(raise_on_errors=False) # save, recalc via LibreOffice, reload
ctx.wb                          # WorkbookProxy (also ctx.workbook)

# WorksheetProxy (ws) - also has all openpyxl Worksheet attrs
ws['A1']                        # get CellProxy
ws['A1'] = value                # set cell value
ws.cell(row, col, value=None)   # get CellProxy by row/col index
ws.range('A1:C3')               # get RangeProxy
ws.title                        # sheet name (read/write)
ws.column_dimensions['A'].width # column width

# RangeProxy
range.values                    # get/set 2D array
range.formulas                  # dict of {coord: formula}
range.shape                     # (rows, cols)
range.apply_style(              # bulk styling
    font=None,                  # Font
    fill=None,                  # PatternFill
    alignment=None,             # Alignment
    border=None,                # Border
    number_format=None,         # str (use NumberFormats.*)
)
range.auto_fill(                # fill formulas like Excel drag
    direction=None,             # 'down'|'right'|'up'|'left' (auto-detects)
    source_rows=1,              # rows to use as pattern
    copy_styles=True,
)
range.clear(styles=True)        # clear values and optionally styles
range.dump(show_formulas=False) # formatted table string for debugging

# Pivot tables (computed by LibreOffice/Excel on sync)
ws.add_pivot(                    # returns a fluent PivotTable
    source,                     # e.g. 'Data!A1:D7' (header in first row)
    anchor='A3',                # top-left cell on this sheet
    name='PivotTable1',
    rows=None, columns=None,    # optional field-name lists
    filters=None,               # optional report-filter fields
    values=None,                # ['Sales'] or [(name, func[, display[, fmt]])]
    style=None,                 # e.g. 'PivotStyleMedium9'
)
pt.rows(*names)                 # fluent equivalents of the kwargs above
pt.columns(*names)
pt.filters(*names)
pt.values(name, func='sum', display_name=None, num_format=None)
pt.values_axis('columns')       # or 'rows' (place value labels)
pt.style(name='PivotStyleMedium9', row_stripes=True, ...)
pt.format_values(ctx, num_format=None)  # force value number format post-sync

# CellProxy
cell.value                      # get materialized value / set value
cell.formula                    # get formula string or None (read-only)
cell.font, cell.fill, cell.border, cell.number_format, cell.alignment
```

## Tips

- Keep each `eval` under 100 LoC; build worksheets incrementally across multiple calls
- `ctx.sheet('Name')` to get existing sheet
- `ctx.create_sheet('Name')` to create new sheet
- `ws.range('A1:Z100').clear()` to clear a range
- Call `ctx.sync()` before reading values you just wrote
- Use `print(ws.range(...).dump())` to verify
