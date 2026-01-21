---
name: xlsx
description: "Creating & editing Excel workbooks via CLI"
---

# headless-excel CLI

All operations are bash calls. Code runs with `ctx` (ExcelContext) and `NumberFormats` available.

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

## Tips

- `ctx.sheet('Name')` to get existing sheet
- `ctx.create_sheet('Name')` to create new sheet
- `ws.range('A1:Z100').clear()` to clear a range
- Call `ctx.sync()` before reading values you just wrote
- Use `print(ws.range(...).dump())` to verify