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

**PREFER BULK APIs** over cell-by-cell:

- `ws.write('A1', [[...], [...]])` — bulk write (simplest, no range math)
- `ws.range('A1:C3').values = [[...]]` — bulk write (lenient, data shape wins)
- `ws.range('A1:C3').apply_style(...)` — bulk style
- `ws.range('A1:A10').auto_fill()` — fill formulas, ala drag to complete
- `print(ws.range('A1:D10').dump())` — verify output

## Write

```bash
headless-excel eval model.xlsx "
ws = ctx.active

# Bulk write with write() - simplest, no range to get wrong
ws.write('A1', [
    ['Revenue', 'Q1', 'Q2'],
    ['Product A', 100, 150],
    ['Product B', 200, '=B3*1.1'],
])
# stderr: [headless-excel] .write() wrote to A1:C3 (3×3)

# Single cell (when needed)
ws['A5'] = '=SUM(B2:B3)'

ctx.sync()
print(ws.range('A1:C5').dump())  # verify
"
```

Bulk write with `range().values` is lenient - data shape determines actual region:
```bash
headless-excel eval model.xlsx "
ws = ctx.active
ws.range('A1:D10').values = [  # range is just a hint
    ['Name', 'Value'],
    ['Alpha', 100],
]
# stderr: [headless-excel] .values wrote to A1:B2 (specified A1:D10, got 2×2)
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

# Bulk style (preferred)
ws.range('A1:D1').apply_style(
    font=Font(bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor='4472C4')
)
ws.range('B2:D10').apply_style(number_format=NumberFormats.ACCOUNTING)

# Single cell (when needed)
ws['A1'].font = Font(bold=True)
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

- `ws.write('A1', [[...]])` for bulk writes (simplest, no range math needed)
- `ws.range('A1:C3').values = [[...]]` is lenient (data shape wins, check stderr)
- `ws.range('A1:A10').auto_fill()` to fill formulas down
- `ctx.sheet('Name')` to get existing sheet
- `ctx.create_sheet('Name')` to create new sheet
- `ws.range('A1:Z100').clear()` to clear a range
- Call `ctx.sync()` before reading values you just wrote
- **Always `print(ws.range(...).dump())` to verify after edits**
