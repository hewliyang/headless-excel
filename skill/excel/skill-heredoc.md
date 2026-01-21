---
name: xlsx
description: "Creating & editing Excel workbooks"
---

# headless-excel

Wraps `openpyxl` with automatic formula recalculation and error detection via LibreOffice.

Use heredocs to run inline python scripts in bash instead of creating entire scripts. Keep them under 100 LoC per batch of changes, you should build up the worksheets incrementally. Call `ctx.sync()` frequently to catch errors early and avoid error propagation.

**CRITICAL**
Remember that formula results do not materialize until either the context manager exits and reenters OR ctx.sync() is called manually.
For any APIs not available in `headless-excel`, remember that `ws` is just a `openpyxl.Worksheet` object and `ctx.wb` is just a `openpyxl.Workbook`.

**PREFER BULK APIs** — avoid cell-by-cell loops:

```py
# ❌ SLOW: cell-by-cell loops
for i, val in enumerate(['A', 'B', 'C']):
    ws.cell(row=1, column=i+1, value=val)

# ✅ FAST: bulk write with .write() (simplest - no range math)
ws.write("A1", [['A', 'B', 'C']])
# stderr: [headless-excel] .write() wrote to A1:C1 (1×3)

# ✅ FAST: bulk write with .values (lenient - data shape wins)
ws.range("A1:C1").values = [['A', 'B', 'C']]
# stderr: [headless-excel] .values wrote to A1:C1 (1×3)

# ✅ FAST: bulk style with .apply_style  
ws.range("A1:C1").apply_style(font=Font(bold=True), fill=PatternFill('solid', fgColor='4472C4'))

# ✅ FAST: auto_fill formulas instead of loops
ws['A1'] = '=B1*2'
ws.range('A1:A100').auto_fill()
```

**DEBUGGING: Always use `.dump()` instead of loops**

```py
# ❌ NEVER do this
for row in ws.iter_rows(min_row=1, max_row=5):
    print([cell.value for cell in row])

# ✅ ALWAYS do this
print(ws.range("A1:D5").dump())                  # values as table
print(ws.range("A1:D5").dump(show_formulas=True)) # formulas as table
```

**⚠️ VERIFY AFTER EVERY EDIT:** Always `print(ws.range(...).dump())` after edits to catch mistakes early.

## Creating Files

```py
from headless_excel import create, NumberFormats
from openpyxl.styles import Font, PatternFill

with create("output.xlsx", overwrite=True) as ctx:
    ws = ctx.active
    
    # Bulk write with .write() (simplest - no range math)
    ws.write("A1", [
        ['Revenue', 'Q1', 'Q2'],
        ['Product A', 100, 150],
        ['Product B', 200, '=B3*1.1'],
    ])
    # stderr: [headless-excel] .write() wrote to A1:C3 (3×3)
    
    # Bulk style with .apply_style
    ws.range("A1:C1").apply_style(font=Font(bold=True), fill=PatternFill('solid', fgColor='4472C4'))
    ws.range("B2:C3").apply_style(number_format=NumberFormats.ACCOUNTING)
    
    # Single cell (when needed)
    ws['A5'] = '=SUM(B2:B3)'
    
    ctx.sync()
    print(ws.range("A1:C5").dump())  # verify
```

## Editing Files

```py
from headless_excel import run

with run("existing.xlsx") as ctx:
    ws = ctx.active
    ws['A1'] = 'New Value'
    ws['B2'] = '=A1*2'

    # create_sheet() returns a WorksheetProxy you can use immediately
    assumptions = ctx.create_sheet("Assumptions")
    assumptions['A1'] = 'Title'  # ✅ Use the returned proxy directly
    cover = ctx.create_sheet("Cover", 0)  # insert at front
    assumptions.title = "Model Assumptions"
    ctx.delete_sheet("OldSheet")
    ctx.workbook.move_sheet("Revenue", -1)

    # ctx.active = "SheetName" only works for EXISTING sheets
    ctx.active = "Revenue"  # switch to an existing sheet
    ws = ctx.active
```

## Reading Values & Formulas

```py
from headless_excel import run

with run("model.xlsx") as ctx:
    ctx.sync()
    # Single value
    result = ctx.active['A1'].value
    # Single cell formula (or None if not a formula)
    formula = ctx.active['A1'].formula
    # Range as 2D array
    data = ctx.sheet("Summary").range("A1:D10").values
    # Formulas (dict of cell->formula)
    formulas = ctx.active.formulas
    formulas = ctx.active.range("A1:C10").formulas
```

## Debugging with dump()

```py
from headless_excel import run

with run("model.xlsx") as ctx:
    ctx.sync()
    # Print range as formatted table (dump returns string, use print())
    print(ctx.active.range("A1:D5").dump())
    # |   A |   B |   C |   D |
    # |-----|-----|-----|-----|
    # | 100 | 200 | 300 | 600 |
    # ...

    # Show formulas instead of values
    print(ctx.active.range("A1:D5").dump(show_formulas=True))
```

## Bulk Read/Write with Ranges

```py
from headless_excel import create

with create("model.xlsx") as ctx:
    ws = ctx.active
    
    # write() - simplest, just anchor + data
    written = ws.write("A1", [
        [10, 20, "=A1+B1"],
        [30, 40, "=A2+B2"],
        [50, 60, "=SUM(C1:C2)"],
    ])
    # stderr: [headless-excel] .write() wrote to A1:C3 (3×3)
    # returns: "A1:C3"
    
    # range().values - lenient, data shape determines actual region
    ws.range("E1:Z100").values = [[1, 2], [3, 4]]
    # stderr: [headless-excel] .values wrote to E1:F2 (specified E1:Z100, got 2×2)
    
    # Read data
    data = ws.range("A1:C3").values
    print(ws.range("B2:E5").shape)  # (4, 4)
```

## Clearing Ranges

```py
ws.range("A1:C3").clear()              # Clear values only
ws.range("A1:C3").clear(styles=True)   # Clear values and reset styles
```

## Auto Fill (Like Excel Drag Handle)

```py
ws["A1"] = "=B1*C1"
ws.range("A1:A100").auto_fill()  # Fills down: =B2*C2, =B3*C3, ...

ws.range("A1:D1").auto_fill(direction="right")  # Fill right
ws.range("A1:A10").auto_fill(direction="up")    # Fill up from last row

# Absolute refs ($) stay fixed, relative refs adjust
ws["A1"] = "=$B$1*C1"
ws.range("A1:A5").auto_fill()  # =$B$1*C2, =$B$1*C3, ...

# Multi-row pattern (like selecting 2 rows and dragging)
ws["A1"], ws["A2"] = "Label", "=SUM(B1:D1)"
ws.range("A1:A10").auto_fill(source_rows=2)  # Alternates pattern

# Copies styles by default; disable with copy_styles=False
```

## Number Formats

Use built-in constants for number formatting:

```python
from headless_excel import NumberFormats

ws['A1'].number_format = NumberFormats.ACCOUNTING      # $1,234.56
ws['A1'].number_format = NumberFormats.PERCENTAGE_2DP  # 15.00%
```

**Formats:** `ACCOUNTING`, `ACCOUNTING_0DP`, `PERCENTAGE`, `PERCENTAGE_1DP`, `PERCENTAGE_2DP`, `NUMBER`, `NUMBER_0DP`, `DATE`, `DATE_LONG`

## Error Handling

```py
from headless_excel import create

with create("model.xlsx") as ctx:
    ctx.active['A1'] = '=B1/C1'
    ctx.active['B1'] = 10
    ctx.active['C1'] = 0  # Division by zero!

    ctx.sync()
    # Errors always printed to stderr on context exit:
    # SyncResult(success=False, total_errors=1):
    #   Sheet!A1: #DIV/0! formula==B1/C1 inputs={Sheet!B1=10, Sheet!C1=0}
```

## Styling Tips

- Use openpyxl's `Font`, `PatternFill`, `Border` directly for styling
- Set `ws.column_dimensions['A'].width` for appropriate column widths
- Minimal fills, use borders sparingly for section separation
- Appropriate widths so data is readable out of the box
