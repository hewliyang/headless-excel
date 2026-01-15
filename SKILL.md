---
name: xlsx
description: "Creating & editing with Excel workbooks"
---

# headless-excel

Wraps `openpyxl` with automatic formula recalculation and error detection via LibreOffice.

Use heredocs to run inline python scripts in bash instead of creating entire scripts. Keep them under 100 LoC per batch of changes, you should build up the worksheets incrementally. Call `ctx.sync()` frequently to catch errors early and avoid error propagation.

**CRITICAL**
Remember that formula results do not materialize until either the context manager exits and reenters OR ctx.sync() is called manually.
For any APIs not available in `headless-excel`, remember that `ws` is just a `openpyxl.Worksheet` object and `ctx.wb` is just a `openpyxl.Workbook`.

**DEBUGGING: Always use `.dump()` instead of loops**

```py
# ❌ NEVER do this - verbose and error-prone
for row in ws.iter_rows(min_row=1, max_row=5):
    print([cell.value for cell in row])

# ✅ ALWAYS do this instead
ws.range("A1:D5").dump()                  # values as table
ws.range("A1:D5").dump(show_formulas=True) # formulas as table
```

**⚠️ VERIFY AFTER EVERY EDIT:** Always call `print(ws.range(...).dump())` after edits or `.sync()` to sanity-check the numbers. This catches off-by-one errors, wrong cell references, and formula mistakes before they propagate. Don't skip this step.

## Creating Files

```py
from headless_excel import create, NumberFormats
from openpyxl.styles import Font, PatternFill

with create(
    "output.xlsx",
    overwrite=True, # to replace an existing file, else to continue use `run` instead of `create
) as ctx:
    ws = ctx.active
    ws['A1'] = 'Revenue'
    ws['A2'] = 100
    ws['A3'] = '=SUM(A2:A2)'

    # Cell styling
    ws['A1'].font = Font(bold=True)
    ws['A1'].fill = PatternFill('solid', start_color='FFFF00')

    # Range styling
    ws.range("A2:A3").apply_style(number_format=NumberFormats.ACCOUNTING)
    ws.column_dimensions['A'].width = 20
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
    ws.range("A1:C3").values = [
        [10, 20, "=A1+B1"],
        [30, 40, "=A2+B2"],
        [50, 60, "=SUM(C1:C2)"],
    ]
    data = ws.range("A1:C3").values
    print(ws.range("B2:E5").shape)  # (4, 4)
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

    # Just print the sync result - shows errors with context
    # Note: if manually syncing - context manager will not raise on exit
    print(ctx.sync())
    # SyncResult(success=False, total_errors=1):
    #   Sheet!A1: #DIV/0! formula==B1/C1 inputs={Sheet!B1=10, Sheet!C1=0}
```

## Styling Tips

- Use openpyxl's `Font`, `PatternFill`, `Border` directly for styling
- Set `ws.column_dimensions['A'].width` for appropriate column widths
- Minimal fills, use borders sparingly for section separation
- Appropriate widths so data is readable out of the box
