---
name: xlsx
description: "Spreadsheet creation, editing, and analysis. Use when working with .xlsx/.xlsm/.csv/.tsv files."
---

# headless-excel

Wraps `openpyxl` with automatic formula recalculation and error detection via LibreOffice.

Keep heredocs under 100 LoC per tool call. Call `ctx.sync()` frequently to catch errors early and avoid error propagation.

**CRITICAL**
Remember that formula results do not materialize until either the context manager exits and reenters OR ctx.sync() is called manually.
For any APIs not available in `headless-excel`, remember that `ws` is just a `openpyxl.Worksheet` object and `ctx.wb` is just a `openpyxl.Workbook`.

## Creating Files

```bash
uv run python <<'EOF'
from headless_excel import create, NumberFormats, Colors
from openpyxl.styles import Font, PatternFill

with create("output.xlsx") as ctx:
    ws = ctx.active
    ws['A1'] = 'Revenue'
    ws['A2'] = 100
    ws['A3'] = '=SUM(A2:A2)'

    # Cell styling
    ws['A1'].font = Font(bold=True)
    ws['A1'].fill = PatternFill('solid', start_color='FFFF00')

    # Range styling
    ws.range("A2:A3").apply_style(
        font=Font(color=Colors.FORMULA),
        number_format=NumberFormats.ACCOUNTING
    )
    ws.column_dimensions['A'].width = 20
EOF
```

## Editing Files

```bash
uv run python <<'EOF'
from headless_excel import run

with run("existing.xlsx") as ctx:
    ws = ctx.active
    ws['A1'] = 'New Value'
    ws['B2'] = '=A1*2'

    # Sheets: create, rename, delete, reorder
    assumptions = ctx.create_sheet("Assumptions")
    cover = ctx.create_sheet("Cover", 0)  # insert at front
    assumptions.title = "Model Assumptions"
    ctx.delete_sheet("OldSheet")
    ctx.workbook.move_sheet("Revenue", -1)

    # Switch active sheet
    ctx.active = "Revenue"
EOF
```

## Reading Values & Formulas

```bash
uv run python <<'EOF'
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
EOF
```

## Debugging with dump()

```bash
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx") as ctx:
    ctx.sync()
    # Print range as formatted table
    ctx.active.range("A1:D5").dump()
    # |   A |   B |   C |   D |
    # |-----|-----|-----|-----|
    # | 100 | 200 | 300 | 600 |
    # ...

    # Show formulas instead of values
    ctx.active.range("A1:D5").dump(show_formulas=True)
EOF
```

## Bulk Read/Write with Ranges

```bash
uv run python <<'EOF'
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
EOF
```

## Number Formats & Colors

Use built-in constants for financial model conventions:

```python
from headless_excel import NumberFormats, Colors
from openpyxl.styles import Font

ws['A1'].number_format = NumberFormats.ACCOUNTING      # $1,234.56
ws['A1'].number_format = NumberFormats.PERCENTAGE_2DP  # 15.00%
ws['A1'].font = Font(color=Colors.HARDCODE)      # Blue for inputs
ws['A1'].font = Font(color=Colors.FORMULA)       # Black for formulas
ws['A1'].font = Font(color=Colors.EXTERNAL_LINK) # Green for links
```

**Formats:** `ACCOUNTING`, `ACCOUNTING_0DP`, `PERCENTAGE`, `PERCENTAGE_1DP`, `PERCENTAGE_2DP`, `NUMBER`, `NUMBER_0DP`, `DATE`, `DATE_LONG`

## Error Handling

```bash
uv run python <<'EOF'
from headless_excel import create, FormulaError

try:
    with create("model.xlsx") as ctx:
        ctx.active['A1'] = '=INVALID_REF'
except FormulaError as e:
    print(e.errors)  # {'#NAME?': ['Sheet1!A1']}
EOF

# Manual error handling
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx", auto_sync=False) as ctx:
    ctx.active['A1'] = '=B1/C1'
    result = ctx.sync()
    if not result.success:
        for d in result.error_details:
            print(f"{d.error} at {d.location}: {d.formula}, refs: {d.neighbors}")
EOF
```

## Common Pitfalls

- Column mapping: column 64 = BL, not BK
- Division by zero: check denominators (#DIV/0!)
- Cross-sheet refs: use `Sheet1!A1` format
