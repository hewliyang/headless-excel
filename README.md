# headless-excel

Excel wrapper with automatic recalculation via LibreOffice — an OfficeJS-style API for headless environments.

## Problem

When manipulating Excel files programmatically with `openpyxl` in a headless environment (no Excel engine running), formulas don't evaluate. You need to:

1. Know to read with `data_only=True` to get materialized values
2. Remember to save after modifications
3. Constantly reload after saves
4. Manually run recalculation to evaluate formulas
5. Check for formula errors (`#REF!`, `#NAME?`, etc.) to close the feedback loop

This is especially painful for AI agents that need immediate feedback on formula validity.

## Solution

`headless-excel` provides an OfficeJS-inspired API that handles all of this automatically:

```python
from headless_excel import run

with run("model.xlsx") as ctx:
    # Make changes using openpyxl workbook
    ws = ctx.active
    ws["A1"] = 100
    ws["A2"] = "=A1*2"
    ws["A3"] = "=SUM(A1:A2)"

    # sync() saves, recalculates via LibreOffice, and reloads
    result = ctx.sync()

    if not result.success:
        print(f"Formula errors: {result.errors}")
        # e.g., {'#REF!': ['Sheet1!B5'], '#NAME?': ['Sheet1!C3']}

    # Read materialized values directly from cells
    print(ctx.active["A3"].value)  # 300
```

## API

### `run(path, create=False, auto_sync=True, raise_on_errors=False)`

Context manager for Excel operations. Similar to OfficeJS `Excel.run()`.

```python
from headless_excel import run

# Create new workbook
with run("new.xlsx", create=True) as ctx:
    ctx.active["A1"] = "Hello"
    # auto-syncs on exit

# Load existing, manual sync
with run("existing.xlsx", auto_sync=False) as ctx:
    ctx.active["A1"] = "=B1+C1"
    result = ctx.sync()
    print(ctx.active["A1"].value)

# Raise on formula errors
with run("model.xlsx", raise_on_errors=True) as ctx:
    ctx.active["A1"] = "=INVALID_REF"
    # Raises FormulaError on exit
```

### `ExcelContext`

The context object provides:

| Property/Method              | Description                                |
| ---------------------------- | ------------------------------------------ |
| `workbook` / `wb`            | The workbook proxy for editing and reading |
| `active`                     | Get/set active worksheet (assignable)      |
| `sheet(name)`                | Get worksheet by name                      |
| `create_sheet(name)`         | Create new worksheet                       |
| `delete_sheet(name)`         | Delete worksheet by name                   |
| `sync(raise_on_errors=True)` | Save, recalc via LibreOffice, reload       |
| `values`                     | Raw workbook with data_only=True           |

#### Sheet Management

```python
with run("model.xlsx", create=True) as ctx:
    # Create sheets (default "Sheet" auto-removed on first create)
    data = ctx.create_sheet("Data")
    summary = ctx.create_sheet("Summary")
    intro = ctx.create_sheet("Intro", 0)  # insert at index

    # Switch active sheet
    ctx.active = ctx.sheet("Data")  # by proxy
    ctx.active = "Summary"          # by name

    # Rename sheet
    ctx.sheet("Data").title = "RawData"

    # Reorder sheets (offset: negative=left, positive=right)
    ctx.workbook.move_sheet("Summary", -1)

    # Delete sheet
    ctx.delete_sheet("RawData")
```

### Ranges

Use `sheet.range(ref)` for bulk read/write operations on cell ranges. This avoids off-by-one errors when working with 2D data.

```python
with run("model.xlsx", create=True) as ctx:
    ws = ctx.active

    # Write 2D array to range (validates dimensions)
    # Can include formulas (strings starting with '=')
    ws.range("A1:C3").values = [
        [10, 20, "=A1+B1"],
        [30, 40, "=A2+B2"],
        [50, 60, "=SUM(C1:C2)"],
    ]

    # Read 2D array from range
    data = ws.range("A1:C3").values  # [[10, 20, "=A1+B1"], ...]

    # Get range shape
    r = ws.range("B2:E5")
    r.shape      # (4, 4)
    r.num_rows   # 4
    r.num_cols   # 4

    # Get formulas in range (dict, only formula cells)
    ws["A1"] = "=B1+C1"
    ws.range("A1:B1").formulas  # {"A1": "=B1+C1"}

    # Get all formulas in sheet
    ws.formulas  # {"A1": "=B1+C1", ...}
```

#### Range Styling

Apply styles to all cells in a range at once:

```python
from openpyxl.styles import Font, PatternFill, Alignment

with run("model.xlsx", create=True) as ctx:
    ws = ctx.active
    ws.range("A1:D1").values = [["Q1", "Q2", "Q3", "Q4"]]

    # Apply multiple styles
    ws.range("A1:D1").apply_style(
        font=Font(bold=True, size=12),
        fill=PatternFill(start_color="CCCCCC", fill_type="solid"),
        alignment=Alignment(horizontal="center"),
        number_format="#,##0",
    )
```

### `SyncResult`

Returned by `sync()`:

```python
result = ctx.sync()
result.success       # bool - True if no formula errors
result.total_errors  # int - count of errors
result.errors        # dict - {'#REF!': ['Sheet1!A1'], ...}
result.error_details # list[ErrorDetail] - detailed error info
result.raise_on_errors()  # raise FormulaError if errors exist

# ErrorDetail provides formula context for debugging
for detail in result.error_details:
    detail.location   # 'Sheet1!A1'
    detail.error      # '#DIV/0!'
    detail.formula    # '=B1/C1'
    detail.neighbors  # {'Sheet1!B1': 100, 'Sheet1!C1': 0}
```

### Number Formatting & Colors

Use built-in constants instead of remembering format strings:

```python
from headless_excel import run, NumberFormats, Colors
from openpyxl.styles import Font

with run("model.xlsx") as ctx:
    ws = ctx.active

    # Apply accounting format (aligns $, negatives in parentheses)
    ws['A1'] = 1234.56
    ws['A1'].number_format = NumberFormats.ACCOUNTING

    # Percentage with 2 decimals
    ws['B1'] = 0.1575
    ws['B1'].number_format = NumberFormats.PERCENTAGE_2DP

    # Apply format to range
    ws.range("C1:C10").apply_style(number_format=NumberFormats.NUMBER)

    # Color conventions for financial models
    ws['D1'] = 100  # Hardcoded input
    ws['D1'].font = Font(color=Colors.HARDCODE)  # Blue

    ws['E1'] = "=D1*2"  # Formula
    ws['E1'].font = Font(color=Colors.FORMULA)  # Black
```

**Available formats:** `ACCOUNTING`, `ACCOUNTING_0DP`, `PERCENTAGE`, `PERCENTAGE_1DP`, `PERCENTAGE_2DP`, `NUMBER`, `NUMBER_0DP`, `DATE`, `DATE_LONG`

**Available colors:** `HARDCODE` (blue), `FORMULA` (black), `EXTERNAL_LINK` (green)

### Exceptions

- `ExcelError` - Base exception
- `FormulaError` - Formula errors found (has `.errors` dict)
- `RecalcError` - LibreOffice recalculation failed
- `SyncError` - Save or sync operation failed

## How It Works

1. **Edit**: Use standard openpyxl operations on `ctx.workbook`
2. **Sync**: Calls `ctx.sync()` which:
   - Saves the workbook to disk
   - Runs LibreOffice in headless mode with a macro to recalculate all formulas
   - Reloads the workbook twice: once for formulas, once with `data_only=True` for values
   - Scans for Excel error values (`#REF!`, `#NAME?`, `#VALUE!`, etc.)
3. **Read**: Access computed values via cell `.value` or `range.values`

## For AI Agents

The key benefit for agentic workflows is the **feedback loop**:

```python
from headless_excel import run, FormulaError

def build_model(path: str):
    with run(path, create=True, auto_sync=False) as ctx:
        ws = ctx.active

        # Agent builds formulas...
        ws["A1"] = "=SomeFormula"

        # Sync and check for errors
        result = ctx.sync()

        if not result.success:
            # Agent can iterate based on specific errors
            for error_type, locations in result.errors.items():
                print(f"{error_type} at {locations}")
            # ... fix and retry

        # Validate computed values
        actual = ws["A1"].value
        if actual != expected:
            # ... adjust formula
```
