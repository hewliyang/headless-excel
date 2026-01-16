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
from headless_excel import create, run

# Create new workbook
with create("model.xlsx") as ctx:
    ws = ctx.active
    ws["A1"] = 100
    ws["A2"] = "=A1*2"
    ws["A3"] = "=SUM(A1:A2)"
    # auto-syncs on exit

# Open existing workbook
with run("model.xlsx") as ctx:
    # sync() saves, recalculates via LibreOffice, and reloads
    result = ctx.sync()

    if not result.success:
        print(f"Formula errors: {result.errors}")
        # e.g., {'#REF!': ['Sheet1!B5'], '#NAME?': ['Sheet1!C3']}

    # Read materialized values directly from cells
    print(ctx.active["A3"].value)  # 300
```

## API

### `create(path, overwrite=False, auto_sync=True, verbose_errors=True, ...)`

Create a new Excel file. By default, auto-applies and lints financial color conventions on exit.

```python
from headless_excel import create

# Create new workbook
with create("new.xlsx") as ctx:
    ctx.active["A1"] = "Hello"
    # auto-syncs on exit

# Overwrite existing file
with create("existing.xlsx", overwrite=True) as ctx:
    ctx.active["A1"] = "Fresh start"
```

### `run(path, auto_sync=True, verbose_errors=True, ...)`

Open an existing Excel file. Similar to OfficeJS `Excel.run()`. By default, auto-applies and lints financial color conventions on exit.

```python
from headless_excel import run

# Load existing, manual sync
with run("existing.xlsx", auto_sync=False) as ctx:
    ctx.active["A1"] = "=B1+C1"
    result = ctx.sync()
    print(ctx.active["A1"].value)

# Formula errors printed to stderr
with run("model.xlsx", verbose_errors=True) as ctx:
    ctx.active["A1"] = "=INVALID_REF"
    # Errors like #NAME? printed to stderr on exit
```

### `ExcelContext`

The context object provides:

| Property/Method      | Description                                |
| -------------------- | ------------------------------------------ |
| `workbook` / `wb`    | The workbook proxy for editing and reading |
| `active`             | Get/set active worksheet (assignable)      |
| `sheet(name)`        | Get worksheet by name                      |
| `create_sheet(name)` | Create new worksheet                       |
| `delete_sheet(name)` | Delete worksheet by name                   |
| `sync()`             | Save, recalc via LibreOffice, reload       |
| `values`             | Raw workbook with data_only=True           |

#### Sheet Management

```python
with create("model.xlsx") as ctx:
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
with create("model.xlsx") as ctx:
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

    # Get formula from individual cell
    ws["A1"].formula  # "=B1+C1"
    ws["B1"].formula  # None (not a formula)

    # Debug: dump range contents as formatted table (returns string)
    print(ws.range("A1:C3").dump())
    # |   A |   B |       C |
    # |-----|-----|---------|
    # |  10 |  20 |      30 |
    # |  30 |  40 |      70 |
    # |  50 |  60 |     100 |

    print(ws.range("A1:C3").dump(show_formulas=True))  # Show formulas instead of values

    # Count formulas
    ws.formula_count          # number of formulas in sheet
    ws.range("A1:C3").formula_count  # number of formulas in range
```

#### Auto Fill

Fill ranges like Excel's drag handle — formulas adjust their references automatically:

```python
with create("model.xlsx") as ctx:
    ws = ctx.active

    # Basic formula fill down (like dragging the fill handle)
    ws["A1"] = "=B1*C1"
    ws.range("A1:A10").auto_fill()  # A2:A10 get =B2*C2, =B3*C3, etc.

    # Fill right
    ws["A1"] = "=A2+A3"
    ws.range("A1:D1").auto_fill(direction="right")  # B1:D1 get =B2+B3, etc.

    # Fill up or left
    ws["A10"] = "=B10"
    ws.range("A1:A10").auto_fill(direction="up")  # Fills from bottom

    # Absolute references stay fixed, relative adjust
    ws["A1"] = "=$B$1*C1"
    ws.range("A1:A5").auto_fill()  # =$B$1*C2, =$B$1*C3, etc.

    # Multi-row source pattern (like selecting 2 rows and dragging)
    ws["A1"] = "Revenue"
    ws["A2"] = "=SUM(B1:D1)"
    ws.range("A1:A8").auto_fill(source_rows=2)  # Alternates pattern

    # Multi-column source pattern
    ws.range("A1:B1").values = [["Q1", "=A1+1"]]
    ws.range("A1:H1").auto_fill(direction="right", source_cols=2)

    # Style copying (enabled by default)
    ws["A1"].font = Font(bold=True)
    ws.range("A1:A5").auto_fill()  # All cells get bold
    ws.range("A1:A5").auto_fill(copy_styles=False)  # Values only
```

#### Range Styling

Apply styles to all cells in a range at once:

```python
from openpyxl.styles import Font, PatternFill, Alignment

with create("model.xlsx") as ctx:
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

### Error Scanning

Find formula errors at any level without triggering sync:

```python
with run("model.xlsx") as ctx:
    # After sync, scan for errors in materialized values
    errors = ctx.find_errors()                    # whole workbook
    errors = ctx.active.find_errors()             # single sheet
    errors = ctx.active.range("A1:D10").find_errors()  # specific range

    # ErrorScanResult has a nice __repr__ for agents
    print(errors)
    # Formula errors (3):
    #   #DIV/0!: Sheet1!A3, Sheet1!B5
    #   #REF!: Sheet1!C10
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

### Financial Color Automation

Automatically apply or lint financial modeling color conventions:

```python
from headless_excel import create, run, Colors, infer_financial_color
from openpyxl.styles import Font

# Infer the correct color for any value
infer_financial_color(100)           # Colors.HARDCODE (blue)
infer_financial_color("=A1+B1")      # Colors.FORMULA (black)
infer_financial_color("=Sheet2!A1")  # Colors.EXTERNAL_LINK (green)

# Manual auto-apply (for when you want control over timing)
# Note: disabling both to show manual usage
with create("model.xlsx", auto_financial_colors=False, lint_financial_colors=False) as ctx:
    ws = ctx.active
    ws["A1"] = 100
    ws["A2"] = "=A1*2"
    ws["A3"] = "=Sheet2!B1"

    # Apply to range
    ws.range("A1:A3").auto_financial_colors()

    # Or apply to entire sheet
    ws.auto_financial_colors()

    # Or apply to entire workbook
    ctx.auto_financial_colors()

# Lint: check for color violations (enabled by default, logs warnings)
with run("model.xlsx", lint_financial_colors=False) as ctx:
    # Check specific range
    violations = ctx.active.range("A1:D10").lint_financial_colors()

    # Check entire workbook - returns ColorLintResult
    result = ctx.lint_financial_colors()

    # ColorLintResult has a nice __repr__ for agents
    print(result)
    # Color violations (2):
    #   Sheet1! need HARDCODE (blue): A1, B2
    #   Sheet1! need FORMULA (black): C3

    # Dict-like access for backwards compatibility
    for sheet, sheet_violations in result.items():
        for v in sheet_violations:
            print(f"{sheet}!{v.cell}: expected {v.expected_color}")

# Financial color linting is enabled by default and logs warnings on violations
# To disable it:
with create("model.xlsx", lint_financial_colors=False) as ctx:
    ws = ctx.active
    ws["A1"] = 100  # No color check on exit
```

### Exceptions

- `ExcelError` - Base exception
- `FormulaError` - Formula errors found (has `.errors` dict)
- `RecalcError` - LibreOffice recalculation failed
- `SyncError` - Save or sync operation failed

### Hooks

Hooks allow you to run custom code at key points in the Excel workflow without modifying your main code. Common use cases: auto-formatting, validation, logging.

**Hook locations** (project-local takes precedence):

- `./.headless-excel/hooks/*.py` — project-local
- `~/.headless-excel/hooks/*.py` — global fallback

**Hook types:**

```python
# ~/.headless-excel/hooks/my_hooks.py
from headless_excel import ExcelContext, SyncResult, pre_sync, post_sync, on_exit

@pre_sync
def before_sync(ctx: ExcelContext) -> None:
    """Called before each sync() — modify workbook before save."""
    pass

@post_sync
def after_sync(ctx: ExcelContext, result: SyncResult) -> None:
    """Called after each sync() — inspect results, log errors."""
    pass

@on_exit
def on_context_exit(ctx: ExcelContext) -> None:
    """Called when context exits — final validation, cleanup."""
    pass
```

**Stateful hooks** using closures:

```python
from headless_excel import ExcelContext, ExtensionAPI, extension

@extension
def my_extension(api: ExtensionAPI):
    sync_count = 0

    @api.pre_sync
    def count_syncs(ctx: ExcelContext) -> None:
        nonlocal sync_count
        sync_count += 1
        print(f"Sync #{sync_count}")
```

See `examples/hooks/` for complete examples including financial color automation and audit logging.

## How It Works

1. **Edit**: Use standard openpyxl operations on `ctx.workbook`
2. **Sync**: Calls `ctx.sync()` which:
   - Saves the workbook to disk
   - Runs LibreOffice in headless mode with a macro to recalculate all formulas
   - Reloads the workbook twice: once for formulas, once with `data_only=True` for values
   - Scans for Excel error values (`#REF!`, `#NAME?`, `#VALUE!`, etc.)
3. **Read**: Access computed values via cell `.value` or `range.values`
