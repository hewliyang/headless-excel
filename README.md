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
    print(ctx.active["A1"].value)  # or ctx.read_value("Sheet1", "A1")

# Raise on formula errors
with run("model.xlsx", raise_on_errors=True) as ctx:
    ctx.active["A1"] = "=INVALID_REF"
    # Raises FormulaError on exit
```

### `ExcelContext`

The context object provides:

| Property/Method               | Description                                 |
| ----------------------------- | ------------------------------------------- |
| `workbook` / `wb`             | The workbook proxy for editing and reading  |
| `active`                      | The active worksheet proxy                  |
| `sheet(name)`                 | Get worksheet by name                       |
| `create_sheet(name)`          | Create new worksheet                        |
| `sync(raise_on_errors=False)` | Save, recalc via LibreOffice, reload        |
| `read_value(sheet, cell)`     | Helper to read computed value               |
| `read_range(sheet, range)`    | Helper to read computed values from range   |
| `values`                      | (Advanced) Raw workbook with data_only=True |

### `SyncResult`

Returned by `sync()`:

```python
result = ctx.sync()
result.success      # bool - True if no formula errors
result.total_errors # int - count of errors
result.errors       # dict - {'#REF!': ['Sheet1!A1'], ...}
result.raise_on_errors()  # raise FormulaError if errors exist
```

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
3. **Read**: Access computed values via `ctx.values` or `ctx.read_value()`

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
        actual = ctx.read_value("Sheet1", "A1")
        if actual != expected:
            # ... adjust formula
```
