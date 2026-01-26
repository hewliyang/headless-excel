# API Reference

This document covers the custom APIs provided by headless-excel. All workbook and worksheet objects are proxies over `openpyxl`. If a method doesn't exist on the proxy, it forwards to the underlying `openpyxl` object.

## Table of Contents

- [Context API](#context-api)
  - [run() / create()](#run--create)
  - [ExcelContext](#excelcontext)
  - [SyncResult](#syncresult)
- [Worksheet API](#worksheet-api)
  - [write()](#write)
  - [range()](#range)
  - [formulas / formula_count](#formulas--formula_count)
  - [find_errors()](#find_errors)
- [Range API](#range-api)
  - [values](#values)
  - [auto_fill()](#auto_fill)
  - [dump()](#dump)
  - [clear()](#clear)
  - [apply_style()](#apply_style)
  - [formulas / formula_count](#formulas--formula_count-1)
  - [find_errors()](#find_errors-1)
- [Cell API](#cell-api)
  - [value](#value)
  - [formula](#formula)
- [NumberFormats](#numberformats)
- [Hooks](#hooks)

---

## Context API

### run() / create()

Entry points for working with Excel files.

```python
from headless_excel import run, create

# Open existing file
with run("model.xlsx") as ctx:
    ws = ctx.active
    ws["A1"] = 100
    # auto-syncs on exit

# Create new file
with create("new.xlsx") as ctx:
    ws = ctx.active
    ws["A1"] = "Hello"

# Overwrite existing file
with create("existing.xlsx", overwrite=True) as ctx:
    ...
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `path` | `str \| Path` | required | Path to Excel file |
| `auto_sync` | `bool` | `True` | Automatically sync on context exit |
| `verbose_errors` | `bool` | `True` | Print formula errors to stderr on exit |
| `recalc_timeout` | `int` | `30` | Timeout in seconds for LibreOffice recalculation |
| `overwrite` | `bool` | `False` | (create only) Overwrite existing file |

### ExcelContext

The context object provides access to workbooks and worksheets.

```python
with run("model.xlsx") as ctx:
    # Access worksheets
    ws = ctx.active              # Active worksheet
    ws = ctx.sheet("Revenue")    # By name
    ws = ctx.workbook["Revenue"] # Also by name

    # Create/delete sheets
    ws = ctx.create_sheet("NewSheet")
    ctx.delete_sheet("OldSheet")

    # Set active sheet
    ctx.active = "Revenue"
    ctx.active = ctx.sheet("Revenue")

    # Sync: save → recalculate → reload
    result = ctx.sync()
    if not result.success:
        print(result.errors)

    # Find errors across all sheets
    errors = ctx.find_errors()
```

**Note:** When creating a new workbook with `create()`, the default empty "Sheet" is automatically removed when you call `create_sheet()` for the first time.

### SyncResult

Returned by `ctx.sync()`. Contains information about formula errors.

```python
result = ctx.sync()

result.success       # bool - True if no formula errors
result.total_errors  # int - Count of errors
result.errors        # dict[str, list[str]] - Errors by type
                     # e.g., {"#DIV/0!": ["Sheet1!A3", "Sheet1!B5"]}

# Raise exception if errors found
result.raise_on_errors()

# Or pass flag to sync()
ctx.sync(raise_on_errors=True)
```

---

## Worksheet API

### write()

Bulk write 2D data starting at an anchor cell. The simplest API for populating data—no range math required.

```python
ws.write("A1", [
    ["Product", "Q1", "Q2", "Total"],
    ["Widget", 100, 150, "=B2+C2"],
    ["Gadget", 200, 250, "=B3+C3"],
])
# Prints: [headless-excel] .write() wrote to A1:D3 (3×4)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `cell` | `str` | Anchor cell (top-left of write region), e.g., `"A1"` |
| `data` | `list[list[Any]]` | 2D list of values to write |

**Returns:** `str` - The actual range written (e.g., `"A1:D3"`)

**Feedback:** Prints to stderr showing actual write region.

### range()

Get a `RangeProxy` for bulk operations on a rectangular region.

```python
r = ws.range("A1:D10")
r = ws.range("A1")  # Single cell also works
```

### formulas / formula_count

Get all formulas in the sheet or count them.

```python
# Dict of coordinate → formula
ws.formulas
# {'A1': '=B1+C1', 'D5': '=SUM(A1:A4)'}

# Just the count
ws.formula_count
# 42
```

### find_errors()

Scan the sheet for formula errors after sync.

```python
ctx.sync()
result = ws.find_errors()

print(result)
# Formula errors (3):
#   #DIV/0!: A3, B5
#   #REF!: C10

result.total_errors   # 3
result.errors_by_type # {"#DIV/0!": ["A3", "B5"], "#REF!": ["C10"]}
```

---

## Range API

### values

Get or set 2D array of values.

```python
# Read values
data = ws.range("A1:C3").values
# [[100, 200, 300], [10, 20, 30], [1, 2, 3]]

# Write values
ws.range("A1:C3").values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
]
```

**Lenient behavior:** The range acts as an anchor—data shape determines the actual write region. This prevents off-by-one errors when specifying ranges.

```python
# Range says A1:D10 but data is 3×4
ws.range("A1:D10").values = [[1,2,3,4], [5,6,7,8], [9,10,11,12]]
# Prints: [headless-excel] .values wrote to A1:D3 (specified A1:D10, got 3×4)
```

### auto_fill()

Fill the range by repeating source cells, adjusting formula references. Mimics Excel's drag-handle behavior.

```python
# Fill down (default for tall ranges)
ws["A1"] = "=B1*C1"
ws.range("A1:A10").auto_fill()
# A2:A10 now have =B2*C2, =B3*C3, etc.

# Fill right
ws["A1"] = "=A2+A3"
ws.range("A1:D1").auto_fill(direction="right")
# B1:D1 now have =B2+B3, =C2+C3, =D2+D3

# Multi-row pattern (like alternating labels)
ws["A1"] = "Revenue"
ws["A2"] = "Expenses"
ws.range("A1:A8").auto_fill(source_rows=2)
# Repeats: Revenue, Expenses, Revenue, Expenses...
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `direction` | `"down" \| "right" \| "up" \| "left" \| None` | `None` | Fill direction. `None` auto-detects (tall=down, wide=right) |
| `source_rows` | `int` | `1` | Number of source rows to repeat (for down/up) |
| `source_cols` | `int` | `1` | Number of source columns to repeat (for left/right) |
| `copy_styles` | `bool` | `True` | Copy cell styles along with values |

**Formula translation:**
- Relative references adjust: `=A1` → `=A2` → `=A3`
- Absolute references stay fixed: `=$A$1` stays `=$A$1`
- Mixed references adjust partially: `=$A1` → `=$A2`, `=A$1` stays `=A$1`

### dump()

Print range contents as a formatted table for debugging.

```python
print(ws.range("A1:C3").dump())
# |     A |     B |     C |
# |-------|-------|-------|
# |   100 |   200 |   300 |
# |    10 |    20 |    30 |
# |     1 |     2 |     3 |

# Show formulas instead of values
print(ws.range("A1:C3").dump(show_formulas=True))
# |       A |       B |         C |
# |---------|---------|-----------|
# |     100 |     200 | =A1+B1    |
# |      10 |      20 | =A2+B2    |
# |       1 |       2 | =SUM(A:B) |
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `show_formulas` | `bool` | `False` | Show formulas instead of computed values |

### clear()

Clear cell values and optionally styles.

```python
ws.range("A1:C3").clear()              # Clear values and styles
ws.range("A1:C3").clear(styles=False)  # Clear values only
```

**Note:** Automatically unmerges any merged cells that overlap with the clear range.

### apply_style()

Apply styles to all cells in the range.

```python
from openpyxl.styles import Font, PatternFill, Alignment, Border

ws.range("A1:D1").apply_style(
    font=Font(bold=True, color="FFFFFF"),
    fill=PatternFill("solid", fgColor="4472C4"),
    alignment=Alignment(horizontal="center"),
)

# Number formatting
from headless_excel import NumberFormats
ws.range("B2:B10").apply_style(number_format=NumberFormats.ACCOUNTING)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `font` | `Font` | Font style |
| `fill` | `PatternFill` | Background fill |
| `gradient_fill` | `GradientFill` | Gradient fill (takes precedence over `fill`) |
| `alignment` | `Alignment` | Cell alignment |
| `border` | `Border` | Border style |
| `protection` | `Protection` | Protection settings |
| `number_format` | `str` | Number format string |

### formulas / formula_count

Same as worksheet-level, but scoped to the range.

```python
ws.range("A1:C10").formulas
# {'B2': '=A1*2', 'C3': '=SUM(A1:B2)'}

ws.range("A1:C10").formula_count
# 5
```

### find_errors()

Same as worksheet-level, but scoped to the range.

```python
ws.range("A1:C10").find_errors()
```

### Shape properties

```python
r = ws.range("A1:D10")
r.num_rows  # 10
r.num_cols  # 4
r.shape     # (10, 4)
```

---

## Cell API

Cells are wrapped in `CellProxy` which provides automatic value materialization.

### value

Returns the computed value after sync, or the formula before sync.

```python
ws["A1"] = "=SUM(B1:B10)"
ctx.sync()
ws["A1"].value  # 450 (computed result, not the formula)
```

### formula

Returns the formula string if the cell contains a formula, otherwise `None`.

```python
ws["A1"] = "=SUM(B1:B10)"
ws["A1"].formula  # "=SUM(B1:B10)"

ws["A2"] = 42
ws["A2"].formula  # None
```

---

## NumberFormats

Pre-defined number format strings for common use cases.

```python
from headless_excel import NumberFormats

ws["A1"].number_format = NumberFormats.ACCOUNTING      # $1,234.56 with aligned decimals
ws["A2"].number_format = NumberFormats.ACCOUNTING_0DP  # $1,235 (no decimals)
ws["A3"].number_format = NumberFormats.PERCENTAGE      # 75%
ws["A4"].number_format = NumberFormats.PERCENTAGE_2DP  # 75.00%
ws["A5"].number_format = NumberFormats.NUMBER          # 1,234.56
ws["A6"].number_format = NumberFormats.NUMBER_0DP      # 1,235
ws["A7"].number_format = NumberFormats.DATE            # 25-Jan-2024
ws["A8"].number_format = NumberFormats.DATE_LONG       # January 25, 2024
```

Or apply to ranges:

```python
ws.range("B2:B100").apply_style(number_format=NumberFormats.ACCOUNTING)
```

---

## Hooks

Hooks allow custom code to run at specific points in the workflow. See [hooks.md](hooks.md) for full documentation.

```python
# ~/.headless-excel/hooks/my_hooks.py
from headless_excel import pre_sync, post_sync, on_exit

@pre_sync
def before_sync(ctx):
    print(f"About to sync {ctx.path}")

@post_sync
def after_sync(ctx, result):
    print(f"Sync complete, errors: {result.total_errors}")

@on_exit
def cleanup(ctx):
    print("Done!")
```

---

## Error Handling

### ErrorScanResult

Returned by `find_errors()` methods. Provides a truncated `__repr__` to avoid polluting LLM context.

```python
result = ws.find_errors()

bool(result)          # True if errors exist
len(result)           # Number of error types
"#DIV/0!" in result   # Check for specific error
result["#DIV/0!"]     # List of locations for that error
result.total_errors   # Total count
result.errors_by_type # Full dict

# Iteration
for error_type in result:
    print(error_type, result[error_type])

for error_type, locations in result.items():
    print(error_type, locations)
```

### Exception Types

```python
from headless_excel import (
    ExcelError,           # Base exception
    SyncError,            # Sync operation failed
    RecalcError,          # LibreOffice recalculation failed
    FormulaError,         # Formula errors found (when raise_on_errors=True)
    LibreOfficeNotFoundError,  # LibreOffice not installed
)
```
