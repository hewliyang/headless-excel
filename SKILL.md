---
name: xlsx
description: "Comprehensive spreadsheet creation, editing, and analysis with support for formulas, formatting, data analysis, and visualization. When Claude needs to work with spreadsheets (.xlsx, .xlsm, .csv, .tsv, etc) for: (1) Creating new spreadsheets with formulas and formatting, (2) Reading or analyzing data, (3) Modify existing spreadsheets while preserving formulas, (4) Data analysis and visualization in spreadsheets, or (5) Recalculating formulas"
---

# Requirements for Outputs

## All Excel files

### Zero Formula Errors

- Every Excel model MUST be delivered with ZERO formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)

### Preserve Existing Templates (when updating templates)

- Study and EXACTLY match existing format, style, and conventions when modifying files
- Never impose standardized formatting on files with established patterns
- Existing template conventions ALWAYS override these guidelines

## Financial models

### Color Coding Standards

Unless otherwise stated by the user or existing template

#### Industry-Standard Color Conventions

- **Blue text (RGB: 0,0,255)**: Hardcoded inputs, and numbers users will change for scenarios
- **Black text (RGB: 0,0,0)**: ALL formulas and calculations
- **Green text (RGB: 0,128,0)**: Links pulling from other worksheets within same workbook
- **Red text (RGB: 255,0,0)**: External links to other files
- **Yellow background (RGB: 255,255,0)**: Key assumptions needing attention or cells that need to be updated

### Number Formatting Standards

#### Required Format Rules

- **Years**: Format as text strings (e.g., "2024" not "2,024")
- **Currency**: Use $#,##0 format; ALWAYS specify units in headers ("Revenue ($mm)")
- **Zeros**: Use number formatting to make all zeros "-", including percentages (e.g., "$#,##0;($#,##0);-")
- **Percentages**: Default to 0.0% format (one decimal)
- **Multiples**: Format as 0.0x for valuation multiples (EV/EBITDA, P/E)
- **Negative numbers**: Use parentheses (123) not minus -123

### Formula Construction Rules

#### Assumptions Placement

- Place ALL assumptions (growth rates, margins, multiples, etc.) in separate assumption cells
- Use cell references instead of hardcoded values in formulas
- Example: Use =B5*(1+$B$6) instead of =B5*1.05

#### Formula Error Prevention

- Verify all cell references are correct
- Check for off-by-one errors in ranges
- Ensure consistent formulas across all projection periods
- Test with edge cases (zero values, negative numbers)
- Verify no unintended circular references

#### Documentation Requirements for Hardcodes

- Comment or in cells beside (if end of table). Format: "Source: [System/Document], [Date], [Specific Reference], [URL if applicable]"
- Examples:
  - "Source: Company 10-K, FY2024, Page 45, Revenue Note, [SEC EDGAR URL]"
  - "Source: Company 10-Q, Q2 2025, Exhibit 99.1, [SEC EDGAR URL]"
  - "Source: Bloomberg Terminal, 8/15/2025, AAPL US Equity"
  - "Source: FactSet, 8/20/2025, Consensus Estimates Screen"

# XLSX creation, editing, and analysis

## Overview

Use `headless-excel` for all Excel operations. It provides automatic formula recalculation and error detection via `libreoffice-calc` and is essentially a wrapper around `openpyxl` plus some utility methods to make life easier.

## CRITICAL: Use Formulas, Not Hardcoded Values

**Always use Excel formulas instead of calculating values in Python and hardcoding them.** This ensures the spreadsheet remains dynamic and updateable.

### ❌ WRONG - Hardcoding Calculated Values

```python
# Bad: Calculating in Python and hardcoding result
total = df['Sales'].sum()
ws['B10'] = total  # Hardcodes 5000

# Bad: Computing growth rate in Python
growth = (df.iloc[-1]['Revenue'] - df.iloc[0]['Revenue']) / df.iloc[0]['Revenue']
ws['C5'] = growth  # Hardcodes 0.15
```

### ✅ CORRECT - Using Excel Formulas

```python
# Good: Let Excel calculate the sum
ws['B10'] = '=SUM(B2:B9)'

# Good: Growth rate as Excel formula
ws['C5'] = '=(C4-C2)/C2'
```

## Basic Usage

CRITICAL:

- Keep your heredoc's under 100 LoC to avoid compounding errors you cannot recover from
- Always use `auto_sync`, or if not the remember to call `ctx.sync()` as frequently as possible

### Creating new Excel files

```bash
uv run python <<'EOF'
from headless_excel import run
from openpyxl.styles import Font, PatternFill, Alignment

with run("output.xlsx", create=True) as ctx:
    ws = ctx.active

    # Add data and formulas
    ws['A1'] = 'Revenue'
    ws['A2'] = 100
    ws['A3'] = 200
    ws['A4'] = '=SUM(A2:A3)'

    # Cell-by-cell formatting
    ws['A1'].font = Font(bold=True, color='0000FF')
    ws['A1'].fill = PatternFill('solid', start_color='FFFF00')

    # Batch styling for ranges
    ctx.apply_style("Sheet", "A2:A4", font=Font(color='000000'), number_format='#,##0')

    # Column width
    ws.column_dimensions['A'].width = 20

    # Auto-syncs on exit (saves, recalculates, checks for errors)
EOF
```

### Editing existing Excel files

```bash
uv run python <<'EOF'
from headless_excel import run

with run("existing.xlsx") as ctx:
    ws = ctx.active

    # Modify cells
    ws['A1'] = 'New Value'
    ws['B2'] = '=A1*2'

    # Work with specific sheets
    revenue_sheet = ctx.sheet("Revenue")
    revenue_sheet['C1'] = '=SUM(A1:B1)'

    # Create new sheet
    new_sheet = ctx.create_sheet("Assumptions")
    new_sheet['A1'] = 'Growth Rate'
    new_sheet['B1'] = 0.15

    # Auto-syncs on exit
EOF
```

### Reading computed values

```bash
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx") as ctx:
    # Read single value
    revenue = ctx.read_value("Revenue", "A1")

    # Read range
    totals = ctx.read_range("Summary", "A1:D10")

    # Access directly after sync
    ctx.active['A1'] = '=SUM(B1:B10)'
    ctx.sync()
    result = ctx.active['A1'].value
EOF
```

### Inspecting formulas

Use `get_formulas()` to audit and debug formula logic without calculating values:

```bash
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx") as ctx:
    # Get all formulas in a specific sheet
    formulas = ctx.get_formulas("Revenue")
    # Returns: {'A3': '=A1+A2', 'B3': '=B1*B2', 'C3': '=SUM(A3:B3)'}

    # Get all formulas in entire workbook
    all_formulas = ctx.get_formulas()
    # Returns: {'Revenue!A3': '=A1+A2', 'Summary!B5': '=Revenue!C3*1.15', ...}

    # Common use cases:
    # 1. Audit formula logic before sync
    for cell, formula in formulas.items():
        if 'HARDCODED_VALUE' in formula:
            print(f"Warning: {cell} has hardcoded value")

    # 2. Find cells referencing a specific cell
    for cell, formula in formulas.items():
        if 'B1' in formula:
            print(f"{cell} references B1: {formula}")

    # 3. Validate formula consistency
    if formulas.get('C3') != formulas.get('C4'):
        print("Warning: Inconsistent formulas in C3 and C4")
EOF
```

## Error Handling

`headless-excel` automatically detects formula errors:

```bash
# Automatic error checking with auto_sync
uv run python <<'EOF'
from headless_excel import run, FormulaError

try:
    with run("model.xlsx", create=True) as ctx:
        ctx.active['A1'] = '=INVALID_REF'
        # Exits with sync, detects error automatically
except FormulaError as e:
    print(f"Formula errors found: {e.errors}")
    # {'#NAME?': ['Sheet1!A1']}
EOF

# Manual error checking with detailed context
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx", auto_sync=False) as ctx:
    ws = ctx.active
    ws['A1'] = '=B1/C1'  # Potential #DIV/0!

    result = ctx.sync()
    if not result.success:
        for detail in result.error_details:
            print(f"{detail.error} at {detail.location}: {detail.formula}")
            print(f"  Referenced values: {detail.neighbors}")
        # Fix errors...
        ws['C1'] = 1
        result = ctx.sync()
EOF

# Raise on errors
uv run python <<'EOF'
from headless_excel import run

with run("model.xlsx", raise_on_errors=True) as ctx:
    ctx.active['A1'] = '=InvalidFormula'
    # Raises FormulaError on exit
EOF
```

## Data Analysis with pandas

For bulk data operations and analysis, use pandas:

```bash
uv run python <<'EOF'
import pandas as pd

# Read Excel
df = pd.read_excel('file.xlsx')  # First sheet
all_sheets = pd.read_excel('file.xlsx', sheet_name=None)  # All sheets

# Analyze
df.head()
df.info()
df.describe()

# Export to Excel for further formula work
df.to_excel('output.xlsx', index=False)
EOF
```

Then use `headless-excel` to add formulas:

```bash
uv run python <<'EOF'
from headless_excel import run

with run("output.xlsx") as ctx:
    ws = ctx.active
    last_row = len(df) + 2  # +1 for header, +1 for next row
    ws[f'A{last_row}'] = '=SUM(A2:A{last_row-1})'
EOF
```

## Best Practices

### Library Selection

- **headless-excel**: For all Excel creation, editing, and formula work
- **pandas**: For data analysis, transformations, and bulk data operations

### Working with headless-excel

- Formulas are automatically recalculated via LibreOffice
- Formula errors are automatically detected on sync
- Use `auto_sync=False` for manual control over recalculation timing
- Use `raise_on_errors=True` to enforce zero-error requirement
- Use `ctx.get_formulas("Sheet1")` to inspect all formulas in a sheet
- Use `ctx.apply_style("Sheet1", "A1:G5", ...)` for batch styling
- Use A1 notation for cell references (e.g., `ws['A1']`, `ws['B5']`), not numeric indices
- Build worksheets incrementally: prefer heredocs for multi-step cell setup, keep each heredoc under 100 lines, and call `ctx.sync()` after each block to catch errors before they compound

### Common Pitfalls to Avoid

- [ ] **Column mapping**: Confirm Excel columns (e.g., column 64 = BL, not BK)
- [ ] **Row offset**: Excel rows are 1-indexed (DataFrame row 5 = Excel row 6)
- [ ] **Division by zero**: Check denominators to avoid #DIV/0!
- [ ] **Invalid references**: Verify cell references point to valid cells (#REF!)
- [ ] **Cross-sheet references**: Use correct format (Sheet1!A1) for linking sheets

## Code Style Guidelines

When generating Python code for Excel operations:

- Write minimal, concise Python code without unnecessary comments
- Avoid verbose variable names and redundant operations
- Avoid unnecessary print statements

For Excel files themselves:

- Add comments to cells with complex formulas or important assumptions
- Document data sources for hardcoded values
- Include notes for key calculations and model sections
