---
name: xlsx-uno
description: "Creating & editing Excel workbooks via UNO (LibreOffice API)"
---

# headless-excel UNO Mode

Direct LibreOffice API access via Python. Faster than openpyxl mode, with full LibreOffice feature support.

## Commands

```bash
headless-excel uno_run file.xlsx "code"      # run UNO code with helpers + hooks
headless-excel uno_run --no-hooks file.xlsx "code"  # skip hooks
headless-excel uno_eval file.xlsx "code"     # raw UNO (no helpers)
```

## Helpers Available

In `uno_run`, these helpers are pre-loaded:

```python
cell("A1")           # -> (col, row) 0-indexed tuple
rng("A1:C3")         # -> (c1, r1, c2, r2) 0-indexed tuple
sheet(0)             # -> XSpreadsheet by index
sheet("Data")        # -> XSpreadsheet by name
sheets()             # -> list of sheet names
recalc()             # -> doc.calculateAll()
doc                  # -> XSpreadsheetDocument
```

## Workflow

1. Open/create file automatically
2. Write UNO code with helpers
3. Set `result` variable to return values
4. Auto-saves on exit

## Write

```bash
headless-excel uno_run --no-hooks model.xlsx '
s = sheet(0)
c, r = cell("A1")
s.getCellByPosition(c, r).setString("Revenue")

c, r = cell("A2")
s.getCellByPosition(c, r).setValue(100)

c, r = cell("A3")
s.getCellByPosition(c, r).setFormula("=SUM(A2:A2)")
'
```

## Read

```bash
headless-excel uno_run --no-hooks model.xlsx '
recalc()
s = sheet(0)
c, r = cell("A3")
result = s.getCellByPosition(c, r).getValue()
'
```

## Bulk Operations

```bash
headless-excel uno_run --no-hooks model.xlsx '
s = sheet(0)

# Write 2D array
data = [
    ["Product", "Q1", "Q2", "Total"],
    ["Widget", 100, 150, "=SUM(B2:C2)"],
    ["Gadget", 200, 250, "=SUM(B3:C3)"],
]

# Use setDataArray for values (fast)
c1, r1, c2, r2 = rng("A1:C3")
value_data = [row[:3] for row in data]  # exclude formulas
s.getCellRangeByPosition(c1, r1, c2, r2-1).setDataArray(tuple(tuple(r) for r in value_data))

# Set formulas individually
for i, row in enumerate(data[1:], start=1):
    s.getCellByPosition(3, i).setFormula(row[3])

recalc()
result = s.getCellRangeByPosition(0, 0, 3, 2).getDataArray()
'
```

## Multi-Sheet Workbooks

```bash
headless-excel uno_run --no-hooks model.xlsx '
all_sheets = doc.getSheets()

# Rename first sheet
all_sheets.getByIndex(0).setName("Data")

# Create new sheet
all_sheets.insertNewByName("Summary", 1)

# Write to sheets
data = sheet("Data")
data.getCellByPosition(0, 0).setValue(100)

summary = sheet("Summary")
summary.getCellByPosition(0, 0).setFormula("=SUM(Data.A:A)")

# Reorder sheets
all_sheets.moveByName("Summary", 0)  # move to first position

recalc()
result = sheets()
'
```

## Styling

```bash
headless-excel uno_run --no-hooks model.xlsx '
s = sheet(0)

# Bold
c, r = cell("A1")
s.getCellByPosition(c, r).setPropertyValue("CharWeight", 150)

# Font color (RGB as integer)
s.getCellByPosition(c, r).setPropertyValue("CharColor", 0x0000FF)  # blue

# Background color
s.getCellByPosition(c, r).setPropertyValue("CellBackColor", 0xFFFF00)  # yellow

# Number format
s.getCellByPosition(1, 1).setPropertyValue("NumberFormat", 4)  # currency

# Range styling
c1, r1, c2, r2 = rng("A1:D1")
header = s.getCellRangeByPosition(c1, r1, c2, r2)
header.setPropertyValue("CharWeight", 150)
header.setPropertyValue("CellBackColor", 0x1F4E79)
header.setPropertyValue("CharColor", 0xFFFFFF)
'
```

## View Settings

```bash
headless-excel uno_run --no-hooks model.xlsx '
controller = doc.getCurrentController()

# Hide gridlines
controller.setPropertyValue("ShowGrid", False)

# Freeze panes (freeze row 1)
controller.freezeAtPosition(0, 1)

# Set column width
s = sheet(0)
s.getColumns().getByIndex(0).setPropertyValue("Width", 3500)  # ~1 inch
'
```

## Conditional Formatting

```bash
headless-excel uno_run --no-hooks model.xlsx '
from com.sun.star.sheet.ConditionOperator import LESS, GREATER
from com.sun.star.beans import PropertyValue

s = sheet(0)
c1, r1, c2, r2 = rng("A1:A100")
target = s.getCellRangeByPosition(c1, r1, c2, r2)

cond = target.getPropertyValue("ConditionalFormat")

# Red if < 0
cond.addNew((
    PropertyValue("Operator", 0, LESS, 0),
    PropertyValue("Formula1", 0, "0", 0),
    PropertyValue("StyleName", 0, "Bad", 0),
))

# Green if > 1000
cond.addNew((
    PropertyValue("Operator", 0, GREATER, 0),
    PropertyValue("Formula1", 0, "1000", 0),
    PropertyValue("StyleName", 0, "Good", 0),
))

target.setPropertyValue("ConditionalFormat", cond)
'
```

## Best Practices

Use formulas, not Python-computed values:

```python
# Bad
total = sum(values)
s.getCellByPosition(0, 9).setValue(total)

# Good
s.getCellByPosition(0, 9).setFormula("=SUM(A1:A9)")
```

## UNO Quick Reference

```python
# Document
doc.getSheets()                          # XSpreadsheets collection
doc.getSheets().getCount()               # number of sheets
doc.getSheets().getByIndex(0)            # XSpreadsheet by index
doc.getSheets().getByName("Sheet1")      # XSpreadsheet by name
doc.getSheets().insertNewByName(name, pos)  # create sheet
doc.getSheets().removeByName(name)       # delete sheet
doc.getSheets().moveByName(name, pos)    # reorder sheet
doc.calculateAll()                       # recalculate

# Sheet (XSpreadsheet)
sheet.getName()                          # sheet name
sheet.setName(name)                      # rename
sheet.getCellByPosition(col, row)        # XCell (0-indexed!)
sheet.getCellRangeByPosition(c1, r1, c2, r2)  # XCellRange
sheet.getColumns().getByIndex(i)         # column object
sheet.getRows().getByIndex(i)            # row object

# Cell (XCell)
cell.getValue()                          # numeric value
cell.getString()                         # string value
cell.getFormula()                        # formula string
cell.setValue(num)                       # set number
cell.setString(str)                      # set string
cell.setFormula("=A1+B1")               # set formula
cell.getType()                           # CellContentType enum
cell.getError()                          # error code (0 = no error)

# Range (XCellRange)
range.getDataArray()                     # 2D tuple of values
range.setDataArray(tuple)                # set 2D data (fast!)
range.getFormulaArray()                  # 2D tuple of formulas

# Properties (XPropertySet) - cells/ranges support this
obj.setPropertyValue("CharWeight", 150)  # bold
obj.setPropertyValue("CharColor", 0xFF0000)  # font color
obj.setPropertyValue("CellBackColor", 0x00FF00)  # background
obj.setPropertyValue("NumberFormat", 4)  # number format
obj.getPropertyValue("PropertyName")     # get property

# Controller
controller = doc.getCurrentController()
controller.setActiveSheet(sheet)         # activate sheet
controller.setPropertyValue("ShowGrid", False)  # hide gridlines
controller.freezeAtPosition(col, row)    # freeze panes
```

## Tips

- All positions are **0-indexed** (A1 = col 0, row 0)
- Use `cell("A1")` helper to convert A1 notation
- Use `setDataArray()` for bulk writes (much faster than cell-by-cell)
- Set `result = ...` to return values from the script
- Call `recalc()` before reading computed values
- LibreOffice uses `;` as formula separator (auto-converted from `,`)
