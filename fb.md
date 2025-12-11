Issues I Hit 🤔

### 1. Post-sync value access is confusing

After calling ctx.sync(), I expected ws['A1'].value to return the calculated value, but it still showed the formula string. I had to use ctx.read_value()  
 instead.

```python
  ws['A1'] = '=SUM(B1:B10)'
  ctx.sync()
  # Expected this to work:
  # result = ws['A1'].value  # Still shows '=SUM(B1:B10)'
  # Had to do this:
  result = ctx.read_value(sheet_name, 'A1')  # Gets calculated value
```

Suggestion: After sync(), automatically reload computed values into the worksheet object so direct cell access works intuitively.

### 2. No explicit refresh without recalc

Sometimes I just want to reload values without triggering LibreOffice recalculation (e.g., after manually editing the file externally).

Suggestion: Add a refresh() or reload() method separate from sync().

### 3. Batch styling is verbose

Setting styles on ranges requires looping:

```python
  for cell in ['B5', 'C5', 'D5', 'E5', 'F5', 'G5']:
      ws[cell].font = Font(color='008000')
      ws[cell].number_format = '#,##0;(#,##0);-'
```

Suggestion: Add batch styling helpers:

```python
  ctx.apply_style('A1:G5', font=Font(color='008000'),
                  number_format='#,##0;(#,##0);-')
```

Nice-to-Haves 🎁

### 4. Formula validation without save

```python
  result = ctx.validate()  # Check formulas without saving
  if result.has_errors:
      # Fix issues before committing
```

### 5. Better error context

When formulas fail, show:

- The actual formula text
- The cell reference
- What type of error (#DIV/0!, #REF!, etc.)
- Neighboring cell values (for debugging)

Current: {'#DIV/0!': ['Sheet1!A1']}  
 Better: {'Sheet1!A1': {'error': '#DIV/0!', 'formula': '=B1/C1', 'B1': 100, 'C1': 0}}

### 6. Helper for financial modeling patterns

```python
  # Common pattern I used repeatedly
  ctx.project_line_item('Revenue', 'Assumptions!G7')  # Growth rate
  ctx.project_as_percent('Cash', 'Revenue', 'Assumptions!G17')
```

### 7. Inspection tools

```python
  formulas = ctx.get_formulas('Sheet1')  # Get all formulas in sheet
  dependencies = ctx.trace_dependents('A1')  # What depends on this cell?
```
