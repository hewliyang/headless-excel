- [ ] Inspection tools

```python
  dependencies = ctx.trace_dependents('A1')  # What depends on this cell?
```

> this seems useful - need to figure out how to build the deps graph. in excel - can be circular

- [ ] Improve default **repr** for FormulaError, show more details about the error. Make it more similar to what you get out of ctx.sync() (if this returns an error result)

- [ ] Investigate if `recalc` can somehow save the cached values to the file instead of having to recalc every time

  - ps: it is. its just that we don't populate `_workbook_values` on enter of the context manager!

# Ranges + setters via 2D arrays

Helps w/ off by one errors via for-looping
Perhaps can support styling as well (refactor `ctx.apply_style` to apply on range instead of globally to keep things consistent.)

```py
B3E6_R = sheet.range("B3:E6")
B3E6_R.set_values([[a1,b1,c1,d1],[a2,b2,c2,d2], ...]) # recommended by clawd
# OR
B3E6_R.values=[[...], [...], ...] # OfficeJS API
```

# Named range utilities

```py
ctx.define_name("GrowthRate", "Assumptions!B6")
ctx.get_named_ranges()
```

# Sheet management utilities

```py
ctx.active_sheet("{{ sheet name }}")
# OR
ctx.active = ctx.sheet("{{ sheet_name }}") # stick with properties or methods?!

ctx.delete_sheet("{{ sheet name }}")
```

# Improve documentation

1. show >1 cross-sheet example
2. sheet creation order not documented
3. emphasize we are just wrapping openpyxl and ws = ctx.active; ws is just a `openpyxl.Worksheet`
