# Pivot Tables

Declare a pivot with `ws.add_pivot(...)`. Pivots are **declared, not computed in Python** — the source cache is flagged `refreshOnLoad`, so `ctx.sync()` (LibreOffice) or Excel materialises every value and grand total. You must `ctx.sync()` to see results.

Fluent form:

```bash
headless-excel eval model.xlsx "
data = ctx.sheet('Sheet'); data.title = 'Data'
data.range('A1:D7').values = [
    ['Region', 'Product', 'Units', 'Sales'],
    ['North', 'A', 10, 100], ['North', 'B', 5, 80],
    ['South', 'A', 7, 70],  ['South', 'B', 3, 60],
    ['North', 'A', 4, 40],  ['South', 'B', 9, 90],
]

pivot = ctx.create_sheet('Pivot')
(pivot.add_pivot('Data!A1:D7', anchor='A3', name='ByRegion')
      .rows('Region')
      .columns('Product')
      .values('Sales', func='sum', num_format='\$#,##0')
      .style('PivotStyleMedium9'))

ctx.sync()  # LibreOffice computes the aggregations
print(pivot.range('A3:E8').dump())
"
```

Declarative form (equivalent):

```python
pivot.add_pivot(
    'Data!A1:D7', anchor='A3', name='ByRegion',
    rows=['Region'], columns=['Product'],
    values=[('Sales', 'sum', 'Total Sales', '$#,##0')],
    style='PivotStyleMedium9',
)
```

`filters=[...]` adds report-filter (page) fields when your source has them. All field names must match the source header row.

Aggregation functions: `sum, count, countNums, average, max, min, product, stdDev, stdDevp, var, varp` (aliases like `avg`/`mean` accepted).

LibreOffice may ignore a value field's number format. To force it onto the rendered cells, call `.format_values(ctx)` after building (it syncs and styles the materialised value region).

## Inspect, edit & delete pivots

Pivots support full CRUD — including pivots loaded from an existing file. `ws.pivots` and `ws.get_pivot(name)` return `PivotTable` wrappers you can read, mutate, or delete:

```bash
headless-excel eval model.xlsx "
ws = ctx.sheet('Pivot')

# Read
for pt in ws.pivots:
    print(pt.name, pt.to_dict())

pt = ws.get_pivot('ByRegion')
print(pt.row_fields, pt.column_fields, pt.value_fields)

# Update (works on loaded pivots too)
pt.remove_columns('Product')          # also remove_rows/filters/values
pt.clear_values().values('Units', 'average', 'Avg Units')
pt.rename('RegionSummary').move('A20')
pt.set_grand_totals(rows=False, columns=True)

# Delete
ws.remove_pivot('OldPivot')           # returns True if removed
# or: pt.delete()
"
```

Edits to the declarative layout (rows/columns/filters/values/style) re-render the pivot in place; you still need `ctx.sync()` for LibreOffice to recompute the values.

## API Reference

```python
# Create (computed by LibreOffice/Excel on sync)
ws.add_pivot(                    # returns a fluent PivotTable
    source,                     # e.g. 'Data!A1:D7' (header in first row)
    anchor='A3',                # top-left cell on this sheet
    name='PivotTable1',
    rows=None, columns=None,    # optional field-name lists
    filters=None,               # optional report-filter fields
    values=None,                # ['Sales'] or [(name, func[, display[, fmt]])]
    style=None,                 # e.g. 'PivotStyleMedium9'
)
pt.rows(*names)                 # fluent equivalents of the kwargs above
pt.columns(*names)
pt.filters(*names)
pt.values(name, func='sum', display_name=None, num_format=None)
pt.values_axis('columns')       # or 'rows' (place value labels)
pt.style(name='PivotStyleMedium9', row_stripes=True, ...)
pt.format_values(ctx, num_format=None)  # force value number format post-sync

# Read / update / delete (incl. pivots loaded from an existing file)
ws.pivots                       # list[PivotTable] wrapping existing pivots
ws.get_pivot(name)              # PivotTable by name (raises KeyError if missing)
ws.remove_pivot(name)           # delete by name -> bool
pt.name, pt.anchor, pt.source   # read-only properties
pt.fields                       # all source header names
pt.row_fields, pt.column_fields, pt.filter_fields
pt.value_fields                 # [{field, func, display_name, num_format}, ...]
pt.style_name
pt.to_dict()                    # JSON-friendly snapshot of the whole layout
pt.remove_rows(*names)          # also remove_columns/filters/values
pt.clear_rows()                 # also clear_columns/filters/values
pt.rename(new_name)             # unique within the sheet
pt.move(anchor)                 # relocate top-left anchor
pt.set_grand_totals(rows=None, columns=None)
pt.delete()                     # remove this pivot (and its cache)
```
