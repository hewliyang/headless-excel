Issues I Hit 🤔

### 6. Helper for financial modeling patterns

```python
  # Common pattern I used repeatedly
  ctx.project_line_item('Revenue', 'Assumptions!G7')  # Growth rate
  ctx.project_as_percent('Cash', 'Revenue', 'Assumptions!G17')
```

> probs has to be some sort of extension/plugin system

### 7. Inspection tools

```python
  dependencies = ctx.trace_dependents('A1')  # What depends on this cell?
```

> this seems useful - need to figure out how to build the deps graph. in excel - can be circular
