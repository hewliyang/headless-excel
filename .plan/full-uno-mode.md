# Full UNO Mode: Agent-First Excel Automation

## Overview

Replace openpyxl-based cell operations with direct UNO code execution inside LibreOffice. Agents write UNO/Python code that runs in the LibreOffice process, giving full access to all spreadsheet features with optimal performance.

## Why

| Current Approach | Full UNO Mode |
|------------------|---------------|
| openpyxl writes → save → UNO reload → calc | Direct UNO writes → calc (instant) |
| ~11s sync for 2MB file | ~100ms sync |
| Limited to openpyxl features | Full LibreOffice feature set |
| We maintain abstraction layer | Agent learns UNO directly |

## Architecture

```
┌─────────────────┐     exec(code)      ┌─────────────────────────┐
│   Agent/User    │ ──────────────────► │   LibreOffice Daemon    │
│   (Python)      │                     │   (soffice --headless)  │
│                 │ ◄────────────────── │                         │
└─────────────────┘   result/error      │   ┌─────────────────┐   │
                                        │   │ Python Macro    │   │
                                        │   │ - exec(code)    │   │
                                        │   │ - pre-loaded    │   │
                                        │   │   helpers       │   │
                                        │   └─────────────────┘   │
                                        └─────────────────────────┘
```

## Implementation Plan

### Phase 1: EXEC Command in Daemon

Add new command to `daemon/base.py`:

```python
elif cmd.startswith("EXEC:"):
    # EXEC:session_id:base64_encoded_code
    parts = cmd.split(":", 2)
    session_id, code_b64 = parts[1], parts[2]
    code = base64.b64decode(code_b64).decode('utf-8')
    
    doc, filepath = sessions[session_id]
    
    # Pre-loaded globals for the exec context
    exec_globals = {
        'doc': doc,
        'desktop': desktop,
        'uno': uno,
        'json': json,
        # Helper functions
        'cell': make_cell_helper(doc),
        'value': make_value_helper(doc),
        'formula': make_formula_helper(doc),
        'recalc': lambda: doc.calculateAll(),
        'sheet': lambda name=None: get_sheet(doc, name),
        'sheets': lambda: get_sheet_names(doc),
    }
    
    exec_locals = {}
    try:
        exec(code, exec_globals, exec_locals)
        result = exec_locals.get('result', None)
        return f"OK:{json.dumps(result)}"
    except Exception as e:
        return f"ERROR:{traceback.format_exc()}"
```

### Phase 2: Helper Functions

Pre-loaded helpers that make common operations easy:

```python
# Cell helper - get/set by A1 reference
def cell(ref, value=None, sheet=None):
    """
    cell("A1")           # get value
    cell("A1", 100)      # set value
    cell("A1", "=B1*2")  # set formula
    """
    sh = get_sheet(doc, sheet)
    col, row = parse_ref(ref)
    c = sh.getCellByPosition(col, row)
    if value is None:
        return c.getValue() if c.getType() == VALUE else c.getString()
    elif isinstance(value, str) and value.startswith("="):
        c.setFormula(value)
    elif isinstance(value, (int, float)):
        c.setValue(value)
    else:
        c.setString(str(value))

# Range helper - bulk operations
def range_values(ref, data=None, sheet=None):
    """
    range_values("A1:C3")              # get 2D array
    range_values("A1:C3", [[1,2,3],    # set 2D array
                          [4,5,6],
                          [7,8,9]])
    """
    sh = get_sheet(doc, sheet)
    rng = sh.getCellRangeByName(ref)
    if data is None:
        return [list(row) for row in rng.getDataArray()]
    else:
        rng.setDataArray(tuple(tuple(r) for r in data))

# Style helper
def style(ref, font_color=None, bg_color=None, number_format=None, sheet=None):
    """
    style("A1", font_color=0x0000FF)  # blue text
    style("A1:A10", number_format="$#,##0.00")
    """
    ...

# Named range helper  
def named(name, value=None):
    """
    named("Revenue")           # get value of named range
    named("Revenue", 1000000)  # set value
    """
    ...

# Conditional formatting helper
def conditional(ref, rule_type, **kwargs):
    """
    conditional("A1:A100", "color_scale", 
                min_color=0xFF0000, max_color=0x00FF00)
    conditional("B1:B100", "data_bar", color=0x0000FF)
    conditional("C1:C100", "icon_set", icons="arrows")
    """
    ...
```

### Phase 3: Python Client API

New `UnoSession` class in `headless_excel/uno_session.py`:

```python
class UnoSession:
    """Direct UNO code execution session."""
    
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._session_id: str | None = None
    
    def __enter__(self):
        # Open or create file in LibreOffice
        if self.path.exists():
            resp = send_daemon_command(f"OPEN:{self.path.absolute()}")
        else:
            resp = send_daemon_command(f"NEW:{self.path.absolute()}")
        self._session_id = resp.split(":")[1]
        return self
    
    def __exit__(self, *args):
        if self._session_id:
            send_daemon_command(f"SAVE:{self._session_id}")
            send_daemon_command(f"CLOSE:{self._session_id}")
    
    def exec(self, code: str) -> Any:
        """Execute Python code inside LibreOffice."""
        code_b64 = base64.b64encode(code.encode()).decode()
        resp = send_daemon_command(f"EXEC:{self._session_id}:{code_b64}")
        if resp.startswith("ERROR:"):
            raise UnoExecError(resp[6:])
        return json.loads(resp[3:]) if resp.startswith("OK:") else None
    
    # Convenience methods that generate code
    def cell(self, ref: str, value: Any = None) -> Any:
        if value is None:
            return self.exec(f"result = cell('{ref}')")
        else:
            self.exec(f"cell('{ref}', {repr(value)})")
    
    def recalc(self):
        self.exec("recalc()")
    
    def range(self, ref: str, data: list[list] = None) -> list[list]:
        if data is None:
            return self.exec(f"result = range_values('{ref}')")
        else:
            self.exec(f"range_values('{ref}', {repr(data)})")
```

### Phase 4: Agent-Friendly Interface

For agents, provide a simple interface:

```python
from headless_excel import uno_session

# Simple usage
with uno_session("model.xlsx") as s:
    # Set inputs
    s.cell("B2", 1000000)  # Revenue
    s.cell("B3", 0.4)      # Margin
    
    # Recalc
    s.recalc()
    
    # Read outputs
    profit = s.cell("B10")
    
# Advanced usage - raw UNO code
with uno_session("model.xlsx") as s:
    s.exec("""
        # Access any LibreOffice feature
        sheet = doc.getSheets().getByName("Inputs")
        
        # Set up conditional formatting
        from com.sun.star.sheet import ConditionOperator
        formats = sheet.getConditionalFormats()
        ...
        
        # Use solver
        solver = desktop.createInstance("com.sun.star.sheet.Solver")
        ...
        
        result = {"status": "ok", "profit": cell("B10")}
    """)
```

### Phase 5: Documentation for Agents

Create `docs/uno-api.md` with:

1. **Quick reference** - common operations
2. **Helper function docs** - cell(), range_values(), style(), etc.
3. **UNO API patterns** - how to do things the UNO way
4. **Examples** - financial modeling snippets
5. **Gotchas** - 0-indexed vs 1-indexed, etc.

```markdown
# UNO Quick Reference

## Cell Operations
cell("A1")              # get value
cell("A1", 100)         # set number
cell("A1", "=B1*2")     # set formula
cell("A1", "Hello")     # set string

## Range Operations  
range_values("A1:C3")   # get 2D array
range_values("A1:C3", [[1,2,3], [4,5,6], [7,8,9]])

## Formatting
style("A1", font_color=0x0000FF)           # blue text
style("A1:A10", number_format="$#,##0.00") # currency
style("B1:B10", number_format="0.0%")      # percentage

## Calculation
recalc()                          # recalculate all
doc.setPropertyValue("IsIterationEnabled", True)  # enable iterative calc

## Sheets
sheet()                 # active sheet
sheet("Inputs")         # sheet by name
sheets()                # list all sheet names
doc.getSheets().insertNewByName("NewSheet", 0)    # create sheet
```

## Migration Path

### Option A: Parallel APIs
Keep both `run()` (openpyxl-based) and `uno_session()` (UNO-based):

```python
# Old way - still works
with run("model.xlsx") as ctx:
    ctx.active["A1"] = 100
    ctx.sync()

# New way - faster
with uno_session("model.xlsx") as s:
    s.cell("A1", 100)
    s.recalc()
```

### Option B: Backend Switch
Add `backend` parameter to existing API:

```python
with run("model.xlsx", backend="uno") as ctx:
    ctx.active["A1"] = 100  # Uses UNO under the hood
    ctx.sync()              # Just recalc(), no file I/O
```

### Recommendation
Start with Option A (parallel APIs), deprecate openpyxl path later if UNO proves stable.

## Testing Strategy

1. **Unit tests** - Helper functions in isolation
2. **Integration tests** - Full exec() round-trips
3. **Parity tests** - Same operations via openpyxl and UNO, compare results
4. **Performance benchmarks** - Track sync times at various file sizes
5. **Feature tests** - Conditional formatting, charts, etc.

## Risks & Mitigations

| Risk | Mitigation |
|------|------------|
| UNO API instability | Pin LibreOffice version, test on CI |
| Security (arbitrary exec) | Document that this is agent-first, not multi-tenant |
| Error messages confusing | Wrap exceptions with context, show code location |
| LibreOffice process crashes | Auto-restart daemon, retry logic |
| Memory leaks from long sessions | Add session timeout, periodic restart |

## Timeline Estimate

| Phase | Effort |
|-------|--------|
| Phase 1: EXEC command | 2-3 hours |
| Phase 2: Helper functions | 3-4 hours |
| Phase 3: Python client | 2-3 hours |
| Phase 4: Agent interface | 1-2 hours |
| Phase 5: Documentation | 2-3 hours |
| Testing & polish | 4-6 hours |
| **Total** | **~2-3 days** |

## Success Metrics

1. **Performance**: sync() < 200ms for any file size
2. **Feature parity**: All openpyxl examples work in UNO mode
3. **Agent usability**: Agent can learn API from docs + examples
4. **Reliability**: < 1% error rate in production use
