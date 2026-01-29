# Typed UNO Hooks

Write real Python code with full IDE autocomplete for UNO hooks, instead of string-based `ctx.exec()`.

## Problem

Current UNO hooks require writing code as strings:

```python
@pre_sync
def my_hook(ctx):
    ctx.exec('''
        sheet = doc.getSheets().getByIndex(0)
        sheet.getCellByPosition(0, 0).setValue(42)
    ''')
```

No autocomplete, no type checking, error-prone.

## Solution

Use `@uno_func` decorator + `ooouno` type stubs:

```python
from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetDocument, XSpreadsheet
    from com.sun.star.table import XCell

@on_open
@uno_func
def hide_gridlines(doc: XSpreadsheetDocument) -> None:
    """Full autocomplete works here!"""
    controller = doc.getCurrentController()
    sheets = doc.getSheets()
    
    for i in range(sheets.getCount()):
        sheet: XSpreadsheet = sheets.getByIndex(i)
        controller.setActiveSheet(sheet)
        controller.setPropertyValue("ShowGrid", False)
```

## How It Works

1. **`from __future__ import annotations`** - Makes all type hints strings (PEP 563), so they're not evaluated at runtime

2. **`TYPE_CHECKING` guard** - Imports UNO types only during static analysis, not at runtime

3. **`ooouno` package** - Provides 4300+ type stubs for all UNO interfaces:
   ```bash
   uv add --dev ooouno
   ```

4. **`@uno_func` decorator**:
   - Extracts function body via `inspect.getsource()`
   - Strips type annotations via AST transformation
   - Sends clean code to `ctx.exec()` at runtime

## Implementation

```python
import ast
import inspect
import textwrap

class TypeAnnotationStripper(ast.NodeTransformer):
    """Remove type annotations from AST."""
    
    def visit_FunctionDef(self, node):
        node.returns = None
        for arg in node.args.args:
            arg.annotation = None
        self.generic_visit(node)
        return node
    
    def visit_AnnAssign(self, node):
        # Convert `x: Type = value` to `x = value`
        if node.value is not None:
            return ast.Assign(targets=[node.target], value=node.value)
        return None

def uno_func(fn):
    """Decorator that converts typed function to UNO exec code."""
    # Get and dedent source
    source = inspect.getsource(fn)
    lines = source.split('\n')
    for i, line in enumerate(lines):
        if line.strip().startswith('def '):
            lines = lines[i:]
            break
    source = textwrap.dedent('\n'.join(lines))
    
    # Strip type annotations
    tree = ast.parse(source)
    tree = TypeAnnotationStripper().visit(tree)
    source = ast.unparse(tree)
    
    # Extract function body (skip def line and docstring)
    tree = ast.parse(source)
    body = tree.body[0].body
    if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant):
        body = body[1:]  # Skip docstring
    
    code = ast.unparse(ast.Module(body=body, type_ignores=[]))
    
    def wrapper(ctx):
        return ctx.exec(code)
    
    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    wrapper._uno_source = code
    return wrapper
```

## Available Types

From `ooouno` package (4300+ interfaces):

| Module | Types |
|--------|-------|
| `com.sun.star.sheet` | XSpreadsheetDocument, XSpreadsheet, XCell, etc. (337 types) |
| `com.sun.star.table` | XCell, XCellRange, XTableColumns, etc. |
| `com.sun.star.beans` | XPropertySet, PropertyValue, etc. |
| `com.sun.star.frame` | XController, XDesktop, etc. |
| `com.sun.star.uno` | XInterface, XComponentContext, etc. |

## Example: Financial Colors Hook

```python
from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetDocument, XSpreadsheet
    from com.sun.star.table import XCell

BLUE = 0x0000FF   # Hardcoded values
BLACK = 0x000000  # Formulas
GREEN = 0x008000  # External refs

@pre_sync
@uno_func
def financial_colors(doc: XSpreadsheetDocument) -> None:
    """Apply financial modeling color conventions."""
    sheets = doc.getSheets()
    
    for sheet_idx in range(sheets.getCount()):
        sheet: XSpreadsheet = sheets.getByIndex(sheet_idx)
        
        for row in range(100):
            for col in range(26):
                cell: XCell = sheet.getCellByPosition(col, row)
                formula = cell.getFormula()
                
                if formula and formula.startswith("="):
                    color = GREEN if "!" in formula else BLACK
                    cell.setPropertyValue("CharColor", color)
                elif cell.getType().value == 1:  # NUMBER
                    cell.setPropertyValue("CharColor", BLUE)
```

## Dual-Mode Support

Support both UnoContext and ExcelContext:

```python
@pre_sync
def my_hook(ctx) -> None:
    if hasattr(ctx, 'exec'):
        # UnoContext - use uno_func logic
        ctx.exec('''...''')
    else:
        # ExcelContext - use openpyxl
        for ws in ctx.workbook.worksheets:
            ...
```

Or create a decorator that handles both:

```python
def dual_mode(uno_fn, openpyxl_fn):
    def wrapper(ctx):
        if hasattr(ctx, 'exec'):
            return uno_fn(ctx)
        else:
            return openpyxl_fn(ctx)
    return wrapper
```

## Limitations

1. **`inspect.getsource()` requirement** - Function must be defined in a file (not stdin/exec)
2. **No closures** - Can't capture variables from outer scope (code runs in LibreOffice)
3. **Limited imports** - Only `doc`, `uno`, `json` available in exec context

## Files

- `/tmp/uno_typed_hooks.py` - Working prototype
- `ooouno` package - Type stubs (installed via `uv add --dev ooouno`)
