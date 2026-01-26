# Hooks

Hooks allow custom code to run at specific points in the Excel workflow without modifying the core library. Common use cases include:

- Applying formatting conventions (e.g., financial modeling colors)
- Validation and linting
- Logging and auditing
- Telemetry and undo/redo (planned — see [roadmap](roadmap.md#cell-change-tracking))

## Discovery

Hooks are auto-discovered from Python files in these locations:

1. `./.headless-excel/hooks/*.py` — Project-local (takes precedence)
2. `~/.headless-excel/hooks/*.py` — Global fallback

Project-local hooks override global hooks entirely (no merging).

## Hook Types

### `@pre_sync`

Called before each `sync()` operation (before save/recalc/reload). Use this to modify the workbook before it's saved.

```python
from headless_excel import ExcelContext, pre_sync

@pre_sync
def auto_format(ctx: ExcelContext) -> None:
    # Apply formatting before save
    ws = ctx.active
    ws.range("A1:Z1").apply_style(font=Font(bold=True))
```

### `@post_sync`

Called after each `sync()` operation (after save/recalc/reload). Use this to inspect results or perform validation.

```python
from headless_excel import ExcelContext, SyncResult, post_sync

@post_sync
def log_errors(ctx: ExcelContext, result: SyncResult) -> None:
    if not result.success:
        print(f"Warning: {result.total_errors} formula errors found")
```

### `@on_exit`

Called when the context manager exits (after final sync if any). Use this for linting, cleanup, or final validation.

```python
from headless_excel import ExcelContext, on_exit

@on_exit
def final_check(ctx: ExcelContext) -> None:
    errors = ctx.find_errors()
    if errors:
        print(f"Workbook has {errors.total_errors} errors")
```

## Output Configuration

Hook output is captured and prefixed. By default, output goes to stderr.

```python
@pre_sync(output="stderr")   # Default - print to stderr with prefix
def hook1(ctx): ...

@pre_sync(output="stdout")   # Print to stdout with prefix
def hook2(ctx): ...

@pre_sync(output="none")     # Suppress output entirely
def hook3(ctx): ...
```

Output is prefixed with the hook type and function name:

```
[pre-sync:auto_format] Applying header formatting
[post-sync:log_errors] Warning: 2 formula errors found
```

## Extension API (Stateful Hooks)

For hooks that need shared state across calls, use the `@extension` decorator:

```python
from headless_excel import ExcelContext, SyncResult, extension

@extension
def sync_counter(api):
    count = 0

    @api.pre_sync
    def increment(ctx: ExcelContext) -> None:
        nonlocal count
        count += 1
        print(f"Sync #{count}")

    @api.post_sync(output="stdout")
    def report(ctx: ExcelContext, result: SyncResult) -> None:
        print(f"Sync #{count} complete: {'OK' if result.success else 'ERRORS'}")
```

## Examples

See `examples/hooks/` for complete hook implementations:

- `financial_colors.py` — Financial modeling color conventions
- `audit_log.py` — Logging and auditing
- `magic_numbers.py` — Detect hardcoded numbers in formulas
- `code_size.py` — Track workbook complexity
