# Multi-Instance Daemon Support

Run multiple concurrent LibreOffice daemon instances with isolated profiles.

## Current State

Single daemon instance with global state:

```
~/.headless-excel/
├── daemon.pid
└── daemon.port
```

Limitations:

- One PID/port file pair
- Shared LO profile (macOS/Windows macro location)
- Hardcoded UNO port 2002 on Linux

## Target State

Per-instance isolation:

```
~/.headless-excel/
├── instances/
│   ├── default/
│   │   ├── daemon.pid
│   │   └── daemon.port
│   └── worker-1/
│       ├── daemon.pid
│       └── daemon.port
└── profiles/
    ├── default/
    │   └── user/Scripts/python/unobridge.py
    └── worker-1/
        └── user/Scripts/python/unobridge.py
```

Key: LibreOffice's `-env:UserInstallation=file:///path` flag enables isolated profiles.

## Verified

Sanity check (`scripts/multi_instance_test.py`) confirms two soffice instances can run simultaneously with separate profiles and ports.

## Implementation

### 1. Instance Path Management

```python
# daemon/config.py

@dataclass
class InstancePaths:
    instance_id: str
    pid_file: Path
    port_file: Path
    profile_dir: Path

    @classmethod
    def for_instance(cls, instance_id: str = "default") -> "InstancePaths":
        base = HEADLESS_EXCEL_DIR / "instances" / instance_id
        profile = HEADLESS_EXCEL_DIR / "profiles" / instance_id
        return cls(
            instance_id=instance_id,
            pid_file=base / "daemon.pid",
            port_file=base / "daemon.port",
            profile_dir=profile,
        )
```

### 2. Update Base Functions

Add `instance_id` parameter to all daemon functions:

```python
def is_daemon_running(instance_id: str = "default") -> bool: ...
def send_daemon_command(cmd: str, instance_id: str = "default") -> str: ...
def cleanup_daemon_files(instance_id: str = "default") -> None: ...
```

### 3. Platform Changes

#### macOS/Windows

Install macro to instance-specific profile, start with isolated profile:

```python
def start_daemon_macos(wait: bool, timeout: float, instance_id: str = "default") -> int:
    paths = InstancePaths.for_instance(instance_id)

    # Install macro to instance profile
    macro_dir = paths.profile_dir / "user/Scripts/python"
    macro_dir.mkdir(parents=True, exist_ok=True)
    (macro_dir / "unobridge.py").write_text(get_unobridge_macro(config))

    cmd = [
        soffice,
        "--headless",
        f"-env:UserInstallation=file://{paths.profile_dir}",
        "vnd.sun.star.script:unobridge.py$start_server?language=Python&location=user",
    ]
```

#### Linux

Pass instance-specific UNO port and profile to helper:

```python
def start_daemon_linux(..., instance_id: str = "default") -> int:
    paths = InstancePaths.for_instance(instance_id)
    daemon_port = find_free_port()
    uno_port = find_free_port()

    env["HEADLESS_EXCEL_PORT"] = str(daemon_port)
    env["HEADLESS_EXCEL_UNO_PORT"] = str(uno_port)
    env["HEADLESS_EXCEL_PROFILE"] = str(paths.profile_dir)
```

Update `linux_helper.py` to read `HEADLESS_EXCEL_UNO_PORT` and `HEADLESS_EXCEL_PROFILE` from env.

### 4. Public API

```python
def start_daemon(instance_id: str = "default", wait: bool = True, timeout: float = 15) -> int: ...
def stop_daemon(instance_id: str = "default") -> bool: ...
def stop_all_daemons() -> int: ...
def list_instances() -> list[str]: ...
```

Context manager for temporary instances:

```python
@contextmanager
def daemon_instance(instance_id: str | None = None) -> Iterator[str]:
    if instance_id is None:
        instance_id = f"temp-{uuid4().hex[:8]}"
    try:
        start_daemon(instance_id)
        yield instance_id
    finally:
        stop_daemon(instance_id)
        cleanup_instance(instance_id)
```

### 5. CLI

```bash
headless-excel daemon start --instance worker-1
headless-excel daemon stop --instance worker-1
headless-excel daemon stop --all
headless-excel daemon list
headless-excel recalc file.xlsx --instance worker-1
headless-excel recalc *.xlsx --parallel 4
```

## Open Questions

1. **Max instances** - Memory limit? (~200-400MB per LO instance)
2. **Profile cleanup** - Delete on stop or keep for faster restart?
3. **Pool implementation** - Round-robin vs least-recently-used?
4. **Auto-restart** - Should pool restart crashed instances?

---

# Cell Change Tracking

Track cell modifications at each sync for telemetry, debugging, and undo/redo systems.

## Current State

Writes trigger a simple dirty flag callback:

```python
# proxy.py
OnWriteCallback = Callable[[], None] | None

# context.py
def _mark_dirty(self) -> None:
    self._dirty = True
```

`SyncResult` only contains error information:

```python
@dataclass
class SyncResult:
    success: bool
    total_errors: int
    errors: dict[str, list[str]]
    error_details: list[ErrorDetail]
```

No visibility into which cells were modified or their before/after states.

## Target State

Full cell snapshots for all modified cells, enabling:

- **Telemetry**: "Agent wrote to A1, B2:C5 in this step"
- **Debugging**: Inspect exact values/formulas/styles written
- **Undo/redo**: Users can implement their own history system

```python
@dataclass
class SyncResult:
    success: bool
    total_errors: int
    errors: dict[str, list[str]]
    error_details: list[ErrorDetail]

    # NEW
    changes: list[CellChange]
```

## Implementation

### 1. Cell Snapshot Serialization

New module `src/headless_excel/telemetry.py`:

```python
from dataclasses import dataclass, asdict
from typing import Any
from openpyxl.descriptors.base import Descriptor

@dataclass
class CellSnapshot:
    """Full state of a cell at a point in time."""
    value: Any
    formula: str | None
    font: dict | None
    fill: dict | None
    border: dict | None
    alignment: dict | None
    number_format: str
    protection: dict | None

    def to_dict(self) -> dict:
        return asdict(self)

@dataclass
class CellChange:
    """Before/after state for a single cell."""
    coordinate: str  # "Sheet1!A1"
    before: CellSnapshot
    after: CellSnapshot

def _serialize_style(obj, depth: int = 0) -> dict | None:
    """Recursively serialize openpyxl style objects."""
    if depth > 5 or obj is None:
        return None
    if isinstance(obj, Descriptor):  # Filter unset descriptor values
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, (list, tuple)):
        items = [_serialize_style(x, depth + 1) for x in obj]
        return [x for x in items if x is not None] or None

    if hasattr(obj, "__elements__") or hasattr(obj, "__attrs__"):
        result = {}
        for attr in list(getattr(obj, "__elements__", [])) + list(getattr(obj, "__attrs__", [])):
            try:
                val = getattr(obj, attr, None)
                if val is not None and not callable(val):
                    serialized = _serialize_style(val, depth + 1)
                    if serialized is not None:
                        result[attr] = serialized
            except (AttributeError, TypeError):
                pass
        return result if result else None
    return str(obj)

def snapshot_cell(cell) -> CellSnapshot:
    """Capture full cell state."""
    val = cell.value
    return CellSnapshot(
        value=val,
        formula=val if isinstance(val, str) and val.startswith("=") else None,
        font=_serialize_style(cell.font),
        fill=_serialize_style(cell.fill),
        border=_serialize_style(cell.border),
        alignment=_serialize_style(cell.alignment),
        number_format=cell.number_format,
        protection=_serialize_style(cell.protection),
    )
```

### 2. Enhanced Write Callback

Change callback signature to pass cell location:

```python
# proxy.py
OnWriteCallback = Callable[[str, str], None] | None  # (sheet_name, coordinate)

# In CellProxy.value setter, font setter, etc:
if self._on_write:
    sheet_name = self._formula_cell.parent.title
    self._on_write(sheet_name, self._formula_cell.coordinate)
```

### 3. Change Tracking in ExcelContext

```python
# context.py

class ExcelContext:
    def __init__(self, ...):
        ...
        self._pending_writes: dict[str, CellSnapshot] = {}  # coord -> before snapshot
        self._baseline_cache: dict[str, CellSnapshot] = {}  # coord -> state at last sync

    def _on_cell_write(self, sheet_name: str, coordinate: str) -> None:
        """Called before each cell write."""
        self._dirty = True
        full_coord = f"{sheet_name}!{coordinate}"

        # Capture "before" state (only first write per cell per sync cycle)
        if full_coord not in self._pending_writes:
            if full_coord in self._baseline_cache:
                self._pending_writes[full_coord] = self._baseline_cache[full_coord]
            else:
                # First time seeing this cell - snapshot current state
                cell = self._get_cell(sheet_name, coordinate)
                self._pending_writes[full_coord] = snapshot_cell(cell)
```

### 4. Capture Changes at Sync

```python
def sync(self, raise_on_errors: bool = False) -> SyncResult:
    ...
    run_pre_sync_hooks(self)

    # Capture "after" snapshots before saving
    changes: list[CellChange] = []
    for coord, before in self._pending_writes.items():
        sheet_name, cell_ref = coord.split("!", 1)
        cell = self._get_cell(sheet_name, cell_ref)
        after = snapshot_cell(cell)
        changes.append(CellChange(coordinate=coord, before=before, after=after))

    # Save, recalc, reload...
    self._workbook.save(self.path)
    recalc(self.path, timeout=self._recalc_timeout)
    self._workbook = load_workbook(self.path)
    ...

    # Update baseline cache and clear pending
    for change in changes:
        self._baseline_cache[change.coordinate] = change.after
    self._pending_writes.clear()

    sync_result = SyncResult(
        success=...,
        total_errors=...,
        errors=...,
        error_details=...,
        changes=changes,
    )
    ...
```

### 5. Hook Usage

```python
from headless_excel import post_sync, ExcelContext, SyncResult

@post_sync
def telemetry(ctx: ExcelContext, result: SyncResult) -> None:
    if result.changes:
        print(f"Modified {len(result.changes)} cell(s):")
        for change in result.changes:
            print(f"  {change.coordinate}: {change.before.value!r} → {change.after.value!r}")

# Undo/redo implementation
class UndoHistory:
    def __init__(self):
        self.stack: list[list[CellChange]] = []

    def record(self, changes: list[CellChange]) -> None:
        self.stack.append(changes)

    def undo(self, ctx: ExcelContext) -> None:
        if not self.stack:
            return
        changes = self.stack.pop()
        for change in changes:
            sheet, cell = change.coordinate.split("!", 1)
            ctx.workbook[sheet][cell].value = change.before.value
            # Restore styles...
```

### 6. Opt-in Flag (Optional)

If performance overhead is a concern:

```python
with ExcelContext("file.xlsx", track_changes=True) as ctx:
    ...  # Changes tracked

with ExcelContext("file.xlsx", track_changes=False) as ctx:
    ...  # No tracking, slightly faster
```

Default: `True` (safer for debugging), or `False` (opt-in for production).

## Open Questions

1. **Default on/off** - Enable by default for discoverability, or opt-in for performance?
2. **Range writes** - `ws.range("A1:C3").values = [[...]]` fires 9 callbacks. Batch into single range change?
3. **Style-only tracking** - Track when only font/fill changes, not value?
4. **Memory** - Large workbooks with many writes could accumulate snapshots. Cap or warn?
5. **Computed values** - Track cells whose VALUE changed due to recalc (not just explicit writes)?
