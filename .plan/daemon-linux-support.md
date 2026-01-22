# LibreOffice Daemon: Linux Support

## Current Status

- **macOS**: ✅ Daemon works via Python macro approach
- **Linux**: ✅ Daemon works via helper process + UNO socket approach

## Problem

On Linux, invoking Python macros via `vnd.sun.star.script:` URL crashes LibreOffice with `std::bad_alloc`:

```
terminate called after throwing an instance of 'std::bad_alloc'
  what():  std::bad_alloc
Fatal exception: Signal 6
```

This happens on:
- Debian Bookworm (LO 7.4.7)
- Ubuntu 24.04 (LO 24.2.7)
- Both arm64 and x86_64 architectures

## Root Cause (CONFIRMED)

**The crash is caused by Python ABI mismatch between distro-packaged LibreOffice and venv Python - NOT a macOS vs Linux difference.**

### The Real Difference: Bundled vs Distro-Packaged

| Platform | `libpyuno` links to | Isolated? |
|----------|---------------------|-----------|
| **macOS** | `@loader_path/LibreOfficePython.framework` (bundled) | ✅ Yes |
| **Linux distro packages** | `/lib/.../libpython3.11.so` (system) | ❌ No |

**macOS LibreOffice** bundles its own Python framework and uses `@loader_path` relative linking - completely self-contained, ignores system PATH.

**Linux distro-packaged LibreOffice** links `libpyuno.so` against system Python. When a venv with a different Python version is active:

1. `libpyuno.so` is compiled against **Python 3.11** ABI (system)
2. Venv puts **Python 3.12** on PATH  
3. LibreOffice finds Python 3.12 interpreter but `libpyuno.so` expects 3.11 ABI
4. ABI mismatch → `std::bad_alloc`

### Confirmation

```bash
# libpyuno.so links to system Python 3.11
$ ldd /usr/lib/libreoffice/program/libpyuno.so
  libpython3.11.so.1.0 => /lib/x86_64-linux-gnu/libpython3.11.so.1.0

# But venv has Python 3.12
$ /app/.venv/bin/python --version
  Python 3.12.12
```

Testing in Docker (Debian bookworm + Python 3.12 venv):
- **Without venv active**: ✅ Works
- **With venv active**: ❌ `std::bad_alloc` crash

### References
- [Bug 163960](https://bugs.documentfoundation.org/show_bug.cgi?id=163960) - Comment 4 identified venv as trigger

### Why Our Helper Approach Works

1. Starts `soffice` in **socket mode** (no Python macro execution inside LO)
2. Uses `/usr/bin/python3` (system Python 3.11 with `uno`) to connect via UNO externally
3. Completely bypasses the Python version conflict

### Alternative Fix (Not Implemented)

Could sanitize environment before launching soffice:
```python
env = os.environ.copy()
if 'VIRTUAL_ENV' in env:
    venv_bin = Path(env['VIRTUAL_ENV']) / 'bin'
    env['PATH'] = ':'.join(p for p in env['PATH'].split(':') if p != str(venv_bin))
    del env['VIRTUAL_ENV']
subprocess.Popen(['soffice', ...], env=env)
```
However, the helper approach is more robust as it also solves the `uno` module availability issue (uno can't be pip-installed).

## Working Alternative: UNO Socket Connection

LibreOffice's socket listening mode works perfectly on Linux:

```bash
# Start LO in socket mode
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager"
```

```python
# Connect via UNO (requires system Python with python3-uno)
import uno

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", localContext)

ctx = resolver.resolve(
    "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")

smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

# Open, calculate, save
url = uno.systemPathToFileUrl('/path/to/file.xlsx')
doc = desktop.loadComponentFromURL(url, '_blank', 0, ())
doc.calculateAll()
doc.store()
doc.close(True)
```

## Challenge: `uno` Module Availability

The `uno` Python module:
- Comes from `python3-uno` system package
- Only available in system Python (`/usr/bin/python3`)
- Cannot be pip-installed into a venv
- Located at `/usr/lib/python3/dist-packages/uno.py`

This means we can't use UNO directly from our venv-based application.

## Potential Solutions

### Option 1: Helper Process (Recommended)

Create a small helper script that runs with system Python:

```python
#!/usr/bin/env python3
# daemon_helper.py - runs with system python
"""UNO bridge that accepts commands via stdin/socket and forwards to LibreOffice."""

import uno
import socket
import json

def main():
    # Start LO, connect via UNO
    # Listen for commands on TCP port
    # Forward recalc requests to LO via UNO
    pass
```

Architecture:
1. `start_daemon()` spawns both `soffice` and `daemon_helper.py` (using `/usr/bin/python3`)
2. Our app sends commands to helper via TCP
3. Helper forwards to LO via UNO

### Option 2: Subprocess Calls

For each recalc, spawn a short-lived system Python process:

```python
def daemon_recalc_linux(filename):
    subprocess.run([
        '/usr/bin/python3', '-c', f'''
import uno
# ... connect and recalc ...
''', filename
    ])
```

Simpler but less efficient (no persistent connection).

### Option 3: Symlink uno into venv

```bash
ln -s /usr/lib/python3/dist-packages/uno.py .venv/lib/python3.12/site-packages/
ln -s /usr/lib/python3/dist-packages/unohelper.py .venv/lib/python3.12/site-packages/
```

Fragile - depends on matching Python versions and may have import issues.

### Option 4: Use unoserver

The `unoserver` package does exactly this but requires running its server with system Python:

```bash
/usr/bin/python3 -m unoserver  # Start server
unoconvert input.xlsx output.xlsx  # Client can be any Python
```

We could potentially use unoserver's client library from our venv.

## Implementation (COMPLETED)

### Phase 1: ✅ Document Current Limitation
- Daemon mode works on macOS only
- Linux falls back to cold-start (graceful degradation already implemented)

### Phase 2: ✅ Linux Support via Helper Process
1. ✅ Ship `daemon_helper.py` that runs with system Python (`LINUX_HELPER_SCRIPT` in libre.py)
2. ✅ Detect Linux in `start_daemon()` and spawn helper
3. ✅ Helper manages LO socket connection via UNO
4. ✅ Same TCP protocol (PING/RECALC/QUIT) for consistency

### Performance Results (Ubuntu 24.04 in Docker)
| Mode | Avg Time | Speedup |
|------|----------|---------|
| Daemon | 0.77s | 4.6x faster |
| Cold-start | 3.55s | baseline |

### Required Dockerfile Changes

```dockerfile
RUN apt-get install -y \
    libreoffice-calc \
    libreoffice-script-provider-python \
    python3-uno
```

## Version Compatibility

| Platform | LO Version | Python Macro | UNO Socket | Daemon |
|----------|------------|--------------|------------|--------|
| macOS    | 25.8.1     | ✅           | ✅         | ✅ (macro) |
| Debian 12| 7.4.7      | ❌ crash     | ✅         | ✅ (helper) |
| Ubuntu 24| 24.2.7     | ❌ crash     | ✅         | ✅ (helper) |

## References

- [LibreOffice UNO API](https://api.libreoffice.org/)
- [unoserver](https://github.com/unoconv/unoserver) - Similar approach
- [python-uno docs](https://wiki.documentfoundation.org/Documentation/DevGuide/Professional_UNO)
