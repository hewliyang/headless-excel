# LibreOffice Daemon: Linux Support

## Current Status

- **macOS**: ✅ Daemon works via Python macro approach
- **Linux**: ❌ Python macro approach crashes, needs UNO socket approach

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

## Root Cause

The crash occurs in `libpyuno.so` when LibreOffice tries to execute Python scripts via command-line macro invocation. The `uno` Python module works fine when imported directly - it's specifically the `vnd.sun.star.script:` URL handler that crashes.

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

## Recommended Implementation

### Phase 1: Document Current Limitation
- Daemon mode works on macOS only
- Linux falls back to cold-start (graceful degradation already implemented)

### Phase 2: Linux Support via Helper Process
1. Ship `daemon_helper.py` that runs with system Python
2. Detect Linux in `start_daemon()` and spawn helper
3. Helper manages LO socket connection
4. Same TCP protocol (PING/RECALC/QUIT) for consistency

### Required Dockerfile Changes

```dockerfile
RUN apt-get install -y \
    libreoffice-calc \
    libreoffice-script-provider-python \
    python3-uno
```

## Version Compatibility

| Platform | LO Version | Python Macro | UNO Socket |
|----------|------------|--------------|------------|
| macOS    | 25.8.1     | ✅           | ✅         |
| Debian 12| 7.4.7      | ❌ crash     | ✅         |
| Ubuntu 24| 24.2.7     | ❌ crash     | ✅         |

## References

- [LibreOffice UNO API](https://api.libreoffice.org/)
- [unoserver](https://github.com/unoconv/unoserver) - Similar approach
- [python-uno docs](https://wiki.documentfoundation.org/Documentation/DevGuide/Professional_UNO)
