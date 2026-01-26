# headless-excel

A light wrapper around `openpyxl` providing ergonomic APIs for agents. Uses LibreOffice in headless mode to evaluate formulas. Can be used in [CLI](#cli) or [SDK](#sdk) mode.

## Requirements

- Python 3.10+
- LibreOffice (`libreoffice-calc`)

## Installation

```bash
pip install headless-excel
```

For the live viewer (`watch` command):

```bash
pip install headless-excel[watch]
```

Install LibreOffice:

```bash
# macOS
brew install --cask libreoffice

# Ubuntu/Debian
sudo apt install libreoffice libreoffice-calc

# Windows
winget install -e --id TheDocumentFoundation.LibreOffice
```

Verify setup:

```bash
headless-excel check
```

## SDK

Use as a Python library. All workbook and worksheet objects are proxies over `openpyxl`. If a proxy method doesn't exist, it forwards to the underlying `openpyxl` object.

```python
from headless_excel import create, run

# Create new workbook
with create("model.xlsx") as ctx:
    ws = ctx.active
    ws["A1"] = 100
    ws["A2"] = 200
    ws["A3"] = "=A1+A2"

    ctx.sync()  # saves, recalculates, reloads
    print(ws["A3"].value)  # 300

# Open existing workbook
with run("model.xlsx") as ctx:
    ws = ctx.active
    ws["A1"] = 500

    ctx.sync()
    print(ws["A3"].value)  # 700
```

## CLI

```bash
headless-excel check                    # verify setup
headless-excel create FILE              # create empty workbook
headless-excel eval FILE CODE           # execute Python code
headless-excel eval FILE -              # read code from stdin
headless-excel watch FILE [--open]      # live viewer (requires [watch] extras)
headless-excel libreoffice start        # start daemon
headless-excel libreoffice stop         # stop daemon
headless-excel libreoffice status       # check daemon status
```

The `eval` command injects the following into globals:

| Name            | Description                                                                                             |
| --------------- | ------------------------------------------------------------------------------------------------------- |
| `ctx`           | ExcelContext for the file (same as returned by `run()` and `create()`)                                  |
| `NumberFormats` | Class with predefined number format constants (e.g., `NumberFormats.CURRENCY`, `NumberFormats.PERCENT`) |

### Example

```bash
headless-excel create model.xlsx
headless-excel eval model.xlsx "
ws = ctx.active
ws.write('A1', [
    ['Product', 'Q1', 'Q2', 'Total'],
    ['Widget', 100, 150, '=B2+C2'],
])
ctx.sync()
print(ws.range('A1:D2').dump())
"
```

## Daemon

Formula recalculation requires `libreoffice-calc`. By default, each `sync()` spawns a new LibreOffice process.

The daemon is a TCP server running inside a persistent LibreOffice instance. It accepts commands to recalculate workbooks via the UNO API, avoiding startup overhead:

```bash
headless-excel libreoffice start
```

The daemon auto-exits after 5 minutes of inactivity. `sync()` automatically tries to start the daemon, falling back to a cold headless LibreOffice run only if that fails.

You can manually control the daemon lifecycle via CLI:

```bash
headless-excel libreoffice start   # start daemon
headless-excel libreoffice stop    # stop daemon
headless-excel libreoffice status  # check if running
```

Or via SDK:

```python
from headless_excel import start_daemon, stop_daemon, is_daemon_running

start_daemon()
stop_daemon()
print(is_daemon_running())
```

## Configuration

| Variable                      | Default | Description                                              |
| ----------------------------- | ------- | -------------------------------------------------------- |
| `HEADLESS_EXCEL_PORT`         | 49152   | Preferred TCP port (auto-finds free port if unavailable) |
| `HEADLESS_EXCEL_IDLE_TIMEOUT` | 300     | Seconds of inactivity before daemon auto-exits           |
| `HEADLESS_EXCEL_MAX_ERRORS`   | 10      | Max formula errors to display when syncing               |

## Documentation

- [API Reference](docs/api.md) — Full SDK documentation
- [Hooks](docs/hooks.md) — Custom hooks for formatting, validation, and linting
- [Roadmap](docs/roadmap.md) — Planned features and improvements

## License

Apache 2.0
