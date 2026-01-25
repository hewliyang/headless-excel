"""CLI for headless-excel."""

import argparse
import asyncio
import subprocess
import sys
import tempfile
from pathlib import Path

from headless_excel import NumberFormats, create, run
from headless_excel.daemon import is_daemon_running, start_daemon, stop_daemon
from headless_excel.daemon.base import PID_FILE, get_soffice_path
from headless_excel.hooks import get_registry
from headless_excel.libre import setup_libreoffice_macro
from headless_excel.watch import watch

GREEN = "\033[32m"
RED = "\033[31m"
DIM = "\033[2m"
RESET = "\033[0m"

# Ensure UTF-8 encoding for Unicode symbols on Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[union-attr]


def _ok(msg: str) -> None:
    print(f"{GREEN}✓{RESET} {msg}")


def _fail(msg: str) -> None:
    print(f"{RED}✗{RESET} {msg}", file=sys.stderr)


def _info(msg: str, indent: int = 1) -> None:
    print(f"{'  ' * indent}{DIM}{msg}{RESET}")


def _bullet(msg: str, indent: int = 2) -> None:
    print(f"{'  ' * indent}{DIM}•{RESET} {msg}")


def _get_hook_name(hook: object) -> str:
    """Get the name of a hook function."""
    return getattr(hook, "__name__", repr(hook))


def _find_hook_files(directory: Path) -> list[Path]:
    """Find all Python hook files in a directory (excluding private files)."""
    if not directory.is_dir():
        return []
    return sorted(f for f in directory.rglob("*.py") if not f.name.startswith("_"))


def _get_libreoffice_version() -> str | None:
    """Get LibreOffice version string, or None if unavailable."""
    soffice = get_soffice_path()
    if not soffice:
        return None
    try:
        result = subprocess.run(
            [soffice, "--version"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        return result.stdout.strip()
    except Exception:
        return None


def _run_recalc_test() -> tuple[bool, str]:
    """Run a quick recalc test. Returns (success, message)."""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            test_file = Path(tmpdir) / "test.xlsx"
            with create(test_file) as ctx:
                ctx.active["A1"] = 2
                ctx.active["A2"] = 3
                ctx.active["A3"] = "=A1+A2"
                ctx.sync()
                result = ctx.active.cell(3, 1).value

            if result == 5:
                return True, "Recalc test passed (2+3=5)"
            else:
                return False, f"Recalc test failed (expected 5, got {result})"
    except Exception as e:
        return False, f"Recalc test failed: {e}"


def _check_libreoffice() -> tuple[bool, str | None]:
    """Check if LibreOffice is available. Returns (found, path)."""
    soffice = get_soffice_path()
    if soffice:
        _ok(f"LibreOffice found: {soffice}")
        if version := _get_libreoffice_version():
            _info(version)
        return True, soffice

    _fail("LibreOffice not found")
    _info("Install it:")
    if sys.platform == "darwin":
        _info("brew install --cask libreoffice", indent=2)
    elif sys.platform == "win32":
        _info("Download from https://www.libreoffice.org/download/", indent=2)
    else:
        _info("sudo apt install libreoffice libreoffice-calc", indent=2)
    return False, None


def _check_macro() -> bool:
    """Check/setup LibreOffice macro. Returns success."""
    if setup_libreoffice_macro():
        _ok("LibreOffice macro configured")
        return True
    _fail("Failed to configure LibreOffice macro")
    return False


def _check_recalc() -> bool:
    """Run a quick recalc test. Returns success."""
    print("\nRunning recalc test...")
    success, msg = _run_recalc_test()
    (_ok if success else _fail)(msg)
    return success


def _print_hook_source(directory: Path, label: str) -> None:
    """Print hook source directory and its files."""
    hook_files = _find_hook_files(directory)
    suffix = "" if hook_files else ", empty"
    _info(f"source: {directory.absolute()} ({label}{suffix})")
    for f in hook_files:
        _bullet(str(f.relative_to(directory)))


def _print_hooks_info() -> None:
    """Print information about discovered hooks."""
    print("\nHooks:")

    local_dir = Path(".headless-excel/hooks")
    global_dir = Path.home() / ".headless-excel" / "hooks"

    if local_dir.is_dir():
        _print_hook_source(local_dir, "project-local")
    elif global_dir.is_dir():
        _print_hook_source(global_dir, "global")
    else:
        _info("(no hooks directory found)")

    # Show registered hooks
    registry = get_registry()
    hook_types = [
        ("pre_sync", registry.pre_sync),
        ("post_sync", registry.post_sync),
        ("on_exit", registry.on_exit),
    ]

    registered = [(name, hooks) for name, hooks in hook_types if hooks]
    if registered:
        print()
        _info("registered:")
        max_len = max(len(name) for name, _ in registered)
        for name, hooks in registered:
            names = ", ".join(_get_hook_name(h.fn) for h in hooks)
            _info(f"{name}:{' ' * (max_len - len(name))}  {names}", indent=2)


def _check_daemon() -> None:
    """Check daemon status."""
    print("\nDaemon:")
    if is_daemon_running():
        _ok("LibreOffice daemon is running")
        if PID_FILE.exists():
            _info(f"PID: {PID_FILE.read_text().strip()}")
    else:
        _info("LibreOffice daemon is not running (optional)")
        _info("Start for faster recalc: headless-excel libreoffice start", indent=2)


def cmd_check() -> int:
    """Check if environment is set up correctly."""
    print("headless-excel environment check\n")

    soffice_ok, _ = _check_libreoffice()
    macro_ok = _check_macro() if soffice_ok else False
    _print_hooks_info()
    _check_daemon()
    recalc_ok = _check_recalc() if soffice_ok and macro_ok else True

    print()

    all_ok = soffice_ok and macro_ok and recalc_ok

    if not all_ok:
        print("Some checks failed. See above for details.")
        return 1

    print("All checks passed! Ready to use.")
    return 0


def cmd_libreoffice_start() -> int:
    """Start the LibreOffice daemon."""
    if is_daemon_running():
        _ok("Daemon is already running")
        if PID_FILE.exists():
            _info(f"PID: {PID_FILE.read_text().strip()}")
        return 0

    print("Starting LibreOffice daemon...")
    try:
        pid = start_daemon(wait=True, timeout=15)
        _ok(f"Daemon started (PID: {pid})")
        _info("Stop with: headless-excel libreoffice stop")
        return 0
    except Exception as e:
        _fail(f"Failed to start daemon: {e}")
        return 1


def cmd_libreoffice_stop() -> int:
    """Stop the LibreOffice daemon."""
    if not is_daemon_running():
        print("Daemon is not running")
        # Still try to clean up any stale processes
        stop_daemon()
        return 0

    print("Stopping LibreOffice daemon...")
    if stop_daemon():
        _ok("Daemon stopped")
        return 0
    else:
        _fail("Failed to stop daemon")
        return 1


def cmd_libreoffice_status() -> int:
    """Check the status of the LibreOffice daemon."""
    if is_daemon_running():
        _ok("Daemon is running")
        if PID_FILE.exists():
            _info(f"PID: {PID_FILE.read_text().strip()}")
        return 0
    else:
        print("Daemon is not running")
        _info("Start with: headless-excel libreoffice start")
        return 1


def main():
    parser = argparse.ArgumentParser(
        prog="headless-excel",
        description="Excel automation tool for headless environments",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # check
    subparsers.add_parser("check", help="Check environment setup")

    # create
    p_create = subparsers.add_parser("create", help="Create new Excel file")
    p_create.add_argument("file", help="Output file path")

    # eval
    p_eval = subparsers.add_parser("eval", help="Eval Python code against file")
    p_eval.add_argument("file", help="Excel file")
    p_eval.add_argument(
        "code", nargs="?", default="-", help="Code to eval (default: stdin)"
    )

    # watch
    p_watch = subparsers.add_parser("watch", help="Live viewer with auto-reload")
    p_watch.add_argument("file", help="Excel file to watch")
    p_watch.add_argument(
        "--port", type=int, default=8080, help="HTTP port (default: 8080)"
    )
    p_watch.add_argument(
        "--ws-port", type=int, default=8765, help="WebSocket port (default: 8765)"
    )
    p_watch.add_argument(
        "--open", action="store_true", help="Open browser automatically"
    )

    # libreoffice
    p_libre = subparsers.add_parser(
        "libreoffice",
        help="Manage LibreOffice daemon for fast recalculation",
        aliases=["lo"],
    )
    libre_sub = p_libre.add_subparsers(dest="libre_command", required=True)
    libre_sub.add_parser("start", help="Start the LibreOffice daemon")
    libre_sub.add_parser("stop", help="Stop the LibreOffice daemon")
    libre_sub.add_parser("status", help="Check daemon status")

    args = parser.parse_args()

    match args.command:
        case "check":
            sys.exit(cmd_check())

        case "create":
            with create(args.file) as ctx:
                ctx.sync()
            print(f"Created {args.file}")

        case "eval":
            file_path = Path(args.file)
            valid_extensions = {".xlsx", ".xls", ".xlsm", ".xlsb"}
            looks_like_code = file_path.suffix.lower() not in valid_extensions and any(
                c in args.file for c in "=()[]{}'\""
            )
            if looks_like_code and not file_path.exists():
                _fail("Missing filename. The first argument should be an Excel file.")
                _info('Example: headless-excel eval file.xlsx "ws = ctx.active"')
                _info(
                    "Example: echo 'print(ctx.active)' | headless-excel eval file.xlsx"
                )
                sys.exit(1)

            if args.code == "-":
                # Check if stdin has data - use isatty() which works cross-platform
                # If stdin is a TTY, no data is being piped in
                if sys.stdin.isatty():
                    _fail("No code provided. Pass code as argument or pipe via stdin.")
                    _info('Example: headless-excel eval file.xlsx "ws = ctx.active"')
                    _info(
                        "Example: echo 'print(ctx.active)' | headless-excel eval file.xlsx"
                    )
                    sys.exit(1)
                code = sys.stdin.read()
            else:
                code = args.code
            with run(args.file, _code=code) as ctx:
                exec(code, {"ctx": ctx, "NumberFormats": NumberFormats})

        case "watch":
            try:
                asyncio.run(
                    watch(
                        args.file,
                        http_port=args.port,
                        ws_port=args.ws_port,
                        open_browser=args.open,
                    )
                )
            except KeyboardInterrupt:
                print("\nStopped")
            except (FileNotFoundError, ValueError) as e:
                _fail(str(e))
                sys.exit(1)

        case "libreoffice" | "lo":
            match args.libre_command:
                case "start":
                    sys.exit(cmd_libreoffice_start())
                case "stop":
                    sys.exit(cmd_libreoffice_stop())
                case "status":
                    sys.exit(cmd_libreoffice_status())


if __name__ == "__main__":
    main()
