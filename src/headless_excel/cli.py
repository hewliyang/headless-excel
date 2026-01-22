"""CLI for headless-excel."""

import argparse
import asyncio
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from headless_excel import NumberFormats, create, run
from headless_excel.hooks import get_registry
from headless_excel.libre import setup_libreoffice_macro
from headless_excel.watch import watch

GREEN = "\033[32m"
RED = "\033[31m"
DIM = "\033[2m"
RESET = "\033[0m"


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
    try:
        result = subprocess.run(
            ["soffice", "--version"],
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
    soffice = shutil.which("soffice")
    if soffice:
        _ok(f"LibreOffice found: {soffice}")
        if version := _get_libreoffice_version():
            _info(version)
        return True, soffice

    _fail("LibreOffice not found in PATH")
    _info("Install it:")
    if sys.platform == "darwin":
        _info("brew install --cask libreoffice", indent=2)
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


def cmd_check() -> int:
    """Check if environment is set up correctly."""
    print("headless-excel environment check\n")

    soffice_ok, _ = _check_libreoffice()
    macro_ok = _check_macro() if soffice_ok else False
    _print_hooks_info()
    recalc_ok = _check_recalc() if soffice_ok and macro_ok else True

    print()

    all_ok = soffice_ok and macro_ok and recalc_ok

    if not all_ok:
        print("Some checks failed. See above for details.")
        return 1

    print("All checks passed! Ready to use.")
    return 0


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

    args = parser.parse_args()

    match args.command:
        case "check":
            sys.exit(cmd_check())

        case "create":
            with create(args.file) as ctx:
                ctx.sync()
            print(f"Created {args.file}")

        case "eval":
            code = sys.stdin.read() if args.code == "-" else args.code
            with run(args.file) as ctx:
                exec(code, {"ctx": ctx, "NumberFormats": NumberFormats})

        case "watch":
            try:
                asyncio.run(watch(args.file, http_port=args.port, ws_port=args.ws_port))
            except KeyboardInterrupt:
                print("\nStopped")
            except (FileNotFoundError, ValueError) as e:
                _fail(str(e))
                sys.exit(1)


if __name__ == "__main__":
    main()
