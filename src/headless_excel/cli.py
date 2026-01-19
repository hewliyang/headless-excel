"""CLI for headless-excel."""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from headless_excel import NumberFormats, create, run
from headless_excel.hooks import get_registry
from headless_excel.libre import setup_libreoffice_macro


def _get_hook_name(hook: object) -> str:
    """Get the name of a hook function."""
    return getattr(hook, "__name__", repr(hook))


def cmd_check() -> int:
    """Check if environment is set up correctly."""
    print("headless-excel environment check\n")
    ok = True

    # 1. Check soffice in PATH
    soffice = shutil.which("soffice")
    if soffice:
        print(f"✓ LibreOffice found: {soffice}")

        # Get version
        try:
            result = subprocess.run(
                ["soffice", "--version"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            version = result.stdout.strip()
            print(f"  {version}")
        except Exception:
            pass
    else:
        print("✗ LibreOffice not found in PATH")
        print("  Install it:")
        if sys.platform == "darwin":
            print("    brew install --cask libreoffice")
        else:
            print("    sudo apt install libreoffice libreoffice-calc")
        ok = False

    # 2. Check/setup macro
    if soffice:
        if setup_libreoffice_macro():
            print("✓ LibreOffice macro configured")
        else:
            print("✗ Failed to configure LibreOffice macro")
            ok = False

    # 3. List hooks
    print("\nHooks:")
    local_dir = Path(".headless-excel/hooks")
    global_dir = Path.home() / ".headless-excel" / "hooks"

    if local_dir.is_dir():
        hook_files = sorted(
            f for f in local_dir.rglob("*.py") if not f.name.startswith("_")
        )
        if hook_files:
            print(f"  source: {local_dir.absolute()} (project-local)")
            for f in hook_files:
                print(f"    • {f.relative_to(local_dir)}")
        else:
            print(f"  source: {local_dir.absolute()} (project-local, empty)")
    elif global_dir.is_dir():
        hook_files = sorted(
            f for f in global_dir.rglob("*.py") if not f.name.startswith("_")
        )
        if hook_files:
            print(f"  source: {global_dir} (global)")
            for f in hook_files:
                print(f"    • {f.relative_to(global_dir)}")
        else:
            print(f"  source: {global_dir} (global, empty)")
    else:
        print("  (no hooks directory found)")

    # Show registered hooks
    registry = get_registry()
    total = len(registry.pre_sync) + len(registry.post_sync) + len(registry.on_exit)
    if total > 0:
        print("\n  registered:")
        if registry.pre_sync:
            names = ", ".join(_get_hook_name(h) for h in registry.pre_sync)
            print(f"    pre_sync:  {names}")
        if registry.post_sync:
            names = ", ".join(_get_hook_name(h) for h in registry.post_sync)
            print(f"    post_sync: {names}")
        if registry.on_exit:
            names = ", ".join(_get_hook_name(h) for h in registry.on_exit)
            print(f"    on_exit:   {names}")

    # 4. Quick recalc test
    if soffice and ok:
        print("\nRunning recalc test...")
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                test_file = Path(tmpdir) / "test.xlsx"
                with create(str(test_file), verbose_errors=False) as ctx:
                    ctx.active["A1"] = 2
                    ctx.active["A2"] = 3
                    ctx.active["A3"] = "=A1+A2"
                    ctx.sync()
                    result = ctx.active.cell(3, 1).value

                if result == 5:
                    print("✓ Recalc test passed (2+3=5)")
                else:
                    print(f"✗ Recalc test failed (expected 5, got {result})")
                    ok = False
        except Exception as e:
            print(f"✗ Recalc test failed: {e}")
            ok = False

    print()
    if ok:
        print("All checks passed! Ready to use.")
        return 0
    else:
        print("Some checks failed. See above for details.")
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

    args = parser.parse_args()

    if args.command == "check":
        sys.exit(cmd_check())

    elif args.command == "create":
        with create(args.file) as ctx:
            ctx.sync()
        print(f"Created {args.file}")

    elif args.command == "eval":
        code = sys.stdin.read() if args.code == "-" else args.code

        with run(args.file) as ctx:
            ns = {
                "ctx": ctx,
                "NumberFormats": NumberFormats,
            }
            exec(code, ns)


if __name__ == "__main__":
    main()
