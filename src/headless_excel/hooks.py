"""Extension hook system for headless-excel.

Hooks allow users to run custom code at specific points in the Excel workflow
without modifying the core library. Common use cases include:
- Applying formatting conventions (e.g., financial modeling colors)
- Validation and linting
- Logging and auditing

Discovery:
    Hooks are auto-discovered from Python files in these locations:
    1. ./.headless-excel/hooks/*.py (project-local, takes precedence)
    2. ~/.headless-excel/hooks/*.py (global fallback)

Usage:
    # ~/.headless-excel/hooks/my_hooks.py
    from headless_excel import ExcelContext, SyncResult, pre_sync, post_sync

    @pre_sync
    def before_sync(ctx: ExcelContext) -> None:
        print(f"About to sync {ctx.path}")

    @post_sync
    def after_sync(ctx: ExcelContext, result: SyncResult) -> None:
        print(f"Sync complete, errors: {result.total_errors}")

Output Configuration:
    Hook output is captured and prefixed. By default, output goes to stderr.
    You can configure this per-hook:

    @pre_sync(output="stdout")  # send to stdout instead
    def my_hook(ctx): ...

    @pre_sync(output="none")  # suppress output entirely
    def quiet_hook(ctx): ...
"""

from __future__ import annotations

import importlib.util
import logging
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Generic, Literal, TypeVar, overload

if TYPE_CHECKING:
    from headless_excel.context import ExcelContext, SyncResult

logger = logging.getLogger(__name__)

# ============================================================================
# Type Aliases
# ============================================================================

PreSyncFn = Callable[["ExcelContext"], None]
PostSyncFn = Callable[["ExcelContext", "SyncResult"], None]
OnExitFn = Callable[["ExcelContext"], None]

F = TypeVar("F", PreSyncFn, PostSyncFn, OnExitFn)

HookOutput = Literal["stderr", "stdout", "none"]

T_Hook = TypeVar("T_Hook")


# ============================================================================
# Hook Registry
# ============================================================================


@dataclass
class HookEntry(Generic[T_Hook]):
    """A registered hook with its output configuration."""

    fn: T_Hook
    output: HookOutput = "stderr"


@dataclass
class HookRegistry:
    """Container for registered hooks."""

    pre_sync: list[HookEntry[PreSyncFn]] = field(default_factory=list)
    post_sync: list[HookEntry[PostSyncFn]] = field(default_factory=list)
    on_exit: list[HookEntry[OnExitFn]] = field(default_factory=list)

    _discovered: bool = field(default=False, repr=False)
    _discovery_errors: list[str] = field(default_factory=list, repr=False)

    def clear(self) -> None:
        """Clear all hooks and reset discovery state."""
        self.pre_sync.clear()
        self.post_sync.clear()
        self.on_exit.clear()
        self._discovered = False
        self._discovery_errors.clear()


# Global registry
_registry = HookRegistry()


# ============================================================================
# Decorator API
# ============================================================================


@overload
def pre_sync(fn: PreSyncFn) -> PreSyncFn: ...


@overload
def pre_sync(*, output: HookOutput = "stderr") -> Callable[[PreSyncFn], PreSyncFn]: ...


def pre_sync(
    fn: PreSyncFn | None = None, *, output: HookOutput = "stderr"
) -> PreSyncFn | Callable[[PreSyncFn], PreSyncFn]:
    """Register a pre-sync hook.

    Called before each sync() operation (before save/recalc/reload).
    Use this to modify the workbook before it's saved.

    Args:
        output: Where to send captured output. Options:
            - "stderr" (default): Print to stderr with prefix
            - "stdout": Print to stdout with prefix
            - "none": Suppress output entirely

    Example:
        from headless_excel import ExcelContext, pre_sync

        @pre_sync
        def auto_format(ctx: ExcelContext) -> None:
            for ws in ctx.workbook.worksheets:
                ws.auto_financial_colors()

        @pre_sync(output="stdout")  # send to stdout instead
        def verbose_hook(ctx: ExcelContext) -> None:
            print("Processing...")
    """

    def decorator(f: PreSyncFn) -> PreSyncFn:
        _registry.pre_sync.append(HookEntry(f, output))
        return f

    if fn is not None:
        return decorator(fn)
    return decorator


@overload
def post_sync(fn: PostSyncFn) -> PostSyncFn: ...


@overload
def post_sync(
    *, output: HookOutput = "stderr"
) -> Callable[[PostSyncFn], PostSyncFn]: ...


def post_sync(
    fn: PostSyncFn | None = None, *, output: HookOutput = "stderr"
) -> PostSyncFn | Callable[[PostSyncFn], PostSyncFn]:
    """Register a post-sync hook.

    Called after each sync() operation (after save/recalc/reload).
    Use this to inspect results or perform validation.

    Args:
        output: Where to send captured output. Options:
            - "stderr" (default): Print to stderr with prefix
            - "stdout": Print to stdout with prefix
            - "none": Suppress output entirely

    Example:
        from headless_excel import ExcelContext, SyncResult, post_sync

        @post_sync
        def log_errors(ctx: ExcelContext, result: SyncResult) -> None:
            if not result.success:
                print(f"Sync had {result.total_errors} errors")
    """

    def decorator(f: PostSyncFn) -> PostSyncFn:
        _registry.post_sync.append(HookEntry(f, output))
        return f

    if fn is not None:
        return decorator(fn)
    return decorator


@overload
def on_exit(fn: OnExitFn) -> OnExitFn: ...


@overload
def on_exit(*, output: HookOutput = "stderr") -> Callable[[OnExitFn], OnExitFn]: ...


def on_exit(
    fn: OnExitFn | None = None, *, output: HookOutput = "stderr"
) -> OnExitFn | Callable[[OnExitFn], OnExitFn]:
    """Register an on-exit hook.

    Called when the context manager exits (after final sync if any).
    Use this for linting, cleanup, or final validation.

    Args:
        output: Where to send captured output. Options:
            - "stderr" (default): Print to stderr with prefix
            - "stdout": Print to stdout with prefix
            - "none": Suppress output entirely

    Example:
        from headless_excel import ExcelContext, on_exit

        @on_exit
        def lint_colors(ctx: ExcelContext) -> None:
            result = ctx.lint_financial_colors()
            if result.total_violations > 0:
                print(result)
    """

    def decorator(f: OnExitFn) -> OnExitFn:
        _registry.on_exit.append(HookEntry(f, output))
        return f

    if fn is not None:
        return decorator(fn)
    return decorator


# ============================================================================
# Factory API (for stateful extensions)
# ============================================================================


class ExtensionAPI:
    """API for extension factory functions.

    Use this pattern when you need shared state across hooks.

    Example:
        from headless_excel import ExcelContext, extension

        @extension
        def my_extension(api):
            sync_count = 0

            @api.pre_sync
            def count_syncs(ctx: ExcelContext) -> None:
                nonlocal sync_count
                sync_count += 1
                print(f"Sync #{sync_count}")

            @api.post_sync(output="stdout")  # configure output
            def log_result(ctx: ExcelContext, result) -> None:
                print(f"Done: {result.success}")
    """

    @overload
    def pre_sync(self, fn: PreSyncFn) -> PreSyncFn: ...

    @overload
    def pre_sync(
        self, *, output: HookOutput = "stderr"
    ) -> Callable[[PreSyncFn], PreSyncFn]: ...

    def pre_sync(
        self, fn: PreSyncFn | None = None, *, output: HookOutput = "stderr"
    ) -> PreSyncFn | Callable[[PreSyncFn], PreSyncFn]:
        """Register a pre-sync hook.

        Args:
            output: Where to send output ("stderr", "stdout", or "none")
        """

        def decorator(f: PreSyncFn) -> PreSyncFn:
            _registry.pre_sync.append(HookEntry(f, output))
            return f

        if fn is not None:
            return decorator(fn)
        return decorator

    @overload
    def post_sync(self, fn: PostSyncFn) -> PostSyncFn: ...

    @overload
    def post_sync(
        self, *, output: HookOutput = "stderr"
    ) -> Callable[[PostSyncFn], PostSyncFn]: ...

    def post_sync(
        self, fn: PostSyncFn | None = None, *, output: HookOutput = "stderr"
    ) -> PostSyncFn | Callable[[PostSyncFn], PostSyncFn]:
        """Register a post-sync hook.

        Args:
            output: Where to send output ("stderr", "stdout", or "none")
        """

        def decorator(f: PostSyncFn) -> PostSyncFn:
            _registry.post_sync.append(HookEntry(f, output))
            return f

        if fn is not None:
            return decorator(fn)
        return decorator

    @overload
    def on_exit(self, fn: OnExitFn) -> OnExitFn: ...

    @overload
    def on_exit(
        self, *, output: HookOutput = "stderr"
    ) -> Callable[[OnExitFn], OnExitFn]: ...

    def on_exit(
        self, fn: OnExitFn | None = None, *, output: HookOutput = "stderr"
    ) -> OnExitFn | Callable[[OnExitFn], OnExitFn]:
        """Register an on-exit hook.

        Args:
            output: Where to send output ("stderr", "stdout", or "none")
        """

        def decorator(f: OnExitFn) -> OnExitFn:
            _registry.on_exit.append(HookEntry(f, output))
            return f

        if fn is not None:
            return decorator(fn)
        return decorator


ExtensionFactory = Callable[[ExtensionAPI], None]


def extension(factory_fn: ExtensionFactory) -> ExtensionFactory:
    """Decorator for extension factory functions.

    Example:
        @extension
        def my_extension(api):
            @api.pre_sync
            def hook(ctx):
                pass
    """
    api = ExtensionAPI()
    factory_fn(api)
    return factory_fn


# ============================================================================
# Discovery
# ============================================================================


def _load_hook_file(path: Path) -> None:
    """Load and execute a single hook file."""
    try:
        spec = importlib.util.spec_from_file_location(
            f"headless_excel_hook_{path.stem}_{id(path)}", path
        )
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            logger.debug(f"Loaded hook file: {path}")
    except Exception as e:
        error_msg = f"{path}: {e}"
        _registry._discovery_errors.append(error_msg)
        logger.warning(f"Failed to load hook file {path}: {e}")


def _load_hooks_from_dir(directory: Path) -> None:
    """Recursively load all .py files from a directory."""
    if not directory.exists():
        return

    for f in sorted(directory.rglob("*.py")):
        # Skip private/dunder files
        if f.name.startswith("_"):
            continue
        _load_hook_file(f)


def discover_hooks() -> None:
    """Auto-discover and load hooks from standard locations.

    Discovery order (first match wins, no merging):
    1. ./.headless-excel/hooks/ (project-local)
    2. ~/.headless-excel/hooks/ (global)
    """
    if _registry._discovered:
        return
    _registry._discovered = True

    # Project-local takes precedence
    local_dir = Path(".headless-excel/hooks")
    if local_dir.is_dir():
        logger.debug(f"Discovering hooks from project-local: {local_dir}")
        _load_hooks_from_dir(local_dir)
        return  # Project-local overrides global entirely

    # Fall back to global
    global_dir = Path.home() / ".headless-excel" / "hooks"
    if global_dir.is_dir():
        logger.debug(f"Discovering hooks from global: {global_dir}")
        _load_hooks_from_dir(global_dir)


# ============================================================================
# Public API for context.py
# ============================================================================


def get_registry() -> HookRegistry:
    """Get the hook registry, discovering hooks if needed."""
    discover_hooks()
    return _registry


def _get_hook_name(hook: object) -> str:
    """Get the name of a hook for logging purposes."""
    return getattr(hook, "__name__", repr(hook))


def _run_hook_with_prefix(
    hook: object,
    run_fn: Callable[[], None],
    hook_type: str,
    output: HookOutput = "stderr",
) -> None:
    """Run a hook, capturing stdout and prefixing each line.

    Args:
        hook: The hook function (for getting name)
        run_fn: Zero-arg function that actually calls the hook
        hook_type: Type of hook for prefix (e.g., "pre-sync", "post-sync", "on-exit")
        output: Where to send captured output ("stderr", "stdout", or "none")
    """
    import io
    import sys

    hook_name = _get_hook_name(hook)
    prefix = f"[{hook_type}:{hook_name}]"

    # Capture stdout
    old_stdout = sys.stdout
    captured = io.StringIO()
    sys.stdout = captured

    try:
        run_fn()
    finally:
        sys.stdout = old_stdout

    # Print captured output with prefix to configured destination
    if output == "none":
        return

    captured_output = captured.getvalue()
    if captured_output:
        dest = sys.stderr if output == "stderr" else sys.stdout
        for line in captured_output.rstrip("\n").split("\n"):
            print(f"{prefix} {line}", file=dest)


def run_pre_sync_hooks(ctx: ExcelContext) -> None:
    """Run all pre-sync hooks."""
    registry = get_registry()
    for entry in registry.pre_sync:
        try:
            _run_hook_with_prefix(
                entry.fn, lambda e=entry: e.fn(ctx), "pre-sync", entry.output
            )
        except Exception as e:
            logger.error(f"Pre-sync hook {_get_hook_name(entry.fn)} failed: {e}")
            raise


def run_post_sync_hooks(ctx: ExcelContext, result: SyncResult) -> None:
    """Run all post-sync hooks."""
    registry = get_registry()
    for entry in registry.post_sync:
        try:
            _run_hook_with_prefix(
                entry.fn, lambda e=entry: e.fn(ctx, result), "post-sync", entry.output
            )
        except Exception as e:
            logger.error(f"Post-sync hook {_get_hook_name(entry.fn)} failed: {e}")
            raise


def run_on_exit_hooks(ctx: ExcelContext) -> None:
    """Run all on-exit hooks."""
    registry = get_registry()
    for entry in registry.on_exit:
        try:
            _run_hook_with_prefix(
                entry.fn, lambda e=entry: e.fn(ctx), "on-exit", entry.output
            )
        except Exception as e:
            logger.error(f"On-exit hook {_get_hook_name(entry.fn)} failed: {e}")
            raise


def clear_hooks() -> None:
    """Clear all hooks and reset discovery state.

    Useful for testing.
    """
    _registry.clear()


def get_discovery_errors() -> list[str]:
    """Get list of errors encountered during hook discovery."""
    return _registry._discovery_errors.copy()
