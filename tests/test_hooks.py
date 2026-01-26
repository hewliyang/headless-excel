import tempfile
from pathlib import Path

from headless_excel import (
    ExcelContext,
    SyncResult,
    clear_hooks,
    create,
    extension,
    on_exit,
    post_sync,
    pre_sync,
)
from headless_excel.hooks import (
    _load_hooks_from_dir,
    _registry,
    get_discovery_errors,
    get_registry,
)


class TestHookDecorators:
    def test_pre_sync_decorator(self):
        calls = []

        @pre_sync
        def my_hook(ctx: ExcelContext) -> None:
            calls.append("pre_sync")

        registry = get_registry()
        hook_fns = [entry.fn for entry in registry.pre_sync]
        assert my_hook in hook_fns

    def test_post_sync_decorator(self):
        calls = []

        @post_sync
        def my_hook(ctx: ExcelContext, result: SyncResult) -> None:
            calls.append("post_sync")

        registry = get_registry()
        hook_fns = [entry.fn for entry in registry.post_sync]
        assert my_hook in hook_fns

    def test_on_exit_decorator(self):
        calls = []

        @on_exit
        def my_hook(ctx: ExcelContext) -> None:
            calls.append("on_exit")

        registry = get_registry()
        hook_fns = [entry.fn for entry in registry.on_exit]
        assert my_hook in hook_fns


class TestExtensionFactory:
    def test_extension_decorator(self):
        @extension
        def my_extension(api):
            @api.pre_sync
            def hook1(ctx: ExcelContext) -> None:
                pass

            @api.on_exit
            def hook2(ctx: ExcelContext) -> None:
                pass

        registry = get_registry()
        hook_names = [getattr(entry.fn, "__name__", "") for entry in registry.pre_sync]
        assert "hook1" in hook_names
        hook_names = [getattr(entry.fn, "__name__", "") for entry in registry.on_exit]
        assert "hook2" in hook_names

    def test_extension_with_state(self):
        call_count = []

        @extension
        def counter_extension(api):
            count = 0

            @api.pre_sync
            def increment(ctx: ExcelContext) -> None:
                nonlocal count
                count += 1
                call_count.append(count)

        registry = get_registry()
        hook_fn = registry.pre_sync[0].fn

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = Path(f.name)

        try:
            with create(path, overwrite=True, auto_sync=False) as ctx:
                hook_fn(ctx)
                hook_fn(ctx)
        finally:
            path.unlink(missing_ok=True)

        assert call_count == [1, 2]


class TestHookExecution:
    def test_pre_sync_called_before_sync(self, tmp_path: Path):
        calls = []

        @pre_sync
        def record_call(ctx: ExcelContext) -> None:
            calls.append("pre_sync")

        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            assert calls == []
            ctx.sync()
            assert calls == ["pre_sync"]

    def test_post_sync_called_after_sync(self, tmp_path: Path):
        results = []

        @post_sync
        def record_result(ctx: ExcelContext, result: SyncResult) -> None:
            results.append(result.success)

        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.sync()

        assert results == [True]

    def test_on_exit_called_on_context_exit(self, tmp_path: Path):
        calls = []

        @on_exit
        def record_exit(ctx: ExcelContext) -> None:
            calls.append("on_exit")

        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            assert calls == []

        assert calls == ["on_exit"]

    def test_hook_order(self, tmp_path: Path):
        calls = []

        @pre_sync
        def first(ctx: ExcelContext) -> None:
            calls.append("first")

        @pre_sync
        def second(ctx: ExcelContext) -> None:
            calls.append("second")

        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.sync()

        assert calls == ["first", "second"]


class TestHookDiscovery:
    def test_load_hooks_from_dir(self, tmp_path: Path):
        hook_file = tmp_path / "my_hook.py"
        hook_file.write_text("""
from headless_excel import pre_sync, ExcelContext

@pre_sync
def discovered_hook(ctx: ExcelContext) -> None:
    pass
""")

        _load_hooks_from_dir(tmp_path)

        registry = get_registry()
        hook_names = [getattr(entry.fn, "__name__", "") for entry in registry.pre_sync]
        assert "discovered_hook" in hook_names

    def test_skip_private_files(self, tmp_path: Path):
        hook_file = tmp_path / "_private.py"
        hook_file.write_text("""
from headless_excel import pre_sync, ExcelContext

@pre_sync
def should_not_load(ctx: ExcelContext) -> None:
    pass
""")

        _load_hooks_from_dir(tmp_path)

        registry = get_registry()
        hook_names = [getattr(entry.fn, "__name__", "") for entry in registry.pre_sync]
        assert "should_not_load" not in hook_names

    def test_discovery_error_handling(self, tmp_path: Path):
        hook_file = tmp_path / "bad_hook.py"
        hook_file.write_text("this is not valid python {{{{")

        _load_hooks_from_dir(tmp_path)

        errors = get_discovery_errors()
        assert len(errors) > 0
        assert "bad_hook.py" in errors[0]


class TestHookOutputConfiguration:
    def test_default_output_is_stderr(self):
        @pre_sync
        def my_hook(ctx: ExcelContext) -> None:
            pass

        registry = get_registry()
        entry = next(e for e in registry.pre_sync if e.fn is my_hook)
        assert entry.output == "stderr"

    def test_output_stdout(self):
        @pre_sync(output="stdout")
        def my_hook(ctx: ExcelContext) -> None:
            pass

        registry = get_registry()
        entry = next(e for e in registry.pre_sync if e.fn is my_hook)
        assert entry.output == "stdout"

    def test_output_none(self):
        @post_sync(output="none")
        def my_hook(ctx: ExcelContext, result: SyncResult) -> None:
            pass

        registry = get_registry()
        entry = next(e for e in registry.post_sync if e.fn is my_hook)
        assert entry.output == "none"

    def test_extension_api_output_config(self):
        @extension
        def my_extension(api):
            @api.pre_sync(output="stdout")
            def hook1(ctx: ExcelContext) -> None:
                pass

            @api.on_exit(output="none")
            def hook2(ctx: ExcelContext) -> None:
                pass

        registry = get_registry()
        entry1 = next(
            e for e in registry.pre_sync if getattr(e.fn, "__name__", "") == "hook1"
        )
        entry2 = next(
            e for e in registry.on_exit if getattr(e.fn, "__name__", "") == "hook2"
        )
        assert entry1.output == "stdout"
        assert entry2.output == "none"


class TestClearHooks:
    def test_clear_hooks_removes_all(self):
        @pre_sync
        def hook1(ctx: ExcelContext) -> None:
            pass

        @post_sync
        def hook2(ctx: ExcelContext, result: SyncResult) -> None:
            pass

        @on_exit
        def hook3(ctx: ExcelContext) -> None:
            pass

        assert len(_registry.pre_sync) > 0

        clear_hooks()

        assert len(_registry.pre_sync) == 0
        assert len(_registry.post_sync) == 0
        assert len(_registry.on_exit) == 0
