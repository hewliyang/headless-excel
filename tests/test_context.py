"""Tests for ExcelContext and run()."""

from pathlib import Path

import pytest

from headless_excel import ExcelContext, FormulaError, run


class TestExcelContext:
    def test_create_new_workbook(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.workbook.save(path)

        assert path.exists()

    def test_load_existing_workbook(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"

        # Create first
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.workbook.save(path)

        # Load and verify
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == "Hello"

    def test_file_not_found(self, tmp_path: Path):
        path = tmp_path / "nonexistent.xlsx"
        with pytest.raises(FileNotFoundError):
            ExcelContext(path)

    def test_workbook_aliases(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            assert ctx.workbook is ctx.wb

    def test_sheet_operations(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.sheet("Data")["A1"] = "Test"
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert "Data" in ctx.workbook.sheetnames
            assert ctx.sheet("Data")["A1"].value == "Test"

    def test_values_before_sync_raises(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(RuntimeError, match="Call sync"):
                _ = ctx.values

    def test_read_value_before_sync_raises(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            with pytest.raises(RuntimeError, match="Call sync"):
                ctx.read_value("Sheet", "A1")


class TestRun:
    def test_run_create(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with run(path, create=True, auto_sync=False) as ctx:
            ctx.active["A1"] = "Created"
            ctx.workbook.save(path)

        assert path.exists()

    def test_run_auto_sync_false(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with run(path, create=True, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.workbook.save(path)
            # No auto sync, so values not available


class TestSyncResult:
    def test_sync_result_success(self, tmp_path: Path):
        """Test that SyncResult correctly reports success."""
        from headless_excel.context import SyncResult

        result = SyncResult(success=True, total_errors=0, errors={})
        assert result.success
        result.raise_on_errors()  # Should not raise

    def test_sync_result_errors(self):
        """Test that SyncResult correctly reports errors."""
        from headless_excel.context import SyncResult

        result = SyncResult(
            success=False,
            total_errors=2,
            errors={"#REF!": ["Sheet1!A1", "Sheet1!B2"]},
        )
        assert not result.success
        assert result.total_errors == 2

        with pytest.raises(FormulaError) as exc_info:
            result.raise_on_errors()

        assert exc_info.value.total == 2
        assert "#REF!" in exc_info.value.errors


class TestFormulaError:
    def test_formula_error_str(self):
        err = FormulaError(
            errors={"#REF!": ["Sheet1!A1"], "#NAME?": ["Sheet1!B2"]},
            total=2,
        )
        s = str(err)
        assert "2 formula error" in s
        assert "#REF!" in s
        assert "#NAME?" in s
