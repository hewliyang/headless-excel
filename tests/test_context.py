"""Tests for ExcelContext, run(), and create()."""

from pathlib import Path

import pytest
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Side,
)

from headless_excel import (
    ErrorDetail,
    ExcelContext,
    FormulaError,
    SyncResult,
    create,
    run,
)
from headless_excel.proxy import CellProxy


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
            assert ctx.active["A1"].value == "Hello"  # type: ignore[union-attr]

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
            assert ctx.sheet("Data")["A1"].value == "Test"  # type: ignore[union-attr]

    def test_values_before_sync_raises(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(RuntimeError, match="Call sync"):
                _ = ctx.values


class TestRun:
    def test_run_opens_existing(self, tmp_path: Path):
        """Test that run() opens an existing file."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].value == "Hello"  # type: ignore[union-attr]

    def test_run_file_not_found(self, tmp_path: Path):
        """Test that run() raises FileNotFoundError for missing file."""
        path = tmp_path / "nonexistent.xlsx"
        with pytest.raises(FileNotFoundError):
            with run(path) as _:
                pass

    def test_run_auto_sync_false(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.workbook.save(path)
            # No auto sync, so values not available


class TestCreate:
    def test_create_new_file(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = "Created"
            ctx.workbook.save(path)

        assert path.exists()

    def test_create_file_exists_raises(self, tmp_path: Path):
        """Test that create() raises FileExistsError when file exists."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.workbook.save(path)

        with pytest.raises(FileExistsError, match="already exists"):
            with create(path) as ctx:
                pass

    def test_create_overwrite(self, tmp_path: Path):
        """Test that create(overwrite=True) replaces existing file."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = "Original"
            ctx.workbook.save(path)

        with create(path, overwrite=True, auto_sync=False) as ctx:
            ctx.active["A1"] = "Replaced"
            ctx.workbook.save(path)

        with run(path) as ctx:
            assert ctx.active["A1"].value == "Replaced"  # type: ignore[union-attr]

    def test_create_manual_sync_prints_errors_on_exit(self, tmp_path: Path, capsys):
        """Test that manual sync with errors prints to stderr on exit.

        This ensures agents can't miss errors by ignoring sync() return value.
        Errors are always printed at context exit when verbose_errors=True.
        File is still saved successfully - these are just formula errors.
        """
        path = tmp_path / "test.xlsx"

        # Should complete without exception - file is saved
        with create(path, verbose_errors=True) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A2"] = "=A1+Sheet2!A1"

            result = ctx.sync(raise_on_errors=False)
            assert not result.success
            assert result.total_errors == 1

        # Errors printed to stderr
        captured = capsys.readouterr()
        assert "#NAME?" in captured.err
        # File was saved
        assert path.exists()

    def test_create_verbose_errors_prints_to_stderr(self, tmp_path: Path, capsys):
        """Test that verbose_errors=True prints errors to stderr."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=True) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A2"] = "=A1+Sheet2!A1"

        # Errors printed to stderr, file still saved
        captured = capsys.readouterr()
        assert "SyncResult(success=False" in captured.err
        assert "#NAME?" in captured.err
        assert path.exists()

    def test_create_verbose_errors_false(self, tmp_path: Path, capsys):
        """Test that verbose_errors=False suppresses error output."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A2"] = "=A1+Sheet2!A1"

        # No errors printed
        captured = capsys.readouterr()
        assert "#NAME?" not in captured.err
        assert path.exists()


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
        assert "Formula errors (2):" in s
        assert "#REF!" in s
        assert "#NAME?" in s

    def test_formula_error_str_with_details(self):
        err = FormulaError(
            errors={"#DIV/0!": ["Sheet1!A3"]},
            total=1,
            error_details=[
                ErrorDetail(
                    location="Sheet1!A3",
                    error="#DIV/0!",
                    formula="=A1/A2",
                    neighbors={"Sheet1!A1": 10, "Sheet1!A2": 0},
                )
            ],
        )
        s = str(err)
        assert "Formula errors (1):" in s
        assert "Sheet1!A3: #DIV/0!" in s
        assert "formula==A1/A2" in s
        assert "inputs={" in s
        assert "Sheet1!A1=10" in s
        assert "Sheet1!A2=0" in s


class TestRangeProxy:
    """Tests for the RangeProxy class and sheet.range() method."""

    def test_range_values_read(self, tmp_path: Path):
        """Test reading values from a range."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 1
            ctx.active["B1"] = 2
            ctx.active["A2"] = 3
            ctx.active["B2"] = 4
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            r = ctx.active.range("A1:B2")
            assert r.values == [[1, 2], [3, 4]]

    def test_range_values_write(self, tmp_path: Path):
        """Test writing values to a range."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:C2")
            r.values = [[1, 2, 3], [4, 5, 6]]
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["C2"].value == 6  # type: ignore[union-attr]

    def test_range_values_write_row_mismatch_lenient(self, tmp_path: Path, capsys):
        """Test that row count mismatch is allowed (lenient behavior).

        Range becomes an anchor - data shape determines actual write region.
        Mismatch is reported to stderr but not an error.
        """
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")  # specified 2 rows
            r.values = [[1, 2]]  # only 1 row - allowed
            ctx.workbook.save(path)

        # Verify what was written
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["B1"].value == 2  # type: ignore[union-attr]
            assert ctx.active["A2"].value is None  # type: ignore[union-attr]

        # Check stderr feedback
        captured = capsys.readouterr()
        assert "[headless-excel] .values wrote to A1:B1" in captured.err
        assert "specified A1:B2" in captured.err
        assert "got 1×2" in captured.err

    def test_range_values_write_col_mismatch_lenient(self, tmp_path: Path, capsys):
        """Test that column count mismatch is allowed (lenient behavior)."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:C2")  # specified 3 cols
            r.values = [[1, 2], [3, 4]]  # only 2 cols - allowed
            ctx.workbook.save(path)

        # Verify what was written
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["B2"].value == 4  # type: ignore[union-attr]
            assert ctx.active["C1"].value is None  # type: ignore[union-attr]

        # Check stderr feedback
        captured = capsys.readouterr()
        assert "[headless-excel] .values wrote to A1:B2" in captured.err
        assert "specified A1:C2" in captured.err
        assert "got 2×2" in captured.err

    def test_range_values_write_overflow(self, tmp_path: Path, capsys):
        """Test that data larger than range is allowed (overflow)."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")  # specified 2x2
            r.values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]  # 3x3 - overflow allowed
            ctx.workbook.save(path)

        # Verify all data was written
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["C3"].value == 9  # type: ignore[union-attr]

        # Check stderr feedback
        captured = capsys.readouterr()
        assert "[headless-excel] .values wrote to A1:C3" in captured.err
        assert "specified A1:B2" in captured.err
        assert "got 3×3" in captured.err

    def test_range_values_write_exact_match(self, tmp_path: Path, capsys):
        """Test that exact match reports without 'specified' note."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")  # 2x2
            r.values = [[1, 2], [3, 4]]  # exact 2x2
            ctx.workbook.save(path)

        captured = capsys.readouterr()
        assert "[headless-excel] .values wrote to A1:B2 (2×2)" in captured.err
        assert "specified" not in captured.err  # no mismatch note

    def test_range_values_write_not_list(self, tmp_path: Path):
        """Test that non-list data raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")
            with pytest.raises(ValueError, match="must be a 2D list"):
                r.values = "not a list"  # type: ignore[assignment]

    def test_range_values_write_row_not_list(self, tmp_path: Path):
        """Test that non-list row raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")
            with pytest.raises(ValueError, match="Row 0 must be a list"):
                r.values = ["not", "nested"]  # type: ignore[assignment]

    def test_range_shape(self, tmp_path: Path):
        """Test range shape property."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("B3:E6")
            assert r.shape == (4, 4)
            assert r.num_rows == 4
            assert r.num_cols == 4

    def test_range_single_cell(self, tmp_path: Path):
        """Test range with single cell reference."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("C5")
            assert r.shape == (1, 1)
            r.values = [[42]]
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["C5"].value == 42  # type: ignore[union-attr]

    def test_range_formulas(self, tmp_path: Path):
        """Test reading formulas from a range."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.active["B1"] = "=A1*2"
            ctx.active["A2"] = "=A1+5"
            ctx.active["B2"] = 30

            r = ctx.active.range("A1:B2")
            formulas = r.formulas
            # Returns dict, not 2D array (more token-efficient for sparse formulas)
            assert formulas == {"B1": "=A1*2", "A2": "=A1+5"}

    def test_range_values_write_with_formulas(self, tmp_path: Path):
        """Test writing formulas via range.values."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            # Write mix of values and formulas
            ctx.active.range("A1:C2").values = [
                [10, 20, "=A1+B1"],
                [30, 40, "=A2+B2"],
            ]
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            # Values are materialized (10, 20 are numbers)
            assert ctx.active["A1"].value == 10  # type: ignore[union-attr]
            assert ctx.active["B2"].value == 40  # type: ignore[union-attr]
            # Formulas are preserved and accessible via .formulas
            assert ctx.active.formulas == {"C1": "=A1+B1", "C2": "=A2+B2"}
            # Range formulas returns dict (same format as sheet.formulas)
            assert ctx.active.range("C1:C2").formulas == {
                "C1": "=A1+B1",
                "C2": "=A2+B2",
            }

    def test_range_invalid_reference(self, tmp_path: Path):
        """Test that invalid range reference raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="Invalid range reference"):
                ctx.active.range("not_a_range")

    def test_range_repr(self, tmp_path: Path):
        """Test range repr."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")
            assert "RangeProxy" in repr(r)
            assert "A1:B2" in repr(r)

    def test_range_values_empty_data(self, tmp_path: Path, capsys):
        """Test that empty data list prints message but doesn't error."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            r = ctx.active.range("A1:B2")
            r.values = []  # empty list

        captured = capsys.readouterr()
        assert "[headless-excel] .values no data to write" in captured.err


class TestWorksheetWrite:
    """Tests for WorksheetProxy.write() method."""

    def test_write_basic(self, tmp_path: Path, capsys):
        """Test basic write operation."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            result = ctx.active.write("A1", [[1, 2, 3], [4, 5, 6]])
            ctx.workbook.save(path)

            assert result == "A1:C2"

        # Verify written values
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["C2"].value == 6  # type: ignore[union-attr]

        captured = capsys.readouterr()
        assert "[headless-excel] .write() wrote to A1:C2 (2×3)" in captured.err

    def test_write_single_cell(self, tmp_path: Path, capsys):
        """Test writing a single value."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            result = ctx.active.write("B5", [[42]])
            ctx.workbook.save(path)

            assert result == "B5"

        # Verify
        with ExcelContext(path) as ctx:
            assert ctx.active["B5"].value == 42  # type: ignore[union-attr]

        captured = capsys.readouterr()
        assert "[headless-excel] .write() wrote to B5 (1×1)" in captured.err

    def test_write_with_formulas(self, tmp_path: Path, capsys):
        """Test writing mix of values and formulas."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            result = ctx.active.write(
                "A1",
                [
                    [10, 20, "=A1+B1"],
                    [30, 40, "=A2+B2"],
                ],
            )
            ctx.workbook.save(path)

            assert result == "A1:C2"

        # Verify
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 10  # type: ignore[union-attr]
            assert ctx.active.formulas == {"C1": "=A1+B1", "C2": "=A2+B2"}

    def test_write_marks_dirty(self, tmp_path: Path):
        """Test that write() marks context as dirty."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.sync()
            assert not ctx._dirty

            ctx.active.write("A1", [[1]])
            assert ctx._dirty

    def test_write_rejects_range(self, tmp_path: Path):
        """Test that write() rejects range references."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="single cell anchor"):
                ctx.active.write("A1:B2", [[1, 2], [3, 4]])

    def test_write_rejects_non_list(self, tmp_path: Path):
        """Test that write() rejects non-list data."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="must be a 2D list"):
                ctx.active.write("A1", "not a list")  # type: ignore[arg-type]

    def test_write_rejects_non_list_rows(self, tmp_path: Path):
        """Test that write() rejects rows that are not lists."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="Row 0 must be a list"):
                ctx.active.write("A1", ["not", "nested"])  # type: ignore[list-item]

    def test_write_empty_data(self, tmp_path: Path, capsys):
        """Test write with empty data."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            result = ctx.active.write("A1", [])

            assert result == "A1"

        captured = capsys.readouterr()
        assert "[headless-excel] .write() no data to write" in captured.err

    def test_write_jagged_arrays(self, tmp_path: Path, capsys):
        """Test write with jagged (uneven) arrays."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            # Rows have different lengths
            result = ctx.active.write(
                "A1",
                [
                    [1, 2, 3],
                    [4, 5],  # shorter row
                    [7, 8, 9, 10],  # longer row
                ],
            )
            ctx.workbook.save(path)

            # Should use max width (4 cols)
            assert result == "A1:D3"

        # Verify
        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value == 1  # type: ignore[union-attr]
            assert ctx.active["D3"].value == 10  # type: ignore[union-attr]
            # Shorter row doesn't touch column D
            assert ctx.active["C2"].value is None  # type: ignore[union-attr]

        captured = capsys.readouterr()
        assert "[headless-excel] .write() wrote to A1:D3 (3×4)" in captured.err


class TestRangeApplyStyle:
    """Tests for RangeProxy.apply_style()."""

    def test_apply_font_to_single_cell(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.active.range("A1").apply_style(font=Font(bold=True, color="FF0000"))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell: CellProxy = ctx.active["A1"]  # type: ignore[assignment]
            assert cell.font.bold is True
            assert cell.font.color.rgb == "00FF0000"  # type: ignore[union-attr]

    def test_apply_font_to_range(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            for col in ["A", "B", "C"]:
                ctx.active[f"{col}1"] = f"Value {col}"
            ctx.active.range("A1:C1").apply_style(font=Font(italic=True))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            for col in ["A", "B", "C"]:
                assert ctx.active[f"{col}1"].font.italic is True  # type: ignore[union-attr]

    def test_apply_fill(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Colored"
            fill = PatternFill(
                start_color="FFFF00", end_color="FFFF00", fill_type="solid"
            )
            ctx.active.range("A1").apply_style(fill=fill)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].fill.start_color.rgb == "00FFFF00"  # type: ignore[union-attr]

    def test_apply_gradient_fill(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Gradient"
            gradient = GradientFill(stop=["FF0000", "0000FF"])
            ctx.active.range("A1").apply_style(gradient_fill=gradient)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell: CellProxy = ctx.active["A1"]  # type: ignore[assignment]
            assert cell.fill.type == "linear"  # type: ignore[union-attr]
            assert len(cell.fill.stop) == 2  # type: ignore[union-attr,arg-type]
            assert cell.fill.stop[0].color.rgb == "00FF0000"  # type: ignore[union-attr,index]
            assert cell.fill.stop[1].color.rgb == "000000FF"  # type: ignore[union-attr,index]

    def test_apply_gradient_fill_to_range(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            for col in ["A", "B", "C"]:
                ctx.active[f"{col}1"] = f"Gradient {col}"
            gradient = GradientFill(stop=["00FF00", "FFFF00"], degree=90)
            ctx.active.range("A1:C1").apply_style(gradient_fill=gradient)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            for col in ["A", "B", "C"]:
                cell: CellProxy = ctx.active[f"{col}1"]  # type: ignore[assignment]
                assert cell.fill.type == "linear"  # type: ignore[union-attr]
                assert cell.fill.degree == 90  # type: ignore[union-attr]
                assert cell.fill.stop[0].color.rgb == "0000FF00"  # type: ignore[union-attr,index]
                assert cell.fill.stop[1].color.rgb == "00FFFF00"  # type: ignore[union-attr,index]

    def test_gradient_fill_takes_precedence_over_fill(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Both fills"
            pattern = PatternFill(start_color="FF0000", fill_type="solid")
            gradient = GradientFill(stop=["00FF00", "0000FF"])
            ctx.active.range("A1").apply_style(fill=pattern, gradient_fill=gradient)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell: CellProxy = ctx.active["A1"]  # type: ignore[assignment]
            assert cell.fill.type == "linear"  # type: ignore[union-attr]
            assert len(cell.fill.stop) == 2  # type: ignore[union-attr,arg-type]

    def test_apply_number_format(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 1234.56
            ctx.active["A2"] = -789.12
            ctx.active.range("A1:A2").apply_style(number_format="#,##0.00;(#,##0.00)")
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].number_format == "#,##0.00;(#,##0.00)"  # type: ignore[union-attr]
            assert ctx.active["A2"].number_format == "#,##0.00;(#,##0.00)"  # type: ignore[union-attr]

    def test_apply_alignment(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Centered"
            ctx.active.range("A1").apply_style(
                alignment=Alignment(horizontal="center", vertical="top")
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].alignment.horizontal == "center"  # type: ignore[union-attr]
            assert ctx.active["A1"].alignment.vertical == "top"  # type: ignore[union-attr]

    def test_apply_border(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Bordered"
            thin_side = Side(style="thin", color="000000")
            border = Border(
                left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
            )
            ctx.active.range("A1").apply_style(border=border)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].border.left.style == "thin"  # type: ignore[union-attr]
            assert ctx.active["A1"].border.right.style == "thin"  # type: ignore[union-attr]

    def test_apply_protection(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Protected"
            ctx.active.range("A1").apply_style(
                protection=Protection(locked=True, hidden=True)
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].protection.locked is True  # type: ignore[union-attr]
            assert ctx.active["A1"].protection.hidden is True  # type: ignore[union-attr]

    def test_apply_multiple_styles(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["B2"] = 42
            ctx.active.range("B2").apply_style(
                font=Font(bold=True),
                fill=PatternFill(start_color="00FF00", fill_type="solid"),
                number_format="0.00",
                alignment=Alignment(horizontal="right"),
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell: CellProxy = ctx.active["B2"]  # type: ignore[assignment]
            assert cell.font.bold is True
            assert cell.fill.start_color.rgb == "0000FF00"  # type: ignore[union-attr]
            assert cell.number_format == "0.00"
            assert cell.alignment.horizontal == "right"

    def test_apply_style_to_2d_range(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            for row in range(1, 4):
                for col in ["A", "B", "C"]:
                    ctx.active[f"{col}{row}"] = f"{col}{row}"
            ctx.active.range("A1:C3").apply_style(font=Font(size=14))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            for row in range(1, 4):
                for col in ["A", "B", "C"]:
                    assert ctx.active[f"{col}{row}"].font.size == 14  # type: ignore[union-attr]


class TestRangeClear:
    """Tests for RangeProxy.clear() method."""

    def test_clear_values(self, tmp_path: Path):
        """Test clearing values only."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active.range("A1:C3").values = [
                [1, 2, 3],
                [4, 5, 6],
                [7, 8, 9],
            ]
            ctx.active.range("A1:C3").clear()

            # All values should be None
            values = ctx.active.range("A1:C3").values
            assert values == [
                [None, None, None],
                [None, None, None],
                [None, None, None],
            ]

    def test_clear_preserves_styles_when_styles_false(self, tmp_path: Path):
        """Test that clear() preserves styles when styles=False."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Test"
            ctx.active.range("A1").apply_style(font=Font(bold=True))
            ctx.active.range("A1").clear(styles=False)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].value is None  # type: ignore[union-attr]
            assert ctx.active["A1"].font.bold is True  # type: ignore[union-attr]

    def test_clear_with_styles(self, tmp_path: Path):
        """Test clearing values and styles."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Test"
            ctx.active.range("A1").apply_style(
                font=Font(bold=True),
                fill=PatternFill(start_color="FF0000", fill_type="solid"),
                number_format="0.00",
            )
            ctx.active.range("A1").clear(styles=True)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell: CellProxy = ctx.active["A1"]  # type: ignore[assignment]
            assert cell.value is None
            # Styles should be reset to defaults
            assert cell.font.bold is not True
            assert cell.number_format == "General"

    def test_clear_formulas(self, tmp_path: Path):
        """Test that clear() removes formulas."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.active["A2"] = "=A1*2"
            ctx.active.range("A2").clear()

            assert ctx.active["A2"].formula is None  # type: ignore[union-attr]
            assert ctx.active["A2"].value is None  # type: ignore[union-attr]

    def test_clear_marks_dirty(self, tmp_path: Path):
        """Test that clear() marks context as dirty."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.sync()
            assert not ctx._dirty

            ctx.active.range("A1").clear()
            assert ctx._dirty

    def test_clear_unmerges_merged_cells(self, tmp_path: Path):
        """Test that clear() unmerges merged cells before clearing."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "Merged Title"
            ws._formula_ws.merge_cells("A1:C1")
            ws["A2"] = "Data"

            # This should not raise an error
            ws.range("A1:C2").clear()

            # Values should be cleared
            assert ws.cell(1, 1).value is None
            assert ws.cell(2, 1).value is None

            # Merged range should be removed
            assert len(list(ws._formula_ws.merged_cells.ranges)) == 0  # type: ignore[arg-type]

    def test_clear_unmerges_only_overlapping_ranges(self, tmp_path: Path):
        """Test that clear() only unmerges ranges that overlap with the clear range."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "Merge 1"
            ws._formula_ws.merge_cells("A1:B1")
            ws["D1"] = "Merge 2"
            ws._formula_ws.merge_cells("D1:E1")

            # Clear only the first merged range
            ws.range("A1:B2").clear()

            # First merge should be gone
            assert ws.cell(1, 1).value is None

            # Second merge should still exist
            merged_ranges = [str(r) for r in ws._formula_ws.merged_cells.ranges]  # type: ignore[union-attr]
            assert "D1:E1" in merged_ranges
            assert "A1:B1" not in merged_ranges

    def test_clear_partial_overlap_unmerges(self, tmp_path: Path):
        """Test that clear() unmerges when range partially overlaps merged cells."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "Wide Merge"
            ws._formula_ws.merge_cells("A1:D1")

            # Clear only part of the merged range
            ws.range("B1:C1").clear()

            # Merge should be removed (can't have partial merge)
            assert len(list(ws._formula_ws.merged_cells.ranges)) == 0  # type: ignore[arg-type]

    def test_clear_with_styles_unmerges(self, tmp_path: Path):
        """Test that clear(styles=True) also handles merged cells."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "Styled Merge"
            ws._formula_ws.merge_cells("A1:C1")
            ws.range("A1").apply_style(font=Font(bold=True))

            ws.range("A1:C1").clear(styles=True)

            assert ws.cell(1, 1).value is None
            assert ws.cell(1, 1).font.bold is not True
            assert len(list(ws._formula_ws.merged_cells.ranges)) == 0  # type: ignore[arg-type]


class TestAutoFill:
    """Tests for RangeProxy.auto_fill() method."""

    def test_auto_fill_formula_down(self, tmp_path: Path):
        """Test basic formula fill downward."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=B1*C1"
            ctx.active.range("A1:A5").auto_fill()

            assert ctx.active.get_cell("A1").formula == "=B1*C1"
            assert ctx.active.get_cell("A2").formula == "=B2*C2"
            assert ctx.active.get_cell("A3").formula == "=B3*C3"
            assert ctx.active.get_cell("A4").formula == "=B4*C4"
            assert ctx.active.get_cell("A5").formula == "=B5*C5"

    def test_auto_fill_formula_right(self, tmp_path: Path):
        """Test formula fill rightward."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=A2+A3"
            ctx.active.range("A1:D1").auto_fill(direction="right")

            assert ctx.active.get_cell("A1").formula == "=A2+A3"
            assert ctx.active.get_cell("B1").formula == "=B2+B3"
            assert ctx.active.get_cell("C1").formula == "=C2+C3"
            assert ctx.active.get_cell("D1").formula == "=D2+D3"

    def test_auto_fill_formula_up(self, tmp_path: Path):
        """Test formula fill upward."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A5"] = "=B5*2"
            ctx.active.range("A1:A5").auto_fill(direction="up")

            assert ctx.active.get_cell("A5").formula == "=B5*2"
            assert ctx.active.get_cell("A4").formula == "=B4*2"
            assert ctx.active.get_cell("A3").formula == "=B3*2"
            assert ctx.active.get_cell("A2").formula == "=B2*2"
            assert ctx.active.get_cell("A1").formula == "=B1*2"

    def test_auto_fill_formula_left(self, tmp_path: Path):
        """Test formula fill leftward."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["D1"] = "=D2*2"
            ctx.active.range("A1:D1").auto_fill(direction="left")

            assert ctx.active.get_cell("D1").formula == "=D2*2"
            assert ctx.active.get_cell("C1").formula == "=C2*2"
            assert ctx.active.get_cell("B1").formula == "=B2*2"
            assert ctx.active.get_cell("A1").formula == "=A2*2"

    def test_auto_fill_absolute_references(self, tmp_path: Path):
        """Test that absolute references stay fixed."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=$B$1*C1"
            ctx.active.range("A1:A3").auto_fill()

            assert ctx.active.get_cell("A1").formula == "=$B$1*C1"
            assert ctx.active.get_cell("A2").formula == "=$B$1*C2"
            assert ctx.active.get_cell("A3").formula == "=$B$1*C3"

    def test_auto_fill_mixed_references(self, tmp_path: Path):
        """Test that mixed references adjust correctly."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=$B1*C$1"
            ctx.active.range("A1:A3").auto_fill()

            assert ctx.active.get_cell("A1").formula == "=$B1*C$1"
            assert ctx.active.get_cell("A2").formula == "=$B2*C$1"
            assert ctx.active.get_cell("A3").formula == "=$B3*C$1"

    def test_auto_fill_values(self, tmp_path: Path):
        """Test that non-formula values are copied as-is."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            ctx.active.range("A1:A5").auto_fill()

            for row in range(1, 6):
                assert ctx.active.get_cell(f"A{row}").value == 100

    def test_auto_fill_cross_sheet_reference(self, tmp_path: Path):
        """Test that cross-sheet references adjust correctly."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=Sheet2!A1+B1"
            ctx.active.range("A1:A3").auto_fill()

            assert ctx.active.get_cell("A1").formula == "=Sheet2!A1+B1"
            assert ctx.active.get_cell("A2").formula == "=Sheet2!A2+B2"
            assert ctx.active.get_cell("A3").formula == "=Sheet2!A3+B3"

    def test_auto_fill_sum_range(self, tmp_path: Path):
        """Test that range references in functions adjust correctly."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=SUM(B1:D1)"
            ctx.active.range("A1:A3").auto_fill()

            assert ctx.active.get_cell("A1").formula == "=SUM(B1:D1)"
            assert ctx.active.get_cell("A2").formula == "=SUM(B2:D2)"
            assert ctx.active.get_cell("A3").formula == "=SUM(B3:D3)"

    def test_auto_fill_multi_row_source(self, tmp_path: Path):
        """Test multi-row source pattern repeats correctly."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=X1"
            ctx.active["A2"] = "=Y1"
            ctx.active.range("A1:A8").auto_fill(source_rows=2)

            assert ctx.active.get_cell("A1").formula == "=X1"
            assert ctx.active.get_cell("A2").formula == "=Y1"
            assert ctx.active.get_cell("A3").formula == "=X3"
            assert ctx.active.get_cell("A4").formula == "=Y3"
            assert ctx.active.get_cell("A5").formula == "=X5"
            assert ctx.active.get_cell("A6").formula == "=Y5"
            assert ctx.active.get_cell("A7").formula == "=X7"
            assert ctx.active.get_cell("A8").formula == "=Y7"

    def test_auto_fill_multi_col_source(self, tmp_path: Path):
        """Test multi-column source pattern repeats correctly."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=A2"
            ctx.active["B1"] = "=B3"
            ctx.active.range("A1:F1").auto_fill(direction="right", source_cols=2)

            assert ctx.active.get_cell("A1").formula == "=A2"
            assert ctx.active.get_cell("B1").formula == "=B3"
            assert ctx.active.get_cell("C1").formula == "=C2"
            assert ctx.active.get_cell("D1").formula == "=D3"
            assert ctx.active.get_cell("E1").formula == "=E2"
            assert ctx.active.get_cell("F1").formula == "=F3"

    def test_auto_fill_copies_styles(self, tmp_path: Path):
        """Test that styles are copied by default."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            ctx.active.get_cell("A1").font = Font(bold=True, color="0000FF")
            ctx.active.get_cell("A1").number_format = "#,##0.00"
            ctx.active.range("A1:A3").auto_fill()

            for row in range(1, 4):
                cell = ctx.active._formula_ws.cell(row, 1)
                assert cell.font.bold is True
                assert cell.font.color.rgb == "000000FF"
                assert cell.number_format == "#,##0.00"

    def test_auto_fill_skip_styles(self, tmp_path: Path):
        """Test that style copying can be disabled."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            ctx.active.get_cell("A1").font = Font(bold=True)
            ctx.active.range("A1:A3").auto_fill(copy_styles=False)

            assert ctx.active._formula_ws.cell(1, 1).font.bold is True
            assert ctx.active._formula_ws.cell(2, 1).font.bold is False
            assert ctx.active._formula_ws.cell(3, 1).font.bold is False

    def test_auto_fill_single_cell_raises(self, tmp_path: Path):
        """Test that single cell range raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="Cannot auto_fill a single cell"):
                ctx.active.range("A1").auto_fill()

    def test_auto_fill_source_rows_too_large_raises(self, tmp_path: Path):
        """Test that source_rows >= range rows raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="source_rows.*must be less than"):
                ctx.active.range("A1:A3").auto_fill(source_rows=3)

    def test_auto_fill_source_cols_too_large_raises(self, tmp_path: Path):
        """Test that source_cols >= range cols raises ValueError."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            with pytest.raises(ValueError, match="source_cols.*must be less than"):
                ctx.active.range("A1:C1").auto_fill(direction="right", source_cols=3)

    def test_auto_fill_auto_direction_tall(self, tmp_path: Path):
        """Test that auto direction chooses 'down' for tall ranges."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=B1"
            ctx.active.range("A1:A5").auto_fill()  # 5 rows, 1 col -> down

            assert ctx.active.get_cell("A2").formula == "=B2"

    def test_auto_fill_auto_direction_wide(self, tmp_path: Path):
        """Test that auto direction chooses 'right' for wide ranges."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=A2"
            ctx.active.range("A1:E1").auto_fill()  # 1 row, 5 cols -> right

            assert ctx.active.get_cell("B1").formula == "=B2"

    def test_auto_fill_marks_dirty(self, tmp_path: Path):
        """Test that auto_fill() marks context as dirty."""
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            ctx.sync()
            assert not ctx._dirty

            ctx.active.range("A1:A3").auto_fill()
            assert ctx._dirty

    def test_auto_fill_after_sync_calculates_values(self, tmp_path: Path):
        """Test that formulas added after sync() are calculated correctly.

        Regression test: formulas added after an initial sync() were not
        being calculated because cached WorksheetProxy objects still
        pointed to the old formula workbook after reload.
        """
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            # Set up data and sync first
            ctx.active["B1"] = 10
            ctx.active["C1"] = 2
            ctx.active["C2"] = 3
            ctx.active["C3"] = 4
            ctx.sync()

            # Now add formula with auto_fill AFTER the sync
            ctx.active["D1"] = "=$B$1*C1"
            ctx.active.range("D1:D3").auto_fill()
            ctx.sync()

            # Values should be calculated: 10*2=20, 10*3=30, 10*4=40
            assert ctx.active["D1"].value == 20  # type: ignore[union-attr]
            assert ctx.active["D2"].value == 30  # type: ignore[union-attr]
            assert ctx.active["D3"].value == 40  # type: ignore[union-attr]

    def test_formula_after_sync_calculates(self, tmp_path: Path):
        """Test that any formula added after sync() is calculated.

        Regression test: this is the simpler case without auto_fill.
        """
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.active["B1"] = 5
            ctx.sync()

            # Add formula after sync
            ctx.active["C1"] = "=A1+B1"
            ctx.sync()

            assert ctx.active["C1"].value == 15  # type: ignore[union-attr]


class TestSheetFormulas:
    """Tests for WorksheetProxy.formulas property."""

    def test_sheet_formulas(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.active["A2"] = 20
            ctx.active["A3"] = "=A1+A2"
            ctx.active["B1"] = "=A1*2"

            formulas = ctx.active.formulas
            assert formulas == {"A3": "=A1+A2", "B1": "=A1*2"}

    def test_sheet_formulas_empty(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            formulas = ctx.active.formulas
            assert formulas == {}

    def test_sheet_formulas_no_formulas(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 100
            ctx.active["A2"] = "text"
            formulas = ctx.active.formulas
            assert formulas == {}

    def test_sheet_formulas_complex(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=SUM(B1:B10)"
            ctx.active["A2"] = "=IF(A1>0,A1,-A1)"
            ctx.active["A3"] = "=VLOOKUP(A1,B:C,2,FALSE)"

            formulas = ctx.active.formulas
            assert "A1" in formulas
            assert "A2" in formulas
            assert "A3" in formulas


class TestExtractCellRefs:
    def test_simple_refs(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs("=A1+B2", "Sheet1")
            assert set(refs) == {"Sheet1!A1", "Sheet1!B2"}

    def test_absolute_refs(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs("=$A$1+$B2+C$3", "Sheet1")
            assert "Sheet1!A1" in refs
            assert "Sheet1!B2" in refs
            assert "Sheet1!C3" in refs

    def test_cross_sheet_refs(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs("=Sheet2!A1+B1", "Sheet1")
            assert "Sheet2!A1" in refs
            assert "Sheet1!B1" in refs

    def test_range_refs(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs("=SUM(A1:A10)", "Sheet1")
            assert "Sheet1!A1" in refs
            assert "Sheet1!A10" in refs

    def test_no_refs(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs("=PI()", "Sheet1")
            assert refs == []

    def test_mixed_content(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            refs = ctx._extract_cell_refs('=IF(A1>0, "YES", B2)', "Sheet1")
            assert "Sheet1!A1" in refs
            assert "Sheet1!B2" in refs


class TestErrorDetail:
    def test_error_detail_creation(self):
        detail = ErrorDetail(
            location="Sheet1!A1",
            error="#DIV/0!",
            formula="=B1/C1",
            neighbors={"Sheet1!B1": 100, "Sheet1!C1": 0},
        )
        assert detail.location == "Sheet1!A1"
        assert detail.error == "#DIV/0!"
        assert detail.formula == "=B1/C1"
        assert detail.neighbors["Sheet1!B1"] == 100
        assert detail.neighbors["Sheet1!C1"] == 0

    def test_error_detail_defaults(self):
        detail = ErrorDetail(location="Sheet1!A1", error="#REF!")
        assert detail.formula is None
        assert detail.neighbors == {}


class TestSyncResultWithDetails:
    def test_sync_result_with_error_details(self):
        details = [
            ErrorDetail(
                location="Sheet1!A1",
                error="#DIV/0!",
                formula="=B1/C1",
                neighbors={"Sheet1!B1": 10, "Sheet1!C1": 0},
            )
        ]
        result = SyncResult(
            success=False,
            total_errors=1,
            errors={"#DIV/0!": ["Sheet1!A1"]},
            error_details=details,
        )
        assert len(result.error_details) == 1
        assert result.error_details[0].formula == "=B1/C1"
        assert result.error_details[0].neighbors["Sheet1!C1"] == 0

    def test_sync_result_empty_details(self):
        result = SyncResult(success=True, total_errors=0, errors={})
        assert result.error_details == []


class TestProxyValueUpdate:
    """Tests that worksheet proxies update correctly after sync."""

    def test_proxy_shows_calculated_value_after_sync(self, tmp_path: Path):
        """Verify that worksheet proxy returns calculated values after sync."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            # Use the proxy returned by create_sheet directly (don't reassign to ctx.active)
            ws = ctx.create_sheet("Sheet1", 0)

            ws["A1"] = 100
            ws["A2"] = 200
            ws["A3"] = "=SUM(A1:A2)"

            # Before sync, should see formula
            assert ws["A3"].value == "=SUM(A1:A2)"  # type: ignore[union-attr]

            ctx.sync()

            # After sync, same proxy should show calculated value
            assert ws["A3"].value == 300  # type: ignore[union-attr]
            # Read via range too
            assert ws.range("A3").values == [[300]]

    def test_proxy_identity_maintained_after_sync(self, tmp_path: Path):
        """Verify that worksheet proxy object identity is maintained after sync."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws_before = ctx.active
            ws_before["A1"] = "=10+20"

            ctx.sync()

            ws_after = ctx.active

            # Should be same object (cached)
            assert ws_before is ws_after
            # And should show calculated value
            assert ws_after["A1"].value == 30  # type: ignore[union-attr]

    def test_create_sheet_proxy_updates_after_sync(self, tmp_path: Path):
        """Verify that proxy from create_sheet gets updated after sync."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            # Keep reference to proxy returned by create_sheet
            ws = ctx.create_sheet("Data", 0)

            # Use range API to set values including formula
            ws.range("A1:A3").values = [[100], [200], ["=A1+A2"]]

            # Before sync, formula is just a string
            assert ws["A3"].value == "=A1+A2"  # type: ignore[union-attr]

            ctx.sync()

            # After sync, should show calculated value
            assert ws["A3"].value == 300  # type: ignore[union-attr]
            assert ws.range("A3").values == [[300]]

    def test_renamed_sheet_proxy_updates_after_sync(self, tmp_path: Path):
        """Verify that proxy still updates after sheet is renamed.

        Regression test: cache key was stale after rename, causing
        _update_values_wb to fail to find the sheet in values_wb.
        """
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            # Access active sheet (cached under "Sheet")
            ws = ctx.active
            # Rename it (cache key becomes stale)
            ws.title = "Renamed"

            ws["A1"] = 100
            ws["A2"] = "=A1*2"

            # Before sync, should see formula
            assert ws["A2"].value == "=A1*2"  # type: ignore[union-attr]

            ctx.sync()

            # After sync, should show calculated value despite rename
            assert ws["A2"].value == 200  # type: ignore[union-attr]
            assert ws.range("A1:A2").values == [[100], [200]]


class TestBuildErrorDetails:
    def test_build_error_details_requires_sync(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=1/0"
            ctx.workbook.save(path)
            details = ctx._build_error_details({"#DIV/0!": ["Sheet!A1"]})
            assert details == []

    def test_build_error_details_basic(self, tmp_path: Path):
        from openpyxl import Workbook

        path = tmp_path / "test.xlsx"
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["B1"] = 0
        ws["C1"] = "=A1/B1"
        wb.save(path)

        with ExcelContext(path) as ctx:
            ctx._workbook = wb
            ctx._values_workbook = Workbook()
            ctx._values_workbook.active["A1"] = 10
            ctx._values_workbook.active["B1"] = 0
            ctx._values_workbook.active["C1"] = "#DIV/0!"

            details = ctx._build_error_details({"#DIV/0!": ["Sheet!C1"]})

            assert len(details) == 1
            assert details[0].location == "Sheet!C1"
            assert details[0].error == "#DIV/0!"
            assert details[0].formula == "=A1/B1"
            assert "Sheet!A1" in details[0].neighbors
            assert "Sheet!B1" in details[0].neighbors

    def test_auto_remove_default_sheet(self, tmp_path: Path):
        """Verify that the default 'Sheet' is auto-removed when creating first custom sheet."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            # Initially should have the default "Sheet"
            assert len(ctx.workbook.sheetnames) == 1
            assert "Sheet" in ctx.workbook.sheetnames

            # Create first custom sheet - should auto-remove default "Sheet"
            ctx.create_sheet("Data")

            # Should now have only "Data", not "Sheet"
            assert len(ctx.workbook.sheetnames) == 1
            assert "Data" in ctx.workbook.sheetnames
            assert "Sheet" not in ctx.workbook.sheetnames

            # Creating another sheet should not remove anything
            ctx.create_sheet("Summary")
            assert len(ctx.workbook.sheetnames) == 2
            assert "Data" in ctx.workbook.sheetnames
            assert "Summary" in ctx.workbook.sheetnames

    def test_keep_default_sheet_if_has_data(self, tmp_path: Path):
        """Verify that default sheet is kept if it contains data."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            # Add data to default sheet
            ctx.active["A1"] = "Important data"

            # Create a custom sheet - should NOT remove default sheet
            ctx.create_sheet("Data")

            # Should have both sheets
            assert len(ctx.workbook.sheetnames) == 2
            assert "Sheet" in ctx.workbook.sheetnames
            assert "Data" in ctx.workbook.sheetnames
            assert ctx.sheet("Sheet")["A1"].value == "Important data"  # type: ignore[union-attr]

    def test_no_auto_remove_for_loaded_workbooks(self, tmp_path: Path):
        """Verify that auto-removal only happens for newly created workbooks."""
        path = tmp_path / "test.xlsx"

        # Create a workbook with default "Sheet"
        with ExcelContext(path, create=True) as ctx:
            ctx.workbook.save(path)

        # Load the existing workbook and create a sheet
        with ExcelContext(path) as ctx:
            assert "Sheet" in ctx.workbook.sheetnames

            ctx.create_sheet("Data")

            # Should keep both sheets (no auto-removal for loaded workbooks)
            assert len(ctx.workbook.sheetnames) == 2
            assert "Sheet" in ctx.workbook.sheetnames
            assert "Data" in ctx.workbook.sheetnames


class TestSheetManagement:
    def test_set_active_by_proxy(self, tmp_path: Path):
        """Test setting active sheet using worksheet proxy."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.create_sheet("Summary")

            # Set active by proxy
            ctx.active = ctx.sheet("Summary")
            assert ctx.active.title == "Summary"

            ctx.active = ctx.sheet("Data")
            assert ctx.active.title == "Data"

    def test_set_active_by_name(self, tmp_path: Path):
        """Test setting active sheet by name string."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.create_sheet("Summary")

            # Set active by name
            ctx.active = "Summary"
            assert ctx.active.title == "Summary"

            ctx.active = "Data"
            assert ctx.active.title == "Data"

    def test_set_active_via_workbook_proxy(self, tmp_path: Path):
        """Test setting active sheet via workbook proxy directly."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.create_sheet("Summary")

            # Set active via workbook proxy
            ctx.wb.active = ctx.wb["Summary"]
            assert ctx.wb.active.title == "Summary"

            ctx.wb.active = "Data"
            assert ctx.wb.active.title == "Data"

    def test_delete_sheet(self, tmp_path: Path):
        """Test deleting a sheet."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.create_sheet("Summary")

            assert "Data" in ctx.workbook.sheetnames
            assert "Summary" in ctx.workbook.sheetnames

            ctx.delete_sheet("Data")

            assert "Data" not in ctx.workbook.sheetnames
            assert "Summary" in ctx.workbook.sheetnames

    def test_delete_sheet_not_found(self, tmp_path: Path):
        """Test that deleting non-existent sheet raises KeyError."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")

            with pytest.raises(KeyError, match="NotFound"):
                ctx.delete_sheet("NotFound")

    def test_delete_only_sheet_raises(self, tmp_path: Path):
        """Test that deleting the only sheet raises ValueError."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            # Only one sheet exists
            with pytest.raises(ValueError, match="Cannot delete the only sheet"):
                ctx.delete_sheet("Sheet")

    def test_delete_sheet_clears_cache(self, tmp_path: Path):
        """Test that deleting a sheet clears it from the proxy cache."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ctx.create_sheet("Data")
            ctx.create_sheet("Summary")

            # Access sheet to cache it
            _ = ctx.sheet("Data")
            assert "Data" in ctx._proxy._sheet_cache  # type: ignore[union-attr]

            ctx.delete_sheet("Data")

            # Should be removed from cache
            assert "Data" not in ctx._proxy._sheet_cache  # type: ignore[union-attr]


class TestCellFormula:
    """Tests for CellProxy.formula property."""

    def test_formula_returns_formula_string(self, tmp_path: Path):
        """Test that .formula returns formula string for formula cells."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "=SUM(B1:B10)"
            ws["A2"] = "=A1*2"

            assert ws["A1"].formula == "=SUM(B1:B10)"  # type: ignore[union-attr]
            assert ws["A2"].formula == "=A1*2"  # type: ignore[union-attr]

    def test_formula_returns_none_for_values(self, tmp_path: Path):
        """Test that .formula returns None for non-formula cells."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 42
            ws["A2"] = "hello"
            ws["A3"] = 3.14

            assert ws["A1"].formula is None  # type: ignore[union-attr]
            assert ws["A2"].formula is None  # type: ignore[union-attr]
            assert ws["A3"].formula is None  # type: ignore[union-attr]

    def test_formula_returns_none_for_empty(self, tmp_path: Path):
        """Test that .formula returns None for empty cells."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            assert ws["A1"].formula is None  # type: ignore[union-attr]


class TestRangeDump:
    """Tests for RangeProxy.dump() method."""

    def test_dump_returns_formatted_string(self, tmp_path: Path):
        """Test that dump() returns a formatted table string."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws.range("A1:C3").values = [
                [1, 2, 3],
                [4, 5, 6],
                [7, 8, 9],
            ]

            result = ws.range("A1:C3").dump()

            assert isinstance(result, str)
            assert "A" in result
            assert "B" in result
            assert "C" in result
            assert "1" in result
            assert "9" in result

    def test_dump_shows_formulas(self, tmp_path: Path):
        """Test that dump(show_formulas=True) shows formula strings."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["B1"] = "=A1*2"

            result = ws.range("A1:B1").dump(show_formulas=True)

            assert "=A1*2" in result

    def test_dump_handles_none_values(self, tmp_path: Path):
        """Test that dump() handles empty cells gracefully."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 1
            # B1 is empty

            result = ws.range("A1:B1").dump()

            assert isinstance(result, str)
            assert "1" in result

    def test_dump_truncates_long_values(self, tmp_path: Path):
        """Test that dump() truncates very long cell values."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "This is a very long string that should be truncated"

            result = ws.range("A1:A1").dump()

            assert "…" in result  # Truncation indicator


class TestFormulaCount:
    """Tests for formula_count property on RangeProxy and WorksheetProxy."""

    def test_range_formula_count(self, tmp_path: Path):
        """Test formula_count on RangeProxy."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "=A1*2"
            ws["A3"] = "=A1+A2"
            ws["B1"] = "text"
            ws["B2"] = "=A1+10"

            # Range A1:A3 has 2 formulas
            assert ws.range("A1:A3").formula_count == 2

            # Range A1:B2 has 2 formulas
            assert ws.range("A1:B2").formula_count == 2

            # Range A1:B3 has 3 formulas
            assert ws.range("A1:B3").formula_count == 3

    def test_range_formula_count_empty(self, tmp_path: Path):
        """Test formula_count returns 0 for range with no formulas."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "text"

            assert ws.range("A1:A2").formula_count == 0

    def test_sheet_formula_count(self, tmp_path: Path):
        """Test formula_count on WorksheetProxy."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "=A1*2"
            ws["B5"] = "=SUM(A1:A2)"
            ws["C10"] = "=A1+B5"

            assert ws.formula_count == 3

    def test_sheet_formula_count_empty(self, tmp_path: Path):
        """Test formula_count returns 0 for sheet with no formulas."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "text"

            assert ws.formula_count == 0


class TestFindErrors:
    """Tests for find_errors() on RangeProxy, WorksheetProxy, and ExcelContext."""

    def test_range_find_errors_after_sync(self, tmp_path: Path):
        """Test find_errors on RangeProxy after sync."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = 0
            ws["A3"] = "=A1/A2"  # Will be #DIV/0!
            ws["B1"] = "=NonExistent()"  # Will be #NAME?

            ctx.sync()

            # Check range with #DIV/0!
            errors = ws.range("A1:A3").find_errors()
            assert "#DIV/0!" in errors
            assert "A3" in errors["#DIV/0!"]

            # Check range with #NAME?
            errors = ws.range("B1:B1").find_errors()
            assert "#NAME?" in errors
            assert "B1" in errors["#NAME?"]

    def test_range_find_errors_no_errors(self, tmp_path: Path):
        """Test find_errors returns empty dict when no errors."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = 2
            ws["A3"] = "=A1/A2"  # Valid: 5

            ctx.sync()

            errors = ws.range("A1:A3").find_errors()
            assert errors == {}

    def test_range_find_errors_before_sync(self, tmp_path: Path):
        """Test find_errors returns empty dict before sync (no values_ws)."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "=1/0"

            # Before sync, values_ws is None
            errors = ws.range("A1:A1").find_errors()
            assert errors == {}

    def test_sheet_find_errors_after_sync(self, tmp_path: Path):
        """Test find_errors on WorksheetProxy after sync."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = 0
            ws["A3"] = "=A1/A2"  # #DIV/0!
            ws["C10"] = "=MissingSheet!A1"  # #REF! or #NAME?

            ctx.sync()

            errors = ws.find_errors()
            assert "#DIV/0!" in errors
            assert "A3" in errors["#DIV/0!"]
            # The other error could be #REF! or #NAME? depending on LibreOffice
            assert len(errors) >= 1

    def test_sheet_find_errors_no_errors(self, tmp_path: Path):
        """Test find_errors returns empty dict when no errors on sheet."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "=A1*2"

            ctx.sync()

            errors = ws.find_errors()
            assert errors == {}

    def test_context_find_errors_single_sheet(self, tmp_path: Path):
        """Test find_errors on ExcelContext with single sheet."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 0
            ws["A2"] = "=1/A1"  # #DIV/0!

            ctx.sync()

            errors = ctx.find_errors()
            assert "#DIV/0!" in errors
            # Should be fully qualified with sheet name
            assert any("A2" in loc for loc in errors["#DIV/0!"])

    def test_context_find_errors_multiple_sheets(self, tmp_path: Path):
        """Test find_errors on ExcelContext aggregates from all sheets."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws1 = ctx.create_sheet("Sheet1")
            ws1["A1"] = 0
            ws1["A2"] = "=1/A1"  # #DIV/0!

            ws2 = ctx.create_sheet("Sheet2")
            ws2["B1"] = "=UnknownFunc()"  # #NAME?

            ctx.sync()

            errors = ctx.find_errors()

            # Should have errors from both sheets
            assert "#DIV/0!" in errors
            assert "#NAME?" in errors

            # Locations should be fully qualified
            assert any("Sheet1!" in loc for loc in errors["#DIV/0!"])
            assert any("Sheet2!" in loc for loc in errors["#NAME?"])

    def test_context_find_errors_no_errors(self, tmp_path: Path):
        """Test find_errors returns empty dict when no errors in workbook."""
        path = tmp_path / "test.xlsx"

        with create(path, verbose_errors=False) as ctx:
            ws = ctx.active
            ws["A1"] = 10
            ws["A2"] = "=A1*2"

            ctx.sync()

            errors = ctx.find_errors()
            assert errors == {}

    def test_context_find_errors_before_sync(self, tmp_path: Path):
        """Test find_errors returns empty dict before sync."""
        path = tmp_path / "test.xlsx"

        with ExcelContext(path, create=True) as ctx:
            ws = ctx.active
            ws["A1"] = "=1/0"

            # Before sync
            errors = ctx.find_errors()
            assert errors == {}


class TestSourceCode:
    """Tests for _code - code passed through for hooks."""

    def test_code_none_for_library_usage(self, tmp_path: Path):
        """Test that _code is None when not using CLI."""
        path = tmp_path / "test.xlsx"

        with create(path) as ctx:
            assert ctx._code is None

    def test_code_with_explicit_param(self, tmp_path: Path):
        """Test _code parameter (used by CLI)."""
        path = tmp_path / "test.xlsx"
        code = "ctx.active['A1'] = 1\nctx.active['A2'] = 2\n"

        with create(path, _code=code) as ctx:
            assert ctx._code == code
