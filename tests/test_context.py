"""Tests for ExcelContext and run()."""

from pathlib import Path

import pytest
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side

from headless_excel import ErrorDetail, ExcelContext, FormulaError, SyncResult, run


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


class TestApplyStyle:
    def test_apply_font_to_single_cell(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.apply_style("Sheet", "A1", font=Font(bold=True, color="FF0000"))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell = ctx.active["A1"]
            assert cell.font.bold is True
            assert cell.font.color.rgb == "00FF0000"

    def test_apply_font_to_range(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            for col in ["A", "B", "C"]:
                ctx.active[f"{col}1"] = f"Value {col}"
            ctx.apply_style("Sheet", "A1:C1", font=Font(italic=True))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            for col in ["A", "B", "C"]:
                assert ctx.active[f"{col}1"].font.italic is True

    def test_apply_fill(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Colored"
            fill = PatternFill(
                start_color="FFFF00", end_color="FFFF00", fill_type="solid"
            )
            ctx.apply_style("Sheet", "A1", fill=fill)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].fill.start_color.rgb == "00FFFF00"

    def test_apply_number_format(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 1234.56
            ctx.active["A2"] = -789.12
            ctx.apply_style("Sheet", "A1:A2", number_format="#,##0.00;(#,##0.00)")
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].number_format == "#,##0.00;(#,##0.00)"
            assert ctx.active["A2"].number_format == "#,##0.00;(#,##0.00)"

    def test_apply_alignment(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Centered"
            ctx.apply_style(
                "Sheet", "A1", alignment=Alignment(horizontal="center", vertical="top")
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].alignment.horizontal == "center"
            assert ctx.active["A1"].alignment.vertical == "top"

    def test_apply_border(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Bordered"
            thin_side = Side(style="thin", color="000000")
            border = Border(
                left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
            )
            ctx.apply_style("Sheet", "A1", border=border)
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].border.left.style == "thin"
            assert ctx.active["A1"].border.right.style == "thin"

    def test_apply_protection(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Protected"
            ctx.apply_style(
                "Sheet", "A1", protection=Protection(locked=True, hidden=True)
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            assert ctx.active["A1"].protection.locked is True
            assert ctx.active["A1"].protection.hidden is True

    def test_apply_multiple_styles(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["B2"] = 42
            ctx.apply_style(
                "Sheet",
                "B2",
                font=Font(bold=True),
                fill=PatternFill(start_color="00FF00", fill_type="solid"),
                number_format="0.00",
                alignment=Alignment(horizontal="right"),
            )
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            cell = ctx.active["B2"]
            assert cell.font.bold is True
            assert cell.fill.start_color.rgb == "0000FF00"
            assert cell.number_format == "0.00"
            assert cell.alignment.horizontal == "right"

    def test_apply_style_to_2d_range(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            for row in range(1, 4):
                for col in ["A", "B", "C"]:
                    ctx.active[f"{col}{row}"] = f"{col}{row}"
            ctx.apply_style("Sheet", "A1:C3", font=Font(size=14))
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            for row in range(1, 4):
                for col in ["A", "B", "C"]:
                    assert ctx.active[f"{col}{row}"].font.size == 14

    def test_apply_style_sets_dirty_flag(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Test"
            ctx.workbook.save(path)
            ctx._dirty = False
            ctx.apply_style("Sheet", "A1", font=Font(bold=True))
            assert ctx._dirty is True


class TestGetFormulas:
    def test_get_formulas_single_sheet(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = 10
            ctx.active["A2"] = 20
            ctx.active["A3"] = "=A1+A2"
            ctx.active["B1"] = "=A1*2"
            ctx.active["C1"] = "Static text"
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            formulas = ctx.get_formulas("Sheet")
            assert formulas == {"A3": "=A1+A2", "B1": "=A1*2"}

    def test_get_formulas_all_sheets(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=SUM(1,2)"
            ctx.create_sheet("Data")
            ctx.sheet("Data")["B2"] = "=AVERAGE(1,2,3)"
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            formulas = ctx.get_formulas()
            assert "Sheet!A1" in formulas
            assert formulas["Sheet!A1"] == "=SUM(1,2)"
            assert "Data!B2" in formulas
            assert formulas["Data!B2"] == "=AVERAGE(1,2,3)"

    def test_get_formulas_empty_sheet(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            formulas = ctx.get_formulas("Sheet")
            assert formulas == {}

    def test_get_formulas_no_formulas(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "Hello"
            ctx.active["A2"] = 123
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            formulas = ctx.get_formulas("Sheet")
            assert formulas == {}

    def test_get_formulas_complex_formulas(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with ExcelContext(path, create=True) as ctx:
            ctx.active["A1"] = "=IF(B1>0, B1*C1, 0)"
            ctx.active["A2"] = "=VLOOKUP(D1, A1:C10, 2, FALSE)"
            ctx.active["A3"] = '=CONCATENATE("Hello", " ", "World")'
            ctx.workbook.save(path)

        with ExcelContext(path) as ctx:
            formulas = ctx.get_formulas("Sheet")
            assert len(formulas) == 3
            assert "IF" in formulas["A1"]
            assert "VLOOKUP" in formulas["A2"]
            assert "CONCATENATE" in formulas["A3"]


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
        with run(path, create=True, auto_sync=False) as ctx:
            ws = ctx.create_sheet("Sheet1", 0)
            ws = ctx.active

            ws["A1"] = 100
            ws["A2"] = 200
            ws["A3"] = "=SUM(A1:A2)"

            # Before sync, should see formula
            assert ws["A3"].value == "=SUM(A1:A2)"

            ctx.sync()

            # After sync, same proxy should show calculated value
            assert ws["A3"].value == 300
            assert ctx.read_value("Sheet1", "A3") == 300

    def test_proxy_identity_maintained_after_sync(self, tmp_path: Path):
        """Verify that worksheet proxy object identity is maintained after sync."""
        path = tmp_path / "test.xlsx"
        with run(path, create=True, auto_sync=False) as ctx:
            ws_before = ctx.active
            ws_before["A1"] = "=10+20"

            ctx.sync()

            ws_after = ctx.active

            # Should be same object (cached)
            assert ws_before is ws_after
            # And should show calculated value
            assert ws_after["A1"].value == 30


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
