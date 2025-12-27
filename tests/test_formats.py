"""Tests for NumberFormats and Colors constants."""

from pathlib import Path

from openpyxl.styles import Font

from headless_excel import (
    ColorLintError,
    ColorLintViolation,
    Colors,
    NumberFormats,
    create,
    infer_financial_color,
    run,
)


class TestNumberFormats:
    """Test NumberFormats constants."""

    def test_accounting_formats_exist(self):
        """Accounting formats should exist and have correct values."""
        assert hasattr(NumberFormats, "ACCOUNTING")
        assert hasattr(NumberFormats, "ACCOUNTING_0DP")
        assert isinstance(NumberFormats.ACCOUNTING, str)
        assert isinstance(NumberFormats.ACCOUNTING_0DP, str)

    def test_percentage_formats_exist(self):
        """Percentage formats should exist and have correct values."""
        assert NumberFormats.PERCENTAGE == "0%"
        assert NumberFormats.PERCENTAGE_1DP == "0.0%"
        assert NumberFormats.PERCENTAGE_2DP == "0.00%"

    def test_number_formats_exist(self):
        """Basic number formats should exist."""
        assert NumberFormats.NUMBER == "#,##0.00"
        assert NumberFormats.NUMBER_0DP == "#,##0"

    def test_date_formats_exist(self):
        """Date formats should exist."""
        assert NumberFormats.DATE == "m/d/yyyy"
        assert NumberFormats.DATE_LONG == "mmmm d, yyyy"

    def test_apply_accounting_format(self, tmp_path: Path):
        """NumberFormats.ACCOUNTING should work with cells."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 1234.56
            ctx.active["A1"].number_format = NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            assert ctx.active["A1"].value == 1234.56  # type: ignore[union-attr]

    def test_apply_percentage_format(self, tmp_path: Path):
        """NumberFormats.PERCENTAGE should work with cells."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 0.1575
            ctx.active["A1"].number_format = NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]

    def test_apply_style_with_format(self, tmp_path: Path):
        """NumberFormats should work with range.apply_style()."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.active["A2"] = 200
            ctx.active.range("A1:A2").apply_style(number_format=NumberFormats.NUMBER)
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.NUMBER  # type: ignore[union-attr]
            assert ctx.active["A2"].number_format == NumberFormats.NUMBER  # type: ignore[union-attr]


class TestColors:
    """Test Colors constants."""

    def test_color_constants_exist(self):
        """Color constants should exist and have correct ARGB values."""
        assert Colors.HARDCODE == "FF0000FF"  # Blue
        assert Colors.FORMULA == "FF000000"  # Black
        assert Colors.EXTERNAL_LINK == "FF008000"  # Green

    def test_colors_are_strings(self):
        """Colors should be strings in ARGB format."""
        assert isinstance(Colors.HARDCODE, str)
        assert isinstance(Colors.FORMULA, str)
        assert isinstance(Colors.EXTERNAL_LINK, str)
        assert len(Colors.HARDCODE) == 8  # ARGB hex: AARRGGBB

    def test_apply_color_to_font(self, tmp_path: Path):
        """Colors should work with openpyxl Font."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.active["A1"].font = Font(color=Colors.HARDCODE)  # type: ignore[union-attr]

            ctx.active["A2"] = "=A1*2"
            ctx.active["A2"].font = Font(color=Colors.FORMULA)  # type: ignore[union-attr]

            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]
            assert ctx.active["A2"].font.color.rgb == Colors.FORMULA  # type: ignore[union-attr]


class TestInferFinancialColor:
    """Test infer_financial_color function."""

    def test_hardcode_number(self):
        """Numbers should be HARDCODE (blue)."""
        assert infer_financial_color(100) == Colors.HARDCODE
        assert infer_financial_color(3.14) == Colors.HARDCODE
        assert infer_financial_color(0) == Colors.HARDCODE

    def test_hardcode_string(self):
        """Non-formula strings should be HARDCODE (blue)."""
        assert infer_financial_color("Hello") == Colors.HARDCODE
        assert infer_financial_color("Revenue") == Colors.HARDCODE

    def test_hardcode_none(self):
        """None should be HARDCODE (blue)."""
        assert infer_financial_color(None) == Colors.HARDCODE

    def test_formula_simple(self):
        """Simple formulas without sheet refs should be FORMULA (black)."""
        assert infer_financial_color("=A1+B1") == Colors.FORMULA
        assert infer_financial_color("=SUM(A1:A10)") == Colors.FORMULA
        assert infer_financial_color("=100*2") == Colors.FORMULA

    def test_external_link_sheet_ref(self):
        """Formulas with sheet references should be EXTERNAL_LINK (green)."""
        assert infer_financial_color("=Sheet2!A1") == Colors.EXTERNAL_LINK
        assert infer_financial_color("=Revenue!B5+Costs!B5") == Colors.EXTERNAL_LINK
        assert infer_financial_color("='Other Sheet'!A1") == Colors.EXTERNAL_LINK

    def test_external_link_workbook_ref(self):
        """Formulas with workbook references should be EXTERNAL_LINK (green)."""
        assert infer_financial_color("=[Book.xlsx]Sheet1!A1") == Colors.EXTERNAL_LINK

    def test_string_with_exclamation_not_external(self):
        """String literals containing '!' should not be detected as external."""
        assert infer_financial_color('="Hello!"') == Colors.FORMULA
        assert infer_financial_color('=A1&"Warning!"') == Colors.FORMULA


class TestAutoFinancialColors:
    """Test auto_financial_colors method."""

    def test_range_auto_financial_colors(self, tmp_path: Path):
        """Range.auto_financial_colors should apply correct colors."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100  # Hardcode
            ws["A2"] = "=A1*2"  # Formula
            ws["B1"] = "Label"  # Hardcode
            ws.range("A1:B2").auto_financial_colors()
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]
            assert ctx.active["A2"].font.color.rgb == Colors.FORMULA  # type: ignore[union-attr]
            assert ctx.active["B1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]

    def test_sheet_auto_financial_colors(self, tmp_path: Path):
        """Sheet.auto_financial_colors should apply to entire sheet."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A2"] = "=A1*2"
            ws["C5"] = "=Sheet2!A1"  # Would be external if Sheet2 existed
            ws.auto_financial_colors()
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]
            assert ctx.active["A2"].font.color.rgb == Colors.FORMULA  # type: ignore[union-attr]
            assert ctx.active["C5"].font.color.rgb == Colors.EXTERNAL_LINK  # type: ignore[union-attr]

    def test_context_auto_financial_colors(self, tmp_path: Path):
        """ctx.auto_financial_colors should apply to all sheets."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            sheet1 = ctx.create_sheet("Data")
            sheet1["A1"] = 100
            sheet2 = ctx.create_sheet("Calcs")
            sheet2["A1"] = "=Data!A1*2"
            ctx.auto_financial_colors()
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.sheet("Data")["A1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]
            assert ctx.sheet("Calcs")["A1"].font.color.rgb == Colors.EXTERNAL_LINK  # type: ignore[union-attr]

    def test_auto_preserves_other_font_properties(self, tmp_path: Path):
        """auto_financial_colors should preserve bold, italic, etc."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A1"].font = Font(bold=True, italic=True, size=14)  # type: ignore[union-attr]
            ws.range("A1").auto_financial_colors()
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            font = ctx.active["A1"].font  # type: ignore[union-attr]
            assert font.color.rgb == Colors.HARDCODE
            assert font.bold is True
            assert font.italic is True
            assert font.size == 14


class TestLintFinancialColors:
    """Test lint_financial_colors method."""

    def test_range_lint_no_violations(self, tmp_path: Path):
        """Range with correct colors should have no violations."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A1"].font = Font(color=Colors.HARDCODE)  # type: ignore[union-attr]
            ws["A2"] = "=A1*2"
            ws["A2"].font = Font(color=Colors.FORMULA)  # type: ignore[union-attr]

            violations = ws.range("A1:A2").lint_financial_colors()
            assert violations == []

    def test_range_lint_with_violations(self, tmp_path: Path):
        """Range with wrong colors should report violations."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ws = ctx.active
            ws["A1"] = 100  # Should be blue, but no color set
            ws["A2"] = "=A1*2"
            ws["A2"].font = Font(
                color=Colors.HARDCODE
            )  # Wrong: should be black  # type: ignore[union-attr]

            violations = ws.range("A1:A2").lint_financial_colors()
            assert len(violations) == 2

            # Check A1 violation
            v1 = next(v for v in violations if v.cell == "A1")
            assert v1.expected_color == Colors.HARDCODE
            assert v1.value == 100

            # Check A2 violation
            v2 = next(v for v in violations if v.cell == "A2")
            assert v2.expected_color == Colors.FORMULA
            assert v2.current_color == Colors.HARDCODE

    def test_context_lint_all_sheets(self, tmp_path: Path):
        """ctx.lint_financial_colors should check all sheets."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            sheet1 = ctx.create_sheet("Data")
            sheet1["A1"] = 100  # Violation: no color

            sheet2 = ctx.create_sheet("Calcs")
            sheet2["A1"] = "=Data!A1"  # Violation: no color

            violations = ctx.lint_financial_colors()
            assert "Data" in violations
            assert "Calcs" in violations
            assert len(violations["Data"]) == 1
            assert len(violations["Calcs"]) == 1

    def test_context_lint_raise_on_violations(self, tmp_path: Path):
        """ctx.lint_financial_colors should raise when requested."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100  # Violation

            try:
                ctx.lint_financial_colors(raise_on_violations=True)
                assert False, "Should have raised ColorLintError"
            except ColorLintError as e:
                assert e.total == 1
                assert "Sheet" in e.violations

    def test_lint_financial_colors_context_param(self, tmp_path: Path):
        """lint_financial_colors=True on context should raise on exit."""
        path = tmp_path / "test.xlsx"
        try:
            with create(path, auto_sync=False, lint_financial_colors=True) as ctx:
                ctx.active["A1"] = 100  # Violation: no color
            assert False, "Should have raised SystemExit"
        except SystemExit as e:
            assert "Financial color violations" in str(e)

    def test_lint_passes_with_correct_colors(self, tmp_path: Path):
        """lint_financial_colors=True should pass when colors are correct."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False, lint_financial_colors=True) as ctx:
            ws = ctx.active
            ws["A1"] = 100
            ws["A1"].font = Font(color=Colors.HARDCODE)  # type: ignore[union-attr]
            ws["A2"] = "=A1*2"
            ws["A2"].font = Font(color=Colors.FORMULA)  # type: ignore[union-attr]
            # Should not raise


class TestColorLintError:
    """Test ColorLintError exception."""

    def test_error_str(self):
        """ColorLintError should have useful string representation."""
        violations = {
            "Sheet1": [
                ColorLintViolation(
                    cell="A1",
                    value=100,
                    expected_color=Colors.HARDCODE,
                    current_color=None,
                )
            ]
        }
        err = ColorLintError(violations=violations, total=1)
        s = str(err)
        assert "Financial color violations (1)" in s
        assert "Sheet1!A1" in s
        assert "HARDCODE" in s

    def test_error_empty(self):
        """Empty ColorLintError should report no violations."""
        err = ColorLintError()
        assert str(err) == "No color lint violations"
