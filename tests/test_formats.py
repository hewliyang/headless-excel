"""Tests for NumberFormats and Colors constants."""

from pathlib import Path

from openpyxl.styles import Font

from headless_excel import (
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
        with create(path, auto_sync=False, lint_financial_colors=False) as ctx:
            ctx.active["A1"] = 1234.56
            ctx.active["A1"].number_format = NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False, lint_financial_colors=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            assert ctx.active["A1"].value == 1234.56  # type: ignore[union-attr]

    def test_apply_percentage_format(self, tmp_path: Path):
        """NumberFormats.PERCENTAGE should work with cells."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False, lint_financial_colors=False) as ctx:
            ctx.active["A1"] = 0.1575
            ctx.active["A1"].number_format = NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False, lint_financial_colors=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]

    def test_apply_style_with_format(self, tmp_path: Path):
        """NumberFormats should work with range.apply_style()."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False, lint_financial_colors=False) as ctx:
            ctx.active["A1"] = 100
            ctx.active["A2"] = 200
            ctx.active.range("A1:A2").apply_style(number_format=NumberFormats.NUMBER)
            ctx.workbook.save(path)

        with run(path, auto_sync=False, lint_financial_colors=False) as ctx:
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

    def test_text_labels_not_colored(self):
        """Non-formula strings (text labels) should return None (no color change)."""
        assert infer_financial_color("Hello") is None
        assert infer_financial_color("Revenue") is None

    def test_none_not_colored(self):
        """None should return None (no color change)."""
        assert infer_financial_color(None) is None

    def test_bool_not_colored(self):
        """Booleans should return None (no color change)."""
        assert infer_financial_color(True) is None
        assert infer_financial_color(False) is None

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
            ws["A1"] = 100  # Hardcode (numeric)
            ws["A2"] = "=A1*2"  # Formula
            ws["B1"] = "Label"  # Text label (not colored)
            ws.range("A1:B2").auto_financial_colors()
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].font.color.rgb == Colors.HARDCODE  # type: ignore[union-attr]
            assert ctx.active["A2"].font.color.rgb == Colors.FORMULA  # type: ignore[union-attr]
            # Text labels are not colored - they keep their default theme color
            assert ctx.active["B1"].font.color.type == "theme"  # type: ignore[union-attr]

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
        with create(path, auto_sync=False, lint_financial_colors=False) as ctx:
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
        with create(path, auto_sync=False, lint_financial_colors=False) as ctx:
            sheet1 = ctx.create_sheet("Data")
            sheet1["A1"] = 100  # Violation: no color

            sheet2 = ctx.create_sheet("Calcs")
            sheet2["A1"] = "=Data!A1"  # Violation: no color

            violations = ctx.lint_financial_colors()
            assert "Data" in violations
            assert "Calcs" in violations
            assert len(violations["Data"]) == 1
            assert len(violations["Calcs"]) == 1

    def test_lint_financial_colors_context_param(self, tmp_path: Path, caplog):
        """lint_financial_colors=True on context should log warning on exit."""
        import logging

        path = tmp_path / "test.xlsx"
        with caplog.at_level(logging.WARNING):
            with create(path, auto_sync=False, lint_financial_colors=True, auto_financial_colors=False) as ctx:
                ctx.active["A1"] = 100  # Violation: no color

        assert "Financial color violations" in caplog.text

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


class TestFormatColorViolations:
    """Test format_color_violations function."""

    def test_format_violations(self):
        """format_color_violations should produce useful string representation."""
        from headless_excel.errors import format_color_violations

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
        s = format_color_violations(violations)
        assert "Financial color violations (1)" in s
        assert "Sheet1!A1" in s
        assert "HARDCODE" in s

    def test_format_empty(self):
        """Empty violations should report no violations."""
        from headless_excel.errors import format_color_violations

        s = format_color_violations({})
        assert s == "No color lint violations"


class TestErrorTruncation:
    """Test error truncation for LLM context management."""

    def test_get_set_max_errors_displayed(self):
        """get/set_max_errors_displayed should work correctly."""
        from headless_excel import get_max_errors_displayed, set_max_errors_displayed

        original = get_max_errors_displayed()
        try:
            assert original == 10  # Default

            set_max_errors_displayed(5)
            assert get_max_errors_displayed() == 5

            set_max_errors_displayed(100)
            assert get_max_errors_displayed() == 100
        finally:
            set_max_errors_displayed(original)

    def test_color_violations_truncated(self):
        """format_color_violations should truncate beyond max_errors_displayed."""
        from headless_excel import get_max_errors_displayed, set_max_errors_displayed
        from headless_excel.errors import format_color_violations

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(3)

            # Create 10 violations
            violations = {
                "Sheet1": [
                    ColorLintViolation(
                        cell=f"A{i}",
                        value=i * 100,
                        expected_color=Colors.HARDCODE,
                        current_color=None,
                    )
                    for i in range(1, 11)
                ]
            }
            s = format_color_violations(violations)

            # Should show total count
            assert "Financial color violations (10)" in s

            # Should show first 3
            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "Sheet1!A3" in s

            # Should NOT show beyond limit
            assert "Sheet1!A4" not in s
            assert "Sheet1!A10" not in s

            # Should show truncation message
            assert "... and 7 more violation(s) (truncated)" in s
        finally:
            set_max_errors_displayed(original)

    def test_color_violations_not_truncated_when_under_limit(self):
        """format_color_violations should not truncate when under limit."""
        from headless_excel import get_max_errors_displayed, set_max_errors_displayed
        from headless_excel.errors import format_color_violations

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(10)

            # Create 5 violations (under limit)
            violations = {
                "Sheet1": [
                    ColorLintViolation(
                        cell=f"A{i}",
                        value=i * 100,
                        expected_color=Colors.HARDCODE,
                        current_color=None,
                    )
                    for i in range(1, 6)
                ]
            }
            s = format_color_violations(violations)

            # Should show all 5
            assert "Financial color violations (5)" in s
            assert "Sheet1!A1" in s
            assert "Sheet1!A5" in s

            # Should NOT show truncation message
            assert "truncated" not in s
        finally:
            set_max_errors_displayed(original)

    def test_formula_error_str_truncated(self):
        """FormulaError.__str__ should truncate beyond max_errors_displayed."""
        from headless_excel import (
            FormulaError,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(3)

            # Create FormulaError with 10 errors
            err = FormulaError(
                errors={
                    "#REF!": [f"Sheet1!A{i}" for i in range(1, 6)],
                    "#DIV/0!": [f"Sheet1!B{i}" for i in range(1, 6)],
                },
                total=10,
            )
            s = str(err)

            # Should show total count
            assert "Formula errors (10)" in s

            # Should only show first 3 total
            error_lines = [line for line in s.split("\n") if line.startswith("  ") and "truncated" not in line]
            assert len(error_lines) == 3

            # Should show truncation message
            assert "... and 7 more error(s) (truncated)" in s
        finally:
            set_max_errors_displayed(original)

    def test_formula_error_str_with_details_truncated(self):
        """FormulaError.__str__ with error_details should truncate."""
        from headless_excel import (
            ErrorDetail,
            FormulaError,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(2)

            # Create FormulaError with 5 detailed errors
            err = FormulaError(
                errors={"#DIV/0!": [f"Sheet1!A{i}" for i in range(1, 6)]},
                total=5,
                error_details=[
                    ErrorDetail(
                        location=f"Sheet1!A{i}",
                        error="#DIV/0!",
                        formula=f"=B{i}/C{i}",
                        neighbors={f"Sheet1!B{i}": 100, f"Sheet1!C{i}": 0},
                    )
                    for i in range(1, 6)
                ],
            )
            s = str(err)

            # Should show total count
            assert "Formula errors (5)" in s

            # Should show first 2 with details
            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "formula=" in s

            # Should NOT show beyond limit
            assert "Sheet1!A3" not in s

            # Should show truncation message
            assert "... and 3 more error(s) (truncated)" in s
        finally:
            set_max_errors_displayed(original)

    def test_formula_error_repr_uses_truncated_str(self):
        """FormulaError.__repr__ should also be truncated to prevent context pollution."""
        from headless_excel import (
            FormulaError,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(2)

            err = FormulaError(
                errors={"#REF!": [f"Sheet1!A{i}" for i in range(1, 6)]},
                total=5,
            )

            # repr should be same as str (truncated)
            assert repr(err) == str(err)
            assert "truncated" in repr(err)
        finally:
            set_max_errors_displayed(original)

    def test_formula_error_not_truncated_when_under_limit(self):
        """FormulaError.__str__ should not truncate when under limit."""
        from headless_excel import (
            FormulaError,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(10)

            # Create FormulaError with 3 errors (under limit)
            err = FormulaError(
                errors={"#REF!": ["Sheet1!A1", "Sheet1!A2", "Sheet1!A3"]},
                total=3,
            )
            s = str(err)

            # Should show all 3
            assert "Formula errors (3)" in s
            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "Sheet1!A3" in s

            # Should NOT show truncation message
            assert "truncated" not in s
        finally:
            set_max_errors_displayed(original)

    def test_formula_error_empty(self):
        """FormulaError with no errors should not show truncation."""
        from headless_excel import FormulaError

        err = FormulaError(errors={}, total=0)
        s = str(err)
        assert s == "No formula errors"
        assert "truncated" not in s
