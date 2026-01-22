"""Tests for NumberFormats constants."""

from pathlib import Path

from headless_excel import (
    NumberFormats,
    create,
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
        assert NumberFormats.DATE == "dd-mmm-yyyy"
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
            error_lines = [
                line
                for line in s.split("\n")
                if line.startswith("  ") and "truncated" not in line
            ]
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

    def test_sync_result_repr_truncated(self):
        """SyncResult.__repr__ should truncate beyond max_errors_displayed."""
        from headless_excel import (
            SyncResult,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(3)

            # Create SyncResult with 10 errors
            result = SyncResult(
                success=False,
                total_errors=10,
                errors={
                    "#REF!": [f"Sheet1!A{i}" for i in range(1, 6)],
                    "#DIV/0!": [f"Sheet1!B{i}" for i in range(1, 6)],
                },
            )
            s = repr(result)

            # Should show total count
            assert "total_errors=10" in s

            # Should only show first 3 total
            error_lines = [
                line
                for line in s.split("\n")
                if line.startswith("  ") and "truncated" not in line
            ]
            assert len(error_lines) == 3

            # Should show truncation message
            assert "... and 7 more error(s) (truncated)" in s
        finally:
            set_max_errors_displayed(original)

    def test_sync_result_repr_with_details_truncated(self):
        """SyncResult.__repr__ with error_details should truncate."""
        from headless_excel import (
            ErrorDetail,
            SyncResult,
            get_max_errors_displayed,
            set_max_errors_displayed,
        )

        original = get_max_errors_displayed()
        try:
            set_max_errors_displayed(2)

            # Create SyncResult with 5 detailed errors
            result = SyncResult(
                success=False,
                total_errors=5,
                errors={"#DIV/0!": [f"Sheet1!A{i}" for i in range(1, 6)]},
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
            s = repr(result)

            # Should show total count
            assert "total_errors=5" in s

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

    def test_sync_result_repr_success(self):
        """SyncResult.__repr__ should be concise on success."""
        from headless_excel import SyncResult

        result = SyncResult(success=True)
        s = repr(result)
        assert s == "SyncResult(success=True)"

    def test_sync_result_str_same_as_repr(self):
        """SyncResult.__str__ should be same as __repr__."""
        from headless_excel import SyncResult

        result = SyncResult(
            success=False,
            total_errors=2,
            errors={"#REF!": ["Sheet1!A1", "Sheet1!A2"]},
        )
        assert str(result) == repr(result)
