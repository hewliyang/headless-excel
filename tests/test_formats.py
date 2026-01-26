from pathlib import Path

import headless_excel.errors as errors
from headless_excel import (
    ErrorDetail,
    FormulaError,
    NumberFormats,
    SyncResult,
    create,
    get_max_errors_displayed,
    run,
)


class TestNumberFormats:
    def test_accounting_formats_exist(self):
        assert hasattr(NumberFormats, "ACCOUNTING")
        assert hasattr(NumberFormats, "ACCOUNTING_0DP")
        assert isinstance(NumberFormats.ACCOUNTING, str)
        assert isinstance(NumberFormats.ACCOUNTING_0DP, str)

    def test_percentage_formats_exist(self):
        assert NumberFormats.PERCENTAGE == "0%"
        assert NumberFormats.PERCENTAGE_1DP == "0.0%"
        assert NumberFormats.PERCENTAGE_2DP == "0.00%"

    def test_number_formats_exist(self):
        assert NumberFormats.NUMBER == "#,##0.00"
        assert NumberFormats.NUMBER_0DP == "#,##0"

    def test_date_formats_exist(self):
        assert NumberFormats.DATE == "dd-mmm-yyyy"
        assert NumberFormats.DATE_LONG == "mmmm d, yyyy"

    def test_apply_accounting_format(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 1234.56
            ctx.active["A1"].number_format = NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.ACCOUNTING  # type: ignore[union-attr]
            assert ctx.active["A1"].value == 1234.56  # type: ignore[union-attr]

    def test_apply_percentage_format(self, tmp_path: Path):
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 0.1575
            ctx.active["A1"].number_format = NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.PERCENTAGE_2DP  # type: ignore[union-attr]

    def test_apply_style_with_format(self, tmp_path: Path):
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
    def test_max_errors_displayed(self):
        original = get_max_errors_displayed()
        try:
            assert original == 10

            errors.MAX_ERRORS_DISPLAYED = 5
            assert get_max_errors_displayed() == 5

            errors.MAX_ERRORS_DISPLAYED = 100
            assert get_max_errors_displayed() == 100
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_formula_error_str_truncated(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 3

            err = FormulaError(
                errors={
                    "#REF!": [f"Sheet1!A{i}" for i in range(1, 6)],
                    "#DIV/0!": [f"Sheet1!B{i}" for i in range(1, 6)],
                },
                total=10,
            )
            s = str(err)

            assert "Formula errors (10)" in s

            error_lines = [
                line
                for line in s.split("\n")
                if line.startswith("  ") and "truncated" not in line
            ]
            assert len(error_lines) == 3

            assert "... and 7 more error(s) (truncated)" in s
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_formula_error_str_with_details_truncated(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 2

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

            assert "Formula errors (5)" in s

            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "formula=" in s

            assert "Sheet1!A3" not in s

            assert "... and 3 more error(s) (truncated)" in s
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_formula_error_repr_uses_truncated_str(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 2

            err = FormulaError(
                errors={"#REF!": [f"Sheet1!A{i}" for i in range(1, 6)]},
                total=5,
            )

            assert repr(err) == str(err)
            assert "truncated" in repr(err)
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_formula_error_not_truncated_when_under_limit(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 10

            err = FormulaError(
                errors={"#REF!": ["Sheet1!A1", "Sheet1!A2", "Sheet1!A3"]},
                total=3,
            )
            s = str(err)

            assert "Formula errors (3)" in s
            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "Sheet1!A3" in s

            assert "truncated" not in s
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_formula_error_empty(self):
        err = FormulaError(errors={}, total=0)
        s = str(err)
        assert s == "No formula errors"
        assert "truncated" not in s

    def test_sync_result_repr_truncated(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 3

            result = SyncResult(
                success=False,
                total_errors=10,
                errors={
                    "#REF!": [f"Sheet1!A{i}" for i in range(1, 6)],
                    "#DIV/0!": [f"Sheet1!B{i}" for i in range(1, 6)],
                },
            )
            s = repr(result)

            assert "total_errors=10" in s

            error_lines = [
                line
                for line in s.split("\n")
                if line.startswith("  ") and "truncated" not in line
            ]
            assert len(error_lines) == 3

            assert "... and 7 more error(s) (truncated)" in s
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_sync_result_repr_with_details_truncated(self):
        original = get_max_errors_displayed()
        try:
            errors.MAX_ERRORS_DISPLAYED = 2

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

            assert "total_errors=5" in s

            assert "Sheet1!A1" in s
            assert "Sheet1!A2" in s
            assert "formula=" in s

            assert "Sheet1!A3" not in s

            assert "... and 3 more error(s) (truncated)" in s
        finally:
            errors.MAX_ERRORS_DISPLAYED = original

    def test_sync_result_repr_success(self):
        result = SyncResult(success=True)
        s = repr(result)
        assert s == "SyncResult(success=True)"

    def test_sync_result_str_same_as_repr(self):
        result = SyncResult(
            success=False,
            total_errors=2,
            errors={"#REF!": ["Sheet1!A1", "Sheet1!A2"]},
        )
        assert str(result) == repr(result)
