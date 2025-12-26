"""Tests for NumberFormats and Colors constants."""

from pathlib import Path

from openpyxl.styles import Font

from headless_excel import Colors, NumberFormats, create, run


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
            ctx.active["A1"].number_format = NumberFormats.ACCOUNTING
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.ACCOUNTING
            assert ctx.active["A1"].value == 1234.56

    def test_apply_percentage_format(self, tmp_path: Path):
        """NumberFormats.PERCENTAGE should work with cells."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 0.1575
            ctx.active["A1"].number_format = NumberFormats.PERCENTAGE_2DP
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.PERCENTAGE_2DP

    def test_apply_style_with_format(self, tmp_path: Path):
        """NumberFormats should work with range.apply_style()."""
        path = tmp_path / "test.xlsx"
        with create(path, auto_sync=False) as ctx:
            ctx.active["A1"] = 100
            ctx.active["A2"] = 200
            ctx.active.range("A1:A2").apply_style(number_format=NumberFormats.NUMBER)
            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].number_format == NumberFormats.NUMBER
            assert ctx.active["A2"].number_format == NumberFormats.NUMBER


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
            ctx.active["A1"].font = Font(color=Colors.HARDCODE)

            ctx.active["A2"] = "=A1*2"
            ctx.active["A2"].font = Font(color=Colors.FORMULA)

            ctx.workbook.save(path)

        with run(path, auto_sync=False) as ctx:
            assert ctx.active["A1"].font.color.rgb == Colors.HARDCODE
            assert ctx.active["A2"].font.color.rgb == Colors.FORMULA
