"""Financial modeling color conventions hook.

This hook automatically applies and validates financial modeling color conventions:
- Blue (HARDCODE): Numeric input values / hardcoded numbers
- Black (FORMULA): Formulas without external sheet references
- Green (EXTERNAL_LINK): Formulas referencing other sheets
"""

from dataclasses import dataclass
from typing import Any

from openpyxl.styles import Font

from headless_excel import ExcelContext, on_exit, pre_sync


class Colors:
    HARDCODE = "FF0000FF"  # Blue - hardcoded input values
    FORMULA = "FF000000"  # Black - formula-driven cells
    EXTERNAL_LINK = "FF008000"  # Green - links to external sources


def _remove_string_literals(formula: str) -> str:
    result = []
    in_string = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if char == '"':
            if in_string:
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                in_string = False
            else:
                in_string = True
            result.append(char)
        elif not in_string:
            result.append(char)
        i += 1
    return "".join(result)


def infer_financial_color(value: Any) -> str | None:
    if isinstance(value, str) and value.startswith("="):
        formula_without_strings = _remove_string_literals(value)
        if "!" in formula_without_strings:
            return Colors.EXTERNAL_LINK
        return Colors.FORMULA
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return Colors.HARDCODE
    return None


@dataclass
class ColorLintViolation:
    cell: str
    value: Any
    expected_color: str
    current_color: str | None


def _color_name(color_code: str | None) -> str:
    if color_code is None:
        return "none"
    color_map = {
        Colors.HARDCODE: "HARDCODE (blue)",
        Colors.FORMULA: "FORMULA (black)",
        Colors.EXTERNAL_LINK: "EXTERNAL_LINK (green)",
    }
    return color_map.get(color_code, color_code)


@pre_sync
def auto_financial_colors(ctx: ExcelContext) -> None:
    """Apply financial modeling colors before each sync."""
    for ws in ctx.workbook.worksheets:
        for row in ws._formula_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    color = infer_financial_color(cell.value)
                    if color is None:
                        continue
                    old_font = cell.font
                    cell.font = Font(
                        name=old_font.name,
                        size=old_font.size,
                        bold=old_font.bold,
                        italic=old_font.italic,
                        underline=old_font.underline,
                        strike=old_font.strike,
                        color=color,
                    )


@on_exit
def lint_financial_colors(ctx: ExcelContext) -> None:
    """Check financial color conventions when context exits."""
    violations: list[ColorLintViolation] = []

    for ws in ctx.workbook.worksheets:
        for row in ws._formula_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    expected = infer_financial_color(cell.value)
                    if expected is None:
                        continue

                    current = None
                    if cell.font and cell.font.color:
                        color = cell.font.color
                        if color.type == "rgb" and color.rgb:
                            current = color.rgb
                        elif color.type == "theme":
                            current = f"theme:{color.theme}"

                    if current != expected:
                        violations.append(
                            ColorLintViolation(
                                cell=f"{ws.title}!{cell.coordinate}",
                                value=cell.value,
                                expected_color=expected,
                                current_color=current,
                            )
                        )

    if violations:
        print(f"Color violations ({len(violations)}):")
        for v in violations[:10]:  # Limit output
            print(f"  {v.cell}: need {_color_name(v.expected_color)}")
        if len(violations) > 10:
            print(f"  ... +{len(violations) - 10} more")
