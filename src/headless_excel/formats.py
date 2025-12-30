"""Number format and color constants for Excel cells."""

from typing import Any


class NumberFormats:
    """Essential number format strings for financial modeling.

    Example:
        >>> ws['A1'].number_format = NumberFormats.ACCOUNTING
        >>> ws['B1'].number_format = NumberFormats.PERCENTAGE
        >>> ctx.apply_style("Sheet", "C1:C10", number_format=NumberFormats.ACCOUNTING_0DP)
    """

    # Accounting formats - aligns $ and decimals, negatives in parentheses
    ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ACCOUNTING_0DP = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)'

    # Percentage formats
    PERCENTAGE = "0%"
    PERCENTAGE_1DP = "0.0%"
    PERCENTAGE_2DP = "0.00%"

    # Basic number format with comma separator
    NUMBER = "#,##0.00"
    NUMBER_0DP = "#,##0"

    # Date formats
    DATE = "m/d/yyyy"
    DATE_LONG = "mmmm d, yyyy"


class Colors:
    """Color codes for financial modeling text conventions.

    Excel colors are in ARGB hex format (Alpha, Red, Green, Blue).

    Example:
        >>> from openpyxl.styles import Font
        >>> ws['A1'].font = Font(color=Colors.HARDCODE)  # Blue text for inputs
        >>> ws['A2'].font = Font(color=Colors.FORMULA)   # Black text for formulas
    """

    HARDCODE = "FF0000FF"  # Blue - hardcoded input values
    FORMULA = "FF000000"  # Black - formula-driven cells
    EXTERNAL_LINK = "FF008000"  # Green - links to external sources


def infer_financial_color(value: Any) -> str | None:
    """Infer the conventional financial modeling color for a cell value.

    Financial modeling conventions:
    - Blue (HARDCODE): Numeric input values (int, float)
    - Black (FORMULA): Formulas that don't reference other sheets
    - Green (EXTERNAL_LINK): Formulas that reference other sheets (contain '!')
    - None: Text labels/strings (no color change, left as-is)

    Args:
        value: The cell value (formula string, number, etc.)

    Returns:
        Color code from Colors class, or None for text labels

    Example:
        >>> infer_financial_color(100)
        'FF0000FF'  # Blue - hardcoded
        >>> infer_financial_color("=A1+B1")
        'FF000000'  # Black - formula
        >>> infer_financial_color("=Sheet2!A1")
        'FF008000'  # Green - external link
        >>> infer_financial_color("Revenue")
        None  # Text labels are not colored
    """
    if isinstance(value, str) and value.startswith("="):
        # It's a formula - check if it references another sheet
        # The '!' indicates a sheet reference like Sheet2!A1 or [Book.xlsx]Sheet!A1
        # We avoid matching '!' inside quoted strings by checking it's outside quotes
        # Simple heuristic: if '!' appears and it's not inside a string literal
        formula_without_strings = _remove_string_literals(value)
        if "!" in formula_without_strings:
            return Colors.EXTERNAL_LINK
        return Colors.FORMULA
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return Colors.HARDCODE
    # Text labels and other values are not colored
    return None


def _remove_string_literals(formula: str) -> str:
    """Remove string literals from formula to avoid false positive '!' matches.

    Handles Excel string literals which use double quotes.

    Example:
        >>> _remove_string_literals('=A1&"Hello!"')
        '=A1&""'
    """
    result = []
    in_string = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if char == '"':
            if in_string:
                # Check for escaped quote ("")
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
