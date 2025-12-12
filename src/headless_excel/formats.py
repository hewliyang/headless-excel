"""Number format and color constants for Excel cells."""


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
