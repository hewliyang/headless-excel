"""Number format constants for Excel cells."""


class NumberFormats:
    """Essential number format strings for Excel.

    Example:
        >>> ws['A1'].number_format = NumberFormats.ACCOUNTING
        >>> ws['B1'].number_format = NumberFormats.PERCENTAGE
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
    DATE = "dd-mmm-yyyy"
    DATE_LONG = "mmmm d, yyyy"
