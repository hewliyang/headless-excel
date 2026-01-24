"""Code size warning hook example.

This hook warns if too much code is executed in a single CLI eval block.
Useful for encouraging modular scripts.

Only active for CLI eval commands (ctx._code is None for library usage).

Usage:
    # Copy to .headless-excel/hooks/code_size.py
    headless-excel eval file.xlsx "
    x = 1
    y = 2
    ...  # warns if > 2 lines
    "
"""

from headless_excel import ExcelContext, on_exit

MAX_LINES = 100


def count_code_lines(code: str) -> int:
    """Count non-empty, non-comment lines."""
    return sum(
        1
        for line in code.splitlines()
        if line.strip() and not line.strip().startswith("#")
    )


@on_exit
def warn_large_blocks(ctx: ExcelContext) -> None:
    """Warn if code block exceeds MAX_LINES."""
    if not ctx._code:
        return

    lines = count_code_lines(ctx._code)
    if lines > MAX_LINES:
        print(f"⚠ Code block has {lines} lines (recommended max: {MAX_LINES})")
        print("  Consider splitting into smaller blocks to avoid mistakes.")
