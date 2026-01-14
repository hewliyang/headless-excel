"""Magic numbers detection hook.

Requires: pip install formulas

Detects numbers in math/comparison operations (*, /, +, -, >, <, etc.)
Skips function arguments to avoid false positives (VLOOKUP column index, etc.)
"""

import formulas
from formulas.tokens.operand import Number
from formulas.tokens.operator import OperatorToken

from headless_excel import ExcelContext, on_exit


def _extract_magic_numbers(formula: str) -> list[float]:
    """Extract magic numbers from a formula using AST parsing."""
    if not formula.startswith("="):
        return []
    try:
        tokens, _ = formulas.Parser().ast(formula)
    except Exception:
        return []

    numbers = []
    prev = None
    for token in tokens:
        if isinstance(token, Number):
            # Magic if preceded by operator, not function argument (separator)
            if prev is None or isinstance(prev, OperatorToken):
                numbers.append(float(token.name))
        prev = token
    return numbers


@on_exit
def lint_magic_numbers(ctx: ExcelContext) -> None:
    """Check for magic numbers when context exits."""
    violations: list[tuple[str, str, float]] = []

    for ws in ctx.workbook.worksheets:
        for row in ws._formula_ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue

                for num in _extract_magic_numbers(value):
                    violations.append((f"{ws.title}!{cell.coordinate}", value, num))

    if violations:
        print(f"\n⚠️  Magic Numbers ({len(violations)}):")
        for cell, formula, num in violations[:10]:
            print(f"   {cell}: {num} in {formula}")
        if len(violations) > 10:
            print(f"   ... +{len(violations) - 10} more")
