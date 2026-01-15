"""Magic numbers detection hook.

Detects numbers in math/comparison operations (*, /, +, -, >, <, etc.)
Skips function arguments to avoid false positives (VLOOKUP column index, etc.)
"""

from openpyxl.formula.tokenizer import Tokenizer

from headless_excel import ExcelContext, on_exit


def _extract_magic_numbers(formula: str) -> list[float]:
    """Extract magic numbers from a formula using openpyxl tokenizer."""
    if not formula.startswith("="):
        return []
    try:
        tok = Tokenizer(formula)
    except Exception:
        return []

    numbers = []
    prev_type = None
    func_depth = 0  # track if we're inside function args

    for t in tok.items:
        if t.type == "FUNC":
            if t.subtype == "OPEN":
                func_depth += 1
            elif t.subtype == "CLOSE":
                func_depth -= 1
        elif t.type == "OPERAND" and t.subtype == "NUMBER":
            # Magic if preceded by infix operator AND not inside function args
            if func_depth == 0 or prev_type == "OPERATOR-INFIX":
                numbers.append(float(t.value))
        prev_type = t.type
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
