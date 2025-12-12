#!/usr/bin/env python3
"""Convert bash heredoc Python style to regular Python code blocks."""

import re
import sys
from pathlib import Path


def convert_heredoc_to_python(content: str) -> str:
    """Convert bash heredoc style to regular Python code blocks."""

    # Pattern to match: uv run python <<'EOF' ... EOF
    # Also matches: uv run python<<'EOF' (without space)
    pattern = r"```bash\s*\n\s*uv run python\s*<<'EOF'\s*\n(.*?)\nEOF\s*\n```"

    def replace_match(match):
        python_code = match.group(1)
        return f"```py\n{python_code}\n```"

    # Replace all occurrences
    converted = re.sub(pattern, replace_match, content, flags=re.DOTALL)

    return converted


def main():
    if len(sys.argv) < 2:
        print("Usage: uv run scripts/convert_heredoc.py <file>")
        print("   or: uv run scripts/convert_heredoc.py <file> --in-place")
        sys.exit(1)

    file_path = Path(sys.argv[1])
    in_place = "--in-place" in sys.argv or "-i" in sys.argv

    if not file_path.exists():
        print(f"Error: File '{file_path}' not found")
        sys.exit(1)

    # Read the file
    content = file_path.read_text()

    # Convert heredoc to Python blocks
    converted = convert_heredoc_to_python(content)

    if in_place:
        file_path.write_text(converted)
        print(f"✓ Converted {file_path} in place")
    else:
        # Output to stdout
        print(converted)


if __name__ == "__main__":
    main()
