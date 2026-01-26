This is a `uv` Python project.

## Running Tests

To run tests, use:

```bash
uv run pytest
```

For test coverage:

```bash
uv run pytest --cov=src --cov-report=term-missing
```

For verbose output:

```bash
uv run pytest -v
```

Tests are located in the `tests/` directory.

## Linting

```bash
uv run ruff check
```

To auto-fix issues:

```bash
uv run ruff check --fix
```

## Formatting

```bash
uv run ruff format
```

## Type Checking

```bash
uv run ty check
```
