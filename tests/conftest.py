"""Pytest configuration and fixtures."""

import pytest

from headless_excel import clear_hooks


@pytest.fixture(autouse=True)
def clean_hooks():
    """Clear hooks before and after each test to ensure isolation."""
    clear_hooks()
    yield
    clear_hooks()
