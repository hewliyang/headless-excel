"""Pytest configuration and fixtures."""

import pytest

from headless_excel import clear_hooks, start_daemon, stop_daemon


@pytest.fixture(scope="session", autouse=True)
def libreoffice_daemon():
    """Start LibreOffice daemon for fast recalc during tests (optional)."""
    try:
        start_daemon()
    except Exception:
        # Daemon failed to start - tests will use cold-start mode
        pass
    yield
    try:
        stop_daemon()
    except Exception:
        pass


@pytest.fixture(autouse=True)
def clean_hooks():
    """Clear hooks before and after each test to ensure isolation."""
    clear_hooks()
    yield
    clear_hooks()
