import pytest

from headless_excel import clear_hooks, start_daemon, stop_daemon


@pytest.fixture(scope="session", autouse=True)
def libreoffice_daemon():
    try:
        start_daemon()
    except Exception:
        pass
    yield
    try:
        stop_daemon()
    except Exception:
        pass


@pytest.fixture(autouse=True)
def clean_hooks():
    clear_hooks()
    yield
    clear_hooks()
