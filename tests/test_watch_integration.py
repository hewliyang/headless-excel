"""Integration tests for watch server - starts real servers."""

import asyncio
import urllib.error
import urllib.request

import pytest
import pytest_asyncio


@pytest.fixture
def xlsx_file(tmp_path):
    """Create a test xlsx file."""
    from openpyxl import Workbook

    path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.active["A1"] = "test"
    wb.save(path)
    wb.close()
    return path


@pytest_asyncio.fixture
async def watch_server(xlsx_file):
    """Start a watch server and yield ports, then cleanup."""
    from headless_excel.watch import watch

    http_port = 18080  # Use high ports to avoid conflicts
    ws_port = 18765

    # Start server in background task
    task = asyncio.create_task(
        watch(str(xlsx_file), http_port=http_port, ws_port=ws_port)
    )

    # Give server time to start
    await asyncio.sleep(0.3)

    yield {"http_port": http_port, "ws_port": ws_port, "file": xlsx_file}

    # Cleanup
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        pass


@pytest.mark.asyncio
async def test_http_server_serves_html(watch_server):
    """HTTP server should serve the viewer HTML."""
    port = watch_server["http_port"]

    loop = asyncio.get_event_loop()

    def fetch():
        with urllib.request.urlopen(f"http://localhost:{port}/", timeout=5) as resp:
            return resp.status, resp.read().decode()

    status, body = await loop.run_in_executor(None, fetch)

    assert status == 200
    assert "<html>" in body.lower()
    assert "test.xlsx" in body


@pytest.mark.asyncio
async def test_http_server_serves_xlsx(watch_server):
    """HTTP server should serve the xlsx file."""
    port = watch_server["http_port"]

    loop = asyncio.get_event_loop()

    def fetch():
        with urllib.request.urlopen(
            f"http://localhost:{port}/file.xlsx", timeout=5
        ) as resp:
            return resp.status, resp.headers.get("Content-Type")

    status, content_type = await loop.run_in_executor(None, fetch)

    assert status == 200
    assert "spreadsheet" in content_type


@pytest.mark.asyncio
async def test_websocket_connects(watch_server):
    """Should be able to connect via WebSocket."""
    import websockets

    port = watch_server["ws_port"]

    async with websockets.connect(f"ws://localhost:{port}", close_timeout=1) as ws:
        # Connection succeeded if we get here without exception
        assert ws is not None


@pytest.mark.asyncio
async def test_websocket_receives_reload_on_file_change(watch_server):
    """WebSocket should receive reload message when file changes."""
    import websockets
    from openpyxl import load_workbook

    port = watch_server["ws_port"]
    xlsx_file = watch_server["file"]

    async with websockets.connect(f"ws://localhost:{port}", close_timeout=1) as ws:
        # Modify the file
        await asyncio.sleep(0.1)  # Let connection stabilize

        wb = load_workbook(xlsx_file)
        wb.active["A1"] = "modified"
        wb.save(xlsx_file)
        wb.close()

        # Should receive reload message
        try:
            msg = await asyncio.wait_for(ws.recv(), timeout=2.0)
            assert msg == "reload"
        except asyncio.TimeoutError:
            pytest.fail("Did not receive reload message after file change")


@pytest.mark.asyncio
async def test_http_404_for_unknown_path(watch_server):
    """HTTP server should return 404 for unknown paths."""
    port = watch_server["http_port"]

    loop = asyncio.get_event_loop()

    def fetch():
        try:
            urllib.request.urlopen(f"http://localhost:{port}/unknown", timeout=5)
            return 200
        except urllib.error.HTTPError as e:
            return e.code

    status = await loop.run_in_executor(None, fetch)
    assert status == 404


def test_watch_rejects_missing_file(tmp_path):
    """watch() should raise for missing files."""
    from headless_excel.watch import watch

    missing = tmp_path / "missing.xlsx"

    with pytest.raises(FileNotFoundError):
        asyncio.run(watch(str(missing)))


def test_watch_rejects_non_xlsx(tmp_path):
    """watch() should raise for non-xlsx files."""
    from headless_excel.watch import watch

    csv_file = tmp_path / "test.csv"
    csv_file.write_text("a,b,c")

    with pytest.raises(ValueError, match=".xlsx"):
        asyncio.run(watch(str(csv_file)))
