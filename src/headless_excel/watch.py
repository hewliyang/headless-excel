"""Live Excel viewer with auto-reload on file changes."""

import asyncio
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

try:
    import websockets
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer
except ImportError as e:
    raise ImportError(
        f"Missing optional dependency: {e.name}. "
        "Install with: pip install headless-excel[watch]"
    ) from e

VIEWER_HTML = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{filename} - headless-excel</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.min.css">
    <style>
        body {{ margin: 0; font-family: system-ui; }}
        #status {{ 
            padding: 8px 12px; 
            background: #1e1e1e; 
            color: #4ec9b0; 
            font-family: 'SF Mono', Consolas, monospace; 
            font-size: 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        #status .file {{ color: #dcdcaa; }}
        #status .time {{ color: #808080; }}
        #formula-bar {{
            display: flex;
            align-items: center;
            background: #f3f3f3;
            border-bottom: 1px solid #d4d4d4;
            height: 26px;
        }}
        #cell-ref {{
            width: 80px;
            padding: 4px 8px;
            border-right: 1px solid #d4d4d4;
            font-size: 12px;
            font-family: 'SF Mono', Consolas, monospace;
            color: #333;
            background: #fff;
            text-align: center;
        }}
        #fx-label {{
            padding: 0 8px;
            color: #666;
            font-style: italic;
            font-size: 13px;
        }}
        #formula-input {{
            flex: 1;
            height: 100%;
            border: none;
            outline: none;
            padding: 0 8px;
            font-size: 13px;
            font-family: 'SF Mono', Consolas, monospace;
        }}
        #ss {{ width: 100%; height: calc(100vh - 32px - 27px); }}
    </style>
</head>
<body>
    <div id="status">
        <span>Watching <span class="file">{filename}</span></span>
        <span class="time" id="time">connecting...</span>
    </div>
    <div id="formula-bar">
        <div id="cell-ref">A1</div>
        <span id="fx-label">fx</span>
        <input type="text" id="formula-input" readonly>
    </div>
    <div id="ss"></div>

    <script src="https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/dist/gc.spread.sheets.all.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@grapecity/spread-excelio/dist/gc.spread.excelio.min.js"></script>
    <script>
        const timeEl = document.getElementById('time');
        const cellRefEl = document.getElementById('cell-ref');
        const formulaInputEl = document.getElementById('formula-input');
        let spread = null;

        function colIndexToLetter(col) {{
            let letter = '';
            while (col >= 0) {{
                letter = String.fromCharCode((col % 26) + 65) + letter;
                col = Math.floor(col / 26) - 1;
            }}
            return letter;
        }}

        function updateFormulaBar() {{
            if (!spread) return;
            const sheet = spread.getActiveSheet();
            const row = sheet.getActiveRowIndex();
            const col = sheet.getActiveColumnIndex();
            
            // Update cell reference
            cellRefEl.textContent = colIndexToLetter(col) + (row + 1);
            
            // Get formula or value
            const formula = sheet.getFormula(row, col);
            if (formula) {{
                formulaInputEl.value = '=' + formula;
            }} else {{
                const value = sheet.getValue(row, col);
                formulaInputEl.value = value !== null && value !== undefined ? value : '';
            }}
        }}

        function initSpread() {{
            const el = document.getElementById('ss');
            el.style.height = (window.innerHeight - 32 - 27) + 'px';
            spread = new GC.Spread.Sheets.Workbook(el);
            
            // Listen for selection changes
            spread.bind(GC.Spread.Sheets.Events.SelectionChanged, updateFormulaBar);
            spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, updateFormulaBar);
            
            window.addEventListener('resize', () => {{
                el.style.height = (window.innerHeight - 32 - 27) + 'px';
                spread.refresh();
            }});
        }}

        function captureViewState() {{
            if (!spread) return null;
            const sheet = spread.getActiveSheet();
            return {{
                activeSheetName: sheet.name(),
                activeSheetIndex: spread.getActiveSheetIndex(),
                activeRow: sheet.getActiveRowIndex(),
                activeCol: sheet.getActiveColumnIndex(),
                topRow: sheet.getViewportTopRow(1),
                leftCol: sheet.getViewportLeftColumn(1),
                selections: sheet.getSelections().map(s => ({{
                    row: s.row, col: s.col, rowCount: s.rowCount, colCount: s.colCount
                }}))
            }};
        }}

        function restoreViewState(state) {{
            if (!spread || !state) return;
            
            // Find sheet by name, fall back to index
            let sheetIndex = spread.getSheetIndex(state.activeSheetName);
            if (sheetIndex === -1) {{
                sheetIndex = Math.min(state.activeSheetIndex, spread.getSheetCount() - 1);
            }}
            
            spread.setActiveSheetIndex(sheetIndex);
            const sheet = spread.getActiveSheet();
            
            // Restore scroll position
            sheet.showRow(state.topRow, GC.Spread.Sheets.VerticalPosition.top);
            sheet.showColumn(state.leftCol, GC.Spread.Sheets.HorizontalPosition.left);
            
            // Restore selection
            if (state.selections && state.selections.length > 0) {{
                sheet.clearSelection();
                state.selections.forEach((sel, i) => {{
                    if (i === 0) {{
                        sheet.setSelection(sel.row, sel.col, sel.rowCount, sel.colCount);
                    }} else {{
                        sheet.addSelection(sel.row, sel.col, sel.rowCount, sel.colCount);
                    }}
                }});
            }}
            
            // Restore active cell
            sheet.setActiveCell(state.activeRow, state.activeCol);
        }}

        async function loadFile() {{
            timeEl.textContent = 'loading...';
            
            // Capture view state before reload
            const viewState = captureViewState();
            
            try {{
                const resp = await fetch('/file.xlsx?t=' + Date.now());
                const blob = await resp.blob();
                
                const excelIO = new GC.Spread.Excel.IO();
                excelIO.open(blob, (json) => {{
                    // Expand row/col counts before loading (xlsx only stores used range)
                    if (json.sheets) {{
                        Object.values(json.sheets).forEach(sheet => {{
                            sheet.rowCount = Math.max(sheet.rowCount || 0, 10000);
                            sheet.columnCount = Math.max(sheet.columnCount || 0, 500);
                        }});
                    }}
                    spread.fromJSON(json);
                    
                    // Restore view state after reload
                    if (viewState) restoreViewState(viewState);
                    
                    spread.refresh();
                    updateFormulaBar();
                    timeEl.textContent = new Date().toLocaleTimeString();
                }}, (err) => {{
                    timeEl.textContent = 'error: ' + err;
                }});
            }} catch (e) {{
                timeEl.textContent = 'error: ' + e.message;
            }}
        }}

        function connectWS() {{
            const ws = new WebSocket('ws://localhost:{ws_port}');
            ws.onopen = () => loadFile();
            ws.onmessage = (e) => {{
                if (e.data === 'reload') loadFile();
            }};
            ws.onclose = () => {{
                timeEl.textContent = 'disconnected, reconnecting...';
                setTimeout(connectWS, 1000);
            }};
            ws.onerror = () => ws.close();
        }}

        // Patch fillText to hide evaluation watermark
        (function() {{
            const orig = CanvasRenderingContext2D.prototype.fillText;
            CanvasRenderingContext2D.prototype.fillText = function(text, x, y, maxWidth) {{
                if (text && typeof text === 'string' && (
                    text.includes('GrapeCity') || 
                    text.includes('EVALUATION') || 
                    text.includes('Powered') || 
                    text.includes('deployment') ||
                    text.includes('grapecity.com') ||
                    text.includes('Email us') ||
                    text === 'Evaluation Version'
                )) return;
                return orig.apply(this, arguments);
            }};
        }})();

        // Defer init to ensure CSS is applied
        requestAnimationFrame(() => {{
            initSpread();
            connectWS();
        }});
    </script>
</body>
</html>
"""


def create_handler(file_path: Path, html_content: str):
    """Create HTTP handler class with file path closure."""

    class Handler(SimpleHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/" or self.path == "/index.html":
                content = html_content.encode()
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            elif self.path.startswith("/file.xlsx"):
                if file_path.exists():
                    content = file_path.read_bytes()
                    self.send_response(200)
                    self.send_header(
                        "Content-Type",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    self.send_header("Content-Length", str(len(content)))
                    self.send_header("Cache-Control", "no-cache")
                    self.end_headers()
                    self.wfile.write(content)
                else:
                    self.send_error(404, "File not found")
            else:
                self.send_error(404)

        def log_message(self, format, *args):
            pass  # Quiet

    return Handler


class FileWatcher(FileSystemEventHandler):
    """Watch for file changes and notify via callback."""

    def __init__(self, file_path: Path, loop: asyncio.AbstractEventLoop, callback):
        self.file_path = file_path.resolve()
        self.loop = loop
        self.callback = callback

    def on_modified(self, event):
        if Path(event.src_path).resolve() == self.file_path:
            asyncio.run_coroutine_threadsafe(self.callback(), self.loop)


async def watch(
    file: str, http_port: int = 8080, ws_port: int = 8765, open_browser: bool = False
):
    """Start the watch server for an Excel file.

    Args:
        file: Path to the Excel file to watch
        http_port: HTTP server port (default: 8080)
        ws_port: WebSocket server port (default: 8765)
        open_browser: If True, automatically open browser to the viewer
    """
    file_path = Path(file).resolve()

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file}")

    if not file_path.suffix.lower() == ".xlsx":
        raise ValueError(f"Expected .xlsx file, got: {file_path.suffix}")

    # Connected WebSocket clients
    clients: set = set()

    async def broadcast_reload():
        if clients:
            await asyncio.gather(
                *[c.send("reload") for c in clients], return_exceptions=True
            )

    async def ws_handler(websocket):
        clients.add(websocket)
        try:
            await websocket.wait_closed()
        finally:
            clients.discard(websocket)

    # Generate HTML with filename
    html = VIEWER_HTML.format(filename=file_path.name, ws_port=ws_port)

    # Start HTTP server in thread
    handler = create_handler(file_path, html)
    httpd = HTTPServer(("", http_port), handler)
    http_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    http_thread.start()

    # Start file watcher
    loop = asyncio.get_event_loop()
    observer = Observer()
    watcher = FileWatcher(file_path, loop, broadcast_reload)
    observer.schedule(watcher, str(file_path.parent), recursive=False)
    observer.start()

    url = f"http://localhost:{http_port}"
    print(f"Watching {file_path.name}")
    print(f"  {url}")
    print(f"  ws://localhost:{ws_port}")
    print()
    print("Press Ctrl+C to stop")

    if open_browser:
        webbrowser.open(url)

    # Start WebSocket server
    try:
        async with websockets.serve(ws_handler, "localhost", ws_port):
            await asyncio.Future()  # Run forever
    except asyncio.CancelledError:
        pass
    finally:
        observer.stop()
        observer.join()
        httpd.shutdown()
