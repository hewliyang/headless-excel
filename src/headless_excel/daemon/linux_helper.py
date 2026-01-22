#!/usr/bin/env python3
"""
Linux daemon helper for headless-excel.

This script runs with system Python (which has access to the uno module)
and bridges TCP commands to LibreOffice via UNO socket connection.

Usage: /usr/bin/python3 linux_helper.py

Protocol:
    PING           -> PONG
    RECALC:/path   -> OK or ERROR:message
    QUIT           -> OK (then exits)
"""

import os
import signal
import socket
import subprocess
import sys
import time

DAEMON_PORT = 2023
UNO_PORT = 2002
UNO_HOST = "127.0.0.1"

soffice_proc = None
desktop = None


def start_soffice():
    """Start LibreOffice in socket listening mode."""
    global soffice_proc
    cmd = [
        "soffice",
        "--headless",
        "--invisible",
        "--nologo",
        "--norestore",
        f"--accept=socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ServiceManager",
    ]
    soffice_proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )
    return soffice_proc.pid


def connect_uno(retries=30, delay=0.5):
    """Connect to LibreOffice via UNO and return the desktop."""
    import uno
    from com.sun.star.connection import NoConnectException

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )

    for _ in range(retries):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext"
            )
            smgr = ctx.ServiceManager
            return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except NoConnectException:
            time.sleep(delay)

    raise RuntimeError("Failed to connect to LibreOffice via UNO")


def handle_recalc(filepath):
    """Open, recalculate, save, and close an Excel file."""
    import uno

    global desktop
    if desktop is None:
        desktop = connect_uno()

    url = uno.systemPathToFileUrl(filepath)
    doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
    doc.calculateAll()
    doc.store()
    doc.close(True)


def cleanup(signum=None, frame=None):
    """Clean up LibreOffice process on exit."""
    global soffice_proc, desktop
    if desktop:
        try:
            desktop.terminate()
        except Exception:
            pass
    if soffice_proc:
        try:
            os.killpg(soffice_proc.pid, signal.SIGTERM)
        except Exception:
            pass
    sys.exit(0)


def main():
    global desktop

    signal.signal(signal.SIGTERM, cleanup)
    signal.signal(signal.SIGINT, cleanup)

    # Start LibreOffice
    start_soffice()

    # Connect via UNO (with retries)
    try:
        desktop = connect_uno()
    except Exception as e:
        print(f"Failed to connect to LibreOffice: {e}", file=sys.stderr)
        cleanup()
        sys.exit(1)

    # Start TCP server
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind((UNO_HOST, DAEMON_PORT))
    server.listen(5)

    while True:
        try:
            conn, addr = server.accept()
            conn.settimeout(30)
            data = conn.recv(4096).decode("utf-8").strip()

            if data == "PING":
                conn.send(b"PONG")
            elif data == "QUIT":
                conn.send(b"OK")
                conn.close()
                server.close()
                cleanup()
                break
            elif data.startswith("RECALC:"):
                filepath = data[7:]
                try:
                    handle_recalc(filepath)
                    conn.send(b"OK")
                except Exception as e:
                    conn.send(f"ERROR:{e}".encode())
            else:
                conn.send(b"ERROR:Unknown command")
            conn.close()
        except TimeoutError:
            continue
        except Exception as e:
            print(f"Server error: {e}", file=sys.stderr)
            break

    cleanup()


if __name__ == "__main__":
    main()
