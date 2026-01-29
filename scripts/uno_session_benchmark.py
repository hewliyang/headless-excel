#!/usr/bin/env python3
"""Benchmark the new UNO session-based sync vs old approach."""

import tempfile
import time
from pathlib import Path

from headless_excel import create, run
from headless_excel.daemon import is_daemon_running


def benchmark_sync():
    """Benchmark sync() with UNO session optimization."""
    print("=== UNO Session Sync Benchmark ===\n")
    print(f"Daemon running: {is_daemon_running()}\n")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = Path(f.name)

    try:
        # Create workbook with formulas
        print("1. Creating workbook with 100 cells + formula...")
        t0 = time.perf_counter()
        with create(path, auto_sync=False, overwrite=True) as ctx:
            ws = ctx.active
            for i in range(100):
                ws.cell(row=i + 1, column=1, value=i + 1)
            ws["B1"] = "=SUM(A1:A100)"
            ctx.sync()
        create_time = time.perf_counter() - t0
        print(f"   Create + first sync: {create_time * 1000:.2f}ms")

        # Modify and sync
        print("\n2. Modify + sync (this uses UNO session if daemon running)...")
        with run(path, auto_sync=False) as ctx:
            ws = ctx.active

            # First sync to establish baseline
            t0 = time.perf_counter()
            ws["A1"] = 1000
            ctx.sync()
            sync1_time = time.perf_counter() - t0
            print(f"   First modify + sync: {sync1_time * 1000:.2f}ms")
            print(f"   B1 value: {ws['B1'].value}")

            # Second sync
            t0 = time.perf_counter()
            ws["A1"] = 2000
            ctx.sync()
            sync2_time = time.perf_counter() - t0
            print(f"   Second modify + sync: {sync2_time * 1000:.2f}ms")
            print(f"   B1 value: {ws['B1'].value}")

            # Third sync
            t0 = time.perf_counter()
            ws["A1"] = 3000
            ctx.sync()
            sync3_time = time.perf_counter() - t0
            print(f"   Third modify + sync: {sync3_time * 1000:.2f}ms")
            print(f"   B1 value: {ws['B1'].value}")

        print("\n=== Summary ===")
        print(f"  First sync:  {sync1_time * 1000:.2f}ms")
        print(f"  Second sync: {sync2_time * 1000:.2f}ms")
        print(f"  Third sync:  {sync3_time * 1000:.2f}ms")
        print(
            f"  Average:     {(sync1_time + sync2_time + sync3_time) / 3 * 1000:.2f}ms"
        )

    finally:
        path.unlink(missing_ok=True)


def benchmark_large_file():
    """Benchmark with larger file."""
    print("\n\n=== Large File Benchmark (1000 rows x 10 cols) ===\n")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = Path(f.name)

    try:
        # Create larger workbook
        print("1. Creating large workbook...")
        t0 = time.perf_counter()
        with create(path, auto_sync=False, overwrite=True) as ctx:
            ws = ctx.active
            for row in range(1, 1001):
                for col in range(1, 11):
                    ws.cell(row=row, column=col, value=row * col)
            # Add formula column
            for row in range(1, 1001):
                ws.cell(row=row, column=11, value=f"=SUM(A{row}:J{row})")
            ctx.sync()
        create_time = time.perf_counter() - t0
        file_size = path.stat().st_size / 1024
        print(f"   Create + sync: {create_time:.2f}s")
        print(f"   File size: {file_size:.1f} KB")

        # Modify and sync multiple times
        print("\n2. Multiple modify + sync cycles...")
        sync_times = []
        with run(path, auto_sync=False) as ctx:
            ws = ctx.active

            for i in range(3):
                t0 = time.perf_counter()
                ws["A1"] = i * 1000
                ctx.sync()
                sync_time = time.perf_counter() - t0
                sync_times.append(sync_time)
                print(f"   Sync {i + 1}: {sync_time * 1000:.2f}ms, K1={ws['K1'].value}")

        print("\n=== Summary ===")
        print(f"  Average sync time: {sum(sync_times) / len(sync_times) * 1000:.2f}ms")

    finally:
        path.unlink(missing_ok=True)


if __name__ == "__main__":
    benchmark_sync()
    benchmark_large_file()
