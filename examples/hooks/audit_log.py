"""Audit logging hook example.

This hook demonstrates the factory pattern for stateful extensions.
It logs sync operations with timing information.

Output configuration:
    By default, hook output goes to stderr. You can configure this per-hook:
    - output="stderr" (default): prefixed output to stderr
    - output="stdout": prefixed output to stdout
    - output="none": suppress output entirely
"""

import time

from headless_excel import ExcelContext, ExtensionAPI, SyncResult, extension


@extension
def audit_extension(api: ExtensionAPI):
    """Factory pattern for stateful hooks."""
    sync_count = 0
    start_time = None

    @api.pre_sync(output="none")  # silent - just records timing
    def record_start(ctx: ExcelContext) -> None:
        nonlocal start_time
        start_time = time.time()

    @api.post_sync  # default: stderr
    def log_sync(ctx: ExcelContext, result: SyncResult) -> None:
        nonlocal sync_count, start_time
        sync_count += 1
        duration = time.time() - start_time if start_time else 0

        status = "✓" if result.success else f"✗ ({result.total_errors} errors)"
        print(f"[Sync #{sync_count}] {ctx.path.name} {status} ({duration:.2f}s)")

    @api.on_exit  # default: stderr
    def summary(ctx: ExcelContext) -> None:
        if sync_count > 0:
            print(f"[Audit] Total syncs: {sync_count}")
