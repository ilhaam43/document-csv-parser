"""Cross-process lock for Microsoft Excel automation jobs."""

from __future__ import annotations

import contextlib
import logging
import sys
import tempfile
import threading
import time
from pathlib import Path


LOGGER = logging.getLogger(__name__)
LOCK_STATE = threading.local()
LOCK_PATH = Path(tempfile.gettempdir()) / "document-csv-parser-excel.lock"


@contextlib.contextmanager
def excel_process_lock(timeout_seconds: int):
    """Serialize Excel automation across threads, workers, and Python processes."""
    lock_depth = getattr(LOCK_STATE, "depth", 0)
    if lock_depth:
        LOCK_STATE.depth = lock_depth + 1
        try:
            yield
        finally:
            LOCK_STATE.depth -= 1
        return

    LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    lock_handle = LOCK_PATH.open("a+b")
    deadline = time.monotonic() + timeout_seconds
    acquired = False
    try:
        if lock_handle.seek(0, 2) == 0:
            lock_handle.write(b"0")
            lock_handle.flush()

        if sys.platform == "win32":
            import msvcrt

            while not acquired:
                try:
                    lock_handle.seek(0)
                    msvcrt.locking(lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
                    acquired = True
                except OSError:
                    if time.monotonic() >= deadline:
                        raise TimeoutError("Timed out waiting for another Excel automation job to finish.")
                    time.sleep(0.5)
        else:
            import fcntl

            while not acquired:
                try:
                    fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    acquired = True
                except BlockingIOError:
                    if time.monotonic() >= deadline:
                        raise TimeoutError("Timed out waiting for another Excel automation job to finish.")
                    time.sleep(0.5)

        LOCK_STATE.depth = 1
        yield
    finally:
        LOCK_STATE.depth = 0
        if acquired:
            try:
                if sys.platform == "win32":
                    import msvcrt

                    lock_handle.seek(0)
                    msvcrt.locking(lock_handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)
            except OSError:
                LOGGER.warning("Could not release the Excel process lock cleanly.", exc_info=True)
        lock_handle.close()
