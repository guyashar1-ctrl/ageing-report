"""
logger.py
=========
Centralised logging for the Ageing Report system.

All modules call get_logger(__name__) to obtain a standard Python logger.
Log records are also accumulated in a list so that the Excel "לוג שגיאות"
sheet can be populated at the end of the run.
"""

import logging
import sys
from typing import List, Tuple
from datetime import datetime

# ---------------------------------------------------------------------------
# Module-level log accumulator  (level_name, message, timestamp)
# ---------------------------------------------------------------------------
_log_records: List[Tuple[str, str, str]] = []


class _AccumulatorHandler(logging.Handler):
    """Handler that appends every log record to _log_records."""

    def emit(self, record: logging.LogRecord) -> None:
        ts = datetime.fromtimestamp(record.created).strftime("%d/%m/%Y %H:%M:%S")
        _log_records.append((record.levelname, record.getMessage(), ts))


# ---------------------------------------------------------------------------
# Root logger configuration  (called once at import time)
# ---------------------------------------------------------------------------
def _configure_root_logger() -> None:
    root = logging.getLogger()
    if root.handlers:
        return  # already configured

    root.setLevel(logging.DEBUG)

    # Console handler (INFO and above)
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    fmt = logging.Formatter("[%(levelname)s] %(name)s – %(message)s")
    console.setFormatter(fmt)
    root.addHandler(console)

    # Accumulator handler (all levels)
    acc = _AccumulatorHandler()
    acc.setLevel(logging.DEBUG)
    root.addHandler(acc)


_configure_root_logger()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_logger(name: str) -> logging.Logger:
    """Return a module-level logger."""
    return logging.getLogger(name)


def get_accumulated_logs() -> List[Tuple[str, str, str]]:
    """
    Return all accumulated log records as a list of (level, message, timestamp).
    Useful for writing to the Excel log sheet.
    """
    return list(_log_records)


def clear_accumulated_logs() -> None:
    """Clear the in-memory log accumulator (call at start of each run)."""
    _log_records.clear()
