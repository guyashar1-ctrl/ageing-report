"""
file_utils.py
=============
File-handling utilities:
  - ZIP extraction to a temp directory
  - Auto-detection of BKMVDATA.TXT (and variants) inside a directory tree
  - Encoding detection / file reading with fallback
"""

import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional, Tuple

from config.constants import (
    BKMVDATA_FILENAMES,
    INI_FILENAMES,
    ENCODINGS_TO_TRY,
    RECORD_TYPE_ACCOUNT,
    RECORD_TYPE_MOVEMENT,
    RECORD_TYPE_FILE_HEADER,
)
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# ZIP extraction
# ---------------------------------------------------------------------------

def extract_zip(zip_bytes: bytes) -> str:
    """
    Extract a ZIP archive (given as raw bytes) into a new temp directory.

    Returns the path to the temp directory.
    The caller is responsible for deleting it when done
    (call cleanup_temp_dir).
    """
    tmp_dir = tempfile.mkdtemp(prefix="ageing_")
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = zf.namelist()
        zf.extractall(tmp_dir)
    log.info("ZIP extracted to %s  (%d entries)", tmp_dir, len(names))
    return tmp_dir


def cleanup_temp_dir(path: str) -> None:
    """Remove the temporary directory created by extract_zip."""
    try:
        shutil.rmtree(path, ignore_errors=True)
        log.debug("Cleaned up temp dir: %s", path)
    except Exception as exc:
        log.warning("Could not clean up temp dir %s: %s", path, exc)


# ---------------------------------------------------------------------------
# BKMVDATA / INI auto-detection
# ---------------------------------------------------------------------------

def find_bkmvdata(root: str) -> Optional[str]:
    """
    Walk the directory tree rooted at *root* and find the BKMVDATA.TXT file.

    Strategy:
    1. Look for files whose name matches known BKMVDATA filenames (case-insensitive).
    2. If not found by name, scan every .TXT file for A100/B100/C100 record prefixes.
    3. Return the path of the first match, or None.
    """
    root_path = Path(root)

    # --- Strategy 1: match by filename ---
    for dirpath, _dirs, files in os.walk(root_path):
        for fname in files:
            if fname.upper() in [n.upper() for n in BKMVDATA_FILENAMES]:
                found = os.path.join(dirpath, fname)
                log.info("Found BKMVDATA by name: %s", found)
                return found

    # --- Strategy 2: content sniffing ---
    known_prefixes = {
        RECORD_TYPE_FILE_HEADER,
        RECORD_TYPE_ACCOUNT,
        RECORD_TYPE_MOVEMENT,
    }
    for dirpath, _dirs, files in os.walk(root_path):
        for fname in files:
            if not fname.upper().endswith(".TXT") and "." in fname:
                continue  # skip non-txt files
            fpath = os.path.join(dirpath, fname)
            try:
                content = _read_first_bytes(fpath, 2048)
                if content and any(content.startswith(p) for p in known_prefixes):
                    log.info("Found BKMVDATA by content sniffing: %s", fpath)
                    return fpath
            except Exception:
                pass

    log.warning("BKMVDATA.TXT not found under %s", root)
    return None


def find_ini(root: str) -> Optional[str]:
    """Find the INI.TXT metadata file (optional – the system works without it)."""
    root_path = Path(root)
    for dirpath, _dirs, files in os.walk(root_path):
        for fname in files:
            if fname.upper() in [n.upper() for n in INI_FILENAMES]:
                found = os.path.join(dirpath, fname)
                log.info("Found INI file: %s", found)
                return found
    log.debug("INI.TXT not found under %s (not required)", root)
    return None


# ---------------------------------------------------------------------------
# Encoding-aware file reading
# ---------------------------------------------------------------------------

def read_text_file(path: str) -> Tuple[str, str]:
    """
    Read a text file, trying multiple encodings in order.

    Returns (content: str, encoding_used: str).
    Raises RuntimeError if no encoding succeeds.
    """
    for enc in ENCODINGS_TO_TRY:
        try:
            with open(path, "r", encoding=enc, errors="strict") as fh:
                content = fh.read()
            log.debug("Read %s with encoding %s", path, enc)
            return content, enc
        except (UnicodeDecodeError, LookupError):
            continue

    # Last resort: read with replacement characters
    with open(path, "r", encoding="latin-1", errors="replace") as fh:
        content = fh.read()
    log.warning("Read %s with latin-1 (replacement) – Hebrew may be garbled", path)
    return content, "latin-1"


def _read_first_bytes(path: str, n: int) -> Optional[str]:
    """Return the first *n* bytes of *path* decoded with the first working encoding."""
    for enc in ENCODINGS_TO_TRY:
        try:
            with open(path, "rb") as fh:
                raw = fh.read(n)
            return raw.decode(enc, errors="strict")
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

def validate_directory(path: str) -> bool:
    """Return True if *path* is an existing readable directory."""
    return os.path.isdir(path) and os.access(path, os.R_OK)


def validate_file(path: str) -> bool:
    """Return True if *path* is an existing readable file."""
    return os.path.isfile(path) and os.access(path, os.R_OK)


def normalize_account_key(key: str) -> str:
    """
    Normalize an account key for fuzzy matching.
    Strips whitespace, converts to upper-case.
    Does NOT strip leading zeros — many Israeli account numbers
    have meaningful leading zeros.  If you need zero-insensitive
    matching, apply lstrip('0') on top of this result.
    """
    return key.strip().upper()
