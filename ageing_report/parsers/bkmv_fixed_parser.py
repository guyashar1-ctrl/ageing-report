"""
bkmv_fixed_parser.py
====================
Parses TXT.BKMVDATA using the EXACT field-extraction logic verified by the user.

B100  –  Account master data
    account_number  = record[172:178].strip()
    account_name    = tab-split field[2]  (reported in debug if empty)
    opening_balance = tab-split fields[4]+[5] (side + amount)

C100  –  Movement records
    account_number  = abs( int( re.findall(r'[+-]\\d+', record)[-1] ) )
    movement        = signed float  (+debit / -credit)
                      sign from the last signed number,
                      magnitude from tab-split field[10]

Public API
----------
parse_bkmvdata_fixed(path)
    -> (b100_records: dict[str, B100Fixed], c100_movements: list[C100Fixed],
        debug_samples: list[str], warnings: list[str])

    b100_records  keyed by the normalised account_number string
    c100_movements  one entry per C100 line
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from utils.file_utils import read_text_file
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class B100Fixed:
    account_number:      str        # from record[172:178].strip()
    account_name:        str        # from tab-split field[2]
    opening_balance:     float      # signed: positive=debit, negative=credit
    raw_line:            str
    # Debug fields – show what positions / fields were used
    debug_name_source:    str = ""
    debug_balance_source: str = ""


@dataclass
class C100Fixed:
    account_number:  int    # abs(int(last_signed_match))  — always positive
    movement:        float  # signed: positive=debit, negative=credit
    raw_line:        str
    debug_last_match: str = ""   # the raw matched token, e.g. "+0000000000020005"
    debug_amount_raw: str = ""   # the raw amount string from tab field[10]


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_bkmvdata_fixed(
    path: str,
) -> Tuple[Dict[str, B100Fixed], List[C100Fixed], List[str], List[str]]:
    """
    Parse BKMVDATA and return:
        b100_records   – dict keyed by account_number string
        c100_movements – list of all C100 entries
        debug_samples  – first 10 B100 extractions for inspection
        warnings       – any parse issues
    """
    log.info("parse_bkmvdata_fixed: %s", path)
    content, encoding = read_text_file(path)
    log.info("Encoding: %s", encoding)

    b100_records:   Dict[str, B100Fixed] = {}
    c100_movements: List[C100Fixed]      = []
    warnings:       List[str]            = []
    debug_samples:  List[str]            = []

    b100_count = 0
    c100_count = 0

    for lineno, raw_line in enumerate(content.splitlines(), start=1):
        if not raw_line.strip():
            continue

        # Determine record type from the first 4 characters of the raw line.
        rec_type = raw_line[:4].upper().strip()

        if rec_type == "B100":
            b100_count += 1
            rec = _parse_b100(raw_line, lineno, warnings)
            if rec:
                key = _norm(rec.account_number)
                if key in b100_records:
                    warnings.append(
                        f"Line {lineno}: duplicate B100 account '{rec.account_number}' – overwriting"
                    )
                b100_records[key] = rec

                # Collect first 10 debug samples
                if len(debug_samples) < 10:
                    preview = raw_line[:200].replace("\t", "→")
                    debug_samples.append(
                        f"B100 line {lineno}: "
                        f"acct=[{rec.account_number!r}] pos172:178  "
                        f"name=[{rec.account_name!r}] {rec.debug_name_source}  "
                        f"ob={rec.opening_balance:.2f} {rec.debug_balance_source}  "
                        f"raw_preview={preview!r}"
                    )

        elif rec_type == "C100":
            c100_count += 1
            mvmt = _parse_c100(raw_line, lineno, warnings)
            if mvmt:
                c100_movements.append(mvmt)

    log.info(
        "Parsed %d B100 lines → %d valid accounts | "
        "%d C100 lines → %d valid movements",
        b100_count, len(b100_records),
        c100_count, len(c100_movements),
    )

    if b100_count == 0:
        warnings.append(
            "No B100 records found. "
            "Check that the file is TXT.BKMVDATA and not a different file."
        )
    if c100_count == 0:
        warnings.append("No C100 records found.")

    return b100_records, c100_movements, debug_samples, warnings


# ---------------------------------------------------------------------------
# B100 parsing  –  exact position logic
# ---------------------------------------------------------------------------

def _parse_b100(
    raw_line: str,
    lineno:   int,
    warnings: List[str],
) -> Optional[B100Fixed]:
    """
    Extract account_number from fixed character position 172:178.
    Extract account_name and opening_balance from tab-split fields.
    """
    # ── Account number (EXACT per spec) ─────────────────────────────────
    if len(raw_line) > 178:
        account_number = raw_line[172:178].strip()
    elif len(raw_line) > 172:
        account_number = raw_line[172:].strip()
    else:
        # Line too short – log and skip
        warnings.append(
            f"Line {lineno}: B100 shorter than 172 chars "
            f"(len={len(raw_line)}) – cannot extract account number"
        )
        return None

    if not account_number:
        warnings.append(f"Line {lineno}: B100 position 172:178 is blank – skipping")
        return None

    # ── Tab-split for remaining fields ───────────────────────────────────
    fields = raw_line.split("\t")

    # Account name (field[2])
    account_name = ""
    name_source  = ""
    if len(fields) > 2 and fields[2].strip():
        account_name = fields[2].strip()
        name_source  = "tab_field[2]"
    elif len(raw_line) > 79:
        # Fallback: fixed positions 29:79 (standard Takken-7 B100 name field)
        candidate = raw_line[29:79].strip()
        if candidate:
            account_name = candidate
            name_source  = "fixed_pos_29:79"
        else:
            name_source = "empty (not found)"
    else:
        name_source = "line too short for fixed fallback"

    # Opening balance (tab fields[4] = side, fields[5] = amount)
    opening_balance  = 0.0
    balance_source   = ""

    if len(fields) > 5:
        side_raw   = fields[4].strip()
        amount_raw = fields[5].strip()
        amount     = _parse_amount(amount_raw)
        if amount != 0.0:
            if side_raw == "2":   # credit → negative
                opening_balance = -amount
                balance_source  = "tab_fields[4,5]: side=credit"
            else:                 # "1" debit or unknown → positive
                opening_balance = amount
                balance_source  = f"tab_fields[4,5]: side={side_raw!r}→debit"
        else:
            # Try signed opening balance at field[4] (some packages use signed only)
            if side_raw and side_raw[0] in ("+", "-"):
                opening_balance = _parse_amount(side_raw)
                balance_source  = "tab_field[4] as signed amount"
            else:
                balance_source = "tab_fields[4,5]: amount=0"
    else:
        balance_source = f"tab_split has only {len(fields)} fields (<6)"

    return B100Fixed(
        account_number=account_number,
        account_name=account_name,
        opening_balance=opening_balance,
        raw_line=raw_line,
        debug_name_source=name_source,
        debug_balance_source=balance_source,
    )


# ---------------------------------------------------------------------------
# C100 parsing  –  exact regex logic
# ---------------------------------------------------------------------------

def _parse_c100(
    raw_line: str,
    lineno:   int,
    warnings: List[str],
) -> Optional[C100Fixed]:
    """
    Extract account_number as the LAST signed numeric token.
    Extract movement amount from tab-split field[10].
    Sign of the last token → debit (+) or credit (-).
    """
    # ── Account number: last signed numeric field ────────────────────────
    matches = re.findall(r'[+-]\d+', raw_line)
    if not matches:
        warnings.append(
            f"Line {lineno}: C100 – no signed numeric fields found – skipping"
        )
        return None

    last_match = matches[-1]
    try:
        last_val = int(last_match)
    except ValueError:
        warnings.append(
            f"Line {lineno}: C100 – cannot parse last signed match {last_match!r} – skipping"
        )
        return None

    account_number = abs(last_val)
    is_debit       = (last_val >= 0)   # + → debit, - → credit

    if account_number == 0:
        warnings.append(
            f"Line {lineno}: C100 – last signed match {last_match!r} gives account 0 – skipping"
        )
        return None

    # ── Amount from tab-split field[10] ─────────────────────────────────
    fields    = raw_line.split("\t")
    amount    = 0.0
    amount_raw = ""
    if len(fields) > 10:
        amount_raw = fields[10].strip()
        amount     = abs(_parse_amount(amount_raw))

    # If tab-split gives no amount, try the second-to-last signed number
    if amount == 0.0 and len(matches) >= 2:
        try:
            amount = abs(int(matches[-2]))
            amount_raw = f"fallback from matches[-2]={matches[-2]!r}"
        except ValueError:
            pass

    signed_movement = amount if is_debit else -amount

    return C100Fixed(
        account_number=account_number,
        movement=signed_movement,
        raw_line=raw_line,
        debug_last_match=last_match,
        debug_amount_raw=amount_raw,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_amount(raw: str) -> float:
    """Parse Israeli-format number: remove commas/dots used as thousands sep."""
    raw = raw.strip()
    if not raw:
        return 0.0
    cleaned = re.sub(r"[,\s]", "", raw)
    # Handle comma as decimal: "1234,56" → "1234.56"
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _norm(account_number: str) -> str:
    """Normalize account key for dict lookup: strip whitespace, upper-case."""
    return account_number.strip().upper()


# ---------------------------------------------------------------------------
# Lookup helpers (used by simple_balance.py)
# ---------------------------------------------------------------------------

def find_b100(
    b100_records: Dict[str, B100Fixed],
    account_number: str,
) -> Optional[B100Fixed]:
    """
    Look up a B100 record by account_number.

    Tries three variants:
      1. Exact normalised match
      2. Zero-stripped match  ("001234" → "1234")
      3. Mutual zero-strip
    """
    key = _norm(account_number)

    if key in b100_records:
        return b100_records[key]

    stripped = key.lstrip("0") or key
    if stripped in b100_records:
        return b100_records[stripped]

    for stored_key, rec in b100_records.items():
        if stored_key.lstrip("0") == stripped:
            return rec

    return None


def get_c100_for_account(
    c100_movements: List[C100Fixed],
    account_number: str,
) -> List[C100Fixed]:
    """
    Return all C100 movements whose account_number matches the given key.
    Handles zero-stripped variants.
    """
    try:
        target = int(account_number.strip())
    except ValueError:
        return []

    result = []
    for m in c100_movements:
        if m.account_number == target:
            result.append(m)

    return result
