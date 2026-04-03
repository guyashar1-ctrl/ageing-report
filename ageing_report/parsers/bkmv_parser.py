"""
bkmv_parser.py
==============
Parser for the Israeli "Uniform Structure" file (BKMVDATA.TXT / TXT.BKMVDATA).

Public API
----------
parse_bkmvdata(path: str) -> BkmvData
    Reads the file and returns a BkmvData dataclass containing:
        - file_header  : A100FileHeader (company / period info)
        - accounts     : dict[str, AccountCard]  keyed by normalised account_key
        - movements    : list[Movement]
        - warnings     : list[str]

Design notes
------------
- Field indices are read from config/field_mappings.py — do NOT hardcode them here.
- All balance amounts are stored as signed floats:
      positive  = debit  (חובה)
      negative  = credit (זכות)
- Movement amounts are stored as signed floats:
      positive  = debit movement  (the movement is a debit to debit_account)
      negative  = credit movement (from the perspective of credit_account)
  BUT the raw C100 amount field is always positive; the sign is derived by
  context (which account we are looking at) in the balance calculator.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

from config.constants import (
    BALANCE_SIDE_DEBIT,
    BALANCE_SIDE_CREDIT,
    BKMV_DELIMITER,
    B100_MIN_FIELDS,
    C100_MIN_FIELDS,
    DATE_FORMAT_BKMV,
    MOVEMENT_TYPE_OPENING,
    RECORD_TYPE_ACCOUNT,
    RECORD_TYPE_FILE_HEADER,
    RECORD_TYPE_MOVEMENT,
    TARGET_YEAR,
    ZERO_THRESHOLD,
)
from config.field_mappings import (
    A100_FIELDS,
    B100_FIELDS,
    B100_FIELDS_SIGNED,
    C100_FIELDS,
    USE_SIGNED_BALANCE_IN_B100,
)
from utils.file_utils import normalize_account_key, read_text_file
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class A100FileHeader:
    """Metadata from the A100 record."""
    tax_id:       str = ""
    company_name: str = ""
    period_start: Optional[date] = None
    period_end:   Optional[date] = None
    b100_count:   int = 0
    c100_count:   int = 0
    raw_line:     str = ""


@dataclass
class AccountCard:
    """
    Represents one B100 account-card record.

    Balances are signed floats: positive = debit, negative = credit.
    """
    account_key:      str            # As stored in the file (original case/spacing stripped)
    account_key_norm: str            # Normalised for matching (upper, stripped)
    account_name:     str
    sort_code:        str            # Chart-of-accounts classification code

    # Signed balances (positive = debit, negative = credit)
    opening_balance:  float = 0.0
    closing_balance:  float = 0.0

    # Foreign currency (informational only; not used in the core calculation)
    currency_code:    str = ""
    opening_balance_fc: float = 0.0
    closing_balance_fc: float = 0.0

    # Raw line for debugging
    raw_line: str = ""


@dataclass
class Movement:
    """
    Represents one C100 journal-movement record.

    IMPORTANT – the raw C100 amount is always POSITIVE.  Whether this movement
    is a debit or credit for a specific account is determined by whether the
    account appears in debit_account or credit_account.

    To query: "Is this a debit movement for account X?"
        → (movement.debit_account_norm == normalize_account_key(X))
    To query: "Is this a credit movement for account X?"
        → (movement.credit_account_norm == normalize_account_key(X))
    """
    entry_number:        str
    line_number:         str
    txn_date:            Optional[date]
    value_date:          Optional[date]
    details:             str
    reference1:          str
    amount:              float            # Always positive (ILS)
    debit_account:       str             # Original key
    credit_account:      str             # Original key
    debit_account_norm:  str             # Normalised
    credit_account_norm: str             # Normalised
    movement_type:       str             # "1" = regular, "2" = opening
    currency_code:       str = ""
    foreign_amount:      float = 0.0
    raw_line:            str = ""


@dataclass
class BkmvData:
    """Complete parsed content of a BKMVDATA.TXT file."""
    file_header: Optional[A100FileHeader]
    accounts:    Dict[str, AccountCard]   # key = normalised account_key
    movements:   List[Movement]
    warnings:    List[str] = field(default_factory=list)

    # Convenience: all movements for year TARGET_YEAR (populated after parsing)
    movements_target_year: List[Movement] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

def parse_bkmvdata(path: str) -> BkmvData:
    """
    Read and parse a BKMVDATA.TXT file.

    Returns a BkmvData instance.  Never raises – errors are recorded as
    warnings in the returned object and logged.
    """
    log.info("Parsing BKMVDATA: %s", path)
    content, encoding = read_text_file(path)
    log.info("File encoding detected: %s", encoding)

    file_header: Optional[A100FileHeader] = None
    accounts:    Dict[str, AccountCard] = {}
    movements:   List[Movement] = []
    warnings:    List[str] = []

    total_lines = 0
    skipped_lines = 0

    for lineno, raw_line in enumerate(content.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        total_lines += 1
        fields = line.split(BKMV_DELIMITER)
        record_type = fields[0].strip().upper() if fields else ""

        if record_type == RECORD_TYPE_FILE_HEADER:
            file_header = _parse_a100(fields, raw_line, warnings)

        elif record_type == RECORD_TYPE_ACCOUNT:
            card = _parse_b100(fields, raw_line, warnings, lineno)
            if card:
                if card.account_key_norm in accounts:
                    w = (f"Line {lineno}: duplicate account key '{card.account_key}' – "
                         f"overwriting previous entry")
                    warnings.append(w)
                    log.warning(w)
                accounts[card.account_key_norm] = card

        elif record_type == RECORD_TYPE_MOVEMENT:
            mvmt = _parse_c100(fields, raw_line, warnings, lineno)
            if mvmt:
                movements.append(mvmt)

        else:
            # D110, M100, or unknown – skip silently
            skipped_lines += 1

    log.info("Parsed %d accounts, %d movements (%d lines skipped)",
             len(accounts), len(movements), skipped_lines)

    if file_header:
        log.info("Period: %s – %s | Company: %s",
                 file_header.period_start, file_header.period_end,
                 file_header.company_name)
        if file_header.b100_count and file_header.b100_count != len(accounts):
            w = (f"A100 declares {file_header.b100_count} B100 records but "
                 f"{len(accounts)} were parsed")
            warnings.append(w)
            log.warning(w)
        if file_header.c100_count and file_header.c100_count != len(movements):
            w = (f"A100 declares {file_header.c100_count} C100 records but "
                 f"{len(movements)} were parsed")
            warnings.append(w)
            log.warning(w)

    # Filter movements to TARGET_YEAR (exclude opening-balance movement type)
    target_mvmts = [
        m for m in movements
        if m.txn_date and m.txn_date.year == TARGET_YEAR
        and m.movement_type != MOVEMENT_TYPE_OPENING
    ]
    log.info("Movements in year %d (excluding opening type): %d",
             TARGET_YEAR, len(target_mvmts))

    return BkmvData(
        file_header=file_header,
        accounts=accounts,
        movements=movements,
        warnings=warnings,
        movements_target_year=target_mvmts,
    )


# ---------------------------------------------------------------------------
# A100 parsing
# ---------------------------------------------------------------------------

def _parse_a100(
    fields: List[str],
    raw_line: str,
    warnings: List[str],
) -> A100FileHeader:
    f = A100_FIELDS
    header = A100FileHeader(raw_line=raw_line)
    try:
        header.tax_id       = _get_field(fields, f["tax_id"]).strip()
        header.company_name = _get_field(fields, f["company_name"]).strip()
        header.period_start = _parse_date(
            _get_field(fields, f["period_start"]), "A100.period_start", warnings)
        header.period_end   = _parse_date(
            _get_field(fields, f["period_end"]), "A100.period_end", warnings)
        b100_raw = _get_field(fields, f["b100_count"])
        c100_raw = _get_field(fields, f["c100_count"])
        header.b100_count = int(b100_raw) if b100_raw.strip().isdigit() else 0
        header.c100_count = int(c100_raw) if c100_raw.strip().isdigit() else 0
    except Exception as exc:
        w = f"A100 parse error: {exc}  |  line: {raw_line[:120]}"
        warnings.append(w)
        log.warning(w)
    return header


# ---------------------------------------------------------------------------
# B100 parsing
# ---------------------------------------------------------------------------

def _parse_b100(
    fields: List[str],
    raw_line: str,
    warnings: List[str],
    lineno: int,
) -> Optional[AccountCard]:
    mapping = B100_FIELDS_SIGNED if USE_SIGNED_BALANCE_IN_B100 else B100_FIELDS

    if len(fields) < B100_MIN_FIELDS:
        w = (f"Line {lineno}: B100 has only {len(fields)} fields "
             f"(expected >= {B100_MIN_FIELDS}) – skipping.  "
             f"Check field_mappings.py B100_MIN_FIELDS.")
        warnings.append(w)
        log.warning(w)
        return None

    try:
        account_key  = _get_field(fields, mapping["account_key"]).strip()
        account_name = _get_field(fields, mapping["account_name"]).strip()
        sort_code    = _get_field(fields, mapping["sort_code"]).strip()

        if USE_SIGNED_BALANCE_IN_B100:
            opening_balance = _parse_signed_amount(
                _get_field(fields, mapping["opening_balance"]))
            closing_balance = _parse_signed_amount(
                _get_field(fields, mapping["closing_balance"]))
        else:
            opening_balance = _parse_balance(
                _get_field(fields, mapping["opening_balance_side"]),
                _get_field(fields, mapping["opening_balance"]),
                f"B100[{account_key}].opening",
                warnings,
            )
            closing_balance = _parse_balance(
                _get_field(fields, mapping["closing_balance_side"]),
                _get_field(fields, mapping["closing_balance"]),
                f"B100[{account_key}].closing",
                warnings,
            )

        currency_code    = _get_field(fields, mapping.get("currency_code", 99))
        opening_fc       = _parse_amount(_get_field(fields, mapping.get("opening_balance_fc", 99)))
        closing_fc       = _parse_amount(_get_field(fields, mapping.get("closing_balance_fc", 99)))

        if not account_key:
            w = f"Line {lineno}: B100 has empty account_key – skipping"
            warnings.append(w)
            log.warning(w)
            return None

        return AccountCard(
            account_key=account_key,
            account_key_norm=normalize_account_key(account_key),
            account_name=account_name,
            sort_code=sort_code,
            opening_balance=opening_balance,
            closing_balance=closing_balance,
            currency_code=currency_code,
            opening_balance_fc=opening_fc,
            closing_balance_fc=closing_fc,
            raw_line=raw_line,
        )

    except Exception as exc:
        w = (f"Line {lineno}: B100 parse error: {exc}  |  "
             f"line: {raw_line[:120]}")
        warnings.append(w)
        log.warning(w)
        return None


# ---------------------------------------------------------------------------
# C100 parsing
# ---------------------------------------------------------------------------

def _parse_c100(
    fields: List[str],
    raw_line: str,
    warnings: List[str],
    lineno: int,
) -> Optional[Movement]:
    f = C100_FIELDS

    if len(fields) < C100_MIN_FIELDS:
        w = (f"Line {lineno}: C100 has only {len(fields)} fields "
             f"(expected >= {C100_MIN_FIELDS}) – skipping.  "
             f"Check field_mappings.py C100_MIN_FIELDS.")
        warnings.append(w)
        log.warning(w)
        return None

    try:
        movement_type  = _get_field(fields, f["movement_type"]).strip()
        entry_number   = _get_field(fields, f["entry_number"]).strip()
        line_number    = _get_field(fields, f["line_number"]).strip()
        txn_date       = _parse_date(_get_field(fields, f["date"]),
                                     f"C100[{entry_number}].date", warnings)
        value_date     = _parse_date(_get_field(fields, f["value_date"]),
                                     f"C100[{entry_number}].value_date", warnings)
        details        = _get_field(fields, f["details"]).strip()
        reference1     = _get_field(fields, f["reference1"]).strip()
        amount         = _parse_amount(_get_field(fields, f["amount"]))
        debit_account  = _get_field(fields, f["debit_account"]).strip()
        credit_account = _get_field(fields, f["credit_account"]).strip()
        currency_code  = _get_field(fields, f.get("currency_code", 99))
        foreign_amount = _parse_amount(_get_field(fields, f.get("foreign_amount", 99)))

        return Movement(
            entry_number=entry_number,
            line_number=line_number,
            txn_date=txn_date,
            value_date=value_date,
            details=details,
            reference1=reference1,
            amount=abs(amount),   # always positive; sign determined by account role
            debit_account=debit_account,
            credit_account=credit_account,
            debit_account_norm=normalize_account_key(debit_account),
            credit_account_norm=normalize_account_key(credit_account),
            movement_type=movement_type,
            currency_code=currency_code,
            foreign_amount=foreign_amount,
            raw_line=raw_line,
        )

    except Exception as exc:
        w = (f"Line {lineno}: C100 parse error: {exc}  |  "
             f"line: {raw_line[:120]}")
        warnings.append(w)
        log.warning(w)
        return None


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _get_field(fields: List[str], index: int) -> str:
    """Safely get a field by index; return empty string if out of range."""
    if index >= len(fields):
        return ""
    return fields[index]


def _parse_date(
    raw: str,
    context: str,
    warnings: List[str],
) -> Optional[date]:
    """Parse YYYYMMDD date string.  Returns None (with warning) on failure."""
    raw = raw.strip()
    if not raw or raw == "0" * len(raw):
        return None
    try:
        return datetime.strptime(raw[:8], DATE_FORMAT_BKMV).date()
    except ValueError:
        w = f"Could not parse date '{raw}' in {context}"
        warnings.append(w)
        log.debug(w)
        return None


def _parse_amount(raw: str) -> float:
    """
    Parse a numeric amount string to float.
    Handles:  "1,234.56"  "1234.56"  "-1234"  ""  etc.
    """
    raw = raw.strip()
    if not raw:
        return 0.0
    # Remove thousand separators (comma in Israeli format)
    cleaned = re.sub(r"[,\s]", "", raw)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _parse_signed_amount(raw: str) -> float:
    """Parse a signed amount where positive=debit, negative=credit."""
    return _parse_amount(raw)


def _parse_balance(
    side_raw: str,
    amount_raw: str,
    context: str,
    warnings: List[str],
) -> float:
    """
    Convert a (side_code, absolute_amount) pair to a signed float.
    side_code: "1" = debit (positive), "2" = credit (negative).
    """
    side   = side_raw.strip()
    amount = abs(_parse_amount(amount_raw))

    if side == BALANCE_SIDE_DEBIT:
        return amount
    elif side == BALANCE_SIDE_CREDIT:
        return -amount
    elif side == "" and amount == 0.0:
        return 0.0
    else:
        # Some software leaves side blank when amount is positive = debit
        w = (f"Unknown balance side '{side}' in {context} – "
             f"assuming debit (positive).  Adjust BALANCE_SIDE_* in constants.py if wrong.")
        warnings.append(w)
        log.debug(w)
        return amount   # assume debit


# ---------------------------------------------------------------------------
# Account lookup utilities (used by processors)
# ---------------------------------------------------------------------------

def find_account(
    bkmv: BkmvData,
    account_key: str,
) -> Optional[AccountCard]:
    """
    Look up an account by key, using:
    1. Exact normalised match
    2. Zero-stripped match  (e.g. "001001" → "1001")
    Returns None if not found.
    """
    norm = normalize_account_key(account_key)

    # 1. Exact match
    if norm in bkmv.accounts:
        return bkmv.accounts[norm]

    # 2. Zero-stripped match
    stripped = norm.lstrip("0") or norm
    if stripped in bkmv.accounts:
        log.debug("Account '%s' matched via zero-strip to '%s'", account_key, stripped)
        return bkmv.accounts[stripped]

    # 3. Reverse: strip the stored key's zeros and compare
    for stored_norm, card in bkmv.accounts.items():
        if stored_norm.lstrip("0") == norm.lstrip("0"):
            log.debug("Account '%s' matched via mutual zero-strip to '%s'",
                      account_key, card.account_key)
            return card

    log.debug("Account key '%s' not found in BKMVDATA", account_key)
    return None


def get_movements_for_account(
    bkmv: BkmvData,
    account_key: str,
    year_only: bool = True,
) -> List[Movement]:
    """
    Return all movements where the given account appears on either side.
    If year_only=True, restrict to TARGET_YEAR movements.
    """
    norm = normalize_account_key(account_key)
    # Also allow zero-stripped variation
    norm_stripped = norm.lstrip("0") or norm

    pool = bkmv.movements_target_year if year_only else bkmv.movements

    result = []
    for m in pool:
        d_norm = m.debit_account_norm
        c_norm = m.credit_account_norm
        d_stripped = d_norm.lstrip("0") or d_norm
        c_stripped = c_norm.lstrip("0") or c_norm

        if norm in (d_norm, c_norm) or norm_stripped in (d_stripped, c_stripped):
            result.append(m)

    return result
