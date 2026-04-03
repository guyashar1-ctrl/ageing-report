"""
simple_balance.py
=================
Simple closing-balance calculator using the fixed parser output.

Formula
-------
    Closing Balance = Opening Balance + Movement Sum

where:
    Movement Sum = sum of signed C100 movements for the account
    positive movement  = debit
    negative movement  = credit

Public API
----------
calculate_simple_results(
    pdf_account_numbers : list[str],       # from 331 PDF
    b100_records        : dict[str, B100Fixed],
    c100_movements      : list[C100Fixed],
) -> list[SimpleAccountResult]
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from parsers.bkmv_fixed_parser import (
    B100Fixed,
    C100Fixed,
    find_b100,
    get_c100_for_account,
)
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Result data class
# ---------------------------------------------------------------------------

@dataclass
class SimpleAccountResult:
    pdf_account_number: str       # exactly as it appeared in the 331 PDF
    b100_account_number: str      # as extracted from B100 pos 172:178 (or "" if not found)
    account_name: str
    opening_balance: float        # from B100 (signed: +debit / -credit)
    movement_sum: float           # sum of signed C100 movements
    closing_balance: float        # = opening_balance + movement_sum
    matched: bool                 # True if B100 record was found
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Calculator
# ---------------------------------------------------------------------------

def calculate_simple_results(
    pdf_account_numbers: List[str],
    b100_records: Dict[str, "B100Fixed"],
    c100_movements: List["C100Fixed"],
) -> List[SimpleAccountResult]:
    """
    For each account number from the 331 PDF:
      1. Match to a B100 record (get name + opening balance).
      2. Collect all matching C100 movements.
      3. Calculate closing_balance = opening_balance + sum(movements).
    """
    results: List[SimpleAccountResult] = []

    for pdf_acct in pdf_account_numbers:
        log.info("Processing PDF account: %s", pdf_acct)

        # ── Match B100 ──────────────────────────────────────────────────
        b100 = find_b100(b100_records, pdf_acct)

        if b100 is None:
            w = (f"Account '{pdf_acct}' from 331 PDF not found in B100 records. "
                 f"Opening balance = 0.")
            log.warning(w)
            results.append(SimpleAccountResult(
                pdf_account_number=pdf_acct,
                b100_account_number="",
                account_name=f"[לא נמצא: {pdf_acct}]",
                opening_balance=0.0,
                movement_sum=0.0,
                closing_balance=0.0,
                matched=False,
                warnings=[w],
            ))
            continue

        opening_balance = b100.opening_balance
        account_name    = b100.account_name or f"[ללא שם: {pdf_acct}]"

        # ── Collect C100 movements ───────────────────────────────────────
        movements = get_c100_for_account(c100_movements, pdf_acct)
        movement_sum = sum(m.movement for m in movements)

        closing_balance = opening_balance + movement_sum

        log.info(
            "  %s → B100='%s' name='%s' ob=%.2f  "
            "c100_count=%d  movement_sum=%.2f  closing=%.2f",
            pdf_acct, b100.account_number, account_name,
            opening_balance, len(movements), movement_sum, closing_balance,
        )

        results.append(SimpleAccountResult(
            pdf_account_number=pdf_acct,
            b100_account_number=b100.account_number,
            account_name=account_name,
            opening_balance=opening_balance,
            movement_sum=movement_sum,
            closing_balance=closing_balance,
            matched=True,
        ))

    return results
