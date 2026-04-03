"""
balance_calculator.py
=====================
Core calculation logic for the "Debt/Credit Balance Summation" column.

This is the most critical module in the system.  Every design decision is
documented inline.

Public API
----------
calculate_account_result(
    account_key : str,
    bkmv        : BkmvData,
    account_line: AccountLine | None,
) -> AccountResult

The returned AccountResult contains everything needed to generate one row in
the Excel output, including the list of selected movements and the start date.

Algorithm (verbatim from the specification)
-------------------------------------------
1.  Determine the sign of closing_balance:
        > 0  → work with DEBIT movements only
        < 0  → work with CREDIT movements only
        = 0  → no accumulation; leave summation/date empty

2.  Load all C100 movements for this account for year TARGET_YEAR.
    Sort oldest → newest, with stable tie-breaking by
    (entry_number, line_number).

3.  Filter to the relevant side only.

4.  Accumulate backward (newest → oldest):
        cumulative_sum += movement.amount
        stop as soon as abs(cumulative_sum) >= abs(closing_balance)

5.  Selected movements = the "tail" of the sorted list that was accumulated.
    They are stored in chronological order (oldest selected first) for
    writing to the Excel helper sheet.

6.  debt_start_date = date of the OLDEST selected movement
    (= the movement that was the LAST added in the backward iteration).

Edge cases handled
------------------
- Balance is exactly zero → skip.
- No movements found at all → warning, empty result.
- Movements found but total < closing_balance threshold →
    use all available movements, flag in warnings.
- Multiple movements on the same date → stable sort by (entry_number, line_number).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import List, Optional

from config.constants import (
    DEBIT_LABEL,
    CREDIT_LABEL,
    TARGET_YEAR,
    ZERO_THRESHOLD,
)
from parsers.bkmv_parser import (
    AccountCard,
    BkmvData,
    Movement,
    find_account,
    get_movements_for_account,
)
from parsers.report331_parser import AccountLine
from utils.file_utils import normalize_account_key
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Result data model
# ---------------------------------------------------------------------------

@dataclass
class SelectedMovement:
    """A single movement chosen by the backward-accumulation algorithm."""
    txn_date:       Optional[date]
    entry_number:   str
    line_number:    str
    details:        str
    reference1:     str
    amount:         float       # Always positive (absolute value)
    side:           str         # "ח" or "ז" – the side relevant to this account


@dataclass
class AccountResult:
    """Full result for one account row in the Excel output."""

    # Identity (prefer 331 report values; fall back to BKMVDATA)
    account_number:   str
    account_name:     str

    # Balances (signed: positive=debit, negative=credit)
    opening_balance:  float
    closing_balance:  float

    # The selected movements in chronological order (for SUM formula)
    selected_movements: List[SelectedMovement] = field(default_factory=list)

    # Sum of selected movements (positive absolute value)
    selected_sum:     float = 0.0

    # Date of the oldest selected movement (= "debt start date")
    debt_start_date:  Optional[date] = None

    # The balance side that was accumulated ("ח" or "ז")
    balance_side:     str = ""

    # Warnings for this specific account
    warnings:         List[str] = field(default_factory=list)

    # Whether all available movements were used without reaching the threshold
    threshold_not_reached: bool = False


# ---------------------------------------------------------------------------
# Main calculator
# ---------------------------------------------------------------------------

def calculate_account_result(
    account_key:  str,
    bkmv:         BkmvData,
    account_line: Optional[AccountLine] = None,
) -> AccountResult:
    """
    Compute the AccountResult for a single account.

    Parameters
    ----------
    account_key  : Account number (as it appears in the 331 report or BKMVDATA).
    bkmv         : Parsed BKMVDATA content.
    account_line : If provided, prefer its account_name and use its balance
                   only as a cross-check.  BKMVDATA closing_balance is used
                   for the algorithm (it is more precise).
    """
    # ------------------------------------------------------------------
    # Step A: Resolve account card from BKMVDATA
    # ------------------------------------------------------------------
    card: Optional[AccountCard] = find_account(bkmv, account_key)

    display_number = account_key
    display_name   = account_line.account_name if account_line else ""

    if card is None:
        w = (f"Account '{account_key}' not found in BKMVDATA.TXT. "
             f"Opening/closing balances will be 0.  Movements cannot be retrieved.")
        log.warning(w)
        return AccountResult(
            account_number=display_number,
            account_name=display_name or f"[לא נמצא: {account_key}]",
            opening_balance=0.0,
            closing_balance=0.0,
            warnings=[w],
        )

    # Prefer 331 display name; fall back to BKMVDATA name
    if not display_name:
        display_name = card.account_name

    opening_balance = card.opening_balance
    closing_balance = card.closing_balance

    result = AccountResult(
        account_number=display_number,
        account_name=display_name,
        opening_balance=opening_balance,
        closing_balance=closing_balance,
    )

    # ------------------------------------------------------------------
    # Step 1: Determine balance side
    # ------------------------------------------------------------------
    if abs(closing_balance) < ZERO_THRESHOLD:
        log.debug("Account %s: closing balance is zero – skipping accumulation",
                  account_key)
        result.balance_side = ""
        return result   # No accumulation needed

    balance_is_debit = closing_balance > 0
    result.balance_side = DEBIT_LABEL if balance_is_debit else CREDIT_LABEL

    # ------------------------------------------------------------------
    # Step 2: Load movements for this account
    # ------------------------------------------------------------------
    all_movements = get_movements_for_account(bkmv, account_key, year_only=True)
    log.debug("Account %s: %d movements in year %d",
              account_key, len(all_movements), TARGET_YEAR)

    if not all_movements:
        w = (f"Account '{account_key}': no movements found in year {TARGET_YEAR}. "
             f"Closing balance = {closing_balance:.2f}.  "
             f"Summation formula will be empty.")
        result.warnings.append(w)
        log.warning(w)
        return result

    # ------------------------------------------------------------------
    # Step 3: Sort oldest → newest; stable tie-break by (entry_number, line_number)
    # ------------------------------------------------------------------
    def _sort_key(m: Movement):
        dt = m.txn_date or date(1900, 1, 1)
        try:
            en = int(m.entry_number) if m.entry_number.isdigit() else 0
        except (ValueError, AttributeError):
            en = 0
        try:
            ln = int(m.line_number) if m.line_number.isdigit() else 0
        except (ValueError, AttributeError):
            ln = 0
        return (dt, en, ln)

    sorted_movements = sorted(all_movements, key=_sort_key)

    # ------------------------------------------------------------------
    # Step 4: Filter to relevant side (debit movements vs credit movements)
    # ------------------------------------------------------------------
    account_norm = normalize_account_key(account_key)
    # Also accept zero-stripped variation
    account_norm_stripped = account_norm.lstrip("0") or account_norm

    def _is_debit_for_account(m: Movement) -> bool:
        dn = m.debit_account_norm
        return (dn == account_norm or
                dn == account_norm_stripped or
                dn.lstrip("0") == account_norm_stripped)

    def _is_credit_for_account(m: Movement) -> bool:
        cn = m.credit_account_norm
        return (cn == account_norm or
                cn == account_norm_stripped or
                cn.lstrip("0") == account_norm_stripped)

    if balance_is_debit:
        relevant = [m for m in sorted_movements if _is_debit_for_account(m)]
        side_label = DEBIT_LABEL
    else:
        relevant = [m for m in sorted_movements if _is_credit_for_account(m)]
        side_label = CREDIT_LABEL

    log.debug("Account %s: %d relevant (%s) movements after filtering",
              account_key, len(relevant), side_label)

    if not relevant:
        w = (f"Account '{account_key}': no {side_label} movements found "
             f"(closing balance = {closing_balance:.2f} {'ח' if balance_is_debit else 'ז'}).  "
             f"This is unexpected.  Check movement data.")
        result.warnings.append(w)
        log.warning(w)
        return result

    # ------------------------------------------------------------------
    # Step 5: Backward accumulation
    # ------------------------------------------------------------------
    threshold = abs(closing_balance)
    cumulative = 0.0
    selected_indices: List[int] = []   # indices into *relevant* list (backward)

    for i in range(len(relevant) - 1, -1, -1):
        m = relevant[i]
        cumulative += m.amount
        selected_indices.append(i)
        if cumulative >= threshold - ZERO_THRESHOLD:
            # Threshold reached or exceeded → stop
            break
    else:
        # Loop completed without reaching threshold
        w = (f"Account '{account_key}': cumulative sum of all {len(relevant)} "
             f"relevant movements = {cumulative:.2f} is less than "
             f"closing balance {threshold:.2f}.  "
             f"All available movements will be included in the formula.  "
             f"This may indicate missing data in BKMVDATA.")
        result.warnings.append(w)
        log.warning(w)
        result.threshold_not_reached = True

    # ------------------------------------------------------------------
    # Step 6: Build selected_movements in chronological order
    # ------------------------------------------------------------------
    # selected_indices are in reverse order (newest first); reverse them
    selected_indices_chrono = list(reversed(selected_indices))

    selected: List[SelectedMovement] = []
    for i in selected_indices_chrono:
        m = relevant[i]
        selected.append(SelectedMovement(
            txn_date=m.txn_date,
            entry_number=m.entry_number,
            line_number=m.line_number,
            details=m.details,
            reference1=m.reference1,
            amount=m.amount,
            side=side_label,
        ))

    result.selected_movements = selected
    result.selected_sum = sum(s.amount for s in selected)

    # Debt start date = date of the OLDEST selected movement
    if selected and selected[0].txn_date:
        result.debt_start_date = selected[0].txn_date

    log.debug(
        "Account %s: selected %d movements, sum=%.2f, threshold=%.2f, start=%s",
        account_key,
        len(selected),
        result.selected_sum,
        threshold,
        result.debt_start_date,
    )

    return result


# ---------------------------------------------------------------------------
# Batch processor
# ---------------------------------------------------------------------------

def process_all_accounts(
    account_lines: List[AccountLine],
    bkmv: BkmvData,
) -> List[AccountResult]:
    """
    Process every account from the 331 report and return a list of results.
    """
    results: List[AccountResult] = []
    for al in account_lines:
        log.info("Processing account: %s  (%s)", al.account_number, al.account_name)
        r = calculate_account_result(
            account_key=al.account_number,
            bkmv=bkmv,
            account_line=al,
        )
        results.append(r)
    return results
