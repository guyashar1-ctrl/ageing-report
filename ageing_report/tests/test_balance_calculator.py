"""
test_balance_calculator.py
==========================
Unit tests for the backward-accumulation algorithm in balance_calculator.py.

Run with:  pytest ageing_report/tests/ -v
"""

from __future__ import annotations

import sys
import os
# Allow imports from the project root when running tests directly
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from typing import List

import pytest

from config.constants import ZERO_THRESHOLD
from parsers.bkmv_parser import AccountCard, BkmvData, Movement
from processors.balance_calculator import (
    AccountResult,
    calculate_account_result,
)


# ---------------------------------------------------------------------------
# Fixtures / builders
# ---------------------------------------------------------------------------

def _make_movement(
    txn_date: date,
    amount: float,
    debit_account: str = "",
    credit_account: str = "",
    entry_number: str = "1",
    line_number: str = "1",
) -> Movement:
    """Build a minimal Movement for testing."""
    return Movement(
        entry_number=entry_number,
        line_number=line_number,
        txn_date=txn_date,
        value_date=txn_date,
        details="Test",
        reference1="",
        amount=abs(amount),
        debit_account=debit_account,
        credit_account=credit_account,
        debit_account_norm=debit_account.strip().upper(),
        credit_account_norm=credit_account.strip().upper(),
        movement_type="1",
    )


def _make_bkmv(
    account_key: str,
    opening: float,
    closing: float,
    movements: List[Movement],
) -> BkmvData:
    """Build a minimal BkmvData for testing."""
    card = AccountCard(
        account_key=account_key,
        account_key_norm=account_key.upper(),
        account_name="Test Account",
        sort_code="1342",
        opening_balance=opening,
        closing_balance=closing,
    )
    return BkmvData(
        file_header=None,
        accounts={account_key.upper(): card},
        movements=movements,
        movements_target_year=movements,
    )


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------

class TestZeroBalance:
    def test_zero_closing_balance_returns_empty_result(self):
        bkmv = _make_bkmv("10001", opening=0.0, closing=0.0, movements=[])
        result = calculate_account_result("10001", bkmv)
        assert result.selected_movements == []
        assert result.debt_start_date is None
        assert result.balance_side == ""


class TestDebitBalance:
    def test_simple_debit_exact_match(self):
        """Cumulative sum equals closing balance exactly → stops immediately."""
        mvmts = [
            _make_movement(date(2025, 1, 10), 1000, debit_account="10001", entry_number="1"),
            _make_movement(date(2025, 2, 15), 2000, debit_account="10001", entry_number="2"),
            _make_movement(date(2025, 3, 20), 3000, debit_account="10001", entry_number="3"),
        ]
        # Closing balance = 3000 (debit) → should select only the last movement
        bkmv = _make_bkmv("10001", opening=0.0, closing=3000.0, movements=mvmts)
        result = calculate_account_result("10001", bkmv)

        assert result.balance_side == "ח"
        assert len(result.selected_movements) == 1
        assert result.selected_movements[0].amount == 3000.0
        assert result.selected_movements[0].txn_date == date(2025, 3, 20)
        assert result.debt_start_date == date(2025, 3, 20)
        assert result.selected_sum == 3000.0

    def test_debit_multiple_movements_needed(self):
        """Need last two movements to exceed threshold."""
        mvmts = [
            _make_movement(date(2025, 1, 1), 500, debit_account="10001", entry_number="1"),
            _make_movement(date(2025, 2, 1), 600, debit_account="10001", entry_number="2"),
            _make_movement(date(2025, 3, 1), 700, debit_account="10001", entry_number="3"),
            _make_movement(date(2025, 4, 1), 800, debit_account="10001", entry_number="4"),
        ]
        # Closing balance 1400 → need last two: 800+700 = 1500 >= 1400
        bkmv = _make_bkmv("10001", opening=0.0, closing=1400.0, movements=mvmts)
        result = calculate_account_result("10001", bkmv)

        assert result.balance_side == "ח"
        assert len(result.selected_movements) == 2
        # Should be in chronological order: March then April
        assert result.selected_movements[0].txn_date == date(2025, 3, 1)
        assert result.selected_movements[1].txn_date == date(2025, 4, 1)
        assert result.debt_start_date == date(2025, 3, 1)
        assert result.selected_sum >= 1400.0

    def test_debit_only_debit_movements_selected(self):
        """Credit movements for the same account must be ignored."""
        mvmts = [
            _make_movement(date(2025, 1, 1), 3000, debit_account="10001", entry_number="1"),
            # Credit movement – must NOT be included in debit selection
            _make_movement(date(2025, 2, 1), 1000, credit_account="10001", entry_number="2"),
            _make_movement(date(2025, 3, 1), 2000, debit_account="10001", entry_number="3"),
        ]
        # Closing balance 2000 (debit) → last debit movement = 2000
        bkmv = _make_bkmv("10001", opening=0.0, closing=2000.0, movements=mvmts)
        result = calculate_account_result("10001", bkmv)

        assert result.balance_side == "ח"
        for sm in result.selected_movements:
            assert sm.side == "ח"
        assert len(result.selected_movements) == 1
        assert result.selected_movements[0].amount == 2000.0

    def test_debit_all_movements_insufficient(self):
        """When all movements sum to less than closing balance, use all and flag."""
        mvmts = [
            _make_movement(date(2025, 1, 1), 100, debit_account="10001", entry_number="1"),
            _make_movement(date(2025, 2, 1), 200, debit_account="10001", entry_number="2"),
        ]
        # Closing balance 9999 → way above total 300
        bkmv = _make_bkmv("10001", opening=0.0, closing=9999.0, movements=mvmts)
        result = calculate_account_result("10001", bkmv)

        assert result.threshold_not_reached is True
        assert len(result.selected_movements) == 2  # all used
        assert result.selected_sum == 300.0


class TestCreditBalance:
    def test_simple_credit_balance(self):
        """Closing balance is credit (negative) → select credit movements."""
        mvmts = [
            _make_movement(date(2025, 1, 1), 500, credit_account="20001", entry_number="1"),
            _make_movement(date(2025, 2, 1), 500, credit_account="20001", entry_number="2"),
            _make_movement(date(2025, 3, 1), 400, credit_account="20001", entry_number="3"),
        ]
        # Closing balance -500 (credit) → last credit movement = 400... need more
        bkmv = _make_bkmv("20001", opening=0.0, closing=-500.0, movements=mvmts)
        result = calculate_account_result("20001", bkmv)

        assert result.balance_side == "ז"
        for sm in result.selected_movements:
            assert sm.side == "ז"
        # 400 < 500, so go back: 400+500 = 900 >= 500
        assert len(result.selected_movements) == 2
        assert result.debt_start_date == date(2025, 2, 1)

    def test_credit_only_credit_movements_selected(self):
        """Debit movements must be ignored when balance is credit."""
        mvmts = [
            _make_movement(date(2025, 1, 1), 1000, credit_account="20001", entry_number="1"),
            # This is a debit for account 20001 – must be ignored
            _make_movement(date(2025, 2, 1), 500, debit_account="20001", entry_number="2"),
            _make_movement(date(2025, 3, 1), 2000, credit_account="20001", entry_number="3"),
        ]
        bkmv = _make_bkmv("20001", opening=0.0, closing=-2000.0, movements=mvmts)
        result = calculate_account_result("20001", bkmv)

        assert result.balance_side == "ז"
        for sm in result.selected_movements:
            assert sm.side == "ז"
        assert len(result.selected_movements) == 1
        assert result.selected_movements[0].amount == 2000.0


class TestStableOrdering:
    def test_same_date_stable_sort(self):
        """Multiple movements on same date → stable order by entry_number."""
        mvmts = [
            _make_movement(date(2025, 6, 1), 300, debit_account="30001",
                          entry_number="3", line_number="1"),
            _make_movement(date(2025, 6, 1), 200, debit_account="30001",
                          entry_number="1", line_number="1"),
            _make_movement(date(2025, 6, 1), 100, debit_account="30001",
                          entry_number="2", line_number="1"),
        ]
        # Closing balance 300 → should pick entry_number=3 (last in stable sort)
        bkmv = _make_bkmv("30001", opening=0.0, closing=300.0, movements=mvmts)
        result = calculate_account_result("30001", bkmv)

        assert len(result.selected_movements) == 1
        assert result.selected_movements[0].entry_number == "3"
        assert result.selected_movements[0].amount == 300.0


class TestEdgeCases:
    def test_no_movements_returns_no_selection(self):
        """No movements at all → empty selection, warning issued."""
        bkmv = _make_bkmv("99999", opening=0.0, closing=5000.0, movements=[])
        result = calculate_account_result("99999", bkmv)

        assert result.selected_movements == []
        assert result.debt_start_date is None
        assert len(result.warnings) > 0

    def test_account_not_in_bkmvdata(self):
        """Account key missing from BKMVDATA → graceful result with warning."""
        bkmv = _make_bkmv("OTHER", opening=0.0, closing=0.0, movements=[])
        result = calculate_account_result("NOTEXIST", bkmv)

        assert len(result.warnings) > 0
        assert result.opening_balance == 0.0
        assert result.closing_balance == 0.0

    def test_single_movement_exactly_at_threshold(self):
        """One movement whose amount equals closing balance exactly."""
        mvmts = [
            _make_movement(date(2025, 12, 31), 7777.77, debit_account="40001",
                          entry_number="1"),
        ]
        bkmv = _make_bkmv("40001", opening=0.0, closing=7777.77, movements=mvmts)
        result = calculate_account_result("40001", bkmv)

        assert len(result.selected_movements) == 1
        assert abs(result.selected_sum - 7777.77) < ZERO_THRESHOLD

    def test_cumulative_stops_as_soon_as_threshold_met(self):
        """
        Verify that accumulation stops IMMEDIATELY when threshold is reached,
        not after one extra iteration.
        """
        mvmts = [
            _make_movement(date(2025, 1, 1), 100, debit_account="50001", entry_number="1"),
            _make_movement(date(2025, 2, 1), 100, debit_account="50001", entry_number="2"),
            _make_movement(date(2025, 3, 1), 100, debit_account="50001", entry_number="3"),
            _make_movement(date(2025, 4, 1), 200, debit_account="50001", entry_number="4"),
            _make_movement(date(2025, 5, 1), 50,  debit_account="50001", entry_number="5"),
        ]
        # Threshold = 200 → last movement (entry 5) = 50, not enough.
        # Go back: 50 + 200 = 250 >= 200 → stop.  Selected: entry 4 and 5.
        bkmv = _make_bkmv("50001", opening=0.0, closing=200.0, movements=mvmts)
        result = calculate_account_result("50001", bkmv)

        assert len(result.selected_movements) == 2
        selected_entry_numbers = {sm.entry_number for sm in result.selected_movements}
        assert "4" in selected_entry_numbers
        assert "5" in selected_entry_numbers
        assert "3" not in selected_entry_numbers   # Must NOT include entry 3
