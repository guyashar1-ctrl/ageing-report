"""
test_aging_calculator.py
========================
Unit tests for the backward-accumulation algorithm in aging_calculator.py.

Run with:  pytest ageing_report/tests/ -v
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest

from ageing_report.config.constants import ZERO_THRESHOLD
from ageing_report.processors.aging_calculator import calculate_aging, process_accounts


# ---------------------------------------------------------------------------
# calculate_aging tests
# ---------------------------------------------------------------------------

class TestZeroBalance:
    def test_zero_closing_returns_empty(self):
        result = calculate_aging(0.0, [], 0.0)
        assert result == ("", [])

    def test_near_zero_closing_returns_empty(self):
        result = calculate_aging(0.003, [], 0.0)
        assert result == ("", [])


class TestDebitBalance:
    def test_simple_debit_exact_match(self):
        """Single transaction covers closing balance exactly."""
        txns = [("20250320", "1", 3000.0)]
        date_str, parts = calculate_aging(3000.0, txns, 0.0)
        assert date_str == "20/03/2025"
        assert parts == [3000.0]

    def test_debit_multiple_transactions_needed(self):
        """Need last two transactions to cover closing balance."""
        txns = [
            ("20250101", "1", 500.0),
            ("20250201", "1", 600.0),
            ("20250301", "1", 700.0),
            ("20250401", "1", 800.0),
        ]
        # closing=1400, last=800 not enough, 800+700=1500 >= 1400
        date_str, parts = calculate_aging(1400.0, txns, 0.0)
        assert date_str == "01/03/2025"
        assert parts == [800.0, 700.0]

    def test_debit_ignores_credit_transactions(self):
        """Only dc='1' transactions counted for debit balance."""
        txns = [
            ("20250101", "1", 3000.0),
            ("20250201", "2", 1000.0),  # credit - ignored
            ("20250301", "1", 2000.0),
        ]
        date_str, parts = calculate_aging(2000.0, txns, 0.0)
        assert date_str == "01/03/2025"
        assert parts == [2000.0]

    def test_debit_all_insufficient_with_opening(self):
        """All transactions insufficient, falls back to opening balance."""
        txns = [
            ("20250101", "1", 100.0),
            ("20250201", "1", 200.0),
        ]
        date_str, parts = calculate_aging(9999.0, txns, 5000.0)
        assert date_str == "כולל יתרת פתיחה"
        assert parts == [200.0, 100.0, 5000.0]

    def test_debit_all_insufficient_mixed_opening(self):
        """All insufficient, opening has different sign."""
        txns = [("20250101", "1", 100.0)]
        date_str, parts = calculate_aging(500.0, txns, -200.0)
        assert date_str == "מורכב ממספר יתרות"


class TestCreditBalance:
    def test_simple_credit(self):
        """Credit balance selects dc='2' transactions."""
        txns = [
            ("20250115", "2", 500.0),
            ("20250215", "2", 500.0),
            ("20250315", "2", 400.0),
        ]
        # closing=-500, last=400 not enough, 400+500=900 >= 500
        date_str, parts = calculate_aging(-500.0, txns, 0.0)
        assert date_str == "15/02/2025"
        assert parts == [400.0, 500.0]

    def test_credit_ignores_debit_transactions(self):
        """Only dc='2' transactions counted for credit balance."""
        txns = [
            ("20250101", "2", 1000.0),
            ("20250201", "1", 500.0),  # debit - ignored
            ("20250301", "2", 2000.0),
        ]
        date_str, parts = calculate_aging(-2000.0, txns, 0.0)
        assert date_str == "01/03/2025"
        assert parts == [2000.0]


class TestCumulativeStops:
    def test_stops_immediately_when_threshold_met(self):
        """Accumulation stops as soon as cumulative >= target."""
        txns = [
            ("20250101", "1", 100.0),
            ("20250201", "1", 100.0),
            ("20250301", "1", 100.0),
            ("20250401", "1", 200.0),
            ("20250501", "1", 50.0),
        ]
        # target=200, last=50 not enough, 50+200=250 >= 200
        date_str, parts = calculate_aging(200.0, txns, 0.0)
        assert len(parts) == 2
        assert parts == [50.0, 200.0]
        assert date_str == "01/04/2025"


class TestProcessAccounts:
    def test_basic_processing(self):
        pdf_accounts = {
            10001: {'balance_pdf': 1000.0, 'bal_type': 'ח', 'name_visual': 'tset'},
        }
        b11_data = {
            10001: {
                'name': 'test',
                'opening_balance': 0.0,
                'total_debits': 1000.0,
                'total_credits': 0.0,
            },
        }
        transactions = {
            10001: [("20250601", "1", 1000.0)],
        }
        results = process_accounts(pdf_accounts, b11_data, transactions)
        assert len(results) == 1
        assert results[0]['acct_num'] == 10001
        assert results[0]['closing'] == 1000.0
        assert results[0]['debt_start_date'] == "01/06/2025"
        assert results[0]['sum_formula'] == "=1000.00"

    def test_fallback_to_pdf_name(self):
        """When B11 data missing, uses reversed visual name from PDF."""
        pdf_accounts = {
            99999: {'balance_pdf': 500.0, 'bal_type': 'ח', 'name_visual': 'tset'},
        }
        results = process_accounts(pdf_accounts, {}, {})
        assert len(results) == 1
        assert results[0]['name'] == 'test'
        assert results[0]['closing'] == 0.0  # no B11 data
