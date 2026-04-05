"""
bkmv_parser.py - פרסור קובץ מבנה אחיד (BKMVDATA.TXT)
=====================================================
Fixed-width CP862 parser עבור רשומות B11 (כרטיסיות) ו-B1 (תנועות).
"""

import re
from collections import defaultdict

from ageing_report.config.constants import BKMV_ENCODING


def parse_b11_records(bkmv_path, target_accounts):
    """
    קריאת רשומות B11 - כרטיסיות חשבון עם יתרות.

    Returns:
        dict: {acct_num: {name, opening_balance, total_debits, total_credits}}
    """
    b11_data = {}

    with open(bkmv_path, 'rb') as f:
        for line in f:
            if not line.startswith(b'B11'):
                continue
            d = line.rstrip(b'\r\n\x85').decode(BKMV_ENCODING, errors='replace')

            acct_str = d[22:37].strip()
            try:
                acct_num = int(acct_str)
            except ValueError:
                continue

            if acct_num not in target_accounts:
                continue

            name_visual = d[37:87].strip()
            name_logical = (name_visual[::-1]
                .replace('(', '\x00').replace(')', '(').replace('\x00', ')')
                .replace('[', '\x00').replace(']', '[').replace('\x00', ']'))

            amts = re.findall(r'[+-]\d{14}', d[270:])
            opening = int(amts[0]) / 100 if len(amts) >= 1 else 0
            total_debits = int(amts[1]) / 100 if len(amts) >= 2 else 0
            total_credits = int(amts[2]) / 100 if len(amts) >= 3 else 0

            b11_data[acct_num] = {
                'name': name_logical,
                'opening_balance': opening,
                'total_debits': total_debits,
                'total_credits': total_credits,
            }

    return b11_data


def parse_b1_transactions(bkmv_path, target_accounts):
    """
    קריאת רשומות B1 - תנועות יומן.

    Returns:
        dict: {acct_num: [(date_str, dc, amount), ...]}
    """
    transactions = defaultdict(list)

    with open(bkmv_path, 'rb') as f:
        for line in f:
            if not (line.startswith(b'B1') and not line.startswith(b'B11')):
                continue
            d = line.rstrip(b'\r\n\x85').decode(BKMV_ENCODING, errors='replace')

            acct_str = d[172:187].strip()
            try:
                acct_num = int(acct_str)
            except ValueError:
                continue

            if acct_num not in target_accounts:
                continue

            date_str = d[156:164]
            dc = d[202:203].strip()

            amt_match = re.search(r'[+-]\d{14}', d[203:])
            if not amt_match:
                continue
            amount = int(amt_match.group()) / 100

            if dc in ('1', '2') and amount != 0:
                transactions[acct_num].append((date_str, dc, amount))

    return transactions
