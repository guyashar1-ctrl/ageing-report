"""
aging_calculator.py - חישוב גיול חובות
=======================================
לוגיקת סכימה אחורה: סוכמים תנועות מסוף השנה להתחלה
עד שהסכום המצטבר עובר את יתרת הסגירה.
"""

import re

from ageing_report.config.constants import ZERO_THRESHOLD


def calculate_aging(closing, acct_txns, opening):
    """
    חישוב תאריך תחילת חוב.

    Args:
        closing: יתרת סגירה
        acct_txns: [(date_str, dc, amount), ...]
        opening: יתרת פתיחה

    Returns:
        (debt_start_date, sum_parts)
    """
    debt_start_date = ""
    sum_parts = []

    if closing > ZERO_THRESHOLD:
        relevant = [(d, amt) for d, dc, amt in acct_txns if dc == '1' and amt > 0]
        relevant.sort(key=lambda x: x[0], reverse=True)
        target = closing
    elif closing < -ZERO_THRESHOLD:
        relevant = [(d, amt) for d, dc, amt in acct_txns if dc == '2' and amt > 0]
        relevant.sort(key=lambda x: x[0], reverse=True)
        target = abs(closing)
    else:
        return "", []

    cumulative = 0
    for date_str, amt in relevant:
        cumulative += amt
        sum_parts.append(amt)
        if cumulative >= target - ZERO_THRESHOLD:
            debt_start_date = f"{date_str[6:8]}/{date_str[4:6]}/{date_str[0:4]}"
            break

    if not debt_start_date and sum_parts:
        if (closing > 0 and opening > 0) or (closing < 0 and opening < 0):
            debt_start_date = "כולל יתרת פתיחה"
            sum_parts.append(abs(opening))
        else:
            debt_start_date = "מורכב ממספר יתרות"

    return debt_start_date, sum_parts


def process_accounts(pdf_accounts, b11_data, transactions):
    """
    עיבוד כל החשבונות וחישוב גיול.

    Returns:
        list[dict]: תוצאות ממוינות לפי מספר חשבון
    """
    results = []

    for acct_num in sorted(pdf_accounts.keys()):
        pdf_info = pdf_accounts[acct_num]
        b11 = b11_data.get(acct_num, {})

        name = b11.get('name', pdf_info['name_visual'][::-1])
        opening = b11.get('opening_balance', 0)
        total_debits = b11.get('total_debits', 0)
        total_credits = b11.get('total_credits', 0)
        closing = opening + total_debits - total_credits

        acct_txns = transactions.get(acct_num, [])
        debt_start_date, sum_parts = calculate_aging(closing, acct_txns, opening)

        sum_formula = ""
        if sum_parts:
            sum_formula = "=" + "+".join(f"{x:.2f}" for x in sum_parts)

        results.append({
            'acct_num': acct_num,
            'name': name,
            'closing': closing,
            'opening': opening,
            'debt_start_date': debt_start_date,
            'sum_formula': sum_formula,
        })

    return results
