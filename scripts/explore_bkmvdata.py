#!/usr/bin/env python3
"""
סקריפט חקירה: ניתוח מבנה קובץ BKMVDATA.TXT
=============================================
סקריפט זה מנתח את מבנה קובץ המבנה האחיד ומציג:
- ספירת רשומות לפי סוג
- מבנה שדות של כל סוג רשומה
- דוגמאות לרשומות

שימוש:
    python3 scripts/explore_bkmvdata.py bkmv_data/BKMVDATA.TXT
"""

import sys
import re
from collections import Counter


def analyze_file(filepath):
    """ניתוח כללי של הקובץ"""
    print(f"מנתח: {filepath}")
    print("=" * 60)

    record_types = Counter()
    total_lines = 0

    with open(filepath, 'rb') as f:
        for line in f:
            total_lines += 1
            d = line.decode('cp862', errors='replace')
            rtype = d[:3] if d[:3] in ('B11', 'D11', 'D12') else d[:2]
            record_types[rtype] += 1

    print(f"\nסה\"כ שורות: {total_lines:,}")
    print(f"\nסוגי רשומות:")
    print(f"{'סוג':<6} {'כמות':>10}  תיאור")
    print("-" * 50)

    descriptions = {
        'A1': 'כותרת קובץ',
        'B1': 'תנועות יומן (פקודות)',
        'B11': 'כרטיסיות חשבון (מאזן בוחן)',
        'C1': 'כותרות מסמכים (חשבוניות)',
        'D11': 'שורות פירוט מסמכים',
        'D12': 'שורות פירוט נוספות',
        'M1': 'פריטים/מוצרים',
    }

    for rtype, count in sorted(record_types.items()):
        desc = descriptions.get(rtype, '?')
        print(f"{rtype:<6} {count:>10,}  {desc}")


def show_b11_samples(filepath, count=5):
    """הצגת דוגמאות לרשומות B11 (כרטיסיות)"""
    print(f"\n{'=' * 60}")
    print(f"דוגמאות B11 (כרטיסיות חשבון) - {count} ראשונות:")
    print("-" * 60)

    with open(filepath, 'rb') as f:
        found = 0
        for line in f:
            if line.startswith(b'B11'):
                raw = line.rstrip(b'\r\n\x85')
                d = raw.decode('cp862', errors='replace')

                acct = d[22:37].strip()
                name = d[37:87].strip()[::-1]  # היפוך לסדר לוגי
                group = d[87:97].strip()

                amts = re.findall(r'[+-]\d{14}', d[270:])
                opening = int(amts[0]) / 100 if len(amts) >= 1 else 0
                debits = int(amts[1]) / 100 if len(amts) >= 2 else 0
                credits = int(amts[2]) / 100 if len(amts) >= 3 else 0
                closing = opening + debits - credits

                print(f"\n  חשבון: {acct}")
                print(f"  שם: {name}")
                print(f"  קבוצה: {group}")
                print(f"  יתרת פתיחה: {opening:>15,.2f}")
                print(f"  סה\"כ חובה:  {debits:>15,.2f}")
                print(f"  סה\"כ זכות:  {credits:>15,.2f}")
                print(f"  יתרת סגירה: {closing:>15,.2f}")

                found += 1
                if found >= count:
                    break


def show_b1_samples(filepath, acct_filter=None, count=5):
    """הצגת דוגמאות לרשומות B1 (תנועות יומן)"""
    print(f"\n{'=' * 60}")
    label = f"דוגמאות B1 (תנועות) - חשבון {acct_filter}" if acct_filter else f"דוגמאות B1 (תנועות) - {count} ראשונות"
    print(f"{label}:")
    print("-" * 60)

    with open(filepath, 'rb') as f:
        found = 0
        for line in f:
            if not (line.startswith(b'B1') and not line.startswith(b'B11')):
                continue
            raw = line.rstrip(b'\r\n\x85')
            d = raw.decode('cp862', errors='replace')

            acct = d[172:187].strip()
            if acct_filter and acct != str(acct_filter):
                continue

            date_str = d[156:164]
            counter_acct = d[187:202].strip()
            dc = d[202:203].strip()
            dc_label = 'חובה' if dc == '1' else 'זכות' if dc == '2' else '?'

            amt_match = re.search(r'[+-]\d{14}', d[203:])
            amount = int(amt_match.group()) / 100 if amt_match else 0

            formatted_date = f"{date_str[6:8]}/{date_str[4:6]}/{date_str[0:4]}"
            print(f"  {formatted_date}  חשבון:{acct:<8} נגדי:{counter_acct:<8} {dc_label}  {amount:>12,.2f}")

            found += 1
            if found >= count:
                break

    if found == 0:
        print("  לא נמצאו תנועות")


def show_field_positions(filepath):
    """הצגת מיקומי שדות בכל סוג רשומה"""
    print(f"\n{'=' * 60}")
    print("מפת שדות:")
    print("-" * 60)

    print("""
B11 (כרטיסיית חשבון):
  pos 0-2:     סוג רשומה ("B11")
  pos 3-12:    מספר רץ (10 ספרות)
  pos 13-21:   מספר עוסק (9 ספרות)
  pos 22-36:   מספר חשבון (15 תווים)
  pos 37-86:   שם חשבון (50 תווים, עברית ויזואלית)
  pos 87-96:   קוד קבוצת מיון
  pos 277-291: יתרת פתיחה (סימן + 14 ספרות, באגורות)
  pos 292-306: סה"כ חובה
  pos 307-321: סה"כ זכות

B1 (תנועת יומן):
  pos 0-1:     סוג רשומה ("B1")
  pos 2-12:    מספר רץ (11 ספרות)
  pos 13-21:   מספר עוסק (9 ספרות)
  pos 156-163: תאריך ערך (YYYYMMDD)
  pos 164-171: תאריך תנועה (YYYYMMDD)
  pos 172-186: מספר חשבון ראשי (15 תווים)
  pos 187-201: מספר חשבון נגדי (15 תווים)
  pos 202:     סוג: 1=חובה, 2=זכות
  pos 206+:    סכום (סימן + 14 ספרות, באגורות)

C1 (מסמך):
  pos 0-1:     סוג רשומה ("C1")
  pos 2-12:    מספר רץ
  pos 13-21:   מספר עוסק
  pos 25-44:   מספר/סוג מסמך (20 תווים)
  pos 45-52:   תאריך מסמך (YYYYMMDD)
  pos 53-56:   שעה (HHMM)
  pos 57-106:  שם לקוח/ספק (50 תווים)

קידוד: CP862 (DOS Hebrew), טקסט עברי בסדר ויזואלי (הפוך)
פורמט סכום: +00000006728422 = 67,284.22 ש"ח (חלוקה ב-100)
""")


def verify_account(filepath, acct_num):
    """אימות חשבון ספציפי - B11 + סיכום B1"""
    print(f"\n{'=' * 60}")
    print(f"אימות חשבון {acct_num}:")
    print("-" * 60)

    # B11
    with open(filepath, 'rb') as f:
        for line in f:
            if line.startswith(b'B11'):
                d = line.rstrip(b'\r\n\x85').decode('cp862', errors='replace')
                acct = d[22:37].strip()
                if acct == str(acct_num):
                    name = d[37:87].strip()[::-1]
                    amts = re.findall(r'[+-]\d{14}', d[270:])
                    opening = int(amts[0]) / 100 if len(amts) >= 1 else 0
                    debits = int(amts[1]) / 100 if len(amts) >= 2 else 0
                    credits = int(amts[2]) / 100 if len(amts) >= 3 else 0
                    print(f"  B11 שם: {name}")
                    print(f"  B11 פתיחה: {opening:,.2f}")
                    print(f"  B11 חובה: {debits:,.2f}")
                    print(f"  B11 זכות: {credits:,.2f}")
                    print(f"  B11 סגירה: {opening + debits - credits:,.2f}")
                    break

    # B1 - sum all transactions
    b1_debit = 0
    b1_credit = 0
    b1_count = 0
    with open(filepath, 'rb') as f:
        for line in f:
            if not (line.startswith(b'B1') and not line.startswith(b'B11')):
                continue
            d = line.rstrip(b'\r\n\x85').decode('cp862', errors='replace')
            acct = d[172:187].strip()
            if acct != str(acct_num):
                continue
            dc = d[202:203].strip()
            amt_match = re.search(r'[+-]\d{14}', d[203:])
            if amt_match:
                amt = int(amt_match.group()) / 100
                if dc == '1':
                    b1_debit += amt
                elif dc == '2':
                    b1_credit += amt
                b1_count += 1

    print(f"\n  B1 תנועות: {b1_count}")
    print(f"  B1 סה\"כ חובה: {b1_debit:,.2f}")
    print(f"  B1 סה\"כ זכות: {b1_credit:,.2f}")

    if abs(b1_debit - debits) < 0.01 and abs(b1_credit - credits) < 0.01:
        print(f"\n  ✓ סכומי B1 תואמים ל-B11")
    else:
        print(f"\n  ✗ אי-התאמה! B1 חובה={b1_debit:,.2f} vs B11={debits:,.2f}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("שימוש: python3 explore_bkmvdata.py <BKMVDATA_PATH> [ACCOUNT_NUM]")
        print("דוגמה: python3 explore_bkmvdata.py bkmv_data/BKMVDATA.TXT 17009")
        sys.exit(1)

    filepath = sys.argv[1]

    if not __import__('os').path.exists(filepath):
        print(f"קובץ לא נמצא: {filepath}")
        sys.exit(1)

    analyze_file(filepath)
    show_field_positions(filepath)
    show_b11_samples(filepath)

    if len(sys.argv) >= 3:
        acct = sys.argv[2]
        show_b1_samples(filepath, acct_filter=acct, count=20)
        verify_account(filepath, acct)
    else:
        show_b1_samples(filepath, count=10)
