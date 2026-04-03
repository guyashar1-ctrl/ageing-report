#!/usr/bin/env python3
"""
סקריפט חילוץ: חילוץ חשבונות מסעיף מסוים בדוח 331
===================================================
מחלץ את רשימת החשבונות מסעיף נתון בדוח 331 ומדפיס אותם כ-CSV.
ניתן להשתמש בכל קוד סעיף (1342, 1302, 2660 וכו').

שימוש:
    python3 scripts/extract_pdf_accounts.py <PDF_PATH> [SECTION_CODE]
    python3 scripts/extract_pdf_accounts.py report.pdf 1342
    python3 scripts/extract_pdf_accounts.py report.pdf 1342 > accounts.csv
"""

import sys
import re

try:
    import pdfplumber
except ImportError:
    print("חסרת חבילה. יש להריץ: pip3 install pdfplumber", file=sys.stderr)
    sys.exit(1)


# מיפוי קודי סעיפים לשמות בסדר ויזואלי (הפוך)
# הוסף סעיפים נוספים כאן לפי הצורך
SECTION_NAMES = {
    '1302': 'ילארשי עבטמ 1302',
    '1304': 'ץוח עבטמ 1304',
    '1342': 'ץראב םיחותפ תובוח 1342',
    '1348': 'יארשא יסיטרכ 1348',
    '1334': 'יארשא יסיטרכ 1334',
    '2660': 'תוחוקלמ תומדקמ 2660',
    '133':  'הייבגל תואחמה 133',
    '135':  'הבוח תורתי םיבייח 135',
}


def extract_section(pdf_path, section_code):
    """חילוץ חשבונות מסעיף מסוים"""
    # בניית מחרוזת זיהוי לסעיף
    if section_code in SECTION_NAMES:
        section_marker = SECTION_NAMES[section_code]
    else:
        # ניסיון גנרי - חיפוש לפי קוד הסעיף
        section_marker = section_code

    accounts = []
    total = 0

    with pdfplumber.open(pdf_path) as pdf:
        in_section = False
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                # תחילת סעיף
                if section_marker in line and 'כ"הס' not in line and not in_section:
                    in_section = True
                    continue
                # המשך
                if in_section and '(ךשמה)' in line:
                    continue
                # סוף סעיף
                if in_section and section_code in line and 'כ"הס' in line:
                    m = re.match(r'^([\d,]+\.\d{2})', line.strip())
                    if m:
                        total = float(m.group(1).replace(',', ''))
                    in_section = False
                    continue
                if in_section:
                    m = re.match(
                        r'^([\d,]+\.\d{2})\s+(ח|ז)\s+(\d+)\s+(.+?)\s+(\d{3,6})$',
                        line.strip()
                    )
                    if m:
                        accounts.append({
                            'acct_num': int(m.group(5)),
                            'name_visual': m.group(4).strip(),
                            'name_logical': m.group(4).strip()[::-1],
                            'group': m.group(3),
                            'balance': float(m.group(1).replace(',', '')),
                            'bal_type': m.group(2),
                        })

    return accounts, total


def main():
    if len(sys.argv) < 2:
        print("שימוש: python3 extract_pdf_accounts.py <PDF_PATH> [SECTION_CODE]")
        print("ברירת מחדל: סעיף 1342")
        sys.exit(1)

    pdf_path = sys.argv[1]
    section_code = sys.argv[2] if len(sys.argv) >= 3 else '1342'

    accounts, total = extract_section(pdf_path, section_code)

    if not accounts:
        print(f"לא נמצאו חשבונות בסעיף {section_code}", file=sys.stderr)
        sys.exit(1)

    # הדפסת כותרת CSV
    print("מספר_חשבון,שם_חשבון,קבוצה,יתרה,סוג")

    for a in accounts:
        name = a['name_logical'].replace(',', ' ')
        sign = '' if a['bal_type'] == 'ח' else '-'
        print(f"{a['acct_num']},{name},{a['group']},{sign}{a['balance']:.2f},{a['bal_type']}")

    # סיכום ל-stderr
    debit_count = sum(1 for a in accounts if a['bal_type'] == 'ח')
    credit_count = sum(1 for a in accounts if a['bal_type'] == 'ז')
    print(f"\nסעיף {section_code}: {len(accounts)} חשבונות ({debit_count} חובה, {credit_count} זכות), סה\"כ {total:,.2f}",
          file=sys.stderr)


if __name__ == '__main__':
    main()
