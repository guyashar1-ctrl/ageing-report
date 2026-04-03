"""
pdf_parser.py - פרסור דוח 331 (PDF)
====================================
חילוץ רשימת חשבונות מסעיף נתון בדוח 331.
"""

import re
import pdfplumber

from ageing_report.config.constants import SECTION_NAMES, SECTION_CONTINUE_MARKER


def parse_pdf_section(pdf_path, section_code):
    """
    חילוץ חשבונות מסעיף מסוים בדוח 331.

    Args:
        pdf_path: נתיב לקובץ PDF
        section_code: קוד סעיף (למשל '1342')

    Returns:
        dict: {acct_num: {balance_pdf, bal_type, name_visual}}
    """
    section_marker = SECTION_NAMES.get(section_code, section_code)
    pdf_accounts = {}

    with pdfplumber.open(pdf_path) as pdf:
        in_section = False
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                if section_marker in line and 'כ"הס' not in line and not in_section:
                    in_section = True
                    continue
                if in_section and SECTION_CONTINUE_MARKER in line:
                    continue
                if in_section and section_code in line and 'כ"הס' in line:
                    in_section = False
                    continue
                if in_section:
                    m = re.match(
                        r'^([\d,]+\.\d{2})\s+(ח|ז)\s+(\d+)\s+(.+?)\s+(\d{3,6})$',
                        line.strip()
                    )
                    if m:
                        acct_num = int(m.group(5))
                        pdf_accounts[acct_num] = {
                            'balance_pdf': float(m.group(1).replace(',', '')),
                            'bal_type': m.group(2),
                            'name_visual': m.group(4).strip(),
                        }

    return pdf_accounts
