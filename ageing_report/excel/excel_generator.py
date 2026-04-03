"""
excel_generator.py - יצירת קובץ Excel מעוצב
=============================================
גיליון בודד RTL עם כותרות כחולות, גבולות, freeze panes ו-auto-filter.
"""

from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from ageing_report.config.constants import EXCEL_HEADERS, SHEET_NAME


def generate_excel(results):
    """
    יצירת קובץ Excel מתוצאות הגיול.

    Args:
        results: list[dict] מ-process_accounts()

    Returns:
        bytes: תוכן קובץ Excel
    """
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.sheet_view.rightToLeft = True

    # סגנונות
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, size=12, name='Arial', color='FFFFFF')
    data_font = Font(size=11, name='Arial')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # כותרות
    for col, header in enumerate(EXCEL_HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # שורות נתונים
    for row_idx, r in enumerate(results, 2):
        cells_data = [
            (1, r['acct_num'], 'center', None),
            (2, r['name'], 'right', None),
            (3, r['closing'], 'center', '#,##0.00'),
            (4, r['opening'], 'center', '#,##0.00'),
            (5, r['debt_start_date'], 'center', None),
            (6, r['sum_formula'] or "", 'center', '#,##0.00'),
        ]
        for col, value, align, fmt in cells_data:
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = data_font
            cell.alignment = Alignment(horizontal=align)
            cell.border = thin_border
            if fmt:
                cell.number_format = fmt

    # רוחב עמודות
    widths = {'A': 15, 'B': 45, 'C': 18, 'D': 18, 'E': 22, 'F': 50}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:F{len(results) + 1}"

    # החזרת bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
