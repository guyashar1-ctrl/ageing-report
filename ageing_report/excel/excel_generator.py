"""
excel_generator.py
==================
Generates the output Excel workbook from a list of AccountResult objects.

Workbook structure
------------------
Sheet 1 – "ריכוז"          (Main summary)
Sheet 2 – "תנועות נבחרות"  (Selected movements – used by SUM formulas)
Sheet 3 – "כל התנועות"     (All movements for all processed accounts – audit)
Sheet 4 – "לוג שגיאות"     (Errors and warnings)

All sheets are right-to-left (RTL).  Headers are in Hebrew.

SUM formula design
------------------
The "סכום ח/ז" column in sheet 1 contains a real Excel SUM formula:

    =SUM('תנועות נבחרות'!D{start_row}:D{end_row})

where D is the "סכום" (Amount) column of the "תנועות נבחרות" sheet.

The selected movements for each account are written in a contiguous block
in "תנועות נבחרות", so a single range formula always suffices.

If an account has zero selected movements (zero balance or no data),
the cell is left as a numeric 0 with a note in the log sheet.

Column layout ("ריכוז")
-----------------------
A: מספר חשבון      (Account Number)
B: שם חשבון        (Account Name)
C: יתרה נוכחית     (Current / Closing Balance – signed)
D: יתרה פתיחה      (Opening Balance – signed)
E: סכום ח/ז        (Debt/Credit Summation – SUM formula)
F: תאריך תחילת חוב (Debt Start Date)
G: הערות           (Notes / flags)

Column layout ("תנועות נבחרות")
--------------------------------
A: מספר חשבון  (Account Number)
B: תאריך       (Date)
C: פרטים       (Details / description)
D: סכום        (Amount – always positive, absolute value)
E: ח/ז         (Side: ח or ז)
F: מספר תנועה  (Transaction / entry number)
"""

from __future__ import annotations

import io
from datetime import date as _date
from typing import List, Optional, Tuple

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter

from config.constants import (
    DATE_FORMAT_DISPLAY,
    SHEET_ALL_MVMT,
    SHEET_LOG,
    SHEET_MAIN,
    SHEET_SELECTED,
)
from parsers.bkmv_parser import BkmvData, get_movements_for_account
from processors.balance_calculator import AccountResult
from utils.file_utils import normalize_account_key
from utils.logger import get_accumulated_logs, get_logger

log = get_logger(__name__)

# ---------------------------------------------------------------------------
# Style constants
# ---------------------------------------------------------------------------

_HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")   # dark blue
_SUBHDR_FILL   = PatternFill("solid", fgColor="2E75B6")   # medium blue
_ALT_ROW_FILL  = PatternFill("solid", fgColor="EBF3FB")   # light blue
_WARN_FILL     = PatternFill("solid", fgColor="FFF2CC")   # yellow
_ERROR_FILL    = PatternFill("solid", fgColor="FFE0E0")   # light red

_HEADER_FONT   = Font(name="Arial", size=11, bold=True, color="FFFFFF")
_BODY_FONT     = Font(name="Arial", size=10)
_BOLD_FONT     = Font(name="Arial", size=10, bold=True)

_RTL_ALIGN     = Alignment(horizontal="right", vertical="center",
                            readingOrder=2)          # 2 = RTL
_CENTER_ALIGN  = Alignment(horizontal="center", vertical="center",
                            readingOrder=2)
_NUMBER_FORMAT = '#,##0.00;[Red]-#,##0.00'           # thousands + 2 decimals
_DATE_FORMAT   = 'DD/MM/YYYY'

_THIN = Side(style="thin", color="BFBFBF")
_THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_excel(
    results:     List[AccountResult],
    bkmv:        BkmvData,
    header_code: str,
    header_name: str,
) -> bytes:
    """
    Build the Excel workbook and return it as a bytes object.

    Parameters
    ----------
    results     : List of AccountResult from balance_calculator.
    bkmv        : Parsed BKMVDATA (needed for "all movements" sheet).
    header_code : The 331 header code entered by the user (for title).
    header_name : The 331 header display name (for title).
    """
    wb = Workbook()

    # Remove default sheet
    if wb.active:
        wb.remove(wb.active)

    # Build sheets in logical order
    ws_main     = wb.create_sheet(SHEET_MAIN)
    ws_selected = wb.create_sheet(SHEET_SELECTED)
    ws_all      = wb.create_sheet(SHEET_ALL_MVMT)
    ws_log      = wb.create_sheet(SHEET_LOG)

    # Collect row-range info for selected movements as we build the sheet,
    # then use it to write SUM formulas in the main sheet.
    selected_row_ranges: List[Optional[Tuple[int, int]]] = []

    _build_selected_sheet(ws_selected, results, selected_row_ranges)
    _build_main_sheet(ws_main, results, selected_row_ranges,
                      header_code, header_name)
    _build_all_movements_sheet(ws_all, results, bkmv)
    _build_log_sheet(ws_log)

    # Return as bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    log.info("Excel workbook generated successfully")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

# ── "תנועות נבחרות" (Selected Movements) ──────────────────────────────────

def _build_selected_sheet(
    ws,
    results: List[AccountResult],
    row_ranges: List[Optional[Tuple[int, int]]],
) -> None:
    """Write selected movements and record (start_row, end_row) per account."""
    ws.sheet_view.rightToLeft = True

    headers = [
        "מספר חשבון",   # A
        "תאריך",        # B
        "פרטים",        # C
        "סכום",         # D  ← referenced by SUM formula
        "ח/ז",          # E
        "מספר תנועה",   # F
    ]
    _write_header_row(ws, 1, headers)

    col_widths = [18, 14, 40, 16, 6, 16]
    _set_column_widths(ws, col_widths)

    current_row = 2

    for result in results:
        if not result.selected_movements:
            row_ranges.append(None)
            continue

        start_row = current_row
        is_alt = False

        for sm in result.selected_movements:
            fill = _ALT_ROW_FILL if is_alt else None
            is_alt = not is_alt

            row_data = [
                result.account_number,
                sm.txn_date if sm.txn_date else "",
                sm.details or sm.reference1,
                sm.amount,
                sm.side,
                sm.entry_number,
            ]
            _write_data_row(ws, current_row, row_data, fill=fill)

            # Apply specific cell formats
            date_cell = ws.cell(current_row, 2)
            if sm.txn_date:
                date_cell.number_format = _DATE_FORMAT
                date_cell.value = sm.txn_date

            amount_cell = ws.cell(current_row, 4)
            amount_cell.number_format = _NUMBER_FORMAT

            current_row += 1

        end_row = current_row - 1
        row_ranges.append((start_row, end_row))

    ws.freeze_panes = "A2"


# ── "ריכוז" (Main Summary) ──────────────────────────────────────────────────

def _build_main_sheet(
    ws,
    results:             List[AccountResult],
    selected_row_ranges: List[Optional[Tuple[int, int]]],
    header_code:         str,
    header_name:         str,
) -> None:
    ws.sheet_view.rightToLeft = True

    # Title row
    title = f"ריכוז גיול חשבונות – כותרת {header_code}  {header_name}"
    title_cell = ws.cell(1, 1, value=title)
    title_cell.font = Font(name="Arial", size=13, bold=True, color="1F4E79")
    title_cell.alignment = _RTL_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)

    # Column headers
    col_headers = [
        "מספר חשבון",          # A
        "שם חשבון",            # B
        "יתרה נוכחית",         # C
        "יתרה פתיחה",          # D
        f"סכום ח/ז",           # E – SUM formula column
        "תאריך תחילת חוב",     # F
        "הערות",               # G
    ]
    _write_header_row(ws, 2, col_headers)

    col_widths = [18, 35, 18, 18, 18, 18, 40]
    _set_column_widths(ws, col_widths)

    is_alt = False
    for row_idx, (result, row_range) in enumerate(
            zip(results, selected_row_ranges), start=3):

        fill = _ALT_ROW_FILL if is_alt else None
        is_alt = not is_alt

        # Notes column
        notes = ""
        if result.threshold_not_reached:
            notes = "⚠ סכום תנועות לא הגיע ליתרה הסגירה"
        if not result.selected_movements and abs(result.closing_balance) > 0.005:
            notes = "⚠ לא נמצאו תנועות"
        if result.warnings:
            notes = (notes + " | " if notes else "") + result.warnings[0][:60]
        if result.threshold_not_reached:
            fill = _WARN_FILL

        ws.cell(row_idx, 1, value=result.account_number).alignment = _RTL_ALIGN
        ws.cell(row_idx, 2, value=result.account_name).alignment  = _RTL_ALIGN

        # Closing balance (signed)
        c_cell = ws.cell(row_idx, 3, value=result.closing_balance)
        c_cell.number_format = _NUMBER_FORMAT
        c_cell.alignment = _RTL_ALIGN

        # Opening balance (signed)
        d_cell = ws.cell(row_idx, 4, value=result.opening_balance)
        d_cell.number_format = _NUMBER_FORMAT
        d_cell.alignment = _RTL_ALIGN

        # SUM formula – column E
        e_cell = ws.cell(row_idx, 5)
        if row_range:
            start_row, end_row = row_range
            # Reference the "Amount" column (D) in "תנועות נבחרות" sheet
            sheet_ref = f"'{SHEET_SELECTED}'"
            if start_row == end_row:
                formula = f"={sheet_ref}!D{start_row}"
            else:
                formula = f"=SUM({sheet_ref}!D{start_row}:D{end_row})"
            e_cell.value = formula
        else:
            # Zero balance or no movements
            e_cell.value = 0
        e_cell.number_format = _NUMBER_FORMAT
        e_cell.alignment = _RTL_ALIGN

        # Debt start date – column F
        f_cell = ws.cell(row_idx, 6)
        if result.debt_start_date:
            f_cell.value = result.debt_start_date
            f_cell.number_format = _DATE_FORMAT
        f_cell.alignment = _RTL_ALIGN

        # Notes – column G
        g_cell = ws.cell(row_idx, 7, value=notes)
        g_cell.alignment = _RTL_ALIGN
        if notes.startswith("⚠"):
            g_cell.font = Font(name="Arial", size=10, color="C00000")

        # Apply row fill
        if fill:
            for col in range(1, 8):
                ws.cell(row_idx, col).fill = fill

        # Apply borders
        for col in range(1, 8):
            ws.cell(row_idx, col).border = _THIN_BORDER
            ws.cell(row_idx, col).font = _BODY_FONT

    # Totals row
    total_row = 3 + len(results)
    ws.cell(total_row, 1, value="סה\"כ").font = _BOLD_FONT
    for col_letter, col_idx in [("C", 3), ("D", 4), ("E", 5)]:
        if total_row > 3:
            total_cell = ws.cell(total_row, col_idx)
            total_cell.value = (
                f"=SUM({col_letter}3:{col_letter}{total_row - 1})"
            )
            total_cell.number_format = _NUMBER_FORMAT
            total_cell.font = _BOLD_FONT
            total_cell.alignment = _RTL_ALIGN

    ws.freeze_panes = "A3"


# ── "כל התנועות" (All Movements) ───────────────────────────────────────────

def _build_all_movements_sheet(
    ws,
    results: List[AccountResult],
    bkmv: BkmvData,
) -> None:
    ws.sheet_view.rightToLeft = True

    headers = [
        "מספר חשבון",   # A
        "שם חשבון",     # B
        "תאריך",        # C
        "פרטים",        # D
        "סכום",         # E
        "ח/ז",          # F
        "מספר תנועה",   # G
        "נבחרה",        # H – Is this movement in the selected set?
    ]
    _write_header_row(ws, 1, headers)
    col_widths = [18, 30, 14, 40, 16, 6, 16, 10]
    _set_column_widths(ws, col_widths)

    current_row = 2

    for result in results:
        # Get ALL movements for this account (not just selected)
        all_mvmts = get_movements_for_account(bkmv, result.account_number, year_only=True)

        # Build set of (entry_number, line_number) for selected movements
        selected_set = {
            (sm.entry_number, sm.line_number)
            for sm in result.selected_movements
        }

        account_norm = normalize_account_key(result.account_number)

        for m in sorted(all_mvmts,
                        key=lambda x: (x.txn_date or _date(1900,1,1),
                                       x.entry_number, x.line_number)):
            # Determine if this movement is debit or credit for this account
            d_norm = m.debit_account_norm
            is_debit = (d_norm == account_norm or
                        d_norm.lstrip("0") == account_norm.lstrip("0"))
            side = "ח" if is_debit else "ז"

            in_selected = "✓" if (m.entry_number, m.line_number) in selected_set else ""

            ws.cell(current_row, 1, result.account_number).alignment = _RTL_ALIGN
            ws.cell(current_row, 2, result.account_name).alignment   = _RTL_ALIGN

            date_cell = ws.cell(current_row, 3)
            if m.txn_date:
                date_cell.value = m.txn_date
                date_cell.number_format = _DATE_FORMAT
            date_cell.alignment = _RTL_ALIGN

            ws.cell(current_row, 4, m.details or m.reference1).alignment = _RTL_ALIGN

            amt_cell = ws.cell(current_row, 5, m.amount)
            amt_cell.number_format = _NUMBER_FORMAT
            amt_cell.alignment = _RTL_ALIGN

            ws.cell(current_row, 6, side).alignment = _CENTER_ALIGN
            ws.cell(current_row, 7, m.entry_number).alignment = _RTL_ALIGN
            ws.cell(current_row, 8, in_selected).alignment = _CENTER_ALIGN

            if in_selected:
                for col in range(1, 9):
                    ws.cell(current_row, col).fill = _ALT_ROW_FILL

            for col in range(1, 9):
                ws.cell(current_row, col).font = _BODY_FONT
                ws.cell(current_row, col).border = _THIN_BORDER

            current_row += 1

    ws.freeze_panes = "A2"


# ── "לוג שגיאות" (Log / Errors) ────────────────────────────────────────────

def _build_log_sheet(ws) -> None:
    ws.sheet_view.rightToLeft = True

    headers = ["רמה", "תאריך ושעה", "הודעה"]
    _write_header_row(ws, 1, headers)
    _set_column_widths(ws, [12, 20, 100])

    current_row = 2
    for level, message, timestamp in get_accumulated_logs():
        fill = None
        if level in ("ERROR", "CRITICAL"):
            fill = _ERROR_FILL
        elif level == "WARNING":
            fill = _WARN_FILL

        ws.cell(current_row, 1, level).alignment = _CENTER_ALIGN
        ws.cell(current_row, 2, timestamp).alignment = _CENTER_ALIGN
        ws.cell(current_row, 3, message).alignment = _RTL_ALIGN

        if fill:
            for col in range(1, 4):
                ws.cell(current_row, col).fill = fill

        for col in range(1, 4):
            ws.cell(current_row, col).font = _BODY_FONT
            ws.cell(current_row, col).border = _THIN_BORDER

        current_row += 1

    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _write_header_row(ws, row: int, labels: List[str]) -> None:
    for col_idx, label in enumerate(labels, start=1):
        cell = ws.cell(row, col_idx, value=label)
        cell.fill   = _HEADER_FILL
        cell.font   = _HEADER_FONT
        cell.alignment = _RTL_ALIGN
        cell.border = _THIN_BORDER
    ws.row_dimensions[row].height = 22


def _write_data_row(ws, row: int, values: list, fill=None) -> None:
    for col_idx, value in enumerate(values, start=1):
        cell = ws.cell(row, col_idx, value=value)
        cell.font      = _BODY_FONT
        cell.alignment = _RTL_ALIGN
        cell.border    = _THIN_BORDER
        if fill:
            cell.fill = fill


def _set_column_widths(ws, widths: List[int]) -> None:
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
