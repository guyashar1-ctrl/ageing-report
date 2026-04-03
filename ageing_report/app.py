"""
app.py
======
Streamlit web application for the Ageing Report system.

Workflow
--------
1.  User uploads a ZIP archive (or enters a folder path) containing the
    Uniform Structure files (BKMVDATA.TXT and friends).
2.  User uploads the 331 PDF report.
3.  User enters the header code (e.g., 1342).
4.  User clicks "אשר / Confirm".
5.  The system processes the data and offers a downloadable Excel file.

Run with:
    streamlit run ageing_report/app.py
or (from the ageing_report/ directory):
    streamlit run app.py
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import streamlit as st

# Ensure project root is on sys.path regardless of working directory
_HERE = Path(__file__).parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from config.constants import TARGET_YEAR
from excel.excel_generator import generate_excel
from parsers.bkmv_parser import parse_bkmvdata
from parsers.report331_parser import parse_331_pdf
from processors.balance_calculator import process_all_accounts
from utils.file_utils import (
    cleanup_temp_dir,
    extract_zip,
    find_bkmvdata,
    validate_directory,
    validate_file,
)
from utils.logger import clear_accumulated_logs, get_accumulated_logs, get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Ageing Report – גיול חשבונות",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Inject right-to-left CSS for the entire page
st.markdown(
    """
    <style>
    body, .stApp { direction: rtl; }
    .stTextInput > label,
    .stFileUploader > label,
    .stButton > button { direction: rtl; }
    h1, h2, h3 { direction: rtl; }
    .stAlert { direction: rtl; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.title("📊 מערכת גיול חשבונות")
st.markdown(
    "**קובץ מבנה אחיד** + **דו\"ח 331** → **קובץ Excel** עם נוסחאות SUM אמיתיות"
)
st.divider()


# ---------------------------------------------------------------------------
# Input section
# ---------------------------------------------------------------------------

col_left, col_right = st.columns([1, 1])

with col_right:
    st.subheader("1. קבצי מבנה אחיד")
    input_method = st.radio(
        "בחר שיטת קלט:",
        options=["העלאת קובץ ZIP", "נתיב תיקייה"],
        horizontal=True,
        key="input_method",
    )

    zip_bytes: bytes | None = None
    folder_path: str = ""

    if input_method == "העלאת קובץ ZIP":
        zip_file = st.file_uploader(
            "העלה קובץ ZIP המכיל את קבצי המבנה האחיד",
            type=["zip"],
            key="zip_uploader",
        )
        if zip_file:
            zip_bytes = zip_file.read()
            st.success(f"ZIP נטען: {zip_file.name}  ({len(zip_bytes):,} bytes)")
    else:
        folder_path = st.text_input(
            "הזן נתיב מלא לתיקייה המכילה את קבצי המבנה האחיד",
            placeholder=r"לדוגמה: C:\Accounting\2025\UniformStructure",
            key="folder_path_input",
        )
        if folder_path:
            if validate_directory(folder_path):
                st.success(f"תיקייה נמצאה: {folder_path}")
            else:
                st.error(f"התיקייה לא נמצאה או אינה נגישה: {folder_path}")

with col_left:
    st.subheader("2. קובץ דו\"ח 331")
    pdf_file = st.file_uploader(
        "העלה את קובץ ה-PDF של דו\"ח 331",
        type=["pdf"],
        key="pdf_uploader",
    )
    if pdf_file:
        st.success(f"PDF נטען: {pdf_file.name}  ({pdf_file.size:,} bytes)")

st.divider()

st.subheader("3. קוד כותרת מדו\"ח 331")
header_code = st.text_input(
    "הזן קוד כותרת (לדוגמה: 1342)",
    placeholder="1342",
    key="header_code_input",
    max_chars=20,
)

st.divider()

# ---------------------------------------------------------------------------
# Validation and processing
# ---------------------------------------------------------------------------

confirm_clicked = st.button(
    "✅ אשר ועבד",
    type="primary",
    key="confirm_btn",
    disabled=(
        (zip_bytes is None and not folder_path.strip()) or
        pdf_file is None or
        not header_code.strip()
    ),
)

if confirm_clicked:
    clear_accumulated_logs()
    tmp_dir = None

    progress = st.progress(0, text="מתחיל עיבוד…")

    try:
        # ----------------------------------------------------------------
        # Step 1: Resolve working directory
        # ----------------------------------------------------------------
        progress.progress(10, text="מאתר קבצי מבנה אחיד…")

        if zip_bytes:
            tmp_dir = extract_zip(zip_bytes)
            working_dir = tmp_dir
            st.info(f"ZIP חולץ לתיקייה זמנית: {working_dir}")
        else:
            working_dir = folder_path.strip()
            if not validate_directory(working_dir):
                st.error(f"התיקייה '{working_dir}' אינה תקינה.")
                st.stop()

        # ----------------------------------------------------------------
        # Step 2: Find BKMVDATA
        # ----------------------------------------------------------------
        bkmvdata_path = find_bkmvdata(working_dir)
        if not bkmvdata_path:
            st.error(
                "לא נמצא קובץ BKMVDATA.TXT בתיקייה.  "
                "וודא שהקובץ קיים (או שם מקביל כגון TXT.BKMVDATA) "
                "ונסה שוב."
            )
            st.stop()

        st.info(f"נמצא BKMVDATA: `{bkmvdata_path}`")

        # ----------------------------------------------------------------
        # Step 3: Save PDF to temp file
        # ----------------------------------------------------------------
        pdf_bytes = pdf_file.read()
        with tempfile.NamedTemporaryFile(
            suffix=".pdf", delete=False, prefix="report331_"
        ) as pdf_tmp:
            pdf_tmp.write(pdf_bytes)
            pdf_path = pdf_tmp.name

        # ----------------------------------------------------------------
        # Step 4: Parse BKMVDATA
        # ----------------------------------------------------------------
        progress.progress(25, text="מפרסר קבצי מבנה אחיד…")
        with st.spinner("מפרסר BKMVDATA.TXT…"):
            bkmv = parse_bkmvdata(bkmvdata_path)

        st.success(
            f"BKMVDATA: {len(bkmv.accounts):,} חשבונות, "
            f"{len(bkmv.movements):,} תנועות "
            f"({len(bkmv.movements_target_year):,} בשנת {TARGET_YEAR})"
        )
        if bkmv.warnings:
            with st.expander(f"⚠ {len(bkmv.warnings)} אזהרות פרסור BKMVDATA"):
                for w in bkmv.warnings[:20]:
                    st.warning(w)

        # ----------------------------------------------------------------
        # Step 5: Parse 331 PDF
        # ----------------------------------------------------------------
        progress.progress(45, text="מפרסר דו\"ח 331…")
        with st.spinner("מפרסר PDF של דו\"ח 331…"):
            report = parse_331_pdf(pdf_path, header_code.strip())

        if report.header_name:
            st.success(
                f"כותרת נמצאה: **{header_code} – {report.header_name}** "
                f"({len(report.accounts)} חשבונות)"
            )
        else:
            st.warning(
                f"קוד הכותרת '{header_code}' לא נמצא בדו\"ח 331.  "
                f"בדוק שהקוד נכון."
            )

        if report.warnings:
            with st.expander(f"⚠ {len(report.warnings)} אזהרות פרסור PDF"):
                for w in report.warnings:
                    st.warning(w)

        if not report.accounts:
            st.error("לא נמצאו חשבונות תחת הכותרת שנבחרה.  בדוק קוד כותרת וקובץ PDF.")
            st.stop()

        # ----------------------------------------------------------------
        # Step 6: Calculate balances
        # ----------------------------------------------------------------
        progress.progress(65, text="מחשב גיול חשבונות…")
        with st.spinner("מחשב יתרות וסיכומי ח/ז…"):
            results = process_all_accounts(report.accounts, bkmv)

        not_found  = sum(1 for r in results if not r.selected_movements
                         and abs(r.closing_balance) > 0.005)
        with_warn  = sum(1 for r in results if r.warnings)
        total_accs = len(results)

        st.success(
            f"עובדו {total_accs} חשבונות | "
            f"{with_warn} עם אזהרות | "
            f"{not_found} ללא תנועות"
        )

        # ----------------------------------------------------------------
        # Step 7: Generate Excel
        # ----------------------------------------------------------------
        progress.progress(80, text="מייצר קובץ Excel…")
        with st.spinner("מייצר Excel עם נוסחאות SUM…"):
            excel_bytes = generate_excel(
                results=results,
                bkmv=bkmv,
                header_code=header_code.strip(),
                header_name=report.header_name,
            )

        progress.progress(100, text="הושלם!")

        # ----------------------------------------------------------------
        # Step 8: Download button
        # ----------------------------------------------------------------
        st.divider()
        st.subheader("📥 הורדת קובץ Excel")

        filename = f"Ageing_{header_code.strip()}_{TARGET_YEAR}.xlsx"

        st.download_button(
            label=f"⬇ הורד קובץ Excel – {filename}",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel",
        )

        # Preview table
        with st.expander("👁 תצוגה מקדימה של תוצאות"):
            import pandas as pd
            preview_data = []
            for r in results:
                preview_data.append({
                    "מספר חשבון": r.account_number,
                    "שם חשבון":   r.account_name,
                    "יתרה נוכחית": r.closing_balance,
                    "יתרה פתיחה":  r.opening_balance,
                    "סכום ח/ז":    r.selected_sum,
                    "תאריך תחילת חוב": (
                        r.debt_start_date.strftime("%d/%m/%Y")
                        if r.debt_start_date else ""
                    ),
                    "הערות": "; ".join(r.warnings[:1]) if r.warnings else "",
                })
            df = pd.DataFrame(preview_data)
            st.dataframe(df, use_container_width=True)

    except Exception as exc:
        log.exception("Unhandled error during processing")
        st.error(f"שגיאה לא צפויה: {exc}")
        st.exception(exc)

    finally:
        # Clean up temp PDF
        try:
            if "pdf_path" in locals() and os.path.exists(pdf_path):
                os.unlink(pdf_path)
        except Exception:
            pass
        # Clean up extracted ZIP
        if tmp_dir:
            cleanup_temp_dir(tmp_dir)


# ---------------------------------------------------------------------------
# Sidebar: help / field mapping notes
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("ℹ עזרה")
    st.markdown(
        """
**זרימת עבודה:**
1. העלה ZIP עם קבצי מבנה אחיד **או** הזן נתיב תיקייה
2. העלה PDF של דו"ח 331
3. הזן קוד כותרת (לדוגמה: 1342)
4. לחץ "אשר ועבד"
5. הורד את קובץ ה-Excel

---
**מבנה קובץ ה-Excel:**
- **ריכוז** – שורה לכל חשבון עם נוסחת SUM
- **תנועות נבחרות** – תנועות שנבחרו לנוסחה
- **כל התנועות** – ביקורת מלאה
- **לוג שגיאות** – אזהרות ושגיאות

---
**כיצד לשנות מיפוי שדות?**

ערוך את הקובץ:
`config/field_mappings.py`

שנה את המספרים ב-`B100_FIELDS`
ו-`C100_FIELDS` בהתאם למיקום
האמיתי של כל שדה בקובץ שלך.
        """
    )

    st.divider()
    log_records = get_accumulated_logs()
    if log_records:
        st.markdown(f"**לוג: {len(log_records)} רשומות**")
        errors = [r for r in log_records if r[0] in ("ERROR", "CRITICAL")]
        warnings = [r for r in log_records if r[0] == "WARNING"]
        if errors:
            st.error(f"{len(errors)} שגיאות")
        if warnings:
            st.warning(f"{len(warnings)} אזהרות")
