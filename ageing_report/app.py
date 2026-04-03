"""
app.py - ממשק Streamlit לדוח גיול לקוחות
==========================================
"""

import re
import tempfile
import os
import streamlit as st
import pandas as pd

from ageing_report.parsers.pdf_parser import parse_pdf_section
from ageing_report.parsers.bkmv_parser import parse_b11_records, parse_b1_transactions
from ageing_report.processors.aging_calculator import process_accounts
from ageing_report.excel.excel_generator import generate_excel
from ageing_report.utils.file_utils import extract_zip, find_bkmvdata, cleanup_temp_dir
from ageing_report.config.constants import ZERO_THRESHOLD

# --- הגדרת עמוד ---
st.set_page_config(
    page_title="דוח גיול לקוחות",
    page_icon="📊",
    layout="wide",
)

# --- RTL CSS ---
st.markdown("""
<style>
    .stApp { direction: rtl; }
    .stMarkdown, .stText, .stAlert { direction: rtl; text-align: right; }
    h1, h2, h3 { text-align: center; }
    .stDataFrame { direction: ltr; }
</style>
""", unsafe_allow_html=True)

st.title("דוח גיול לקוחות")
st.markdown("---")


# --- פונקציות עם cache ---

@st.cache_data(show_spinner=False)
def cached_parse_pdf(pdf_bytes, section_code):
    """פרסור PDF עם cache לפי תוכן הקובץ וקוד הסעיף."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        return parse_pdf_section(tmp_path, section_code)
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def cached_parse_b11(bkmv_bytes, bkmv_name, target_accounts_tuple):
    """פרסור B11 עם cache."""
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(bkmv_bytes)
        tmp_path = tmp.name
    try:
        return parse_b11_records(tmp_path, set(target_accounts_tuple))
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def cached_parse_b1(bkmv_bytes, bkmv_name, target_accounts_tuple):
    """פרסור B1 עם cache."""
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(bkmv_bytes)
        tmp_path = tmp.name
    try:
        return parse_b1_transactions(tmp_path, set(target_accounts_tuple))
    finally:
        os.unlink(tmp_path)


# --- שדות קלט ---
col1, col2 = st.columns(2)

with col1:
    zip_file = st.file_uploader("קובץ מבנה אחיד (ZIP)", type=["zip"],
                                 help="קובץ ZIP המכיל BKMVDATA.TXT. אם ה-ZIP מכיל גם PDF של דוח 331, הוא יזוהה אוטומטית.")

with col2:
    pdf_file = st.file_uploader("דוח 331 (PDF)", type=["pdf"],
                                 help="אופציונלי אם ה-PDF נמצא בתוך ה-ZIP")

# שמירת קוד סעיף ב-session_state
if 'section_code' not in st.session_state:
    st.session_state.section_code = "1342"

section_code_input = st.text_input(
    "קוד/י סעיף (ניתן להזין מספר קודים מופרדים בפסיק)",
    value=st.session_state.section_code,
    help="לדוגמה: 1342 או 1342,1302,2660",
)
st.session_state.section_code = section_code_input

st.markdown("---")

# --- כפתור הפקה ---
if st.button("הפק דוח", type="primary", use_container_width=True):
    if not zip_file:
        st.error("יש להעלות קובץ מבנה אחיד (ZIP)")
    elif not section_code_input.strip():
        st.error("יש להזין קוד סעיף")
    else:
        # פרסור קודי סעיפים
        section_codes = [s.strip() for s in section_code_input.split(',') if s.strip()]
        if not section_codes:
            st.error("יש להזין לפחות קוד סעיף אחד")
            st.stop()

        temp_dir = None
        try:
            progress = st.progress(0)
            status = st.empty()

            # שלב 1: חילוץ ZIP
            status.info("מחלץ קובץ מבנה אחיד...")
            progress.progress(5)
            temp_dir, pdf_inside_zip = extract_zip(zip_file)
            bkmv_path = find_bkmvdata(temp_dir)
            if not bkmv_path:
                st.error("לא נמצא קובץ BKMVDATA.TXT בתוך ה-ZIP")
                st.stop()

            # קביעת מקור ה-PDF
            pdf_source = None
            if pdf_file:
                pdf_source = pdf_file
            elif pdf_inside_zip:
                pdf_source = pdf_inside_zip
                st.info(f"נמצא PDF בתוך ה-ZIP: {os.path.basename(pdf_inside_zip)}")
            else:
                st.error("לא הועלה דוח 331 ולא נמצא PDF בתוך ה-ZIP")
                st.stop()

            progress.progress(10)

            # קריאת BKMVDATA לזיכרון (עבור cache)
            with open(bkmv_path, 'rb') as f:
                bkmv_bytes = f.read()
            bkmv_name = os.path.basename(bkmv_path)

            # קריאת PDF לזיכרון (עבור cache)
            if hasattr(pdf_source, 'read'):
                pdf_bytes = pdf_source.read()
                pdf_source.seek(0)
            else:
                with open(pdf_source, 'rb') as f:
                    pdf_bytes = f.read()

            # שלב 2: PDF - כל הסעיפים
            all_pdf_accounts = {}
            section_totals = {}
            step_pct = 20 // max(len(section_codes), 1)

            for i, code in enumerate(section_codes):
                status.info(f"קורא דוח 331 - סעיף {code}...")
                progress.progress(10 + (i + 1) * step_pct)
                try:
                    accounts, total = cached_parse_pdf(pdf_bytes, code)
                except ValueError as e:
                    st.error(str(e))
                    st.stop()

                if not accounts:
                    st.warning(f"לא נמצאו חשבונות בסעיף {code}")
                    continue
                all_pdf_accounts.update(accounts)
                section_totals[code] = total

            if not all_pdf_accounts:
                st.error("לא נמצאו חשבונות באף סעיף שהוזן")
                st.stop()

            progress.progress(30)
            target_accounts = set(all_pdf_accounts.keys())
            target_tuple = tuple(sorted(target_accounts))

            # שלב 3: B11
            status.info("קורא כרטיסיות חשבון...")
            b11_data = cached_parse_b11(bkmv_bytes, bkmv_name, target_tuple)
            progress.progress(50)

            # שלב 4: B1
            status.info("קורא תנועות יומן...")
            transactions = cached_parse_b1(bkmv_bytes, bkmv_name, target_tuple)
            progress.progress(65)

            # שלב 5: חישוב גיול
            status.info("מחשב גיול...")
            results = process_accounts(all_pdf_accounts, b11_data, transactions)
            progress.progress(80)

            # שלב 6: Excel
            status.info("מייצר קובץ Excel...")
            excel_bytes = generate_excel(results)
            progress.progress(100)
            status.success("הדוח מוכן!")

            # --- אימות מול PDF ---
            if section_totals:
                st.markdown("### אימות מול דוח 331")
                total_closing = sum(r['closing'] for r in results)
                for code, pdf_total in section_totals.items():
                    if pdf_total is not None:
                        diff = abs(total_closing - pdf_total)
                        if diff < 0.1:
                            st.success(f"סעיף {code}: סה\"כ {total_closing:,.2f} = סה\"כ PDF {pdf_total:,.2f} ✓")
                        else:
                            st.warning(f"סעיף {code}: סה\"כ מחושב {total_closing:,.2f} ≠ סה\"כ PDF {pdf_total:,.2f} (הפרש: {diff:,.2f})")

            # --- סיכום ---
            st.markdown("### סיכום")
            debit_count = sum(1 for r in results if r['closing'] > ZERO_THRESHOLD)
            credit_count = sum(1 for r in results if r['closing'] < -ZERO_THRESHOLD)
            zero_count = sum(1 for r in results if abs(r['closing']) <= ZERO_THRESHOLD)
            with_date = sum(1 for r in results if re.match(r'\d{2}/\d{2}/\d{4}', r['debt_start_date']))
            with_opening = sum(1 for r in results if r['debt_start_date'] == 'כולל יתרת פתיחה')
            no_date = sum(1 for r in results if r['debt_start_date'] in ('', 'מורכב ממספר יתרות'))
            total_closing = sum(r['closing'] for r in results)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("חשבונות", len(results))
            c2.metric("יתרות חובה", debit_count)
            c3.metric("יתרות זכות", credit_count)
            c4.metric("סה\"כ יתרות", f"{total_closing:,.2f}")

            # --- חשבונות ללא תאריך מדויק ---
            if with_opening > 0 or no_date > 0:
                st.markdown("### חשבונות ללא תאריך מדויק")
                cols = st.columns(3)
                cols[0].metric("עם תאריך מדויק", with_date)
                cols[1].metric("כולל יתרת פתיחה", with_opening)
                cols[2].metric("ללא תאריך / מורכב", no_date)

                # הצגת הפירוט
                problem_accounts = [
                    r for r in results
                    if r['debt_start_date'] in ('כולל יתרת פתיחה', 'מורכב ממספר יתרות', '')
                    and abs(r['closing']) > ZERO_THRESHOLD
                ]
                if problem_accounts:
                    with st.expander(f"הצג {len(problem_accounts)} חשבונות ללא תאריך מדויק"):
                        df_prob = pd.DataFrame(problem_accounts)
                        df_prob = df_prob[['acct_num', 'name', 'closing', 'debt_start_date']]
                        df_prob.columns = ['מספר חשבון', 'שם חשבון', 'יתרה', 'סטטוס']
                        st.dataframe(df_prob, use_container_width=True, hide_index=True)

            # --- תצוגה מקדימה ---
            st.markdown("### תצוגה מקדימה")
            df = pd.DataFrame(results)
            df.columns = ['מספר חשבון', 'שם חשבון', 'יתרה נוכחית', 'יתרת פתיחה', 'תאריך תחילת חוב', 'פירוט סכימה']
            st.dataframe(df, use_container_width=True, hide_index=True)

            # --- הורדה ---
            codes_str = "_".join(section_codes)
            st.download_button(
                label="הורד קובץ Excel",
                data=excel_bytes,
                file_name=f"דוח_גיול_{codes_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"שגיאה לא צפויה: {e}")
        finally:
            if temp_dir:
                cleanup_temp_dir(temp_dir)
