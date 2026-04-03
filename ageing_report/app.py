"""
app.py - ממשק Streamlit לדוח גיול לקוחות
==========================================
"""

import re
import streamlit as st
import pandas as pd

from ageing_report.parsers.pdf_parser import parse_pdf_section
from ageing_report.parsers.bkmv_parser import parse_b11_records, parse_b1_transactions
from ageing_report.processors.aging_calculator import process_accounts
from ageing_report.excel.excel_generator import generate_excel
from ageing_report.utils.file_utils import extract_zip, find_bkmvdata, cleanup_temp_dir

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

# --- שדות קלט ---
col1, col2, col3 = st.columns(3)

with col1:
    zip_file = st.file_uploader("קובץ מבנה אחיד (ZIP)", type=["zip"])

with col2:
    pdf_file = st.file_uploader("דוח 331 (PDF)", type=["pdf"])

with col3:
    section_code = st.text_input("קוד סעיף", value="1342", help="לדוגמה: 1342, 1302, 2660")

st.markdown("---")

# --- כפתור הפקה ---
if st.button("הפק דוח", type="primary", use_container_width=True):
    if not zip_file:
        st.error("יש להעלות קובץ מבנה אחיד (ZIP)")
    elif not pdf_file:
        st.error("יש להעלות דוח 331 (PDF)")
    elif not section_code.strip():
        st.error("יש להזין קוד סעיף")
    else:
        temp_dir = None
        try:
            progress = st.progress(0)
            status = st.empty()

            # שלב 1: PDF
            status.info("קורא דוח 331...")
            progress.progress(10)
            pdf_accounts = parse_pdf_section(pdf_file, section_code.strip())
            if not pdf_accounts:
                st.error(f"לא נמצאו חשבונות בסעיף {section_code}")
                st.stop()
            progress.progress(25)
            target_accounts = set(pdf_accounts.keys())

            # שלב 2: חילוץ ZIP
            status.info("מחלץ קובץ מבנה אחיד...")
            progress.progress(30)
            temp_dir = extract_zip(zip_file)
            bkmv_path = find_bkmvdata(temp_dir)
            if not bkmv_path:
                st.error("לא נמצא קובץ BKMVDATA.TXT בתוך ה-ZIP")
                st.stop()
            progress.progress(35)

            # שלב 3: B11
            status.info("קורא כרטיסיות חשבון...")
            b11_data = parse_b11_records(bkmv_path, target_accounts)
            progress.progress(50)

            # שלב 4: B1
            status.info("קורא תנועות יומן...")
            transactions = parse_b1_transactions(bkmv_path, target_accounts)
            progress.progress(65)

            # שלב 5: חישוב גיול
            status.info("מחשב גיול...")
            results = process_accounts(pdf_accounts, b11_data, transactions)
            progress.progress(80)

            # שלב 6: Excel
            status.info("מייצר קובץ Excel...")
            excel_bytes = generate_excel(results)
            progress.progress(100)
            status.success("הדוח מוכן!")

            # --- סיכום ---
            st.markdown("### סיכום")
            debit_count = sum(1 for r in results if r['closing'] > 0.005)
            credit_count = sum(1 for r in results if r['closing'] < -0.005)
            with_date = sum(1 for r in results if re.match(r'\d{2}/\d{2}/\d{4}', r['debt_start_date']))
            total_closing = sum(r['closing'] for r in results)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("חשבונות", len(results))
            c2.metric("יתרות חובה", debit_count)
            c3.metric("יתרות זכות", credit_count)
            c4.metric("סה\"כ יתרות", f"{total_closing:,.2f}")

            # --- תצוגה מקדימה ---
            st.markdown("### תצוגה מקדימה")
            df = pd.DataFrame(results)
            df.columns = ['מספר חשבון', 'שם חשבון', 'יתרה נוכחית', 'יתרת פתיחה', 'תאריך תחילת חוב', 'פירוט סכימה']
            st.dataframe(df, use_container_width=True, hide_index=True)

            # --- הורדה ---
            st.download_button(
                label="הורד קובץ Excel",
                data=excel_bytes,
                file_name=f"דוח_גיול_{section_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"שגיאה: {e}")
        finally:
            if temp_dir:
                cleanup_temp_dir(temp_dir)
