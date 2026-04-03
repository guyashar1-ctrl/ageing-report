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
from ageing_report.config.constants import ZERO_THRESHOLD, SECTION_NAMES

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="דוח גיול לקוחות",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────
# Custom CSS - polished RTL design
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* RTL base */
    .stApp { direction: rtl; }
    .stMarkdown, .stText, .stAlert { direction: rtl; text-align: right; }
    .stDataFrame { direction: ltr; }

    /* Hero header */
    .hero {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6da8 50%, #3a8fd4 100%);
        border-radius: 16px;
        padding: 2.5rem 2rem;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    .hero h1 {
        color: white;
        font-size: 2.4rem;
        margin: 0 0 0.5rem 0;
        font-weight: 700;
    }
    .hero p {
        color: rgba(255,255,255,0.85);
        font-size: 1.1rem;
        margin: 0;
    }

    /* Step cards */
    .step-card {
        background: #f8fafc;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: border-color 0.2s, box-shadow 0.2s;
    }
    .step-card:hover {
        border-color: #3a8fd4;
        box-shadow: 0 2px 12px rgba(58,143,212,0.15);
    }
    .step-card.active {
        border-color: #2d6da8;
        background: #f0f7ff;
    }
    .step-card.done {
        border-color: #22c55e;
        background: #f0fdf4;
    }
    .step-number {
        display: inline-block;
        width: 32px;
        height: 32px;
        line-height: 32px;
        text-align: center;
        border-radius: 50%;
        background: #2d6da8;
        color: white;
        font-weight: 700;
        font-size: 0.9rem;
        margin-left: 0.75rem;
    }
    .step-number.done {
        background: #22c55e;
    }
    .step-title {
        display: inline;
        font-size: 1.15rem;
        font-weight: 600;
        color: #1e293b;
    }
    .step-desc {
        color: #64748b;
        font-size: 0.9rem;
        margin-top: 0.5rem;
        margin-right: 2.75rem;
    }

    /* Stat cards */
    .stat-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 1rem;
        margin: 1rem 0;
    }
    .stat-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    .stat-card .value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e293b;
        line-height: 1.2;
    }
    .stat-card .label {
        font-size: 0.85rem;
        color: #64748b;
        margin-top: 0.3rem;
    }
    .stat-card.blue .value { color: #2d6da8; }
    .stat-card.green .value { color: #16a34a; }
    .stat-card.red .value { color: #dc2626; }
    .stat-card.amber .value { color: #d97706; }

    /* Validation badge */
    .validation-pass {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        color: #166534;
        font-weight: 500;
        margin: 0.5rem 0;
    }
    .validation-fail {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        color: #92400e;
        font-weight: 500;
        margin: 0.5rem 0;
    }

    /* Section pills */
    .section-pills {
        display: flex;
        gap: 0.5rem;
        flex-wrap: wrap;
        justify-content: center;
        margin: 0.75rem 0;
    }
    .section-pill {
        background: #e0f2fe;
        color: #0369a1;
        border-radius: 20px;
        padding: 0.3rem 1rem;
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* Process timeline */
    .timeline-step {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding: 0.5rem 0;
        font-size: 0.95rem;
    }
    .timeline-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
        background: #cbd5e1;
        flex-shrink: 0;
    }
    .timeline-dot.active {
        background: #3a8fd4;
        box-shadow: 0 0 0 3px rgba(58,143,212,0.3);
    }
    .timeline-dot.done {
        background: #22c55e;
    }

    /* Download area */
    .download-area {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        border: 2px solid #86efac;
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        margin: 1.5rem 0;
    }
    .download-area h3 {
        color: #166534;
        margin: 0 0 0.5rem 0;
    }
    .download-area p {
        color: #15803d;
        margin: 0 0 1rem 0;
    }

    /* File upload styling */
    [data-testid="stFileUploader"] {
        border-radius: 12px;
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.5rem;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Cached parsing functions
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def cached_parse_pdf(pdf_bytes, section_code):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        return parse_pdf_section(tmp_path, section_code)
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def cached_parse_b11(bkmv_bytes, bkmv_name, target_accounts_tuple):
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(bkmv_bytes)
        tmp_path = tmp.name
    try:
        return parse_b11_records(tmp_path, set(target_accounts_tuple))
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def cached_parse_b1(bkmv_bytes, bkmv_name, target_accounts_tuple):
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(bkmv_bytes)
        tmp_path = tmp.name
    try:
        return parse_b1_transactions(tmp_path, set(target_accounts_tuple))
    finally:
        os.unlink(tmp_path)


# ─────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────

def render_stat_card(value, label, color="blue"):
    return f'<div class="stat-card {color}"><div class="value">{value}</div><div class="label">{label}</div></div>'


def get_step_state(step_num):
    """Get visual state for a step: done / active / pending."""
    current = st.session_state.get('current_step', 0)
    if step_num < current:
        return "done"
    elif step_num == current:
        return "active"
    return ""


# ─────────────────────────────────────────────
# Hero Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>📊 דוח גיול לקוחות</h1>
    <p>חילוץ אוטומטי מדוח 331 ומבנה אחיד &bull; חישוב תאריך תחילת חוב &bull; הורדת Excel מעוצב</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Initialize session state
# ─────────────────────────────────────────────
if 'section_code' not in st.session_state:
    st.session_state.section_code = "1342"
if 'results' not in st.session_state:
    st.session_state.results = None
if 'excel_bytes' not in st.session_state:
    st.session_state.excel_bytes = None
if 'section_totals' not in st.session_state:
    st.session_state.section_totals = {}
if 'section_codes_used' not in st.session_state:
    st.session_state.section_codes_used = []


# ─────────────────────────────────────────────
# Input Section - Step Cards
# ─────────────────────────────────────────────

col_input, col_preview = st.columns([1, 1], gap="large")

with col_input:

    # Step 1: ZIP Upload
    st.markdown("""
    <div class="step-card">
        <span class="step-number">1</span>
        <span class="step-title">העלאת קובץ מבנה אחיד</span>
        <div class="step-desc">קובץ ZIP המכיל BKMVDATA.TXT &mdash; אם ה-ZIP מכיל גם PDF, הוא יזוהה אוטומטית</div>
    </div>
    """, unsafe_allow_html=True)

    zip_file = st.file_uploader(
        "ZIP file", type=["zip"],
        label_visibility="collapsed",
        key="zip_uploader",
    )
    if zip_file:
        st.success(f"✓ {zip_file.name} ({zip_file.size / 1024 / 1024:.1f} MB)")

    # Step 2: PDF Upload
    st.markdown("""
    <div class="step-card">
        <span class="step-number">2</span>
        <span class="step-title">העלאת דוח 331</span>
        <div class="step-desc">אופציונלי אם ה-PDF נמצא בתוך ה-ZIP</div>
    </div>
    """, unsafe_allow_html=True)

    pdf_file = st.file_uploader(
        "PDF file", type=["pdf"],
        label_visibility="collapsed",
        key="pdf_uploader",
    )
    if pdf_file:
        st.success(f"✓ {pdf_file.name} ({pdf_file.size / 1024:.0f} KB)")

    # Step 3: Section Code
    st.markdown("""
    <div class="step-card">
        <span class="step-number">3</span>
        <span class="step-title">בחירת קוד סעיף</span>
        <div class="step-desc">ניתן להזין מספר קודים מופרדים בפסיק</div>
    </div>
    """, unsafe_allow_html=True)

    # Quick-pick buttons for common sections
    section_code_input = st.text_input(
        "קוד סעיף",
        value=st.session_state.section_code,
        label_visibility="collapsed",
        placeholder="הזן קוד סעיף...",
    )
    st.session_state.section_code = section_code_input

    # Quick section pills
    quick_cols = st.columns(len(SECTION_NAMES))
    for i, (code, _name) in enumerate(SECTION_NAMES.items()):
        with quick_cols[i]:
            if st.button(code, key=f"quick_{code}", use_container_width=True):
                current = st.session_state.section_code.strip()
                if current and code not in current.split(','):
                    st.session_state.section_code = f"{current},{code}"
                else:
                    st.session_state.section_code = code
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Generate button
    can_generate = zip_file is not None and section_code_input.strip() != ""
    generate_clicked = st.button(
        "🚀 הפק דוח גיול",
        type="primary",
        use_container_width=True,
        disabled=not can_generate,
    )

    if not can_generate:
        missing = []
        if not zip_file:
            missing.append("קובץ ZIP")
        if not section_code_input.strip():
            missing.append("קוד סעיף")
        st.caption(f"חסר: {', '.join(missing)}")

with col_preview:
    # Right side - shows status/results
    if st.session_state.results is None and not generate_clicked:
        # Empty state
        st.markdown("""
        <div style="
            background: #f8fafc;
            border: 2px dashed #cbd5e1;
            border-radius: 16px;
            padding: 4rem 2rem;
            text-align: center;
            margin-top: 1rem;
        ">
            <div style="font-size: 3rem; margin-bottom: 1rem;">📋</div>
            <h3 style="color: #475569; margin: 0 0 0.5rem 0;">התוצאות יופיעו כאן</h3>
            <p style="color: #94a3b8; margin: 0;">העלו קבצים ולחצו "הפק דוח גיול"</p>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Processing
# ─────────────────────────────────────────────

if generate_clicked:
    section_codes = [s.strip() for s in section_code_input.split(',') if s.strip()]
    if not section_codes:
        st.error("יש להזין לפחות קוד סעיף אחד")
        st.stop()

    temp_dir = None
    try:
        # Processing container in the right column
        with col_preview:
            progress_container = st.container()
            with progress_container:
                st.markdown("### ⏳ מעבד נתונים...")
                progress = st.progress(0)
                log_area = st.empty()

                steps_status = {
                    'zip': '⏳',
                    'pdf': '⏳',
                    'b11': '⏳',
                    'b1': '⏳',
                    'aging': '⏳',
                    'excel': '⏳',
                }

                def update_log():
                    lines = []
                    labels = {
                        'zip': 'חילוץ ZIP ואיתור קבצים',
                        'pdf': 'קריאת דוח 331',
                        'b11': 'קריאת כרטיסיות חשבון (B11)',
                        'b1': 'קריאת תנועות יומן (B1)',
                        'aging': 'חישוב גיול',
                        'excel': 'יצירת קובץ Excel',
                    }
                    for key, label in labels.items():
                        icon = steps_status[key]
                        lines.append(f"{icon} {label}")
                    log_area.markdown("\n\n".join(lines))

                update_log()

                # Step 1: Extract ZIP
                steps_status['zip'] = '🔄'
                update_log()
                progress.progress(5)
                temp_dir, pdf_inside_zip = extract_zip(zip_file)
                bkmv_path = find_bkmvdata(temp_dir)
                if not bkmv_path:
                    st.error("❌ לא נמצא קובץ BKMVDATA.TXT בתוך ה-ZIP")
                    st.stop()

                pdf_source = None
                if pdf_file:
                    pdf_source = pdf_file
                elif pdf_inside_zip:
                    pdf_source = pdf_inside_zip
                else:
                    st.error("❌ לא הועלה דוח 331 ולא נמצא PDF בתוך ה-ZIP")
                    st.stop()

                steps_status['zip'] = '✅'
                update_log()
                progress.progress(10)

                # Read files into memory for caching
                with open(bkmv_path, 'rb') as f:
                    bkmv_bytes = f.read()
                bkmv_name = os.path.basename(bkmv_path)

                if hasattr(pdf_source, 'read'):
                    pdf_bytes = pdf_source.read()
                    pdf_source.seek(0)
                else:
                    with open(pdf_source, 'rb') as f:
                        pdf_bytes = f.read()

                # Step 2: Parse PDF
                steps_status['pdf'] = '🔄'
                update_log()

                all_pdf_accounts = {}
                section_totals = {}

                for i, code in enumerate(section_codes):
                    progress.progress(10 + int(20 * (i + 1) / len(section_codes)))
                    try:
                        accounts, total = cached_parse_pdf(pdf_bytes, code)
                    except ValueError as e:
                        st.error(str(e))
                        st.stop()
                    if accounts:
                        all_pdf_accounts.update(accounts)
                        section_totals[code] = total

                if not all_pdf_accounts:
                    steps_status['pdf'] = '❌'
                    update_log()
                    st.error("לא נמצאו חשבונות באף סעיף שהוזן")
                    st.stop()

                steps_status['pdf'] = '✅'
                update_log()
                progress.progress(30)

                target_accounts = set(all_pdf_accounts.keys())
                target_tuple = tuple(sorted(target_accounts))

                # Step 3: B11
                steps_status['b11'] = '🔄'
                update_log()
                b11_data = cached_parse_b11(bkmv_bytes, bkmv_name, target_tuple)
                steps_status['b11'] = '✅'
                update_log()
                progress.progress(50)

                # Step 4: B1
                steps_status['b1'] = '🔄'
                update_log()
                transactions = cached_parse_b1(bkmv_bytes, bkmv_name, target_tuple)
                steps_status['b1'] = '✅'
                update_log()
                progress.progress(65)

                # Step 5: Aging
                steps_status['aging'] = '🔄'
                update_log()
                results = process_accounts(all_pdf_accounts, b11_data, transactions)
                steps_status['aging'] = '✅'
                update_log()
                progress.progress(80)

                # Step 6: Excel
                steps_status['excel'] = '🔄'
                update_log()
                excel_bytes = generate_excel(results)
                steps_status['excel'] = '✅'
                update_log()
                progress.progress(100)

                # Save to session state
                st.session_state.results = results
                st.session_state.excel_bytes = excel_bytes
                st.session_state.section_totals = section_totals
                st.session_state.section_codes_used = section_codes

                st.rerun()

    except ValueError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"שגיאה לא צפויה: {e}")
    finally:
        if temp_dir:
            cleanup_temp_dir(temp_dir)


# ─────────────────────────────────────────────
# Results Display (persistent via session state)
# ─────────────────────────────────────────────

if st.session_state.results is not None:
    results = st.session_state.results
    excel_bytes = st.session_state.excel_bytes
    section_totals = st.session_state.section_totals
    section_codes = st.session_state.section_codes_used

    with col_preview:
        st.markdown("### ✅ הדוח מוכן!")

    # Compute stats
    debit_count = sum(1 for r in results if r['closing'] > ZERO_THRESHOLD)
    credit_count = sum(1 for r in results if r['closing'] < -ZERO_THRESHOLD)
    zero_count = sum(1 for r in results if abs(r['closing']) <= ZERO_THRESHOLD)
    with_date = sum(1 for r in results if re.match(r'\d{2}/\d{2}/\d{4}', r['debt_start_date']))
    with_opening = sum(1 for r in results if r['debt_start_date'] == 'כולל יתרת פתיחה')
    no_date = sum(1 for r in results if r['debt_start_date'] in ('', 'מורכב ממספר יתרות'))
    total_closing = sum(r['closing'] for r in results)

    # ── Stats Dashboard ──
    st.markdown("---")

    st.markdown(f"""
    <div class="stat-grid">
        {render_stat_card(len(results), "סה״כ חשבונות", "blue")}
        {render_stat_card(debit_count, "יתרות חובה", "red")}
        {render_stat_card(credit_count, "יתרות זכות", "green")}
        {render_stat_card(zero_count, "יתרות אפס", "amber")}
        {render_stat_card(f"{total_closing:,.2f}", "סה״כ יתרות", "blue")}
        {render_stat_card(with_date, "עם תאריך מדויק", "green")}
        {render_stat_card(with_opening, "כולל יתרת פתיחה", "amber")}
        {render_stat_card(no_date, "ללא תאריך", "red")}
    </div>
    """, unsafe_allow_html=True)

    # ── Validation ──
    if section_totals:
        for code, pdf_total in section_totals.items():
            if pdf_total is not None:
                diff = abs(total_closing - pdf_total)
                if diff < 0.1:
                    st.markdown(f"""
                    <div class="validation-pass">
                        ✅ <strong>אימות סעיף {code}:</strong> סה"כ מחושב {total_closing:,.2f} = סה"כ PDF {pdf_total:,.2f}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="validation-fail">
                        ⚠️ <strong>סעיף {code}:</strong> סה"כ מחושב {total_closing:,.2f} ≠ סה"כ PDF {pdf_total:,.2f} (הפרש: {diff:,.2f})
                    </div>
                    """, unsafe_allow_html=True)

    # ── Tabs for data views ──
    st.markdown("---")

    tab_all, tab_debit, tab_credit, tab_problems = st.tabs([
        f"📋 כל החשבונות ({len(results)})",
        f"🔴 יתרות חובה ({debit_count})",
        f"🟢 יתרות זכות ({credit_count})",
        f"⚠️ ללא תאריך ({with_opening + no_date})",
    ])

    df_all = pd.DataFrame(results)
    display_cols = {
        'acct_num': 'מספר חשבון',
        'name': 'שם חשבון',
        'closing': 'יתרה נוכחית',
        'opening': 'יתרת פתיחה',
        'debt_start_date': 'תאריך תחילת חוב',
        'sum_formula': 'פירוט סכימה',
    }
    df_display = df_all.rename(columns=display_cols)

    column_config = {
        "יתרה נוכחית": st.column_config.NumberColumn(format="%.2f"),
        "יתרת פתיחה": st.column_config.NumberColumn(format="%.2f"),
    }

    with tab_all:
        # Search filter
        search = st.text_input("🔍 חיפוש לפי שם או מספר חשבון", key="search_all", placeholder="הקלד לחיפוש...")
        filtered = df_display
        if search:
            mask = (
                df_display['שם חשבון'].str.contains(search, na=False) |
                df_display['מספר חשבון'].astype(str).str.contains(search, na=False)
            )
            filtered = df_display[mask]
            st.caption(f"נמצאו {len(filtered)} תוצאות")

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config=column_config,
        )

    with tab_debit:
        df_debit = df_display[df_all['closing'] > ZERO_THRESHOLD]
        search_d = st.text_input("🔍 חיפוש", key="search_debit", placeholder="הקלד לחיפוש...")
        if search_d:
            mask = (
                df_debit['שם חשבון'].str.contains(search_d, na=False) |
                df_debit['מספר חשבון'].astype(str).str.contains(search_d, na=False)
            )
            df_debit = df_debit[mask]
        st.dataframe(df_debit, use_container_width=True, hide_index=True, height=500, column_config=column_config)

    with tab_credit:
        df_credit = df_display[df_all['closing'] < -ZERO_THRESHOLD]
        search_c = st.text_input("🔍 חיפוש", key="search_credit", placeholder="הקלד לחיפוש...")
        if search_c:
            mask = (
                df_credit['שם חשבון'].str.contains(search_c, na=False) |
                df_credit['מספר חשבון'].astype(str).str.contains(search_c, na=False)
            )
            df_credit = df_credit[mask]
        st.dataframe(df_credit, use_container_width=True, hide_index=True, height=500, column_config=column_config)

    with tab_problems:
        problem_mask = df_all['debt_start_date'].isin(['כולל יתרת פתיחה', 'מורכב ממספר יתרות', ''])
        balance_mask = df_all['closing'].abs() > ZERO_THRESHOLD
        df_problems = df_display[problem_mask & balance_mask]

        if len(df_problems) == 0:
            st.success("כל החשבונות עם יתרה פעילה קיבלו תאריך מדויק!")
        else:
            st.caption(f"{len(df_problems)} חשבונות דורשים בדיקה ידנית")
            st.dataframe(
                df_problems[['מספר חשבון', 'שם חשבון', 'יתרה נוכחית', 'תאריך תחילת חוב']],
                use_container_width=True,
                hide_index=True,
                height=400,
                column_config=column_config,
            )

    # ── Download Section ──
    st.markdown("---")
    codes_str = "_".join(section_codes)

    st.markdown("""
    <div class="download-area">
        <h3>📥 הורדת הדוח</h3>
        <p>קובץ Excel מעוצב עם כל הנתונים, מוכן לשימוש</p>
    </div>
    """, unsafe_allow_html=True)

    dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
    with dl_col2:
        st.download_button(
            label=f"⬇️  הורד דוח גיול - {codes_str}.xlsx",
            data=excel_bytes,
            file_name=f"דוח_גיול_{codes_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    # Reset button
    st.markdown("<br>", unsafe_allow_html=True)
    reset_col1, reset_col2, reset_col3 = st.columns([1, 2, 1])
    with reset_col2:
        if st.button("🔄 הפק דוח חדש", use_container_width=True):
            st.session_state.results = None
            st.session_state.excel_bytes = None
            st.session_state.section_totals = {}
            st.session_state.section_codes_used = []
            st.rerun()
