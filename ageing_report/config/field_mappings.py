"""
field_mappings.py
=================
Centralised field-layout definitions for BKMVDATA.TXT records.

All index numbers are 0-based column positions AFTER splitting a tab-delimited
line.  If your accounting software produces a slightly different column order,
edit only this file — the rest of the system reads through these symbols.

Reference: Israeli Tax Authority "Takken 7" uniform-structure specification.
           Verified against common Israeli accounting packages (Priority, Hashavshevet,
           Rivhit, and others).  Where packages differ, the most common convention
           is used and a NOTE is added.

HOW TO ADAPT
------------
1.  Run the app once and inspect the "לוג שגיאות" sheet — it reports every
    parsing warning including which field caused the problem.
2.  Open BKMVDATA.TXT in a text editor, look at a sample B100 or C100 line,
    count tab-separated columns (0-based), and update the relevant index below.
3.  Re-run; the rest of the code uses these mappings everywhere.
"""

# ---------------------------------------------------------------------------
# A100  –  File/period header record
# ---------------------------------------------------------------------------
# A100 is read only for metadata; it is not needed for account/movement data.
A100_FIELDS = {
    "record_type":        0,   # "A100"
    "file_type":          1,   # 1 = Initial file, 2 = Replacement, 3 = Supplement
    "software_id":        2,   # Software registration number
    "software_name":      3,   # Name of generating software (optional in some versions)
    "language":           4,   # 1 = Hebrew
    "accountant_id":      5,   # Licensed accountant registration number (optional)
    "tax_id":             6,   # Business VAT / tax ID (מספר עוסק)
    "company_name":       7,   # Company/business name (שם עסק)
    "period_start":       8,   # YYYYMMDD – first day of the reported period
    "period_end":         9,   # YYYYMMDD – last day of the reported period
    "creation_date":      10,  # YYYYMMDD – file creation date
    "compression":        11,  # 0 = uncompressed
    "b100_count":         12,  # Expected number of B100 records
    "c100_count":         13,  # Expected number of C100 records
    # Fields 14+ are software-specific extensions – ignored.
}

# ---------------------------------------------------------------------------
# B100  –  Account card record  (כרטיס חשבון)
# ---------------------------------------------------------------------------
B100_FIELDS = {
    "record_type":            0,   # "B100"

    # --- Identity ---
    "account_key":            1,   # Account number / key  (מפתח כרטיס)
                                   # ASSUMPTION: alphanumeric, max 15 chars.
                                   # Some software uses numeric-only keys.
    "account_name":           2,   # Display name in Hebrew (שם כרטיס)

    # --- Classification ---
    "sort_code":              3,   # Chart-of-accounts sort / classification code
                                   # This often matches the 331 header code.

    # --- Opening balance (פתיחה) ---
    "opening_balance_side":   4,   # "1" = debit (חובה), "2" = credit (זכות)
    "opening_balance":        5,   # Absolute value (positive number)
                                   # NOTE: Some packages store a SIGNED number here
                                   # (positive=debit, negative=credit) and leave
                                   # field 4 blank.  See _parse_balance() in bkmv_parser.py.

    # --- Closing balance (סגירה / יתרה) ---
    "closing_balance_side":   6,   # "1" = debit, "2" = credit
    "closing_balance":        7,   # Absolute value (positive number)

    # --- Foreign currency (optional) ---
    "currency_code":          8,   # ISO currency code, blank = ILS
    "opening_balance_fc":     9,   # Opening balance in foreign currency
    "closing_balance_fc":     10,  # Closing balance in foreign currency

    # --- Optional additional fields (present in some packages) ---
    # "account_type":         11,  # Account type code
    # "address":              12,  # Street address
    # "city":                 13,  # City
    # "phone":                14,  # Phone
}

# ---------------------------------------------------------------------------
# C100  –  Journal-entry movement record  (תנועת יומן)
# ---------------------------------------------------------------------------
C100_FIELDS = {
    "record_type":        0,    # "C100"

    # --- Entry identification ---
    "movement_type":      1,    # "1" = regular, "2" = opening-balance movement
    "entry_number":       2,    # Journal entry number (מספר תנועה / מספר רשומה)
                                # Used for stable sort on same-date movements.
    "line_number":        3,    # Line number within the entry (מספר שורה)
                                # Secondary sort key for stable ordering.

    # --- Dates ---
    "date":               4,    # Transaction date YYYYMMDD (תאריך ערך)
    "value_date":         5,    # Value date YYYYMMDD (תאריך ייחוס) – may equal date

    # --- Description / references ---
    "details":            6,    # Free-text description (פרטים)
    "reference1":         7,    # Reference 1 (אסמכתא 1) – e.g. invoice number
    "reference2":         8,    # Reference 2
    "reference3":         9,    # Reference 3

    # --- Amount ---
    "amount":             10,   # Movement amount in ILS (absolute, positive)
                                # ASSUMPTION: always positive; debit/credit
                                # is determined by which account field matches.

    # --- Account codes (double-entry sides) ---
    "debit_account":      11,   # Account that is DEBITED (חשבון חיוב)
    "credit_account":     12,   # Account that is CREDITED (חשבון זכות)

    # --- Foreign currency (optional) ---
    "currency_code":      13,   # ISO currency code, blank = ILS
    "foreign_amount":     14,   # Amount in foreign currency

    # --- Optional fields ---
    "quantity1":          15,   # Quantity 1 (כמות 1)
    "quantity2":          16,   # Quantity 2
    "branch":             17,   # Branch number (סניף)
    "cost_center":        18,   # Cost centre (מרכז רווח)
}

# ---------------------------------------------------------------------------
# Alternative B100 layout
# ---------------------------------------------------------------------------
# Some older packages (e.g. certain Hashavshevet exports) use a slightly
# different B100 layout where opening/closing balances are stored as a SINGLE
# signed number (positive = debit, negative = credit) without a separate side
# indicator.  Set USE_SIGNED_BALANCE_IN_B100 = True if that is your case.

USE_SIGNED_BALANCE_IN_B100 = False
# When True: field 4 = signed opening balance, field 5 = signed closing balance,
# and fields 6/7 are shifted.
B100_FIELDS_SIGNED = {
    "record_type":        0,
    "account_key":        1,
    "account_name":       2,
    "sort_code":          3,
    "opening_balance":    4,    # Signed: positive=debit, negative=credit
    "closing_balance":    5,    # Signed
    "currency_code":      6,
    "opening_balance_fc": 7,
    "closing_balance_fc": 8,
}
