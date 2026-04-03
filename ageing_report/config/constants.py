"""
constants.py
============
Global constants for the Ageing Report system.

These values should not need to change between runs, but are centralized here
so that any adjustment is made in a single location.
"""

# ---------------------------------------------------------------------------
# BKMVDATA record-type identifiers
# ---------------------------------------------------------------------------
RECORD_TYPE_FILE_HEADER = "A100"   # Company / period header record
RECORD_TYPE_ACCOUNT = "B100"       # Account card (כרטיס חשבון)
RECORD_TYPE_MOVEMENT = "C100"      # Journal movement (תנועה)
RECORD_TYPE_ASSET = "D110"         # Fixed asset (רכוש קבוע)
RECORD_TYPE_INVENTORY = "M100"     # Inventory (מלאי)

# ---------------------------------------------------------------------------
# Balance-side codes used inside B100 records
# 1 = Debit (חובה), 2 = Credit (זכות)
# ---------------------------------------------------------------------------
BALANCE_SIDE_DEBIT = "1"
BALANCE_SIDE_CREDIT = "2"

# Hebrew labels used in output
DEBIT_LABEL = "ח"    # חובה
CREDIT_LABEL = "ז"   # זכות

# ---------------------------------------------------------------------------
# Movement-type codes inside C100 records
# 1 = Regular movement, 2 = Opening balance movement
# ---------------------------------------------------------------------------
MOVEMENT_TYPE_REGULAR = "1"
MOVEMENT_TYPE_OPENING = "2"

# ---------------------------------------------------------------------------
# Date & year settings
# ---------------------------------------------------------------------------
DATE_FORMAT_BKMV = "%Y%m%d"          # YYYYMMDD as used in BKMVDATA
DATE_FORMAT_DISPLAY = "%d/%m/%Y"     # DD/MM/YYYY for display in Hebrew Excel
TARGET_YEAR = 2025                   # The fiscal year we are analysing
PERIOD_START = "20250101"
PERIOD_END   = "20251231"

# ---------------------------------------------------------------------------
# File detection
# ---------------------------------------------------------------------------
# Primary names first, then TXT.BKMVDATA variant (some Israeli software uses it)
BKMVDATA_FILENAMES = [
    "BKMVDATA.TXT",
    "bkmvdata.txt",
    "TXT.BKMVDATA",
    "txt.bkmvdata",
    "BKMVDATA",
]
INI_FILENAMES = [
    "INI.TXT",
    "ini.txt",
    "TXT.INI",
    "txt.ini",
]

# ---------------------------------------------------------------------------
# Encoding detection order (Israeli Hebrew files use Windows-1255 most often)
# ---------------------------------------------------------------------------
ENCODINGS_TO_TRY = [
    "windows-1255",
    "cp1255",
    "iso-8859-8",
    "utf-8-sig",    # UTF-8 with BOM
    "utf-8",
    "latin-1",      # fallback – never fails but may mis-decode Hebrew
]

# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
BKMV_DELIMITER = "\t"    # Tab-delimited records (standard)
ZERO_THRESHOLD = 0.005   # Treat abs(balance) < this as zero

# Minimum number of fields expected per record type (for validation)
B100_MIN_FIELDS = 8
C100_MIN_FIELDS = 13

# ---------------------------------------------------------------------------
# Excel sheet names (Hebrew)
# ---------------------------------------------------------------------------
SHEET_MAIN     = "ריכוז"               # Summary / main result sheet
SHEET_SELECTED = "תנועות נבחרות"       # Selected movements (for SUM formulas)
SHEET_ALL_MVMT = "כל התנועות"          # All movements (audit trail)
SHEET_LOG      = "לוג שגיאות"          # Errors & warnings
