"""
report331_parser.py  –  Israeli 331 PDF parser (rewritten)
===========================================================

Design principles
-----------------
1.  TEXT-PATTERN classification only – no reliance on x-position / indentation.
2.  A line is classified by what it *contains*, not where it sits on the page.
3.  Hierarchical context stack: headers push/pop as we walk through the file.
4.  Full debug trace emitted when no accounts are found (or always in DEBUG mode).

Line classification rules
--------------------------
SUMMARY  – line contains "סה"כ", "סהכ", "סכום NNNN", "total" (case-insensitive)
HEADER   – starts with a short numeric code AND contains no monetary-amount+side pair
ACCOUNT  – starts with a numeric code AND has at least one amount AND a ח/ז indicator
           (the amount+side pair can appear anywhere on the same line, or the
           ח/ז can be on a continuation fragment of the same logical row)
OTHER    – everything else (page headings, column labels, blank lines, etc.)

Stop condition when collecting under header H
----------------------------------------------
Stop when a HEADER line appears whose code is NOT a string-prefix of H.
Examples (H = "1342"):
    • "13421"  → child  → keep going (sub-header, its accounts also belong to H)
    • "1342"   → self   → keep going (duplicate header line)
    • "1343"   → sibling → STOP
    • "134"    → parent  → STOP
    • "28300"  → this is an account number, never a stop trigger by itself

Public API
----------
parse_331_pdf(pdf_path, header_code) -> Report331Result
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from utils.logger import get_logger

log = get_logger(__name__)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class AccountLine:
    account_number: str
    account_name:   str
    balance:        float       # signed: positive = debit, negative = credit
    balance_raw:    str
    page_number:    int = 0
    raw_line:       str = ""    # normalised raw PDF line for debugging


@dataclass
class Report331Result:
    header_code:  str
    header_name:  str
    accounts:          List[AccountLine] = field(default_factory=list)
    warnings:          List[str]         = field(default_factory=list)
    debug_lines:       List[str]         = field(default_factory=list)

    # The header row itself (raw text from PDF)
    header_line_raw:   str   = ""

    # Total line found in the PDF for this header (the סה"כ row).
    # None = not found (never treat as 0 – those are different situations).
    header_total:      Optional[float] = None   # signed: +debit / -credit
    header_total_raw:  str             = ""     # full normalised text of the chosen summary line
    header_total_side: str             = ""     # "ח" or "ז" or ""
    header_total_code_matched: bool    = False  # True if summary was found by code match

    # Why collection stopped
    stop_reason: str = ""

    # Every summary (סה"כ) line seen inside the section, in order: (page_no, text)
    section_summary_lines: List[Tuple[int, str]] = field(default_factory=list)

    # Every line in the section: (page_no, normalised_text, TYPE_*)
    section_lines: List[Tuple[int, str, str]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

TYPE_HEADER  = "HEADER"
TYPE_ACCOUNT = "ACCOUNT"
TYPE_SUMMARY = "SUMMARY"
TYPE_OTHER   = "OTHER"

# Detects the Hebrew/English summary markers
_SUMMARY_RE = re.compile(
    r'סה["\u05f4\u2019\u0022]?כ'   # סה"כ  (with various quote chars)
    r'|סהכ'
    r'|סה כ'
    r'|סכום\s*\d'
    r'|total[\s:]',
    re.IGNORECASE | re.UNICODE,
)

# A "money amount": digits with optional thousands-comma/dot and decimals.
# Matches:  1,234.56  |  1234.56  |  1,234  |  1234  |  0
_AMOUNT_RE = re.compile(
    r'\b\d{1,3}(?:[,\.]\d{3})*(?:[,\.]\d{1,2})?\b'
    r'|\b\d+\b',
    re.UNICODE,
)

# ח = debit indicator, ז = credit indicator (or the full words)
_SIDE_RE = re.compile(r'(?<!\w)([חז])(?!\w)|חובה|זכות', re.UNICODE)

# A numeric "code" token: 2+ consecutive digits (possibly with leading/trailing spaces)
_CODE_RE = re.compile(r'^\s*(\d{2,})\s+(.*?)\s*$', re.DOTALL)


# ---------------------------------------------------------------------------
# Line normalisation helper
# ---------------------------------------------------------------------------

def _normalise(text: str) -> str:
    """Strip control chars, collapse whitespace, normalise Hebrew quotation marks."""
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r'[\x00-\x1f\x7f]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# ---------------------------------------------------------------------------
# Amount parsing
# ---------------------------------------------------------------------------

def _parse_amount_str(s: str) -> float:
    """Convert an Israeli-format amount string to float (strips commas/dots used as separators)."""
    s = s.strip()
    # Decide separator style: if last non-digit separator has 2-digit group → decimal
    # Simple heuristic: remove the rightmost separator that splits into <=2 digits,
    # treat the rest as thousands.
    # E.g.  "1,234.56" → 1234.56   "1.234,56" → 1234.56   "1,234" → 1234
    s2 = re.sub(r'[,\.](?=\d{3}(?!\d))', '', s)   # remove thousands separator
    s2 = s2.replace(',', '.').strip()               # normalise decimal separator
    try:
        return float(s2)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# Core line classifier
# ---------------------------------------------------------------------------

@dataclass
class _LineInfo:
    ltype:        str
    code:         str        # leading numeric code found (empty if none)
    name:         str        # everything after the code, before the balance part
    balance_raw:  str        # the matched amount string (empty if none)
    balance_side: str        # "ח" or "ז" or ""
    balance:      float      # signed float


def _classify(raw: str) -> _LineInfo:
    line = _normalise(raw)
    empty = _LineInfo(TYPE_OTHER, "", "", "", "", 0.0)

    if not line:
        return empty

    # ── 1. Summary? ──────────────────────────────────────────────────────
    if _SUMMARY_RE.search(line):
        return _LineInfo(TYPE_SUMMARY, "", line, "", "", 0.0)

    # ── 2. Does the line start with a numeric code? ───────────────────────
    m = _CODE_RE.match(line)
    if not m:
        # No leading code → not a header or account
        return empty

    code = m.group(1)
    rest = m.group(2)   # everything after the code

    # ── 3. Is there a ח/ז indicator anywhere on the line? ────────────────
    side_m = _SIDE_RE.search(line)
    side = ""
    if side_m:
        raw_side = side_m.group()
        if raw_side in ("ח", "חובה"):
            side = "ח"
        elif raw_side in ("ז", "זכות"):
            side = "ז"

    # ── 4. Are there monetary amounts on the line? ────────────────────────
    amounts = _AMOUNT_RE.findall(line)
    # Remove the code itself from amounts list (it's a number too)
    amounts_without_code = [a for a in amounts if a.replace(',','').replace('.','') != code]

    has_amounts = len(amounts_without_code) > 0

    # ── 5. Classify ───────────────────────────────────────────────────────
    if has_amounts and side:
        # Account line: code + name + balance amount + ח/ז
        # Extract the name: strip the trailing amounts+side from `rest`
        # Strategy: split `rest` on the first occurrence of an amount that
        # looks like a balance (last large amount near the ח/ז).
        name = _extract_name_from_account_line(rest, side, amounts_without_code)
        balance_raw = amounts_without_code[-1] if amounts_without_code else "0"
        balance_val = _parse_amount_str(balance_raw)
        if side == "ז":
            balance_val = -balance_val
        return _LineInfo(TYPE_ACCOUNT, code, name.strip(), balance_raw, side, balance_val)

    elif has_amounts and not side:
        # Has amounts but no side indicator.
        # This is ambiguous. Treat as HEADER for now (could be a subtotal without ח/ז).
        name = rest.strip()
        return _LineInfo(TYPE_HEADER, code, name, "", "", 0.0)

    else:
        # No amounts (or only the code itself) → header line
        name = rest.strip()
        return _LineInfo(TYPE_HEADER, code, name, "", "", 0.0)


def _extract_name_from_account_line(rest: str, side: str, amounts: List[str]) -> str:
    """
    Given the text after the leading code, extract just the account name,
    stripping the trailing amounts and ח/ז.
    """
    s = rest

    # Remove ח/ז words
    s = re.sub(r'(?<!\w)(ח|ז|חובה|זכות)(?!\w)', '', s, flags=re.UNICODE)

    # Remove all amounts (greedy, from right)
    for amt in sorted(set(amounts), key=len, reverse=True):
        s = s.replace(amt, '', 1)   # remove first occurrence from the right
        # We strip right-to-left so iterate in descending order

    # Clean up
    s = re.sub(r'\s{2,}', ' ', s).strip()
    # Remove leading/trailing punctuation artifacts
    s = s.strip('- \t,.')
    return s


# ---------------------------------------------------------------------------
# Hierarchy helpers
# ---------------------------------------------------------------------------

def _is_descendant_or_self(code: str, parent: str) -> bool:
    """True if code starts with parent (child-or-self in prefix hierarchy)."""
    return code.startswith(parent)


def _is_stop_header(code: str, target: str) -> bool:
    """
    True if seeing a HEADER with this code should stop collection when
    we are inside target header.
    Rules:
      - Stop if code does NOT start with target  (sibling or uncle)
      - Stop if code IS a strict prefix of target (parent)
      - Do NOT stop for code == target (duplicate) or code starts with target (child)
    """
    if code == target:
        return False                        # same header repeated
    if code.startswith(target):
        return False                        # child header, keep collecting
    return True                             # sibling or parent → stop


# ---------------------------------------------------------------------------
# PDF text extraction strategies
# ---------------------------------------------------------------------------

def _extract_lines(pdf_path: str, warnings: List[str]) -> List[Tuple[int, str]]:
    """
    Extract (page_no, line_text) pairs using the best available strategy.
    Tries three strategies and picks the richest result.
    """
    # Strategy A: pdfplumber extract_text  (logical order, usually good for RTL)
    lines_a = _strategy_extract_text(pdf_path, warnings)

    # Strategy B: pdfplumber word-based, sorted x0 DESC (explicit RTL reconstruction)
    lines_b = _strategy_word_rtl(pdf_path, warnings)

    # Pick the strategy that finds more account-looking lines
    def account_score(lines):
        return sum(1 for _, l in lines if _classify(l).ltype == TYPE_ACCOUNT)

    score_a = account_score(lines_a)
    score_b = account_score(lines_b)

    log.info("PDF extraction – strategy A found %d account-like lines, B found %d",
             score_a, score_b)

    # Return the better one; if tied, prefer A (simpler)
    chosen = lines_b if score_b > score_a else lines_a
    strategy_name = "B (word RTL)" if score_b > score_a else "A (extract_text)"
    log.info("Using strategy %s", strategy_name)
    return chosen


def _strategy_extract_text(pdf_path: str, warnings: List[str]) -> List[Tuple[int, str]]:
    try:
        import pdfplumber   # type: ignore
    except ImportError:
        return []
    result: List[Tuple[int, str]] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pno, page in enumerate(pdf.pages, 1):
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                for line in text.splitlines():
                    result.append((pno, line))
    except Exception as exc:
        warnings.append(f"extract_text error: {exc}")
        log.warning("extract_text failed: %s", exc)
    return result


def _strategy_word_rtl(pdf_path: str, warnings: List[str]) -> List[Tuple[int, str]]:
    """
    Extract words, group by row (y-tolerance), sort each row x0 DESC (RTL),
    join into line strings.
    """
    try:
        import pdfplumber   # type: ignore
    except ImportError:
        return []
    result: List[Tuple[int, str]] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pno, page in enumerate(pdf.pages, 1):
                words = page.extract_words(
                    x_tolerance=4, y_tolerance=4, keep_blank_chars=False
                )
                if not words:
                    continue
                # Group by row
                rows: dict = {}
                for w in words:
                    key = round(w["top"] / 3) * 3
                    rows.setdefault(key, []).append(w)
                for top_key in sorted(rows):
                    row_words = sorted(rows[top_key],
                                       key=lambda w: w["x0"], reverse=True)
                    line_text = " ".join(w["text"] for w in row_words)
                    result.append((pno, line_text))
    except Exception as exc:
        warnings.append(f"word-RTL error: {exc}")
        log.warning("word-RTL extraction failed: %s", exc)
    return result


# ---------------------------------------------------------------------------
# Main public entry point
# ---------------------------------------------------------------------------

def parse_331_pdf(pdf_path: str, header_code: str) -> Report331Result:
    header_code = header_code.strip()
    result = Report331Result(header_code=header_code, header_name="")

    log.info("parse_331_pdf: %s  |  target header='%s'", pdf_path, header_code)

    # ── Extract lines ────────────────────────────────────────────────────
    raw_lines = _extract_lines(pdf_path, result.warnings)
    if not raw_lines:
        result.warnings.append(
            "לא ניתן היה לחלץ טקסט מה-PDF. "
            "ייתכן שמדובר ב-PDF סרוק. נדרש OCR."
        )
        return result

    log.info("Total raw lines extracted: %d", len(raw_lines))

    # ── Run state machine ────────────────────────────────────────────────
    _run_state_machine(raw_lines, header_code, result)

    # ── Identify the best summary / total line ────────────────────────────
    pg, summary_txt, code_matched = _find_best_summary(
        result.section_summary_lines, header_code
    )
    if summary_txt:
        total_val, total_side = _parse_summary_amount(summary_txt, header_code)
        result.header_total               = total_val
        result.header_total_raw           = summary_txt
        result.header_total_side          = total_side
        result.header_total_code_matched  = code_matched
        log.info(
            "Summary total: %.2f (%s)  code_matched=%s  raw='%s'",
            total_val, total_side, code_matched, summary_txt[:60],
        )
        if not code_matched:
            result.warnings.append(
                f"Summary total taken from last summary line (header code '{header_code}' "
                f"not found in any summary). Verify manually."
            )
    else:
        result.header_total = None   # explicitly NOT FOUND
        result.warnings.append(
            f"No summary (סה\"כ) line found inside section '{header_code}'. "
            f"Stop reason: {result.stop_reason}"
        )

    # ── Debug dump when nothing found ────────────────────────────────────
    if not result.accounts:
        _emit_debug(raw_lines, header_code, result)

    log.info("Result: %d accounts under header '%s' ('%s')",
             len(result.accounts), header_code, result.header_name)
    return result


# ---------------------------------------------------------------------------
# Summary-line amount extraction
# ---------------------------------------------------------------------------

def _parse_summary_amount(raw: str, header_code: str = "") -> Tuple[float, str]:
    """
    Extract the signed balance and side from a summary line (סה"כ row).

    Strategy:
      1. Locate the ח/ז indicator.
      2. Search for amounts ONLY in the text that appears before the ח/ז
         indicator (the balance is always left of the side-char in logical RTL).
      3. Strip out the header_code itself so "סה\"כ 1342 8,000 ח" doesn't
         confuse 1342 for the balance.
      4. Take the last remaining amount (closest to ח/ז).
      5. Apply sign: ח → positive, ז → negative.

    Returns (signed_float, side_char).  (0.0, "") if nothing found.
    """
    line = _normalise(raw)

    # Find ח/ז and restrict amount search to text before it
    side_m = _SIDE_RE.search(line)
    side = ""
    if side_m:
        rs = side_m.group()
        side = "ח" if rs in ("ח", "חובה") else "ז"
        search_text = line[:side_m.start()]   # only look left of the side indicator
    else:
        search_text = line

    amounts = _AMOUNT_RE.findall(search_text)

    # Remove the header code from candidates so it isn't mistaken for a balance
    if header_code:
        amounts = [a for a in amounts
                   if a.replace(",", "").replace(".", "") != header_code]

    if not amounts:
        return 0.0, side

    # Last amount before ח/ז = the balance
    amount = _parse_amount_str(amounts[-1])
    if side == "ז":
        amount = -amount

    return amount, side


def _find_best_summary(
    summary_lines: List[Tuple[int, str]],
    header_code: str,
) -> Tuple[int, str, bool]:
    """
    From all summary lines collected inside the section, pick the one
    most likely to be the total for header_code.

    Priority:
      1. A summary line that contains header_code as a standalone number
         (e.g. "סה\"כ 1342 8,000 ח" contains "1342").
      2. Fall back to the LAST summary line in the section (most inclusive
         total in a hierarchical report — the parent summary comes after
         child summaries, so the last one before the section ends belongs
         to the outermost level we collected).

    Returns (page_no, text, code_was_matched).
    Returns (-1, "", False) if the list is empty.
    """
    if not summary_lines:
        return -1, "", False

    code_re = re.compile(r"(?<!\d)" + re.escape(header_code) + r"(?!\d)")
    for pg, txt in summary_lines:
        if code_re.search(txt):
            return pg, txt, True

    # No code match – use last summary
    pg, txt = summary_lines[-1]
    return pg, txt, False


# ---------------------------------------------------------------------------
# State machine
# ---------------------------------------------------------------------------

def _run_state_machine(
    lines:       List[Tuple[int, str]],
    header_code: str,
    result:      Report331Result,
) -> None:
    STATE_SEARCHING  = "SEARCHING"
    STATE_COLLECTING = "COLLECTING"
    STATE_DONE       = "DONE"

    state = STATE_SEARCHING
    seen  : set = set()

    for page_no, raw in lines:
        if state == STATE_DONE:
            break

        info = _classify(raw)
        norm = _normalise(raw)

        if state == STATE_SEARCHING:
            if info.ltype == TYPE_HEADER and info.code == header_code:
                state = STATE_COLLECTING
                result.header_name     = info.name
                result.header_line_raw = norm
                log.info("  Header '%s' found on page %d: '%s'",
                         header_code, page_no, info.name)

        elif state == STATE_COLLECTING:
            # Record every line in the section for debug (including OTHER)
            result.section_lines.append((page_no, norm, info.ltype))

            if info.ltype == TYPE_ACCOUNT:
                key = info.code
                if key in seen:
                    log.debug("  Duplicate account %s – skipped", key)
                    continue
                seen.add(key)
                result.accounts.append(AccountLine(
                    account_number=info.code,
                    account_name=info.name,
                    balance=info.balance,
                    balance_raw=f"{info.balance_raw} {info.balance_side}".strip(),
                    page_number=page_no,
                    raw_line=norm[:120],
                ))
                log.debug("  ✓ account  %s  |  %s  |  %s %s",
                          info.code, info.name, info.balance_raw, info.balance_side)

            elif info.ltype == TYPE_HEADER:
                if _is_stop_header(info.code, header_code):
                    result.stop_reason = (
                        f"HEADER '{info.code}' on page {page_no} "
                        f"(sibling/parent of '{header_code}')"
                    )
                    log.info("  Stop: %s", result.stop_reason)
                    state = STATE_DONE
                else:
                    log.debug("  Sub-header %s – continue collecting", info.code)

            elif info.ltype == TYPE_SUMMARY:
                # Record the summary line but DO NOT stop.
                # Stopping on the first summary breaks hierarchical sections where
                # sub-headers each have their own summary before the parent total.
                result.section_summary_lines.append((page_no, norm))
                log.debug("  Summary recorded (continuing): %s", norm[:80])

            # TYPE_OTHER: recorded in section_lines, otherwise ignored

    # If we exhausted all lines without hitting a stop-header
    if state == STATE_COLLECTING:
        result.stop_reason = "END_OF_DOCUMENT (no sibling/parent header found after section)"


# ---------------------------------------------------------------------------
# Debug output
# ---------------------------------------------------------------------------

def _emit_debug(
    lines:       List[Tuple[int, str]],
    header_code: str,
    result:      Report331Result,
) -> None:
    """
    When no accounts are found, build a detailed diagnostic report
    and store it in result.debug_lines and the Python logger.
    """
    debug: List[str] = []
    debug.append(f"=== DEBUG: No accounts found under header '{header_code}' ===")

    # Collect all header lines
    all_headers = [
        (pno, _classify(raw))
        for pno, raw in lines
        if _classify(raw).ltype == TYPE_HEADER
    ]
    debug.append(f"\nAll detected HEADER rows ({len(all_headers)} total):")
    for pno, info in all_headers[:60]:
        debug.append(f"  p{pno:02d} | code={info.code!r:12s} | name={info.name[:40]!r}")

    # Find the target header
    target_idx = next(
        (i for i, (_, raw) in enumerate(lines)
         if _classify(raw).ltype == TYPE_HEADER and _classify(raw).code == header_code),
        None,
    )
    if target_idx is None:
        debug.append(f"\n*** Header '{header_code}' was NOT found in the PDF ***")
        debug.append("Possible causes:")
        debug.append("  1. The entered header code doesn't match any code in the PDF.")
        debug.append("  2. The PDF extracts the header without the code on the same line.")
        debug.append("  3. The PDF is scanned (image-based) and cannot be parsed.")
        result.warnings.append(
            f"הכותרת '{header_code}' לא זוהתה בדוח 331.  "
            "בדוק את הקוד ואת פלט debug_pdf.py."
        )
    else:
        debug.append(f"\nHeader '{header_code}' found at line index {target_idx}.")
        debug.append("Next 40 lines after it (with classification):")
        for pno, raw in lines[target_idx + 1: target_idx + 41]:
            info = _classify(raw)
            tag  = f"[{info.ltype:8s}]"
            extra = ""
            if info.ltype == TYPE_ACCOUNT:
                extra = f"  code={info.code}  balance={info.balance_raw}{info.balance_side}"
            elif info.ltype == TYPE_HEADER:
                extra = f"  code={info.code}  name={info.name[:30]}"
            line_repr = _normalise(raw)[:70]
            debug.append(f"  p{pno:02d} {tag} {line_repr!r}{extra}")

        result.warnings.append(
            f"נמצאה כותרת '{header_code}' אך לא זוהו שורות חשבון מתחתיה.  "
            "ראה debug_lines לפרטים."
        )

    result.debug_lines = debug

    # Also emit to logger
    for line in debug:
        log.warning("  %s", line)
