"""
debug_pdf.py  –  run this manually to inspect what pdfplumber actually
extracts from your 331 PDF.

Usage:
    python debug_pdf.py "path/to/your/report331.pdf"
"""
import sys, re, os
sys.path.insert(0, os.path.dirname(__file__))

def dump_pdf(path):
    try:
        import pdfplumber
    except ImportError:
        print("pdfplumber not installed"); return

    print(f"\n{'='*70}")
    print(f"FILE: {path}")
    print(f"{'='*70}\n")

    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"\n{'─'*60}")
            print(f"PAGE {page_num}")
            print(f"{'─'*60}")

            # Strategy A: plain extract_text
            text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
            print("\n[A] extract_text() output:")
            for i, line in enumerate(text.splitlines(), 1):
                if line.strip():
                    print(f"  {i:3d} | {repr(line)}")

            # Strategy B: words sorted x0 DESC (RTL)
            words = page.extract_words(x_tolerance=4, y_tolerance=4,
                                       keep_blank_chars=False)
            if words:
                # Group into lines by rounded top
                rows = {}
                for w in words:
                    key = round(w["top"] / 3) * 3
                    rows.setdefault(key, []).append(w)

                print("\n[B] word-based RTL lines (x0 descending):")
                for top in sorted(rows):
                    line_words = sorted(rows[top], key=lambda w: w["x0"], reverse=True)
                    line_text  = " ".join(w["text"] for w in line_words)
                    if line_text.strip():
                        # Annotate tokens
                        codes    = re.findall(r'\b\d{3,}\b', line_text)
                        has_side = bool(re.search(r'[חז]', line_text))
                        amounts  = re.findall(r'\b[\d,]+(?:\.\d+)?\b', line_text)
                        tag = ""
                        if re.search(r'סה["\u05f4]?כ|סכום\s*\d|סהכ|total', line_text, re.I):
                            tag = "[SUMMARY]"
                        elif codes and amounts and has_side:
                            tag = "[ACCOUNT?]"
                        elif codes and not has_side:
                            tag = "[HEADER?]"
                        print(f"  y={top:4.0f} | {tag:12s} | {repr(line_text)}")

            if page_num >= 5:
                print("\n(stopping at page 5 – pass --all to see all pages)")
                break

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python debug_pdf.py <pdf_path>")
        sys.exit(1)
    dump_pdf(sys.argv[1])
