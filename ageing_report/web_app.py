"""
web_app.py  –  Ageing Report FastAPI backend
=============================================

POST /process  runs the full pipeline in four verified stages:

  STAGE 1  Parse 331 PDF          → count account rows found
  STAGE 2  Match to BKMVDATA      → count matched accounts
  STAGE 3  Generate & save Excel  → verify file exists on disk, print path
  STAGE 4  Auto-open with Excel   → os.startfile(), fallback: print path

CRITICAL DESIGN RULE:
  The Excel file is ALWAYS created (stages 3-4) regardless of how many
  accounts were found in stages 1-2.  Zero-account runs still produce a
  workbook so the user can see what happened.

Every stage logs to:
  (a) the uvicorn console window
  (b) output/process.log  (appended on every run)
  (c) the JSON response returned to the browser
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

# ── Paths ────────────────────────────────────────────────────────────────
_HERE       = Path(__file__).parent.resolve()
OUTPUT_DIR  = _HERE / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE    = OUTPUT_DIR / "process.log"

sys.path.insert(0, str(_HERE))

from config.constants import TARGET_YEAR
from excel.simple_excel import generate_simple_excel
from parsers.bkmv_fixed_parser import parse_bkmvdata_fixed
from parsers.report331_parser import parse_331_pdf
from processors.simple_balance import calculate_simple_results
from utils.file_utils import cleanup_temp_dir, extract_zip, find_bkmvdata, validate_directory
from utils.logger import clear_accumulated_logs, get_logger

log = get_logger(__name__)

app = FastAPI(title="Ageing Report – מערכת גיול יתרות")


# ---------------------------------------------------------------------------
# Logging helper  (writes to console + file)
# ---------------------------------------------------------------------------

def _log(msg: str, stage: str | None = None) -> None:
    prefix = f"[{stage}] " if stage else ""
    full   = f"{datetime.now().strftime('%H:%M:%S')}  {prefix}{msg}"
    print(full, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(full + "\n")
    except Exception:
        pass


def _run_started() -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _log(f"\n{'='*60}")
    _log(f"NEW RUN  {ts}")
    _log(f"OUTPUT_DIR: {OUTPUT_DIR}")
    _log(f"{'='*60}")


# ---------------------------------------------------------------------------
# GET /
# ---------------------------------------------------------------------------

def _load_html() -> str:
    return (_HERE / "templates" / "index.html").read_text(encoding="utf-8")


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(content=_load_html())


# ---------------------------------------------------------------------------
# GET /test-excel  –  quick smoke test (no uploads needed)
# ---------------------------------------------------------------------------

@app.get("/test-excel")
async def test_excel():
    """
    Creates a minimal 2-row Excel and auto-opens it.
    Verifies stages 3-4 independently of any PDF/BKMVDATA parsing.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment

    _run_started()
    _log("RUNNING SMOKE TEST (no real data)", "TEST")

    output_path = OUTPUT_DIR / f"test_smoke_{datetime.now().strftime('%H%M%S')}.xlsx"

    # Build
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Smoke Test"
    ws.sheet_view.rightToLeft = True
    ws.append(["Account", "Name", "Balance"])
    ws.append(["10001",   "Test Customer A",  5000.00])
    ws.append(["10002",   "Test Customer B", -2000.00])
    _log(f"Workbook built (2 rows)", "STAGE 3")

    # Save
    try:
        wb.save(str(output_path))
        size = output_path.stat().st_size
        _log(f"File saved: {output_path}  ({size:,} bytes)", "STAGE 3")
    except Exception as e:
        _log(f"SAVE FAILED: {e}", "STAGE 3")
        return JSONResponse({"error": f"Excel save failed: {e}"}, status_code=500)

    # Verify
    if not output_path.exists():
        _log("File does NOT exist after save!", "STAGE 3")
        return JSONResponse({"error": "File not found after save"}, status_code=500)
    _log("File existence verified", "STAGE 3")

    # Open
    opened, open_err = _open_file(output_path)
    _log(f"Auto-open: {'OK' if opened else 'FAILED: ' + str(open_err)}", "STAGE 4")

    return JSONResponse({
        "success":   True,
        "test_mode": True,
        "filename":  output_path.name,
        "path":      str(output_path),
        "opened":    opened,
        "open_error": str(open_err) if open_err else None,
    })


# ---------------------------------------------------------------------------
# POST /process  –  main pipeline  (4 verified stages)
# ---------------------------------------------------------------------------

@app.post("/process")
async def process_files(
    pdf_file:    UploadFile = File(...),
    bkmv_file:   UploadFile = File(None),
    folder_path: str        = Form(None),
    header_code: str        = Form(...),
):
    clear_accumulated_logs()
    _run_started()

    # Accumulated stage results (sent back to browser)
    stages: dict = {
        "stage1": {"ok": False, "detail": ""},
        "stage2": {"ok": False, "detail": ""},
        "stage3": {"ok": False, "detail": ""},
        "stage4": {"ok": False, "detail": ""},
    }

    tmp_dir:      str | None = None
    pdf_tmp_path: str | None = None

    try:
        header_code = header_code.strip()
        _log(f"header_code={header_code!r}  pdf={pdf_file.filename!r}  "
             f"bkmv={getattr(bkmv_file,'filename','—')!r}  "
             f"folder={folder_path!r}")

        # ── Validate ─────────────────────────────────────────────────────
        if not header_code:
            return _err("קוד הכותרת לא יכול להיות ריק", stages)
        if not pdf_file or not pdf_file.filename:
            return _err("יש להעלות קובץ PDF של דוח 331", stages)
        has_zip    = bkmv_file and bkmv_file.filename
        has_folder = folder_path and folder_path.strip()
        if not has_zip and not has_folder:
            return _err("יש להעלות ZIP או נתיב תיקייה", stages)

        # ── Resolve working directory ────────────────────────────────────
        if has_zip:
            zip_bytes = await bkmv_file.read()
            tmp_dir   = extract_zip(zip_bytes)
            working_dir = tmp_dir
            _log(f"ZIP extracted → {working_dir}")
        else:
            working_dir = folder_path.strip()
            if not validate_directory(working_dir):
                return _err(f"התיקייה '{working_dir}' לא נמצאה", stages)

        # ── Locate BKMVDATA ──────────────────────────────────────────────
        bkmvdata_path = find_bkmvdata(working_dir)
        if not bkmvdata_path:
            return _err("לא נמצא BKMVDATA.TXT", stages)
        _log(f"BKMVDATA: {bkmvdata_path}")

        # ── PDF → temp file ──────────────────────────────────────────────
        pdf_bytes = await pdf_file.read()
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, prefix="r331_") as f:
            f.write(pdf_bytes)
            pdf_tmp_path = f.name
        _log(f"PDF temp: {pdf_tmp_path}  ({len(pdf_bytes):,} bytes)")

        # ════════════════════════════════════════════════════════════════
        # STAGE 1  –  Parse 331 PDF
        # ════════════════════════════════════════════════════════════════
        _log("─── STAGE 1: Parse 331 PDF ───", "STAGE 1")
        try:
            report = parse_331_pdf(pdf_tmp_path, header_code)
            n_found = len(report.accounts)
            detail1 = (
                f"header_name='{report.header_name}'  "
                f"accounts_found={n_found}  "
                f"warnings={len(report.warnings)}"
            )
            _log(detail1, "STAGE 1")
            if report.warnings:
                for w in report.warnings[:3]:
                    _log(f"  warning: {w}", "STAGE 1")
            stages["stage1"]["ok"]     = True
            stages["stage1"]["detail"] = detail1
        except Exception as e:
            tb = traceback.format_exc()
            _log(f"FAILED: {e}\n{tb}", "STAGE 1")
            stages["stage1"]["detail"] = f"Exception: {e}"
            # Still try to build an empty Excel
            report = None
            n_found = 0

        # ════════════════════════════════════════════════════════════════
        # STAGE 2  –  Parse BKMVDATA (B100 + C100 only, exact mapping)
        # ════════════════════════════════════════════════════════════════
        _log("─── STAGE 2: Parse BKMVDATA (fixed parser) + match ───", "STAGE 2")
        b100_records   = {}
        c100_movements = []
        debug_samples  = []
        results        = []
        try:
            b100_records, c100_movements, debug_samples, bkmv_warnings = \
                parse_bkmvdata_fixed(bkmvdata_path)

            _log(
                f"BKMVDATA: {len(b100_records)} B100 accounts, "
                f"{len(c100_movements)} C100 movements",
                "STAGE 2",
            )
            for w in bkmv_warnings[:5]:
                _log(f"  bkmv warning: {w}", "STAGE 2")

            # Extract just the account number strings from the 331 PDF
            pdf_account_numbers = (
                [al.account_number for al in report.accounts]
                if report else []
            )

            results = calculate_simple_results(
                pdf_account_numbers, b100_records, c100_movements
            )

            n_matched = sum(1 for r in results if r.matched)
            detail2 = (
                f"b100_accounts={len(b100_records)}  "
                f"c100_movements={len(c100_movements)}  "
                f"331_accounts={n_found}  "
                f"matched={n_matched}"
            )
            _log(detail2, "STAGE 2")
            stages["stage2"]["ok"]     = True
            stages["stage2"]["detail"] = detail2
        except Exception as e:
            tb = traceback.format_exc()
            _log(f"FAILED: {e}\n{tb}", "STAGE 2")
            stages["stage2"]["detail"] = f"Exception: {e}"

        # ════════════════════════════════════════════════════════════════
        # STAGE 3  –  Generate Excel  (ALWAYS runs regardless of above)
        # ════════════════════════════════════════════════════════════════
        _log("─── STAGE 3: Generate Excel ───", "STAGE 3")
        ts           = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename     = f"Ageing_{header_code}_{TARGET_YEAR}_{ts}.xlsx"
        output_path  = OUTPUT_DIR / filename
        _log(f"Target path: {output_path}", "STAGE 3")

        excel_bytes: bytes | None = None
        try:
            excel_bytes = generate_simple_excel(
                results        = results,
                debug_samples  = debug_samples,
                c100_movements = c100_movements,
                b100_records   = b100_records,
                header_code    = header_code,
                header_name    = report.header_name if report else "",
            )
            _log(f"generate_simple_excel() returned {len(excel_bytes):,} bytes", "STAGE 3")
        except Exception as e:
            tb = traceback.format_exc()
            _log(f"generate_simple_excel() FAILED: {e}\n{tb}", "STAGE 3")
            stages["stage3"]["detail"] = f"generate_simple_excel() failed: {e}"
            # Fall back to a minimal workbook so stages 3-4 can still run
            excel_bytes = _build_fallback_workbook(
                header_code, results, report, str(e)
            )
            _log(f"Fallback workbook built ({len(excel_bytes):,} bytes)", "STAGE 3")

        # Write to disk
        try:
            output_path.write_bytes(excel_bytes)
            _log(f"write_bytes() done", "STAGE 3")
        except Exception as e:
            tb = traceback.format_exc()
            _log(f"write_bytes() FAILED: {e}\n{tb}", "STAGE 3")
            stages["stage3"]["detail"] += f"  |  write_bytes failed: {e}"
            return _err(f"לא ניתן לשמור קובץ Excel: {e}", stages)

        # Verify
        if not output_path.exists():
            msg = f"File does NOT exist after write: {output_path}"
            _log(msg, "STAGE 3")
            stages["stage3"]["detail"] = msg
            return _err("הקובץ לא נשמר על הדיסק", stages)

        actual_size = output_path.stat().st_size
        _log(f"FILE VERIFIED: {output_path}  ({actual_size:,} bytes)", "STAGE 3")
        stages["stage3"]["ok"]     = True
        stages["stage3"]["detail"] = f"path={output_path}  size={actual_size:,} bytes"

        # ════════════════════════════════════════════════════════════════
        # STAGE 4  –  Auto-open
        # ════════════════════════════════════════════════════════════════
        _log("─── STAGE 4: Auto-open ───", "STAGE 4")
        opened, open_err = _open_file(output_path)
        if opened:
            _log("os.startfile() called successfully", "STAGE 4")
            stages["stage4"]["ok"]     = True
            stages["stage4"]["detail"] = "opened"
        else:
            _log(f"Auto-open failed: {open_err}", "STAGE 4")
            stages["stage4"]["detail"] = str(open_err)

        # ── Response ─────────────────────────────────────────────────────
        response_payload = {
            "success":      True,
            "filename":     filename,
            "path":         str(output_path),
            "opened":       opened,
            "open_error":   str(open_err) if open_err else None,
            "accounts":     len(results),
            "header_name":  report.header_name if report else "",
            "stages":       stages,
            "debug_lines":  (report.debug_lines if report and not report.accounts else []),
        }
        _log(f"Returning success response: accounts={len(results)}  opened={opened}")
        return JSONResponse(response_payload)

    except Exception as exc:
        tb = traceback.format_exc()
        _log(f"UNHANDLED EXCEPTION: {exc}\n{tb}")
        return JSONResponse(
            status_code=500,
            content={"error": f"שגיאה לא צפויה: {exc}", "stages": stages,
                     "traceback": tb},
        )

    finally:
        if pdf_tmp_path:
            try:
                os.unlink(pdf_tmp_path)
            except Exception:
                pass
        if tmp_dir:
            cleanup_temp_dir(tmp_dir)


# ---------------------------------------------------------------------------
# GET /download/{filename}
# ---------------------------------------------------------------------------

@app.get("/download/{filename}")
async def download_file(filename: str):
    filepath = _safe_output_path(filename)
    if filepath is None or not filepath.exists():
        return JSONResponse(status_code=404, content={"error": "קובץ לא נמצא"})
    return FileResponse(
        path=str(filepath),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# POST /open/{filename}
# ---------------------------------------------------------------------------

@app.post("/open/{filename}")
async def reopen_file(filename: str):
    filepath = _safe_output_path(filename)
    if filepath is None or not filepath.exists():
        return JSONResponse(status_code=404, content={"error": "קובץ לא נמצא"})
    opened, err = _open_file(filepath)
    return JSONResponse({"opened": opened, "path": str(filepath),
                         "error": str(err) if err else None})


# ---------------------------------------------------------------------------
# POST /debug-pdf
# ---------------------------------------------------------------------------

@app.post("/debug-pdf")
async def debug_pdf(pdf_file: UploadFile = File(...)):
    pdf_bytes = await pdf_file.read()
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    try:
        from parsers.report331_parser import (
            _classify, _extract_lines, TYPE_HEADER, _normalise,
        )
        lines = _extract_lines(tmp, [])
        classified = []
        for pno, raw in lines[:400]:
            info = _classify(raw)
            classified.append({
                "page": pno, "type": info.ltype, "code": info.code,
                "name": info.name[:60], "balance": info.balance_raw,
                "side": info.balance_side, "raw": _normalise(raw)[:100],
            })
        headers = [c for c in classified if c["type"] == TYPE_HEADER]
        return JSONResponse({"total_lines": len(lines), "headers": headers,
                             "classified": classified})
    finally:
        try:
            os.unlink(tmp)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _open_file(path: Path):
    """Open path with OS default app. Returns (opened: bool, error: Exception|None)."""
    try:
        if sys.platform == "win32":
            os.startfile(str(path))
        elif sys.platform == "darwin":
            import subprocess; subprocess.Popen(["open", str(path)])
        else:
            import subprocess; subprocess.Popen(["xdg-open", str(path)])
        return True, None
    except Exception as e:
        return False, e


def _build_fallback_workbook(
    header_code: str,
    results,
    report,
    error_msg: str,
) -> bytes:
    """Minimal workbook written when generate_simple_excel() fails."""
    import io
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ageing"
    ws.sheet_view.rightToLeft = True
    ws.append(["מספר חשבון", "שם חשבון", "יתרת פתיחה", "יתרת סגירה", "הערה"])
    ws.append(["", f"שגיאה: {error_msg[:80]}", "", "", ""])
    if report:
        for al in (report.accounts or []):
            ws.append([al.account_number, al.account_name, "", "", ""])
    elif results:
        for r in results:
            ws.append([
                r.pdf_account_number,
                r.account_name,
                r.opening_balance,
                r.closing_balance,
                "" if r.matched else "לא נמצא",
            ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _safe_output_path(filename: str) -> Path | None:
    if not re.fullmatch(r"[\w\-]+\.xlsx", filename):
        return None
    path = (OUTPUT_DIR / filename).resolve()
    if not str(path).startswith(str(OUTPUT_DIR.resolve())):
        return None
    return path


def _err(message: str, stages: dict | None = None) -> JSONResponse:
    _log(f"Returning error: {message}")
    content: dict = {"error": message}
    if stages:
        content["stages"] = stages
    return JSONResponse(status_code=422, content=content)


# ---------------------------------------------------------------------------
# Entry-point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8501, reload=False)
