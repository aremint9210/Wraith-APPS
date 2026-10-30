"""
WRAITH // PDF Testsheet Verification & Audit Engine
Scans PDF reports to verify whether testsheets (CBM Condition Testsheet & Visual Inspection Checklist)
are properly attached at the back of each report.
"""

import os
import sys
import re
from pathlib import Path

def _safe_log(log_fn, msg: str, level: str = "INFO"):
    if not log_fn:
        return
    if log_fn is print:
        prefix = f"[{level}] " if level not in ("INFO", "OK") else ""
        print(f"{prefix}{msg}")
        return
    try:
        import inspect
        sig = inspect.signature(log_fn)
        if len(sig.parameters) >= 2:
            log_fn(msg, level)
        else:
            log_fn(msg)
    except Exception:
        try:
            print(msg)
        except Exception:
            pass


def is_cbm_testsheet_page(text: str) -> bool:
    """Checks if text contains CBM condition / IR / US / TEV test sheet markers."""
    t = text.upper()
    if "CBM CONDITION" in t:
        return True
    if "FDR. NO" in t and ("STATUS" in t or "LOAD" in t or "BREAKER" in t):
        return True
    if "TEST SHEET" in t and ("CBM" in t or "SUBSTATION" in t or "DELTA" in t):
        return True
    return False


def is_vi_checklist_page(text: str, has_images: bool = False) -> bool:
    """Checks if text contains Visual Inspection Checklist markers."""
    t = text.upper()
    if "VISUAL INSPECTION CHECKLIST" in t or "CHECKLIST FOR 11KV" in t:
        return True
    if "PE NAME (ERMS)" in t or "F/L NO (ERMS)" in t or "PE NAME (SITE)" in t:
        return True
    if "VER. : 1-2019" in t or "VER.:1-2019" in t or "VER. : 1 - 2019" in t:
        return True
    if "PCE" in t and "VISUAL INSPECTION" in t:
        return True
    return False


def check_pdf_has_testsheet(pdf_path: str) -> dict:
    """
    Analyzes a single PDF file to check if testsheet(s) are attached at the back.
    Returns a dictionary with detailed inspection metrics.
    """
    pdf_path = str(pdf_path)
    if not os.path.exists(pdf_path):
        return {
            "file": os.path.basename(pdf_path),
            "path": pdf_path,
            "status": "ERROR",
            "error": "File not found",
            "pages": 0,
            "has_cbm": False,
            "has_vi": False,
            "has_any_testsheet": False,
            "last_page_header": ""
        }

    try:
        import fitz
        doc = fitz.open(pdf_path)
        num_pages = len(doc)

        has_cbm = False
        has_vi = False
        cbm_page = None
        vi_page = None
        last_pages_headers = []

        # Check last 6 pages
        start_check = max(0, num_pages - 6)
        for i in range(start_check, num_pages):
            page = doc[i]
            txt = page.get_text()
            img_count = len(page.get_images())
            
            lines = [line.strip() for line in txt.splitlines() if line.strip()]
            header = lines[0] if lines else ("[Scanned Image]" if img_count > 0 else "[Blank Page]")
            last_pages_headers.append(f"P{i+1}: {header[:35]}")

            if not has_cbm and is_cbm_testsheet_page(txt):
                has_cbm = True
                cbm_page = i + 1

            if not has_vi and is_vi_checklist_page(txt, has_images=(img_count > 0)):
                has_vi = True
                vi_page = i + 1

        last_page_txt = doc[-1].get_text().strip()
        last_page_lines = [l.strip() for l in last_page_txt.splitlines() if l.strip()]
        last_page_hdr = last_page_lines[0] if last_page_lines else ("[Scanned Image]" if len(doc[-1].get_images()) > 0 else "[Empty]")
        
        doc.close()

        # Determine overall status
        if has_cbm and has_vi:
            status = "COMPLETE"
            summary_desc = f"Both Attached (CBM: P{cbm_page}, VI: P{vi_page})"
        elif has_cbm and not has_vi:
            status = "CBM_ONLY"
            summary_desc = f"CBM Testsheet attached (P{cbm_page}), VI Checklist Missing"
        elif not has_cbm and has_vi:
            status = "VI_ONLY"
            summary_desc = f"VI Checklist attached (P{vi_page}), CBM Testsheet Missing"
        else:
            status = "MISSING"
            summary_desc = "No Testsheet attached (Ends prematurely)"

        return {
            "file": os.path.basename(pdf_path),
            "path": pdf_path,
            "status": status,
            "pages": num_pages,
            "has_cbm": has_cbm,
            "has_vi": has_vi,
            "cbm_page": cbm_page,
            "vi_page": vi_page,
            "has_any_testsheet": (has_cbm or has_vi),
            "summary_desc": summary_desc,
            "last_page_header": last_page_hdr,
            "last_pages_trail": " | ".join(last_pages_headers)
        }

    except Exception as e:
        return {
            "file": os.path.basename(pdf_path),
            "path": pdf_path,
            "status": "ERROR",
            "error": str(e),
            "pages": 0,
            "has_cbm": False,
            "has_vi": False,
            "has_any_testsheet": False,
            "last_page_header": ""
        }


def scan_folder_for_testsheets(
    folder_path: str,
    ignore_group: bool = True,
    log_fn=print
) -> dict:
    """
    Scans an entire directory for PDF reports and evaluates their testsheet attachments.
    """
    if not os.path.exists(folder_path):
        _safe_log(log_fn, f"❌ Target directory not found: {folder_path}", "ERR")
        return {
            "folder": folder_path,
            "total_scanned": 0,
            "complete": [],
            "missing": [],
            "cbm_only": [],
            "vi_only": [],
            "errors": [f"Directory not found: {folder_path}"]
        }

    _safe_log(log_fn, f"🔍 Scanning PDF files in: {folder_path}...", "CMD")
    if ignore_group:
        _safe_log(log_fn, "ℹ️ Filtering out group files (*group*.pdf)...", "INFO")

    all_files = sorted(os.listdir(folder_path))
    pdf_files = []

    for f in all_files:
        if not f.lower().endswith(".pdf"):
            continue
        if ignore_group and "group" in f.lower():
            continue
        pdf_files.append(os.path.join(folder_path, f))

    total = len(pdf_files)
    _safe_log(log_fn, f"📊 Found {total} PDF reports to inspect.", "INFO")

    complete_list = []
    missing_list = []
    cbm_only_list = []
    vi_only_list = []
    error_list = []

    for idx, pdf_p in enumerate(pdf_files, 1):
        res = check_pdf_has_testsheet(pdf_p)
        fname = res["file"]

        if res["status"] == "COMPLETE":
            complete_list.append(res)
        elif res["status"] == "MISSING":
            missing_list.append(res)
            _safe_log(log_fn, f"  ❌ [{idx}/{total}] MISSING TESTSHEET: {fname} ({res['pages']} pages) - Ends at: {res['last_page_header']}", "WARN")
        elif res["status"] == "CBM_ONLY":
            cbm_only_list.append(res)
            _safe_log(log_fn, f"  ⚠️ [{idx}/{total}] CBM ONLY: {fname} (Missing VI Checklist)", "WARN")
        elif res["status"] == "VI_ONLY":
            vi_only_list.append(res)
            _safe_log(log_fn, f"  ⚠️ [{idx}/{total}] VI ONLY: {fname} (Missing CBM Testsheet)", "WARN")
        else:
            error_list.append(res)
            _safe_log(log_fn, f"  ⚠️ [{idx}/{total}] ERROR reading {fname}: {res.get('error')}", "ERR")

    _safe_log(log_fn, f"\n{'='*60}", "INFO")
    _safe_log(log_fn, f"📋 SCAN COMPLETE: {total} total files checked", "INFO")
    _safe_log(log_fn, f"  ✅ Complete (Both attached): {len(complete_list)}", "OK")
    if missing_list:
        _safe_log(log_fn, f"  ❌ MISSING Testsheets: {len(missing_list)}", "ERR")
    if cbm_only_list:
        _safe_log(log_fn, f"  ⚠️ CBM Testsheet Only: {len(cbm_only_list)}", "WARN")
    if vi_only_list:
        _safe_log(log_fn, f"  ⚠️ VI Checklist Only: {len(vi_only_list)}", "WARN")
    _safe_log(log_fn, f"{'='*60}\n", "INFO")

    return {
        "folder": folder_path,
        "total_scanned": total,
        "complete": complete_list,
        "missing": missing_list,
        "cbm_only": cbm_only_list,
        "vi_only": vi_only_list,
        "errors": error_list
    }


if __name__ == "__main__":
    default_dir = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42234207 - JOHOR - JBU - 400 PE IR US TEV\01. FULL REPORT"
    target = sys.argv[1] if len(sys.argv) > 1 else default_dir
    scan_folder_for_testsheets(target, ignore_group=True)
