"""
WRAITH Lossless Word to PDF Conversion Engine
Features:
- Native Microsoft Word COM Automation (Word.Application)
- 100% Lossless Image Resolution (DoNotCompressPictures = True)
- High-Fidelity 300+ DPI Print Optimization for Radiometric IR & DC Visual Photos
- Optional Automatic Testsheet PDF Append (04. TESTSHEET/PDF) for QUICK REPORT or custom workflows
- Robust PDF Merging via PyMuPDF (fitz) with fallback to pypdf.PdfWriter
- Flexible Single-File, Multi-File, or Directory-Wide Conversion
"""

import os
import sys
import time
import glob
import re
from pathlib import Path

def _safe_log(log_fn, msg: str):
    if not log_fn:
        return
    try:
        log_fn(msg)
    except (UnicodeEncodeError, Exception):
        try:
            clean_msg = msg.encode("ascii", "replace").decode("ascii")
            log_fn(clean_msg)
        except Exception:
            pass


def find_matching_testsheet_pdf(docx_path: str, testsheet_pdf_dir: str) -> str:
    """
    Finds the corresponding Testsheet PDF in 04. TESTSHEET/PDF matching the substation number.
    """
    if not testsheet_pdf_dir or not os.path.exists(testsheet_pdf_dir):
        return None

    docx_base = os.path.basename(docx_path)
    # Extract numerical prefix (e.g. '178' from '178. 28072026 ... .docx')
    match = re.match(r'^(\d+)', docx_base.strip())
    if not match:
        return None

    sub_num = match.group(1)

    # 1. Look for exact prefix match (e.g. '178. ... .pdf' or '178 ... .pdf')
    for f in sorted(os.listdir(testsheet_pdf_dir)):
        if f.lower().endswith(".pdf"):
            f_clean = f.strip()
            if f_clean.startswith(f"{sub_num}.") or f_clean.startswith(f"{sub_num} ") or f_clean.startswith(f"{int(sub_num)}."):
                return os.path.join(testsheet_pdf_dir, f)

    # 2. Fallback contains match
    for f in sorted(os.listdir(testsheet_pdf_dir)):
        if f.lower().endswith(".pdf") and sub_num in f:
            return os.path.join(testsheet_pdf_dir, f)

    return None


def merge_pdf_files(primary_pdf: str, secondary_pdf: str, output_pdf: str) -> bool:
    """
    Losslessly appends secondary_pdf to the back of primary_pdf using PyMuPDF (fitz) or pypdf.PdfWriter.
    """
    try:
        import fitz
        doc1 = fitz.open(primary_pdf)
        doc2 = fitz.open(secondary_pdf)
        doc1.insert_pdf(doc2)
        doc1.save(output_pdf, garbage=4, deflate=True)
        doc1.close()
        doc2.close()
        return True
    except Exception:
        try:
            from pypdf import PdfWriter
            writer = PdfWriter()
            writer.append(primary_pdf)
            writer.append(secondary_pdf)
            with open(output_pdf, "wb") as f_out:
                writer.write(f_out)
            writer.close()
            return True
        except Exception as err:
            print(f"[Merge Error]: {err}")
            return False


def convert_word_files_to_pdf(
    docx_files: list,
    destination_folder: str = None,
    append_testsheet: bool = False,
    testsheet_pdf_dir: str = None,
    log_fn=print
) -> dict:
    """
    Converts a list of Word (.docx) files into high-resolution lossless PDF files.
    Optionally appends the corresponding testsheet PDF at the back.
    """
    if not docx_files:
        _safe_log(log_fn, "⚠️ No Word (.docx) files specified for conversion.")
        return {"success": False, "converted_count": 0, "errors": ["No files selected"]}

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word_app = None
    converted_files = []
    failed_files = []

    try:
        _safe_log(log_fn, "📄 Initializing Microsoft Word COM Engine (Headless)...")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        # Word Constants
        wdExportFormatPDF = 17
        wdExportOptimizeForPrint = 0  # High quality 300+ DPI
        wdExportAllDocument = 0
        wdExportDocumentContent = 0
        wdExportCreateWordBookmarks = 1

        total = len(docx_files)
        mode_label = "Lossless Word -> PDF" + (" + Append Testsheet PDF" if append_testsheet else "")
        _safe_log(log_fn, f"🔄 Converting {total} selected Word document(s) [{mode_label}]...")

        for idx, docx_p in enumerate(docx_files, 1):
            if not os.path.exists(docx_p):
                _safe_log(log_fn, f"  ❌ [{idx}/{total}] File not found: {docx_p}")
                failed_files.append((docx_p, "File not found"))
                continue

            base_name = os.path.splitext(os.path.basename(docx_p))[0]
            if destination_folder:
                os.makedirs(destination_folder, exist_ok=True)
                pdf_p = os.path.join(destination_folder, base_name + ".pdf")
            else:
                pdf_p = os.path.join(os.path.dirname(docx_p), base_name + ".pdf")

            _safe_log(log_fn, f"  [{idx}/{total}] Converting: {os.path.basename(docx_p)}...")

            try:
                # Open document in ReadOnly mode
                doc = word_app.Documents.Open(os.path.abspath(docx_p), ReadOnly=True, ConfirmConversions=False)

                # ENFORCE LOSSLESS IMAGE RESOLUTION (Zero Compression for IR & DC photos)
                try:
                    doc.DoNotCompressPictures = True
                except Exception:
                    pass

                # Temporary staging path if merging testsheet
                temp_word_pdf = pdf_p + ".tmp.pdf" if append_testsheet else pdf_p

                # Export to PDF
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(temp_word_pdf),
                    ExportFormat=wdExportFormatPDF,
                    OpenAfterExport=False,
                    OptimizeFor=wdExportOptimizeForPrint,
                    Range=wdExportAllDocument,
                    From=1,
                    To=1,
                    Item=wdExportDocumentContent,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=wdExportCreateWordBookmarks,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False
                )

                doc.Close(SaveChanges=False)

                # Phase 2: Append Testsheet PDF at the back if requested
                if append_testsheet:
                    ts_pdf = find_matching_testsheet_pdf(docx_p, testsheet_pdf_dir)
                    if ts_pdf and os.path.exists(ts_pdf):
                        _safe_log(log_fn, f"    📎 Found Testsheet PDF: {os.path.basename(ts_pdf)}")
                        final_merged_pdf = pdf_p
                        temp_merged_pdf = pdf_p + ".merged.pdf"

                        success = merge_pdf_files(temp_word_pdf, ts_pdf, temp_merged_pdf)
                        if success:
                            if os.path.exists(final_merged_pdf):
                                try: os.remove(final_merged_pdf)
                                except Exception: pass
                            os.rename(temp_merged_pdf, final_merged_pdf)
                            try: os.remove(temp_word_pdf)
                            except Exception: pass
                            _safe_log(log_fn, f"    ✓ [MERGED] Quick Report + Testsheet PDF -> {os.path.basename(final_merged_pdf)} ({round(os.path.getsize(final_merged_pdf)/1024, 1)} KB)")
                        else:
                            if os.path.exists(temp_word_pdf):
                                if os.path.exists(final_merged_pdf):
                                    try: os.remove(final_merged_pdf)
                                    except Exception: pass
                                os.rename(temp_word_pdf, final_merged_pdf)
                            _safe_log(log_fn, f"    ⚠️ Warning: Failed to merge testsheet PDF. Kept report PDF.")
                    else:
                        # Fallback if no matching testsheet found
                        if os.path.exists(temp_word_pdf):
                            if os.path.exists(pdf_p):
                                try: os.remove(pdf_p)
                                except Exception: pass
                            os.rename(temp_word_pdf, pdf_p)
                        _safe_log(log_fn, f"    ⚠️ No matching testsheet found in {testsheet_pdf_dir}. Generated report PDF only.")
                else:
                    _safe_log(log_fn, f"    ✓ [DONE] Lossless PDF created ({round(os.path.getsize(pdf_p)/1024, 1)} KB)")

                converted_files.append(pdf_p)

            except Exception as doc_err:
                _safe_log(log_fn, f"    ❌ Failed: {doc_err}")
                failed_files.append((docx_p, str(doc_err)))

    except Exception as app_err:
        _safe_log(log_fn, f"❌ Word COM initialization failed: {app_err}")
        return {"success": False, "converted_count": len(converted_files), "errors": [str(app_err)]}

    finally:
        if word_app:
            try:
                word_app.Quit(SaveChanges=False)
            except Exception:
                pass
            word_app = None
        pythoncom.CoUninitialize()

    _safe_log(log_fn, f"\n🎉 Completed: {len(converted_files)}/{len(docx_files)} PDF(s) created successfully.")
    return {
        "success": len(failed_files) == 0,
        "converted_count": len(converted_files),
        "converted_files": converted_files,
        "failed_files": failed_files
    }

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
        dest = sys.argv[2] if len(sys.argv) > 2 else None
        append_ts = "--append-testsheet" in sys.argv
        ts_dir = sys.argv[3] if len(sys.argv) > 3 and not sys.argv[3].startswith("--") else None
        if os.path.isfile(target):
            files = [target]
        elif os.path.isdir(target):
            files = sorted(glob.glob(os.path.join(target, "*.docx")))
        else:
            files = []
        convert_word_files_to_pdf(files, destination_folder=dest, append_testsheet=append_ts, testsheet_pdf_dir=ts_dir)
    else:
        print("Usage: python convert_word_to_pdf.py <docx_file_or_folder> [destination_folder] [--append-testsheet] [testsheet_dir]")
