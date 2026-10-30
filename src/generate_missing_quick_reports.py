import os
import sys
import re
import time
import subprocess
from pathlib import Path
import openpyxl
import docx
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR = Path(r"C:\Users\aremi\Desktop\REPORT CBM\PO 42312390 - TNBD BANGI - IR US - TIANU")
CHECKLIST_PATH = BASE_DIR / "SUBSTATION_CHECKLIST.xlsx"
TEMPLATE_PATH = BASE_DIR / "99. GENERATE REPORT" / "TEMPLATE" / "99. template-quickreport-novi.docx"
RM_DIR = BASE_DIR / "03. RAW MATERIAL"
QR_DIR = BASE_DIR / "02. QUICK REPORT"
SCRIPT_DIR = BASE_DIR / "99. GENERATE REPORT" / "SCRIPT"

def sanitize_filename(name):
    """Clean filename by removing or replacing invalid Windows path characters."""
    name = str(name).replace('/', '-').replace('\\', '-')
    return re.sub(r'[<>:"|?*]', '', name).strip()

def remove_table_if_instruction(doc):
    """Remove instruction table at top of template if present."""
    if doc.tables:
        t0 = doc.tables[0]
        full_t0_text = "".join(cell.text for row in t0.rows for cell in row.cells)
        if "GENERATE CHECKLIST.bat" in full_t0_text or "C:\\Users" in full_t0_text:
            t0._element.getparent().remove(t0._element)

def save_docx_safe(doc, target_path):
    """Try saving docx; retry or fallback to alternate filename if file is locked."""
    try:
        doc.save(str(target_path))
        return target_path
    except PermissionError:
        time.sleep(0.5)
        try:
            doc.save(str(target_path))
            return target_path
        except PermissionError:
            # Fallback if locked by Word
            alt_path = target_path.parent / (target_path.stem + "_new.docx")
            doc.save(str(alt_path))
            return alt_path

def generate_missing_quick_reports():
    print("=" * 60)
    print("      GENERATE MISSING QUICK REPORTS (.DOCX)")
    print("=" * 60)
    
    # 1. Ensure SUBSTATION_CHECKLIST.xlsx exists
    if not CHECKLIST_PATH.exists():
        print("Substation checklist not found. Generating Excel checklist first...")
        subprocess.run([sys.executable, str(SCRIPT_DIR / "generate_excel.py")], check=True)

    # 2. Read Checklist Excel
    wb = openpyxl.load_workbook(CHECKLIST_PATH, data_only=True)
    
    if "Substation Checklist" not in wb.sheetnames or "Substation Info" not in wb.sheetnames:
        print("Error: Required sheets missing from SUBSTATION_CHECKLIST.xlsx")
        wb.close()
        return

    ws_check = wb["Substation Checklist"]
    ws_info = wb["Substation Info"]

    # Map info from Sheet 2 (Substation Info)
    info_map = {}
    for r in range(4, ws_info.max_row + 1):
        num_val = ws_info.cell(row=r, column=1).value
        if num_val is not None and str(num_val).isdigit():
            num = int(num_val)
            info_map[num] = {
                'num': num,
                'name': str(ws_info.cell(row=r, column=2).value or "").strip(),
                'type': str(ws_info.cell(row=r, column=3).value or "").strip(),
                'date': str(ws_info.cell(row=r, column=4).value or "").strip(),
                'gps': str(ws_info.cell(row=r, column=5).value or "").strip(),
                'erms_name': str(ws_info.cell(row=r, column=6).value or "").strip(),
                'site_name': str(ws_info.cell(row=r, column=7).value or "").strip(),
                'fl_no': str(ws_info.cell(row=r, column=8).value or ws_info.cell(row=r, column=9).value or "").strip()
            }

    # Find missing quick reports (Column AQ / Col 43 is "NO")
    missing_list = []
    for r in range(5, ws_check.max_row + 1):
        num_val = ws_check.cell(row=r, column=1).value
        qr_val = ws_check.cell(row=r, column=43).value # AQ is Col 43
        if num_val is not None and str(num_val).isdigit():
            num = int(num_val)
            if str(qr_val).strip().upper() == "NO":
                if num in info_map:
                    missing_list.append(info_map[num])

    wb.close()

    print(f"\nFound {len(missing_list)} substations with missing Quick Reports (AQ = NO).\n")

    if not missing_list:
        print("All Quick Reports are already present! Nothing to generate.")
        return

    if not TEMPLATE_PATH.exists():
        print(f"Error: Quick report template not found at:\n  {TEMPLATE_PATH}")
        return

    QR_DIR.mkdir(parents=True, exist_ok=True)

    # 3. Process each missing substation
    for idx, info in enumerate(missing_list, start=1):
        num = info['num']
        name = info['name']
        print(f"[{idx}/{len(missing_list)}] Processing Substation {num}: {name}...")

        # Find Raw Material directory
        clean_num = str(num).lstrip('0') or str(num)
        for entry in os.listdir(RM_DIR):
            m_d = re.match(r'^0*(\d+[a-zA-Z]?)[\.\s\-_]', entry)
            if m_d and m_d.group(1).upper() == clean_num.upper():
                if (RM_DIR / entry).is_dir():
                    rm_sub_folder = RM_DIR / entry
                    break
            elif (entry.startswith(f"{num}.") or entry.startswith(f"{clean_num}.")) and (RM_DIR / entry).is_dir():
                rm_sub_folder = RM_DIR / entry
                break

        digital_dir = None
        if rm_sub_folder:
            for sub in os.listdir(rm_sub_folder):
                if sub.upper() in ["DIGITAL", "DG"]:
                    digital_dir = rm_sub_folder / sub
                    break
            if not digital_dir and (rm_sub_folder / "IR").exists():
                pass # Direct under rm_sub_folder if no DIGITAL
            elif not digital_dir:
                digital_dir = rm_sub_folder

        # Map images 1..24
        image_map = {}
        if digital_dir and digital_dir.exists():
            all_files = os.listdir(digital_dir)
            def num_sort(f):
                base = os.path.splitext(f)[0]
                return int(base) if base.isdigit() else 9999
            sorted_files = [f for f in sorted(all_files, key=num_sort) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]

            for i in range(1, 25):
                found_path = None
                for ext in ['.jpg', '.JPG', '.jpeg', '.JPEG', '.png', '.PNG']:
                    p = digital_dir / f"{i}{ext}"
                    if p.exists():
                        found_path = p
                        break
                if not found_path and (i - 1) < len(sorted_files):
                    found_path = digital_dir / sorted_files[i - 1]
                image_map[i] = found_path

        # Open docx template
        doc = docx.Document(str(TEMPLATE_PATH))
        remove_table_if_instruction(doc)

        fl_val = str(info.get('fl_no', '')).strip()
        if not fl_val or fl_val in ['-', 'N/A', '', '.', '/', 'nan', 'None', 'none', '0', '0.0']:
            fl_val = "N/A"

        text_replacements = {
            "{{ substation_type }}": info['type'],
            "{{ fl_no }}": fl_val,
            "{{ substation_name_erms }}": info['erms_name'],
            "{{ substation_name_site }}": info['site_name'],
            "{{ gps_coordinate }}": info['gps'],
            "{{ date }}": info['date']
        }

        def replace_in_paragraph(p, key, val):
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
                        run.font.name = 'Tahoma'
                        run.font.size = docx.shared.Pt(10)
                # Fallback if placeholder was split across runs
                if key in p.text:
                    p.text = p.text.replace(key, str(val))
                    for run in p.runs:
                        run.font.name = 'Tahoma'
                        run.font.size = docx.shared.Pt(10)
            # Ensure zero spacing before and after paragraph
            p.paragraph_format.space_before = docx.shared.Pt(0)
            p.paragraph_format.space_after = docx.shared.Pt(0)

        # Text replacements in paragraphs
        for p in doc.paragraphs:
            p.paragraph_format.space_before = docx.shared.Pt(0)
            p.paragraph_format.space_after = docx.shared.Pt(0)
            for key, val in text_replacements.items():
                replace_in_paragraph(p, key, val)

        # Text and Image replacements in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        p.paragraph_format.space_before = docx.shared.Pt(0)
                        p.paragraph_format.space_after = docx.shared.Pt(0)
                        for key, val in text_replacements.items():
                            replace_in_paragraph(p, key, val)

                    # Replace image placeholders {{ visual_image1 }} .. {{ visual_image24 }}
                    for i in range(1, 25):
                        tag = f"{{{{ visual_image{i} }}}}"
                        if tag in cell.text:
                            cell.text = "" # Clear placeholder text
                            p = cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p.paragraph_format.space_before = docx.shared.Pt(0)
                            p.paragraph_format.space_after = docx.shared.Pt(0)
                            img_path = image_map.get(i)
                            if img_path and img_path.exists():
                                run = p.add_run()
                                run.add_picture(str(img_path), height=docx.shared.Cm(6))

        # Format output filename
        date_clean = info['date'].replace('/', '').replace('-', '')
        raw_sub_name = info['site_name'] or info['name']
        clean_sub_name = sanitize_filename(raw_sub_name)
        out_filename = f"{num}. {date_clean} {clean_sub_name} (VI).docx"
        out_docx_path = QR_DIR / out_filename

        saved_path = save_docx_safe(doc, out_docx_path)
        print(f"  Created: {saved_path.name}")

    # 4. Automatically re-run generate_excel.py to update checklist AQ values
    print("\nUpdating Excel checklist with newly generated reports...")
    try:
        subprocess.run([sys.executable, str(SCRIPT_DIR / "generate_excel.py")], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Note: Checklist update finished with notice ({e}).")

    print("\n" + "=" * 60)
    print("  ALL MISSING QUICK REPORTS GENERATED SUCCESSFULLY!")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    generate_missing_quick_reports()
