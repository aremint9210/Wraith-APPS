"""
Inject FLIR Radiometric Images and Visual Photos into CBM Word Reports.
Supports Part 01 (Summary/Overview) and Part 02-09 (Switchgear Panels 1-8).
Integrated with WRAITH APP and Standalone Execution.

Features:
- Substation-level single-pass Word COM batching (10x faster cold runs)
- Dynamic workspace detection via --workspace or BASE_DIR environment
- Safety backup to Fresh/ folder (preserves pristine original reports)
- Auto-extracts image mappings from Excel testsheets
- Invokes FLIR COM Word Add-In via 32-bit PowerShell engine in a single batch
- Injects FLIR CIRViewer ActiveX controls (.bin) and thumbnails (.wmf)
- Injects paired digital visual camera photos
- Two-way dynamic link synchronization: DOCVARIABLE fields, settings.xml docVars, docProps/custom.xml
- Strictly leaves reports in editable Word format (no PDF conversion)
"""

import os
import sys
import re
import json
import shutil
import zipfile
import subprocess
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Optional

# Ensure standard output and error support full Unicode
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr.encoding != 'utf-8':
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl

# ==============================================================================
# CONFIGURATION & DYNAMIC PATH RESOLUTION
# ==============================================================================
SCRIPT_DIR = Path(__file__).resolve().parent

def resolve_paths(workspace_override: Optional[Path] = None):
    """Dynamically resolves workspace and folder locations."""
    if workspace_override and Path(workspace_override).exists():
        base = Path(workspace_override).resolve()
    elif os.environ.get("BASE_DIR") and Path(os.environ.get("BASE_DIR")).exists():
        base = Path(os.environ.get("BASE_DIR")).resolve()
    else:
        cwd = Path.cwd().resolve()
        if cwd.name == "01. FULL REPORT":
            base = cwd.parent
        else:
            base = cwd

    # Check if base itself is 01. FULL REPORT or PO root
    if base.name == "01. FULL REPORT":
        po_root = base.parent
        full_report_dir = base
    else:
        po_root = base
        full_report_dir = po_root / "01. FULL REPORT"
        if not full_report_dir.exists():
            full_report_dir = po_root

    fresh_dir = full_report_dir / "Fresh"
    temp_dir = full_report_dir / ".flir_temp"
    cache_dir = temp_dir / "cache"

    testsheet_dir = po_root / "04. TESTSHEET"
    raw_material_dir = po_root / "03. RAW MATERIAL"
    
    # Template directory search
    template_dir = po_root / "05. TEMPLATE" / "33kV"
    if not template_dir.exists():
        template_dir = po_root / "05. TEMPLATE"
    if not template_dir.exists():
        # Fallback to WRAITH APP templates
        cand = SCRIPT_DIR.parent / "templates" / "33kV"
        if cand.exists():
            template_dir = cand
        else:
            template_dir = SCRIPT_DIR.parent / "templates"

    # FLIR helper script search
    flir_helper = SCRIPT_DIR / "flir_word_report_helper.ps1"
    if not flir_helper.exists():
        cand = Path(r"C:\Users\aremi\Desktop\REPORT CBM\WRAITH APP\src\flir_word_report_helper.ps1")
        if cand.exists():
            flir_helper = cand

    ps32 = Path(r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")

    return {
        "po_root": po_root,
        "full_report_dir": full_report_dir,
        "fresh_dir": fresh_dir,
        "temp_dir": temp_dir,
        "cache_dir": cache_dir,
        "testsheet_dir": testsheet_dir,
        "raw_material_dir": raw_material_dir,
        "template_dir": template_dir,
        "flir_helper": flir_helper,
        "ps32": ps32
    }


PANEL_COMPARTMENT_NAMES = [
    "BREAKER COMPARTMENT",
    "CABLE COMPARTMENT",
    "BUSBAR COMPARTMENT",
    "PT COMPARTMENT",
    "CABLE ENTRY",
    "LINK BOX",
    "SECONDARY COMPARTMENT"
]

OVERVIEW_NAMES = [
    "LHS Front",
    "RHS Front",
    "LHS Back",
    "RHS Back",
    "LHS Top",
    "RHS Top",
    "LHS Cable Cellar",
    "RHS Cable Cellar",
    "LHS MBB",
    "RHS MBB"
]


# ==============================================================================
# SAFETY BACKUP & UTILITIES
# ==============================================================================
def ensure_fresh_backups(paths: dict):
    """Ensures original reports in 01. FULL REPORT are backed up to Fresh/ untouched."""
    fresh_dir = paths["fresh_dir"]
    full_report_dir = paths["full_report_dir"]
    fresh_dir.mkdir(parents=True, exist_ok=True)
    all_docx = [f for f in full_report_dir.glob("*.docx") if not f.name.startswith("~")]
    backed_up_count = 0
    for doc_path in all_docx:
        target_backup = fresh_dir / doc_path.name
        if not target_backup.exists():
            shutil.copy2(doc_path, target_backup)
            backed_up_count += 1
    if backed_up_count > 0:
        print(f"[SAFETY] Archived {backed_up_count} fresh report(s) into: Fresh\\")
    else:
        print(f"[SAFETY] Verified all reports in Fresh\\ backup folder.")


def is_valid_flir_docx(file_path: Path) -> bool:
    """Verify that a generated FLIR single-page docx exists and has the activeX1.bin binary."""
    if not file_path.exists() or file_path.stat().st_size < 100000:
        return False
    try:
        with zipfile.ZipFile(file_path, "r") as z:
            return "word/activeX/activeX1.bin" in z.namelist()
    except Exception:
        return False


def clean_page_breaks(doc):
    """
    Cleans up empty paragraphs between tables and page-break paragraphs,
    and sets page-break paragraphs to a tiny font (2pt) with zero spacing
    to strictly maintain a clean 1-page-per-compartment layout.
    """
    body = doc.element.body
    children = list(body)
    i = 0
    while i < len(children):
        child = children[i]
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            txt = child.text or ""
            brs = child.findall('.//' + qn('w:br'))
            has_page_break = any(b.get(qn('w:type')) == 'page' for b in brs)
            prev_is_tbl = (i > 0 and children[i-1].tag.split('}')[-1] == 'tbl')
            next_is_p_with_br = (i + 1 < len(children) and any(b.get(qn('w:type')) == 'page' for b in children[i+1].findall('.//' + qn('w:br'))))
            if prev_is_tbl and not has_page_break and next_is_p_with_br and not txt.strip():
                body.remove(child)
                children = list(body)
                continue
        i += 1

    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            brs = child.findall('.//' + qn('w:br'))
            if any(b.get(qn('w:type')) == 'page' for b in brs):
                pPr = child.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    child.insert(0, pPr)
                spacing = pPr.find(qn('w:spacing'))
                if spacing is None:
                    spacing = OxmlElement('w:spacing')
                    pPr.append(spacing)
                spacing.set(qn('w:after'), '0')
                spacing.set(qn('w:before'), '0')
                spacing.set(qn('w:line'), '20')
                spacing.set(qn('w:lineRule'), 'exact')
                rPr = pPr.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    pPr.append(rPr)
                sz = rPr.find(qn('w:sz'))
                if sz is None:
                    sz = OxmlElement('w:sz')
                    rPr.append(sz)
                sz.set(qn('w:val'), '2')

    children = list(body)
    last_tbl_indices = [idx for idx, c in enumerate(children) if c.tag.endswith('tbl')]
    if last_tbl_indices:
        last_tbl_idx = max(last_tbl_indices)
        empty_after = [c for c in children[last_tbl_idx+1:] if c.tag.endswith('p') and not (c.text or '').strip()]
        if len(empty_after) > 1:
            for ep in empty_after[1:]:
                try:
                    body.remove(ep)
                except Exception:
                    pass


# ==============================================================================
# TESTSHEET & IMAGE RESOLUTION
# ==============================================================================
def resolve_image_pair(ir_dir: Path, img_val) -> Tuple[Optional[Path], Optional[Path]]:
    """Resolves matching IR (.jpg) and visual photo (.jpg) files."""
    if not img_val or str(img_val).strip() == "" or str(img_val).lower() == "none":
        return None, None

    m = re.search(r"(\d+)", str(img_val))
    if not m:
        return None, None
    num_int = int(m.group(1))

    candidates = [
        f"FLIR0{num_int:03d}.jpg",
        f"FLIR{num_int:04d}.jpg",
        f"FLIR0{num_int}.jpg",
        f"FLIR{num_int}.jpg"
    ]
    ir_file = None
    for cand in candidates:
        p = ir_dir / cand
        if p.exists():
            ir_file = p
            break

    if not ir_file:
        for f in ir_dir.glob("*.jpg"):
            if "- photo" not in f.name.lower() and "_photo" not in f.name.lower():
                num_match = re.search(r"(\d+)", f.name)
                if num_match and int(num_match.group(1)) == num_int:
                    ir_file = f
                    break

    if not ir_file:
        return None, None

    dc_file = None
    stem = ir_file.stem
    for dc_name in [f"{stem}- photo.jpg", f"{stem}-photo.jpg", f"{stem}_photo.jpg", f"{stem} - photo.jpg"]:
        p = ir_dir / dc_name
        if p.exists():
            dc_file = p
            break

    if not dc_file:
        for f in ir_dir.glob("*.jpg"):
            if "- photo" in f.name.lower() or "_photo" in f.name.lower():
                num_match = re.search(r"(\d+)", f.name)
                if num_match and int(num_match.group(1)) == num_int:
                    dc_file = f
                    break

    return ir_file, dc_file


def parse_testsheet(xlsx_path: Path):
    """Parses a CBM Switchgear Excel testsheet to extract overview & panel mappings."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_name = None
    for name in wb.sheetnames:
        if '33kv' in name.lower() or 'main' in name.lower() or 'switchgear' in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    overviews = {}
    overview_map = {
        ('lhs', 'front'): 0,
        ('rhs', 'front'): 1,
        ('lhs', 'back'): 2,
        ('rhs', 'back'): 3,
        ('lhs', 'top'): 4,
        ('rhs', 'top'): 5,
        ('lhs', 'cable cellar'): 6,
        ('rhs', 'cable cellar'): 7,
        ('lhs', 'mbb'): 8,
        ('rhs', 'mbb'): 9,
    }
    for r in range(63, 75):
        side_val = str(ws.cell(r, 3).value or '').lower()
        comp_val = str(ws.cell(r, 11).value or '').lower().strip()
        img_val = ws.cell(r, 16).value
        side = 'lhs' if 'lhs' in side_val else ('rhs' if 'rhs' in side_val else None)
        if not side:
            for prev_r in range(r, 62, -1):
                p_side = str(ws.cell(prev_r, 3).value or '').lower()
                if 'lhs' in p_side: side = 'lhs'; break
                if 'rhs' in p_side: side = 'rhs'; break
        for (k_side, k_comp), idx in overview_map.items():
            if k_side == side and k_comp in comp_val and img_val:
                overviews[idx] = img_val

    panels = {}
    for p_idx in range(1, 9):
        r_start = 15 + (p_idx - 1) * 6
        panel_no = ws.cell(r_start, 3).value
        feeder = ws.cell(r_start, 4).value
        comps = {}
        for off, comp_name in enumerate(['BREAKER', 'CABLE', 'BUSBAR', 'PT', 'CABLE ENTRY', 'LINK BOX']):
            curr_r = r_start + off
            img = ws.cell(curr_r, 16).value
            if img:
                comps[off] = img
        rem = ws.cell(r_start, 17).value
        if rem and 'sp:' in str(rem).lower():
            comps[6] = rem
        panels[p_idx] = {'panel_no': panel_no, 'feeder': feeder, 'comps': comps}

    return overviews, panels


def find_testsheet_for_substation(testsheet_dir: Path, substation_name: str) -> Optional[Path]:
    """Finds matching .xlsx testsheet in 04. TESTSHEET folder."""
    if not testsheet_dir.exists():
        return None
    norm_sub = re.sub(r'^\d+\.\s*', '', substation_name).strip().lower()
    norm_sub = re.sub(r'\s*\([^)]*\)', '', norm_sub).strip()

    best_match = None
    for f in testsheet_dir.glob("*.xlsx"):
        if f.name.startswith("~"):
            continue
        f_norm = f.stem.lower()
        if norm_sub in f_norm:
            return f
        if any(w in f_norm for w in norm_sub.split() if len(w) > 3):
            best_match = f
    return best_match


def find_raw_ir_dir_for_substation(raw_material_dir: Path, substation_name: str) -> Optional[Path]:
    """Finds matching RAW DATA/IR folder in 03. RAW MATERIAL."""
    if not raw_material_dir.exists():
        return None
    norm_sub = re.sub(r'^\d+\.\s*', '', substation_name).strip().lower()
    norm_sub = re.sub(r'\s*\([^)]*\)', '', norm_sub).strip()

    matched_dir = None
    for d in raw_material_dir.iterdir():
        if d.is_dir():
            d_norm = d.name.lower()
            if norm_sub in d_norm:
                matched_dir = d
                break

    if not matched_dir:
        for d in raw_material_dir.iterdir():
            if d.is_dir():
                d_norm = d.name.lower()
                if any(w in d_norm for w in norm_sub.split() if len(w) > 3):
                    matched_dir = d
                    break

    if matched_dir:
        ir_candidate = matched_dir / "RAW DATA" / "IR"
        if ir_candidate.exists():
            return ir_candidate
        ir_candidate2 = matched_dir / "IR"
        if ir_candidate2.exists():
            return ir_candidate2
        return matched_dir

    return None


def scan_available_reports(paths: dict):
    """Scans 01. FULL REPORT for .docx files and groups them by Substation and Part."""
    full_report_dir = paths["full_report_dir"]
    all_docx = sorted(list(full_report_dir.glob("*.docx")))
    substations = {}

    for doc_path in all_docx:
        if doc_path.name.startswith("~"):
            continue

        m = re.match(r"^(.*?)\s*-\s*Part\s+(\d+)\s*-\s*(.*?)\.docx$", doc_path.name, re.IGNORECASE)
        if m:
            sub_name = m.group(1).strip()
            part_num = int(m.group(2))
            part_title = m.group(3).strip()
        else:
            sub_name = doc_path.stem.split(" - Part ")[0].strip()
            part_num = 99
            part_title = doc_path.stem

        if sub_name not in substations:
            testsheet = find_testsheet_for_substation(paths["testsheet_dir"], sub_name)
            raw_ir = find_raw_ir_dir_for_substation(paths["raw_material_dir"], sub_name)
            substations[sub_name] = {
                "name": sub_name,
                "testsheet": testsheet,
                "raw_ir": raw_ir,
                "parts": {}
            }

        substations[sub_name]["parts"][part_num] = {
            "part_num": part_num,
            "title": f"Part {part_num:02d} - {part_title}",
            "path": doc_path,
            "is_injectable": (1 <= part_num <= 9)
        }

    return substations


# ==============================================================================
# BATCH FLIR COM ENGINE (OPTIMIZED SINGLE WORD SESSION)
# ==============================================================================
def execute_batch_flir_generation(tasks: List[dict], paths: dict):
    """
    Executes FLIR page generation in a SINGLE Microsoft Word COM session
    for all uncached items, slashing cold-run time by up to 10x!
    """
    cache_dir = paths["cache_dir"]
    cache_dir.mkdir(parents=True, exist_ok=True)
    ps32 = paths["ps32"] if paths["ps32"].exists() else Path("powershell.exe")
    flir_helper = paths["flir_helper"]

    # Filter out already valid cached pages
    uncached_tasks = []
    for t in tasks:
        out_path = cache_dir / f"{t['cache_key']}.docx"
        t["out_path"] = out_path
        if not is_valid_flir_docx(out_path):
            uncached_tasks.append(t)

    if not uncached_tasks:
        return

    print(f"\n[BATCH COM] Launching single Word session for {len(uncached_tasks)} missing FLIR page(s)...")

    # Build manifest array
    manifest_array = []
    for t in uncached_tasks:
        manifest_array.append({
            "template_path": str(t["template_path"].resolve()),
            "output_path": str(t["out_path"].resolve()),
            "pages": [{
                "template_page_no": 1,
                "ir_path": str(t["ir_path"].resolve()),
                "visual_path": str(t["visual_path"].resolve()) if (t["visual_path"] and t["visual_path"].exists()) else "",
                "field_overrides": {}
            }]
        })

    manifest_json_path = cache_dir / f"batch_manifest_{os.getpid()}.json"
    with open(manifest_json_path, "w", encoding="utf-8") as f:
        json.dump(manifest_array, f, indent=2)

    cmd = [
        str(ps32),
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", str(flir_helper.resolve()),
        "-ManifestPath", str(manifest_json_path.resolve())
    ]

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print(f"[BATCH COM] Completed FLIR generation pass successfully.")
    except subprocess.CalledProcessError as e:
        print(f"[BATCH COM ERROR] {e.stderr or e.stdout}")
    finally:
        if manifest_json_path.exists():
            try:
                manifest_json_path.unlink()
            except Exception:
                pass


def extract_flir_data(flir_docx_path: Path):
    """Extracts FLIR binary stream, WMF thumbnail, and DOCVARIABLEs with temperatures."""
    with zipfile.ZipFile(flir_docx_path, "r") as z:
        ax_bin = z.read("word/activeX/activeX1.bin")
        wmf_bytes = z.read("word/media/image2.wmf")
        doc_xml_str = z.read("word/document.xml").decode("utf-8")
        settings_xml_str = z.read("word/settings.xml").decode("utf-8")

    vars_found = re.findall(r'DOCVARIABLE\s+(_Fd\w+)', doc_xml_str)
    var_list = []
    for v_name in vars_found:
        m_val = re.search(r'w:name="' + re.escape(v_name) + r'"\s+w:val="([^"]*)"', settings_xml_str)
        val_str = m_val.group(1) if m_val else "- "
        var_list.append((v_name, val_str))

    return {
        "bin": ax_bin,
        "wmf": wmf_bytes,
        "vars": var_list
    }


# ==============================================================================
# INJECTION ENGINES (PART 01 SUMMARY & PART 02-09 PANELS)
# ==============================================================================
def safe_save_target(temp_final: Path, target_docx: Path):
    """Safely overwrites target_docx with temp_final, handling Word file locks gracefully."""
    import time
    for attempt in range(1, 4):
        try:
            shutil.copy2(temp_final, target_docx)
            print(f"\n[SUCCESS] Document saved: {target_docx.name} ({target_docx.stat().st_size:,} bytes)")
            return True
        except PermissionError:
            if attempt < 3:
                print(f"  [FILE LOCKED] '{target_docx.name}' is open in Microsoft Word. Retrying in 2 seconds (please close Word)...")
                time.sleep(2)
            else:
                fallback = target_docx.parent / f"{target_docx.stem}_INJECTED{target_docx.suffix}"
                try:
                    shutil.copy2(temp_final, fallback)
                    print(f"\n[WARNING] Could not overwrite '{target_docx.name}' (file is open in Word).")
                    print(f"          Saved injected report to: {fallback.name}")
                    return True
                except Exception as ex:
                    print(f"\n[ERROR] Failed to save document: {ex}")
                    return False

def inject_part01_summary(target_docx: Path, raw_ir_dir: Path, overviews_map: Dict[int, any], paths: dict):
    """Injects 10 FLIR Overview thermograms and digital photos into Part 01 Summary."""
    print(f"\n=======================================================================")
    print(f" INJECTING PART 01: SUMMARY / OVERVIEW INSPECTION SHEETS")
    print(f"=======================================================================")

    fresh_backup = paths["fresh_dir"] / target_docx.name
    base_source = fresh_backup if fresh_backup.exists() else target_docx
    cache_dir = paths["cache_dir"]
    temp_dir = paths["temp_dir"]
    tmpl_overview = paths["template_dir"] / "swg-overview.docx"

    # Step 1: Resolve all required overview items
    overview_tasks = []
    for table_idx in range(10):
        img_val = overviews_map.get(table_idx)
        if not img_val:
            continue
        ir_path, dc_path = resolve_image_pair(raw_ir_dir, img_val)
        if not ir_path:
            print(f"  Overview {table_idx+1:2d} ({OVERVIEW_NAMES[table_idx]}): IR missing ({img_val})")
            continue
        cache_key = f"overview_{table_idx}_{ir_path.stem}"
        overview_tasks.append({
            "table_idx": table_idx,
            "template_path": tmpl_overview,
            "ir_path": ir_path,
            "visual_path": dc_path,
            "cache_key": cache_key,
            "out_path": cache_dir / f"{cache_key}.docx"
        })

    # Step 2: Batch generate any uncached FLIR pages
    execute_batch_flir_generation(overview_tasks, paths)

    # Step 3: Extract FLIR data from cache
    flir_results = {}
    for task in overview_tasks:
        if is_valid_flir_docx(task["out_path"]):
            data = extract_flir_data(task["out_path"])
            var_name, val_str = data["vars"][0] if data["vars"] else ("_FdPlaceholder", "- ")
            flir_results[task["table_idx"]] = {
                "bin": data["bin"],
                "wmf": data["wmf"],
                "var_name": var_name,
                "val_str": val_str,
                "dc_path": task["visual_path"],
                "ir_name": task["ir_path"].name
            }
            print(f"  Overview {task['table_idx']+1:2d} ({OVERVIEW_NAMES[task['table_idx']]}): {var_name} = '{val_str.strip()}' | {len(data['bin']):,} bytes")

    # Step 4: Inject Visual Digital Photos
    print("\n[STEP 2] Injecting Visual Digital Photos into Inspection Sheets...")
    doc = docx.Document(base_source)

    inspection_tbl_indices = []
    for i, t in enumerate(doc.tables):
        if t._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
            inspection_tbl_indices.append(i)

    for slot_idx, res in flir_results.items():
        if res["dc_path"] and res["dc_path"].exists():
            actual_tbl_idx = inspection_tbl_indices[slot_idx] if slot_idx < len(inspection_tbl_indices) else slot_idx
            tbl = doc.tables[actual_tbl_idx]
            cell_vis = tbl.rows[9].cells[12]
            cell_vis.text = ""
            p = cell_vis.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(str(res["dc_path"]), width=Cm(8.37), height=Cm(5.20))
            print(f"  Overview {slot_idx+1:2d} ({OVERVIEW_NAMES[slot_idx]}): Injected {res['dc_path'].name}")

    clean_page_breaks(doc)
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_base = temp_dir / "temp_part01_with_visuals.docx"
    doc.save(temp_base)

    # Step 5: OpenXML Assembly & Two-Way FLIR Link Splicing
    print("\n[STEP 3] Synchronizing Two-Way FLIR-to-Table Link & Assembling Package...")
    temp_final = temp_dir / "temp_part01_assembled.docx"

    with zipfile.ZipFile(temp_base, "r") as zin:
        import xml.etree.ElementTree as ET
        rels_xml = zin.read("word/_rels/document.xml.rels")
        rels_tree = ET.fromstring(rels_xml)
        r_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels_tree}

        doc_xml_raw = zin.read("word/document.xml")
        doc_tree = ET.fromstring(doc_xml_raw)
        tables_xml = doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")

        doc_inspection_indices = []
        for i, t in enumerate(tables_xml):
            if t.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
                doc_inspection_indices.append(i)

        ax_targets = {}
        wmf_targets = {}
        for slot_idx, t_idx in enumerate(doc_inspection_indices):
            t_el = tables_xml[t_idx]
            for c in t_el.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
                r_id = c.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = r_map.get(r_id)
                if target and 'activeX' in target:
                    bin_target = 'word/' + target.replace('.xml', '.bin') if not target.startswith('word/') else target.replace('.xml', '.bin')
                    ax_targets[slot_idx] = bin_target
            for v in t_el.findall('.//{urn:schemas-microsoft-com:vml}imagedata'):
                r_id = v.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = r_map.get(r_id)
                if target and 'image' in target and target.endswith('.wmf'):
                    norm_target = 'word/' + target if not target.startswith('word/') else target
                    wmf_targets[slot_idx] = norm_target

        with zipfile.ZipFile(temp_final, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item_info in zin.infolist():
                fname = item_info.filename
                data = zin.read(fname)

                matched_ax = False
                for slot_idx, ax_target in ax_targets.items():
                    if fname == ax_target and slot_idx in flir_results:
                        zout.writestr(item_info, flir_results[slot_idx]["bin"])
                        matched_ax = True
                        break
                if matched_ax: continue

                matched_wmf = False
                for slot_idx, w_target in wmf_targets.items():
                    if fname == w_target and slot_idx in flir_results:
                        zout.writestr(item_info, flir_results[slot_idx]["wmf"])
                        matched_wmf = True
                        break
                if matched_wmf: continue

                if fname == "word/document.xml":
                    doc_xml = data.decode("utf-8")
                    field_pattern = re.compile(
                        r'(<w:instrText[^>]*>\s*DOCVARIABLE\s+)(_Fd\w+)(.*?w:fldCharType=.separate..*?<w:t[^>]*>)([^<]*)(</w:t>)',
                        re.DOTALL
                    )
                    matches = list(field_pattern.finditer(doc_xml))
                    if len(matches) == 10:
                        parts = []
                        last_idx = 0
                        for idx, m in enumerate(matches):
                            parts.append(doc_xml[last_idx:m.start()])
                            if idx in flir_results:
                                new_var = flir_results[idx]["var_name"]
                                new_val = flir_results[idx]["val_str"]
                            else:
                                new_var = m.group(2)
                                new_val = m.group(4)
                            parts.append(m.group(1) + new_var + m.group(3) + new_val + m.group(5))
                            last_idx = m.end()
                        parts.append(doc_xml[last_idx:])
                        doc_xml = "".join(parts)
                        print(f"  [LINK] Synchronized {len(flir_results)} DOCVARIABLE fields to FLIR ActiveX controls")
                    data = doc_xml.encode("utf-8")

                elif fname == "word/settings.xml":
                    settings_xml = data.decode("utf-8")
                    new_docvars = []
                    for t_idx, res in flir_results.items():
                        new_docvars.append(f'<w:docVar w:name="{res["var_name"]}" w:val="{res["val_str"]}"/>')
                    docvar_tags = "".join(new_docvars)
                    if "<w:docVars>" in settings_xml:
                        settings_xml = settings_xml.replace("<w:docVars>", f"<w:docVars>{docvar_tags}")
                    else:
                        settings_xml = settings_xml.replace("</w:settings>", f"<w:docVars>{docvar_tags}</w:docVars></w:settings>")
                    print(f"  [LINK] Registered {len(new_docvars)} dynamic variables in word/settings.xml")
                    data = settings_xml.encode("utf-8")

                elif fname == "_rels/.rels":
                    rels_xml = data.decode("utf-8")
                    if "custom-properties" not in rels_xml:
                        rels_xml = rels_xml.replace(
                            "</Relationships>",
                            '<Relationship Id="rIdFlirCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'
                        )
                    data = rels_xml.encode("utf-8")

                elif fname == "[Content_Types].xml":
                    ct_xml = data.decode("utf-8")
                    if "/docProps/custom.xml" not in ct_xml:
                        ct_xml = ct_xml.replace(
                            "</Types>",
                            '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'
                        )
                    data = ct_xml.encode("utf-8")

                zout.writestr(item_info, data)

            custom_props_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="FLIR"><vt:lpwstr>FLIR</vt:lpwstr></property>'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="FlirTempUnit"><vt:i4>0</vt:i4></property>'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="FlirDistUnit"><vt:i4>0</vt:i4></property>'
                '</Properties>'
            )
            zout.writestr("docProps/custom.xml", custom_props_xml.encode("utf-8"))

    safe_save_target(temp_final, target_docx)


def inject_part_panel(target_docx: Path, panel_idx: int, panel_info: dict, raw_ir_dir: Path, paths: dict):
    """Injects 7 compartment FLIR thermograms, digital photos, and 28 variables into Panel reports."""
    panel_no = panel_info["panel_no"]
    feeder = panel_info["feeder"]
    comps_map = panel_info["comps"]

    print(f"\n=======================================================================")
    print(f" INJECTING PANEL {panel_idx}: {panel_no} ({feeder})")
    print(f"=======================================================================")

    fresh_backup = paths["fresh_dir"] / target_docx.name
    base_source = fresh_backup if fresh_backup.exists() else target_docx
    cache_dir = paths["cache_dir"]
    temp_dir = paths["temp_dir"]

    tmpl_main = paths["template_dir"] / "swg-panel.docx"
    tmpl_linkbox = paths["template_dir"] / "swg-panel - Linkbox.docx"
    tmpl_secondary = paths["template_dir"] / "swg-panel - Secondary Panel.docx"

    # Step 1: Resolve all required compartment items
    panel_tasks = []
    for comp_idx in range(7):
        img_val = comps_map.get(comp_idx)
        if not img_val:
            continue
        ir_path, dc_path = resolve_image_pair(raw_ir_dir, img_val)
        if not ir_path:
            print(f"  Comp {comp_idx+1:2d} ({PANEL_COMPARTMENT_NAMES[comp_idx]}): IR missing ({img_val})")
            continue

        if comp_idx == 5:
            tmpl = tmpl_linkbox
        elif comp_idx == 6:
            tmpl = tmpl_secondary
        else:
            tmpl = tmpl_main

        cache_key = f"panel{panel_idx}_comp{comp_idx}_{ir_path.stem}"
        panel_tasks.append({
            "comp_idx": comp_idx,
            "template_path": tmpl,
            "ir_path": ir_path,
            "visual_path": dc_path,
            "cache_key": cache_key,
            "out_path": cache_dir / f"{cache_key}.docx"
        })

    # Step 2: Batch generate any uncached FLIR pages
    execute_batch_flir_generation(panel_tasks, paths)

    # Step 3: Extract FLIR data from cache
    flir_results = {}
    for task in panel_tasks:
        if is_valid_flir_docx(task["out_path"]):
            data = extract_flir_data(task["out_path"])
            flir_results[task["comp_idx"]] = {
                "bin": data["bin"],
                "wmf": data["wmf"],
                "vars": data["vars"],
                "dc_path": task["visual_path"],
                "ir_name": task["ir_path"].name
            }
            var_summary = ", ".join(f"{v[0]}='{v[1].strip()}'" for v in data["vars"])
            print(f"  Comp {task['comp_idx']+1:2d} ({PANEL_COMPARTMENT_NAMES[task['comp_idx']]}): Injected {task['ir_path'].name} | [{var_summary}]")

    # Step 4: Inject Visual Digital Photos
    print("\n[STEP 2] Injecting Visual Digital Photos into Compartment Sheets...")
    doc = docx.Document(base_source)

    inspection_tbl_indices = []
    for i, t in enumerate(doc.tables):
        if t._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
            inspection_tbl_indices.append(i)

    for comp_idx, res in flir_results.items():
        if res["dc_path"] and res["dc_path"].exists():
            actual_tbl_idx = inspection_tbl_indices[comp_idx] if comp_idx < len(inspection_tbl_indices) else comp_idx
            tbl = doc.tables[actual_tbl_idx]
            cell_vis = tbl.rows[9].cells[12]
            cell_vis.text = ""
            p = cell_vis.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(str(res["dc_path"]), width=Cm(8.37), height=Cm(5.20))
            print(f"  Comp {comp_idx+1:2d} ({PANEL_COMPARTMENT_NAMES[comp_idx]}): Injected {res['dc_path'].name}")

    clean_page_breaks(doc)
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_base = temp_dir / f"temp_panel{panel_idx}_with_visuals.docx"
    doc.save(temp_base)

    # Step 5: OpenXML Dynamic Object Discovery & Two-Way Synchronization
    print("\n[STEP 3] Synchronizing Two-Way FLIR-to-Table Link & Assembling Package...")
    temp_final = temp_dir / f"temp_panel{panel_idx}_assembled.docx"

    with zipfile.ZipFile(temp_base, "r") as zin:
        rels_xml = zin.read("word/_rels/document.xml.rels")
        import xml.etree.ElementTree as ET
        rels_tree = ET.fromstring(rels_xml)
        r_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels_tree}

        doc_xml_raw = zin.read("word/document.xml")
        doc_tree = ET.fromstring(doc_xml_raw)
        tables_xml = doc_tree.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")

        doc_inspection_indices = []
        for i, t in enumerate(tables_xml):
            if t.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
                doc_inspection_indices.append(i)

        ax_targets = {}
        wmf_targets = {}
        for comp_idx, t_idx in enumerate(doc_inspection_indices):
            t_el = tables_xml[t_idx]
            for c in t_el.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}control'):
                r_id = c.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = r_map.get(r_id)
                if target and 'activeX' in target:
                    bin_target = 'word/' + target.replace('.xml', '.bin') if not target.startswith('word/') else target.replace('.xml', '.bin')
                    ax_targets[comp_idx] = bin_target
            for v in t_el.findall('.//{urn:schemas-microsoft-com:vml}imagedata'):
                r_id = v.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = r_map.get(r_id)
                if target and 'image' in target and target.endswith('.wmf'):
                    norm_target = 'word/' + target if not target.startswith('word/') else target
                    wmf_targets[comp_idx] = norm_target

        with zipfile.ZipFile(temp_final, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item_info in zin.infolist():
                fname = item_info.filename
                data = zin.read(fname)

                matched_ax = False
                for comp_idx, ax_target in ax_targets.items():
                    if fname == ax_target and comp_idx in flir_results:
                        zout.writestr(item_info, flir_results[comp_idx]["bin"])
                        matched_ax = True
                        break
                if matched_ax: continue

                matched_wmf = False
                for comp_idx, w_target in wmf_targets.items():
                    if fname == w_target and comp_idx in flir_results:
                        zout.writestr(item_info, flir_results[comp_idx]["wmf"])
                        matched_wmf = True
                        break
                if matched_wmf: continue

                if fname == "word/document.xml":
                    doc_xml = data.decode("utf-8")
                    field_pattern = re.compile(
                        r'(<w:instrText[^>]*>\s*DOCVARIABLE\s+)(_Fd\w+)(.*?w:fldCharType=.separate..*?<w:t[^>]*>)([^<]*)(</w:t>)',
                        re.DOTALL
                    )
                    matches = list(field_pattern.finditer(doc_xml))

                    if len(matches) == 28:
                        parts = []
                        last_idx = 0
                        for idx, m in enumerate(matches):
                            parts.append(doc_xml[last_idx:m.start()])
                            c_idx = idx // 4
                            var_slot = idx % 4

                            if c_idx in flir_results and var_slot < len(flir_results[c_idx]["vars"]):
                                new_var, new_val = flir_results[c_idx]["vars"][var_slot]
                            else:
                                new_var = m.group(2)
                                new_val = m.group(4)

                            parts.append(m.group(1) + new_var + m.group(3) + new_val + m.group(5))
                            last_idx = m.end()
                        parts.append(doc_xml[last_idx:])
                        doc_xml = "".join(parts)
                        total_fields_updated = sum(len(res["vars"]) for res in flir_results.values())
                        print(f"  [LINK] Synchronized {total_fields_updated} DOCVARIABLE fields to FLIR ActiveX controls")
                    data = doc_xml.encode("utf-8")

                elif fname == "word/settings.xml":
                    settings_xml = data.decode("utf-8")
                    new_docvars = []
                    for c_idx, res in flir_results.items():
                        for v_name, v_val in res["vars"]:
                            new_docvars.append(f'<w:docVar w:name="{v_name}" w:val="{v_val}"/>')
                    docvar_tags = "".join(new_docvars)

                    if "<w:docVars>" in settings_xml:
                        settings_xml = settings_xml.replace("<w:docVars>", f"<w:docVars>{docvar_tags}")
                    else:
                        settings_xml = settings_xml.replace("</w:settings>", f"<w:docVars>{docvar_tags}</w:docVars></w:settings>")
                    print(f"  [LINK] Registered {len(new_docvars)} dynamic variables in word/settings.xml")
                    data = settings_xml.encode("utf-8")

                elif fname == "_rels/.rels":
                    rels_xml_str = data.decode("utf-8")
                    if "custom-properties" not in rels_xml_str:
                        rels_xml_str = rels_xml_str.replace(
                            "</Relationships>",
                            '<Relationship Id="rIdFlirCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'
                        )
                    data = rels_xml_str.encode("utf-8")

                elif fname == "[Content_Types].xml":
                    ct_xml_str = data.decode("utf-8")
                    if "/docProps/custom.xml" not in ct_xml_str:
                        ct_xml_str = ct_xml_str.replace(
                            "</Types>",
                            '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'
                        )
                    data = ct_xml_str.encode("utf-8")

                zout.writestr(item_info, data)

            custom_props_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="FLIR"><vt:lpwstr>FLIR</vt:lpwstr></property>'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="FlirTempUnit"><vt:i4>0</vt:i4></property>'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="FlirDistUnit"><vt:i4>0</vt:i4></property>'
                '</Properties>'
            )
            zout.writestr("docProps/custom.xml", custom_props_xml.encode("utf-8"))

    safe_save_target(temp_final, target_docx)


# ==============================================================================
# PIPELINE EXECUTION
# ==============================================================================
def process_substation(sub_info: dict, selected_part_nums: Optional[List[int]], paths: dict):
    """Executes image injection for selected parts of a given substation."""
    sub_name = sub_info["name"]
    testsheet_path = sub_info["testsheet"]
    raw_ir_dir = sub_info["raw_ir"]

    print("\n" + "=" * 75)
    print(f" PROCESSING SUBSTATION: {sub_name}")
    print("=" * 75)

    if not testsheet_path or not testsheet_path.exists():
        print(f"ERROR: Testsheet (.xlsx) not found for {sub_name} in 04. TESTSHEET\\")
        return

    if not raw_ir_dir or not raw_ir_dir.exists():
        print(f"ERROR: Raw material IR directory not found for {sub_name} in 03. RAW MATERIAL\\")
        return

    print(f"  Testsheet: {testsheet_path.name}")
    print(f"  Raw IR   : {raw_ir_dir}")

    overviews_map, panels_map = parse_testsheet(testsheet_path)

    parts_dict = sub_info["parts"]
    if selected_part_nums is None:
        target_parts = sorted([p for p in parts_dict.keys() if parts_dict[p]["is_injectable"]])
    else:
        target_parts = sorted([p for p in selected_part_nums if p in parts_dict and parts_dict[p]["is_injectable"]])

    for part_num in target_parts:
        part_meta = parts_dict[part_num]
        target_docx = part_meta["path"]

        if part_num == 1:
            inject_part01_summary(target_docx, raw_ir_dir, overviews_map, paths)
        elif 2 <= part_num <= 9:
            panel_idx = part_num - 1
            if panel_idx in panels_map:
                inject_part_panel(target_docx, panel_idx, panels_map[panel_idx], raw_ir_dir, paths)
            else:
                print(f"WARNING: Panel {panel_idx} not found in testsheet data. Skipping {target_docx.name}")

    print("\n" + "=" * 75)
    print(f" COMPLETED INJECTION FOR: {sub_name}")
    print(f" [NOTE] Zero PDF conversion performed: Documents remain editable Word reports.")
    print("=" * 75)


# ==============================================================================
# MAIN & INTERACTIVE MENU
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Inject FLIR Radiometric Images and Visual Photos into CBM Word Reports.")
    parser.add_argument("--workspace", type=str, default=None, help="Root directory of the PO/workspace")
    parser.add_argument("--substation", type=str, default="all", help="Substation name or 'all'")
    parser.add_argument("--parts", type=str, default="all", help="Comma-separated part numbers (e.g. '1,2,3') or 'all'")
    parser.add_argument("--interactive", action="store_true", help="Force interactive console menu")

    args = parser.parse_args()

    paths = resolve_paths(Path(args.workspace) if args.workspace else None)
    ensure_fresh_backups(paths)

    substations = scan_available_reports(paths)
    if not substations:
        print(f"\n[ERROR] No CBM report Word documents (.docx) found in: {paths['full_report_dir']}")
        return 1

    # Check if we should run interactively
    is_interactive = args.interactive or (len(sys.argv) == 1 and sys.stdin.isatty())

    if is_interactive:
        sub_list = sorted(list(substations.keys()))
        while True:
            print("\n" + "-" * 75)
            print(" SCANNED SUBSTATIONS IN 01. FULL REPORT:")
            print("-" * 75)
            for idx, sub_name in enumerate(sub_list, 1):
                info = substations[sub_name]
                injectable_count = sum(1 for p in info["parts"].values() if p["is_injectable"])
                ts_status = "Testsheet OK" if info["testsheet"] else "Missing Testsheet"
                ir_status = "Raw IR OK" if info["raw_ir"] else "Missing Raw IR"
                print(f"  [{idx}] {sub_name}")
                print(f"      ({injectable_count} report parts | {ts_status} | {ir_status})")

            print("  " + "-" * 71)
            print("  [A] Process ALL Substations (All Parts)")
            print("  [Q] Quit / Exit")
            print("-" * 75)

            choice = input(f"\nSelect Substation to process [1-{len(sub_list)}, A, or Q]: ").strip()
            if choice.lower() == 'q':
                return 0
            if choice.lower() == 'a':
                for sub_name in sub_list:
                    process_substation(substations[sub_name], None, paths)
                print("\n[DONE] All substations processed successfully.")
                return 0

            if not choice.isdigit() or int(choice) < 1 or int(choice) > len(sub_list):
                print("Invalid selection.")
                continue

            sub_name = sub_list[int(choice) - 1]
            sub_info = substations[sub_name]
            parts_dict = sub_info["parts"]
            injectable_parts = sorted([p for p in parts_dict.keys() if parts_dict[p]["is_injectable"]])

            print("\n" + "=" * 75)
            print(f" SUBSTATION: {sub_name}")
            print("=" * 75)
            print("  [A] All Parts (Part 01 to Part 09)")
            for p_num in injectable_parts:
                print(f"  [{p_num:2d}] {parts_dict[p_num]['title']}")
            print("  [B] Back")
            print("  [Q] Quit")

            p_choice = input(f"\nSelect Part(s) [A, 1-{max(injectable_parts)}, comma-separated, or B]: ").strip()
            if p_choice.lower() == 'q': return 0
            if p_choice.lower() == 'b': continue
            if p_choice.lower() == 'a':
                process_substation(sub_info, None, paths)
                input("\nPress Enter to return to main menu...")
                continue

            sel_parts = [int(x.strip()) for x in p_choice.split(',') if x.strip().isdigit() and int(x.strip()) in injectable_parts]
            if sel_parts:
                process_substation(sub_info, sel_parts, paths)
                input("\nPress Enter to return to main menu...")
    else:
        # Non-interactive batch mode (invoked via GUI or CLI)
        target_sub = args.substation.strip()
        parts_str = args.parts.strip().lower()
        sel_parts = None if parts_str in ["all", "a", ""] else [int(x.strip()) for x in parts_str.split(',') if x.strip().isdigit()]

        if target_sub.lower() in ["all", "a", ""]:
            for s_name, s_info in substations.items():
                process_substation(s_info, sel_parts, paths)
        else:
            matched = False
            for s_name, s_info in substations.items():
                if target_sub.lower() in s_name.lower():
                    process_substation(s_info, sel_parts, paths)
                    matched = True
                    break
            if not matched:
                print(f"[ERROR] Substation matching '{target_sub}' not found.")
                return 1

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nProcess interrupted by user.")
        sys.exit(0)
