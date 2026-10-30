"""
WRAITH Auto-Rename Engine :: CBM Substation Report Image Matcher
Automatically renames raw thermal (IR) and visual (DC) inspection images
according to the substation template sequence defined in the CSV mapping file.
"""

import os
import re
import csv
import sys
import glob
from pathlib import Path
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def extract_substation_number(folder_or_file_name: str) -> str:
    """
    Extracts the substation identifier/number (e.g. '192', '193', '079', '113A', '113B')
    from folder name, CSV name, or path string.
    """
    base = os.path.basename(os.path.normpath(folder_or_file_name))
    # If the folder ends in /IR or /DG, look at the parent folder name
    if base.upper() in ["IR", "DG", "DIGITAL", "NOT DONE", "PROCESSED_EXCEL", "RAW_IMAGES"]:
        parent = os.path.dirname(os.path.normpath(folder_or_file_name))
        base = os.path.basename(parent)

    # Match patterns like "192. 31072026 MERLIN INN", "113A. 31072026...", "PE 192", "192"
    m = re.match(r'^(\d+[A-Za-z]?)', base.strip())
    if m:
        return m.group(1)

    m2 = re.search(r'(?:PE|SS|SUBSTATION|SUB|PO)?[_\s-]*(\d+[A-Za-z]?)', base, re.IGNORECASE)
    if m2:
        return m2.group(1)

    # Clean digits
    digits = "".join(c for c in base if c.isdigit())
    return digits if digits else base


def find_matching_csv(substation_num: str, mapping_dir: str) -> str:
    """
    Finds the exact or closest CSV file for the given substation number in mapping_dir.
    """
    if not os.path.exists(mapping_dir):
        return None

    csv_files = [f for f in os.listdir(mapping_dir) if f.lower().endswith(".csv")]
    if not csv_files:
        return None

    num_norm = substation_num.strip().lower().lstrip("0") or "0"

    # Exact filename match (e.g. "192.csv" or "0192.csv" or "192_something.csv")
    for f in csv_files:
        f_base = os.path.splitext(f)[0].lower()
        f_num = f_base.lstrip("0") or "0"
        if f_base == substation_num.lower() or f_num == num_norm:
            return os.path.join(mapping_dir, f)

    # Prefix match
    for f in csv_files:
        m = re.match(r'^(\d+[A-Za-z]?)', f)
        if m:
            m_num = m.group(1).lower().lstrip("0") or "0"
            if m_num == num_norm:
                return os.path.join(mapping_dir, f)

    # Substring match
    for f in csv_files:
        if substation_num.lower() in f.lower():
            return os.path.join(mapping_dir, f)

    return None


def parse_csv_required_images(csv_path: str) -> list:
    """
    Reads the CSV file and extracts the required image file slots in page sequence order.
    Returns a list of dicts: [
      {
        "page_no": 1,
        "page_name": "page1",
        "doc_template": "01. swg-overview.docx",
        "component": "OVERVIEW",
        "type": "Overview",
        "visual_img": "1.jpg",
        "ir_img": "2.jpg"
      }, ...
    ]
    """
    required_slots = []
    if not os.path.exists(csv_path):
        return required_slots

    try:
        with open(csv_path, mode="r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    except Exception:
        try:
            with open(csv_path, mode="r", encoding="latin-1", errors="replace") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
        except Exception:
            return required_slots

    # Group consecutive rows by Output File Name / Template Page No
    pages = {}
    page_order = []

    for r in rows:
        page_id = r.get("Output File Name") or r.get("Template Page No") or ""
        page_id = page_id.strip()
        if not page_id:
            continue

        if page_id not in pages:
            pages[page_id] = {
                "page_id": page_id,
                "page_no": r.get("Template Page No", ""),
                "doc_template": r.get("Word Template", ""),
                "type": r.get("Type", ""),
                "area": r.get("{{ swg.area }}") or r.get("swg.area") or r.get("Type", ""),
                "panel_name": r.get("{{ panel.name }}") or r.get("panel.name") or "",
                "visual_img": None,
                "ir_img": None
            }
            page_order.append(page_id)

        attr = (r.get("Attribute") or "").strip().lower()
        img_name = (r.get("Image File Name") or "").strip()

        if "visual" in attr or attr == "visual image path":
            pages[page_id]["visual_img"] = img_name
        elif "ir" in attr or attr == "ir image path":
            pages[page_id]["ir_img"] = img_name

    for pid in page_order:
        slot = pages[pid]
        # Only include slots that actually require both or either image
        if slot["visual_img"] or slot["ir_img"]:
            required_slots.append(slot)

    return required_slots


def discover_raw_image_pairs(target_dir: str) -> list:
    """
    Discovers all raw thermal (IR) and visible (DC) photos in target_dir.
    Handles FLIR format (IR_XXXX.jpg + DC_XXXX.jpg), consecutive camera captures,
    and timestamps.
    Returns a sorted list of pairs: [ {"ir": "IR_0005.jpg", "dc": "DC_0006.jpg", "index": 5}, ... ]
    """
    valid_exts = {".jpg", ".jpeg", ".png", ".bmp"}
    all_files = [f for f in os.listdir(target_dir) if os.path.splitext(f)[1].lower() in valid_exts]

    ir_files = {}
    dc_files = {}
    other_files = []

    for f in all_files:
        base_upper = f.upper()
        # Pattern 1: IR_XXXX.jpg and DC_XXXX.jpg
        m_ir = re.match(r'^IR[_-]?0*(\d+)', base_upper)
        m_dc = re.match(r'^DC[_-]?0*(\d+)', base_upper)
        m_flir = re.match(r'^FLIR[_-]?0*(\d+)', base_upper)

        if m_ir:
            num = int(m_ir.group(1))
            ir_files[num] = f
        elif m_dc:
            num = int(m_dc.group(1))
            dc_files[num] = f
        elif m_flir:
            num = int(m_flir.group(1))
            ir_files[num] = f
        else:
            other_files.append(f)

    pairs = []
    matched_dc = set()

    # Pair IR with DC by proximity (typically IR is N and DC is N+1, or same number N)
    for ir_num in sorted(ir_files.keys()):
        ir_file = ir_files[ir_num]
        dc_file = None

        # Check exact match or +1 / -1
        candidates = [ir_num + 1, ir_num, ir_num - 1, ir_num + 2]
        for c in candidates:
            if c in dc_files and c not in matched_dc:
                dc_file = dc_files[c]
                matched_dc.add(c)
                break

        pairs.append({
            "ir_name": ir_file,
            "dc_name": dc_file,
            "key": ir_num,
            "full_ir_path": os.path.join(target_dir, ir_file),
            "full_dc_path": os.path.join(target_dir, dc_file) if dc_file else None
        })

    # If some DC files were not matched with an IR, add them or pair with leftovers
    for dc_num, dc_file in sorted(dc_files.items()):
        if dc_num not in matched_dc:
            pairs.append({
                "ir_name": None,
                "dc_name": dc_file,
                "key": dc_num,
                "full_ir_path": None,
                "full_dc_path": os.path.join(target_dir, dc_file)
            })

    # If no standard IR/DC prefix found, sort remaining files chronologically
    if not pairs and other_files:
        other_files.sort(key=lambda x: os.path.getmtime(os.path.join(target_dir, x)))
        # Group consecutively in 2s (assuming pairs)
        for i in range(0, len(other_files), 2):
            f1 = other_files[i]
            f2 = other_files[i+1] if i+1 < len(other_files) else None
            pairs.append({
                "ir_name": f2,
                "dc_name": f1,
                "key": i,
                "full_ir_path": os.path.join(target_dir, f2) if f2 else None,
                "full_dc_path": os.path.join(target_dir, f1)
            })

    # Sort final pairs by sequence key
    pairs.sort(key=lambda x: x["key"] if x["key"] is not None else 999999)
    return pairs


def auto_rename_substation_folder(raw_folder_path: str, mapping_dir: str = None, dry_run: bool = False, log_fn=print) -> dict:
    """
    Auto-renames a single substation raw material directory based on its matching CSV.
    """
    raw_folder_path = os.path.abspath(raw_folder_path)
    if not os.path.exists(raw_folder_path):
        return {"success": False, "error": f"Folder not found: {raw_folder_path}"}

    # Determine target image directory (handle root vs /IR subfolder)
    ir_subfolder = os.path.join(raw_folder_path, "IR")
    target_img_dir = ir_subfolder if os.path.exists(ir_subfolder) and os.path.isdir(ir_subfolder) else raw_folder_path

    # Extract substation number
    sub_num = extract_substation_number(raw_folder_path)
    if not sub_num:
        return {"success": False, "error": f"Could not determine substation number from: {raw_folder_path}"}

    # Locate mapping dir if omitted
    if not mapping_dir or not os.path.exists(mapping_dir):
        # Traverse upwards to find "06. EXCELL MAPPING"
        curr = Path(raw_folder_path)
        for parent in [curr] + list(curr.parents):
            cand = parent / "06. EXCELL MAPPING"
            if cand.exists() and cand.is_dir():
                mapping_dir = str(cand)
                break

    if not mapping_dir or not os.path.exists(mapping_dir):
        return {"success": False, "error": f"06. EXCELL MAPPING folder not found for {raw_folder_path}"}

    csv_path = find_matching_csv(sub_num, mapping_dir)
    if not csv_path:
        return {"success": False, "error": f"No CSV found in {mapping_dir} matching substation #{sub_num}"}

    log_fn(f"📂 Substation #{sub_num}: Found CSV mapping -> {os.path.basename(csv_path)}")
    required_slots = parse_csv_required_images(csv_path)
    if not required_slots:
        return {"success": False, "error": f"CSV {os.path.basename(csv_path)} has no image slots defined."}

    raw_pairs = discover_raw_image_pairs(target_img_dir)
    if not raw_pairs:
        return {"success": False, "error": f"No image files found in: {target_img_dir}"}

    log_fn(f"  · Found {len(raw_pairs)} raw image pair(s) in {os.path.basename(target_img_dir)}")
    log_fn(f"  · Found {len(required_slots)} required report page slot(s) in CSV")

    # Build renaming operations with Vision Brain
    rename_plan = {}  # { current_filename: new_filename }
    staged_mapping = []

    try:
        from image_vision_brain import BRAIN
        has_brain = True
    except ImportError:
        has_brain = False

    if has_brain and BRAIN:
        staged_mapping = BRAIN.optimize_plan_with_vision(raw_pairs, required_slots, log_fn=log_fn)
        for item in staged_mapping:
            if item.get("raw_dc") and item.get("new_dc"):
                rename_plan[item["raw_dc"]] = item["new_dc"]
            if item.get("raw_ir") and item.get("new_ir"):
                rename_plan[item["raw_ir"]] = item["new_ir"]
            if item.get("match_reason"):
                log_fn(f"    ✓ {item['component']}: {item.get('raw_dc')} -> {item.get('new_dc')} ({item['match_reason']})")
    else:
        spare_counter = 1
        for idx, pair in enumerate(raw_pairs):
            if idx < len(required_slots):
                slot = required_slots[idx]
                target_dc = slot["visual_img"]
                target_ir = slot["ir_img"]
                comp_label = slot.get("panel_name") or slot.get("area") or slot.get("type") or f"Slot #{idx+1}"

                if pair["dc_name"] and target_dc:
                    rename_plan[pair["dc_name"]] = target_dc
                if pair["ir_name"] and target_ir:
                    rename_plan[pair["ir_name"]] = target_ir

                staged_mapping.append({
                    "type": "MAPPED",
                    "component": comp_label,
                    "raw_dc": pair["dc_name"],
                    "new_dc": target_dc,
                    "raw_ir": pair["ir_name"],
                    "new_ir": target_ir
                })
            else:
                spare_ir = f"A{spare_counter:04d}.jpg"
                spare_dc = f"A{spare_counter+1:04d}.jpg"
                spare_counter += 2

                if pair["ir_name"]:
                    rename_plan[pair["ir_name"]] = spare_ir
                if pair["dc_name"]:
                    rename_plan[pair["dc_name"]] = spare_dc

                staged_mapping.append({
                    "type": "SPARE",
                    "component": "Alternate / Spare Take",
                    "raw_dc": pair["dc_name"],
                    "new_dc": spare_dc,
                    "raw_ir": pair["ir_name"],
                    "new_ir": spare_ir
                })

            staged_mapping.append({
                "type": "SPARE",
                "component": "Alternate / Spare Take",
                "raw_dc": pair["dc_name"],
                "new_dc": spare_dc,
                "raw_ir": pair["ir_name"],
                "new_ir": spare_ir
            })

    if dry_run:
        return {
            "success": True,
            "substation": sub_num,
            "csv": os.path.basename(csv_path),
            "plan": rename_plan,
            "staged": staged_mapping
        }

    # Two-phase safe rename to prevent file collision
    temp_prefix = f"__tmp_cbm_{int(datetime.now().timestamp())}_"
    temp_map = {}

    try:
        # Phase 1: Rename all to unique temporary names
        for i, (orig, final) in enumerate(rename_plan.items()):
            src = os.path.join(target_img_dir, orig)
            if os.path.exists(src):
                tmp_name = f"{temp_prefix}{i:03d}_{orig}"
                tmp_path = os.path.join(target_img_dir, tmp_name)
                os.rename(src, tmp_path)
                temp_map[tmp_path] = final

        # Phase 2: Rename all temporary files to target final names
        for tmp_path, final in temp_map.items():
            final_path = os.path.join(target_img_dir, final)
            if os.path.exists(final_path):
                os.remove(final_path)
            os.rename(tmp_path, final_path)

        log_fn(f"✅ Successfully renamed {len(rename_plan)} file(s) for Substation #{sub_num}!")
        return {
            "success": True,
            "substation": sub_num,
            "csv": os.path.basename(csv_path),
            "count": len(rename_plan),
            "target_dir": target_img_dir,
            "staged": staged_mapping
        }

    except Exception as e:
        log_fn(f"❌ Error during renaming Substation #{sub_num}: {str(e)}")
        # Attempt recovery if temporary files exist
        for tmp_path, final in temp_map.items():
            if os.path.exists(tmp_path):
                orig_name = os.path.basename(tmp_path).replace(temp_prefix, "").split("_", 1)[-1]
                orig_path = os.path.join(target_img_dir, orig_name)
                try:
                    os.rename(tmp_path, orig_path)
                except Exception:
                    pass
        return {"success": False, "error": str(e)}


def auto_rename_batch_folders(folder_paths: list, mapping_dir: str = None, log_fn=print) -> dict:
    """
    Processes a list of raw substation folders.
    """
    results = []
    log_fn("=" * 60)
    log_fn(f"🏷️ AUTO RENAME PIPELINE STARTED ({len(folder_paths)} folder(s) selected)")
    log_fn("=" * 60)

    for p in folder_paths:
        res = auto_rename_substation_folder(p, mapping_dir=mapping_dir, log_fn=log_fn)
        results.append(res)

    success_count = sum(1 for r in results if r.get("success"))
    log_fn("=" * 60)
    log_fn(f"🎉 AUTO RENAME PIPELINE FINISHED: {success_count}/{len(folder_paths)} substations completed successfully.")
    log_fn("=" * 60)

    return {
        "total": len(folder_paths),
        "success_count": success_count,
        "results": results
    }


if __name__ == "__main__":
    if len(sys.argv) > 1:
        folders = [sys.argv[1]]
        mapping_arg = sys.argv[2] if len(sys.argv) > 2 else None
        auto_rename_batch_folders(folders, mapping_dir=mapping_arg)
    else:
        print("Usage: python auto_rename_images.py <folder_path> [mapping_dir]")
