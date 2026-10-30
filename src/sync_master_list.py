import os
import re
import sys
import time
import shutil
import warnings
import openpyxl
from copy import copy
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

# Suppress openpyxl's data validation user warning
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Fix Windows console encoding and enable unbuffered output
if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

# ── Paths ──────────────────────────────────────────────────────────────────────
DEFAULT_BASE = os.getcwd()
BASE = os.environ.get("BASE_DIR", DEFAULT_BASE)
CHECKLIST_PATH = os.path.join(BASE, "SUBSTATION_CHECKLIST.xlsx")
MASTER_LIST_PATH = os.path.join(BASE, "06. EXCELL MAPPING", "Master List.xlsx")

def create_timestamped_backup(file_path):
    if not os.path.exists(file_path):
        return
    try:
        backup_dir = os.path.join(os.path.dirname(file_path), "_backups")
        os.makedirs(backup_dir, exist_ok=True)
        base_name, ext = os.path.splitext(os.path.basename(file_path))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(backup_dir, f"{base_name}_{timestamp}{ext}")
        shutil.copy2(file_path, backup_file)
        print(f"  [BACKUP] Saved safety snapshot: {os.path.basename(backup_file)}", flush=True)
        backups = sorted([os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.startswith(base_name)], key=os.path.getmtime)
        while len(backups) > 10:
            os.remove(backups.pop(0))
    except Exception as e:
        print(f"  [BACKUP WARNING] Could not create backup: {e}", flush=True)

def is_blank(v):
    if v is None:
        return True
    s = str(v).strip()
    return s in ["", "-", "N/A", "n/a", "None", "none", "nan", "NAN", "null", "NULL", ".", "/"]

def normalize_compare(v):
    if is_blank(v):
        return ""
    s = str(v).strip().lower()
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return f"{f:.1f}"
    except (ValueError, TypeError):
        pass
    return s

def values_differ(v_existing, v_extracted):
    if is_blank(v_existing):
        return False
    if is_blank(v_extracted):
        return True
    return normalize_compare(v_existing) != normalize_compare(v_extracted)

def resolve_field(v_existing, v_new, field_name="", sub_num=""):
    """
    Sync incoming data from SUBSTATION_CHECKLIST.xlsx to Master List.xlsx:
    - Incoming non-blank data from SUBSTATION_CHECKLIST.xlsx is the source of truth and overrides Master List.xlsx.
    - If incoming value is blank/placeholder, preserve existing Master List.xlsx data if present.
    """
    # Guardrail: Cable types are not valid FP Types
    if field_name in ["FP1 Type", "FP2 Type"]:
        if str(v_new).strip().upper() in ["XLPE", "PILC", "PVC"]:
            v_new = ""
        if str(v_existing).strip().upper() in ["XLPE", "PILC", "PVC"]:
            v_existing = ""

    if not is_blank(v_new):
        return v_new

    if not is_blank(v_existing):
        return v_existing

    return v_existing if v_existing is not None else v_new

def determine_fp_lvdb_types(has_tx1, has_tx2, fp1_data, fp2_data):
    """
    Implements naming conventions for Feeder Pillar (FP) and LVDB based on Transformer (TX) setup:
    - 1 TX + 2 FP/LVDB: Use sequential numbering (FP 1/FP 2, LVDB 1/LVDB 2) or distinct types (FP/LVDB).
    - 2 TX + 1 FP/LVDB: Append TX association (FP TX1/FP TX2 or LVDB TX1/LVDB TX2).
    - 2 TX + 2 FP/LVDB: Match each FP/LVDB to its corresponding TX (FP TX1/FP TX2, LVDB TX1/LVDB TX2, or mixed LVDB TX1/FP TX2).
    """
    def _is_present(d):
        if not d:
            return False
        if isinstance(d, dict):
            mfr = d.get('manufacturer')
            sn = d.get('serial')
            rtg = d.get('rating')
            name = d.get('name')
            for v in [mfr, sn, rtg]:
                if v and str(v).strip() not in ['', '-', 'N/A', 'n/a', 'None', 'nan', '.', '/']:
                    return True
            if name and str(name).strip() not in ['', '-', 'N/A', 'n/a', 'None', 'nan', '.', '/', 'LVDB / FP 1', 'LVDB / FP 2']:
                return True
            return False
        s = str(d).strip()
        return s not in ['', '-', 'N/A', 'n/a', 'None', 'nan', '.', '/', 'LVDB / FP 1', 'LVDB / FP 2']

    def _get_kind(d):
        text = ''
        if isinstance(d, dict):
            parts = []
            for k in ['name', 'model', 'type', 'manufacturer']:
                val = d.get(k)
                if val and str(val).strip() not in ['LVDB / FP 1', 'LVDB / FP 2']:
                    parts.append(str(val))
            text = ' '.join(parts)
        else:
            text = str(d)
        text_upper = text.upper()
        if 'LVDB' in text_upper or 'DISTRIBUTION BOARD' in text_upper or 'LV BOARD' in text_upper or re.search(r'\bLV\s*DB\b', text_upper):
            return 'LVDB'
        return 'FP'

    has_fp1 = _is_present(fp1_data)
    has_fp2 = _is_present(fp2_data)

    if not has_fp1 and not has_fp2:
        return '', ''

    kind1 = _get_kind(fp1_data) if has_fp1 else ''
    kind2 = _get_kind(fp2_data) if has_fp2 else ''

    tx_count = (1 if has_tx1 else 0) + (1 if has_tx2 else 0)
    if tx_count == 0:
        tx_count = 1
        has_tx1 = True

    # 1. 1 TX + 2 FP/LVDB: Use sequential numbering (FP 1/FP 2, LVDB 1/LVDB 2) or distinct types (FP/LVDB).
    if tx_count == 1 and (has_fp1 and has_fp2):
        if kind1 == kind2:
            return f'{kind1} 1', f'{kind2} 2'
        else:
            return kind1, kind2

    # 2. 2 TX + 1 FP/LVDB: Append TX association (FP TX1/FP TX2 or LVDB TX1/LVDB TX2).
    if tx_count >= 2 and (has_fp1 ^ has_fp2):
        if has_fp1:
            raw_text = ' '.join([str(v) for v in fp1_data.values() if v]).upper() if isinstance(fp1_data, dict) else str(fp1_data).upper()
            assoc = 'TX2' if ('TX2' in raw_text or 'TX 2' in raw_text) else 'TX1'
            return f'{kind1} {assoc}', ''
        else:
            raw_text = ' '.join([str(v) for v in fp2_data.values() if v]).upper() if isinstance(fp2_data, dict) else str(fp2_data).upper()
            assoc = 'TX1' if ('TX1' in raw_text or 'TX 1' in raw_text) else 'TX2'
            return '', f'{kind2} {assoc}'

    # 3. 2 TX + 2 FP/LVDB: Match each FP/LVDB to its corresponding TX (FP TX1/FP TX2, LVDB TX1/LVDB TX2, or mixed LVDB TX1/FP TX2).
    if tx_count >= 2 and (has_fp1 and has_fp2):
        return f'{kind1} TX1', f'{kind2} TX2'

    # 4. 1 TX + 1 FP/LVDB: Single setup
    if tx_count == 1:
        if has_fp1:
            raw_text = ' '.join([str(v) for v in fp1_data.values() if v]).upper() if isinstance(fp1_data, dict) else str(fp1_data).upper()
            if '1' in raw_text and 'TX' not in raw_text:
                return f'{kind1} 1', ''
            elif 'TX' in raw_text:
                return f'{kind1} TX1', ''
            return kind1, ''
        else:
            raw_text = ' '.join([str(v) for v in fp2_data.values() if v]).upper() if isinstance(fp2_data, dict) else str(fp2_data).upper()
            if '2' in raw_text and 'TX' not in raw_text:
                return '', f'{kind2} 2'
            elif 'TX' in raw_text:
                return '', f'{kind2} TX2'
            return '', kind2

    t1 = f'{kind1} 1' if has_fp1 else ''
    t2 = f'{kind2} 2' if has_fp2 else ''
    return t1, t2

def parse_substation_filter(input_str):
    if not input_str:
        return None, "ALL"
    s = str(input_str).strip()
    if s.upper() in ["ALL", "*", ""]:
        return None, "ALL"

    parts = [p.strip() for p in s.split(",") if p.strip()]
    ranges = []
    singles = set()

    for part in parts:
        part_upper = part.upper()
        if part_upper in ["ALL", "*"]:
            return None, "ALL"
        m_range = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m_range:
            low, high = int(m_range.group(1)), int(m_range.group(2))
            if low > high:
                low, high = high, low
            ranges.append((low, high))
            continue
        m_plus = re.match(r"^(?:>=|>)?\s*(\d+)\+?$", part)
        if "+" in part or ">=" in part:
            low = int(m_plus.group(1))
            ranges.append((low, 999999))
            continue
        m_single = re.match(r"^(?:PE\s*)?(\d+)([A-Z]?)$", part, re.IGNORECASE)
        if m_single:
            singles.add(int(m_single.group(1)))
            continue

    if not ranges and not singles:
        return None, "ALL"

    def matches(num):
        m = re.match(r"^\d+", str(num).strip())
        if not m:
            return False
        n = int(m.group(0))
        if n in singles:
            return True
        for low, high in ranges:
            if low <= n <= high:
                return True
        return False

    label_parts = []
    if singles:
        if len(singles) == 1:
            label_parts.append(f"PE {list(singles)[0]} only")
        else:
            label_parts.append(f"PE {', '.join(map(str, sorted(singles)))}")
    for low, high in ranges:
        if high == 999999:
            label_parts.append(f"PE {low} onwards")
        else:
            label_parts.append(f"PE {low}-{high}")

    return matches, " & ".join(label_parts)

def clean_val(val):
    if val is None:
        return ""
    val = str(val).strip()
    if val in ['-', 'N/A', '', '.', '/'] or val.lower() == 'n/a':
        return ""
    return val

def format_temp(val):
    if val is None:
        return ""
    s = str(val).strip()
    if not s or s in ['-', 'N/A', '.', '/']:
        return ""
    try:
        f = float(s)
        return f"{f:.1f}"
    except ValueError:
        return s

def format_heater_amp(val):
    if val is None:
        return ""
    s = str(val).strip()
    if not s or s in ['-', 'N/A', '', '.', '/'] or s.lower() in ['n/a', 'nan', 'none', 'null']:
        return ""
    m = re.search(r'[-+]?\d*\.?\d+', s)
    if m:
        try:
            f = float(m.group(0))
            return f"{f:.2f}A"
        except ValueError:
            pass
    return s

def natural_sort_key(key):
    s = str(key).strip()
    m = re.match(r'^(\d+)(.*)$', s)
    if m:
        return (int(m.group(1)), m.group(2))
    return (float('inf'), s)


def sync_and_merge_to_master(target_scope=None):
    start_time = time.time()
    scope_str = target_scope or (sys.argv[1] if len(sys.argv) > 1 else "ALL")
    matches_scope, scope_label = parse_substation_filter(scope_str)

    print("=" * 60)
    print(f"  COMMAND 2: SYNC & MERGE TO MASTER LIST | SCOPE: {scope_label.upper()}")
    print("=" * 60)
    print(f"  Workspace: {BASE}")

    if not os.path.exists(CHECKLIST_PATH):
        print(f"  [ERROR] Checklist file not found: {CHECKLIST_PATH}")
        return False

    if not os.path.exists(MASTER_LIST_PATH):
        print(f"  [ERROR] Master List file not found: {MASTER_LIST_PATH}")
        return False

    print("  Loading SUBSTATION_CHECKLIST.xlsx...")
    wb_cl = openpyxl.load_workbook(CHECKLIST_PATH)

    if 'Substation Checklist' not in wb_cl.sheetnames:
        print("  [ERROR] 'Substation Checklist' sheet not found in checklist workbook.")
        return False

    ws_sc = wb_cl['Substation Checklist']
    has_backup = 'Backup' in wb_cl.sheetnames
    ws_bk = wb_cl['Backup'] if has_backup else None

    # Read Headers from Row 4
    cols = [ws_sc.cell(4, c).value for c in range(1, ws_sc.max_column + 1)]
    print(f"  Found {len(cols)} columns in Checklist.")

    # 1. Read existing rows from Substation Checklist
    def read_sheet_rows(ws):
        data = {}
        ordered_keys = []
        for r in range(5, ws.max_row + 1):
            num = ws.cell(r, 1).value
            name = ws.cell(r, 2).value
            if num is None and name is None:
                continue
            key = str(num).strip() if num is not None else f"UNKNOWN_{r}"
            row_dict = {}
            for c_idx, col_name in enumerate(cols, start=1):
                if col_name:
                    row_dict[col_name] = ws.cell(r, c_idx).value
            data[key] = row_dict
            ordered_keys.append(key)
        return data, ordered_keys

    sc_data, sc_keys = read_sheet_rows(ws_sc)
    bk_data, bk_keys = read_sheet_rows(ws_bk) if has_backup else ({}, [])

    print(f"  Loaded {len(sc_data)} substations from 'Substation Checklist'.")
    if has_backup:
        print(f"  Loaded {len(bk_data)} substations from 'Backup'.")

    # 2. Smart Merge: Only process live substations discovered from Substation Checklist
    merged_data = {}
    final_keys = [k for k in sc_keys]
    final_keys.sort(key=natural_sort_key)



    merge_count = 0
    for k in final_keys:
        d_sc = sc_data.get(k, {})
        d_bk = bk_data.get(k, {})
        
        LIVE_SCAN_COLS = {'FULL REPORT', 'QUICK REPORT', 'RAW MATERIAL', 'TEST SHEET'}

        merged_row = {}
        for col in cols:
            if not col:
                continue
            v_sc = d_sc.get(col)
            v_bk = d_bk.get(col) if has_backup else None
            
            if col in LIVE_SCAN_COLS:
                # Document status must ALWAYS reflect the live folder scan
                final_v = v_sc if v_sc is not None and str(v_sc).strip() != "" else (v_bk or "NO")
            elif col == 'SUBSTATION NAME':
                # Prefer live scanned name from RAW MATERIAL over old guessed names
                if v_sc and not str(v_sc).lower().endswith(('.docx', '.pdf', '.doc', '.xlsx')):
                    final_v = str(v_sc).strip()
                elif v_bk is not None and str(v_bk).strip() != "":
                    final_v = re.sub(r'\.(docx|doc|pdf|xlsx)$', '', str(v_bk).strip(), flags=re.IGNORECASE).strip()
            elif v_sc is not None and str(v_sc).strip() != "":
                final_v = v_sc
            elif v_bk is not None and str(v_bk).strip() != "":
                final_v = v_bk
            else:
                final_v = None
                
            # Rule 4: Ambient Temperature 1 decimal precision
            if col == 'Ambient Temp' and final_v is not None:
                final_v = format_temp(final_v)

            merged_row[col] = final_v


        # Single switchgear guardrail: Check if Switchgear Manufacturer is year e.g. 2024
        if str(merged_row.get('Switchgear Manufacturer', '')).strip() in ['2024', '2024.0', '2025', '2025.0', '2026', '2026.0']:
            merged_row['Switchgear Manufacturer'] = None

        # FP / LVDB Type guardrail: clean cable types ('XLPE'/'PILC') and enforce conventions
        for fp_col in ['FP1 Type', 'FP2 Type']:
            fp_val = str(merged_row.get(fp_col, '') or '').strip()
            if fp_val.upper() in ['XLPE', 'PILC', 'PVC']:
                merged_row[fp_col] = None

        # User rule: FP1 & FP2 model always left blank
        merged_row['FP1 Model'] = None
        merged_row['FP2 Model'] = None

        has_tx1_m = bool(merged_row.get('TX1 Manufacturer') or merged_row.get('TX1 Serial Number') or merged_row.get('TX1 Rating'))
        has_tx2_m = bool(merged_row.get('TX2 Manufacturer') or merged_row.get('TX2 Serial Number') or merged_row.get('TX2 Rating'))
        fp1_info = {
            'manufacturer': merged_row.get('FP1 Manufacturer'), 'serial': merged_row.get('FP1 Serial No'),
            'rating': merged_row.get('FP1 Rating'), 'model': merged_row.get('FP1 Model'),
            'type': merged_row.get('FP1 Type')
        }
        fp2_info = {
            'manufacturer': merged_row.get('FP2 Manufacturer'), 'serial': merged_row.get('FP2 Serial No'),
            'rating': merged_row.get('FP2 Rating'), 'model': merged_row.get('FP2 Model'),
            'type': merged_row.get('FP2 Type')
        }
        c1, c2 = determine_fp_lvdb_types(has_tx1_m, has_tx2_m, fp1_info, fp2_info)
        if not merged_row.get('FP1 Type') and c1:
            merged_row['FP1 Type'] = c1
        if not merged_row.get('FP2 Type') and c2:
            merged_row['FP2 Type'] = c2

        merged_data[k] = merged_row
        merge_count += 1

    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
    HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    YES_FONT = Font(name="Calibri", bold=True, color="1E8449", size=11)
    NO_FONT = Font(name="Calibri", bold=True, color="C0392B", size=11)
    YES_FILL = PatternFill("solid", fgColor="D5F5E3")
    NO_FILL = PatternFill("solid", fgColor="FADBD8")
    ALT_FILL = PatternFill("solid", fgColor="D6E4F0")
    PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")
    INFO_FONT = Font(name="Calibri", size=10, color="1A5276")
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B0B0B0")
    thick = Side(style="medium", color="1F4E79")
    BORDER_INNER = Border(left=thin, right=thin, top=thin, bottom=thin)
    BORDER_HEAD = Border(left=thick, right=thick, top=thick, bottom=thick)

    # 3. Synchronize both sheets in SUBSTATION_CHECKLIST.xlsx
    def write_sheet_data(ws, data_dict, key_list):
        target_max_row = len(key_list) + 4
        # If worksheet has more rows than required, delete excess rows
        if ws.max_row > target_max_row:
            ws.delete_rows(target_max_row + 1, ws.max_row - target_max_row)

        # Ensure header styling on row 4
        for c_idx in range(1, len(cols) + 1):
            c_hdr = ws.cell(4, c_idx)
            c_hdr.font = HEADER_FONT
            c_hdr.fill = HEADER_FILL
            c_hdr.alignment = CENTER
            c_hdr.border = BORDER_HEAD

        LIVE_SCAN_COLS = {'FULL REPORT', 'QUICK REPORT', 'RAW MATERIAL', 'TEST SHEET'}
        INFO_COLS = {'Time', 'Date', 'Ambient Temp', 'Humidity'}
        LEFT_COLS = {
            'SUBSTATION NAME', 'Switchgear Manufacturer', 'Switchgear Model',
            'TX1 Manufacturer', 'TX1 Model', 'TX1 Serial Number',
            'TX2 Manufacturer', 'TX2 Model', 'TX2 Serial Number',
            'FP1 Manufacturer', 'FP1 Model', 'FP1 Serial No',
            'FP2 Manufacturer', 'FP2 Model', 'FP2 Serial No',
            'Battery Bank 1 Manufactrurer', 'Battery Bank 1 Model', 'Battery Bank 1 Serial Number',
            'Battery Bank 2 Manufactrurer', 'Battery Bank 2 Model', 'Battery Bank 2 Serial Number',
            'Battery Bank 3 Manufactrurer', 'Battery Bank 3 Model', 'Battery Bank 3 Serial Number',
            'Battery Bank 4 Manufactrurer', 'Battery Bank 4 Model', 'Battery Bank 4 Serial Number'
        }

        for idx, k in enumerate(key_list, start=5):
            r = idx
            ws.row_dimensions[r].height = 20
            row_fill = ALT_FILL if (idx - 5) % 2 == 0 else PLAIN_FILL
            r_data = data_dict[k]
            num_val = r_data.get('NUMBER', k)
            if str(num_val).isdigit():
                num_val = int(num_val)

            c1 = ws.cell(r, 1, value=num_val)
            c1.font = Font(name="Calibri", size=10, bold=True)
            c1.alignment = CENTER
            c1.fill = row_fill
            c1.border = BORDER_INNER

            for c_idx, col_name in enumerate(cols[1:], start=2):
                if not col_name:
                    continue
                val = r_data.get(col_name)
                c = ws.cell(r, c_idx, value=val)
                c.border = BORDER_INNER

                if col_name in LIVE_SCAN_COLS:
                    val_str = str(val).strip().upper() if val is not None else "NO"
                    c.font = YES_FONT if val_str == "YES" else NO_FONT
                    c.fill = YES_FILL if val_str == "YES" else NO_FILL
                    c.alignment = CENTER
                elif col_name in INFO_COLS:
                    c.font = INFO_FONT
                    c.fill = row_fill
                    c.alignment = CENTER
                elif col_name in LEFT_COLS:
                    c.font = Font(name="Calibri", size=10)
                    c.fill = row_fill
                    c.alignment = LEFT
                else:
                    c.font = Font(name="Calibri", size=10)
                    c.fill = row_fill
                    c.alignment = CENTER

    write_sheet_data(ws_sc, merged_data, final_keys)
    if has_backup:
        write_sheet_data(ws_bk, merged_data, final_keys)
    else:
        ws_bk = wb_cl.copy_worksheet(ws_sc)
        ws_bk.title = "Backup"

    # Synchronize 'Substation Info' sheet
    if 'Substation Info' in wb_cl.sheetnames:
        ws_info = wb_cl['Substation Info']
        info_cols = [ws_info.cell(3, c).value for c in range(1, ws_info.max_column + 1)]
        target_info_max = len(final_keys) + 3
        if ws_info.max_row > target_info_max:
            ws_info.delete_rows(target_info_max + 1, ws_info.max_row - target_info_max)

        existing_info = {}
        for r in range(4, ws_info.max_row + 1):
            num_v = ws_info.cell(r, 1).value
            if num_v is not None:
                k_str = str(num_v).strip()
                existing_info[k_str] = {
                    info_cols[c - 1]: ws_info.cell(r, c).value
                    for c in range(1, len(info_cols) + 1) if info_cols[c - 1]
                }

        for idx, k in enumerate(final_keys, start=4):
            r = idx
            ws_info.row_dimensions[r].height = 20
            row_fill = ALT_FILL if (idx - 4) % 2 == 0 else PLAIN_FILL
            r_data = merged_data[k]
            i_data = existing_info.get(k, {})

            num_val = r_data.get('NUMBER', k)
            if str(num_val).isdigit():
                num_val = int(num_val)

            sub_name = r_data.get('SUBSTATION NAME', '')
            sub_type = r_data.get('SUBSTATION TYPE', '')
            dt_val = i_data.get('DATE') or r_data.get('Date', '')
            gps_val = i_data.get('GPS COORDINATES', '')
            erms_name = i_data.get('SUBSTATION NAME (ERMS)') or sub_name
            site_name = i_data.get('SUBSTATION NAME (SITE)') or sub_name
            fl_erms = i_data.get('F/L NO (ERMS)') or "N/A"
            fl_site = i_data.get('F/L NO (SITE)') or "N/A"

            row_vals = [num_val, sub_name, sub_type, dt_val, gps_val, erms_name, site_name, fl_erms, fl_site]
            for c_idx, val in enumerate(row_vals, start=1):
                c = ws_info.cell(r, c_idx, value=val)
                c.font = Font(name="Calibri", size=10)
                if c_idx == 1:
                    c.font = Font(name="Calibri", size=10, bold=True)
                    c.alignment = CENTER
                elif c_idx in (3, 4):
                    c.alignment = CENTER
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c.fill = row_fill
                c.border = BORDER_INNER

    # 4. Read Panel Schedule from Checklist Workbook
    panel_schedules_by_sub = {}
    if 'Panel Schedule' in wb_cl.sheetnames:
        ws_ps = wb_cl['Panel Schedule']
        for r in range(2, ws_ps.max_row + 1):
            sub_no = ws_ps.cell(r, 1).value
            if sub_no is None:
                continue
            sub_key = str(sub_no).strip()
            if sub_key.isdigit():
                norm_sub_key = f"{int(sub_key):03d}"
                raw_int_key = str(int(sub_key))
            else:
                norm_sub_key = sub_key
                raw_int_key = sub_key

            p_link = clean_val(ws_ps.cell(r, 4).value)
            p_name = clean_val(ws_ps.cell(r, 5).value)
            p_cable = clean_val(ws_ps.cell(r, 6).value)
            p_status = clean_val(ws_ps.cell(r, 7).value)
            p_busbar = clean_val(ws_ps.cell(r, 8).value)
            p_load = clean_val(ws_ps.cell(r, 9).value)
            p_heater = format_heater_amp(ws_ps.cell(r, 10).value)
            p_sn = clean_val(ws_ps.cell(r, 11).value)

            p_name_upper = p_name.upper().strip()

            # Check switchgear type for this substation
            sub_sg_type = str(merged_data.get(sub_key, {}).get('Switchgear Type', '') or '').upper()
            is_vcb = "VCB" in sub_sg_type

            is_transition = (
                p_name_upper in ["TRANSITION", "TRANSITION PANEL", "TRANSITION-PANEL", "TRANSITION_PANEL"]
                or bool(re.match(r'^TRANSITION(\s+PANEL)?$', p_name_upper))
                or "TRANSITION" in p_name_upper
            )
            is_bus_section = (
                p_name_upper in ["BUS SECTION", "BUS-SECTION", "BUSSECTION", "BUS SECTION PANEL"]
                or bool(re.match(r'^BUS\s+SECTION(\s+PANEL)?$', p_name_upper))
                or "BUS SECTION" in p_name_upper
            )

            if is_vcb and is_transition:
                p_cable = "-"
                p_link = "-"
            elif is_bus_section:
                p_cable = "-"

            p_vals = [
                p_link if p_link != "" else None,
                p_name if p_name != "" else None,
                p_cable if p_cable != "" else None,
                p_status if p_status != "" else None,
                p_busbar if p_busbar != "" else "-",
                p_load if p_load != "" else None,
                p_heater if p_heater != "" else None,
                p_sn if p_sn != "" else None
            ]
            
            panel_schedules_by_sub.setdefault(norm_sub_key, []).append(p_vals)
            if raw_int_key != norm_sub_key:
                panel_schedules_by_sub.setdefault(raw_int_key, []).append(p_vals)

    create_timestamped_backup(CHECKLIST_PATH)
    try:
        wb_cl.save(CHECKLIST_PATH)
        print("  [OK] SUBSTATION_CHECKLIST.xlsx synced (Checklist & Backup 100% matched)!")
    except PermissionError:
        print(f"  [CRITICAL ERROR] Cannot save: {os.path.basename(CHECKLIST_PATH)} is currently open in Excel!", flush=True)
        print(f"  --> Please close {os.path.basename(CHECKLIST_PATH)} and run sync again.", flush=True)
        return False

    # 5. Update Master List.xlsx
    print("  Updating 06. EXCELL MAPPING/Master List.xlsx...")
    master_wb = openpyxl.load_workbook(MASTER_LIST_PATH)

    if 'CSV TEMPLATE' not in master_wb.sheetnames:
        print("  [ERROR] 'CSV TEMPLATE' sheet not found in Master List.xlsx.")
        return False

    template_ws = master_wb['CSV TEMPLATE']

    # 5a. Update SUBSTATION LIST Tab
    if 'SUBSTATION LIST' in master_wb.sheetnames:
        ws_sub_list = master_wb['SUBSTATION LIST']
        target_max_row = len(final_keys) + 1
        if matches_scope is None and ws_sub_list.max_row > target_max_row:
            ws_sub_list.delete_rows(target_max_row + 1, ws_sub_list.max_row - target_max_row)

        for idx, k in enumerate(final_keys, start=2):
            num_str = str(k).strip()
            m_k = re.match(r'^\d+', num_str)
            base_k = int(m_k.group(0)) if m_k else num_str
            is_sub_in_scope = True if matches_scope is None else matches_scope(base_k)
            if not is_sub_in_scope:
                continue

            r_data = merged_data[k]
            num_val = r_data.get('NUMBER', k)
            if str(num_val).isdigit():
                num_val = int(num_val)

            def _clean_spec(v):
                if v is None:
                    return "-"
                s = str(v).strip()
                if s in ["", "-", "N/A", ".", "/", "0", "0.0", "nan", "None", "null"] or s.lower() in ["nan", "none", "null", "n/a"]:
                    return "-"
                if s.upper() in ["XLPE", "PILC", "PVC"]:
                    return "-"
                return s

            col_values_40 = [
                num_val,
                _clean_spec(r_data.get('SUBSTATION NAME')),
                _clean_spec(r_data.get('Time')),
                _clean_spec(r_data.get('Date')),
                _clean_spec(r_data.get('Ambient Temp')),
                _clean_spec(r_data.get('Humidity')),
                _clean_spec(r_data.get('Switchgear Manufacturer')),
                _clean_spec(r_data.get('Switchgear Type')),
                _clean_spec(r_data.get('Switchgear Rating')),
                _clean_spec(r_data.get('Switchgear Model')),
                _clean_spec(r_data.get('TX1 Manufacturer')),
                _clean_spec(r_data.get('TX1 Model')),
                _clean_spec(r_data.get('TX1 Rating')),
                _clean_spec(r_data.get('TX1 Serial Number')),
                _clean_spec(r_data.get('TX2 Manufacturer')),
                _clean_spec(r_data.get('TX2 Model')),
                _clean_spec(r_data.get('TX2 Rating')),
                _clean_spec(r_data.get('TX2 Serial Number')),
                _clean_spec(r_data.get('FP1 Manufacturer')),
                _clean_spec(r_data.get('FP1 Model')),
                _clean_spec(r_data.get('FP1 Rating')),
                _clean_spec(r_data.get('FP1 Serial No')),
                _clean_spec(r_data.get('FP1 Type')),
                _clean_spec(r_data.get('FP2 Manufacturer')),
                _clean_spec(r_data.get('FP2 Model')),
                _clean_spec(r_data.get('FP2 Rating')),
                _clean_spec(r_data.get('FP2 Serial No')),
                _clean_spec(r_data.get('FP2 Type')),
                _clean_spec(r_data.get('Battery Bank 1 Manufactrurer')),
                _clean_spec(r_data.get('Battery Bank 1 Model')),
                _clean_spec(r_data.get('Battery Bank 1 Serial Number')),
                _clean_spec(r_data.get('Battery Bank 2 Manufactrurer')),
                _clean_spec(r_data.get('Battery Bank 2 Model')),
                _clean_spec(r_data.get('Battery Bank 2 Serial Number')),
                _clean_spec(r_data.get('Battery Bank 3 Manufactrurer')),
                _clean_spec(r_data.get('Battery Bank 3 Model')),
                _clean_spec(r_data.get('Battery Bank 3 Serial Number')),
                _clean_spec(r_data.get('Battery Bank 4 Manufactrurer')),
                _clean_spec(r_data.get('Battery Bank 4 Model')),
                _clean_spec(r_data.get('Battery Bank 4 Serial Number'))
            ]

            for c_i, val in enumerate(col_values_40, start=1):
                if c_i == 1:
                    ws_sub_list.cell(row=idx, column=c_i, value=val)
                else:
                    existing_val = ws_sub_list.cell(row=idx, column=c_i).value
                    f_name = "FP1 Type" if c_i == 23 else ("FP2 Type" if c_i == 28 else "")
                    final_val = resolve_field(existing_val, val, field_name=f_name)
                    ws_sub_list.cell(row=idx, column=c_i, value=final_val)

        print("  [OK] 'SUBSTATION LIST' sheet in Master List.xlsx updated with clean specifications!")

    # Prune non-existent / ghost substation sheets from Master List.xlsx
    valid_sub_sheets = set()
    for k in final_keys:
        num_str = str(k).strip()
        if num_str.isdigit():
            valid_sub_sheets.add(f"{int(num_str):03d}")
            valid_sub_sheets.add(str(int(num_str)))
        valid_sub_sheets.add(num_str)

    SYSTEM_SHEETS = {
        'Image_Rename', 'Sheet1', 'Image_Rename_VCB', 'Image_Rename_RMU',
        'Image_Mapping_VCB (2)', 'Image_Mapping_VCB', 'Image_Mapping_RMU', 'Image_Mapping_H-POLE',
        'Master_List', 'Sheet2', 'SUBSTATION LIST', 'CSV TEMPLATE', 'Raw Start Number'
    }

    for s_name in list(master_wb.sheetnames):
        if s_name not in SYSTEM_SHEETS and s_name not in valid_sub_sheets:
            del master_wb[s_name]
            print(f"  [CLEANUP] Removed non-existent substation sheet '{s_name}' from Master List.xlsx")

    # 5c. Update individual substation sheets with Panel Schedules (Fast in-place update)
    print("  Updating individual substation tabs with Panel Schedules...", flush=True)
    total_keys = len(final_keys)
    step_s_interval = 10 if total_keys <= 50 else (25 if total_keys <= 150 else 50)

    for idx_s, k in enumerate(final_keys, start=1):
        num_str = str(k).strip()
        m_k = re.match(r'^\d+', num_str)
        base_k = int(m_k.group(0)) if m_k else num_str
        is_sub_in_scope = True if matches_scope is None else matches_scope(base_k)

        if num_str.isdigit():
            sheet_name = f"{int(num_str):03d}"
        else:
            sheet_name = num_str

        # If outside scope and sheet already exists in Master List.xlsx, DO NOT TOUCH IT!
        if not is_sub_in_scope and (sheet_name in master_wb.sheetnames or num_str in master_wb.sheetnames):
            continue

        panels = panel_schedules_by_sub.get(sheet_name, panel_schedules_by_sub.get(num_str, []))

        if sheet_name in master_wb.sheetnames:
            target_ws = master_wb[sheet_name]
        elif num_str in master_wb.sheetnames:
            target_ws = master_wb[num_str]
        else:
            target_ws = None

        if target_ws:
            # Fast in-place cell update preserving existing edits if data is not same
            start_row = 2
            for p_idx, p_vals in enumerate(panels):
                r = start_row + p_idx
                for offset_idx, val in enumerate(p_vals):
                    target_col = 2 + offset_idx
                    existing_val = target_ws.cell(row=r, column=target_col).value
                    final_val = resolve_field(existing_val, val)
                    target_ws.cell(row=r, column=target_col, value=final_val)
        else:
            # Create fresh sheet from CSV TEMPLATE if it's a new substation
            new_ws = master_wb.create_sheet(title=sheet_name)

            for row in template_ws.iter_rows(values_only=False):
                for cell in row:
                    new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = cell.number_format
                        new_cell.alignment = copy(cell.alignment)

            for col_letter, dim in template_ws.column_dimensions.items():
                new_ws.column_dimensions[col_letter].width = dim.width

            for row_num, dim in template_ws.row_dimensions.items():
                new_ws.row_dimensions[row_num].height = dim.height

            for merged_range in template_ws.merged_cells.ranges:
                new_ws.merge_cells(str(merged_range))

            if template_ws.tables:
                for tbl_name, tbl_ref in template_ws.tables.items():
                    new_tbl = openpyxl.worksheet.table.Table(displayName=f"Table_{sheet_name}_{tbl_name}", ref=tbl_ref)
                    new_ws.add_table(new_tbl)

            start_row = 2
            for p_vals in panels:
                for offset_idx, val in enumerate(p_vals):
                    new_ws.cell(row=start_row, column=2 + offset_idx, value=val)
                start_row += 1

        if idx_s % step_s_interval == 0 or idx_s == total_keys or idx_s == 1:
            print(f"    Synced {idx_s:03d}/{total_keys:03d} substation tabs (Sheet '{sheet_name}')...", flush=True)

    # 5d. Reorder sheets in Master List so substation tabs are sequentially ordered
    system_sheet_names = [
        'Image_Rename', 'Sheet1', 'Image_Rename_VCB', 'Image_Rename_RMU',
        'Image_Mapping_VCB (2)', 'Image_Mapping_VCB', 'Image_Mapping_RMU',
        'Image_Mapping_H-POLE', 'Master_List', 'Sheet2', 'SUBSTATION LIST',
        'CSV TEMPLATE', 'Raw Start Number'
    ]

    def sheet_sort_key(ws):
        title = ws.title.strip()
        if title in system_sheet_names:
            return (0, system_sheet_names.index(title), "")
        m = re.match(r'^(\d+)(.*)$', title)
        if m:
            return (1, int(m.group(1)), m.group(2))
        return (2, 0, title)

    print("  Reordering all sheets in natural numerical sequence...", flush=True)
    master_wb._sheets.sort(key=sheet_sort_key)

    print("  Saving Master List.xlsx (this may take a few seconds)...", flush=True)
    create_timestamped_backup(MASTER_LIST_PATH)
    try:
        master_wb.save(MASTER_LIST_PATH)
    except PermissionError:
        print(f"  [CRITICAL ERROR] Cannot save: {os.path.basename(MASTER_LIST_PATH)} is currently open in Excel!", flush=True)
        print(f"  --> Please close {os.path.basename(MASTER_LIST_PATH)} and run sync again.", flush=True)
        return False

    elapsed = time.time() - start_time
    print(f"  [OK] Master List.xlsx updated successfully with individual sheets & panel schedules!")
    print(f"  [TIME] COMPLETED IN {elapsed:.2f} SECONDS", flush=True)
    print("=" * 60, flush=True)
    return True

if __name__ == "__main__":
    scope_arg = sys.argv[1] if len(sys.argv) > 1 else "ALL"
    sync_and_merge_to_master(scope_arg)
