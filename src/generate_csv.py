import os
import sys
import re
import pandas as pd
import win32com.client

# ── Dynamic Path Setup ─────────────────────────────────────────────────────────
DEFAULT_BASE = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42234207 - JOHOR - JBU - 400 PE IR US TEV"
BASE_DIR = os.environ.get("BASE_DIR", DEFAULT_BASE)

MAPPING_DIR = os.environ.get("MAPPING_DIR", os.path.join(BASE_DIR, "06. EXCELL MAPPING"))
RM_DIR = os.environ.get("RM_DIR", os.path.join(BASE_DIR, "03. RAW MATERIAL"))
EXCEL_PATH = os.path.join(MAPPING_DIR, "Master List.xlsx")


def resolve_rm_subfolder(rm_dir, input_val):
    if not rm_dir or not os.path.exists(rm_dir):
        return str(input_val)

    val_str = str(input_val).strip()
    m = re.match(r'^0*(\d+)([a-zA-Z]*)', val_str)
    if m:
        clean_num = m.group(1) + m.group(2).upper()
        base_digits = m.group(1)
        suffix = m.group(2).upper()
    else:
        clean_num = val_str
        base_digits = val_str
        suffix = ''

    try:
        dir_entries = [d for d in os.listdir(rm_dir) if os.path.isdir(os.path.join(rm_dir, d))]
    except Exception:
        return str(input_val)

    # 1. Exact match
    for d in dir_entries:
        if d.lower() == val_str.lower() or d.lower() == clean_num.lower():
            return d

    # 2. Strict prefix match with suffix e.g. 113A. or 113A
    if suffix:
        for d in dir_entries:
            m_d = re.match(r'^0*(\d+[a-zA-Z]?)[\.\s\-_]', d)
            if m_d and m_d.group(1).upper() == clean_num:
                return d

    # 3. Numeric prefix match (e.g. "48. ..." matches "048" and "48")
    for d in dir_entries:
        m_d = re.match(r'^0*(\d+[a-zA-Z]?)[\.\s\-_]', d)
        if m_d and m_d.group(1).upper() == clean_num:
            return d

    # 4. Fallback startswith
    for d in dir_entries:
        if (d.startswith(f"{clean_num}.") or d.startswith(f"{clean_num} ") or 
            d.startswith(f"{val_str}.") or d.startswith(f"{val_str} ") or
            d.startswith(f"{base_digits}.") or d.startswith(f"{base_digits} ")):
            return d

    return str(input_val)


def generate_csv_for_inputs(input_string):
    inputs = [x.strip() for x in input_string.replace(',', ' ').split() if x.strip()]

    if not inputs:
        print("No valid inputs provided.")
        return

    if not os.path.exists(EXCEL_PATH):
        print(f"[CRITICAL ERROR] Master List.xlsx not found at:\n{EXCEL_PATH}")
        return

    print(f"Active Workspace: {BASE_DIR}")
    print(f"Opening Excel via COM to process {len(inputs)} input(s)...")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(EXCEL_PATH)

        # Ensure Master_List sheet exists
        sheet_names = [s.Name for s in wb.Sheets]
        master_sheet_name = "Master_List" if "Master_List" in sheet_names else sheet_names[0]
        ws_master = wb.Sheets(master_sheet_name)

        for input_val in inputs:
            csv_path = os.path.join(MAPPING_DIR, f"{input_val}.csv")
            print(f"\n--- Processing: {input_val} ---")

            # 2. Resolve RAW MATERIAL folder path matching leading number
            subfolder = resolve_rm_subfolder(RM_DIR, input_val)
            cand_ir = os.path.join(RM_DIR, subfolder, "IR")
            if os.path.isdir(cand_ir):
                target_path = cand_ir
            elif os.path.isdir(os.path.join(RM_DIR, subfolder)):
                target_path = os.path.join(RM_DIR, subfolder)
            else:
                target_path = os.path.join(RM_DIR, subfolder, "IR")
                print(f"[WARNING] Expected folder not found: {target_path}")

            # Format AY1 cell as text so leading zeros (e.g. 079) are preserved
            ws_master.Range("AY1").NumberFormat = "@"
            ws_master.Range("AY1").Value = "'" + str(input_val)

            try:
                ws_master.Range("AZ1").Value = target_path
            except Exception:
                pass

            excel.Calculate()
            wb.Save()

            try:
                # 1. Load default sheet: Image_Mapping_RMU
                sheet_to_read = "Image_Mapping_RMU" if "Image_Mapping_RMU" in sheet_names else sheet_names[0]
                df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_to_read)

                # 3. Format Column G (index 6) to HH:MM AM/PM
                if df.shape[1] > 6:
                    col_g_name = df.columns[6]

                    def _fmt(val):
                        try:
                            t = pd.to_datetime(str(val), errors='coerce')
                            if pd.isna(t):
                                return val
                            return t.strftime('%I:%M %p').lstrip('0')
                        except Exception:
                            return val

                    df[col_g_name] = df[col_g_name].apply(_fmt)

                # 4. Inject target path into Column BG (index 58 / 'Folder')
                if 'Folder' in df.columns:
                    df['Folder'] = target_path
                else:
                    required_idx = 58
                    if df.shape[1] <= required_idx:
                        for idx_add in range(required_idx - df.shape[1] + 1):
                            df[f"extra_{idx_add}"] = None
                    df.iloc[:, required_idx] = target_path

                if df.empty:
                    print(f"Error: Sheet '{sheet_to_read}' is empty.")
                    continue

                swg_type_col = '{{ swg.type }}'
                if swg_type_col not in df.columns:
                    print(f"Error: Column '{swg_type_col}' not found in sheet.")
                    continue

                swg_type = str(df[swg_type_col].iloc[0]).strip()
                print(f"Detected Switchgear Type: {swg_type}")

                # 5. Reload dataframe if H-POLE or VCB
                if "H-POLE" in swg_type.upper() and "Image_Mapping_H-POLE" in sheet_names:
                    df = pd.read_excel(EXCEL_PATH, sheet_name="Image_Mapping_H-POLE")
                    print("Reloaded dataframe using Image_Mapping_H-POLE sheet")
                    if 'Folder' in df.columns:
                        df['Folder'] = target_path
                    elif df.shape[1] > 58:
                        df.iloc[:, 58] = target_path

                elif "VCB" in swg_type.upper() and "Image_Mapping_VCB" in sheet_names:
                    df = pd.read_excel(EXCEL_PATH, sheet_name="Image_Mapping_VCB")
                    print("Reloaded dataframe using Image_Mapping_VCB sheet")
                    if 'Folder' in df.columns:
                        df['Folder'] = target_path
                    elif df.shape[1] > 58:
                        df.iloc[:, 58] = target_path

                # 6. Apply heater amp formatting & transition / bus section panel rules
                def _fmt_heater_amp(val):
                    if val is None or pd.isna(val):
                        return "-"
                    s = str(val).strip()
                    if not s or s in ['-', 'N/A', '', '.', '/'] or s.lower() in ['n/a', 'nan', 'none', 'null']:
                        return "-"
                    m = re.search(r'[-+]?\d*\.?\d+', s)
                    if m:
                        try:
                            f = float(m.group(0))
                            return f"{f:.2f}A"
                        except ValueError:
                            pass
                    return s

                if '{{ panel.heateramp }}' in df.columns:
                    df['{{ panel.heateramp }}'] = df['{{ panel.heateramp }}'].apply(_fmt_heater_amp)

                is_vcb_df = "VCB" in swg_type.upper()
                if '{{ panel.name }}' in df.columns:
                    def _apply_panel_rules(row):
                        p_name = str(row.get('{{ panel.name }}', '') or '').upper().strip()
                        is_trans = ("TRANSITION" in p_name)
                        is_bus_sec = ("BUS SECTION" in p_name)

                        if is_vcb_df and is_trans:
                            if '{{ panel.cabletype }}' in row.index: row['{{ panel.cabletype }}'] = '-'
                            if '{{ panel.linknumber }}' in row.index: row['{{ panel.linknumber }}'] = '-'
                        elif is_bus_sec:
                            if '{{ panel.cabletype }}' in row.index: row['{{ panel.cabletype }}'] = '-'
                        return row

                    df = df.apply(_apply_panel_rules, axis=1)

                # 7. Save as CSV
                valid_types = ["RMU SF6", "MRMU", "H-POLE", "VCB", "RMU"]
                if any(vt in swg_type.upper() for vt in valid_types):
                    df.to_csv(csv_path, index=False)
                    print(f"Success: CSV generated at {csv_path}")
                else:
                    print(f"Skipped: Switchgear type is '{swg_type}', not in supported list.")

            except Exception as e:
                print(f"Error reading/saving CSV for {input_val}: {e}")

    except Exception as e:
        print(f"Critical Error with Excel COM Automation: {e}")
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()
        print("\nFinished processing all inputs.")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        generate_csv_for_inputs(" ".join(sys.argv[1:]))
    else:
        user_input = input("Enter the input values (e.g., 060, 062, 079): ")
        if user_input.strip():
            generate_csv_for_inputs(user_input.strip())