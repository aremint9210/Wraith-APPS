import asyncio
import os
import zipfile
import re
import pandas as pd
from PIL import Image
from playwright.async_api import async_playwright

# --- Configuration ---
ROOT_FOLDER = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42312390 - TNBD BANGI - IR US - TIANU\03. RAW MATERIAL"
EXPORT_ROOT = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42312390 - TNBD BANGI - IR US - TIANU\03. UE SCREENSHOT"
MASTER_EXCEL_PATH = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42312390 - TNBD BANGI - IR US - TIANU\99. GENERATE REPORT\EXCELL MAPPING\Master List.xlsx"

# Full HD Resolution (16:9 Ratio)
RENDER_WIDTH = 1920
RENDER_HEIGHT = 1080
DEVICE_SCALE_FACTOR = 1  # Standard 1:1 pixel rendering at 1920x1080

# Calculated DPI to force 3.51 cm height in Word/Excel/PDFs:
# Formula: (1080 px / 3.51 cm) * 2.54 cm/inch = ~781 DPI
# With Scale Factor 1: 781 DPI -> 3.51cm Height x 6.24cm Width
TARGET_DPI = int((RENDER_HEIGHT / 3.51) * 2.54)

MAX_CONCURRENT_PAGES = 4
sem = asyncio.Semaphore(MAX_CONCURRENT_PAGES)


def extract_number(folder_name):
    """Extracts leading or isolated number from folder name (e.g. '277' from '277. PE BANTING TOWN SSU NO 2')."""
    match = re.search(r'(\d+)', folder_name)
    return int(match.group(1)) if match else -1


def load_sheet_mapping(excel_path, sheet_name):
    """
    Loads a specific sheet from Master List.xlsx.
    Column B = Panel Code (e.g. g05151)
    Column A = Panel Prefix (e.g. Panel9 or 09)
    Returns a dictionary mapping lower-cased Panel Code -> Cleaned Panel Prefix.
    """
    if not os.path.exists(excel_path):
        return {}

    try:
        xls = pd.ExcelFile(excel_path)
        sheet_to_load = None
        for s in xls.sheet_names:
            if str(s).strip() == str(sheet_name).strip():
                sheet_to_load = s
                break

        if not sheet_to_load:
            return {}

        df = pd.read_excel(xls, sheet_name=sheet_to_load)

        col_a = df.iloc[:, 0].astype(str).str.strip()
        col_b = df.iloc[:, 1].astype(str).str.strip()

        mapping = {}
        for panel_id, panel_code in zip(col_a, col_b):
            if panel_code and panel_code.lower() != 'nan':
                match = re.search(r'(\d+)', panel_id)
                num_str = f"{int(match.group(1)):02d}" if match else panel_id
                mapping[panel_code.lower()] = num_str

        return mapping

    except Exception as e:
        print(f"    [WARN EXCEL] Failed to load Sheet '{sheet_name}': {e}")
        return {}


async def convert_html_to_jpg(browser_context, html_path, export_station_path, custom_filename):
    async with sem:
        final_path = os.path.join(export_station_path, custom_filename)

        if os.path.exists(final_path):
            return

        page = await browser_context.new_page()
        try:
            file_url = f"file:///{os.path.abspath(html_path)}".replace('\\', '/')
            await page.goto(file_url, wait_until="networkidle")

            await page.evaluate("document.fonts.ready")
            await asyncio.sleep(0.5)

            # 1. Capture raw 1920x1080 PNG
            await page.screenshot(
                path=final_path,
                type="png",
                full_page=True,
                animations="disabled"
            )

            # 2. Inject DPI metadata without modifying or compressing raw pixels
            with Image.open(final_path) as img:
                img.save(final_path, dpi=(TARGET_DPI, TARGET_DPI), compress_level=0)

            print(f"    [SAVED] {custom_filename}")

        except Exception as e:
            print(f"    [ERR] {html_path}: {e}")
        finally:
            await page.close()


async def process_station(browser_context, pe_folder):
    pe_path = os.path.join(ROOT_FOLDER, pe_folder)

    us_path = None
    for item in os.listdir(pe_path):
        if item.upper() == "US" and os.path.isdir(os.path.join(pe_path, item)):
            us_path = os.path.join(pe_path, item)
            break

    if not us_path:
        print(f"\n[SKIP] No 'US' folder found in: {pe_folder}")
        return

    export_station_path = os.path.join(EXPORT_ROOT, pe_folder)
    os.makedirs(export_station_path, exist_ok=True)

    station_num = extract_number(pe_folder)
    print(f"\n>>> Processing Station {station_num}: {pe_folder}")

    panel_mapping = load_sheet_mapping(MASTER_EXCEL_PATH, station_num)

    # 1. Unzip any archives inside US folder
    for root, dirs, files in os.walk(us_path):
        for file in files:
            if file.lower().endswith(".zip"):
                zip_p = os.path.join(root, file)
                ext_p = os.path.join(root, os.path.splitext(file)[0])
                if not os.path.exists(ext_p):
                    try:
                        print(f"    [UNZIP] Extracting {file}...")
                        with zipfile.ZipFile(zip_p, 'r') as z:
                            z.extractall(ext_p)
                    except Exception as e:
                        print(f"    [ERR UNZIP] {file}: {e}")

    # 2. Process Panel Folders
    tasks = []

    for root, dirs, files in os.walk(us_path):
        if "resources" in root.lower():
            continue

        subdirs = [d for d in dirs if not d.lower() == "resources"]
        if not subdirs:
            continue

        valid_measurement_dirs = []
        for d in subdirs:
            subdir_path = os.path.join(root, d)
            has_html = any(f.lower().endswith('.html') for f in os.listdir(subdir_path) if
                           os.path.isfile(os.path.join(subdir_path, f)))
            if has_html:
                valid_measurement_dirs.append(d)

        valid_measurement_dirs.sort()

        panel_folder_name = os.path.basename(root)
        panel_key = panel_folder_name.lower().strip()

        prefix = panel_mapping.get(panel_key, panel_folder_name)

        for idx, meas_dir in enumerate(valid_measurement_dirs):
            suffix = f".{idx + 1}"

            meas_path = os.path.join(root, meas_dir)
            for file in os.listdir(meas_path):
                if file.lower().endswith(".html"):
                    html_full_path = os.path.join(meas_path, file)

                    png_name = f"{prefix}{suffix}.png"
                    png_name = re.sub(r'[\\/*?:"<>|]', '_', png_name)

                    tasks.append(convert_html_to_jpg(browser_context, html_full_path, export_station_path, png_name))

    if tasks:
        await asyncio.gather(*tasks)


async def main():
    os.makedirs(EXPORT_ROOT, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--enable-font-antialiasing",
                "--high-dpi-support=1"
            ]
        )

        context = await browser.new_context(
            viewport={'width': RENDER_WIDTH, 'height': RENDER_HEIGHT},
            device_scale_factor=DEVICE_SCALE_FACTOR
        )

        all_folders = [f for f in os.listdir(ROOT_FOLDER) if os.path.isdir(os.path.join(ROOT_FOLDER, f))]
        all_folders.sort(key=extract_number)

        print(f"Starting Processing via Master List.xlsx Sheet Lookup")
        print(f"Export Location: {EXPORT_ROOT}")

        for folder in all_folders:
            await process_station(context, folder)

        await browser.close()
        print(f"\n--- Batch Finished. 1920x1080 PNGs with 3.51cm Height Metadata saved in '03. UE SCREENSHOT' ---")


if __name__ == "__main__":
    asyncio.run(main())