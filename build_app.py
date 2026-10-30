import os
import sys
import shutil
import subprocess

APPS_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(APPS_DIR, "src")
TEMPLATES_DIR = os.path.join(APPS_DIR, "templates")
GUI_PATH = os.path.join(SRC_DIR, "app_gui.py")
ICON_PATH = os.path.join(SRC_DIR, "app_icon.ico")

import customtkinter
CTK_DIR = os.path.dirname(customtkinter.__file__)
CTK_ASSETS = os.path.join(CTK_DIR, "assets")

print("--- Step 1: Preparing Standalone App Source ---")

dist_dir = os.path.join(APPS_DIR, "dist")
build_dir = os.path.join(APPS_DIR, "build")

# Command to build standalone GUI app with CustomTkinter assets
cmd = [
    sys.executable, "-m", "PyInstaller",
    "-y",
    "--noconsole",
    "--name=WRAITH_APP",
    f"--icon={ICON_PATH}",
    f"--distpath={dist_dir}",
    f"--workpath={build_dir}",
    f"--specpath={APPS_DIR}",
    "--collect-all=customtkinter",
    f"--add-data={CTK_ASSETS};customtkinter/assets",
    f"--add-data={SRC_DIR};src",
    f"--add-data={TEMPLATES_DIR};templates",
    f"--add-data={ICON_PATH};src",
    GUI_PATH
]

print("Running command:", " ".join(cmd))
result = subprocess.run(cmd, cwd=APPS_DIR)

if result.returncode == 0:
    app_folder = os.path.join(dist_dir, "WRAITH_APP")
    
    # Extra safety: ensure customtkinter assets directory is physically present in _internal
    target_ctk_assets = os.path.join(app_folder, "_internal", "customtkinter", "assets")
    if not os.path.exists(target_ctk_assets) and os.path.exists(CTK_ASSETS):
        shutil.copytree(CTK_ASSETS, target_ctk_assets, dirs_exist_ok=True)
        print(f"[ASSET SYNC] Copied CustomTkinter assets to: {target_ctk_assets}")

    zip_path = os.path.join(dist_dir, "WRAITH_APP_Portable")
    
    print("\n--- Step 2: Creating Portable ZIP Package ---")
    shutil.make_archive(zip_path, "zip", dist_dir, "WRAITH_APP")
    
    print("\n[SUCCESS] BUILD COMPLETED SUCCESSFULLY!")
    print(f"Standalone App Location: {os.path.join(app_folder, 'WRAITH_APP.exe')}")
    print(f"Portable Zip Package: {zip_path}.zip")
else:
    print(f"\n[FAILED] BUILD FAILED with exit code {result.returncode}")
