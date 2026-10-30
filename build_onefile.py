import os
import subprocess

APP_DIR = r"C:\Users\aremi\Desktop\REPORT CBM\WRAITH APP"
SRC_DIR = os.path.join(APP_DIR, "src")
DIST_DIR = os.path.join(APP_DIR, "dist")
WORK_DIR = os.path.join(APP_DIR, "build")
SPEC_DIR = APP_DIR
MAIN_GUI = os.path.join(SRC_DIR, "app_gui.py")
ICON_PATH = os.path.join(SRC_DIR, "app_icon.ico")
TEMPLATES_DIR = os.path.join(APP_DIR, "templates")

def build_onefile():
    print("--- Building Single-File Executable (onefile) ---")
    cmd = [
        "python", "-m", "PyInstaller",
        "-y",
        "--noconsole",
        "--onefile",
        "--name=WRAITH_APP_SingleFile",
        f"--icon={ICON_PATH}",
        f"--distpath={DIST_DIR}",
        f"--workpath={WORK_DIR}",
        f"--specpath={SPEC_DIR}",
        f"--add-data={SRC_DIR};src",
        f"--add-data={TEMPLATES_DIR};templates",
        f"--add-data={ICON_PATH};src",
        MAIN_GUI
    ]
    res = subprocess.run(cmd, cwd=APP_DIR, capture_output=True, text=True)
    if res.returncode == 0:
        print("[SUCCESS] Single-file executable built successfully!")
        target_exe = os.path.join(DIST_DIR, "WRAITH_APP_SingleFile.exe")
        if os.path.exists(target_exe):
            print(f"Location: {target_exe}")
            print(f"Size: {round(os.path.getsize(target_exe)/(1024*1024), 2)} MB")
    else:
        print("[FAILED] Build failed:")
        print(res.stderr[-1000:])

if __name__ == "__main__":
    build_onefile()
