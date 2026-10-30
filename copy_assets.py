import os
import shutil

BASE_DIR = r"C:\Users\aremi\Desktop\REPORT CBM"
APPS_DIR = r"C:\Users\aremi\Desktop\REPORT CBM\WRAITH APP"
SRC_DIR = os.path.join(APPS_DIR, "src")
TEMPLATES_DIR = os.path.join(APPS_DIR, "templates")

os.makedirs(SRC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 1. Copy SCRIPT files to APPS IDEA/src
script_source = os.path.join(BASE_DIR, "SCRIPT")
if os.path.exists(script_source):
    for item in os.listdir(script_source):
        s = os.path.join(script_source, item)
        d = os.path.join(SRC_DIR, item)
        if os.path.isfile(s):
            shutil.copy2(s, d)
        elif os.path.isdir(s) and item != "__pycache__":
            shutil.copytree(s, d, dirs_exist_ok=True)
    print("Copied SCRIPT files to src/")

# 2. Copy TEMPLATE files to APPS IDEA/templates
template_source = os.path.join(BASE_DIR, "TEMPLATE")
if os.path.exists(template_source):
    for item in os.listdir(template_source):
        s = os.path.join(template_source, item)
        d = os.path.join(TEMPLATES_DIR, item)
        if os.path.isfile(s):
            shutil.copy2(s, d)
        elif os.path.isdir(s):
            shutil.copytree(s, d, dirs_exist_ok=True)
    print("Copied TEMPLATE files to templates/")
