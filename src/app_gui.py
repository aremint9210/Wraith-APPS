import csv
import json
import os
import re
import shutil
import subprocess
import threading
import time
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import customtkinter as ctk

# Suppress openpyxl user warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# Session state file — lives beside the substation folder, named per-workspace
SESSION_FILENAME = ".wraith_session.json"

# ==============================================================================
#  DESIGN SYSTEM :: GLASS / BENTO DARK THEME
# ==============================================================================
COLOR_CANVAS = "#000000"   # True OLED Black — pixels OFF
COLOR_RAIL   = "#000000"   # Left Dock — same true black
COLOR_CARD   = "#0A0A0A"   # Glass Surface — barely off-black
COLOR_CARD_HOVER  = "#141414"   # Card Hover
COLOR_INNER  = "#050505"   # Inset Terminal / List Background
COLOR_BORDER = "#1A1A1A"   # Hairline Border — neutral dark, no blue tint
COLOR_BORDER_FOCUS = "#38BDF8"  # Active Focus Ring (Cyan Glow)

# Typography & Accents
COLOR_CYAN        = "#38BDF8"   # Primary Execution Accent
COLOR_CYAN_HOVER  = "#0284C7"   # Primary Execution Hover State
COLOR_INDIGO      = "#818CF8"   # Secondary Process Accent
COLOR_VIOLET      = "#A78BFA"   # Tertiary Accent
COLOR_EMERALD     = "#34D399"   # Ready / Status Healthy Accent
COLOR_AMBER       = "#FBBF24"   # Warning / KIV State
COLOR_ROSE        = "#F43F5E"   # Error / Destructive State
COLOR_TEXT_MAIN   = "#F8FAFC"   # High Contrast Text
COLOR_TEXT_MUTED  = "#64748B"   # Secondary Descriptive Labels
COLOR_TEXT_DIM    = "#334155"   # Deeply muted / disabled

# Console log-level palette
LOG_COLORS = {
    "INFO":  COLOR_CYAN,
    "OK":    COLOR_EMERALD,
    "WARN":  COLOR_AMBER,
    "ERROR": COLOR_ROSE,
    "CMD":   COLOR_INDIGO,
    "SYS":   COLOR_VIOLET,
}

FONT_FAMILY = "Segoe UI"
FONT_MONO   = "Consolas"

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

import sys

# Dynamic root & script directory detection for Standalone App
if getattr(sys, 'frozen', False):
    APP_DIR = sys._MEIPASS
    CENTRAL_SCRIPT_DIR = os.path.join(APP_DIR, "src")
    DEFAULT_ROOT_DIR = r"C:\Users\aremi\Desktop\REPORT CBM"
    if not os.path.exists(DEFAULT_ROOT_DIR):
        DEFAULT_ROOT_DIR = os.path.dirname(sys.executable)
else:
    CURRENT_FILE_DIR = os.path.dirname(os.path.abspath(__file__))
    CENTRAL_SCRIPT_DIR = CURRENT_FILE_DIR
    DEFAULT_ROOT_DIR = r"C:\Users\aremi\Desktop\REPORT CBM"
    if not os.path.exists(DEFAULT_ROOT_DIR):
        DEFAULT_ROOT_DIR = os.path.abspath(os.path.join(CURRENT_FILE_DIR, "..", ".."))

if not os.path.exists(CENTRAL_SCRIPT_DIR):
    CENTRAL_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))



# ==============================================================================
#  PERSISTENT CONFIGURATION & SETTINGS SYSTEM
# ==============================================================================
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".wraith_config.json")

def load_wraith_config():
    default_root = r"C:\Users\aremi\Desktop\REPORT CBM"
    if not os.path.exists(default_root):
        curr = os.path.abspath(os.path.dirname(__file__))
        parent = os.path.abspath(os.path.join(curr, ".."))
        if os.path.exists(parent) and any(os.path.isdir(os.path.join(parent, f)) for f in os.listdir(parent) if f.upper().startswith("PO") or "REPORT" in f.upper()):
            default_root = parent
        else:
            default_root = os.getcwd()
            
    config = {
        "root_workspace_dir": default_root,
        "active_po_dir": "",
        "custom_paths": {}
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    config.update(loaded)
        except Exception:
            pass
            
    if not os.path.exists(config.get("root_workspace_dir", "")):
        config["root_workspace_dir"] = default_root

    return config

def save_wraith_config(config):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

INITIAL_CONFIG = load_wraith_config()

def auto_detect_scope_templates(sub_dir, root_dir):
    """Auto-detects template folders for IR, IR+UE, and IR+UE+TEV scopes."""
    candidates_roots = []
    if sub_dir:
        candidates_roots.append(os.path.join(sub_dir, "05. TEMPLATE"))
        candidates_roots.append(os.path.join(sub_dir, "TEMPLATE"))
    if root_dir:
        candidates_roots.append(os.path.join(root_dir, "TEMPLATE"))

    found_ir, found_ir_ue, found_ir_ue_tev = "", "", ""

    for c_root in candidates_roots:
        if not os.path.exists(c_root):
            continue
        subdirs = [os.path.join(c_root, d) for d in os.listdir(c_root) if os.path.isdir(os.path.join(c_root, d))]
        for d in subdirs:
            name = os.path.basename(d).upper()
            if "TEV" in name:
                if not found_ir_ue_tev:
                    found_ir_ue_tev = d
            elif "UE" in name or "US" in name:
                if not found_ir_ue:
                    found_ir_ue = d
            elif "IR" in name and not any(x in name for x in ["UE", "US", "TEV"]):
                if not found_ir:
                    found_ir = d
        if not found_ir and any(f.endswith(".docx") for f in os.listdir(c_root)):
            found_ir = c_root

    return found_ir, found_ir_ue, found_ir_ue_tev


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("WRAITH // Settings & Path Configuration")
        self.geometry("760x680")
        self.resizable(False, False)
        self.configure(fg_color=COLOR_CANVAS)
        self.attributes("-topmost", True)

        self.entries = {}
        self._build_ui()

    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=0, border_width=1, border_color=COLOR_BORDER)
        header.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(
            header, text="⚙️ WORKSPACE & PATH CONFIGURATION",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", padx=20, pady=(16, 4))
        ctk.CTkLabel(
            header, text="Configure main workspace location and custom subfolder overrides.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 16))

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=20, pady=16)

        # 1. Main Workspace Root
        self._add_setting_row(scroll, "root_workspace_dir", "Workspace Root Directory (REPORT CBM):", self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR))

        # 2. Active PO Directory
        self._add_setting_row(scroll, "active_po_dir", "Active PO / Substation Directory:", self.parent.target_substation_dir or "")
        if "active_po_dir" in self.entries:
            self.entries["active_po_dir"].bind("<FocusOut>", lambda e: self._update_subfolders_for_po(self.entries["active_po_dir"].get().strip()))
            self.entries["active_po_dir"].bind("<Return>", lambda e: self._update_subfolders_for_po(self.entries["active_po_dir"].get().strip()))

        ctk.CTkLabel(
            scroll, text="SUBFOLDER PATH OVERRIDES (OPTIONAL)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_INDIGO
        ).pack(anchor="w", pady=(16, 2))

        ctk.CTkLabel(
            scroll, text="Automatically detected from Active PO Directory. Standard folders auto-sync when workspace changes.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", pady=(0, 8))

        sub_dir = self.parent.target_substation_dir or ""
        active_po_name = os.path.basename(sub_dir) if sub_dir else "default"
        customs = self.parent.config.get("custom_paths", {})
        po_customs = customs.get(active_po_name, {}) if isinstance(customs.get(active_po_name), dict) else customs

        def _get_init_val(key, standard_name):
            val = po_customs.get(key)
            if val and os.path.exists(val):
                return val
            return os.path.join(sub_dir, standard_name) if sub_dir else ""

        def _get_tmpl_init_val():
            val = po_customs.get("template_dir")
            if val and os.path.exists(val):
                return val
            local_tmpl = os.path.join(sub_dir, "05. TEMPLATE") if sub_dir else ""
            if local_tmpl and os.path.exists(local_tmpl):
                return local_tmpl
            root_tmpl = os.path.join(self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR), "TEMPLATE")
            if root_tmpl and os.path.exists(root_tmpl):
                return root_tmpl
            return local_tmpl

        folders = [
            ("raw_material", "03. RAW MATERIAL Folder:", _get_init_val("raw_material", "03. RAW MATERIAL")),
            ("quick_report", "02. QUICK REPORT Folder:", _get_init_val("quick_report", "02. QUICK REPORT")),
            ("full_report", "01. FULL REPORT Folder:", _get_init_val("full_report", "01. FULL REPORT")),
            ("testsheet", "04. TESTSHEET Folder:", _get_init_val("testsheet", "04. TESTSHEET")),
            ("template_dir", "05. TEMPLATE Folder (Word .docx Templates):", _get_tmpl_init_val()),
            ("excel_mapping", "06. EXCELL MAPPING Folder:", _get_init_val("excel_mapping", "06. EXCELL MAPPING")),
            ("output_file", "07. OUTPUT FILE Folder:", _get_init_val("output_file", "07. OUTPUT FILE")),
        ]

        for key, label_text, default_val in folders:
            self._add_setting_row(scroll, key, label_text, default_val)

        # 3. Gemini Vision AI API Configuration
        ctk.CTkLabel(
            scroll, text="🤖 AI VISION & REASONING (GEMINI PRO / FLASH)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", pady=(18, 4))

        ctk.CTkLabel(
            scroll, text="Get your free API Key from aistudio.google.com to unlock 100% multimodal visual reasoning.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", pady=(0, 8))

        saved_key = self.parent.config.get("gemini_api_key") or os.environ.get("GEMINI_API_KEY") or ""
        self._add_api_key_row(scroll, "gemini_api_key", "Google Gemini API Key (AIzaSy...):", saved_key)

        # 4. Executive Summary Scope Ranges (Manual Configuration)
        ctk.CTkLabel(
            scroll, text="📊 EXECUTIVE SUMMARY SCOPE RANGES (MANUAL CONFIGURATION)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", pady=(18, 4))

        ctk.CTkLabel(
            scroll, text="Specify substation number ranges for each test type (e.g. '1-400', '51-120, 200-250', 'ALL', or leave blank if none).",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", pady=(0, 8))

        active_po_name = os.path.basename(sub_dir) if sub_dir else "default"
        po_scopes = self.parent.config.get("scope_ranges", {}).get(active_po_name, {})
        if sub_dir and os.path.exists(os.path.join(sub_dir, "scope_config.json")):
            try:
                with open(os.path.join(sub_dir, "scope_config.json"), "r", encoding="utf-8") as f_sc:
                    file_scopes = json.load(f_sc)
                    for k in ["ir_range", "us_range", "tev_range"]:
                        if k in file_scopes and k not in po_scopes:
                            po_scopes[k] = file_scopes[k]
            except Exception:
                pass

        self._add_text_setting_row(scroll, "scope_ir_range", "IR Substation Range (e.g. 1-400 or ALL):", po_scopes.get("ir_range", "ALL"))
        self._add_text_setting_row(scroll, "scope_us_range", "Ultrasound (US) Range (e.g. 51-120):", po_scopes.get("us_range", ""))
        self._add_text_setting_row(scroll, "scope_tev_range", "TEV Substation Range (e.g. 101-150):", po_scopes.get("tev_range", ""))

        # 5. Template Paths by Test Scope (IR / IR+UE / IR+UE+TEV)
        ctk.CTkLabel(
            scroll, text="📑 TEMPLATE PATHS BY TEST SCOPE (AUTO-ROUTING)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", pady=(18, 4))

        ctk.CTkLabel(
            scroll, text="Assign specific template folders for each test scope. If folder names change, browse to update them here without altering any scripts.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", pady=(0, 8))

        det_ir, det_ir_ue, det_ir_ue_tev = auto_detect_scope_templates(sub_dir, self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR))
        val_ir = po_scopes.get("template_ir") or po_customs.get("template_ir") or det_ir
        val_ir_ue = po_scopes.get("template_ir_ue") or po_customs.get("template_ir_ue") or det_ir_ue
        val_ir_ue_tev = po_scopes.get("template_ir_ue_tev") or po_customs.get("template_ir_ue_tev") or det_ir_ue_tev

        self._add_setting_row(scroll, "template_ir", "1. IR Only Template Folder (e.g. IR - TIANU):", val_ir)
        self._add_setting_row(scroll, "template_ir_ue", "2. IR + UE / Ultrasound Template Folder (e.g. IR UE - TIANU):", val_ir_ue)
        self._add_setting_row(scroll, "template_ir_ue_tev", "3. IR + UE + TEV Template Folder (e.g. IR UE TEV):", val_ir_ue_tev)

        actions = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=0)
        actions.pack(fill="x", side="bottom", padx=0, pady=0)

        reset_btn = ctk.CTkButton(
            actions, text="↺ Reset to Substation Defaults",
            fg_color="transparent", hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MUTED, border_width=1, border_color=COLOR_BORDER,
            command=self._reset_defaults
        )
        reset_btn.pack(side="left", padx=20, pady=16)

        save_btn = ctk.CTkButton(
            actions, text="Save & Apply Settings",
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER,
            text_color="#000000", font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            command=self._save_settings
        )
        save_btn.pack(side="right", padx=20, pady=16)

    def _add_api_key_row(self, container, key, label_text, current_val):
        frame = ctk.CTkFrame(container, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        frame.pack(fill="x", pady=6)

        ctk.CTkLabel(
            frame, text=label_text,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w", padx=12, pady=(10, 4))

        input_frame = ctk.CTkFrame(frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=12, pady=(0, 10))

        entry = ctk.CTkEntry(
            input_frame, fg_color=COLOR_INNER, border_color=COLOR_BORDER,
            text_color=COLOR_CYAN, font=ctk.CTkFont(family=FONT_MONO, size=11),
            show="•"
        )
        entry.insert(0, current_val)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entries[key] = entry

        def _toggle_show():
            if entry.cget("show") == "•":
                entry.configure(show="")
                show_btn.configure(text="Hide")
            else:
                entry.configure(show="•")
                show_btn.configure(text="Show")

        show_btn = ctk.CTkButton(
            input_frame, text="Show", width=60,
            fg_color=COLOR_BORDER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=_toggle_show
        )
        show_btn.pack(side="right")

    def _add_text_setting_row(self, container, key, label_text, current_val):
        frame = ctk.CTkFrame(container, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        frame.pack(fill="x", pady=6)

        ctk.CTkLabel(
            frame, text=label_text,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w", padx=12, pady=(10, 4))

        input_frame = ctk.CTkFrame(frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=12, pady=(0, 10))

        entry = ctk.CTkEntry(
            input_frame, fg_color=COLOR_INNER, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MAIN, font=ctk.CTkFont(family=FONT_MONO, size=11)
        )
        entry.insert(0, str(current_val) if current_val is not None else "")
        entry.pack(side="left", fill="x", expand=True)
        self.entries[key] = entry

    def _add_setting_row(self, container, key, label_text, current_val):
        frame = ctk.CTkFrame(container, fg_color=COLOR_CARD, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        frame.pack(fill="x", pady=6)

        ctk.CTkLabel(
            frame, text=label_text,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w", padx=12, pady=(10, 4))

        input_frame = ctk.CTkFrame(frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=12, pady=(0, 10))

        entry = ctk.CTkEntry(
            input_frame, fg_color=COLOR_INNER, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MAIN, font=ctk.CTkFont(family=FONT_MONO, size=11)
        )
        entry.insert(0, current_val)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entries[key] = entry

        browse_btn = ctk.CTkButton(
            input_frame, text="📁 Browse...", width=80,
            fg_color=COLOR_BORDER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=lambda k=key: self._browse(k)
        )
        browse_btn.pack(side="right")

        clear_btn = ctk.CTkButton(
            input_frame, text="✕ Clear", width=60,
            fg_color="transparent", hover_color=COLOR_CARD_HOVER,
            border_width=1, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MUTED, command=lambda k=key: self._clear(k)
        )
        clear_btn.pack(side="right", padx=(0, 6))

    def _clear(self, key):
        if key in self.entries:
            self.entries[key].delete(0, tk.END)

    def _browse(self, key):
        current = self.entries[key].get().strip() if key in self.entries else ""
        if current and os.path.exists(current):
            init = current
        elif key in ["template_dir", "template_ir", "template_ir_ue", "template_ir_ue_tev"]:
            local_tmpl = os.path.join(self.parent.target_substation_dir or "", "05. TEMPLATE")
            root_tmpl = os.path.join(self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR), "TEMPLATE")
            if os.path.exists(local_tmpl):
                init = local_tmpl
            elif os.path.exists(root_tmpl):
                init = root_tmpl
            else:
                init = self.parent.target_substation_dir or DEFAULT_ROOT_DIR
        else:
            init = self.parent.target_substation_dir or DEFAULT_ROOT_DIR

        selected = filedialog.askdirectory(initialdir=init, title=f"Select Path for {key}")
        if selected:
            norm_sel = os.path.normpath(selected)
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, norm_sel)
            if key == "active_po_dir":
                self._update_subfolders_for_po(norm_sel)

    def _update_subfolders_for_po(self, po_dir):
        if not po_dir or not os.path.exists(po_dir):
            return
        local_tmpl = os.path.join(po_dir, "05. TEMPLATE")
        root_tmpl = os.path.join(self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR), "TEMPLATE")
        tmpl_val = local_tmpl if os.path.exists(local_tmpl) else (root_tmpl if os.path.exists(root_tmpl) else local_tmpl)

        defaults = {
            "raw_material": os.path.join(po_dir, "03. RAW MATERIAL"),
            "quick_report": os.path.join(po_dir, "02. QUICK REPORT"),
            "full_report": os.path.join(po_dir, "01. FULL REPORT"),
            "testsheet": os.path.join(po_dir, "04. TESTSHEET"),
            "template_dir": tmpl_val,
            "excel_mapping": os.path.join(po_dir, "06. EXCELL MAPPING"),
            "output_file": os.path.join(po_dir, "07. OUTPUT FILE"),
        }
        for k, v in defaults.items():
            if k in self.entries:
                self.entries[k].delete(0, tk.END)
                self.entries[k].insert(0, v)

        det_ir, det_ir_ue, det_ir_ue_tev = auto_detect_scope_templates(po_dir, self.parent.config.get("root_workspace_dir", DEFAULT_ROOT_DIR))
        if "template_ir" in self.entries:
            self.entries["template_ir"].delete(0, tk.END)
            self.entries["template_ir"].insert(0, det_ir)
        if "template_ir_ue" in self.entries:
            self.entries["template_ir_ue"].delete(0, tk.END)
            self.entries["template_ir_ue"].insert(0, det_ir_ue)
        if "template_ir_ue_tev" in self.entries:
            self.entries["template_ir_ue_tev"].delete(0, tk.END)
            self.entries["template_ir_ue_tev"].insert(0, det_ir_ue_tev)

    def _reset_defaults(self):
        sub = self.entries.get("active_po_dir").get().strip() or self.parent.target_substation_dir
        if not sub:
            return
        self._update_subfolders_for_po(sub)
        if "scope_ir_range" in self.entries:
            self.entries["scope_ir_range"].delete(0, tk.END)
            self.entries["scope_ir_range"].insert(0, "ALL")
        if "scope_us_range" in self.entries:
            self.entries["scope_us_range"].delete(0, tk.END)
        if "scope_tev_range" in self.entries:
            self.entries["scope_tev_range"].delete(0, tk.END)

    def _save_settings(self):
        global DEFAULT_ROOT_DIR
        new_root = self.entries.get("root_workspace_dir").get().strip()
        new_po = self.entries.get("active_po_dir").get().strip()

        if new_root and os.path.exists(new_root):
            self.parent.config["root_workspace_dir"] = new_root
            DEFAULT_ROOT_DIR = new_root

        if new_po and os.path.exists(new_po):
            self.parent.target_substation_dir = new_po
            self.parent.config["active_po_dir"] = new_po

        sub_dir = self.parent.target_substation_dir or new_po
        active_po_name = os.path.basename(sub_dir) if sub_dir else "default"

        standard_defaults = {
            "raw_material": os.path.join(sub_dir, "03. RAW MATERIAL") if sub_dir else "",
            "quick_report": os.path.join(sub_dir, "02. QUICK REPORT") if sub_dir else "",
            "full_report": os.path.join(sub_dir, "01. FULL REPORT") if sub_dir else "",
            "testsheet": os.path.join(sub_dir, "04. TESTSHEET") if sub_dir else "",
            "template_dir": os.path.join(sub_dir, "05. TEMPLATE") if sub_dir else "",
            "excel_mapping": os.path.join(sub_dir, "06. EXCELL MAPPING") if sub_dir else "",
            "output_file": os.path.join(sub_dir, "07. OUTPUT FILE") if sub_dir else "",
        }

        # Only store custom overrides if user entered a path that exists AND differs from the standard auto-detected default
        po_customs = {}
        for k, std_val in standard_defaults.items():
            user_val = self.entries.get(k).get().strip() if k in self.entries else ""
            if user_val and os.path.exists(user_val):
                if not std_val or os.path.normpath(user_val).lower() != os.path.normpath(std_val).lower():
                    po_customs[k] = user_val

        if "custom_paths" not in self.parent.config or not isinstance(self.parent.config["custom_paths"], dict):
            self.parent.config["custom_paths"] = {}

        # Save under workspace-specific key
        if po_customs:
            self.parent.config["custom_paths"][active_po_name] = po_customs
        else:
            if active_po_name in self.parent.config["custom_paths"]:
                del self.parent.config["custom_paths"][active_po_name]

        # Clean up any legacy flat keys if present to prevent cross-workspace locking
        for k in list(standard_defaults.keys()):
            self.parent.config["custom_paths"].pop(k, None)

        # Save Gemini API Key
        api_k = self.entries.get("gemini_api_key").get().strip() if "gemini_api_key" in self.entries else ""
        self.parent.config["gemini_api_key"] = api_k
        if api_k:
            os.environ["GEMINI_API_KEY"] = api_k
            try:
                from image_vision_brain import BRAIN
                BRAIN.api_key = api_k
            except Exception:
                pass

        # Save Executive Summary Scope Ranges
        active_po_name = os.path.basename(sub_dir) if sub_dir else "default"
        if "scope_ranges" not in self.parent.config:
            self.parent.config["scope_ranges"] = {}
        
        ir_r = self.entries.get("scope_ir_range").get().strip() if "scope_ir_range" in self.entries else "ALL"
        us_r = self.entries.get("scope_us_range").get().strip() if "scope_us_range" in self.entries else ""
        tev_r = self.entries.get("scope_tev_range").get().strip() if "scope_tev_range" in self.entries else ""

        tmpl_ir = self.entries.get("template_ir").get().strip() if "template_ir" in self.entries else ""
        tmpl_ir_ue = self.entries.get("template_ir_ue").get().strip() if "template_ir_ue" in self.entries else ""
        tmpl_ir_ue_tev = self.entries.get("template_ir_ue_tev").get().strip() if "template_ir_ue_tev" in self.entries else ""

        scope_dict = {
            "ir_range": ir_r,
            "us_range": us_r,
            "tev_range": tev_r,
            "template_ir": tmpl_ir,
            "template_ir_ue": tmpl_ir_ue,
            "template_ir_ue_tev": tmpl_ir_ue_tev
        }
        if tmpl_ir: po_customs["template_ir"] = tmpl_ir
        if tmpl_ir_ue: po_customs["template_ir_ue"] = tmpl_ir_ue
        if tmpl_ir_ue_tev: po_customs["template_ir_ue_tev"] = tmpl_ir_ue_tev
        self.parent.config["scope_ranges"][active_po_name] = scope_dict
        
        if sub_dir and os.path.exists(sub_dir):
            try:
                with open(os.path.join(sub_dir, "scope_config.json"), "w", encoding="utf-8") as f_sc:
                    json.dump(scope_dict, f_sc, indent=4)
            except Exception as e:
                print(f"Error saving scope_config.json: {e}")

        save_wraith_config(self.parent.config)

        self.parent.apply_custom_paths()
        if hasattr(self.parent, "update_po_metrics"):
            self.parent.update_po_metrics()
        self.parent.log("[SYS] Workspace, folder settings & scope ranges saved successfully!", "OK")
        if api_k:
            self.parent.log("[AI] 🤖 Google Gemini Vision API Key active! 100% Multimodal Vision enabled.", "OK")
        messagebox.showinfo("Settings Applied", "Settings updated successfully! Scope ranges and dashboard metrics refreshed.")
        self.destroy()


class ToolTip:
    """Lightweight tooltip that follows the cursor. Appears after a short hover delay."""

    def __init__(self, widget, text: str, delay_ms: int = 500):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self._tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._cancel,   add="+")
        widget.bind("<ButtonPress>", self._cancel, add="+")

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self, event=None):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        self._hide()

    def _show(self):
        if self._tip_window:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 8
        y = self.widget.winfo_rooty() + (self.widget.winfo_height() // 2) - 12
        self._tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.wm_attributes("-topmost", True)
        frame = tk.Frame(tw, bg="#1E2638", bd=1, relief="flat")
        frame.pack()
        label = tk.Label(
            frame, text=self.text,
            bg="#111520", fg="#F8FAFC",
            font=(FONT_FAMILY, 11),
            padx=10, pady=6,
            justify="left",
            wraplength=220
        )
        label.pack()

    def _hide(self):
        if self._tip_window:
            self._tip_window.destroy()
            self._tip_window = None


class BentoButton(ctk.CTkButton):
    """Action button with animated glow border on hover."""

    def __init__(self, master, title, subtitle, icon="⚡", command=None, accent_color=None):
        self._accent = accent_color if accent_color else COLOR_BORDER
        super().__init__(
            master,
            text=f"  {icon}   {title}\n       {subtitle}",
            corner_radius=10,
            height=62,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=COLOR_INNER,
            hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN,
            border_width=1,
            border_color=self._accent,
            anchor="w",
            command=command
        )
        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")

    def _on_enter(self, _=None):
        glow = self._accent if self._accent != COLOR_BORDER else COLOR_CYAN
        self.configure(border_color=glow, border_width=2)

    def _on_leave(self, _=None):
        self.configure(border_color=self._accent, border_width=1)


class ZoomFolderCard(ctk.CTkFrame):
    def __init__(self, master, folder_name, on_select_cb, on_double_click_cb):
        super().__init__(
            master,
            fg_color=COLOR_CARD,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            height=60
        )
        self.folder_name = folder_name
        self.on_select_cb = on_select_cb
        self.on_double_click_cb = on_double_click_cb
        self.is_selected = False

        self.pack_propagate(False)

        self.icon_lbl = ctk.CTkLabel(
            self, text="📁", font=ctk.CTkFont(size=20), text_color=COLOR_CYAN
        )
        self.icon_lbl.pack(side="left", padx=(18, 12))

        text_container = ctk.CTkFrame(self, fg_color="transparent")
        text_container.pack(side="left", fill="both", expand=True, pady=6)

        self.name_lbl = ctk.CTkLabel(
            text_container,
            text=folder_name,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_TEXT_MAIN,
            anchor="w"
        )
        self.name_lbl.pack(fill="x", anchor="w")

        self.sub_lbl = ctk.CTkLabel(
            text_container,
            text="Click to open workspace",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.sub_lbl.pack(fill="x", anchor="w")

        # Bind ALL descendant widgets to capture clicks inside CTkScrollableFrame
        self._bind_all_children(self)

    def _bind_all_children(self, widget):
        """Recursively bind click/hover events to all child widgets."""
        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<Button-1>", self._on_click, add="+")
        widget.bind("<Double-Button-1>", self._on_double_click, add="+")
        for child in widget.winfo_children():
            self._bind_all_children(child)

    def _on_enter(self, event=None):
        if not self.is_selected:
            self.configure(height=72, border_color=COLOR_CYAN, border_width=2, fg_color="#182338")
            self.name_lbl.configure(font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"))
            self.sub_lbl.configure(text_color=COLOR_CYAN)
            self.icon_lbl.configure(font=ctk.CTkFont(size=24))

    def _on_leave(self, event=None):
        if not self.is_selected:
            self.configure(height=60, border_color=COLOR_BORDER, border_width=1, fg_color=COLOR_CARD)
            self.name_lbl.configure(font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"))
            self.sub_lbl.configure(text_color=COLOR_TEXT_MUTED)
            self.icon_lbl.configure(font=ctk.CTkFont(size=20))

    def set_selected(self, state: bool):
        self.is_selected = state
        if state:
            self.configure(height=66, border_color=COLOR_CYAN, border_width=2, fg_color="#0284C7")
            self.name_lbl.configure(text_color="#FFFFFF", font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"))
            self.sub_lbl.configure(text="SELECTED SUBSTATION", text_color="#E0F2FE")
        else:
            self._on_leave()

    def _on_click(self, event=None):
        # Single click selects AND immediately opens the workspace
        self.on_double_click_cb(self)

    def _on_double_click(self, event=None):
        self.on_double_click_cb(self)


class ExplorerListWindow(ctk.CTkToplevel):
    def __init__(self, parent, title, header_text, items):
        super().__init__(parent)
        self.title(f"WRAITH // {title}")
        self.geometry("600x680")
        self.minsize(450, 500)
        self.resizable(True, True)
        self.configure(fg_color=COLOR_CANVAS)

        self.selected_items = []
        self.transient(parent)
        self.grab_set()

        lbl = ctk.CTkLabel(
            self, text=header_text,
            font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        )
        lbl.pack(anchor="w", padx=24, pady=(24, 4))

        sub_lbl = ctk.CTkLabel(
            self, text="Hold Ctrl or Shift to select multiple items",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=COLOR_TEXT_MUTED
        )
        sub_lbl.pack(anchor="w", padx=24, pady=(0, 12))

        list_frame = ctk.CTkFrame(
            self, fg_color=COLOR_CARD, corner_radius=12,
            border_width=1, border_color=COLOR_BORDER
        )
        list_frame.pack(fill="both", expand=True, padx=24, pady=5)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y", padx=(0, 6), pady=6)

        self.listbox = tk.Listbox(
            list_frame, selectmode=tk.EXTENDED, bg=COLOR_CARD, fg=COLOR_TEXT_MAIN,
            selectbackground=COLOR_CYAN, selectforeground="#000000",
            font=(FONT_FAMILY, 12), bd=0, highlightthickness=0,
            yscrollcommand=scrollbar.set, activestyle="none"
        )
        self.listbox.pack(side="left", fill="both", expand=True, padx=12, pady=12)
        scrollbar.config(command=self.listbox.yview)

        for item in items:
            self.listbox.insert(tk.END, item)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(12, 24))

        confirm_btn = ctk.CTkButton(
            btn_frame, text="Confirm Selection", fg_color=COLOR_CYAN,
            hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            corner_radius=8, command=self._on_confirm
        )
        confirm_btn.pack(side="right", padx=(8, 0))

        cancel_btn = ctk.CTkButton(
            btn_frame, text="Cancel", fg_color=COLOR_CARD,
            hover_color=COLOR_INNER, text_color=COLOR_TEXT_MUTED,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            border_width=1, border_color=COLOR_BORDER,
            corner_radius=8, command=self.destroy
        )
        cancel_btn.pack(side="right")

    def _on_confirm(self):
        selected_indices = self.listbox.curselection()
        self.selected_items = [self.listbox.get(i) for i in selected_indices]
        self.destroy()


class AutoRenameDialog(ctk.CTkToplevel):
    def __init__(self, parent, raw_dir, example_dir, on_confirm_cb):
        super().__init__(parent)
        self.title("WRAITH // Auto-Rename Raw Images")
        self.geometry("780x640")
        self.minsize(650, 480)
        self.resizable(True, True)
        self.configure(fg_color=COLOR_CANVAS)
        self.attributes("-topmost", True)

        self.parent = parent
        self.raw_dir = raw_dir
        self.example_dir = example_dir
        self.on_confirm_cb = on_confirm_cb
        self.selected_paths = []

        self._build_ui()

    def _build_ui(self):
        self.grab_set()

        hdr = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        hdr.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(
            hdr, text="🏷️ AUTO RENAME RAW IMAGES (CSV MAPPING)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", padx=16, pady=(12, 2))

        ctk.CTkLabel(
            hdr, text="Select single or multiple substation folders to auto-match & rename images based on CSV report pages.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=16, pady=(0, 12))

        list_frame = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        list_frame.pack(fill="both", expand=True, padx=20, pady=5)

        top_bar = ctk.CTkFrame(list_frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=16, pady=(10, 4))

        ctk.CTkLabel(
            top_bar, text="📁 Available Substation Raw Folders:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(side="left")

        sel_all_btn = ctk.CTkButton(
            top_bar, text="Select All", width=70, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_CYAN, command=self._select_all
        )
        sel_all_btn.pack(side="right", padx=(4, 0))

        clr_btn = ctk.CTkButton(
            top_bar, text="Clear", width=60, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MUTED, command=self._clear_all
        )
        clr_btn.pack(side="right")

        box_wrap = ctk.CTkFrame(list_frame, fg_color=COLOR_INNER, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        box_wrap.pack(fill="both", expand=True, padx=16, pady=(4, 12))

        scrollbar = tk.Scrollbar(box_wrap, orient="vertical")
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)

        self.listbox = tk.Listbox(
            box_wrap, selectmode=tk.EXTENDED, bg=COLOR_INNER, fg=COLOR_TEXT_MAIN,
            selectbackground=COLOR_CYAN, selectforeground="#000000",
            font=(FONT_FAMILY, 11), bd=0, highlightthickness=0,
            yscrollcommand=scrollbar.set, activestyle="none"
        )
        self.listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scrollbar.config(command=self.listbox.yview)

        self._folder_items = []
        self._populate_substations()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(10, 20))

        browse_btn = ctk.CTkButton(
            btn_frame, text="📁 Browse Custom Folder...", fg_color=COLOR_CARD,
            hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13), border_width=1,
            border_color=COLOR_BORDER, corner_radius=8, height=36,
            command=self._browse_custom_folder
        )
        browse_btn.pack(side="left")

        confirm_btn = ctk.CTkButton(
            btn_frame, text="▶ Run Auto Rename", fg_color=COLOR_CYAN,
            hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            corner_radius=8, height=36, command=self._on_confirm
        )
        confirm_btn.pack(side="right", padx=(8, 0))

        cancel_btn = ctk.CTkButton(
            btn_frame, text="Cancel", fg_color=COLOR_CARD,
            hover_color=COLOR_INNER, text_color=COLOR_TEXT_MUTED,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            border_width=1, border_color=COLOR_BORDER,
            corner_radius=8, height=36, command=self.destroy
        )
        cancel_btn.pack(side="right")

    def _populate_substations(self):
        self.listbox.delete(0, tk.END)
        self._folder_items.clear()

        if self.raw_dir and os.path.exists(self.raw_dir):
            for f in sorted(os.listdir(self.raw_dir)):
                full_p = os.path.join(self.raw_dir, f)
                if os.path.isdir(full_p) and f.upper() not in ["DG", "IR", "PDF", "NOT DONE", "PROCESSED_EXCEL"]:
                    if re.match(r'^\d+', f):
                        self.listbox.insert(tk.END, f" [03. RAW]   {f}")
                        self._folder_items.append(full_p)

        if self.example_dir and os.path.exists(self.example_dir):
            for f in sorted(os.listdir(self.example_dir)):
                full_p = os.path.join(self.example_dir, f)
                if os.path.isdir(full_p):
                    self.listbox.insert(tk.END, f" [EXAMPLE]   {f}")
                    self._folder_items.append(full_p)

    def _select_all(self):
        self.listbox.select_set(0, tk.END)

    def _clear_all(self):
        self.listbox.selection_clear(0, tk.END)

    def _browse_custom_folder(self):
        chosen = filedialog.askdirectory(title="Select Substation Raw Images Folder")
        if chosen:
            self.selected_paths = [chosen]
            self.on_confirm_cb(self.selected_paths)
            self.destroy()

    def _on_confirm(self):
        sel_indices = self.listbox.curselection()
        if not sel_indices:
            messagebox.showwarning("Selection Required", "Please select at least one substation folder from the list or click 'Browse Custom Folder'.")
            return

        self.selected_paths = [self._folder_items[i] for i in sel_indices]
        self.on_confirm_cb(self.selected_paths)
        self.destroy()


class ConvertWordToPdfDialog(ctk.CTkToplevel):
    def __init__(self, parent, default_dir, on_convert_cb):
        super().__init__(parent)
        self.title("WRAITH // Convert Word to PDF (Lossless Mode)")
        self.geometry("840x690")
        self.minsize(700, 520)
        self.resizable(True, True)
        self.configure(fg_color=COLOR_CANVAS)
        self.attributes("-topmost", True)

        self.parent = parent
        self.current_dir = default_dir if default_dir and os.path.exists(default_dir) else (getattr(parent, "target_substation_dir", "") or DEFAULT_ROOT_DIR)
        self.on_convert_cb = on_convert_cb

        self._all_docx_files = []
        self._filtered_files = []

        self._build_ui()

    def _build_ui(self):
        self.grab_set()

        hdr = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        hdr.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(
            hdr, text="📄 CONVERT WORD TO PDF (LOSSLESS MODE)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", padx=16, pady=(12, 2))

        ctk.CTkLabel(
            hdr, text="Select single or multiple Word (.docx) files from your dedicated workspace to convert to high-resolution PDF without quality loss.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=16, pady=(0, 12))

        # Folder Selection Bar
        folder_bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        folder_bar.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(
            folder_bar, text="📁 Dedicated Workspace Folder:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(side="left", padx=(14, 8), pady=10)

        self.path_entry = ctk.CTkEntry(
            folder_bar, fg_color=COLOR_INNER, border_color=COLOR_BORDER,
            text_color=COLOR_CYAN, font=ctk.CTkFont(family=FONT_MONO, size=11)
        )
        self.path_entry.insert(0, self.current_dir)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=10)

        browse_btn = ctk.CTkButton(
            folder_bar, text="📁 Browse Folder...", width=130, height=28,
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MAIN, command=self._browse_folder
        )
        browse_btn.pack(side="right", padx=(0, 14), pady=10)

        # File List Section
        list_frame = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        list_frame.pack(fill="both", expand=True, padx=20, pady=0)

        top_bar = ctk.CTkFrame(list_frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=16, pady=(10, 4))

        self.count_lbl = ctk.CTkLabel(
            top_bar, text="Word Documents:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        )
        self.count_lbl.pack(side="left")

        # Quick actions
        sel_all_btn = ctk.CTkButton(
            top_bar, text="Select All", width=75, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_CYAN, command=self._select_all
        )
        sel_all_btn.pack(side="right", padx=(4, 0))

        clr_btn = ctk.CTkButton(
            top_bar, text="Clear", width=60, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MUTED, command=self._clear_all
        )
        clr_btn.pack(side="right", padx=(4, 0))

        # Search box
        self.search_entry = ctk.CTkEntry(
            top_bar, placeholder_text="🔍 Filter files...", width=160, height=24,
            fg_color=COLOR_INNER, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MAIN, font=ctk.CTkFont(family=FONT_FAMILY, size=11)
        )
        self.search_entry.pack(side="right", padx=(0, 8))
        self.search_entry.bind("<KeyRelease>", lambda e: self._filter_files())

        # Listbox container
        box_wrap = ctk.CTkFrame(list_frame, fg_color=COLOR_INNER, corner_radius=8, border_width=1, border_color=COLOR_BORDER)
        box_wrap.pack(fill="both", expand=True, padx=16, pady=(4, 10))

        scrollbar = tk.Scrollbar(box_wrap, orient="vertical")
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)

        self.listbox = tk.Listbox(
            box_wrap, selectmode=tk.EXTENDED, bg=COLOR_INNER, fg=COLOR_TEXT_MAIN,
            selectbackground=COLOR_CYAN, selectforeground="#000000",
            font=(FONT_FAMILY, 11), bd=0, highlightthickness=0,
            yscrollcommand=scrollbar.set, activestyle="none"
        )
        self.listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scrollbar.config(command=self.listbox.yview)

        # Output options bar
        dest_bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        dest_bar.pack(fill="x", padx=20, pady=(10, 0))

        ctk.CTkLabel(
            dest_bar, text="💾 Save PDF To:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(side="left", padx=(14, 8), pady=8)

        self.dest_var = tk.StringVar(value="same")
        ctk.CTkRadioButton(
            dest_bar, text="Same Folder as Word Files", variable=self.dest_var, value="same",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=COLOR_TEXT_MAIN,
            fg_color=COLOR_CYAN, radiobutton_width=14, radiobutton_height=14
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            dest_bar, text="01. FULL REPORT", variable=self.dest_var, value="full_report",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=COLOR_TEXT_MAIN,
            fg_color=COLOR_CYAN, radiobutton_width=14, radiobutton_height=14
        ).pack(side="left", padx=(0, 16))

        # Testsheet Append Options Bar
        ts_bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        ts_bar.pack(fill="x", padx=20, pady=(8, 0))

        self.append_ts_var = tk.BooleanVar(value="QUICK REPORT" in self.current_dir.upper())
        self.ts_chk = ctk.CTkCheckBox(
            ts_bar, text="📎 Append Matching Testsheet PDF (04. TESTSHEET/PDF) at the back of each converted PDF",
            variable=self.append_ts_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_CYAN, fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER
        )
        self.ts_chk.pack(side="left", padx=(14, 8), pady=8)

        # Bottom Actions
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(10, 20))

        badge_lbl = ctk.CTkLabel(
            btn_frame, text="💎 Lossless 300+ DPI (DoNotCompressPictures = True)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_CYAN
        )
        badge_lbl.pack(side="left")

        confirm_btn = ctk.CTkButton(
            btn_frame, text="▶ Convert Selected to PDF", fg_color=COLOR_CYAN,
            hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            corner_radius=8, height=36, command=self._on_confirm
        )
        confirm_btn.pack(side="right", padx=(8, 0))

        cancel_btn = ctk.CTkButton(
            btn_frame, text="Cancel", fg_color=COLOR_CARD,
            hover_color=COLOR_INNER, text_color=COLOR_TEXT_MUTED,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            border_width=1, border_color=COLOR_BORDER,
            corner_radius=8, height=36, command=self.destroy
        )
        cancel_btn.pack(side="right")

        self._refresh_files()

    def _browse_folder(self):
        chosen = filedialog.askdirectory(initialdir=self.current_dir, title="Select Workspace Folder with Word Files")
        if chosen:
            self.current_dir = chosen
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, chosen)
            if "QUICK REPORT" in chosen.upper():
                self.append_ts_var.set(True)
            self._refresh_files()

    def _refresh_files(self):
        self._all_docx_files.clear()
        if self.current_dir and os.path.exists(self.current_dir):
            for root, _, files in os.walk(self.current_dir):
                for f in sorted(files):
                    if f.lower().endswith(".docx") and not f.startswith("~$"):
                        full_p = os.path.join(root, f)
                        self._all_docx_files.append(full_p)
        self._filter_files()

    def _filter_files(self):
        query = self.search_entry.get().strip().lower()
        self.listbox.delete(0, tk.END)
        self._filtered_files.clear()

        for fp in self._all_docx_files:
            fname = os.path.basename(fp)
            if not query or query in fname.lower():
                sz = round(os.path.getsize(fp) / 1024, 1)
                self.listbox.insert(tk.END, f" 📄 {fname:<55} ({sz} KB)")
                self._filtered_files.append(fp)

        self.count_lbl.configure(text=f"Word Documents ({len(self._filtered_files)} found):")
        # Default select all
        self.listbox.select_set(0, tk.END)

    def _select_all(self):
        self.listbox.select_set(0, tk.END)

    def _clear_all(self):
        self.listbox.selection_clear(0, tk.END)

    def _on_confirm(self):
        sel_indices = self.listbox.curselection()
        if not sel_indices:
            messagebox.showwarning("Selection Required", "Please select at least one Word document (.docx) to convert.")
            return

        selected_files = [self._filtered_files[i] for i in sel_indices]
        dest_choice = self.dest_var.get()
        dest_folder = None
        if dest_choice == "full_report":
            if getattr(self.parent, "target_substation_dir", None):
                dest_folder = os.path.join(self.parent.target_substation_dir, "01. FULL REPORT")
            else:
                dest_folder = os.path.join(self.current_dir, "01. FULL REPORT")

        append_ts = self.append_ts_var.get()
        ts_dir = None
        if append_ts:
            if getattr(self.parent, "target_substation_dir", None):
                ts_dir = os.path.join(self.parent.target_substation_dir, "04. TESTSHEET", "PDF")
                if not os.path.exists(ts_dir):
                    ts_dir = os.path.join(self.parent.target_substation_dir, "04. TESTSHEET")
            else:
                ts_dir = os.path.join(os.path.dirname(self.current_dir), "04. TESTSHEET", "PDF")

        self.on_convert_cb(selected_files, dest_folder, append_ts, ts_dir)
        self.destroy()



class FolderSelectionScreen(ctk.CTkFrame):
    def __init__(self, parent, on_select_callback):
        super().__init__(parent, fg_color=COLOR_CANVAS)
        self.parent = parent
        self.on_select_callback = on_select_callback
        self.selected_card = None
        self.cards = []

        card = ctk.CTkFrame(
            self, fg_color=COLOR_CARD, corner_radius=16,
            border_width=1, border_color=COLOR_BORDER
        )
        card.pack(padx=60, pady=40, fill="both", expand=True)

        title = ctk.CTkLabel(
            card, text="Select Substation Work Folder",
            font=ctk.CTkFont(family=FONT_FAMILY, size=28, weight="bold"), text_color=COLOR_TEXT_MAIN
        )
        title.pack(anchor="w", padx=36, pady=(32, 4))

        self.subtitle = ctk.CTkLabel(
            card, text=f"Main Root: {DEFAULT_ROOT_DIR}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=COLOR_TEXT_MUTED
        )
        self.subtitle.pack(anchor="w", padx=36, pady=(0, 16))

        self.scroll_container = ctk.CTkScrollableFrame(
            card, fg_color=COLOR_INNER, corner_radius=12,
            border_width=1, border_color=COLOR_BORDER
        )
        self.scroll_container.pack(fill="both", expand=True, padx=36, pady=10)

        self.populate_folders()

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=36, pady=(16, 32))

        browse_btn = ctk.CTkButton(
            btn_frame, text="📁 Browse Folder...", fg_color=COLOR_CARD,
            hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14), border_width=1,
            border_color=COLOR_BORDER, corner_radius=8, command=self._browse_folder
        )
        browse_btn.pack(side="left")

        root_btn = ctk.CTkButton(
            btn_frame, text="⚙️ Change Root Workspace...", fg_color=COLOR_CARD,
            hover_color=COLOR_CARD_HOVER, text_color=COLOR_CYAN,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14), border_width=1,
            border_color=COLOR_BORDER, corner_radius=8, command=self._change_workspace_root
        )
        root_btn.pack(side="left", padx=(10, 0))

        proceed_btn = ctk.CTkButton(
            btn_frame, text="Launch CBM Workspace  ➔", fg_color=COLOR_CYAN,
            hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            corner_radius=8, command=self._confirm_selection
        )
        proceed_btn.pack(side="right")

    def populate_folders(self):
        for widget in self.scroll_container.winfo_children():
            widget.destroy()

        self.cards.clear()
        self.selected_card = None

        if os.path.exists(DEFAULT_ROOT_DIR):
            subfolders = [
                f for f in os.listdir(DEFAULT_ROOT_DIR)
                if os.path.isdir(os.path.join(DEFAULT_ROOT_DIR, f)) and f.upper() != "SCRIPT"
            ]
            for folder in subfolders:
                card_item = ZoomFolderCard(
                    self.scroll_container, folder,
                    self._on_card_selected, self._on_card_double_clicked
                )
                card_item.pack(fill="x", padx=10, pady=6)
                self.cards.append(card_item)

    def _on_card_selected(self, card_item):
        if self.selected_card:
            self.selected_card.set_selected(False)
        self.selected_card = card_item
        self.selected_card.set_selected(True)

    def _on_card_double_clicked(self, card_item):
        self._on_card_selected(card_item)
        self._confirm_selection()

    def _confirm_selection(self):
        if self.selected_card:
            full_path = os.path.join(DEFAULT_ROOT_DIR, self.selected_card.folder_name)
            self.on_select_callback(full_path)
        else:
            messagebox.showwarning("Selection Required", "Please click to select a folder first.")

    def _browse_folder(self):
        selected = filedialog.askdirectory(initialdir=DEFAULT_ROOT_DIR, title="Select Substation Work Folder")
        if selected:
            self.on_select_callback(selected)

    def _change_workspace_root(self):
        global DEFAULT_ROOT_DIR
        selected = filedialog.askdirectory(initialdir=DEFAULT_ROOT_DIR, title="Select Main REPORT CBM Workspace Root")
        if selected:
            DEFAULT_ROOT_DIR = selected
            if hasattr(self, 'parent') and hasattr(self.parent, 'config') and isinstance(self.parent.config, dict):
                self.parent.config["root_workspace_dir"] = selected
                save_wraith_config(self.parent.config)
            if hasattr(self, 'subtitle'):
                self.subtitle.configure(text=f"Main Root: {DEFAULT_ROOT_DIR}")
            self.populate_folders()

class ReportManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("WRAITH // CBM Automation Hub v2026")
        self.geometry("1550x980")
        self.minsize(900, 650)
        self.resizable(True, True)
        self.configure(fg_color=COLOR_CANVAS)

        self.config = load_wraith_config()
        self.active_processes = []   # list[subprocess.Popen]
        self.target_substation_dir = self.config.get("active_po_dir") or None
        self.gen_report_dir = self.target_substation_dir
        self.mapping_dir = None
        self.script_dir = CENTRAL_SCRIPT_DIR
        self.output_dir = None
        self.rm_dir = None
        self.qr_dir = None
        self.fr_dir = None
        self.ts_dir = None
        self.insert_sound_dir = None
        self.master_list_path = None

        # Pause / resume session state
        self._is_paused = False          # True while processes are suspended
        self._session_file = None        # absolute path to .rms_session.json
        self._resume_poll_id = None      # after() id for resume polling loop

        self.selection_screen = FolderSelectionScreen(self, self.load_substation_dashboard)
        self.selection_screen.pack(fill="both", expand=True)

        # Graceful close — suspend instead of kill when there are active processes
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def apply_custom_paths(self):
        sub = self.target_substation_dir
        if not sub:
            return

        active_po_name = os.path.basename(sub)
        customs = self.config.get("custom_paths", {})

        # Support both workspace-specific dict (preferred) and legacy flat dict
        po_customs = {}
        if isinstance(customs, dict):
            if active_po_name in customs and isinstance(customs[active_po_name], dict):
                po_customs = customs[active_po_name]
            elif any(k in customs for k in ["raw_material", "quick_report", "full_report", "testsheet", "excel_mapping", "output_file"]):
                po_customs = customs

        root_dir = self.config.get("root_workspace_dir", DEFAULT_ROOT_DIR)
        root_norm = os.path.normcase(os.path.abspath(root_dir)) if root_dir else ""
        sub_norm = os.path.normcase(os.path.abspath(sub))

        def _resolve(key, standard_subfolder):
            override = po_customs.get(key, "")
            if isinstance(override, str):
                override = override.strip()
            else:
                override = ""

            if override and os.path.exists(override):
                ov_norm = os.path.normcase(os.path.abspath(override))
                # Check for stale cross-PO path:
                # If the override is located inside root_workspace_dir, but NOT inside this current active sub,
                # then it belongs to a different PO workspace -> ignore and auto-detect!
                if root_norm and ov_norm.startswith(root_norm) and not ov_norm.startswith(sub_norm):
                    return os.path.join(sub, standard_subfolder)
                return override
            return os.path.join(sub, standard_subfolder)

        self.rm_dir = _resolve("raw_material", "03. RAW MATERIAL")
        self.qr_dir = _resolve("quick_report", "02. QUICK REPORT")
        self.fr_dir = _resolve("full_report", "01. FULL REPORT")
        self.ts_dir = _resolve("testsheet", "04. TESTSHEET")
        self.mapping_dir = _resolve("excel_mapping", "06. EXCELL MAPPING")
        self.output_dir = _resolve("output_file", "07. OUTPUT FILE")
        self.gen_report_dir = sub
        self.master_list_path = os.path.join(self.mapping_dir, "Master List.xlsx")
        self.insert_sound_dir = os.path.join(sub, "08. INSERT SOUND")

        # Dynamic template folder detection: check custom override first, then substation's own '05. TEMPLATE', fallback to Root 'TEMPLATE'
        custom_tmpl = po_customs.get("template_dir")
        if custom_tmpl and os.path.exists(custom_tmpl):
            self.template_dir = custom_tmpl
        else:
            local_tmpl = os.path.join(sub, "05. TEMPLATE") if sub else ""
            if local_tmpl and os.path.exists(local_tmpl):
                self.template_dir = local_tmpl
            else:
                root_tmpl = os.path.join(root_dir, "TEMPLATE") if root_dir else ""
                self.template_dir = root_tmpl if (root_tmpl and os.path.exists(root_tmpl)) else local_tmpl

        det_ir, det_ir_ue, det_ir_ue_tev = auto_detect_scope_templates(sub, root_dir)
        self.template_ir_dir = po_customs.get("template_ir") or det_ir
        self.template_ir_ue_dir = po_customs.get("template_ir_ue") or det_ir_ue
        self.template_ir_ue_tev_dir = po_customs.get("template_ir_ue_tev") or det_ir_ue_tev

    def open_settings_dialog(self):
        SettingsDialog(self)

    def load_substation_dashboard(self, substation_path):
        self.target_substation_dir = substation_path
        self.config["active_po_dir"] = substation_path
        save_wraith_config(self.config)
        self.apply_custom_paths()
        self._session_file = os.path.join(substation_path, SESSION_FILENAME)

        self.selection_screen.pack_forget()

        self.grid_columnconfigure(0, weight=0, minsize=72)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0, minsize=380)
        self.grid_rowconfigure(0, weight=1)

        self._create_left_dock()
        self._create_center_dashboard()
        self._create_right_panel()

        # After UI is built, check for a frozen session from a previous run and start live metrics loop
        self.after(300, self._check_for_saved_session)
        self._start_live_metrics_loop()

    def _start_live_metrics_loop(self):
        if hasattr(self, "_live_metrics_timer_id") and self._live_metrics_timer_id:
            try:
                self.after_cancel(self._live_metrics_timer_id)
            except Exception:
                pass
        self._live_update_metrics()

    def _live_update_metrics(self):
        try:
            if hasattr(self, "target_substation_dir") and self.target_substation_dir and hasattr(self, "lbl_ts"):
                self.update_po_metrics(log_msg=False)
        except Exception:
            pass
        finally:
            self._live_metrics_timer_id = self.after(3000, self._live_update_metrics)

    def _create_left_dock(self):
        dock = ctk.CTkFrame(self, fg_color=COLOR_RAIL, corner_radius=0, border_width=1, border_color=COLOR_BORDER)
        dock.grid(row=0, column=0, sticky="nsew")
        dock.pack_propagate(False)

        # ── Logo Mark ──────────────────────────────────────────────────────────
        logo_frame = ctk.CTkFrame(dock, fg_color=COLOR_CARD, corner_radius=10, width=54, height=54)
        logo_frame.pack(pady=(20, 8))
        logo_frame.pack_propagate(False)
        ctk.CTkLabel(
            logo_frame, text="⬡",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=COLOR_CYAN
        ).place(relx=0.5, rely=0.5, anchor="center")

        # ── Separator ─────────────────────────────────────────────────────────
        sep = ctk.CTkFrame(dock, fg_color=COLOR_BORDER, height=1, width=42)
        sep.pack(pady=(4, 14))

        # ── Nav Tab Entries ───────────────────────────────────────────────────
        nav_items = [
            ("◈", "Dashboard",   None,                       True),
            ("📋", "CSV Mapping", self.open_csv_mapping_folder, False),
            ("📑", "Templates",   self.open_template_folder,   False),
            ("📁", "Output Dir",  self.open_output_folder,     False),
            ("⚙️", "Settings",   self.open_settings_dialog,   False),
        ]
        self._dock_btns = []
        for icon, label, cmd, active in nav_items:
            # Fixed-size centering wrapper so all buttons align identically
            wrapper = ctk.CTkFrame(dock, fg_color="transparent", width=58, height=64)
            wrapper.pack(pady=2)
            wrapper.pack_propagate(False)

            btn_frame = ctk.CTkFrame(
                wrapper, fg_color=COLOR_CARD if active else "transparent",
                corner_radius=10, width=54, height=58
            )
            btn_frame.place(relx=0.5, rely=0.5, anchor="center")
            btn_frame.pack_propagate(False)

            # Active indicator: bottom cyan bar
            if active:
                accent_bar = ctk.CTkFrame(btn_frame, fg_color=COLOR_CYAN, corner_radius=2, width=28, height=3)
                accent_bar.place(relx=0.5, rely=1.0, anchor="s", y=-3)

            icon_lbl = ctk.CTkLabel(
                btn_frame, text=icon,
                font=ctk.CTkFont(size=20),
                text_color=COLOR_CYAN if active else COLOR_TEXT_MUTED
            )
            icon_lbl.place(relx=0.5, rely=0.38, anchor="center")

            sub_lbl = ctk.CTkLabel(
                btn_frame, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=8),
                text_color=COLOR_CYAN if active else COLOR_TEXT_DIM
            )
            sub_lbl.place(relx=0.5, rely=0.78, anchor="center")

            if cmd:
                for w in (wrapper, btn_frame, icon_lbl, sub_lbl):
                    w.bind("<Button-1>", lambda e, c=cmd: c())
                    w.bind("<Enter>", lambda e, f=btn_frame: f.configure(fg_color=COLOR_CARD_HOVER), add="+")
                    w.bind("<Leave>", lambda e, f=btn_frame: f.configure(fg_color="transparent"), add="+")

            ToolTip(btn_frame, text=label, delay_ms=400)

        # ── Bottom Actions ─────────────────────────────────────────────────────
        bottom_sep = ctk.CTkFrame(dock, fg_color=COLOR_BORDER, height=1, width=42)
        bottom_sep.pack(side="bottom", pady=(0, 8))

        # Kill button
        kill_frame = ctk.CTkFrame(dock, fg_color=COLOR_INNER, corner_radius=10, width=58, height=58)
        kill_frame.pack(side="bottom", pady=4)
        kill_frame.pack_propagate(False)
        kill_icon = ctk.CTkLabel(kill_frame, text="🛑", font=ctk.CTkFont(size=20), text_color=COLOR_ROSE)
        kill_icon.place(relx=0.5, rely=0.38, anchor="center")
        kill_sub = ctk.CTkLabel(kill_frame, text="Kill All", font=ctk.CTkFont(family=FONT_FAMILY, size=8), text_color=COLOR_ROSE)
        kill_sub.place(relx=0.5, rely=0.78, anchor="center")
        for w in (kill_frame, kill_icon, kill_sub):
            w.bind("<Button-1>", lambda e: self.terminate_all_processes())
            w.bind("<Enter>", lambda e: kill_frame.configure(fg_color="#3B0A14"), add="+")
            w.bind("<Leave>", lambda e: kill_frame.configure(fg_color=COLOR_INNER), add="+")
        ToolTip(kill_frame, text="Terminate All Processes", delay_ms=400)

        # Switch Folder button
        switch_frame = ctk.CTkFrame(dock, fg_color="transparent", corner_radius=10, width=58, height=58)
        switch_frame.pack(side="bottom", pady=4)
        switch_frame.pack_propagate(False)
        sw_icon = ctk.CTkLabel(switch_frame, text="🔄", font=ctk.CTkFont(size=20), text_color=COLOR_TEXT_MUTED)
        sw_icon.place(relx=0.5, rely=0.38, anchor="center")
        sw_sub = ctk.CTkLabel(switch_frame, text="Switch", font=ctk.CTkFont(family=FONT_FAMILY, size=8), text_color=COLOR_TEXT_DIM)
        sw_sub.place(relx=0.5, rely=0.78, anchor="center")
        for w in (switch_frame, sw_icon, sw_sub):
            w.bind("<Button-1>", lambda e: self._switch_folder())
            w.bind("<Enter>", lambda e: switch_frame.configure(fg_color=COLOR_CARD), add="+")
            w.bind("<Leave>", lambda e: switch_frame.configure(fg_color="transparent"), add="+")
        ToolTip(switch_frame, text="Switch Substation Folder", delay_ms=400)

    def _create_center_dashboard(self):
        center = ctk.CTkFrame(self, fg_color="transparent")
        center.grid(row=0, column=1, sticky="nsew", padx=24, pady=24)

        folder_name = os.path.basename(self.target_substation_dir)
        header_card = ctk.CTkFrame(center, fg_color=COLOR_CARD, corner_radius=12, border_width=1,
                                   border_color=COLOR_BORDER)
        header_card.pack(fill="x", pady=(0, 16))

        title_lbl = ctk.CTkLabel(
            header_card, text=f"Substation Workspace  •  {folder_name}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=22, weight="bold"), text_color=COLOR_TEXT_MAIN
        )
        title_lbl.pack(anchor="w", padx=24, pady=(16, 2))

        sub_path_lbl = ctk.CTkLabel(
            header_card, text=f"Path: {self.target_substation_dir}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=COLOR_TEXT_MUTED
        )
        sub_path_lbl.pack(anchor="w", padx=24, pady=(0, 16))

        # ── PO STATUS HUD LIVE CARD ──────────────────────────────────────────
        status_card = ctk.CTkFrame(center, fg_color=COLOR_CARD, corner_radius=12, border_width=1,
                                   border_color=COLOR_BORDER)
        status_card.pack(fill="x", pady=(0, 16))

        status_top = ctk.CTkFrame(status_card, fg_color="transparent")
        status_top.pack(fill="x", padx=20, pady=(12, 4))
        ctk.CTkLabel(
            status_top, text="📊 LIVE PO STATUS & METRICS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"), text_color=COLOR_CYAN
        ).pack(side="left")

        refresh_metrics_btn = ctk.CTkButton(
            status_top, text="🔄 Refresh", width=68, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), fg_color=COLOR_INNER,
            text_color=COLOR_TEXT_MUTED, hover_color=COLOR_BORDER, command=self.update_po_metrics
        )
        refresh_metrics_btn.pack(side="right")

        metrics_grid = ctk.CTkFrame(status_card, fg_color="transparent")
        metrics_grid.pack(fill="x", padx=16, pady=(4, 16))
        for col_idx in range(5):
            metrics_grid.columnconfigure(col_idx, weight=1)

        ts_count, qr_count, rm_count, fr_done_count, not_done_count = self.get_po_counts()

        self.lbl_ts = self._create_metric_badge(metrics_grid, 0, "TEST SHEETS", str(ts_count), COLOR_CYAN,
                                                command=lambda: self._open_subfolder("04. TESTSHEET"),
                                                tooltip="Click to open 04. TESTSHEET folder")
        self.lbl_qr = self._create_metric_badge(metrics_grid, 1, "QUICK REPORTS", str(qr_count), COLOR_INDIGO,
                                                command=lambda: self._open_subfolder("02. QUICK REPORT"),
                                                tooltip="Click to open 02. QUICK REPORT folder")
        self.lbl_rm = self._create_metric_badge(metrics_grid, 2, "RAW MATERIAL", str(rm_count), COLOR_VIOLET,
                                                command=lambda: self._open_subfolder("03. RAW MATERIAL"),
                                                tooltip="Click to open 03. RAW MATERIAL folder")
        self.lbl_fr = self._create_metric_badge(metrics_grid, 3, "FULL REPORT DONE", str(fr_done_count), COLOR_EMERALD,
                                                command=lambda: self._open_subfolder("01. FULL REPORT"),
                                                tooltip="Click to open 01. FULL REPORT folder")
        self.lbl_nd = self._create_metric_badge(metrics_grid, 4, "PENDING REPORTS", str(not_done_count), COLOR_AMBER,
                                                command=lambda: self._open_subfolder(r"01. FULL REPORT\NOT DONE"),
                                                tooltip="Click to open 01. FULL REPORT\\NOT DONE folder")

        # ── CONSOLE HUD ──────────────────────────────────────────────────────
        console_card = ctk.CTkFrame(center, fg_color=COLOR_CARD, corner_radius=14, border_width=1,
                                    border_color=COLOR_BORDER)
        console_card.pack(fill="both", expand=True, padx=0, pady=(2, 0))

        console_top = ctk.CTkFrame(console_card, fg_color="transparent")
        console_top.pack(fill="x", padx=18, pady=(12, 4))

        ctk.CTkLabel(
            console_top, text="◈  SYSTEM HUD CONSOLE LOG",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"), text_color=COLOR_TEXT_MUTED
        ).pack(side="left")

        btn_row = ctk.CTkFrame(console_top, fg_color="transparent")
        btn_row.pack(side="right")

        # Legend chips
        for tag_name, col in [("INFO", COLOR_CYAN), ("OK", COLOR_EMERALD), ("WARN", COLOR_AMBER), ("ERR", COLOR_ROSE)]:
            ctk.CTkLabel(
                btn_row, text=f" {tag_name} ",
                font=ctk.CTkFont(family=FONT_MONO, size=9, weight="bold"),
                fg_color=COLOR_INNER, text_color=col,
                corner_radius=4, width=32, height=16
            ).pack(side="left", padx=2)

        clear_btn = ctk.CTkButton(
            btn_row, text="⌫ Clear", width=64, height=24,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), fg_color=COLOR_INNER,
            text_color=COLOR_TEXT_MUTED, hover_color=COLOR_BORDER, command=self._clear_console
        )
        clear_btn.pack(side="left", padx=(6, 0))

        # Console textbox with colour tags
        self.console = ctk.CTkTextbox(
            console_card, font=ctk.CTkFont(family=FONT_MONO, size=12),
            fg_color=COLOR_INNER, text_color=COLOR_TEXT_MUTED,
            border_width=1, border_color=COLOR_BORDER, wrap="word"
        )
        self.console.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        # Register text colour tags on the underlying tk.Text widget
        tw = self.console._textbox
        for tag_key, col in LOG_COLORS.items():
            tw.tag_configure(tag_key, foreground=col)
        tw.tag_configure("TS",   foreground=COLOR_TEXT_DIM)
        tw.tag_configure("BOLD", font=(FONT_MONO, 12, "bold"))

    def _create_metric_badge(self, parent, col, title, val_str, accent_col, command=None, tooltip=None):
        frame = ctk.CTkFrame(parent, fg_color=COLOR_INNER, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        frame.grid(row=0, column=col, padx=6, sticky="nsew")

        lbl_title = ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
                                 text_color=COLOR_TEXT_MUTED)
        lbl_title.pack(anchor="w", padx=12, pady=(10, 2))
        val_lbl = ctk.CTkLabel(frame, text=val_str, font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
                               text_color=accent_col)
        val_lbl.pack(anchor="w", padx=12, pady=(0, 10))

        if command:
            for w in (frame, lbl_title, val_lbl):
                w.bind("<Button-1>", lambda e: command())
                w.bind("<Enter>", lambda e: frame.configure(border_color=COLOR_CYAN, border_width=2, fg_color=COLOR_CARD_HOVER), add="+")
                w.bind("<Leave>", lambda e: frame.configure(border_color=COLOR_BORDER, border_width=1, fg_color=COLOR_INNER), add="+")
            if tooltip:
                ToolTip(frame, text=tooltip, delay_ms=400)

        return val_lbl

    def _open_subfolder(self, sub_rel_path: str):
        if not self.target_substation_dir:
            return
        target_path = os.path.join(self.target_substation_dir, sub_rel_path)
        if os.path.exists(target_path):
            os.startfile(target_path)
            self.log(f"📁 Opened folder: {target_path}", level="SYS")
        else:
            messagebox.showwarning("Folder Missing", f"Folder does not exist yet:\n{target_path}")

    def get_po_counts(self):
        if not self.target_substation_dir or not os.path.exists(self.target_substation_dir):
            return 0, 0, 0, 0, 0

        self.apply_custom_paths()
        ts_dir = self.ts_dir or os.path.join(self.target_substation_dir, "04. TESTSHEET")
        ts_count = 0
        if os.path.exists(ts_dir):
            ts_subs = set()
            for f in os.listdir(ts_dir):
                if f.startswith("~") or not f.lower().endswith(".xlsx"):
                    continue
                m = re.match(r'^(\d+)', f)
                if m:
                    ts_subs.add(int(m.group(1)))
            ts_count = len(ts_subs)

        qr_dir = self.qr_dir or os.path.join(self.target_substation_dir, "02. QUICK REPORT")
        qr_count = 0
        if os.path.exists(qr_dir):
            qr_subs = set()
            for f in os.listdir(qr_dir):
                if f.startswith("~$"):
                    continue
                m = re.match(r'^(\d+[A-Za-z]?)', f)
                if m:
                    qr_subs.add(m.group(1))
            qr_count = len(qr_subs)

        rm_dir = self.rm_dir or os.path.join(self.target_substation_dir, "03. RAW MATERIAL")
        rm_count = 0
        if os.path.exists(rm_dir):
            rm_subs = set()
            for f in os.listdir(rm_dir):
                p = os.path.join(rm_dir, f)
                if os.path.isdir(p):
                    if f.upper() in ["DG", "IR", "PDF", "NOT DONE", "PROCESSED_EXCEL"]:
                        continue
                    m = re.match(r'^(\d+)', f)
                    if m:
                        rm_subs.add(int(m.group(1)))
            rm_count = len(rm_subs)

        fr_dir = self.fr_dir or os.path.join(self.target_substation_dir, "01. FULL REPORT")
        done_substations = set()

        if os.path.exists(fr_dir):
            for root, dirs, files in os.walk(fr_dir):
                rel = os.path.relpath(root, fr_dir)
                if "NOT DONE" in rel.upper():
                    continue

                prefix_files = {}
                for f in files:
                    if f.startswith("~$"):
                        continue
                    m = re.match(r'^0*(\d+)', f)
                    if m:
                        num = int(m.group(1))
                        prefix_files.setdefault(num, []).append(f)

                for num, f_list in prefix_files.items():
                    has_g0_docx = any("group 0.docx" in f.lower() for f in f_list)
                    has_g0_pdf  = any("group 0.pdf" in f.lower() for f in f_list)
                    has_g1_docx = any("group 1.docx" in f.lower() for f in f_list)
                    has_g1_pdf  = any("group 1.pdf" in f.lower() for f in f_list)
                    has_main_pdf = any(f.lower().endswith(".pdf") and "group 0" not in f.lower() and "group 1" not in f.lower() for f in f_list)
                    if (has_g0_docx and has_g0_pdf and has_g1_docx and has_g1_pdf and has_main_pdf) or has_main_pdf:
                        done_substations.add(num)

        fr_done_count = len(done_substations)

        not_done_dir = os.path.join(fr_dir, "NOT DONE")
        not_done_count = 0
        if os.path.exists(not_done_dir):
            sub_set = set()
            for f in os.listdir(not_done_dir):
                if f.startswith("~$"):
                    continue
                p = os.path.join(not_done_dir, f)
                m = re.match(r'^(\d+[A-Za-z]?)', f)
                if m:
                    sub_set.add(m.group(1))
            not_done_count = len(sub_set)

        return ts_count, qr_count, rm_count, fr_done_count, not_done_count

    def update_po_metrics(self, log_msg=True):
        ts, qr, rm, fr_done, not_done = self.get_po_counts()
        if hasattr(self, "lbl_ts"): self.lbl_ts.configure(text=str(ts))
        if hasattr(self, "lbl_qr"): self.lbl_qr.configure(text=str(qr))
        if hasattr(self, "lbl_rm"): self.lbl_rm.configure(text=str(rm))
        if hasattr(self, "lbl_fr"): self.lbl_fr.configure(text=str(fr_done))
        if hasattr(self, "lbl_nd"): self.lbl_nd.configure(text=str(not_done))
        if log_msg:
            self.log("📊 PO metrics successfully refreshed.")

    def _right_section(self, parent, title: str, accent: str) -> ctk.CTkFrame:
        """Helper: creates a collapsible section card for the right panel."""
        outer = ctk.CTkFrame(parent, fg_color=COLOR_CARD, corner_radius=12,
                             border_width=1, border_color=COLOR_BORDER)
        outer.pack(fill="x", padx=16, pady=(0, 14))

        header = ctk.CTkFrame(outer, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(10, 4))

        indicator = ctk.CTkLabel(header, text="▾",
                                 font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                                 text_color=accent, width=14)
        indicator.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(header, text=title,
                     font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
                     text_color=accent).pack(side="left")

        # Divider
        ctk.CTkFrame(outer, fg_color=COLOR_BORDER, height=1).pack(fill="x", padx=12, pady=(0, 6))

        body = ctk.CTkFrame(outer, fg_color="transparent")
        body.pack(fill="x")

        # Collapse/expand toggle
        def _toggle(e=None, b=body, ind=indicator):
            if b.winfo_ismapped():
                b.pack_forget()
                ind.configure(text="▸")
            else:
                b.pack(fill="x")
                ind.configure(text="▾")

        header.bind("<Button-1>", _toggle)
        indicator.bind("<Button-1>", _toggle)

        return body

    def _right_btn(self, parent, icon: str, label: str, cmd, accent=None):
        """Helper: single utility button row for the right panel."""
        col = accent if accent else COLOR_TEXT_MAIN
        btn = ctk.CTkButton(
            parent, text=f"{icon}  {label}",
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=col,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            border_width=1, border_color=COLOR_BORDER,
            corner_radius=8, anchor="w", height=34,
            command=cmd
        )
        btn.pack(fill="x", padx=12, pady=3)
        btn.bind("<Enter>", lambda e: btn.configure(border_color=COLOR_BORDER_FOCUS, border_width=2), add="+")
        btn.bind("<Leave>", lambda e: btn.configure(border_color=COLOR_BORDER, border_width=1), add="+")
        return btn

    def _create_right_panel(self):
        right_outer = ctk.CTkFrame(self, fg_color=COLOR_RAIL, corner_radius=0,
                                   border_width=1, border_color=COLOR_BORDER)
        right_outer.grid(row=0, column=2, sticky="nsew")

        # ── Panel Header ──────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(right_outer, fg_color=COLOR_CARD, corner_radius=0,
                           border_width=0, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        ctk.CTkLabel(
            hdr, text="⚙  Quick Pipeline Controls",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(side="left", padx=20, pady=16)

        # Live process count badge
        self.proc_badge = ctk.CTkLabel(
            hdr, text="● 0 running",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            fg_color=COLOR_INNER, text_color=COLOR_EMERALD,
            corner_radius=6, padx=8, pady=3
        )
        self.proc_badge.pack(side="right", padx=16)

        # ── Scrollable body ───────────────────────────────────────────────────
        right = ctk.CTkScrollableFrame(right_outer, fg_color="transparent", corner_radius=0)
        right.pack(fill="both", expand=True, pady=(8, 0))

        # -- AUTOPILOT ---------------------------------------------------------
        ap_card = ctk.CTkFrame(right, fg_color=COLOR_CARD, corner_radius=12,
                               border_width=1, border_color=COLOR_BORDER)
        ap_card.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(
            ap_card, text="◈  AUTOPILOT",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_CYAN
        ).pack(anchor="w", padx=14, pady=(12, 6))

        ctk.CTkFrame(ap_card, fg_color=COLOR_BORDER, height=1).pack(fill="x", padx=12, pady=(0, 8))

        ap_btn = ctk.CTkButton(
            ap_card, text="🚀  Run Auto-Pilot Pipeline",
            corner_radius=8, height=40,
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self.run_autopilot
        )
        ap_btn.pack(padx=10, pady=(0, 6), fill="x")

        pause_row = ctk.CTkFrame(ap_card, fg_color="transparent")
        pause_row.pack(fill="x", padx=10, pady=(0, 6))
        pause_row.columnconfigure(0, weight=1)
        pause_row.columnconfigure(1, weight=1)
        pause_row.columnconfigure(2, weight=1)

        self.pause_btn = ctk.CTkButton(
            pause_row, text="⏸ Pause",
            corner_radius=8, height=32,
            fg_color=COLOR_INNER, hover_color="#2A2A14",
            text_color=COLOR_AMBER,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            border_width=1, border_color="#3A3010",
            command=self.pause_pipeline
        )
        self.pause_btn.grid(row=0, column=0, padx=(0, 2), sticky="ew")

        self.resume_btn = ctk.CTkButton(
            pause_row, text="▶ Resume",
            corner_radius=8, height=32,
            fg_color=COLOR_INNER, hover_color="#0A2A14",
            text_color=COLOR_EMERALD,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            border_width=1, border_color="#103020",
            state="disabled",
            command=self.resume_pipeline
        )
        self.resume_btn.grid(row=0, column=1, padx=2, sticky="ew")

        clear_session_btn = ctk.CTkButton(
            pause_row, text="✕ Clear",
            corner_radius=8, height=32,
            fg_color=COLOR_INNER, hover_color="#1A0808",
            text_color=COLOR_TEXT_MUTED,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            border_width=1, border_color="#2A1010",
            command=self.clear_session
        )
        clear_session_btn.grid(row=0, column=2, padx=(2, 0), sticky="ew")

        ToolTip(self.pause_btn,    text="Suspend all running processes.\nSafe to close UI - state is saved.", delay_ms=400)
        ToolTip(self.resume_btn,   text="Resume a previously paused / saved session.", delay_ms=400)
        ToolTip(clear_session_btn, text="Clear the saved session file.\nUse if a paused session is stale or stuck.", delay_ms=400)

        kill_btn = ctk.CTkButton(
            ap_card, text="🛑   Terminate All Processes",
            corner_radius=8, height=34,
            fg_color=COLOR_INNER, hover_color="#3B0A14", text_color=COLOR_ROSE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            border_width=1, border_color="#3B1020",
            command=self.terminate_all_processes
        )
        kill_btn.pack(padx=12, pady=(0, 12), fill="x")

        # ── SECTION 1: DATA & GENERATION ─────────────────────────────────────
        s1 = self._right_section(right, "DATA & GENERATION", COLOR_CYAN)
        self._right_btn(s1, "📑", "Split Testsheets - TIANU",      self.split_testsheets_tianu)
        self._right_btn(s1, "📊", "Extract Substation Data (Scan)", self.extract_substation_data)
        self._right_btn(s1, "🔄", "Sync & Merge to Master List",   self.sync_to_master_list)
        self._right_btn(s1, "⚡", "Generate Master CSV",            self.generate_csv)
        self._right_btn(s1, "📄", "Generate Reports (Page)",        self.run_reports)
        self._right_btn(s1, "🔥", "Inject FLIR & Digital Images",   self.inject_flir_images_gui)
        self._right_btn(s1, "🎯", "Regenerate by Page Path / Link", self.regenerate_pages_gui)
        self._right_btn(s1, "📝", "Executive Summary",              self.gen_summary)
        ctk.CTkFrame(s1, fg_color="transparent", height=6).pack()  # bottom padding

        # ── SECTION 2: MERGE & EXPORT ─────────────────────────────────────────
        s2 = self._right_section(right, "MERGE & EXPORT", COLOR_INDIGO)
        self._right_btn(s2, "📑", "Merge Word Reports",       self.merge_word)
        self._right_btn(s2, "🧩", "Merge PDF Files",          self.merge_pdfs)
        self._right_btn(s2, "🔄", "Convert Word → PDF",       self.convert_pdf)
        self._right_btn(s2, "🔊", "Insert Sound (WAV) → PDF", self.insert_sound)
        self._right_btn(s2, "📂", "Open 08. INSERT SOUND",    self.open_insert_sound_folder)
        ctk.CTkFrame(s2, fg_color="transparent", height=6).pack()

        # ── SECTION 3: TOOLS & UTILITIES ──────────────────────────────────────
        s3 = self._right_section(right, "TOOLS & UTILITIES", COLOR_VIOLET)
        self._right_btn(s3, "🌐", "Convert HTML → JPG",    self.run_html_to_jpg)
        self._right_btn(s3, "🧹", "Cleanup Word Format",    self.cleanup_word)
        self._right_btn(s3, "🔍", "Check Missing Reports",  self.check_missing)
        self._right_btn(s3, "✏️", "Convert QUICK → FULL PDF", self.convert_quick_to_full)
        ctk.CTkFrame(s3, fg_color="transparent", height=6).pack()

        # ── STATUS BAR ────────────────────────────────────────────────────────
        self.status_frame = ctk.CTkFrame(
            right_outer, fg_color=COLOR_CARD, corner_radius=0,
            border_width=0, height=46
        )
        self.status_frame.pack(side="bottom", fill="x")
        self.status_frame.pack_propagate(False)

        self.status_dot = ctk.CTkLabel(
            self.status_frame, text="●",
            font=ctk.CTkFont(size=14), text_color=COLOR_EMERALD
        )
        self.status_dot.pack(side="left", padx=(16, 6), pady=14)

        self.status_label = ctk.CTkLabel(
            self.status_frame, text="System Ready",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        )
        self.status_label.pack(side="left", pady=14)

        # Separator before clock
        ctk.CTkFrame(self.status_frame, fg_color=COLOR_BORDER, width=1, height=22).pack(side="right", padx=10)

        self.clock_label = ctk.CTkLabel(
            self.status_frame, text="",
            font=ctk.CTkFont(family=FONT_MONO, size=10),
            text_color=COLOR_TEXT_DIM
        )
        self.clock_label.pack(side="right", padx=8)
        self._tick_clock()
        self._animate_status_dot()

    def terminate_all_processes(self):
        """Force terminates running child subprocesses and kills active background automation engines."""
        if not messagebox.askyesno("Confirm Termination",
                                   "Are you sure you want to forcibly stop all pipeline processes and background engines?"):
            return

        self.log("\n🛑 INITIATING SYSTEM PROCESS TERMINATION...")

        for proc in list(self.active_processes):
            try:
                proc.terminate()
                proc.kill()
            except Exception:
                pass
        self.active_processes.clear()

        targets = ["python.exe", "powershell.exe", "WINWORD.EXE"]
        for target in targets:
            try:
                subprocess.run(f"taskkill /F /IM {target}", shell=True, stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL)
                self.log(f"  - Killed all active '{target}' instances.")
            except Exception as e:
                self.log(f"  - Failed killing '{target}': {str(e)}")

        self.log("🛑 ALL PROCESSES TERMINATED SUCCESSFULLY.")
        self._delete_session()
        self._hide_pause_banner()
        self._is_paused = False
        self.pause_btn.configure(state="normal", text="⏸ Pause")
        self.resume_btn.configure(state="disabled")
        self.set_status("System Ready", state="ready")
        self.update_po_metrics()

    def _switch_folder(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.selection_screen = FolderSelectionScreen(self, self.load_substation_dashboard)
        self.selection_screen.pack(fill="both", expand=True)

    # ==========================================================================
    #  PAUSE / RESUME SESSION ENGINE
    # ==========================================================================

    def _session_path(self) -> str:
        return self._session_file or os.path.join(CENTRAL_SCRIPT_DIR, SESSION_FILENAME)

    def _save_session(self):
        """Persist all active PIDs + their commands to disk so we can resume later."""
        records = []
        for proc in self.active_processes:
            try:
                records.append({
                    "pid": proc.pid,
                    "cmd": " ".join(proc.args) if isinstance(proc.args, list) else str(proc.args),
                    "paused_at": time.strftime("%Y-%m-%d %H:%M:%S")
                })
            except Exception:
                pass
        data = {
            "workspace": self.target_substation_dir,
            "paused_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "processes": records
        }
        try:
            with open(self._session_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.log(f"💾 Session state saved → {self._session_path()}", level="SYS")
        except Exception as e:
            self.log(f"⚠ Could not save session: {e}", level="WARN")

    def _load_session(self) -> dict | None:
        """Read a previously saved session JSON, return None if not found / invalid."""
        path = self._session_path()
        if not os.path.exists(path):
            return None
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _delete_session(self):
        path = self._session_path()
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass

    def clear_session(self):
        """Manually clear/delete the saved session state file."""
        path = self._session_path()
        self._delete_session()
        self._hide_pause_banner()
        self._is_paused = False
        self.pause_btn.configure(state="normal", text="⏸ Pause")
        self.resume_btn.configure(state="disabled")
        self.set_status("System Ready", state="ready")
        self.log("🗑 Session state cleared & banner removed.", level="SYS")
        messagebox.showinfo("Session Cleared", "Saved session state file and pause banner cleared.")

    def pause_pipeline(self):
        """Suspend all running subprocesses and save PIDs to disk."""
        if not self.active_processes:
            messagebox.showinfo("Nothing Running", "No active pipeline processes to pause.")
            return

        if not HAS_PSUTIL:
            messagebox.showerror(
                "psutil Missing",
                "Pause requires psutil.\n\nInstall it with:\n    pip install psutil"
            )
            return

        suspended = 0
        for proc in self.active_processes:
            try:
                ps = psutil.Process(proc.pid)
                # Suspend the entire process tree (children too)
                for child in ps.children(recursive=True):
                    try:
                        child.suspend()
                    except Exception:
                        pass
                ps.suspend()
                suspended += 1
            except Exception as e:
                self.log(f"⚠ Could not suspend PID {proc.pid}: {e}", level="WARN")

        self._is_paused = True
        self._save_session()

        # Update UI state
        self.pause_btn.configure(state="disabled", text="⏸  Paused")
        self.resume_btn.configure(state="normal")
        self.set_status("⏸ Pipeline Paused", state="running")
        self.status_dot.configure(text_color=COLOR_AMBER)
        self._show_pause_banner()

        self.log(f"⏸ Pipeline PAUSED — {suspended} process(es) suspended.", level="SYS")
        self.log("   Safe to close the UI. Reopen & pick the same workspace to resume.", level="SYS")

    def resume_pipeline(self):
        """Un-suspend all previously paused processes and start monitoring them."""
        if not HAS_PSUTIL:
            messagebox.showerror("psutil Missing", "Resume requires psutil.\n\nInstall it with:\n    pip install psutil")
            return

        # Try to find live processes from memory first, then fall back to session file
        targets = list(self.active_processes)
        session = self._load_session()
        restored_pids = []

        if not targets and session:
            # Re-attach to saved PIDs
            for rec in session.get("processes", []):
                pid = rec.get("pid")
                if pid:
                    try:
                        ps = psutil.Process(pid)
                        if ps.is_running():
                            restored_pids.append((pid, rec.get("cmd", "??")))
                    except psutil.NoSuchProcess:
                        self.log(f"⚠ PID {pid} no longer exists (may have finished).", level="WARN")

        resumed = 0
        # Resume processes still in memory
        for proc in targets:
            try:
                ps = psutil.Process(proc.pid)
                for child in ps.children(recursive=True):
                    try:
                        child.resume()
                    except Exception:
                        pass
                ps.resume()
                resumed += 1
            except Exception as e:
                self.log(f"⚠ Could not resume PID {proc.pid}: {e}", level="WARN")

        # Resume re-attached PIDs (from session file, not in memory)
        for pid, cmd in restored_pids:
            try:
                ps = psutil.Process(pid)
                for child in ps.children(recursive=True):
                    try:
                        child.resume()
                    except Exception:
                        pass
                ps.resume()
                resumed += 1
                self.log(f"▶ Re-attached & resumed PID {pid}  [{cmd[:60]}]", level="SYS")
            except Exception as e:
                self.log(f"⚠ Could not resume re-attached PID {pid}: {e}", level="WARN")

        self._is_paused = False
        self._delete_session()
        self._hide_pause_banner()

        self.pause_btn.configure(state="normal", text="⏸  Pause")
        self.resume_btn.configure(state="disabled")
        self.set_status("▶ Pipeline Resumed", state="running")

        self.log(f"▶ Pipeline RESUMED — {resumed} process(es) un-suspended.", level="SYS")

        if restored_pids:
            self.log("⚠ Note: stdout from re-attached processes cannot be streamed.", level="WARN")
            self.log("   Monitoring completion via polling…", level="SYS")
            self._start_resume_poll([pid for pid, _ in restored_pids])

    def _start_resume_poll(self, pids: list):
        """Poll psutil every 3 s to detect when resumed processes finish."""
        if self._resume_poll_id:
            self.after_cancel(self._resume_poll_id)

        def _poll():
            still_alive = []
            for pid in pids:
                try:
                    ps = psutil.Process(pid)
                    if ps.is_running() and ps.status() != psutil.STATUS_ZOMBIE:
                        still_alive.append(pid)
                except psutil.NoSuchProcess:
                    self.log(f"✅ PID {pid} completed.", level="OK")

            if still_alive:
                self._resume_poll_id = self.after(3000, lambda: _poll_wrap(still_alive))
            else:
                self.log("🎉 All resumed processes have finished!", level="OK")
                self.set_status("System Ready", state="ready")
                self.update_po_metrics()

        def _poll_wrap(remaining):
            still = []
            for pid in remaining:
                try:
                    ps = psutil.Process(pid)
                    if ps.is_running() and ps.status() != psutil.STATUS_ZOMBIE:
                        still.append(pid)
                except psutil.NoSuchProcess:
                    self.log(f"✅ PID {pid} completed.", level="OK")
            if still:
                self._resume_poll_id = self.after(3000, lambda: _poll_wrap(still))
            else:
                self.log("🎉 All resumed processes have finished!", level="OK")
                self.set_status("System Ready", state="ready")
                self.update_po_metrics()

        _poll()

    def _check_for_saved_session(self):
        """Called after dashboard loads — if a session file exists for this workspace, show resume prompt."""
        session = self._load_session()
        if not session:
            return

        ws = session.get("workspace", "")
        if ws != self.target_substation_dir:
            return  # Session belongs to a different workspace

        procs = session.get("processes", [])
        paused_at = session.get("paused_at", "unknown time")
        n = len(procs)

        if not HAS_PSUTIL:
            self.log("⚠ Saved session found but psutil is not installed — cannot resume.", level="WARN")
            return

        # Check if any saved PIDs are still alive
        alive = []
        for rec in procs:
            try:
                ps = psutil.Process(rec["pid"])
                if ps.is_running():
                    alive.append(rec)
            except psutil.NoSuchProcess:
                pass

        if not alive:
            self.log("ℹ Saved session found but all processes have already completed.", level="SYS")
            self._delete_session()
            return

        # Show the banner and auto-enable resume button
        self._show_pause_banner(session_info=f"{len(alive)} process(es) paused since {paused_at}")
        self.resume_btn.configure(state="normal")
        self._is_paused = True

        self.log(f"🔔 Saved session detected: {len(alive)} suspended process(es) from {paused_at}", level="SYS")
        for rec in alive:
            self.log(f"   · PID {rec['pid']}  →  {rec.get('cmd','??')[:80]}", level="SYS")
        self.log("   Click ▶ Resume in the right panel to continue.", level="SYS")

    def _show_pause_banner(self, session_info: str = ""):
        """Inject an amber banner below the header card in the center column."""
        self._hide_pause_banner()  # Remove any old banner first

        try:
            center = self.grid_slaves(row=0, column=1)[0]
        except (IndexError, AttributeError):
            return

        msg = session_info or "Pipeline is PAUSED — safe to close the UI."
        self._pause_banner = ctk.CTkFrame(
            center, fg_color="#2A1800", corner_radius=10,
            border_width=1, border_color="#5A3800"
        )
        # Insert at position 1 (below the header card)
        self._pause_banner.pack(fill="x", pady=(0, 10), before=center.winfo_children()[1])

        ctk.CTkLabel(
            self._pause_banner,
            text=f"⏸  PAUSED  —  {msg}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_AMBER
        ).pack(side="left", padx=16, pady=10)

        resume_link = ctk.CTkButton(
            self._pause_banner,
            text="▶ Resume Now",
            fg_color="#3A2200", hover_color="#5A3800",
            text_color=COLOR_AMBER,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            corner_radius=6, height=28,
            border_width=1, border_color="#5A3800",
            command=self.resume_pipeline
        )
        resume_link.pack(side="right", padx=12, pady=10)

    def _hide_pause_banner(self):
        if hasattr(self, "_pause_banner") and self._pause_banner:
            try:
                self._pause_banner.destroy()
            except Exception:
                pass
            self._pause_banner = None

    def _on_close(self):
        """Override window close: offer to suspend rather than kill if processes are active."""
        if self.active_processes and not self._is_paused:
            choice = messagebox.askyesnocancel(
                "Active Pipeline",
                f"{len(self.active_processes)} process(es) are still running.\n\n"
                "• Yes  — Suspend & save session (resume next time)\n"
                "• No   — Terminate all processes and quit\n"
                "• Cancel — Stay in the app"
            )
            if choice is None:   # Cancel
                return
            elif choice:         # Yes — suspend
                self.pause_pipeline()
                self.destroy()
                return
            else:                # No — kill
                self.terminate_all_processes()

        elif self._is_paused:
            if not messagebox.askyesno(
                "Paused Session",
                "A paused session is saved.\nThe processes will stay suspended until you resume.\n\nClose anyway?"
            ):
                return

        self.destroy()

    def _clear_console(self):
        tw = self.console._textbox
        tw.configure(state="normal")
        tw.delete("1.0", tk.END)
        tw.configure(state="disabled")

    def _tick_clock(self):
        """Update the clock label every second."""
        try:
            self.clock_label.configure(text=time.strftime("%H:%M:%S"))
            self.after(1000, self._tick_clock)
        except Exception:
            pass

    def _animate_status_dot(self):
        """Pulse the status dot opacity when a process is running."""
        try:
            if self.active_processes:
                current = self.status_dot.cget("text_color")
                # Toggle between full amber and dim amber for a pulse effect
                next_col = COLOR_TEXT_DIM if current != COLOR_TEXT_DIM else COLOR_AMBER
                self.status_dot.configure(text_color=next_col)
            self.after(600, self._animate_status_dot)
        except Exception:
            pass

    def _update_proc_badge(self):
        """Keep the running-process badge in the right panel header in sync."""
        try:
            count = len(self.active_processes)
            if count:
                self.proc_badge.configure(
                    text=f"● {count} running",
                    text_color=COLOR_AMBER
                )
            else:
                self.proc_badge.configure(
                    text="● 0 running",
                    text_color=COLOR_EMERALD
                )
        except Exception:
            pass

    def log(self, text: str, level: str = "INFO"):
        """Write a timestamped, colour-coded line to the console."""
        tw = self.console._textbox
        tw.configure(state="normal")

        ts = time.strftime("%H:%M:%S")
        tw.insert(tk.END, f"[{ts}] ", "TS")

        # Auto-detect level from text content if not explicit
        if level == "INFO":
            t_lower = text.lower()
            if any(k in t_lower for k in ("error", "❌", "fail", "critical")):
                level = "ERROR"
            elif any(k in t_lower for k in ("warn", "⚠", "missing", "kiv")):
                level = "WARN"
            elif any(k in t_lower for k in ("done", "success", "✅", "complete", "🎉", "ok")):
                level = "OK"
            elif text.strip().startswith(">"):
                level = "CMD"
            elif any(k in t_lower for k in ("terminated", "🛑", "system", "init", "launch")):
                level = "SYS"

        tag = level if level in LOG_COLORS else "INFO"
        tw.insert(tk.END, f"[{tag:<5}] ", tag)
        tw.insert(tk.END, text + "\n")
        tw.configure(state="disabled")
        self.console.see(tk.END)

    def set_status(self, text, state="ready"):
        if state == "running":
            self.status_dot.configure(text_color=COLOR_AMBER)
            self.status_label.configure(text=text)
        else:
            self.status_dot.configure(text_color=COLOR_EMERALD)
            self.status_label.configure(text="System Ready")
        self._update_proc_badge()

    def run_async(self, func, *args):
        self.set_status("Executing...", state="running")
        threading.Thread(target=self._execution_wrapper, args=(func, *args), daemon=True).start()

    def _execution_wrapper(self, func, *args):
        try:
            func(*args)
        except Exception as e:
            self.log(f"❌ CRITICAL ERROR: {str(e)}", level="ERROR")
        finally:
            self.set_status("System Ready", state="ready")
            self.update_po_metrics()

    def run_cmd(self, command):
        self.log(f"> {command}", level="CMD")
        self.log("─" * 60, level="INFO")

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        if self.target_substation_dir:
            env["BASE_DIR"] = self.target_substation_dir
            env["MAPPING_DIR"] = self.mapping_dir
            env["TEMPLATE_DIR"] = self.template_dir
            if hasattr(self, "template_ir_dir") and self.template_ir_dir:
                env["TEMPLATE_IR_DIR"] = self.template_ir_dir
            if hasattr(self, "template_ir_ue_dir") and self.template_ir_ue_dir:
                env["TEMPLATE_IR_UE_DIR"] = self.template_ir_ue_dir
            if hasattr(self, "template_ir_ue_tev_dir") and self.template_ir_ue_tev_dir:
                env["TEMPLATE_IR_UE_TEV_DIR"] = self.template_ir_ue_tev_dir
            env["OUTPUT_DIR"] = self.output_dir
            env["RM_DIR"] = os.path.join(self.target_substation_dir, "03. RAW MATERIAL")

        process = subprocess.Popen(
            command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, cwd=self.target_substation_dir, env=env, bufsize=1
        )
        self.active_processes.append(process)

        for line in process.stdout:
            cleaned_line = line.strip()
            if cleaned_line:
                # Filter out noisy openpyxl user warnings if any slip through
                if "Data Validation extension is not supported" in cleaned_line or cleaned_line == "warn(msg)":
                    continue
                self.log(cleaned_line)

        process.wait()
        if process in self.active_processes:
            self.active_processes.remove(process)

    def split_testsheets_tianu(self):
        script_path = os.path.join(self.script_dir, "split_testsheets_tianu.py")
        if not os.path.exists(script_path):
            alt_path = os.path.join(r"C:\Users\aremi\Desktop\REPORT CBM\SCRIPT", "split_testsheets_tianu.py")
            if os.path.exists(alt_path):
                script_path = alt_path
            else:
                messagebox.showerror("Script Missing", f"split_testsheets_tianu.py not found at:\n{script_path}")
                return

        if not self.target_substation_dir:
            messagebox.showwarning("Workspace Not Selected", "Please select a substation workspace first.")
            return

        confirm = messagebox.askyesno(
            "Split Testsheets - TIANU",
            "This will split daily .xlsm test sheets into individual substation Excel files (2 sheets each: Substation Data + VI).\n\n"
            f"Workspace:\n{self.target_substation_dir}\n\n"
            "Output will be saved in: 04. TESTSHEET\\INDIVIDUAL EXCEL\n\n"
            "Do you want to proceed?",
            parent=self
        )
        if not confirm:
            return

        def task():
            self.log("📑 Starting Daily Testsheets Splitter (TIANU)...", level="INFO")
            self.run_cmd(f'python "{script_path}" --dir "{self.target_substation_dir}"')
            self.log("🎉 TESTSHEET SEPARATION COMPLETED!", level="OK")
            self.update_po_metrics()

        self.run_async(task)

    def extract_substation_data(self):
        script_path = os.path.join(self.script_dir, "generate_excel.py")
        if not os.path.exists(script_path):
            messagebox.showerror("Script Missing", f"generate_excel.py not found at:\n{script_path}")
            return

        # Check existing checklist to offer smart default
        cl_path = os.path.join(self.target_substation_dir, "SUBSTATION_CHECKLIST.xlsx")
        default_scope = "ALL"
        if os.path.exists(cl_path):
            try:
                import openpyxl
                wb_tmp = openpyxl.load_workbook(cl_path, read_only=True)
                if 'Substation Checklist' in wb_tmp.sheetnames:
                    ws_tmp = wb_tmp['Substation Checklist']
                    nums = [int(str(ws_tmp.cell(r, 1).value).strip()) 
                            for r in range(5, ws_tmp.max_row + 1) 
                            if ws_tmp.cell(r, 1).value is not None and str(ws_tmp.cell(r, 1).value).strip().isdigit()]
                    if nums:
                        max_pe = max(nums)
                        default_scope = f"{max_pe + 1}-210"
                wb_tmp.close()
            except Exception:
                pass

        dialog = ctk.CTkToplevel(self)
        dialog.title("Extract Substation Data (Scan Testsheets)")
        dialog.geometry("540x440")
        dialog.resizable(False, False)
        dialog.configure(fg_color=COLOR_CANVAS)
        dialog.transient(self)
        dialog.grab_set()

        # Header
        hdr = ctk.CTkFrame(dialog, fg_color="transparent")
        hdr.pack(fill="x", padx=24, pady=(20, 10))
        ctk.CTkLabel(
            hdr, text="📊 Extract Substation Specifications",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w")
        ctk.CTkLabel(
            hdr, text="Extract specs, circuit panels, ratings, and serials from test sheets into SUBSTATION_CHECKLIST.xlsx.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED, wraplength=480
        ).pack(anchor="w", pady=(4, 0))

        # Main Card Frame
        card = ctk.CTkFrame(dialog, fg_color=COLOR_CARD, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="both", expand=True, padx=24, pady=(5, 15))

        # 1. Company Format Selector
        ctk.CTkLabel(
            card, text="🏢 Company Testsheet Format:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w", padx=16, pady=(14, 6))

        format_var = ctk.StringVar(value="AUTO-DETECT")
        seg_btn = ctk.CTkSegmentedButton(
            card,
            values=["AUTO-DETECT", "TIANU", "EET"],
            variable=format_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            selected_color=COLOR_CYAN,
            selected_hover_color=COLOR_CYAN_HOVER,
            unselected_color=COLOR_INNER,
            text_color="#000000"
        )
        seg_btn.pack(fill="x", padx=16, pady=(0, 6))

        lbl_desc = ctk.CTkLabel(
            card,
            text="✨ Auto-Detect automatically identifies EET ('PCE Testsheet') or TIANU ('Substation Data') per file.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=COLOR_TEXT_MUTED, wraplength=460, justify="left"
        )
        lbl_desc.pack(anchor="w", padx=16, pady=(0, 12))

        def update_format_desc(val):
            if val == "TIANU":
                lbl_desc.configure(text="📌 TIANU Engine: Reads 'Substation Data' (CBA) and 'VI' checklist sheets.")
            elif val == "EET":
                lbl_desc.configure(text="📌 EET Engine: Reads 'PCE Testsheet' and 'PCE VI' checklist sheets.")
            else:
                lbl_desc.configure(text="✨ Auto-Detect automatically identifies EET ('PCE Testsheet') or TIANU ('Substation Data') per file.")

        seg_btn.configure(command=update_format_desc)

        # 2. Scope Selection
        ctk.CTkLabel(
            card, text="🎯 Substation Scope / Filter:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w", padx=16, pady=(4, 4))

        ctk.CTkLabel(
            card, text="Examples: ALL (all substations)  •  1-210  •  35 (exact PE)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=16, pady=(0, 4))

        scope_entry = ctk.CTkEntry(
            card,
            font=ctk.CTkFont(family=FONT_MONO, size=12),
            fg_color=COLOR_INNER,
            border_width=1,
            border_color=COLOR_BORDER,
            height=34
        )
        scope_entry.pack(fill="x", padx=16, pady=(0, 10))
        scope_entry.insert(0, default_scope)

        # Bottom Action Buttons
        btn_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_bar.pack(fill="x", padx=24, pady=(0, 16))

        def on_cancel():
            dialog.destroy()

        def on_submit():
            user_scope = scope_entry.get().strip() or "ALL"
            chosen_format = format_var.get().strip()
            dialog.destroy()

            def task():
                self.log(f"📊 Extracting Substation Data (Scope: {user_scope}, Format: {chosen_format})...")
                if chosen_format in ["TIANU", "EET"]:
                    self.run_cmd(f'python "{script_path}" "{user_scope}" "{chosen_format}"')
                else:
                    self.run_cmd(f'python "{script_path}" "{user_scope}"')
                self.log("🎉 EXTRACTION COMPLETED SUCCESSFULLY!", level="OK")

            self.run_async(task)

        cancel_btn = ctk.CTkButton(
            btn_bar, text="Cancel", width=90, height=34,
            fg_color="transparent", hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MUTED, border_width=1, border_color=COLOR_BORDER,
            command=on_cancel
        )
        cancel_btn.pack(side="left")

        submit_btn = ctk.CTkButton(
            btn_bar, text="⚡ Start Extraction", height=34,
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER,
            text_color="#000000", font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=on_submit
        )
        submit_btn.pack(side="right")

    def sync_to_master_list(self):
        script_path = os.path.join(self.script_dir, "sync_master_list.py")
        if not os.path.exists(script_path):
            messagebox.showerror("Script Missing", f"sync_master_list.py not found at:\n{script_path}")
            return

        prompt_text = (
            "Select Substation Scope to Sync to Master List:\n\n"
            "• Exact substation only: e.g. 35   or   12\n"
            "• Custom range:          e.g. 35-400\n"
            "• All substations:       ALL"
        )
        user_scope = simpledialog.askstring(
            "Sync to Master List",
            prompt_text,
            initialvalue="ALL",
            parent=self
        )

        if user_scope is None:
            return

        user_scope = user_scope.strip() or "ALL"

        def task():
            self.log(f"🔄 Merging & Transferring to Master List.xlsx (Scope: {user_scope})...")
            self.run_cmd(f'python "{script_path}" "{user_scope}"')
            self.log("\n🎉 MASTER LIST SYNC COMPLETED SUCCESSFULLY!", level="OK")

        self.run_async(task)

    def raw_material_declaration(self):
        if not os.path.exists(self.mapping_dir):
            messagebox.showerror("Error", f"EXCELL MAPPING directory not found:\n{self.mapping_dir}")
            return

        csv_files = [f for f in os.listdir(self.mapping_dir) if f.lower().endswith(".csv")]
        if not csv_files:
            messagebox.showinfo("Info", "No CSV files found in EXCELL MAPPING.")
            return

        dlg = ExplorerListWindow(self, title="Raw Material Declaration", header_text="Select Equipment Mapping CSV:",
                                 items=csv_files)
        self.wait_window(dlg)

        if dlg.selected_items:
            chosen_csv = dlg.selected_items[0]
            output_img_dir = os.path.join(self.gen_report_dir, "RENAMED_IMAGES")

            try:
                from raw_material_wizard import RawMaterialDeclarationWindow
                wizard = RawMaterialDeclarationWindow(
                    self, self.target_substation_dir, self.mapping_dir,
                    chosen_csv, self.master_list_path, output_img_dir
                )
                self.wait_window(wizard)
            except ImportError:
                messagebox.showerror("Module Missing", "Could not find 'raw_material_wizard.py' in the script folder.")

    def auto_rename_images_gui(self):
        if not self.target_substation_dir or not os.path.exists(self.target_substation_dir):
            messagebox.showerror("Error", "Please select a valid substation workspace first.")
            return

        self.apply_custom_paths()
        rm_dir = self.rm_dir or os.path.join(self.target_substation_dir, "03. RAW MATERIAL")
        ex_dir = os.path.join(self.target_substation_dir, "EXAMPLE IMAGE")

        def handle_selected_folders(folders):
            if not folders:
                return

            def task():
                try:
                    from auto_rename_images import auto_rename_batch_folders
                    self.log(f"🏷️ Launching Auto Rename on {len(folders)} folder(s)...", level="SYS")
                    res = auto_rename_batch_folders(folders, mapping_dir=self.mapping_dir, log_fn=self.log)
                    if res.get("success_count", 0) > 0:
                        self.log(f"🎉 AUTO RENAME COMPLETE: {res['success_count']}/{len(folders)} substation(s) successfully renamed!", level="OK")
                    else:
                        self.log("⚠ Auto rename finished with warnings/errors. Check logs above.", level="WARN")
                except ImportError:
                    self.log("❌ Could not import 'auto_rename_images.py'", level="ERROR")
                except Exception as ex:
                    self.log(f"❌ Auto rename error: {str(ex)}", level="ERROR")

            self.run_async(task)

        dialog = AutoRenameDialog(self, raw_dir=rm_dir, example_dir=ex_dir, on_confirm_cb=handle_selected_folders)
        self.wait_window(dialog)

    def run_html_to_jpg(self):
        script_path = os.path.join(self.script_dir, "html_to_jpg.py")
        if not os.path.exists(script_path):
            messagebox.showerror("Script Missing", f"html_to_jpg.py not found at:\n{script_path}")
            return

        def task():
            self.log("🌐 Launching Playwright HTML -> JPG conversion...")
            self.run_cmd(f'python "{script_path}"')

        self.run_async(task)

    def generate_csv(self):
        if not os.path.exists(self.master_list_path):
            messagebox.showerror("Error", f"Master List.xlsx not found at:\n{self.master_list_path}")
            return

        if not HAS_OPENPYXL:
            messagebox.showwarning("Package Missing", "openpyxl is not installed. Please run:\npip install openpyxl")
            return

        try:
            wb = openpyxl.load_workbook(self.master_list_path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception as e:
            messagebox.showerror("Error Reading Excel", str(e))
            return

        if not sheets:
            messagebox.showinfo("Info", "No sheets found in Master List.xlsx.")
            return

        dlg = ExplorerListWindow(self, title="Generate CSV", header_text="Select Master List Sheets:", items=sheets)
        self.wait_window(dlg)

        if dlg.selected_items:
            input_val = ", ".join(dlg.selected_items)

            def task():
                script_path = os.path.join(self.script_dir, "generate_csv.py")
                self.run_cmd(f'python "{script_path}" "{input_val}"')

            self.run_async(task)

    def merge_pdfs(self):
        if not os.path.exists(self.output_dir):
            messagebox.showerror("Error", "OUTPUT FILE directory does not exist.")
            return

        folders = [f for f in os.listdir(self.output_dir) if os.path.isdir(os.path.join(self.output_dir, f))]
        if not folders:
            messagebox.showinfo("Info", "No subfolders found in OUTPUT FILE.")
            return

        dlg = ExplorerListWindow(self, title="Merge PDF Files", header_text="Select Folders to Merge:", items=folders)
        self.wait_window(dlg)

        if dlg.selected_items:
            def task():
                script_path = os.path.join(self.script_dir, "merge_pdfs.py")
                for folder in dlg.selected_items:
                    self.run_cmd(f'python "{script_path}" "{folder}"')

            self.run_async(task)

    def cleanup_word(self):
        indiv_dir = os.path.join(self.output_dir, "INDIVIDUAL")
        if not os.path.exists(indiv_dir):
            messagebox.showerror("Error", "INDIVIDUAL directory does not exist.")
            return

        folders = [f for f in os.listdir(indiv_dir) if os.path.isdir(os.path.join(indiv_dir, f))]
        if not folders:
            messagebox.showinfo("Info", "No folders found in INDIVIDUAL.")
            return

        dlg = ExplorerListWindow(self, title="Cleanup Word Formatting", header_text="Select Folders to Cleanup:",
                                 items=folders)
        self.wait_window(dlg)

        if dlg.selected_items:
            def task():
                ps_script = os.path.join(self.script_dir, "cleanup_reports.ps1")
                for folder in dlg.selected_items:
                    cmd = f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{ps_script}" -SubFolder "{folder}"'
                    self.run_cmd(cmd)

            self.run_async(task)

    def run_reports(self):
        csv_files = filedialog.askopenfilenames(
            initialdir=self.mapping_dir, title="Select CSV Files (Hold Ctrl / Shift)",
            filetypes=[("CSV Files", "*.csv")]
        )
        if not csv_files: return

        pages = simpledialog.askstring("Page Selection", "Enter pages (e.g., page1, page2) or 'all':",
                                       initialvalue="all")
        if not pages: return

        def task():
            for csv_file in csv_files:
                csv_name = os.path.basename(csv_file)
                self.log(f"📋 Processing CSV (Page Mode): {csv_name}")
                script_path = os.path.join(self.script_dir, "generate_page_report.py")
                self.run_cmd(f'python "{script_path}" "{csv_file}" "{pages}"')

        self.run_async(task)

    def inject_flir_images_gui(self):
        """
        Launches the high-speed FLIR Radiometric and Digital Photo Injector
        for Word reports in 01. FULL REPORT.
        """
        active_dir = self.target_substation_dir
        if not active_dir or not os.path.exists(active_dir):
            messagebox.showerror("No Workspace", "Please select or configure an Active Substation / PO Workspace first.")
            return

        fr_dir = self.fr_dir or os.path.join(active_dir, "01. FULL REPORT")
        if not os.path.exists(fr_dir):
            messagebox.showerror("Folder Missing", f"01. FULL REPORT folder not found in:\n{active_dir}")
            return

        docx_files = [f for f in os.listdir(fr_dir) if f.lower().endswith(".docx") and not f.startswith("~")]
        if not docx_files:
            messagebox.showwarning("No Reports", f"No Word reports (.docx) found in:\n{fr_dir}")
            return

        prompt_msg = (
            f"Active Workspace:\n{os.path.basename(active_dir)}\n\n"
            f"Found {len(docx_files)} report document(s) in 01. FULL REPORT.\n\n"
            "Select Parts to Inject:\n"
            "  • 'all' or 'A'   : All Injectable Parts (Part 01 to Part 09)\n"
            "  • '1'            : Part 01 - Summary / Overviews only\n"
            "  • '2,3,4'        : Specific Panels / Parts\n"
        )
        user_choice = simpledialog.askstring(
            "Inject FLIR & Digital Images",
            prompt_msg,
            initialvalue="all",
            parent=self
        )
        if user_choice is None:
            return

        choice_clean = user_choice.strip().lower()
        parts_arg = "all" if choice_clean in ["all", "a", ""] else choice_clean

        def task():
            self.log("🔥 Starting High-Speed FLIR & Digital Image Injector...", "INFO")
            script_path = os.path.join(self.script_dir, "inject_flir_images.py")
            if not os.path.exists(script_path):
                fallback = os.path.join(fr_dir, "inject_flir_images.py")
                if os.path.exists(fallback):
                    script_path = fallback
                else:
                    self.log(f"[ERROR] inject_flir_images.py not found in {self.script_dir}", "ERR")
                    return

            cmd = f'python "{script_path}" --workspace "{active_dir}" --substation "all" --parts "{parts_arg}"'
            self.run_cmd(cmd)
            self.log("🎉 FLIR & Digital Image Injection complete!", "OK")
            if hasattr(self, "update_po_metrics"):
                self.after(0, self.update_po_metrics)

        self.run_async(task)

    def regenerate_pages_gui(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Regenerate & Merge by Page Link / Path")
        dialog.geometry("700x520")
        dialog.minsize(550, 400)
        dialog.configure(fg_color=COLOR_CANVAS)
        dialog.transient(self)
        dialog.grab_set()

        # Header
        hdr = ctk.CTkFrame(dialog, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16, 8))
        ctk.CTkLabel(
            hdr, text="🎯 Regenerate & Merge Pages",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(anchor="w")
        ctk.CTkLabel(
            hdr, text="Paste one or more .docx file paths / links (or click Browse) to regenerate specific pages & re-merge.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", pady=(2, 0))

        # Text Box for paths
        txt_frame = ctk.CTkFrame(dialog, fg_color=COLOR_CARD, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        txt_frame.pack(fill="both", expand=True, padx=20, pady=8)

        txt_box = ctk.CTkTextbox(
            txt_frame, font=ctk.CTkFont(family=FONT_MONO, size=11),
            fg_color=COLOR_INNER, text_color=COLOR_TEXT_MAIN,
            border_width=0
        )
        txt_box.pack(fill="both", expand=True, padx=10, pady=10)

        # Checkbox & Browse bar
        opts_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        opts_bar.pack(fill="x", padx=20, pady=6)

        auto_merge_var = ctk.BooleanVar(value=True)
        chk_merge = ctk.CTkCheckBox(
            opts_bar, text="Auto-merge Word Reports by Substation after generation",
            variable=auto_merge_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLOR_TEXT_MAIN,
            fg_color=COLOR_CYAN, checkmark_color="#000000"
        )
        chk_merge.pack(side="left", padx=4)

        def browse_docx():
            init_dir = os.path.join(self.output_dir, "INDIVIDUAL") if self.output_dir and os.path.exists(os.path.join(self.output_dir, "INDIVIDUAL")) else self.output_dir
            picked = filedialog.askopenfilenames(
                parent=dialog,
                initialdir=init_dir,
                title="Select .docx Page Files (Hold Ctrl / Shift)",
                filetypes=[("Word Documents", "*.docx")]
            )
            if picked:
                current_text = txt_box.get("1.0", "end").strip()
                new_lines = "\n".join(f'"{p}"' for p in picked)
                if current_text:
                    txt_box.insert("end", "\n" + new_lines + "\n")
                else:
                    txt_box.insert("1.0", new_lines + "\n")

        browse_btn = ctk.CTkButton(
            opts_bar, text="📁 Browse .docx", width=120, height=28,
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MAIN, font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            command=browse_docx
        )
        browse_btn.pack(side="right", padx=4)

        # Action Buttons
        btn_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_bar.pack(fill="x", padx=20, pady=(6, 16))

        def on_submit():
            content = txt_box.get("1.0", "end").strip()
            if not content:
                messagebox.showwarning("No Input", "Please paste or select at least one .docx path.", parent=dialog)
                return
            
            should_merge = auto_merge_var.get()
            dialog.destroy()

            def task():
                self.log("🎯 Starting Targeted Page Regeneration & Re-Merge Pipeline...")
                script_path = os.path.join(self.script_dir, "generate_page_report.py")
                
                temp_input_file = os.path.join(self.script_dir, "_temp_page_targets.txt")
                try:
                    with open(temp_input_file, "w", encoding="utf-8") as f:
                        f.write(content)
                    
                    merge_flag = "--merge" if should_merge else "--no-merge"
                    cmd = f'python "{script_path}" --file "{temp_input_file}" {merge_flag}'
                    self.run_cmd(cmd)
                finally:
                    if os.path.exists(temp_input_file):
                        try:
                            os.remove(temp_input_file)
                        except Exception:
                            pass

            self.run_async(task)

        submit_btn = ctk.CTkButton(
            btn_bar, text="⚡ Start Regeneration", width=160, height=34,
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=on_submit
        )
        submit_btn.pack(side="right", padx=4)

        cancel_btn = ctk.CTkButton(
            btn_bar, text="Cancel", width=90, height=34,
            fg_color=COLOR_INNER, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT_MUTED, font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=dialog.destroy
        )
        cancel_btn.pack(side="right", padx=4)

    def merge_word(self):
        csv_files = filedialog.askopenfilenames(
            initialdir=self.mapping_dir, title="Select Mapping CSV Files (Hold Ctrl / Shift)",
            filetypes=[("CSV Files", "*.csv")]
        )
        if not csv_files: return

        groups = simpledialog.askstring("Target Groups", "Enter Group numbers (e.g. 1,2,5) or 'all':",
                                        initialvalue="all")
        if not groups: return

        def task():
            ps_script = os.path.join(self.script_dir, "dynamic_merge_tool.ps1")
            for csv_file in csv_files:
                csv_name = os.path.basename(csv_file)
                self.log(f"📑 Merging Word Reports for CSV: {csv_name}")
                cmd = f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{ps_script}" -csvName "{csv_name}" -targetGroupsRaw "{groups}"'
                self.run_cmd(cmd)

        self.run_async(task)

    def check_missing(self):
        def task():
            script_path = os.path.join(self.script_dir, "check_missing_reports.py")
            self.run_cmd(f'python "{script_path}"')

        self.run_async(task)

    def gen_summary(self):
        csv_files = filedialog.askopenfilenames(
            initialdir=self.mapping_dir, title="Select CSV Files for Summary",
            filetypes=[("CSV Files", "*.csv")]
        )
        if not csv_files: return

        def task():
            script_path = os.path.join(self.script_dir, "generate_summary.py")
            for csv_file in csv_files:
                self.log(f"📝 Generating Executive Summary for: {os.path.basename(csv_file)}")
                self.run_cmd(f'python "{script_path}" "{csv_file}"')

        self.run_async(task)

    def insert_sound(self):
        def task():
            script_path = os.path.join(self.script_dir, "final_insert_sounds.py")
            target_dir = self.insert_sound_dir or (os.path.join(self.target_substation_dir, "08. INSERT SOUND") if self.target_substation_dir else "")
            self.log("🔊 Starting Sound (WAV) Insertion into PDF...", "CMD")
            if target_dir and os.path.exists(target_dir):
                self.run_cmd(f'python "{script_path}" "{target_dir}"')
            else:
                self.run_cmd(f'python "{script_path}"')

        self.run_async(task)

    def open_insert_sound_folder(self):
        target = self.insert_sound_dir or (os.path.join(self.target_substation_dir, "08. INSERT SOUND") if self.target_substation_dir else None)
        if target and os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showerror("Error", f"08. INSERT SOUND directory does not exist:\n{target}")

    def convert_pdf(self):
        default_dir = ""
        if self.target_substation_dir:
            cand1 = os.path.join(self.target_substation_dir, "07. OUTPUT FILE", "INDIVIDUAL")
            cand2 = os.path.join(self.target_substation_dir, "07. OUTPUT FILE")
            if os.path.exists(cand1):
                default_dir = cand1
            elif os.path.exists(cand2):
                default_dir = cand2
            else:
                default_dir = self.target_substation_dir
        else:
            default_dir = DEFAULT_ROOT_DIR

        def on_confirm_convert(selected_files, dest_folder, append_ts=False, ts_dir=None):
            def task():
                self.log(f"📄 Starting Lossless Word to PDF conversion for {len(selected_files)} file(s)...", "CMD")
                try:
                    from convert_word_to_pdf import convert_word_files_to_pdf
                    res = convert_word_files_to_pdf(
                        selected_files,
                        destination_folder=dest_folder,
                        append_testsheet=append_ts,
                        testsheet_pdf_dir=ts_dir,
                        log_fn=self.log
                    )
                    if res.get("success"):
                        self.log(f"🎉 Successfully converted {res.get('converted_count')} file(s) to PDF with 100% Lossless Quality!", "OK")
                        msg = f"Successfully converted {res.get('converted_count')} document(s) to PDF!"
                        if append_ts:
                            msg += "\n(Attached matching Testsheet PDFs at the back)"
                        msg += f"\n\nDestination: {dest_folder or 'Same folder as Word files'}"
                        messagebox.showinfo("Conversion Complete", msg)
                    else:
                        self.log(f"⚠️ Conversion completed with {len(res.get('failed_files', []))} failure(s).", "WARN")
                except Exception as e:
                    self.log(f"❌ PDF conversion error: {e}", "ERR")

            self.run_async(task)

        ConvertWordToPdfDialog(self, default_dir=default_dir, on_convert_cb=on_confirm_convert)

    def gen_missing_quick(self):
        def task():
            script_path = os.path.join(self.script_dir, "generate_missing_quick_reports.py")
            self.run_cmd(f'python "{script_path}"')

        self.run_async(task)

    def convert_quick_to_full(self):
        def task():
            script_path = os.path.join(self.script_dir, "convert_quick_to_full.py")
            not_done_dir = os.path.join(self.target_substation_dir, "01. FULL REPORT", "NOT DONE")
            self.run_cmd(f'python "{script_path}" "{not_done_dir}"')

        self.run_async(task)

    def run_autopilot(self):
        csv_files = filedialog.askopenfilenames(
            initialdir=self.mapping_dir, title="Auto-Pilot: Select CSV Files",
            filetypes=[("CSV Files", "*.csv")]
        )
        if not csv_files: return

        pages = simpledialog.askstring("Auto-Pilot Pages", "Enter pages or 'all':", initialvalue="all")
        if not pages: return

        all_csvs = sorted([f for f in os.listdir(self.mapping_dir) if f.lower().endswith(".csv")])
        selected_indices = []
        for selected_file in csv_files:
            base_name = os.path.basename(selected_file)
            if base_name in all_csvs:
                idx = all_csvs.index(base_name) + 1
                selected_indices.append(str(idx))

        if not selected_indices:
            messagebox.showerror("Error", "No matching CSV files found in EXCELL MAPPING.")
            return

        choices_str = " ".join(selected_indices)

        def task():
            ps_script = os.path.join(self.script_dir, "orchestrate_all.ps1")
            cmd = f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{ps_script}" -choices "{choices_str}" -pages "{pages}" -baseDir "{self.target_substation_dir}"'
            self.run_cmd(cmd)

        self.run_async(task)

    def open_output_folder(self):
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showerror("Error", "OUTPUT FILE directory does not exist.")

    def open_csv_mapping_folder(self):
        if self.mapping_dir and os.path.exists(self.mapping_dir):
            os.startfile(self.mapping_dir)
        else:
            messagebox.showerror("Error", f"EXCELL MAPPING directory does not exist:\n{self.mapping_dir}")

    def open_template_folder(self):
        if self.template_dir and os.path.exists(self.template_dir):
            os.startfile(self.template_dir)
        else:
            messagebox.showerror("Error", f"Template directory does not exist:\n{self.template_dir}")

    def open_thermal_tagger_beta(self):
        """Opens the interactive Thermal Coordinate and FLIR Tagging Tool (BETA)."""
        try:
            from thermal_tagger_beta import ThermalTaggingDialog
            dialog = ThermalTaggingDialog(self, workspace_dir=self.target_substation_dir)
            dialog.focus()
            self.log("Thermal Coordinate & Tagging Tool (BETA) opened.", "INFO")
        except Exception as e:
            self.log(f"Error opening Thermal Tagging Tool: {e}", "ERROR")
            messagebox.showerror("Thermal Tagging Error", f"Could not launch tool:\n{e}")


if __name__ == "__main__":
    app = ReportManagerApp()
    app.mainloop()