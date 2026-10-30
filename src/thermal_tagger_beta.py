"""
WRAITH Thermal Tagging & Coordinate Injection System (BETA)
Supports:
1. Exact physical FLIR IR image resolution & loading.
2. Spot 1 (Sp1), Spot 2 (Sp2), and on-canvas Delta T (ΔT) measurement tags.
3. Freeform Polygon (Ar1) AND Rectangle Box (Ar1) for average area calculation.
4. Mouse scroll-wheel navigation for Substations, Pages, and Canvas.
5. Individual measurement clear buttons (Clear Sp1, Clear Sp2, Clear Ar1, Clear ΔT) + Clear All.
6. Right-click context menu to delete tags on canvas.
7. Automated injection into FLIR SDK & DOCX report generators.
"""

import os
import sys
import json
import re
import math
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import customtkinter as ctk
from pathlib import Path
from PIL import Image, ImageTk, ImageDraw

# ── Color Palette (Glass / Bento Dark Theme) ─────────────────────────────────
COLOR_CANVAS = "#000000"
COLOR_CARD = "#0A0A0A"
COLOR_CARD_HOVER = "#141414"
COLOR_INNER = "#050505"
COLOR_BORDER = "#1A1A1A"
COLOR_BORDER_FOCUS = "#38BDF8"

COLOR_CYAN = "#38BDF8"
COLOR_CYAN_HOVER = "#0284C7"
COLOR_INDIGO = "#818CF8"
COLOR_VIOLET = "#A78BFA"
COLOR_EMERALD = "#34D399"
COLOR_AMBER = "#FBBF24"
COLOR_ROSE = "#F43F5E"
COLOR_TEXT_MAIN = "#F8FAFC"
COLOR_TEXT_MUTED = "#64748B"
COLOR_TEXT_DIM = "#334155"

FONT_FAMILY = "Segoe UI"
FONT_MONO = "Consolas"

COORDS_FILE_SUFFIX = "_thermal_coords_beta.json"


class ThermalCoordinatesManager:
    """Manages persistence of Sp1, Sp2, Ar1 (Box/Polygon), and Delta T coordinates."""

    @staticmethod
    def get_coords_file_path(workspace_dir: str, csv_prefix: str = "") -> Path:
        base = Path(workspace_dir)
        mapping_dir = base / "06. EXCELL MAPPING"
        if not mapping_dir.exists():
            mapping_dir = base
        if csv_prefix:
            return mapping_dir / f"{csv_prefix}{COORDS_FILE_SUFFIX}"
        return mapping_dir / f"substation{COORDS_FILE_SUFFIX}"

    @classmethod
    def load_coordinates(cls, workspace_dir: str, csv_prefix: str = "") -> dict:
        coords_path = cls.get_coords_file_path(workspace_dir, csv_prefix)
        if coords_path.exists():
            try:
                with open(coords_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[ThermalCoords] Error loading {coords_path}: {e}")
        return {}

    @classmethod
    def save_coordinates(cls, workspace_dir: str, coords_data: dict, csv_prefix: str = ""):
        coords_path = cls.get_coords_file_path(workspace_dir, csv_prefix)
        coords_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(coords_path, "w", encoding="utf-8") as f:
                json.dump(coords_data, f, indent=2)
            print(f"[ThermalCoords] Saved {len(coords_data)} page coordinates to {coords_path.name}")
            return True
        except Exception as e:
            print(f"[ThermalCoords] Error saving {coords_path}: {e}")
            return False

    @staticmethod
    def get_default_coords_for_type(equipment_type: str = "MRMU", area: str = "CABLE COMPARTMENT", width: int = 320, height: int = 240) -> dict:
        eq = str(equipment_type).upper()
        ar = str(area).upper()

        if "OVERVIEW" in ar or "OVERVIEW" in eq:
            return {
                "sp1": {"x": int(width * 0.40), "y": int(height * 0.45), "enabled": False, "label": "Sp1", "temp": ""},
                "sp2": {"x": int(width * 0.60), "y": int(height * 0.45), "enabled": False, "label": "Sp2", "temp": ""},
                "delta_t_tag": {"x": int(width * 0.50), "y": int(height * 0.40), "enabled": False, "val": ""},
                "ar1": {
                    "type": "box",
                    "x": int(width * 0.15), "y": int(height * 0.15), "w": int(width * 0.70), "h": int(height * 0.70),
                    "points": [],
                    "enabled": True, "label": "Ar1", "temp_avg": ""
                },
                "ambient": 31.7,
                "delta_t": ""
            }
        elif "BUSHING" in ar or "BUSHING" in eq:
            return {
                "sp1": {"x": int(width * 0.42), "y": int(height * 0.40), "enabled": True, "label": "Sp1", "temp": ""},
                "sp2": {"x": int(width * 0.58), "y": int(height * 0.40), "enabled": True, "label": "Sp2", "temp": ""},
                "delta_t_tag": {"x": int(width * 0.50), "y": int(height * 0.35), "enabled": True, "val": ""},
                "ar1": {
                    "type": "box",
                    "x": int(width * 0.20), "y": int(height * 0.20), "w": int(width * 0.60), "h": int(height * 0.55),
                    "points": [],
                    "enabled": True, "label": "Ar1", "temp_avg": ""
                },
                "ambient": 31.7,
                "delta_t": ""
            }
        else:
            return {
                "sp1": {"x": int(width * 0.38), "y": int(height * 0.50), "enabled": True, "label": "Sp1", "temp": ""},
                "sp2": {"x": int(width * 0.62), "y": int(height * 0.50), "enabled": True, "label": "Sp2", "temp": ""},
                "delta_t_tag": {"x": int(width * 0.50), "y": int(height * 0.45), "enabled": True, "val": ""},
                "ar1": {
                    "type": "box",
                    "x": int(width * 0.18), "y": int(height * 0.18), "w": int(width * 0.64), "h": int(height * 0.64),
                    "points": [],
                    "enabled": True, "label": "Ar1", "temp_avg": ""
                },
                "ambient": 31.7,
                "delta_t": ""
            }



class ScrollableSearchPopup(ctk.CTkToplevel):
    """
    A modern, glass-themed searchable and scrollable dropdown popup.
    Supports full mouse-wheel scrolling and instant live search filtering.
    """
    def __init__(self, master, target_widget, values: list, on_select_cb, width=320, height=380, placeholder="🔍 Search..."):
        super().__init__(master)
        self.master = master
        self.target_widget = target_widget
        self.all_values = list(values)
        self.filtered_values = list(values)
        self.on_select_cb = on_select_cb

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(fg_color="#0A0A0A")

        # Position right beneath the target widget
        target_widget.update_idletasks()
        rx = target_widget.winfo_rootx()
        ry = target_widget.winfo_rooty() + target_widget.winfo_height() + 4
        
        # Ensure it fits on screen
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        if rx + width > screen_w:
            rx = screen_w - width - 10
        if ry + height > screen_h:
            ry = target_widget.winfo_rooty() - height - 4

        self.geometry(f"{width}x{height}+{rx}+{ry}")

        # Main frame with border
        container = ctk.CTkFrame(self, fg_color="#0A0A0A", corner_radius=10, border_width=2, border_color="#38BDF8")
        container.pack(fill="both", expand=True)

        # Header / Search row
        search_box = ctk.CTkFrame(container, fg_color="transparent")
        search_box.pack(fill="x", padx=10, pady=(10, 6))

        self.search_entry = ctk.CTkEntry(
            search_box, placeholder_text=placeholder,
            height=32, font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#050505", border_color="#1A1A1A", text_color="#F8FAFC"
        )
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        self.search_entry.bind("<Return>", self._on_enter_pressed)
        self.search_entry.bind("<Escape>", lambda e: self.destroy())

        close_btn = ctk.CTkButton(
            search_box, text="✕", width=28, height=32,
            fg_color="#141414", hover_color="#3B0A14", text_color="#64748B",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self.destroy
        )
        close_btn.pack(side="right", padx=(6, 0))

        # Count indicator
        self.count_lbl = ctk.CTkLabel(
            container, text=f"Showing {len(self.filtered_values)} items (Use mouse wheel to scroll):",
            font=ctk.CTkFont(family="Segoe UI", size=9), text_color="#64748B"
        )
        self.count_lbl.pack(anchor="w", padx=12, pady=(0, 4))

        # Scrollable list container
        self.scroll = ctk.CTkScrollableFrame(container, fg_color="#050505", corner_radius=8, border_width=1, border_color="#1A1A1A")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._bind_mousewheel_recursive(self)
        self._bind_mousewheel_recursive(self.scroll)

        self._render_items()

        self.search_entry.focus_set()
        self.bind("<FocusOut>", self._on_focus_loss)

    def _bind_mousewheel_recursive(self, widget):
        widget.bind("<MouseWheel>", self._on_mousewheel, add="+")
        for child in getattr(widget, "winfo_children", lambda: [])():
            self._bind_mousewheel_recursive(child)

    def _on_mousewheel(self, event):
        # Scroll the internal canvas
        try:
            if hasattr(self.scroll, "_parent_canvas"):
                self.scroll._parent_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _on_focus_loss(self, event):
        self.after(150, self._check_focus_and_destroy)

    def _check_focus_and_destroy(self):
        try:
            cur_focus = self.focus_get()
            if not cur_focus or not str(cur_focus).startswith(str(self)):
                self.destroy()
        except Exception:
            pass

    def _on_search_key(self, event):
        query = self.search_entry.get().strip().lower()
        if not query:
            self.filtered_values = list(self.all_values)
        else:
            self.filtered_values = [v for v in self.all_values if query in str(v).lower()]
        self.count_lbl.configure(text=f"Showing {len(self.filtered_values)} items:")
        self._render_items()

    def _on_enter_pressed(self, event):
        if self.filtered_values:
            self._select_item(self.filtered_values[0])

    def _render_items(self):
        for w in self.scroll.winfo_children():
            w.destroy()

        if not self.filtered_values:
            ctk.CTkLabel(
                self.scroll, text="No matches found",
                font=ctk.CTkFont(family="Segoe UI", size=11), text_color="#64748B"
            ).pack(pady=20)
            return

        for idx, val in enumerate(self.filtered_values):
            btn = ctk.CTkButton(
                self.scroll, text=str(val), height=28, anchor="w",
                font=ctk.CTkFont(family="Segoe UI", size=11),
                fg_color="transparent", hover_color="#1E293B", text_color="#F8FAFC",
                command=lambda v=val: self._select_item(v)
            )
            btn.pack(fill="x", pady=1)
            self._bind_mousewheel_recursive(btn)

    def _select_item(self, val):
        self.on_select_cb(val)
        self.destroy()


class ThermalTaggingDialog(ctk.CTkToplevel):
    """
    Advanced Interactive Thermal Coordinate and FLIR SDK Tagging Tool (BETA).
    """

    def __init__(self, parent, workspace_dir: str = None, on_saved_cb=None):
        super().__init__(parent)
        self.parent = parent
        self.workspace_dir = workspace_dir or getattr(parent, "target_substation_dir", "") or os.getcwd()
        self.on_saved_cb = on_saved_cb

        self.title("WRAITH // Thermal Coordinate & FLIR SDK Tagging (BETA)")
        self.geometry("1260x860")
        self.minsize(1100, 720)
        self.configure(fg_color=COLOR_CANVAS)
        self.attributes("-topmost", True)

        # State
        self.available_substations = []
        self.current_substation_idx = 0
        self.pages_catalog = []
        self.current_page_idx = 0
        self.csv_prefix = ""
        self.coords_db = {}
        self.current_coords = {}

        # Canvas & Image State
        self.current_pil_image = None
        self.display_photo = None
        self.img_scale = 1.0
        self.img_offset_x = 0
        self.img_offset_y = 0
        self.native_w = 320
        self.native_h = 240

        # Interactive Drag / Placement State
        self.active_tool = "sp1"  # "sp1", "sp2", "delta_t", "ar1_box", "ar1_poly"
        self.dragging_target = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.poly_drawing_points = []

        self._discover_substations()
        self._build_ui()
        self._bind_mousewheel_global()
        self._load_substation(self.current_substation_idx)

    def _discover_substations(self):
        """Discovers all substation CSVs in 06. EXCELL MAPPING or root."""
        base = Path(self.workspace_dir)
        map_dir = base / "06. EXCELL MAPPING"
        if not map_dir.exists():
            map_dir = base

        csv_files = sorted(list(map_dir.glob("*.csv")), key=lambda p: int(re.search(r'\d+', p.stem).group(0)) if re.search(r'\d+', p.stem) else 999)
        self.available_substations = [p.stem for p in csv_files]
        if not self.available_substations:
            self.available_substations = ["Default"]

    def _scan_substation_pages(self, csv_prefix: str):
        """Scans and loads ONLY pages that have exact physical image files on disk."""
        self.pages_catalog = []
        self.csv_prefix = csv_prefix
        base = Path(self.workspace_dir)
        map_dir = base / "06. EXCELL MAPPING"
        rm_dir = base / "03. RAW MATERIAL"
        csv_file = (map_dir / f"{csv_prefix}.csv") if map_dir.exists() else None

        self.coords_db = ThermalCoordinatesManager.load_coordinates(self.workspace_dir, self.csv_prefix)

        if csv_file and csv_file.exists():
            try:
                import pandas as pd
                try:
                    df = pd.read_csv(csv_file, encoding='utf-8')
                except Exception:
                    df = pd.read_csv(csv_file, encoding='latin1')

                if 'Output File Name' in df.columns:
                    pages_group = df.groupby('Output File Name')
                    sorted_names = sorted(
                        [p for p in pages_group.groups.keys() if str(p).lower() != "page0" and str(p).startswith("page")],
                        key=lambda x: int(re.search(r'\d+', str(x)).group(0)) if re.search(r'\d+', str(x)) else 999
                    )

                    for page_name in sorted_names:
                        g = pages_group.get_group(page_name)
                        row_ir = g[g['Attribute'].str.contains('IR', case=False, na=False)] if 'Attribute' in g.columns else None
                        row_vis = g[g['Attribute'].str.contains('Visual', case=False, na=False)] if 'Attribute' in g.columns else None

                        ir_fn = str(row_ir['Image File Name'].iloc[0]).strip() if row_ir is not None and not row_ir.empty and 'Image File Name' in row_ir.columns else ""
                        vis_fn = str(row_vis['Image File Name'].iloc[0]).strip() if row_vis is not None and not row_vis.empty and 'Image File Name' in row_vis.columns else ""

                        ir_path = None
                        if row_ir is not None and not row_ir.empty and 'Folder' in row_ir.columns:
                            cand_folder = str(row_ir['Folder'].iloc[0]).strip()
                            if cand_folder and cand_folder.lower() != 'nan':
                                cand_p = os.path.join(cand_folder, ir_fn)
                                if os.path.exists(cand_p):
                                    ir_path = Path(cand_p)

                        if not ir_path and ir_fn and rm_dir.exists():
                            for sub in rm_dir.iterdir():
                                if sub.is_dir() and (sub.name.startswith(csv_prefix) or csv_prefix.startswith(sub.name.split('.')[0])):
                                    c1 = sub / "IR" / ir_fn
                                    if c1.exists():
                                        ir_path = c1
                                        break
                                    c2 = sub / ir_fn
                                    if c2.exists():
                                        ir_path = c2
                                        break

                        if ir_path and ir_path.exists():
                            eq_desc = page_name
                            for col in g.columns:
                                if any(k in col.lower() for k in ['name', 'equipment', 'area']):
                                    v = str(g[col].iloc[0]).strip()
                                    if v and v != '-' and v.lower() != 'nan':
                                        eq_desc = v
                                        break

                            self.pages_catalog.append({
                                "page_name": str(page_name),
                                "description": eq_desc,
                                "ir_filename": ir_fn,
                                "ir_path": str(ir_path),
                                "visual_filename": vis_fn
                            })
            except Exception as e:
                print(f"[ThermalTagger] Error parsing CSV {csv_file}: {e}")

        if not self.pages_catalog and rm_dir.exists():
            for sub in rm_dir.iterdir():
                if sub.is_dir() and (sub.name.startswith(csv_prefix) or csv_prefix.startswith(sub.name.split('.')[0])):
                    ir_folder = sub / "IR"
                    if ir_folder.exists():
                        for ir_img in sorted(ir_folder.glob("*.jpg"), key=lambda x: int(re.search(r'\d+', x.name).group(0)) if re.search(r'\d+', x.name) else 999):
                            num = re.search(r'\d+', ir_img.name)
                            p_name = f"page{num.group(0)}" if num else ir_img.stem
                            self.pages_catalog.append({
                                "page_name": p_name,
                                "description": f"IR Photo: {ir_img.name}",
                                "ir_filename": ir_img.name,
                                "ir_path": str(ir_img),
                                "visual_filename": ""
                            })
                        break

    def _build_ui(self):
        # ── TOP CONTROL BAR ───────────────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=0, height=58, border_width=1, border_color=COLOR_BORDER)
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        title_box = ctk.CTkFrame(top_bar, fg_color="transparent")
        title_box.pack(side="left", padx=(14, 8))

        ctk.CTkLabel(
            title_box, text="🎯 FLIR Coordinate & Tagging",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_TEXT_MAIN
        ).pack(side="left")

        ctk.CTkLabel(
            title_box, text=" BETA ",
            font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
            fg_color=COLOR_AMBER, text_color="#000000", corner_radius=4
        ).pack(side="left", padx=6)

        # ── SUBSTATION SELECTOR ─────────────────────────────────────────────
        ctk.CTkLabel(top_bar, text="Substation 🖱️:", font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"), text_color=COLOR_CYAN).pack(side="left", padx=(8, 2))
        
        # Substation Prev
        ctk.CTkButton(
            top_bar, text="◀", width=28, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=self._prev_substation
        ).pack(side="left", padx=(0, 1))

        # Substation Dropdown Button (Opens Searchable & Scrollable Popup)
        self.sub_btn = ctk.CTkButton(
            top_bar, text=self.available_substations[0] if self.available_substations else "Select Substation ▾",
            width=135, height=30, font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            border_width=1, border_color=COLOR_BORDER,
            command=self._open_substation_popup
        )
        self.sub_btn.pack(side="left", padx=1)
        self._bind_mousewheel_widget(self.sub_btn, self._on_substation_scroll)

        # Substation Next
        ctk.CTkButton(
            top_bar, text="▶", width=28, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=self._next_substation
        ).pack(side="left", padx=(1, 6))

        # ── PAGE SELECTOR ───────────────────────────────────────────────────
        ctk.CTkLabel(top_bar, text="Page 🖱️:", font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"), text_color=COLOR_CYAN).pack(side="left", padx=(4, 2))

        # Page Prev
        ctk.CTkButton(
            top_bar, text="◀", width=28, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=self._prev_page
        ).pack(side="left", padx=(0, 1))

        # Page Dropdown Button (Opens Searchable & Scrollable Popup)
        self.page_btn = ctk.CTkButton(
            top_bar, text="Select Page ▾",
            width=210, height=30, font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            border_width=1, border_color=COLOR_BORDER,
            command=self._open_page_popup
        )
        self.page_btn.pack(side="left", padx=1)
        self._bind_mousewheel_widget(self.page_btn, self._on_page_scroll)

        # Page Next
        ctk.CTkButton(
            top_bar, text="▶", width=28, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, command=self._next_page
        ).pack(side="left", padx=(1, 4))

        # Browse Manual Image
        ctk.CTkButton(
            top_bar, text="📂 Browse Image", width=110, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER,
            text_color=COLOR_TEXT_MAIN, border_width=1, border_color=COLOR_BORDER,
            command=self._browse_custom_image
        ).pack(side="left", padx=8)

        # Top Action Buttons (Right)
        # Open DOCX button
        self.open_doc_btn = ctk.CTkButton(
            top_bar, text="📂 Open Word", width=105, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            border_width=1, border_color=COLOR_BORDER,
            command=self._open_current_docx
        )
        self.open_doc_btn.pack(side="right", padx=(4, 14))

        # Inject to Word button
        self.inject_btn = ctk.CTkButton(
            top_bar, text="⚡ Inject to Word", width=135, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color="#F59E0B", hover_color="#D97706", text_color="#000000",
            command=self._inject_current_page_to_word
        )
        self.inject_btn.pack(side="right", padx=4)

        # Save Button
        save_btn = ctk.CTkButton(
            top_bar, text="💾 Save Coords", width=115, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            command=self._save_current_coordinates
        )
        save_btn.pack(side="right", padx=4)

        # ── MAIN BODY ─────────────────────────────────────────────────────────
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=14, pady=12)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=1)

        # ── LEFT: CANVAS CARD ─────────────────────────────────────────────────
        canvas_card = ctk.CTkFrame(content, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        canvas_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # Tool selector row
        tool_bar = ctk.CTkFrame(canvas_card, fg_color="transparent", height=42)
        tool_bar.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            tool_bar, text="Active Tool:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"), text_color=COLOR_TEXT_MUTED
        ).pack(side="left", padx=(0, 6))

        self.tool_btns = {}
        tools = [
            ("sp1", "🔴 Sp1 (Spot 1)", COLOR_ROSE),
            ("sp2", "🔵 Sp2 (Spot 2)", COLOR_CYAN),
            ("delta_t", "📐 ΔT Badge", COLOR_AMBER),
            ("ar1_box", "🟩 Ar1 (Box)", COLOR_EMERALD),
            ("ar1_poly", "🔷 Ar1 (Polygon)", COLOR_VIOLET),
        ]
        for t_id, t_lbl, t_col in tools:
            btn = ctk.CTkButton(
                tool_bar, text=t_lbl, width=115, height=28,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
                fg_color=t_col if t_id == self.active_tool else COLOR_INNER,
                text_color="#000000" if t_id == self.active_tool else COLOR_TEXT_MAIN,
                border_width=1, border_color=t_col,
                command=lambda tid=t_id: self._set_active_tool(tid)
            )
            btn.pack(side="left", padx=2)
            self.tool_btns[t_id] = btn

        # Reset Poly & Clear All in Tool Bar
        ctk.CTkButton(
            tool_bar, text="✕ Reset Poly", width=78, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            fg_color=COLOR_INNER, hover_color="#2A1010", text_color=COLOR_TEXT_MUTED,
            border_width=1, border_color="#3A1515", command=self._clear_ar1
        ).pack(side="right", padx=2)

        ctk.CTkButton(
            tool_bar, text="🗑️ Clear All", width=78, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            fg_color=COLOR_INNER, hover_color="#3B0A14", text_color=COLOR_ROSE,
            border_width=1, border_color="#3B1020", command=self._clear_all_measurements
        ).pack(side="right", padx=2)

        # Image info strip
        self.img_info_lbl = ctk.CTkLabel(
            canvas_card, text="Image: Loading...",
            font=ctk.CTkFont(family=FONT_MONO, size=10), text_color=COLOR_TEXT_MUTED
        )
        self.img_info_lbl.pack(anchor="w", padx=14, pady=(0, 4))

        self.canvas = tk.Canvas(
            canvas_card, bg="#050505", highlightthickness=1,
            highlightbackground=COLOR_BORDER, cursor="crosshair"
        )
        self.canvas.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<Double-Button-1>", self._on_canvas_double_click)
        self.canvas.bind("<Button-3>", self._on_canvas_right_click)
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self._bind_mousewheel_widget(self.canvas, self._on_page_scroll)

        # ── RIGHT: PROPERTIES & TEMPERATURE CARD ──────────────────────────────
        ctrl_card = ctk.CTkFrame(content, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        ctrl_card.grid(row=0, column=1, sticky="nsew")

        header_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        header_row.pack(fill="x", padx=12, pady=(12, 4))

        ctk.CTkLabel(
            header_row, text="📊 Live Measurements",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"), text_color=COLOR_CYAN
        ).pack(side="left")

        ctk.CTkButton(
            header_row, text="🗑️ Clear Page", width=76, height=22,
            font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
            fg_color=COLOR_INNER, hover_color="#3B0A14", text_color=COLOR_ROSE,
            border_width=1, border_color="#3B1020", command=self._clear_all_measurements
        ).pack(side="right")

        ctk.CTkFrame(ctrl_card, fg_color=COLOR_BORDER, height=1).pack(fill="x", padx=12, pady=(0, 8))

        # Spot 1 Card with Clear Button
        self.sp1_box = self._create_param_input_card(ctrl_card, "🔴 Spot 1 (Sp1)", "sp1", COLOR_ROSE, on_clear=self._clear_sp1)
        self.sp2_box = self._create_param_input_card(ctrl_card, "🔵 Spot 2 (Sp2)", "sp2", COLOR_CYAN, on_clear=self._clear_sp2)
        self.ar1_box = self._create_area_input_card(ctrl_card, "🟩 Area 1 (Ar1 Avg)", "ar1", COLOR_EMERALD, on_clear=self._clear_ar1)

        # Delta T Display Card with Clear Button
        dt_card = ctk.CTkFrame(ctrl_card, fg_color=COLOR_INNER, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        dt_card.pack(fill="x", padx=12, pady=6)

        dt_top = ctk.CTkFrame(dt_card, fg_color="transparent")
        dt_top.pack(fill="x", padx=10, pady=(6, 2))

        ctk.CTkLabel(
            dt_top, text="DELTA T (ΔT = |Sp1 - Sp2|)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"), text_color=COLOR_TEXT_MUTED
        ).pack(side="left")

        ctk.CTkButton(
            dt_top, text="✕", width=22, height=18,
            font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
            fg_color=COLOR_CARD, hover_color="#2A1010", text_color=COLOR_TEXT_MUTED,
            border_width=1, border_color=COLOR_BORDER, command=self._clear_delta_t
        ).pack(side="right")

        self.delta_t_lbl = ctk.CTkLabel(
            dt_card, text="- °C",
            font=ctk.CTkFont(family=FONT_MONO, size=18, weight="bold"), text_color=COLOR_EMERALD
        )
        self.delta_t_lbl.pack(anchor="w", padx=10, pady=(0, 2))

        self.defect_status_lbl = ctk.CTkLabel(
            dt_card, text="● Status: No Anomaly (< 1.0 °C)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color=COLOR_EMERALD
        )
        self.defect_status_lbl.pack(anchor="w", padx=10, pady=(0, 6))

        # Presets selector
        preset_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        preset_row.pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(preset_row, text="Preset:", font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color=COLOR_TEXT_MUTED).pack(side="left")
        self.preset_combo = ctk.CTkComboBox(
            preset_row, values=["Cable Compartment", "Transformer Bushing", "Overview Center Box", "3-Phase Spots"],
            height=28, font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            fg_color=COLOR_INNER, border_color=COLOR_BORDER, command=self._apply_preset
        )
        self.preset_combo.pack(side="right", fill="x", expand=True, padx=(6, 0))

        # Action Buttons
        btn_frame = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=12, pady=(8, 12), side="bottom")

        ctk.CTkButton(
            btn_frame, text="Save & Next Page ⏩", height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=COLOR_CYAN, hover_color=COLOR_CYAN_HOVER, text_color="#000000",
            command=self._save_and_next
        ).pack(fill="x", pady=2)

        ctk.CTkButton(
            btn_frame, text="⚡ Apply to All Pages", height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            fg_color=COLOR_INNER, hover_color=COLOR_CARD_HOVER, text_color=COLOR_TEXT_MAIN,
            border_width=1, border_color=COLOR_BORDER, command=self._apply_to_all_pages
        ).pack(fill="x", pady=2)

    def _create_param_input_card(self, parent, title: str, key: str, color: str, on_clear=None):
        card = ctk.CTkFrame(parent, fg_color=COLOR_INNER, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=12, pady=2)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(4, 2))

        ctk.CTkLabel(top, text=title, font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"), text_color=color).pack(side="left")

        if on_clear:
            ctk.CTkButton(
                top, text="✕", width=22, height=18,
                font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
                fg_color=COLOR_CARD, hover_color="#2A1010", text_color=COLOR_TEXT_MUTED,
                border_width=1, border_color=COLOR_BORDER, command=on_clear
            ).pack(side="right", padx=(4, 0))

        coord_lbl = ctk.CTkLabel(top, text="X: 0 | Y: 0", font=ctk.CTkFont(family=FONT_MONO, size=10), text_color=COLOR_TEXT_MUTED)
        coord_lbl.pack(side="right")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=(2, 4))

        ctk.CTkLabel(row, text="Temp (°C):", font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color=COLOR_TEXT_MUTED).pack(side="left", padx=(0, 6))
        entry = ctk.CTkEntry(row, width=80, height=24, font=ctk.CTkFont(family=FONT_MONO, size=11), fg_color=COLOR_CARD, border_color=COLOR_BORDER)
        entry.pack(side="left")
        entry.bind("<KeyRelease>", lambda e: self._recalculate_delta_t())

        return {"card": card, "coord_lbl": coord_lbl, "temp_entry": entry}

    def _create_area_input_card(self, parent, title: str, key: str, color: str, on_clear=None):
        card = ctk.CTkFrame(parent, fg_color=COLOR_INNER, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=12, pady=2)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(4, 2))

        ctk.CTkLabel(top, text=title, font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"), text_color=color).pack(side="left")

        if on_clear:
            ctk.CTkButton(
                top, text="✕", width=22, height=18,
                font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
                fg_color=COLOR_CARD, hover_color="#2A1010", text_color=COLOR_TEXT_MUTED,
                border_width=1, border_color=COLOR_BORDER, command=on_clear
            ).pack(side="right", padx=(4, 0))

        coord_lbl = ctk.CTkLabel(top, text="[0, 0, 0, 0]", font=ctk.CTkFont(family=FONT_MONO, size=9), text_color=COLOR_TEXT_MUTED)
        coord_lbl.pack(side="right")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=(2, 4))

        ctk.CTkLabel(row, text="Avg Temp (°C):", font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color=COLOR_TEXT_MUTED).pack(side="left", padx=(0, 6))
        entry = ctk.CTkEntry(row, width=80, height=24, font=ctk.CTkFont(family=FONT_MONO, size=11), fg_color=COLOR_CARD, border_color=COLOR_BORDER)
        entry.pack(side="left")

        return {"card": card, "coord_lbl": coord_lbl, "temp_entry": entry}

    # ── MOUSE WHEEL SCROLLING HANDLERS ───────────────────────────────────────
    def _bind_mousewheel_widget(self, widget, callback):
        widget.bind("<MouseWheel>", callback)
        for child in getattr(widget, "winfo_children", lambda: [])():
            child.bind("<MouseWheel>", callback)

    def _bind_mousewheel_global(self):
        self.bind("<MouseWheel>", self._on_global_scroll)

    def _on_global_scroll(self, event):
        # If hovering over top bar or canvas, route to page scroll
        widget = event.widget
        if widget == self.sub_combo or str(widget).endswith("sub_combo"):
            self._on_substation_scroll(event)
        else:
            self._on_page_scroll(event)

    def _on_substation_scroll(self, event):
        if not self.available_substations:
            return
        delta = 1 if event.delta < 0 else -1
        new_idx = max(0, min(len(self.available_substations) - 1, self.current_substation_idx + delta))
        if new_idx != self.current_substation_idx:
            self._load_substation(new_idx)

    def _on_page_scroll(self, event):
        if not self.pages_catalog:
            return
        if event.delta < 0:
            self._next_page()
        else:
            self._prev_page()

    # ── CLEAR MEASUREMENT HANDLERS ───────────────────────────────────────────
    def _clear_sp1(self):
        if "sp1" in self.current_coords:
            self.current_coords["sp1"] = {"x": 0, "y": 0, "enabled": False, "label": "Sp1", "temp": ""}
        self.sp1_box["temp_entry"].delete(0, "end")
        self._update_ui_inputs()
        self._draw_canvas()

    def _clear_sp2(self):
        if "sp2" in self.current_coords:
            self.current_coords["sp2"] = {"x": 0, "y": 0, "enabled": False, "label": "Sp2", "temp": ""}
        self.sp2_box["temp_entry"].delete(0, "end")
        self._update_ui_inputs()
        self._draw_canvas()

    def _clear_ar1(self):
        self.poly_drawing_points = []
        if "ar1" in self.current_coords:
            self.current_coords["ar1"] = {
                "type": "box", "x": 0, "y": 0, "w": 0, "h": 0,
                "points": [], "enabled": False, "label": "Ar1", "temp_avg": ""
            }
        self.ar1_box["temp_entry"].delete(0, "end")
        self._update_ui_inputs()
        self._draw_canvas()

    def _clear_delta_t(self):
        if "delta_t_tag" in self.current_coords:
            self.current_coords["delta_t_tag"]["enabled"] = False
            self.current_coords["delta_t_tag"]["val"] = ""
        self.current_coords["delta_t"] = ""
        self.delta_t_lbl.configure(text="- °C", text_color=COLOR_TEXT_MUTED)
        self.defect_status_lbl.configure(text="● Status: -", text_color=COLOR_TEXT_MUTED)
        self._draw_canvas()

    def _clear_all_measurements(self):
        self._clear_sp1()
        self._clear_sp2()
        self._clear_ar1()
        self._clear_delta_t()

    def _on_canvas_right_click(self, event):
        """Right click menu to quickly delete tag under cursor."""
        menu = tk.Menu(self, tearoff=0, bg="#111111", fg="#FFFFFF", activebackground=COLOR_CYAN, activeforeground="#000000")
        menu.add_command(label="🔴 Clear Spot 1 (Sp1)", command=self._clear_sp1)
        menu.add_command(label="🔵 Clear Spot 2 (Sp2)", command=self._clear_sp2)
        menu.add_command(label="🟩 Clear Area 1 (Ar1)", command=self._clear_ar1)
        menu.add_command(label="📐 Clear Delta T Badge", command=self._clear_delta_t)
        menu.add_separator()
        menu.add_command(label="🗑️ Clear All Tags", command=self._clear_all_measurements)
        menu.tk_popup(event.x_root, event.y_root)

    # ── TOOL & RENDERING ─────────────────────────────────────────────────────
    def _set_active_tool(self, tool_id: str):
        self.active_tool = tool_id
        for tid, btn in self.tool_btns.items():
            if tid == tool_id:
                col = COLOR_ROSE if tid == "sp1" else COLOR_CYAN if tid == "sp2" else COLOR_AMBER if tid == "delta_t" else COLOR_EMERALD if tid == "ar1_box" else COLOR_VIOLET
                btn.configure(fg_color=col, text_color="#000000")
            else:
                btn.configure(fg_color=COLOR_INNER, text_color=COLOR_TEXT_MAIN)

    def _load_substation(self, sub_idx: int):
        if not self.available_substations:
            return
        self.current_substation_idx = sub_idx
        sub_name = self.available_substations[sub_idx]
        self.sub_btn.configure(text=f'{sub_name} ▾') if hasattr(self, 'sub_btn') else None

        self._scan_substation_pages(sub_name)

        page_options = [f"{p['page_name']} • {p['description'][:20]}" for p in self.pages_catalog] or ["No Pages Found"]
        
        self.current_page_idx = 0
        if self.pages_catalog:
            self._load_current_page()
        else:
            self.canvas.delete("all")
            self.img_info_lbl.configure(text="No IR images found for this substation.")

    def _load_current_page(self):
        if not self.pages_catalog or self.current_page_idx >= len(self.pages_catalog):
            return

        page_data = self.pages_catalog[self.current_page_idx]
        p_name = page_data["page_name"]
        self.page_btn.configure(text=f"{p_name} • {page_data['description'][:22]} ▾") if hasattr(self, 'page_btn') else None

        if p_name in self.coords_db:
            self.current_coords = self.coords_db[p_name].copy()
        else:
            self.current_coords = ThermalCoordinatesManager.get_default_coords_for_type(
                page_data["description"], page_data["description"]
            )

        if "ar1" in self.current_coords:
            if "points" not in self.current_coords["ar1"]:
                self.current_coords["ar1"]["points"] = []
            if "type" not in self.current_coords["ar1"]:
                self.current_coords["ar1"]["type"] = "box"

        self.poly_drawing_points = list(self.current_coords.get("ar1", {}).get("points", []))

        ir_path = page_data.get("ir_path", "")
        if ir_path and os.path.exists(ir_path):
            try:
                self.current_pil_image = Image.open(ir_path)
                self.native_w, self.native_h = self.current_pil_image.size
                self.img_info_lbl.configure(text=f"Image: {os.path.basename(ir_path)} ({self.native_w}x{self.native_h}) | Path: {ir_path}")
            except Exception as e:
                self.img_info_lbl.configure(text=f"Error loading image: {e}")
                self.current_pil_image = self._create_placeholder_thermal_image()
                self.native_w, self.native_h = 320, 240
        else:
            self.img_info_lbl.configure(text=f"Missing file: {page_data.get('ir_filename', '')}")
            self.current_pil_image = self._create_placeholder_thermal_image()
            self.native_w, self.native_h = 320, 240

        self._update_ui_inputs()
        self._draw_canvas()

    def _browse_custom_image(self):
        file_path = filedialog.askopenfilename(
            title="Select Exact Thermal IR Image",
            filetypes=[("FLIR Thermal Image", "*.jpg;*.jpeg;*.png"), ("All Files", "*.*")]
        )
        if file_path and os.path.exists(file_path):
            try:
                self.current_pil_image = Image.open(file_path)
                self.native_w, self.native_h = self.current_pil_image.size
                if self.pages_catalog and self.current_page_idx < len(self.pages_catalog):
                    self.pages_catalog[self.current_page_idx]["ir_path"] = file_path
                    self.pages_catalog[self.current_page_idx]["ir_filename"] = os.path.basename(file_path)
                self.img_info_lbl.configure(text=f"Custom: {os.path.basename(file_path)} ({self.native_w}x{self.native_h})")
                self._draw_canvas()
            except Exception as e:
                messagebox.showerror("Image Error", f"Could not load image: {e}")

    def _create_placeholder_thermal_image(self) -> Image.Image:
        img = Image.new("RGB", (320, 240), color=(20, 10, 40))
        draw = ImageDraw.Draw(img)
        for y in range(240):
            r = int(255 * (y / 240.0))
            b = int(255 * (1.0 - y / 240.0))
            g = int(128 * math.sin(y / 40.0))
            draw.line([(0, y), (320, y)], fill=(r, max(0, g), b))
        draw.text((80, 110), "FLIR IR Thermal Preview", fill=(255, 255, 255))
        return img

    def _update_ui_inputs(self):
        sp1 = self.current_coords.get("sp1", {})
        sp2 = self.current_coords.get("sp2", {})
        ar1 = self.current_coords.get("ar1", {})

        if sp1.get("enabled", True) and (sp1.get("x", 0) > 0 or sp1.get("y", 0) > 0):
            self.sp1_box["coord_lbl"].configure(text=f"X: {sp1.get('x', 0)} | Y: {sp1.get('y', 0)}")
        else:
            self.sp1_box["coord_lbl"].configure(text="None")

        if sp2.get("enabled", True) and (sp2.get("x", 0) > 0 or sp2.get("y", 0) > 0):
            self.sp2_box["coord_lbl"].configure(text=f"X: {sp2.get('x', 0)} | Y: {sp2.get('y', 0)}")
        else:
            self.sp2_box["coord_lbl"].configure(text="None")

        if ar1.get("enabled", True):
            if ar1.get("type") == "polygon" and ar1.get("points"):
                self.ar1_box["coord_lbl"].configure(text=f"Poly ({len(ar1['points'])} pts)")
            elif ar1.get("w", 0) > 0:
                self.ar1_box["coord_lbl"].configure(text=f"Box [{ar1.get('x', 0)}, {ar1.get('y', 0)}, {ar1.get('w', 0)}, {ar1.get('h', 0)}]")
            else:
                self.ar1_box["coord_lbl"].configure(text="None")
        else:
            self.ar1_box["coord_lbl"].configure(text="None")

        self.sp1_box["temp_entry"].delete(0, "end")
        if sp1.get("temp"):
            self.sp1_box["temp_entry"].insert(0, str(sp1.get("temp")))

        self.sp2_box["temp_entry"].delete(0, "end")
        if sp2.get("temp"):
            self.sp2_box["temp_entry"].insert(0, str(sp2.get("temp")))

        self.ar1_box["temp_entry"].delete(0, "end")
        if ar1.get("temp_avg"):
            self.ar1_box["temp_entry"].insert(0, str(ar1.get("temp_avg")))

        self._recalculate_delta_t()

    def _recalculate_delta_t(self):
        t1_str = self.sp1_box["temp_entry"].get().strip()
        t2_str = self.sp2_box["temp_entry"].get().strip()

        try:
            t1 = float(t1_str)
            t2 = float(t2_str)
            dt = round(abs(t1 - t2), 1)
            self.delta_t_lbl.configure(text=f"{dt:.1f} °C")
            if dt < 1.0:
                self.delta_t_lbl.configure(text_color=COLOR_EMERALD)
                self.defect_status_lbl.configure(text="● Status: No Anomaly (< 1.0 °C)", text_color=COLOR_EMERALD)
            elif 1.0 <= dt < 5.0:
                self.delta_t_lbl.configure(text_color=COLOR_AMBER)
                self.defect_status_lbl.configure(text="● Status: Minor Heating (1.0 - 5.0 °C)", text_color=COLOR_AMBER)
            else:
                self.delta_t_lbl.configure(text_color=COLOR_ROSE)
                self.defect_status_lbl.configure(text="● Status: Hotspot / Defect (> 5.0 °C)", text_color=COLOR_ROSE)
        except ValueError:
            self.delta_t_lbl.configure(text="- °C", text_color=COLOR_TEXT_MUTED)
            self.defect_status_lbl.configure(text="● Status: -", text_color=COLOR_TEXT_MUTED)

        self._draw_canvas()

    def _draw_canvas(self):
        self.canvas.delete("all")
        if not self.current_pil_image:
            return

        cw = self.canvas.winfo_width() or 640
        ch = self.canvas.winfo_height() or 480

        scale_w = (cw - 40) / float(self.native_w)
        scale_h = (ch - 40) / float(self.native_h)
        self.img_scale = min(scale_w, scale_h, 3.5)
        self.img_scale = max(self.img_scale, 0.5)

        disp_w = int(self.native_w * self.img_scale)
        disp_h = int(self.native_h * self.img_scale)

        self.img_offset_x = (cw - disp_w) // 2
        self.img_offset_y = (ch - disp_h) // 2

        resized_pil = self.current_pil_image.resize((disp_w, disp_h), Image.Resampling.BILINEAR)
        self.display_photo = ImageTk.PhotoImage(resized_pil)
        self.canvas.create_image(self.img_offset_x, self.img_offset_y, anchor="nw", image=self.display_photo)

        self.canvas.create_rectangle(
            self.img_offset_x, self.img_offset_y,
            self.img_offset_x + disp_w, self.img_offset_y + disp_h,
            outline="#333333", width=1
        )

        # ── 1. Draw Ar1 Area (Box or Polygon) ─────────────────────────────────
        ar1 = self.current_coords.get("ar1", {})
        if ar1.get("enabled", True):
            poly_pts = ar1.get("points", []) or self.poly_drawing_points

            if ar1.get("type") == "polygon" and len(poly_pts) >= 2:
                canvas_pts = []
                for pt in poly_pts:
                    px = self.img_offset_x + int(pt["x"] * self.img_scale)
                    py = self.img_offset_y + int(pt["y"] * self.img_scale)
                    canvas_pts.extend([px, py])
                    self.canvas.create_oval(px - 4, py - 4, px + 4, py + 4, fill=COLOR_VIOLET, outline="#FFFFFF", width=1, tags="poly_handle")

                if len(poly_pts) >= 3:
                    self.canvas.create_polygon(canvas_pts, outline=COLOR_VIOLET, fill="", width=2, dash=(4, 2), tags="ar1_poly")
                    first_x = self.img_offset_x + int(poly_pts[0]["x"] * self.img_scale)
                    first_y = self.img_offset_y + int(poly_pts[0]["y"] * self.img_scale)
                    self.canvas.create_text(first_x + 6, first_y - 12, text="Ar1 (Polygon)", fill=COLOR_VIOLET, font=(FONT_FAMILY, 9, "bold"), anchor="nw")
                else:
                    self.canvas.create_line(canvas_pts, fill=COLOR_VIOLET, width=2, tags="ar1_poly_line")
            elif ar1 and ar1.get("w", 0) > 0:
                rx1 = self.img_offset_x + int(ar1["x"] * self.img_scale)
                ry1 = self.img_offset_y + int(ar1["y"] * self.img_scale)
                rx2 = rx1 + int(ar1["w"] * self.img_scale)
                ry2 = ry1 + int(ar1["h"] * self.img_scale)

                self.canvas.create_rectangle(rx1, ry1, rx2, ry2, outline=COLOR_EMERALD, width=2, dash=(4, 2), tags="ar1_box")
                self.canvas.create_text(rx1 + 4, ry1 - 12, text="Ar1 (Area)", fill=COLOR_EMERALD, font=(FONT_FAMILY, 9, "bold"), anchor="nw")

        # ── 2. Draw Sp1 & Sp2 ─────────────────────────────────────────────────
        sp1 = self.current_coords.get("sp1", {})
        sp2 = self.current_coords.get("sp2", {})

        sx1, sy1 = None, None
        sx2, sy2 = None, None

        if sp1 and sp1.get("enabled", True) and (sp1.get("x", 0) > 0 or sp1.get("y", 0) > 0):
            sx1 = self.img_offset_x + int(sp1.get("x", 0) * self.img_scale)
            sy1 = self.img_offset_y + int(sp1.get("y", 0) * self.img_scale)
            t1_val = str(sp1.get("temp", ""))
            lbl1 = f"Sp1: {t1_val} °C" if t1_val else "Sp1"
            self._draw_crosshair(sx1, sy1, COLOR_ROSE, lbl1, "sp1_pin")

        if sp2 and sp2.get("enabled", True) and (sp2.get("x", 0) > 0 or sp2.get("y", 0) > 0):
            sx2 = self.img_offset_x + int(sp2.get("x", 0) * self.img_scale)
            sy2 = self.img_offset_y + int(sp2.get("y", 0) * self.img_scale)
            t2_val = str(sp2.get("temp", ""))
            lbl2 = f"Sp2: {t2_val} °C" if t2_val else "Sp2"
            self._draw_crosshair(sx2, sy2, COLOR_CYAN, lbl2, "sp2_pin")

        # ── 3. Draw Delta T Line & Badge ──────────────────────────────────────
        if sx1 is not None and sx2 is not None:
            self.canvas.create_line(sx1, sy1, sx2, sy2, fill=COLOR_AMBER, width=1, dash=(3, 3), tags="dt_line")

            dt_pos = self.current_coords.get("delta_t_tag", {})
            if dt_pos and dt_pos.get("x"):
                dx = self.img_offset_x + int(dt_pos["x"] * self.img_scale)
                dy = self.img_offset_y + int(dt_pos["y"] * self.img_scale)
            else:
                dx = (sx1 + sx2) // 2
                dy = (sy1 + sy2) // 2 - 14

            dt_text = self.delta_t_lbl.cget("text")
            dt_badge_text = f"ΔT: {dt_text}" if dt_text != "- °C" else "ΔT Tag"
            self._draw_delta_t_badge(dx, dy, dt_badge_text)

    def _draw_crosshair(self, cx: int, cy: int, color: str, label: str, tag: str):
        size = 14
        self.canvas.create_oval(cx - size, cy - size, cx + size, cy + size, outline=color, width=2, tags=tag)
        self.canvas.create_line(cx - size - 4, cy, cx + size + 4, cy, fill=color, width=2, tags=tag)
        self.canvas.create_line(cx, cy - size - 4, cx, cy + size + 4, fill=color, width=2, tags=tag)
        w_box = max(40, len(label) * 7 + 10)
        self.canvas.create_rectangle(cx + 10, cy - 18, cx + 10 + w_box, cy - 2, fill="#0A0A0A", outline=color, width=1, tags=tag)
        self.canvas.create_text(cx + 14, cy - 10, text=label, fill=color, font=(FONT_FAMILY, 9, "bold"), anchor="w", tags=tag)

    def _draw_delta_t_badge(self, bx: int, by: int, text: str):
        w_box = max(56, len(text) * 7 + 14)
        h_box = 20
        self.canvas.create_rectangle(bx - w_box // 2, by - h_box // 2, bx + w_box // 2, by + h_box // 2, fill="#0A0A0A", outline=COLOR_AMBER, width=2, tags="dt_badge")
        self.canvas.create_text(bx, by, text=text, fill=COLOR_AMBER, font=(FONT_FAMILY, 9, "bold"), anchor="center", tags="dt_badge")

    def _canvas_to_native(self, cx: int, cy: int) -> tuple[int, int]:
        nx = int((cx - self.img_offset_x) / self.img_scale)
        ny = int((cy - self.img_offset_y) / self.img_scale)
        nx = max(0, min(self.native_w, nx))
        ny = max(0, min(self.native_h, ny))
        return nx, ny

    # ── CANVAS INTERACTION HANDLERS ──────────────────────────────────────────
    def _on_canvas_press(self, event):
        nx, ny = self._canvas_to_native(event.x, event.y)
        self.drag_start_x = nx
        self.drag_start_y = ny

        if self.active_tool == "sp1":
            self.current_coords.setdefault("sp1", {})["x"] = nx
            self.current_coords["sp1"]["y"] = ny
            self.current_coords["sp1"]["enabled"] = True
            self.dragging_target = "sp1"
        elif self.active_tool == "sp2":
            self.current_coords.setdefault("sp2", {})["x"] = nx
            self.current_coords["sp2"]["y"] = ny
            self.current_coords["sp2"]["enabled"] = True
            self.dragging_target = "sp2"
        elif self.active_tool == "delta_t":
            self.current_coords.setdefault("delta_t_tag", {})["x"] = nx
            self.current_coords["delta_t_tag"]["y"] = ny
            self.current_coords["delta_t_tag"]["enabled"] = True
            self.dragging_target = "delta_t_tag"
        elif self.active_tool == "ar1_box":
            self.current_coords.setdefault("ar1", {})["type"] = "box"
            self.current_coords["ar1"]["x"] = nx
            self.current_coords["ar1"]["y"] = ny
            self.current_coords["ar1"]["w"] = 0
            self.current_coords["ar1"]["h"] = 0
            self.current_coords["ar1"]["enabled"] = True
            self.dragging_target = "ar1_box"
        elif self.active_tool == "ar1_poly":
            self.poly_drawing_points.append({"x": nx, "y": ny})
            self.current_coords.setdefault("ar1", {})["type"] = "polygon"
            self.current_coords["ar1"]["points"] = list(self.poly_drawing_points)
            self.current_coords["ar1"]["enabled"] = True
            self.dragging_target = None

        self._update_ui_inputs()
        self._draw_canvas()

    def _on_canvas_drag(self, event):
        if not self.dragging_target:
            return

        nx, ny = self._canvas_to_native(event.x, event.y)

        if self.dragging_target in ["sp1", "sp2", "delta_t_tag"]:
            self.current_coords[self.dragging_target]["x"] = nx
            self.current_coords[self.dragging_target]["y"] = ny
        elif self.dragging_target == "ar1_box":
            x0 = min(self.drag_start_x, nx)
            y0 = min(self.drag_start_y, ny)
            w = abs(nx - self.drag_start_x)
            h = abs(ny - self.drag_start_y)
            self.current_coords["ar1"]["x"] = x0
            self.current_coords["ar1"]["y"] = y0
            self.current_coords["ar1"]["w"] = w
            self.current_coords["ar1"]["h"] = h

        self._update_ui_inputs()
        self._draw_canvas()

    def _on_canvas_release(self, event):
        self.dragging_target = None
        self._update_ui_inputs()
        self._draw_canvas()

    def _on_canvas_double_click(self, event):
        if self.active_tool == "ar1_poly" and len(self.poly_drawing_points) >= 3:
            self.current_coords.setdefault("ar1", {})["type"] = "polygon"
            self.current_coords["ar1"]["points"] = list(self.poly_drawing_points)
            self._update_ui_inputs()
            self._draw_canvas()

    def _on_canvas_resize(self, event):
        self._draw_canvas()

    # ── NAVIGATION & SUBSTATION HANDLERS ──────────────────────────────────────
    def _on_substation_selected(self, choice: str):
        if choice in self.available_substations:
            idx = self.available_substations.index(choice)
            self._load_substation(idx)

    def _on_page_selected(self, choice: str):
        p_name = choice.split(" • ")[0].strip()
        for i, p in enumerate(self.pages_catalog):
            if p["page_name"] == p_name:
                self.current_page_idx = i
                self._load_current_page()
                break

    def _prev_page(self):
        if self.current_page_idx > 0:
            self.current_page_idx -= 1
            self._load_current_page()

    def _next_page(self):
        if self.current_page_idx < len(self.pages_catalog) - 1:
            self.current_page_idx += 1
            self._load_current_page()

    def _apply_preset(self, choice: str):
        self.current_coords = ThermalCoordinatesManager.get_default_coords_for_type(
            choice, choice, self.native_w, self.native_h
        )
        self.poly_drawing_points = []
        self._update_ui_inputs()
        self._draw_canvas()

    def _save_current_coordinates(self):
        if not self.pages_catalog or self.current_page_idx >= len(self.pages_catalog):
            return

        p_name = self.pages_catalog[self.current_page_idx]["page_name"]

        t1 = self.sp1_box["temp_entry"].get().strip()
        t2 = self.sp2_box["temp_entry"].get().strip()
        ta = self.ar1_box["temp_entry"].get().strip()

        if "sp1" in self.current_coords:
            self.current_coords["sp1"]["temp"] = t1
        if "sp2" in self.current_coords:
            self.current_coords["sp2"]["temp"] = t2
        if "ar1" in self.current_coords:
            self.current_coords["ar1"]["temp_avg"] = ta

        try:
            if t1 and t2:
                dt_calc = round(abs(float(t1) - float(t2)), 1)
                self.current_coords["delta_t"] = dt_calc
                if "delta_t_tag" in self.current_coords:
                    self.current_coords["delta_t_tag"]["val"] = str(dt_calc)
        except ValueError:
            pass

        self.coords_db[p_name] = self.current_coords.copy()
        success = ThermalCoordinatesManager.save_coordinates(self.workspace_dir, self.coords_db, self.csv_prefix)
        if success and self.on_saved_cb:
            self.on_saved_cb(self.coords_db)

    def _save_and_next(self):
        self._save_current_coordinates()
        self._next_page()

    def _apply_to_all_pages(self):
        if not messagebox.askyesno("Confirm Apply to All", f"Apply these Sp1, Sp2, and Ar1 coordinates to ALL {len(self.pages_catalog)} pages in substation {self.csv_prefix}?"):
            return

        for p in self.pages_catalog:
            p_name = p["page_name"]
            self.coords_db[p_name] = self.current_coords.copy()

        ThermalCoordinatesManager.save_coordinates(self.workspace_dir, self.coords_db, self.csv_prefix)
        messagebox.showinfo("Success", f"Coordinates successfully applied to all {len(self.pages_catalog)} pages!")

    def _inject_current_page_to_word(self):
        """Immediately saves coordinates and runs FLIR injection for the current page, then opens the .docx."""
        if not self.pages_catalog or self.current_page_idx >= len(self.pages_catalog):
            messagebox.showwarning("No Page", "Please select a valid page to inject.")
            return

        # 1. Save current page coordinates & temperatures
        self._save_current_coordinates()

        page_data = self.pages_catalog[self.current_page_idx]
        p_name = page_data["page_name"]
        sub_prefix = self.csv_prefix

        if hasattr(self, "inject_btn"):
            self.inject_btn.configure(text="⏳ Injecting...", state="disabled")

        def run_injection_thread():
            try:
                base = Path(self.workspace_dir)
                map_dir = base / "06. EXCELL MAPPING"
                out_dir = base / "07. OUTPUT FILE" / "INDIVIDUAL"
                csv_file = map_dir / f"{sub_prefix}.csv"

                if not csv_file.exists():
                    for cand in map_dir.glob("*.csv"):
                        if cand.stem == sub_prefix or sub_prefix.startswith(cand.stem):
                            csv_file = cand
                            break

                if not csv_file.exists():
                    self.after(0, lambda: [
                        self.inject_btn.configure(text="⚡ Inject to Word", state="normal") if hasattr(self, "inject_btn") else None,
                        messagebox.showerror("CSV Not Found", f"Could not find CSV for substation {sub_prefix} in {map_dir}")
                    ])
                    return

                # Set required environment variables for the report generator
                os.environ["BASE_DIR"] = str(base)
                os.environ["MAPPING_DIR"] = str(map_dir)
                os.environ["TEMPLATE_DIR"] = str(base / "05. TEMPLATE")
                os.environ["OUTPUT_DIR"] = str(base / "07. OUTPUT FILE")
                os.environ["RM_DIR"] = str(base / "03. RAW MATERIAL")

                # In-process direct execution (works in both standalone exe and dev mode)
                from generate_report import process_csv
                process_csv(str(csv_file), target_pages_input=str(p_name), no_merge=True)

                expected_docx = out_dir / f"{sub_prefix}. {p_name}.docx"

                def on_done():
                    if hasattr(self, "inject_btn"):
                        self.inject_btn.configure(text="⚡ Inject to Word", state="normal")

                    if expected_docx.exists():
                        try:
                            os.startfile(str(expected_docx))
                        except Exception as open_err:
                            print(f"[ThermalTagger] Could not startfile: {open_err}")
                        messagebox.showinfo(
                            "Injection Complete",
                            f"🎉 Page {p_name} successfully injected with FLIR tags and thermal table measurements!\n\nSaved to:\n{expected_docx.name}\n\nOpened in Microsoft Word."
                        )
                    else:
                        messagebox.showerror("Injection Error", f"FLIR injection failed for {p_name}. File not found:\n{expected_docx}")

                self.after(0, on_done)
            except Exception as ex:
                err_str = str(ex)
                self.after(0, lambda: [
                    self.inject_btn.configure(text="⚡ Inject to Word", state="normal") if hasattr(self, "inject_btn") else None,
                    messagebox.showerror("Error", f"Could not complete injection: {err_str}")
                ])

        threading.Thread(target=run_injection_thread, daemon=True).start()

    def _open_current_docx(self):
        """Opens the rendered docx file for the current page if it exists."""
        if not self.pages_catalog or self.current_page_idx >= len(self.pages_catalog):
            return
        p_name = self.pages_catalog[self.current_page_idx]["page_name"]
        out_dir = Path(self.workspace_dir) / "07. OUTPUT FILE" / "INDIVIDUAL"
        expected_docx = out_dir / f"{self.csv_prefix}. {p_name}.docx"
        if expected_docx.exists():
            try:
                os.startfile(str(expected_docx))
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")
        else:
            messagebox.showinfo("File Not Generated", f"Word file has not been generated yet for {p_name}.\n\nClick '⚡ Inject to Word' to generate it now.")

    def _open_substation_popup(self):
        """Opens searchable, scrollable dropdown popup for substations."""
        if not self.available_substations:
            return
        ScrollableSearchPopup(
            self, self.sub_btn, self.available_substations,
            on_select_cb=self._on_substation_selected,
            width=280, height=420, placeholder="🔍 Search Substation (e.g. 301)..."
        )

    def _open_page_popup(self):
        """Opens searchable, scrollable dropdown popup for pages."""
        if not self.pages_catalog:
            return
        page_options = [f"{p['page_name']} • {p['description'][:25]}" for p in self.pages_catalog]
        ScrollableSearchPopup(
            self, self.page_btn, page_options,
            on_select_cb=self._on_page_selected,
            width=360, height=420, placeholder="🔍 Search Page (e.g. page2, cable)..."
        )

    def _prev_substation(self):
        if self.current_substation_idx > 0:
            self._load_substation(self.current_substation_idx - 1)

    def _next_substation(self):
        if self.current_substation_idx < len(self.available_substations) - 1:
            self._load_substation(self.current_substation_idx + 1)


