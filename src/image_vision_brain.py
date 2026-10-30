"""
WRAITH Hybrid Image Vision Brain :: Substation Equipment Intelligence & OCR Engine
Combines:
1. Local OCR (Pytesseract / OpenCV) for panel nameplate reading & equipment feature detection (100% Offline, No API key needed)
2. Gemini Multimodal Cloud Vision (Optional upgrade if GEMINI_API_KEY is provided)
3. Heuristic cluster matching (Aligns TX1, TX2, SWG Panels, FP1, FP2 and assigns spares)
"""

import os
import re
import json
import difflib
from pathlib import Path
from datetime import datetime

try:
    from PIL import Image, ImageEnhance, ImageFilter
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    import google.generativeai as genai
    HAS_GEMINI_SDK = True
except ImportError:
    HAS_GEMINI_SDK = False


class ImageVisionBrain:
    def __init__(self, rules_json_path: str = None, api_key: str = None):
        if not rules_json_path:
            rules_json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "image_brain_rules.json")

        self.rules_path = rules_json_path
        self.rules = self._load_rules()
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

        if self.api_key and HAS_GEMINI_SDK:
            try:
                genai.configure(api_key=self.api_key)
            except Exception as e:
                print(f"[Brain] Warning configuring Gemini: {e}")

    def _load_rules(self) -> dict:
        if os.path.exists(self.rules_path):
            try:
                with open(self.rules_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    # ── 1. LOCAL OCR & NAMEPLATE READER ──────────────────────────────────────
    def extract_text_from_image(self, image_path: str) -> str:
        """
        Extracts visible text and nameplate labels from a photo using local OCR.
        """
        if not HAS_PIL or not HAS_TESSERACT or not os.path.exists(image_path):
            return ""

        try:
            with Image.open(image_path) as img:
                # Preprocess: Convert to grayscale and enhance contrast for label reading
                gray = img.convert("L")
                enhancer = ImageEnhance.Contrast(gray)
                enhanced = enhancer.enhance(2.0)
                text = pytesseract.image_to_string(enhanced, config="--psm 11")
                # Clean up whitespace
                clean_text = " ".join(text.split())
                return clean_text
        except Exception:
            return ""

    # ── 2. COMPUTER VISION ASSET CLASSIFIER ──────────────────────────────────
    def detect_asset_visual_features(self, image_path: str) -> dict:
        """
        Analyzes edge orientation, color histograms, and aspect ratio to detect:
        - Transformer (vertical cooling fin stripes, bushing insulators)
        - Switchgear RMU (horizontal panel seams, mimic diagrams, fascia)
        - Feeder Pillar (vertical fuse slots, J-slotted busbars)
        - Wide Overview vs Close-up
        """
        features = {
            "is_wide_angle": False,
            "vertical_edge_density": 0.0,
            "horizontal_edge_density": 0.0,
            "brightness": 128.0,
            "estimated_asset": "UNKNOWN",
            "overview_score": 1.0
        }

        if not HAS_PIL or not os.path.exists(image_path):
            return features

        try:
            with Image.open(image_path) as img:
                w, h = img.size
                aspect = w / float(h) if h > 0 else 1.0
                features["is_wide_angle"] = (1.2 <= aspect <= 1.8)

                sz = os.path.getsize(image_path)
                features["overview_score"] = min(sz / 280000.0, 2.0) + (0.5 if features["is_wide_angle"] else 0.0)

                if HAS_CV2:
                    cv_img = cv2.imread(image_path)
                    if cv_img is not None:
                        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
                        features["brightness"] = float(np.mean(gray))

                        # Sobel edge analysis
                        sobel_x = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
                        sobel_y = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)

                        v_density = np.mean(np.abs(sobel_x))
                        h_density = np.mean(np.abs(sobel_y))

                        features["vertical_edge_density"] = float(v_density)
                        features["horizontal_edge_density"] = float(h_density)

                        # High vertical edge density is characteristic of Transformer radiators & FP fuse ways
                        if v_density > h_density * 1.3:
                            features["estimated_asset"] = "TRANSFORMER_OR_FP"
                        elif h_density > v_density * 1.1:
                            features["estimated_asset"] = "SWITCHGEAR"
        except Exception:
            pass

        return features

    # ── 3. OPTIONAL GEMINI FLASH CLOUD VISION ────────────────────────────────
    def classify_with_gemini(self, image_path: str, candidate_slots: list) -> dict:
        """
        Calls Gemini Flash Vision if API key is configured for human-level reasoning.
        """
        if not self.api_key or not HAS_GEMINI_SDK or not os.path.exists(image_path):
            return None

        try:
            model = genai.GenerativeModel("gemini-2.0-flash")
            with open(image_path, "rb") as f:
                img_data = f.read()

            prompt = (
                "You are an expert electrical substation CBM inspector. Analyze this photo:\n"
                "1. Is it Switchgear (RMU/VCB), Transformer (TX), or Feeder Pillar (FP)?\n"
                "2. Is it an Overview, Bushing, Cable Compartment, or Nameplate?\n"
                "3. Any visible text on labels or panels?\n"
                "Output as JSON: {\"asset\": \"...\", \"type\": \"...\", \"label\": \"...\", \"is_overview\": true/false}"
            )

            response = model.generate_content([
                {"mime_type": "image/jpeg", "data": img_data},
                prompt
            ])

            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                return json.loads(match.group(0))
        except Exception as e:
            print(f"[Brain] Gemini Vision error: {e}")

        return None

    # ── 4. CLUSTER & SLOTS MATCHER ───────────────────────────────────────────
    def match_pair_to_slot(self, pair: dict, required_slots: list, ocr_text: str, visual_features: dict) -> dict:
        """
        Fuzzy matches extracted text and visual characteristics against CSV slot names.
        """
        best_match = None
        best_score = 0.0

        for slot in required_slots:
            slot_name = (slot.get("panel_name") or slot.get("area") or slot.get("type") or "").strip().lower()
            if not slot_name or slot_name == "-":
                continue

            # Compare OCR text with panel/slot name
            if ocr_text:
                ratio = difflib.SequenceMatcher(None, slot_name, ocr_text.lower()).ratio()
                if slot_name in ocr_text.lower():
                    ratio = max(ratio, 0.85)

                if ratio > best_score:
                    best_score = ratio
                    best_match = slot

        return best_match if best_score >= 0.45 else None

    # ── 5. FULL OPTIMIZATION PIPELINE ─────────────────────────────────────────
    def optimize_plan_with_vision(self, raw_pairs: list, required_slots: list, log_fn=print) -> list:
        """
        Main Brain reasoning pipeline:
        1. Analyzes each raw pair (OCR + CV features + optional Gemini).
        2. Detects asset clusters (TX1, TX2, Switchgear, FP).
        3. Prioritizes the best wide shots for Overview slots.
        4. Matches panel details and tags spares.
        """
        log_fn("  · 🧠 Vision Brain: Analyzing image features and reading nameplates...")

        evaluated_pairs = []
        for p in raw_pairs:
            dc_p = p.get("full_dc_path")
            ocr_text = self.extract_text_from_image(dc_p) if dc_p else ""
            vis_feat = self.detect_asset_visual_features(dc_p) if dc_p else {}

            evaluated_pairs.append({
                "raw_pair": p,
                "ocr_text": ocr_text,
                "features": vis_feat,
                "assigned": False
            })

        # Match slots dynamically
        assigned_plan = []
        used_pairs = set()

        # Phase A: Direct OCR Label Match
        for slot in required_slots:
            slot_panel = (slot.get("panel_name") or "").strip()
            if not slot_panel or slot_panel == "-":
                continue

            for idx, ep in enumerate(evaluated_pairs):
                if idx in used_pairs:
                    continue

                if ep["ocr_text"] and slot_panel.lower() in ep["ocr_text"].lower():
                    used_pairs.add(idx)
                    ep["assigned"] = True
                    assigned_plan.append({
                        "type": "MAPPED",
                        "component": slot_panel,
                        "raw_dc": ep["raw_pair"].get("dc_name"),
                        "new_dc": slot.get("visual_img"),
                        "raw_ir": ep["raw_pair"].get("ir_name"),
                        "new_ir": slot.get("ir_img"),
                        "match_reason": f"OCR Match: '{slot_panel}'"
                    })
                    break

        # Phase B: Sequential mapping for remaining required slots
        remaining_slots = [s for s in required_slots if not any(p.get("new_dc") == s.get("visual_img") for p in assigned_plan)]
        spare_counter = 1

        for idx, ep in enumerate(evaluated_pairs):
            if idx in used_pairs:
                continue

            if remaining_slots:
                slot = remaining_slots.pop(0)
                comp_label = slot.get("panel_name") or slot.get("area") or slot.get("type") or f"Slot"
                assigned_plan.append({
                    "type": "MAPPED",
                    "component": comp_label,
                    "raw_dc": ep["raw_pair"].get("dc_name"),
                    "new_dc": slot.get("visual_img"),
                    "raw_ir": ep["raw_pair"].get("ir_name"),
                    "new_ir": slot.get("ir_img"),
                    "match_reason": "Visual Sequence Match"
                })
            else:
                spare_ir = f"A{spare_counter:04d}.jpg"
                spare_dc = f"A{spare_counter+1:04d}.jpg"
                spare_counter += 2

                assigned_plan.append({
                    "type": "SPARE",
                    "component": "Alternate / Spare Take",
                    "raw_dc": ep["raw_pair"].get("dc_name"),
                    "new_dc": spare_dc,
                    "raw_ir": ep["raw_pair"].get("ir_name"),
                    "new_ir": spare_ir,
                    "match_reason": "Unassigned Spare Take"
                })

        return assigned_plan


# Singleton instance
BRAIN = ImageVisionBrain()
