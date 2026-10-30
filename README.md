# WRAITH APP (Condition-Based Monitoring Reporting Suite)

**WRAITH APP** is a specialized automation and reporting platform engineered for Electrical Substation Condition-Based Monitoring (CBM). It streamlines testsheet ingestion, data extraction, Master List synchronization, Word report synthesis, high-speed thermal/digital image injection, ultrasound sound embedding, and final multi-format document merging.

---

## 🚀 Key Modules & Capabilities

### Section 1: Data & Generation
- **Split Testsheets (TIANU)**: Automated slicing and organizing of multi-page electrical test records.
- **Extract Substation Data (Scan)**: OCR / tabular extraction from raw equipment testsheets into structured Excel datasets.
- **Sync & Merge to Master List**: Bidirectional reconciliation of scanned field parameters with central PO tracking registries.
- **Generate Master CSV**: Standardized schema mapping for substation equipment components and status flags.
- **Generate Reports (Page-by-Page)**: Templated Microsoft Word document assembly for individual substation panels and switchgear parts.
- **Regenerate by Page Path / Link**: Targeted on-demand regeneration of specific equipment chapters.
- **Executive Summary Generation**: Automated synthesis of high-level asset health and defect summaries.
- **🔥 Inject FLIR & Digital Images**: Fast, direct thermal radiometric (`.jpg` / FLIR) and visual digital photo injection with automated scaling and retry resilience.

### Section 2: Merge & Export
- **Merge Word Reports**: Programmatic joining of individual section Word files into unified, formatted inspection dossiers.
- **Merge PDF Files**: Bookmarked multi-substation PDF compilation.
- **Convert Word → PDF**: High-fidelity COM-level PDF conversion.
- **Insert Sound (WAV) → PDF**: Interactive acoustic ultrasound recording attachment into final PDF deliverables.

---

## 📁 Repository Structure

```text
WRAITH APP/
├── src/
│   ├── app_gui.py                      # Main Tkinter Desktop HUD Application
│   ├── inject_flir_images.py           # Unified Thermal & Visual Image Injection Engine
│   ├── extract_substation_data.py      # Substation extraction & parsing scripts
│   ├── generate_page_report.py         # Word report generation engine
│   ├── generate_csv.py                 # Substation CSV builder
│   ├── sync_master_list.py             # Master list reconciliation
│   ├── merge_word_reports.py           # Word document merger
│   ├── merge_pdfs.py                   # PDF collation engine
│   ├── convert_word_to_pdf.py          # Word to PDF converter
│   ├── final_insert_sounds.py          # Ultrasound WAV embedder
│   ├── flir_libs/                      # Native FLIR SDK & Report runtime libraries
│   └── ...
├── FOLDER TEMPLATE/                    # Standardized CBM PO Folder & Scaffolding Blueprints
│   ├── 33kV (By Location)/             # Image sorting & folder schemes
│   ├── 33kV (By Panel)/
│   ├── RMU/
│   ├── VCB (By Location)/
│   ├── VCB (By Panel)/
│   ├── 06. EXCELL MAPPING/
│   │   └── Master List.xlsx            # Standard Master List schema template
│   └── create_33kv_template.ps1        # Dynamic multi-switchgear PO folder generator
├── templates/                          # Standardized CBM Word (.docx) & Excel templates
│   ├── Master List.xlsx                # Official Master List with intact formula coordinates
│   ├── IR - EET - PAHANG/
│   ├── IR - TIANU/
│   ├── IR UE - EET - PAHANG/
│   ├── IR UE - TIANU/
│   └── IR UE TEV - EET - PAHANG/
├── PO FOLDER EXAMPLE/                  # Full Reference PO Workspace (Standard 01-08 Structure)
│   ├── 01. FULL REPORT/
│   ├── 02. QUICK REPORT/               # Reference Quick Report docx & pdf files
│   ├── 03. RAW MATERIAL/               # Reference site photos & radiometric thermal images
│   ├── 04. TESTSHEET/                  # Reference test records (xlsx & pdf)
│   ├── 05. TEMPLATE/                   # Substation docx template set
│   ├── 06. EXCELL MAPPING/             # Working Master List.xlsx with mapped substations
│   ├── 07. OUTPUT FILE/
│   └── 08. INSERT SOUND/
├── build_app.py                        # Standalone PyInstaller build script (Directory mode)
├── build_onefile.py                    # PyInstaller Single-File build script
├── copy_assets.py                      # Post-build asset distribution helper
├── Launch WRAITH App.vbs               # Silent VBScript launcher (no console window)
├── WRAITH_APP.spec                     # PyInstaller specification file
└── .gitignore
```

---

## 🛠️ Installation & Development

### Prerequisites
- Windows 10 / 11 (64-bit)
- Python 3.10+
- Microsoft Office Word (for COM automation features)

### Running from Source
```bash
# Clone the repository
git clone https://github.com/aremint9210/Wraith-APPS.git
cd Wraith-APPS

# Install dependencies
pip install -r requirements.txt   # or install python-docx, pywin32, pymupdf, pillow, openpyxl

# Launch the Application
python src/app_gui.py
```

### Building the Standalone Application
```bash
python build_app.py
```
Compiled output will be placed in `dist/WRAITH_APP/WRAITH_APP.exe` with bundled runtimes.

---

## 📄 License
Internal Engineering Tool — All rights reserved.
