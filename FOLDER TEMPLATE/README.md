# 33kV Substation Folder Templates & Batch Automation Suite

Dynamic, multi-system folder template suite and automated image management pipeline for **Substation Type 33kV** (and multi-switchgear configurations up to 4 systems and 30 panels per system).

---

## 1. Official 1000-Block Image Naming Standard

Aligned directly with [`Master List.xlsx`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/PO%2042394550%20-%20THUNDER%2033kV/06.%20EXCELL%20MAPPING/Master%20List.xlsx) sheet **`Image_Mapping`** and the automated Word report generation pipeline:

| Block Range | System / Section | Unit / Equipment | Compartment / View | Image Filenames |
| :--- | :--- | :--- | :--- | :--- |
| **1 – 1000** | **Sistem 1** *(33kV)* | `Overview 33` | 12 Overview Positions | `1.jpg` – `24.jpg` |
| | | `33panel1` | 8 Compartments | `51.jpg` – `66.jpg` |
| | | `33panel2` | 8 Compartments | `67.jpg` – `82.jpg` |
| | | `33panel...` | 8 Compartments | `51 + (N - 1) * 16` |
| | | `33panel30` | 8 Compartments | `515.jpg` – `530.jpg` |
| **1001 – 2000** | **Sistem 2** *(11kV)* | `Overview 11` | 12 Overview Positions | `1001.jpg` – `1024.jpg` |
| | | `11panel1` | 8 Compartments | `1051.jpg` – `1066.jpg` |
| | | `11panel2` | 8 Compartments | `1067.jpg` – `1082.jpg` |
| | | `11panel...` | 8 Compartments | `1051 + (N - 1) * 16` |
| | | `11panel30` | 8 Compartments | `1515.jpg` – `1530.jpg` |
| **2001 – 3000** | **Sistem 3** | `Overview S3` | 12 Overview Positions | `2001.jpg` – `2024.jpg` |
| | | `s3panel1` | 8 Compartments | `2051.jpg` – `2066.jpg` |
| | | `s3panel...` | 8 Compartments | `2051 + (N - 1) * 16` |
| | | `s3panel30` | 8 Compartments | `2515.jpg` – `2530.jpg` |
| **3001 – 4000** | **Sistem 4** | `Overview S4` | 12 Overview Positions | `3001.jpg` – `3024.jpg` |
| | | `s4panel1` | 8 Compartments | `3051.jpg` – `3066.jpg` |
| | | `s4panel...` | 8 Compartments | `3051 + (N - 1) * 16` |
| | | `s4panel30` | 8 Compartments | `3515.jpg` – `3530.jpg` |
| **4001+** | **Ancillary / Others**| `POWER TX1` | 8 Compartments | `4001.jpg` – `4016.jpg` |
| | | `POWER TX2` | 8 Compartments | `4017.jpg` – `4032.jpg` |
| | | `LOCALT TX` | 8 Compartments | `4033.jpg` – `4048.jpg` |
| | | `FEEDER PILLAR` | Overview | `4049.jpg` – `4052.jpg` |
| | | `BATTERY BANK` | Overview 1..4 | `4053.jpg` – `4060.jpg` |
| | | `NER` | Overview | `4061.jpg` – `4062.jpg` |
| | | `BLACK BOX` | Overview 1..2 | `4063.jpg` – `4066.jpg` |

### 8 Standard Panel Compartments Sequence
Inside every panel, 8 compartments are captured with 2 images each (Visual = Odd, IR = Even):
1. `BREAKER COMPARTMENT` (Visual: $+0$, IR: $+1$)
2. `BUSBAR DISCONNECTING SW` (Visual: $+2$, IR: $+3$)
3. `CABLE COMPARTMENT` (Visual: $+4$, IR: $+5$)
4. `BUSBAR COMPARTMENT` (Visual: $+6$, IR: $+7$)
5. `PT COMPARTMENT` (Visual: $+8$, IR: $+9$)
6. `CABLE ENTRY` (Visual: $+10$, IR: $+11$)
7. `LINK BOX` (Visual: $+12$, IR: $+13$)
8. `SECONDARY COMPARTMENT` (Visual: $+14$, IR: $+15$)

---

## 2. Pre-Generated Standard Templates

The following ready-to-use template directories are generated in this folder:

1. [`33kV (By Location)`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/FOLDER_TEMPLATE/33kV%20(By%20Location))
   - **Configuration**: 1 System (33kV), 30 Panels + Ancillary (285 numbered folders).
   - **Layout**: Grouped by Compartment first (e.g. `013. BREAKER COMPARTMENT - PANEL 1`, `014. BREAKER COMPARTMENT - PANEL 2`...).
2. [`33kV (By Panel)`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/FOLDER_TEMPLATE/33kV%20(By%20Panel))
   - **Configuration**: 1 System (33kV), 30 Panels + Ancillary (285 numbered folders).
   - **Layout**: Grouped by Panel first (e.g. `013. PANEL 1 - BREAKER COMPARTMENT` through `020. PANEL 1 - SECONDARY COMPARTMENT`...).
3. [`33kV - 2 Sistem (By Location)`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/FOLDER_TEMPLATE/33kV%20-%202%20Sistem%20(By%20Location))
   - **Configuration**: 2 Systems (Sistem 1 & Sistem 2), 15 Panels each + Ancillary (297 numbered folders).
   - **Layout**: By Location with System prefixes.
4. [`33kV - 2 Sistem (By Panel)`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/FOLDER_TEMPLATE/33kV%20-%202%20Sistem%20(By%20Panel))
   - **Configuration**: 2 Systems (Sistem 1 & Sistem 2), 15 Panels each + Ancillary (297 numbered folders).
   - **Layout**: By Panel with System prefixes.

---

## 3. Dynamic Template Generator

To generate any custom layout on demand:
- Double click [`CREATE_33kV_TEMPLATE.bat`](file:///C:/Users/aremi/Desktop/REPORT%20CBM/FOLDER_TEMPLATE/CREATE_33kV_TEMPLATE.bat), or
- Run via PowerShell:
```powershell
& .\create_33kv_template.ps1 -SystemCount 3 -PanelsPerSystem @(10, 10, 8) -SystemNames @('132kV', '33kV', '11kV') -Mode 'Both'
```

---

## 4. Automation Batch Suite Inside Every Template

Inside every generated template folder, 5 batch scripts automate the whole workflow:

```
[Template Root]
│── 1_RENAME_IMAGES.bat
│── 2_MOVE_TO_SORT.bat
│── 3_MOVE_BACK_FROM_SORT.bat
│── 4_SORT_BY_PANEL.bat
│── 5_MOVE_BACK_FROM_PANEL.bat
│── SORT\
└── _engine\
    ├── rename_images.ps1
    ├── move_to_sort.ps1
    ├── move_back.ps1
    ├── sort_by_panel.ps1
    └── move_back_from_panel.ps1
```

* **BAT 1 (`1_RENAME_IMAGES.bat`)**: Automatically renames images inside their respective subfolders according to the **1000-block standard** (`1.jpg`/`2.jpg` for Overview, `51.jpg`/`52.jpg` for S1 Panel 1, `1051.jpg`/`1052.jpg` for S2 Panel 1, `4001.jpg`+ for Ancillary).
* **BAT 2 (`2_MOVE_TO_SORT.bat`)**: Moves all renamed images into `SORT` and creates `sort_manifest.json`.
* **BAT 3 (`3_MOVE_BACK_FROM_SORT.bat`)**: Restores all images from `SORT` back into their originating subfolders.
* **BAT 4 (`4_SORT_BY_PANEL.bat`)**: Organizes images into categorized inspection folders (`SORT_BY_PANEL\01. SISTEM 1 - OVERVIEW`, `02. SISTEM 1 - PANEL 1`, etc.).
* **BAT 5 (`5_MOVE_BACK_FROM_PANEL.bat`)**: Completely restores images from `SORT_BY_PANEL` back into their respective numbered subfolders.
