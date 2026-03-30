# Engineering Tool Hub

FoxFab internal engineering utilities combined into a single Windows desktop application.

---

## Table of Contents

- [Overview](#overview)
- [Quick Start](#quick-start)
- [Tools](#tools)
  - [Bom Filler](#bom-filler)
  - [Doc Prep & Print](#doc-prep--print)
  - [SW Batch Update](#sw-batch-update)
- [App Features](#app-features)
- [Requirements](#requirements)
- [Build](#build)
- [Project Structure](#project-structure)
- [Configuration Reference](#configuration-reference)

---

## Overview

Engineering Tool Hub wraps three manufacturing workflow tools into one professional Windows desktop app — sidebar navigation, per-tool terminal output, progress bars, stop controls, and a master log.

```
┌─────────────────────────────────────────────────────────┐
│  Engineering Tool Hub                                   │
│  ┌──────────────┬────────────────────────────────────┐  │
│  │              │                                    │  │
│  │  Bom Filler   │   [ Tool Panel ]                   │  │
│  │              │   Config  ─────────────────────    │  │
│  │  Doc Prep    │   ▶ Run   ■ Stop   [======   ]     │  │
│  │  & Print     │                                    │  │
│  │              │   ┌────────────────────────────┐   │  │
│  │  SW Batch    │   │ terminal output...         │   │  │
│  │  Update      │   │ [INFO]  pass 1 complete    │   │  │
│  │              │   │ [OK]    3 files copied     │   │  │
│  └──────────────┴───┴────────────────────────────┘   │  │
└─────────────────────────────────────────────────────────┘
```

Each tool runs in a background thread — the UI stays responsive throughout. Output is streamed to the terminal panel in real time and appended to a master log file after each run.

---

## Quick Start

1. Install Python 3.10+, Microsoft Excel, and Adobe Acrobat 2017+
2. Place `es.exe` (Everything CLI) at `tools\BomCheck\es.exe`
3. Make sure the **Everything** service is running and has indexed `Z:\`
4. Double-click **`build.bat`** — it installs dependencies and produces the `.exe`
5. Run `dist\Engineering Tool Hub\Engineering Tool Hub.exe`

> First run on a new machine? See [Requirements](#requirements) for the full list.

---

## Tools

### Bom Filler

Processes a FoxFab manufacturing BOM (`.xlsx` / `.xlsm`) against the stock parts library and copies the correct-revision files for non-stock parts.

#### Workflow

```
User selects BOM file and target folder
              │
              ▼
    ┌─────────────────────┐
    │  PASS 1             │
    │  Stock Parts Check  │
    └─────────────────────┘
    For each part in column A:
      └─ es.exe searches Stock Parts folder (Z:\...300 Stock Parts\...)
            ├─ FOUND  → col B = X, col G = S, col H = S
            └─ not found → continue to Pass 2
              │
              ▼
    ┌─────────────────────────┐
    │  PASS 2                 │
    │  Non-Stock File Copy    │
    └─────────────────────────┘
    For each non-stock part:
      └─ es.exe finds all filenames starting with that part number
            └─ extract revisions (rA, rB, rC …) → pick highest
                  ├─ copy <part> <rev>.pdf  → target folder
                  ├─ copy <part> <rev>.dxf  → target folder
                  └─ col G = X if either file copied or already existed
              │
              ▼
    Workbook saved (xlwings COM — all formatting preserved)
    Summary printed to terminal
```

#### BOM Sheet Markings

| Column | Value | Meaning |
|--------|-------|---------|
| B — STOCK PART | `X` | Part found in stock library |
| G — PDF | `S` | Stock part — PDF sourced from stock |
| G — PDF | `X` | Non-stock — PDF/DXF copied to target folder |
| H — CNC | `S` | Stock part — CNC sourced from stock |

#### Revision Detection

Files are matched by the pattern `<part>[-_ ]r[A-Z]`. Given:
```
240-2202 rA.pdf   240-2202 rA.dxf
240-2202 rC.pdf   240-2202 rC.dxf
```
The script resolves `rC` (highest letter) and copies both files. If no revision suffix exists, the bare part number is used as a fallback.

> **Requirement:** Excel must be **closed** before running. The app warns you and asks for confirmation before proceeding.

---

### Doc Prep & Print

Builds a complete manufacturing packet from a job folder, marks the BOM's CNC column, and sends all documents to the FoxFab printer in the correct order — or saves them as PDFs in Simulation Mode.

#### Workflow

```
1. SELECT JOB FOLDER
   User picks the job variant folder containing 204 BOM\, 205 CNC\, etc.
              │
              ▼
2. BUILD PLAN  (▶ Build Plan button)
   App scans the job folder and reports:
     • BOM Excel file found (highest-revision auto-selected)
     • PRF Excel file found
     • FWO template found
     • CNC files listed (duplex vs simplex sorted)
     • PDFs_Flats files listed
     • Assemblies listed
              │
              ▼
3. REVIEW → GENERATE DOCUMENTS  (▶ Generate Documents button)
   ┌─────────────────────────────────────────────────────────┐
   │ a. CNC Column Marking                                   │
   │    Scans 205 CNC\ folder, marks col H = X in FFMPL     │
   │    sheet for every BOM part with a matching CNC file    │
   │    Saves workbook via xlwings COM                       │
   ├─────────────────────────────────────────────────────────┤
   │ b. Document Generation (in print order)                 │
   │    01  FWO PDF        — auto-filled from PRF data       │
   │    02  BOM PDF        — exported from Excel (all sheets)│
   │    03–N  CNC duplex   — one PDF per duplex CNC file     │
   │    N+1  CNC simplex   — all simplex CNCs merged         │
   │    N+2  PDFs_Flats    — all flat PDFs merged            │
   │    N+3  PRF           — first sheet only                │
   │    N+4  Electrical    — pages 1–2 only                  │
   │    N+5  Assemblies    — all assembly PDFs merged        │
   └─────────────────────────────────────────────────────────┘
   Manual Printing list auto-populated with all generated PDFs
              │
              ▼
4. AMBER CONFIRM BAR appears
   Lists every document queued for print — review before committing
              │
        ┌─────┴───────┐
        ▼             ▼
5a. SEND TO PRINTER   5b. CANCEL PRINT
    All docs sent         Queue cleared,
    via Acrobat COM       no pages printed
    in order
```

#### Printing via Acrobat COM

Instead of spawning a new Acrobat process per file (3–5 s startup each), the app opens **one `AcroExch.App` COM instance** for the entire run:

```
AcroExch.App dispatched once (hidden)
    │
    ├─ Doc 1: DEVMODE set (simplex/duplex) → PDDoc.Open → PrintPages → close
    ├─ Doc 2: DEVMODE set → PDDoc.Open → PrintPages → close
    │   ...
    └─ Doc N: DEVMODE set → PDDoc.Open → PrintPages → close
    │
AcroExch.App.Exit()
```

Print order is guaranteed by COM call order — no spooler polling needed. If Acrobat COM is unavailable, the app falls back to `subprocess.run /t` with spooler polling between documents.

#### CNC Column Marking (inside Generate Documents)

Before exporting the BOM to PDF, the app scans `205 CNC\` and writes `X` to column H for every matched BOM part:

| CNC filename pattern | How it's matched |
|---|---|
| `90004.pdf` | Bare digits → treated as `240-90004` |
| `250-90002.pdf` | Direct `NNN-XXXXX` prefix → matched as-is |
| `240-31209_31211.pdf` | Multi-part → matched as `240-31209` **and** `240-31211` |
| `250-30089 rA.pdf` | Revision suffix stripped before matching |
| `J15302-01_14GALV.pdf` | GALV / J-prefix → PyMuPDF extracts `DRAWING NUMBER:` lines |
| `CNC_Simplex_Merged.pdf` | Skipped — filename contains "Merged" |
| Files in subfolders | Skipped — only top-level files scanned |

Rules: rows where column H already has `S` (stock) are skipped. Non-matched non-stock rows are listed as warnings.

**CNC folder discovery:** the app looks for a `205*` sibling folder next to `204 BOM\`. If not found, a directory picker opens. Cancelling the picker skips CNC marking and continues the print run.

#### FWO Auto-Fill

The Fabrication Work Order PDF is filled automatically from the PRF Excel file:

| PRF cell(s) | FWO field |
|---|---|
| C4 | Job No. |
| C8 | Job Name |
| G9 | Model No. |
| G18–G20, G22 | Enclosure |
| Qty | Units |

Use **Preview FWO** to see the filled result in `logs/FWO_preview.pdf` without running a full print job. Adjust position constants at the top of `app.py` if text lands in the wrong spot (see [FWO Position Tuning](#fwo-position-tuning)).

#### BOM Revision Auto-Selection

If multiple BOM Excel files exist in `204 BOM\` (e.g. after a re-release), the app automatically picks the one with the **highest revision letter** (`rA < rB < rC …`). A pick-list prompt is shown as a fallback if no revision suffix is found.

#### Manual Printing (collapsible)

A **▶ Manual Printing** toggle expands a file list at the bottom of the panel:

- After **Generate Documents**, the list is auto-populated with every generated PDF in print order, duplex pre-set per file
- Add individual PDFs manually via file picker, or remove any entry with ✕
- **Print** sends a single selected file immediately (background thread)
- **Print All** sends all files in list order via the same Acrobat COM session

Use this to resend any individual document or the full batch without re-running the whole plan.

#### Simulation Mode

Toggle the **Simulation Mode** checkbox before clicking **Send to Printer**. Instead of printing, all documents are saved to a timestamped folder next to the `.exe`:

```
Simulated_Print_Output_<job>_<YYYYMMDD_HHMMSS>\
    01_FWO_...pdf
    02_BOM_...pdf
    03_CNC_...pdf
    ...
```

---

### SW Batch Update

References the SolidWorks VB macro for batch-updating custom properties and exporting DXFs.

The panel displays:
- The macro file path (`tools\SolidworksBatchUpdate\SoldworksBatchUpdate.bas`)
- An **Open Folder** shortcut
- Step-by-step instructions for running the macro inside SolidWorks (Tools → Macro → Run)

The macro updates `DrawnBy`, `DwgDrawnBy`, drawing properties, and exports flat-pattern DXFs.

> **Requirement:** SolidWorks must be installed on the machine.

---

## App Features

| Feature | Detail |
|---|---|
| Sidebar navigation | Switch between tools instantly — no re-launch |
| Per-tool terminal | Colour-coded log levels: `[INFO]`, `[OK]`, `[WARN]`, `[ERROR]` |
| Progress bar | Per-tool, updates during each pass |
| Stop button | Hard stop via `threading.Event` checked between iterations |
| Inline confirm bars | Plan review and print gate — no popups |
| Simulation Mode | PDFs saved locally instead of printed |
| Preview FWO | Fills and opens FWO PDF without a print run |
| Preview BOM | Runs CNC marking, exports BOM to PDF, opens it |
| Master log | `logs/ETH_master_YYYY-MM-DD.log` appended after each run |
| Temp folder cleanup | Deferred to next run so spooler finishes safely |

---

## Requirements

| Dependency | Purpose |
|---|---|
| Python 3.10+ | Runtime |
| `xlwings` | Excel COM automation (Bom Filler, CNC marking, PDF export) |
| `pypdf` | PDF merging |
| `pymupdf` | FWO text overlay + GALV drawing number extraction |
| `openpyxl` | PRF data reading (no COM required) |
| `pywin32` | Windows COM, DEVMODE duplex, print API |
| `pyinstaller` | Building the `.exe` |
| `es.exe` | Everything CLI — fast filename search for Bom Filler |
| Microsoft Excel | Bom Filler, CNC column marking, BOM/PRF export |
| Adobe Acrobat 2017+ | Single-instance COM printing with per-document duplex |

Install Python packages:
```
pip install xlwings pypdf pywin32 pyinstaller pymupdf openpyxl
```

`es.exe` (Everything CLI) belongs at `tools\BomCheck\es.exe`.
Download from [voidtools.com](https://www.voidtools.com/support/everything/command_line_interface/).
The **Everything** service must be running and have indexed `Z:\`.

---

## Build

Double-click **`build.bat`** or run it from a terminal at the repo root:

```
build.bat
```

The script runs four steps:

```
[1/4]  Install / upgrade Python dependencies (pip)
[2/4]  Clean dist\ and build\ from any previous run
[3/4]  Run PyInstaller --onedir --windowed
         Bundles: app.py, es.exe, all hidden imports
[4/4]  Copy es.exe into dist\Engineering Tool Hub\ (fallback copy)
```

**Output:** `dist\Engineering Tool Hub\Engineering Tool Hub.exe`

> **To distribute:** copy the **entire** `dist\Engineering Tool Hub\` folder.
> Do **not** move just the `.exe` — it requires the `_internal\` folder alongside it.

---

## Project Structure

```
Engineering Tool Hub\
├── app.py                              # Combined application (single file)
├── build.bat                           # PyInstaller build script
├── .gitignore
├── tools\
│   ├── BomCheck\
│   │   ├── main.py                     # Original standalone BOM script
│   │   ├── build.bat                   # Standalone BomCheck build (legacy)
│   │   ├── BomCheck.spec               # PyInstaller spec (legacy)
│   │   └── es.exe                      # Everything CLI binary
│   ├── DocPrepPrint\
│   │   ├── DocPrepPrint.py             # Original print script (reference)
│   │   └── DocPrepPrint_Test(...).py
│   └── SolidworksBatchUpdate\
│       └── SoldworksBatchUpdate.bas    # VB macro (run inside SolidWorks)
└── logs\                               # Created at runtime — excluded from git
    └── ETH_master_YYYY-MM-DD.log
```

---

## Configuration Reference

### Constants in `app.py`

| Constant | Default | Purpose |
|---|---|---|
| `PREFERRED_PRINTER` | `\\NPSVR05\FoxFab (Konica Bizhub C360i)` | Default printer name |
| `STOCK_PARTS_FOLDER` | `Z:\FOXFAB_DATA\...\300 Stock Parts\PDFs & Flats` | Stock parts search root |
| `ACROBAT_SEARCH_PATHS` | List of common install paths | Where to find Acrobat.exe |
| `BOM_SHEET_NAME` | `FFMPL` | Sheet name in BOM workbook |

### FWO Position Tuning

If auto-filled text lands in the wrong position after a test print, adjust these constants near the top of `app.py`:

```python
FWO_JOB_NO_X    = 165   # x position for Job No. value
FWO_JOB_NO_Y    = 145   # y position (decrease = move up)
FWO_JOB_NAME_X  = 165
FWO_JOB_NAME_Y  = 165
FWO_DATE_X      = 165
FWO_DATE_Y      = 182
FWO_ENCLOSURE_X = 165
FWO_ENCLOSURE_Y = 210
FWO_UNITS_X     = 165
FWO_UNITS_Y     = 245
FWO_FONT_SIZE   = 11    # pt
```

Use **Preview FWO** to verify changes without a full print run.

### Notes

- Bom Filler and CNC marking both open the BOM via xlwings COM — close Excel first.
- Acrobat is located automatically from `ACROBAT_SEARCH_PATHS`. Add a custom path there if installed in a non-standard location.
- The `logs\` folder is excluded from git. Log files are local only.
- The preferred printer name is `PREFERRED_PRINTER` in `app.py` — update if the printer is renamed.
