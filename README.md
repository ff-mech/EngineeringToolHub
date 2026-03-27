# Engineering Tool Hub

FoxFab internal engineering utilities combined into a single Windows desktop application.

---

## Tools

### BOM Check
Processes a FoxFab manufacturing BOM (`.xlsx` / `.xlsm`) in two passes:

- **Pass 1 — Stock Parts Check:** Searches the Stock Parts folder (`Z:\...`) using Everything CLI (`es.exe`) and marks matching rows in the `FFMPL` sheet (columns B, G, H).
- **Pass 2 — Non-Stock PDF/DXF Copy:** For every non-stock part, finds the highest-revision PDF and DXF and copies them to a target folder (typically `202 PDFs_Flats`). Column G is marked if either file was found.

> **Requirement:** Excel must be **closed** before running. The app warns you and asks for confirmation.

---

### Doc Prep & Print
Builds a manufacturing packet from a job folder, marks the BOM's CNC column, and sends everything to the FoxFab printer in the correct order.

#### Workflow

| Step | Action | What happens |
|---|---|---|
| 1 | Select job folder → **Build Plan** | App scans for all required documents and shows a summary in the terminal |
| 2 | Review plan → **Generate Documents** | CNC column is marked in the BOM, then all PDFs are generated into a temp folder |
| 3 | *(auto)* | Manual Printing list is populated with every generated PDF, duplex pre-set |
| 4 | Amber bar appears | Lists all documents ready to send — review before committing |
| 5 | **Send to Printer** | Prints all documents in order via Acrobat COM, or **Cancel Print** to abort |

> In **Simulation Mode** (checkbox), PDFs are saved to a timestamped folder next to the `.exe` instead of printing — useful for review away from the printer.

#### CNC Column Marker (automatic, runs inside Generate Documents)

Before the BOM is exported to PDF, the app automatically scans the `205 CNC` folder and marks column H in the `FFMPL` sheet with `X` for every BOM part that has a matching CNC file. This means the printed BOM already has the CNC column filled in.

**How the CNC folder is found:**
The app looks for a folder starting with `205` alongside `204 BOM` inside the job variant folder:

```
<variant>/
    204 BOM/        ← BOM Excel lives here
    205 CNC/        ← found automatically as a sibling
    202 PDFs_Flats/
```

If the folder cannot be found automatically, a directory picker opens so you can point to it manually. Cancelling the picker skips CNC marking and continues with the rest of the print job.

**Filename parsing rules:**

| File pattern | How it's handled |
|---|---|
| `90004.pdf` | Bare digits → interpreted as `240-90004` |
| `250-90002.pdf` | Direct prefix `NNN-XXXXX` → matched as-is |
| `240-31209_31211.pdf` | Multi-part → parsed as `240-31209` **and** `240-31211` |
| `250-30089 rA.pdf` | Revision suffix stripped → matched as `250-30089` |
| `J15302-01_14GALV.pdf` | GALV / J-prefix → PDF opened with PyMuPDF, all `DRAWING NUMBER:` lines extracted |
| `CNC_Simplex_Merged.pdf` | Skipped — filename contains "Merged" |
| Subfolders (e.g. `archive\`) | Skipped — only top-level files are scanned |

**Behaviour after matching:**
- Rows where column H already contains `S` (stock) are skipped entirely.
- Every matched row gets `X` written to column H.
- Any non-S BOM rows with no matching CNC file are listed in the terminal as warnings.
- The workbook is saved via xlwings (COM) so all table formatting and structured styles are preserved.

#### Print order

| # | Document | Duplex | Notes |
|---|---|---|---|
| 01 | Fabrication Work Order | Simplex | Auto-filled from PRF data (job no., name, date, enclosure, qty) |
| 02 | BOM | Simplex | All sheets — CNC column H already marked |
| 03–N | CNC files | **Duplex** | One PDF per duplex CNC file — prints double-sided |
| N+1 | CNC Simplex (merged) | Simplex | All simplex CNC files merged |
| N+2 | PDFs_Flats (merged) | Simplex | All flat PDFs merged |
| N+3 | Production Release Form | Simplex | First sheet only |
| N+4 | Electrical Pack (pages 1–2) | Simplex | First two pages only |
| N+5 | Assemblies (merged) | Simplex | All assembly PDFs merged |

#### Printing via Acrobat COM (single instance)

Instead of launching a new Acrobat process for each document (which costs 3–5 s of startup per file), the app opens **one `AcroExch.App` COM instance** at the start of the print run and sends all documents through it:

1. `AcroExch.App` is dispatched once and hidden.
2. For each document:
   - DEVMODE is set (duplex or simplex) via `win32print` level-9 — no admin rights required.
   - `PDDoc.Open` → `OpenAVDoc` → `PrintPages` (silent, no dialog) → close doc.
   - 0.2 s breath between documents.
3. `AcroExch.App.Exit()` is called when all documents are done.

Because all jobs are submitted sequentially through the same Acrobat instance, print order is guaranteed by the COM call order — no spooler polling is needed between documents.

**Fallback:** if `win32com` is unavailable or Acrobat's COM interface is not registered, the app falls back to the original per-document `subprocess.run /t` approach with spooler polling between documents. Both Doc Prep & Print and Manual Printing use the same COM path with the same fallback.

#### BOM revision auto-selection
If a job folder contains multiple BOM Excel files (e.g. after a re-release), the app automatically picks the one with the **highest revision letter** (`rA < rB < rC …`). If no revision suffix is found in any filename, a pick-list prompt is shown as a fallback.

#### FWO auto-fill
The Fabrication Work Order PDF is filled automatically using data read from the job's PRF Excel file:
- Reads: Job No. (C4), Model No. (G9), Job Name (C8), Enclosure (G18–G20, G22), Qty
- Overlays text using PyMuPDF; field positions are tunable via `FWO_*` constants at the top of `app.py`
- **Preview FWO** button: fills and saves `logs/FWO_preview.pdf` and opens it — use this to tune coordinates without a full print run

#### Preview BOM button
Runs the full CNC column marking pass first, then exports the BOM to `logs/BOM_preview.pdf` and opens it. Use this after **Build Plan** to verify column H before committing to a full print run:

1. App locates the `205 CNC` folder (prompts if not found).
2. CNC marking runs and the BOM is saved.
3. BOM is exported to PDF via Excel COM.
4. `logs/BOM_preview.pdf` opens automatically.

#### Manual Printing (collapsible)

A **▶ Manual Printing** toggle at the bottom of the panel expands a file list:

- After **Generate Documents**, the list is **automatically populated** with every generated PDF in print order, with duplex pre-set per file (CNC duplex files = on, all others = off).
- Add individual PDFs manually via file picker, or remove any entry with ✕.
- **Print** button sends a single file immediately (background thread — UI stays responsive).
- **Print All** sends all files in list order via the same Acrobat COM session.
- Shares the printer field and terminal with Doc Prep & Print.

Use Manual Printing to resend any individual document or the full batch without re-running the whole plan.

---

### SW Batch Update
References the SolidWorks macro for batch updating custom properties and exporting DXFs:

- Displays the macro file path and an Open Folder shortcut
- Explains the steps to run it inside SolidWorks (Tools → Macro → Run)
- The macro updates `DrawnBy`, `DwgDrawnBy`, drawing properties, and exports flat-pattern DXFs

> **Requirement:** SolidWorks must be installed on the machine.

---

## Features

- Sidebar navigation — switch between tools instantly, no re-launch
- Per-tool progress bar and terminal output with colour-coded log levels
- **Stop** button — hard stop via `threading.Event`, checked between iterations
- Inline confirm bars for plan review and print gate (no popups)
- Master log written to `logs/ETH_master_YYYY-MM-DD.log` next to the `.exe` after each run
- Temp print folder deferred cleanup — files persist until next run so the spooler finishes safely

---

## Requirements

| Dependency | Purpose |
|---|---|
| Python 3.10+ | Runtime |
| `xlwings` | Excel COM automation (BOM Check, CNC marking, PDF export) |
| `pypdf` | PDF merging |
| `pymupdf` | FWO text overlay + GALV PDF part number extraction |
| `openpyxl` | PRF Excel data reading (no COM needed) |
| `pywin32` | Windows COM + DEVMODE duplex + print API |
| `pyinstaller` | Building the `.exe` |
| `es.exe` | Everything CLI for fast file search (BOM Check) |
| Adobe Acrobat 2017+ (installed) | Single-instance COM printing with per-document duplex |
| Microsoft Excel (installed) | BOM Check, CNC marking, BOM/PRF export via xlwings |

Install Python dependencies:
```
pip install xlwings pypdf pywin32 pyinstaller pymupdf openpyxl
```

`es.exe` (Everything CLI) should be placed at `tools\BomCheck\es.exe`. Download from [voidtools.com](https://www.voidtools.com/support/everything/command_line_interface/).

---

## Build

Double-click **`build.bat`** or run it from a terminal:

```bat
build.bat
```

The script will:
1. Install / upgrade all Python dependencies via pip
2. Clean any previous `dist\` and `build\` folders
3. Bundle the app with PyInstaller (`--onedir`, no extraction at launch)
4. Copy `es.exe` into the output folder

Output: `dist\Engineering Tool Hub\Engineering Tool Hub.exe`

> **To distribute:** copy the entire `dist\Engineering Tool Hub\` folder. Do **not** move just the `.exe` — it needs the surrounding `_internal\` folder.

---

## Project Structure

```
Engineering Tool Hub\
├── app.py                          # Combined application (single file)
├── build.bat                       # PyInstaller build script
├── .gitignore
├── tools\
│   ├── BomCheck\
│   │   ├── main.py                 # Original standalone BOM script
│   │   └── es.exe                  # Everything CLI binary
│   ├── DocPrepPrint\
│   │   ├── DocPrepPrint.py         # Original print script reference
│   │   └── DocPrepPrint_Test(...).py
│   └── SolidworksBatchUpdate\
│       └── SoldworksBatchUpdate.bas   # Original VB macro (reference)
└── logs\                           # Created at runtime — excluded from git
    └── ETH_master_YYYY-MM-DD.log
```

---

## Tuning FWO Field Positions

If the auto-filled text on the Fabrication Work Order lands in the wrong position after a test print, adjust the constants near the top of `app.py`:

```python
FWO_JOB_NO_X    = 165   # x position for Job No. value
FWO_JOB_NO_Y    = 145   # y position for Job No. value  (decrease = move up)
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

Use the **Preview FWO** button to see changes instantly without running a full print job. The preview is saved to `logs/FWO_preview.pdf`.

---

## Notes

- The preferred printer name is hardcoded as `PREFERRED_PRINTER` in `app.py`. Update this constant if the printer name changes.
- BOM Check requires Excel to be **closed** — xlwings opens the file via COM in the background.
- CNC marking (inside Doc Prep & Print) also opens the BOM via xlwings — close Excel before clicking **Generate Documents**.
- Acrobat is searched in common install paths automatically. If installed in a non-standard location, add the path to `ACROBAT_SEARCH_PATHS` in `app.py`.
- The `logs/` folder is excluded from git (see `.gitignore`). Log files are local only.
