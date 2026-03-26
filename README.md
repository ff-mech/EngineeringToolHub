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
Builds a manufacturing packet from a job folder and sends it to the FoxFab printer in the correct order.

#### Workflow
1. Select a job folder and click **Build Plan** — the app scans for all required documents and shows a summary.
2. Review the plan in the terminal, then click **Generate Documents** (or **Start Simulation** in sim mode).
3. All PDFs are generated first into a temp folder.
4. In print mode, a second amber confirmation bar appears listing the documents ready to send.
5. Click **Send to Printer** to print them in order with a 3.5-second gap between each, or **Cancel Print** to abort.

#### Print order
| # | Document | Notes |
|---|---|---|
| 01 | Fabrication Work Order | Auto-filled from PRF data (job no., name, date, enclosure, qty) |
| 02 | BOM | All sheets |
| 03–N | CNC files (duplex) | One PDF per duplex CNC file — prints double-sided |
| N+1 | CNC Simplex (merged) | All simplex CNC files merged — single-sided |
| N+2 | PDFs_Flats (merged) | All flat PDFs merged |
| N+3 | Production Release Form | First sheet only |
| N+4 | Electrical Pack (pages 1–2) | First two pages only |
| N+5 | Assemblies (merged) | All assembly PDFs merged |

#### Duplex control
Printing uses **Adobe Acrobat** (`Acrobat.exe /t`) for per-document duplex control:
- CNC duplex files → **duplex** (long-edge binding)
- All other documents → **simplex**

Duplex is set via Windows per-user DEVMODE (level 9 — no admin rights required). Acrobat is located automatically from common installation paths. If not found, falls back to `os.startfile`.

#### FWO auto-fill
The Fabrication Work Order PDF is filled automatically using data read from the job's PRF Excel file:
- Reads: Job No. (C4), Model No. (G9), Job Name (C8), Enclosure (G18–G20, G22), Qty
- Overlays text using PyMuPDF; field positions are tunable via `FWO_*` constants at the top of `app.py`
- **Preview FWO** button: fills and saves `logs/FWO_preview.pdf` and opens it — use this to tune coordinates without a full print run

#### Simulation Mode
Checkbox toggle — saves all generated PDFs to a timestamped folder next to the `.exe` instead of printing. Useful for review or testing away from the printer.

#### Manual Printing (collapsible)
A **▶ Manual Printing** toggle at the bottom of the panel expands a file list for debug/testing:
- Add individual PDFs via file picker
- Set duplex per file with a checkbox
- **Print** button sends a single file immediately
- **Print All** sends all files in list order with 3.5-second gaps
- Shares the same printer field and terminal as Doc Prep & Print

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
| `xlwings` | Excel COM automation (BOM Check, PDF export) |
| `pypdf` | PDF merging |
| `pymupdf` | FWO text overlay (auto-fill) |
| `openpyxl` | PRF Excel data reading (no COM needed) |
| `pywin32` | Windows COM + print API |
| `pyinstaller` | Building the `.exe` |
| `es.exe` | Everything CLI for fast file search (BOM Check) |
| Adobe Acrobat 2017+ (installed) | Per-document duplex printing via `/t` flag |
| Microsoft Excel (installed) | BOM Check + PDF export via xlwings |

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
FWO_JOB_NO_X    = 152   # x position for Job No. value
FWO_JOB_NO_Y    = 138   # y position for Job No. value  (decrease = move up)
FWO_JOB_NAME_X  = 148
FWO_JOB_NAME_Y  = 160
FWO_DATE_X      = 140
FWO_DATE_Y      = 178
FWO_ENCLOSURE_X = 140
FWO_ENCLOSURE_Y = 213
FWO_UNITS_X     = 140
FWO_UNITS_Y     = 261
FWO_FONT_SIZE   = 11    # pt
```

Use the **Preview FWO** button to see changes instantly without running a full print job. The preview is saved to `logs/FWO_preview.pdf`.

---

## Notes

- The preferred printer name is hardcoded as `PREFERRED_PRINTER` in `app.py`. Update this constant if the printer name changes.
- BOM Check requires Excel to be **closed** — xlwings opens the file via COM in the background.
- Acrobat is searched in common install paths automatically. If installed in a non-standard location, add the path to `ACROBAT_SEARCH_PATHS` in `app.py`.
- The `logs/` folder is excluded from git (see `.gitignore`). Log files are local only.
