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
Builds and prints a manufacturing packet from a job folder:

- Detects job PDFs, Excel sheets, and drawing files
- Merges and orders pages into a single print job
- Sends to the FoxFab Konica Bizhub (`NPSVR05`) with correct duplex settings
- **Simulation Mode** (checkbox toggle): saves PDFs to disk instead of printing — useful for review or off-site use

---

### SW Batch Update
Batch-processes a folder of SolidWorks files (`.sldprt`, `.sldasm`, `.slddrw`):

- Updates custom properties (e.g. material, finish, drawn by) across all files
- Exports flat DXFs for all parts
- Uses the SolidWorks COM API via `win32com`
- Stats box shows part/assembly/drawing counts in the selected folder before you run

> **Requirement:** SolidWorks must be installed on the machine.

---

## Features

- Sidebar navigation — switch between tools instantly, no re-launch
- Per-tool progress bar (fills as work progresses), terminal output, and Stop button
- Hard stop via `threading.Event` — checked between iterations, no hung processes
- Master log written to `logs/ETH_master_YYYY-MM-DD.log` next to the `.exe` after each run

---

## Requirements

| Dependency | Purpose |
|---|---|
| Python 3.10+ | Runtime |
| `xlwings` | Excel COM automation (BOM Check) |
| `pypdf` | PDF merging (Doc Prep & Print) |
| `pywin32` | Windows COM + print API |
| `pyinstaller` | Building the `.exe` |
| `es.exe` | Everything CLI for fast file search (BOM Check) |
| SolidWorks (installed) | SW Batch Update COM API |
| Microsoft Excel (installed) | BOM Check xlwings backend |

Install Python dependencies:
```
pip install xlwings pypdf pywin32 pyinstaller
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
├── tools\
│   ├── BomCheck\
│   │   ├── main.py                 # Original standalone BOM script
│   │   └── es.exe                  # Everything CLI binary
│   ├── DocPrepPrint\
│   │   ├── DocPrepPrint.py         # Original print script
│   │   └── DocPrepPrint_Test(...).py  # Original simulation script
│   └── SolidworksBatchUpdate\
│       └── SoldworksBatchUpdate.bas   # Original VB macro (reference)
└── logs\                           # Created at runtime
    └── ETH_master_YYYY-MM-DD.log
```

---

## Notes

- The SW Batch Update tool is a Python rewrite of the original VB macro. It requires a machine with SolidWorks installed for live testing.
- BOM Check requires Excel to be **closed** — xlwings opens the file via COM in the background.
- The preferred printer name is hardcoded in `app.py` as `PREFERRED_PRINTER`. Update this constant if the printer name changes.
