# Changelog

All notable changes to Engineering Tool Hub are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [1.1.0] — 2026-03-30

### Added

**How to Use tab**
- New sidebar tab that launches on startup by default
- Embedded PDF viewer renders `Engineering_Tool_Hub.pdf` directly inside the app using PyMuPDF (fitz) — no external viewer required
- Prev / Next page navigation with a live page counter
- Vertically scrollable canvas with mousewheel support

**SW Batch PDF Export tab**
- Step-by-step guide for using SolidWorks Task Scheduler to batch-export drawings as PDFs
- Tips card covering folder-mode export, sub-folder recursion, and paper size behaviour

**Training Materials tab**
- Lists all documents in `tools/EngineeringDesignPackage/` and prints them double-sided to the preferred printer in a single background thread
- Per-file status label updates in real time while printing; button disabled until the job completes
- Reuses existing `_dpp_acrobat_print` + `_dpp_set_devmode_duplex` infrastructure — no paper size override, letting the printer decide
- "Open FoxFab Design Tips" button opens `tools/FoxFab_Design_Tips.docx` in the default application

### Changed

- **Tab order redesigned** for workflow intuitiveness: How to Use → Bom Filler → Doc Prep & Print → SW Batch Update → SW Batch PDF Export → File Logger → Training Materials
- **Default landing tab** changed from Bom Filler to How to Use so new users are immediately oriented

---

## [1.0.0] — 2026-03-27

### Added

**Application**
- Combined app (`app.py`) wrapping three manufacturing tools into a single Windows desktop application
- Dark navy sidebar navigation — switch between tools instantly without re-launching
- Per-tool terminal output with colour-coded log levels (`[INFO]`, `[OK]`, `[WARN]`, `[ERROR]`)
- Per-tool progress bar and **Stop** button (hard stop via `threading.Event`)
- Inline confirm bars for plan review and print gate — no modal popups
- Master log appended to `logs/ETH_master_YYYY-MM-DD.log` after each run
- Splash screen on startup shown before heavy imports to signal the app is loading
- Lazy panel building — only the active tool panel is constructed at launch

**Bom Filler**
- Two-pass BOM processor: stock parts check (Pass 1) then non-stock PDF/DXF copy (Pass 2)
- Everything CLI (`es.exe`) integration for fast indexed file search
- Revision detection — picks highest `rA/rB/rC…` revision automatically
- Handles both space-separated and dash-separated revision naming conventions
- Excel markup via xlwings COM (preserves all formatting, table styles, named ranges)

**Doc Prep & Print**
- Full manufacturing packet builder from a job variant folder
- CNC column auto-marking — scans `205 CNC\` folder and writes `X` to column H before print
- FWO auto-fill — reads job/model/enclosure/qty from PRF Excel via openpyxl, overlays text with PyMuPDF
- BOM revision auto-selection — picks highest-revision Excel file automatically
- Acrobat COM single-instance printing — one `AcroExch.App` session for all documents, guaranteed print order
- Per-document duplex control via DEVMODE (no admin rights required)
- Fallback to `subprocess /t` + spooler polling if Acrobat COM is unavailable
- Simulation Mode — saves all documents to a timestamped local folder instead of printing
- **Preview FWO** button — fills and opens FWO PDF without a full print run
- **Preview BOM** button — runs CNC marking, exports BOM to PDF, opens it
- Manual Printing panel — auto-populated after Generate Documents; supports add/remove/reorder

**SW Batch Update**
- Reference panel displaying the SolidWorks VB macro path and run instructions

**Build**
- `build.bat` — one-step PyInstaller `--onedir` build with dependency install, clean, and es.exe bundling
