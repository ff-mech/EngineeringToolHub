---
name: bom-filler
description: >
  Fill and process FoxFab BOMs — stock part checking and PDF/DXF file copying. Use this skill whenever the user
  mentions filling a BOM, completing a BOM, BOM filler, stock check, or wants to process a BOM for a job number
  (e.g., "fill the bom for J16204", "/bom-filler J16204", "complete the bom at <path>", "run bom filler").
---

# BOM Filler Skill

This skill automates FoxFab's BOM processing workflow. It checks each part number against the Stock Parts folder,
copies the highest-revision PDFs and DXFs for non-stock parts into the target folder, and highlights any missing
parts with red fill in the BOM.

## Two-Pass Workflow

### Pass 1 — Stock Parts Check
- Reads part numbers from column A of the FFMPL sheet
- Searches the Stock Parts folder (`Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\300 Stock Parts\PDFs & Flats`) via Everything CLI
- Marks stock parts: column B = "X", column G = "S", column H = "S"
- Handles config-variant parts (strips `_###` suffix and retries with base number)

### Pass 2 — Non-Stock PDF + DXF Copy
- For each non-stock part, detects the highest revision independently for PDFs and DXFs
- PDFs support numeric sub-revisions (rA1, rA2, rB12); DXFs use letter-only revisions (rA, rB)
- Prefers standalone PDFs over combined-part files (e.g., `240-90123.pdf` over `240-90123_124.pdf`)
- Copies matching files to the target folder (typically `202 PDFs_Flats`)
- Marks column G = "X" if either PDF or DXF was found/copied
- Flexibar parts (245- prefix): skips DXF copy, marks column H = "N/A"
- Combined-part files: cross-marks all covered BOM rows
- Filters out PUNCH PROGRAM paths (grabs the clean copy from the job folder instead)
- **Missing parts**: rows where both PDF and DXF are not found get highlighted with red fill

## Trigger Phrases

- `/bom-filler J16204`
- "fill the bom for J16204"
- "complete the bom for J16204"
- "run bom filler on J16204"
- "fill the bom at Z:\path\to\BOM.xlsm"

## Input Options

The skill accepts either:
1. **Job number** (`--job J16204`): Auto-locates the BOM in the job folder structure
2. **BOM path** (`--bom "Z:\path\to\BOM.xlsm"`): Direct path to the workbook or its directory

An optional `--target` flag overrides the auto-detected target folder.

## Job Folder Resolution

When given a job number, the script searches `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\` and walks these folder
structure patterns to find the `204 BOM` directory:

### Pattern 1: Variant-Selected
User selected a numbered variant directly (e.g., `J15302-01`):
```
J15302-01/
  204 BOM/
  202 PDFs_Flats/
```

### Pattern 2: Standard 200 Mech
```
J15302/
  200 Mech/
    204 BOM/
    202 PDFs_Flats/
```

### Pattern 3: Numbered Variants
```
J15302/
  200 Mech/
    J15302-01/
      204 BOM/
      202 PDFs_Flats/
    J15302-02/
      204 BOM/ ...
```

### Pattern 4: Named Subfolders (Internal/Enclosure)
```
J16204/
  200 Mech/
    Internal/
      204 BOM/
      202 PDFs_Flats/
    Enclosure/
      204 BOM/ ...
```

Also nested inside numbered variants:
```
J16204/
  200 Mech/
    J16204-01/
      Internal/
        204 BOM/ ...
```

When multiple BOMs are found, the script lists them and prompts for selection via stdin.

## Missing File Handling

- Parts where both PDF and DXF are not found: entire row highlighted red in BOM, tagged `[MISSING]`
- Parts where only PDF or only DXF is missing: row highlighted red, tagged `[MISSING PDF]` or `[MISSING DXF]`
- A summary section at the end lists all missing parts with their row numbers

## BOM Sheet Layout

| Column | Name | Values |
|--------|------|--------|
| A (1) | Part Number | Part number string |
| B (2) | Stock Part | "X" if stock |
| G (7) | PDF | "S" (stock), "X" (found/copied) |
| H (8) | CNC | "S" (stock), "X" (marked), "N/A" (flexibar) |

- Header row: 5
- Data starts at row: 6
- Sheet name: FFMPL

## Dependencies

- **xlwings** — Excel COM automation (preserves formatting, formulas, data connections)
- **Everything CLI (es.exe)** — fast file search (auto-detected from project paths)
- **Microsoft Excel** — must be installed; BOM file must be CLOSED before running

## Configuration

```python
JOBS_ROOT          = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"
STOCK_PARTS_FOLDER = r"Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\300 Stock Parts\PDFs & Flats"
SHEET_NAME         = "FFMPL"
```

## How Claude Should Use This Skill

### Running the BOM Filler

Use the standalone script at `.claude/skills/bom-filler/scripts/bomfiller.py`:

```bash
# By job number — auto-locates BOM
python .claude/skills/bom-filler/scripts/bomfiller.py --job J16204

# By explicit BOM path
python .claude/skills/bom-filler/scripts/bomfiller.py --bom "Z:\path\to\BOM.xlsm"

# With custom target folder
python .claude/skills/bom-filler/scripts/bomfiller.py --job J16204 --target "Z:\path\to\target"
```

### Workflow for Claude

1. Parse the user's request to extract the job number or BOM path
2. Run the script and capture stdout
3. If the script prompts for variant selection (multiple BOMs found), ask the user which one and pipe the answer
4. Report the results to the user:
   - Pass 1 summary: how many stock parts found
   - Pass 2 summary: PDFs/DXFs copied, existed, not found
   - Missing parts list (if any) — these need the user's attention
5. If there are missing parts, suggest the user check whether the files exist or need to be created

### If the script prompts for input

When multiple BOM locations or job folders are found, the script outputs a numbered list and waits for stdin.
Use `echo "1" | python bomfiller.py --job J16204` to pipe the selection, or ask the user which option to choose.

### If dependencies are missing

Tell the user to install them:
```
pip install xlwings
```
Everything CLI (es.exe) must be installed and indexed, or bundled in the project's `tools/BomFiller/` directory.

## Code Locations

- `.claude/skills/bom-filler/scripts/bomfiller.py` — standalone headless script (used by this skill)
- `tools/BomFiller/main.py` — standalone GUI/CLI script (original tool)
- `app.py` — main GUI app with integrated BOM filler panel
