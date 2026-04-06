---
name: print-help
description: >
  Quick-print the electrical drawing pack (pages 1-2) and PRF for a new FoxFab job. Use this skill whenever the user
  mentions starting a new job, having a new job, beginning work on a job number, or wants to print the electrical
  drawing and PRF for a job (e.g., "I have a new job J16204", "I'm starting a new job J15302", "new job J16204",
  "print help for J16204", "print the electrical and PRF for J16204").
---

# Print Help Skill

This skill quickly prints the two key documents an engineer needs when starting a new job:
1. **PRF (Production Release Form)** — first sheet only, printed on **letter** paper
2. **Electrical Drawing Pack** — first 2 pages only, printed on **tabloid (11x17)** paper

Both documents are printed simplex on `\\NPSVR05\FoxFab (Konica Bizhub C360i)`.

The bundled output PDF (for viewing) combines both into one file. When printing, they are sent as
separate print jobs with correct paper sizes (letter for PRF, tabloid for electrical).

## Trigger Phrases

- "I have a new job J16204"
- "I'm starting a new job J15302"
- "new job J16204"
- "starting J16204"
- "print help for J16204"
- "print the electrical and PRF for J16204"
- Any phrase mentioning a new job with a J##### number

## Variant Selection — PRF-Based

Variants are selected by **PRF**, not by folder structure. If the job has multiple PRF files
(e.g., `J15302-01 PRF.xlsx`, `J15302-02 PRF.xlsx`), the script raises `CHOOSE_PRF` and Claude
asks the user which PRF to use via AskUserQuestion. The selected PRF's model number (cell G9) is
then used to match the correct electrical pack PDF.

## Model-Aware Electrical Pack Matching

The script reads the PRF (cell G9 = model number), then uses the model number to pick the correct
electrical pack PDF when there are multiple PACK files in `102 Drawings`. If the model number matches
one pack exactly, that one is used. If no model match is found and multiple packs exist, the script
raises `CHOOSE_PACK` so Claude can ask the user.

## Output Directory

By default, generated PDFs are saved to `.claude/skills/print-help/output/`. The bundled file is named
`{job_number}_Print_Help.pdf`. Use `--output` to override.

## Folder Structure

The script uses the same job folder resolution as bom-filler and doc-prep (Patterns 1-4).

### Document Paths (relative to job root)
- Electrical Pack: `100 Elec/102 Drawings/` — PDF with "PACK" in filename
- PRF: `300 Inputs/302 Production Release Form/` — Excel file with "prf" in filename

## How Claude Should Use This Skill

### Parsing the User's Request
1. Extract the job number (J##### pattern) from the user's message

### Running the Script
```bash
# Step 1: Simulation mode (generates bundled PDF, no printing)
python .claude/skills/print-help/scripts/printhelp.py --job J16204

# With specific PRF selection (after CHOOSE_PRF)
python .claude/skills/print-help/scripts/printhelp.py --job J15302 --prf "J15302-01 PRF.xlsx"

# Override jobs root (for testing)
python .claude/skills/print-help/scripts/printhelp.py --job J16204 --jobs-root "//NPSVR05/path/to/Testing"

# Step 2: After user confirms, print (PRF on letter, electrical on tabloid)
python .claude/skills/print-help/scripts/printhelp.py --job J16204 --print
```

### Workflow for Claude
1. Parse the job number from the user's message
2. Run the script **without** `--print` first (simulation mode)
3. If `CHOOSE_PRF` error: use **AskUserQuestion** with PRF filenames as clickable options, then re-run with `--prf "filename"`
4. If `CHOOSE_PACK` error: use **AskUserQuestion** with pack filenames as clickable options
5. If `CHOOSE_FOLDER` error: use **AskUserQuestion** with folder names as clickable options
6. Show the user a summary of what will be printed (as a markdown table)
7. Ask for confirmation to print using **AskUserQuestion**
8. If confirmed, run again **with** `--print` (and `--prf` if previously selected)

### Error Handling
- `CHOOSE_FOLDER:` — multiple job folders matched; ask user via AskUserQuestion
- `CHOOSE_PRF:` — multiple PRF files found; ask user via AskUserQuestion
- `CHOOSE_PACK:` — multiple PACK PDFs, no model match; ask user via AskUserQuestion
- Missing electrical pack or PRF: warn but continue with whatever is available

## Printing Details

When `--print` is used, documents are sent as **separate print jobs** with correct paper sizes:
- **PRF** → letter (8.5 x 11) via `DMPAPER_LETTER = 1`
- **Electrical Pack** → tabloid (11 x 17) via `DMPAPER_TABLOID = 3`

Paper size is set in the printer devmode before each job. Both are simplex.

## Dependencies
- **pypdf** — extracting pages 1-2 from electrical pack PDF, merging bundle
- **openpyxl** — reading PRF model number for pack matching
- **pywin32** — Acrobat COM printing, Excel COM for PRF-to-PDF export, devmode paper size
- **Adobe Acrobat 2017+** — print dispatch
- **Microsoft Excel** — PRF export to PDF

## Configuration
```python
PREFERRED_PRINTER = r"\\NPSVR05\FoxFab (Konica Bizhub C360i)"
JOBS_ROOT         = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"
DMPAPER_LETTER    = 1   # 8.5 x 11
DMPAPER_TABLOID   = 3   # 11 x 17
```
