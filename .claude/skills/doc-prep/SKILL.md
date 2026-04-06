---
name: doc-prep
description: >
  Automate FoxFab manufacturing document preparation and printing. Use this skill whenever the user mentions doc prep,
  printing a job packet, preparing documents for a job number (e.g., "do docprep for J16204", "prep J15302",
  "prep this for vikram"), checking a BOM against CNC files (e.g., "check the bom for J16204"), or anything related
  to assembling FWO, BOM, CNC, flats, PRF, electrical packs, or assemblies from the FOXFAB_DATA job folder structure.
  Also use this skill when the user wants to add or modify folder structure patterns for the doc prep system, or
  troubleshoot doc prep errors.
---

# Doc Prep Skill

This skill automates FoxFab's manufacturing document preparation. It locates job folders, gathers all required
documents, auto-fills the Fabrication Work Order (FWO) from PRF data, marks the CNC column in the BOM, generates
ordered PDFs, and prints them.

## Two Workflows

### 1. Standard Doc Prep
**Trigger phrases:** "do docprep for J#####", "prep J#####", "doc prep J#####"

Flow:
1. Navigate to `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS`
2. Find the job folder by matching the job number prefix (e.g., `J16204` matches `J16204 DataBank -Docking Station- IAD5`)
3. If multiple folders match, ask the user which one
4. Resolve the folder structure (see Folder Structure Patterns below)
5. Build the document plan — gather all files
6. Generate all PDFs in simulation mode (to a timestamped output folder)
7. Show the user a breakdown by document type
8. Wait for user confirmation
9. Send to printer: `\\NPSVR05\FoxFab (Konica Bizhub C360i)`

### 2. Vikram Workflow (Reference Job)
**Trigger phrases:** "prep this for vikram", "vikram J##### reference J#####"

This workflow uses **two jobs**:
- **Main job** — provides: FWO, Electrical Drawing Pack, PRF
- **Reference job** — provides: BOM, CNC files, PDFs_Flats, Assemblies

The FWO is still auto-filled using the **main job's** PRF data. The reference job's BOM still gets CNC-marked
using the reference job's CNC folder. Everything else about document generation and ordering stays the same.

Flow:
1. Find both job folders in `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS`
2. Resolve folder structures for both jobs
3. Gather documents from the appropriate job (main vs reference)
4. Generate all PDFs in simulation mode
5. Show breakdown with clear labels of which job each document came from
6. Wait for user confirmation
7. Send to printer

## Folder Structure Patterns

The tool needs to find the mechanical subfolders (`204 BOM`, `205 CNC`, `202 PDFs_Flats`, `203 Assemblies`).
Here are the known patterns, checked in order:

### Pattern 1: Variant-Selected
User selected a numbered variant directly (e.g., `J15302-01`):
```
J15302-01/
  204 BOM/
  205 CNC/
  202 PDFs_Flats/
  203 Assemblies/
```
The job root is two levels up (grandparent of the variant folder).

### Pattern 2: Standard 200 Mech
The standard subfolders exist directly inside `200 Mech`:
```
J15302/
  200 Mech/
    204 BOM/
    205 CNC/
    202 PDFs_Flats/
    203 Assemblies/
```

### Pattern 3: Numbered Variants
`200 Mech` contains numbered variant folders (`*-01`, `*-02`):
```
J15302/
  200 Mech/
    J15302-01/
      204 BOM/
      205 CNC/
      202 PDFs_Flats/
      203 Assemblies/
    J15302-02/
      204 BOM/ ...
```

### Pattern 4: Named Subfolders (Internal/Enclosure)
`200 Mech` contains named subfolders like `Enclosure` and `Internal`. **Default to `Internal`** — that's
where the BOM, CNC, flats, and assemblies live.

This can appear at two levels:
- Directly inside `200 Mech`:
```
J16204/
  200 Mech/
    Enclosure/
    Internal/
      204 BOM/
      205 CNC/
      202 PDFs_Flats/
      203 Assemblies/
```
- Inside a numbered variant:
```
J16204/
  200 Mech/
    J16204-01/
      Enclosure/
      Internal/
        204 BOM/
        205 CNC/ ...
```

### Adding New Patterns
When a new folder variant is encountered and the build plan fails, update `_dpp_get_context()` in `app.py`
and the corresponding `get_selected_context()` in `DocPrepPrint.py`. The function should:
1. Check for the new pattern after the existing ones
2. Return the same `{"job_root", "mech_roots", "variant_only"}` dict
3. If the pattern is ambiguous (multiple possible subfolders), ask the user which one to use

## Document Order and Sources

Documents are generated and printed in this exact order:

| # | Document | Source Folder | Duplex | Notes |
|---|----------|--------------|--------|-------|
| 1 | FWO (Fabrication Work Order) | `300 Inputs/` | Simplex | Auto-filled from PRF data if available |
| 2 | BOM | `204 BOM/` | Simplex | All sheets, CNC column marked first |
| 3 | CNC (duplex files) | `205 CNC/` | Duplex | Files starting with `J` or `NNN-` prefix |
| 4 | CNC (simplex files) | `205 CNC/` | Simplex | All remaining CNC PDFs merged into one |
| 5 | PDFs_Flats | `202 PDFs_Flats/` | Simplex | All flat PDFs merged into one |
| 6 | PRF | `302 Production Release Form/` | Simplex | First sheet only |
| 7 | Electrical Pack | `100 Elec/102 Drawings/` | Simplex | Pages 1-2 only, PDF with "PACK" in name |
| 8 | Assemblies | `203 Assemblies/` | Simplex | All PDFs merged, `-LAY` files excluded |

## Key Business Logic

### FWO Auto-Fill
Reads PRF Excel fields via openpyxl (no COM needed):
- **C4** → Job No.
- **C8** → Job Name
- **G9** → Model No.
- **G18** → Size, **G19** → Material, **G20** → Rating, **G22** → Qty
- Material abbreviations: `aluminum/aluminium` → `ALU`, `stainless` → `SS`
- Rating abbreviations: `Type 3R` → `N3R`
- Enclosure = `"Size Material Rating"` (e.g., `72H ALU N3R`)

Overlay positions on FWO PDF (PyMuPDF, Helvetica 11pt):
```
JOB NO:    (165, 145)
JOB NAME:  (165, 165)
DATE:      (165, 182)    Format: "Month DD, YYYY"
ENCLOSURE: (165, 210)
UNITS:     (165, 245)
```

### CNC Column Marking
Before exporting BOM to PDF, scan `205 CNC` for PDFs and mark column H with `X`:
- **J-prefix** (GALV files) → extract part numbers from `DRAWING NUMBER:` lines inside the PDF
- **NNN-XXXXX prefix** (e.g., `240-90004`) → match as-is; additional `_DIGIT` segments = more part numbers
- **Bare digits** (e.g., `90004.pdf`) → prepend `240-`
- **Revision suffixes** (`rA`, `rB`) → stripped before matching
- Skip files with "Merged" in the name
- Skip rows where column H = `S` (stock parts)

### BOM Revision Selection
When multiple Excel files in `204 BOM`, sort by revision: `rA < rA1 < rA2 < rB < rB1 ...`

### Electrical Pack Selection (Model-Aware)
If PRF has a model number, prefer exact match in `100 Elec/102 Drawings`. Fall back to any PACK PDF, ask if ambiguous.

### PRF Variant Matching
For multi-variant jobs, prefer PRF with matching variant suffix (e.g., `-01`).

## Error Handling
Skip missing documents with a warning and continue. Never stop the entire run for a single missing file.

## Dependencies
- **xlwings** — BOM CNC marking, Excel-to-PDF export (COM)
- **PyMuPDF (fitz)** — FWO text overlay, GALV drawing number extraction
- **openpyxl** — PRF data reading (no COM)
- **pypdf** — PDF merging
- **pywin32** — Acrobat COM printing, duplex control
- **Adobe Acrobat 2017+** — print dispatch
- **Microsoft Excel** — BOM/PRF export

## Configuration
```python
PREFERRED_PRINTER  = r"\\NPSVR05\FoxFab (Konica Bizhub C360i)"
BOM_SHEET_NAME     = "FFMPL"
JOBS_ROOT          = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"
```

## How Claude Should Use This Skill

### Running Doc Prep
Use the standalone script at `.claude/skills/doc-prep/scripts/docprep.py`:

```bash
# Standard workflow — simulation mode (generates PDFs, no printing)
python .claude/skills/doc-prep/scripts/docprep.py --job J16204

# With variant selection
python .claude/skills/doc-prep/scripts/docprep.py --job J16204 --variant 02

# Vikram workflow — main job + reference job
python .claude/skills/doc-prep/scripts/docprep.py --job J16204 --ref J15302

# With printing (ONLY after user explicitly says "print" or "send it")
python .claude/skills/doc-prep/scripts/docprep.py --job J16204 --variant 02 --print

# With explicit electrical pack selection
python .claude/skills/doc-prep/scripts/docprep.py --job J16204 --pack "Z:\path\to\pack.pdf"

# BOM Check — read-only comparison of BOM vs CNC folder
python .claude/skills/doc-prep/scripts/docprep.py --job J16204 --check-bom
```

### Workflow for Claude

1. Parse the user's request to extract job number(s) and detect which mode:
   - **"check the bom for J#####"** → BOM check mode
   - **"do docprep for J#####"** → standard doc prep
   - **"prep this for vikram, J##### ref J#####"** → Vikram workflow

2. **ALWAYS run in simulation mode first** (no `--print`). NEVER pass `--print` unless the user
   explicitly says "print", "send it", "print it", or similar. After showing the breakdown, do NOT
   ask "Ready to print?" — just show the results and wait for the user to say "print" on their own.

3. **Handle script errors by parsing the error prefix:**

   - **`CHOOSE_VARIANT:`** — Multiple variants detected. The error message contains comma-separated
     variant folder names after the prefix. Present these as clickable AskUserQuestion options. If
     a PRF exists, check which variant it matches and mark that as "(Recommended)" in the first
     option position. Re-run the script with `--variant <selected>`.

   - **`UNKNOWN_STRUCTURE:`** — Folder pattern not recognized. The error message includes a
     directory tree. Display the tree to the user and ask them to identify which subfolder(s) to
     use for BOM/CNC/Flats/Assemblies via AskUserQuestion. Then guide the user on how to proceed
     (may need manual folder restructuring or a new pattern added to the code).

4. Show the breakdown output formatted as markdown (see formatting rules below).

5. For BOM check: run with `--check-bom`, format the report as a markdown table in chat, AND
   mention the saved report file path.

### Breakdown Formatting (Chat Presentation)

Present results as clean markdown, NOT raw script output:

**Standard layout:**
- Bold header: `**DOC PREP — {job_no} {job_name}**`
- Job info as a small markdown table (Job No., Enclosure, Model No., Qty)
- BOM summary as one bold line: `**BOM:** X parts | Y CNC marked | Z stock | W unmatched`
- Unmatched parts with row numbers and categories:
  - `(flexibar)` for N/A parts (245-/295- prefix)
  - `(stock)` for stock-found parts
  - `(missing CNC)` for truly unmatched parts
- Documents as a markdown table with columns: #, Document, Mode (simplex/duplex)
- Show which variant was selected (e.g., "Variant: J15302-02")
- Excluded LAY files count
- Output path in backticks

**Vikram layout:**
- Same as standard but header includes `(ref: {ref_job})`
- Documents table adds a **Source** column showing which job each doc comes from

### If dependencies are missing
Tell the user to install them:
```
pip install pywin32 pypdf pymupdf openpyxl xlwings
```

## Code Locations
- `.claude/skills/doc-prep/scripts/docprep.py` — standalone headless script (used by this skill)
- `tools/DocPrepPrint/DocPrepPrint.py` — standalone GUI script
- `app.py` — main GUI app with integrated doc prep

When modifying doc prep logic, keep all three in sync.
