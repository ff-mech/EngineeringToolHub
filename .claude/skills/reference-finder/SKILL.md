---
name: reference-finder
description: >
  Find reference jobs by matching a job's PRF (model number, enclosure size, amperage) against the
  Quote Index, then comparing electrical BOMs for component overlap. Use whenever the user asks
  "find references for J#####", "reference jobs for J16204", "similar jobs", or any phrasing about
  finding past jobs to use as a reference for a new one.
---

# Reference Job Finder Skill

This skill answers "find me reference jobs for J#####" by:

1. Reading the job's PRF to extract model number, enclosure size, and amperage.
2. Searching the Quote Index for shipped/won jobs with matching product family, configuration type, amperage (+-200A), and enclosure size (+-2" per dimension).
3. Validating that candidate jobs have a populated `200 Mech` folder.
4. Comparing electrical BOMs by component type/description overlap.
5. Presenting the top 3 reference jobs with match details.

## Trigger Phrases

- "find references for J15874"
- "reference jobs for J16204"
- "find similar jobs to J#####"
- "what jobs can I use as reference for this job"
- Any phrase asking about finding reference jobs or similar past jobs

## Required Job Number

A job number is **required**. If the user doesn't provide one, ask via **AskUserQuestion** before running the script.

## Workflow for Claude

### Step 1 -- Run the script

```bash
python .claude/skills/reference-finder/scripts/reffinder.py --job J15874
```

For testing, use the `--jobs-root` flag:

```bash
python .claude/skills/reference-finder/scripts/reffinder.py --job J15874 \
    --jobs-root "//NPSVR05/FOXFAB_REDIRECT$/lbadong/Desktop/AGENT ARMY/Testing"
```

The script:
- Resolves the job folder (same patterns as print-help / bom-filler)
- Reads the PRF from `300 Inputs/302 Production Release Form/`
- Extracts: Model No (C5), Enclosure Size (G18), Current (C11)
- Parses the model number into: product family, config type, amperage
- Searches the Quote Index (`\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\Quote Index - LAMBERT.xlsx`, sheet `Index`)
- Filters to Status (col Q) containing "Shipped" or "Won"
- Matches product family, config type, amperage (+-200A), enclosure size (+-2" per dimension)
- For each candidate, checks if the job folder exists and `200 Mech` has files
- If both the source job and candidate have electrical BOMs, compares them by DESCRIPTION column overlap
- Returns a JSON result with up to 10 candidates ranked by match score

### Step 2 -- Handle script output

- `UNKNOWN_JOB:` -- job folder not found. Ask the user to confirm the job number.
- `NO_PRF:` -- PRF file not found. Tell the user.
- `NO_MATCHES:` -- no matching jobs found. Tell the user and suggest relaxing criteria.
- Normal JSON output -- proceed to Step 3.

### Step 3 -- BOM visual comparison (top 10 candidates)

The script returns up to 10 candidates (each with a `bom_image` path if available) so that the final
selection of 3 references is driven by visual BOM similarity, not just model-code matching.

The electrical BOMs are DWG-embedded xlsx files (same format as parts-needed skill). The script
extracts EMF previews and converts them to PNG images in `.claude/skills/reference-finder/output/<JOB>/`.
Existing PNGs are reused from prior runs.

Workflow:

1. Read the **source BOM image** with the Read tool to extract component descriptions.
2. Read each of the **up to 10 candidate BOM images** with the Read tool.
3. Compare the component types between source and each candidate to determine overlap.
4. Use these categories: breakers, terminal blocks, lugs, heaters, meters, relays, transfer switches,
   disconnect switches, transformers, motors, shunts, fuses, interlocks, camlocks, etc.
5. Pick the **top 3 candidates** whose BOMs most closely match the source.

If a candidate has no `bom_image` (empty BOM or non-DWG xlsx), mark BOM overlap as "N/A" and rely on
model-code matching alone for that candidate.

### Step 4 -- Present the final 3 references

Show the **top 3** (chosen via BOM similarity from the 10 candidates) as a **markdown table** with columns:

| Rank | Job # | Job Name | Enclosure Size | Amperage | BOM Overlap | Matching Components |
|------|-------|----------|----------------|----------|-------------|---------------------|

- BOM Overlap = percentage of component types shared between source and reference
- Matching Components = list of shared component type categories (e.g., breakers, terminal blocks, lugs)
- If BOM data is unavailable for a candidate, show "N/A" for BOM columns

After the table, show the source job's PRF details for context:
- Model No, Enclosure Size, Amperage

## Important Notes

- The job folder at `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS` is **read-only** -- never create, modify, or delete anything there.
- BOMs are always at `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\J#####\100 Elec\101 Bill of Materials\`
- BOMs are typically DWG-embedded xlsx files -- the xlsx sheet is empty but contains an OLE-embedded AutoCAD drawing with an EMF preview image.

## Configuration

```python
QUOTE_INDEX    = r"\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\Quote Index - LAMBERT.xlsx"
JOBS_ROOT      = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"  # production path, read-only
PRF_SUBPATH    = r"300 Inputs\302 Production Release Form"
BOM_SUBPATH    = r"100 Elec\101 Bill of Materials"
MECH_SUBPATH   = r"200 Mech"
INDEX_SHEET    = "Index"
```

## Dependencies

- **openpyxl** -- read xlsx files (PRF, Quote Index, BOMs)
- **Pillow** -- EMF to PNG conversion
- **olefile** -- OLE embedding detection (optional, for DWG detection)
- Python stdlib `re`, `json`, `argparse`, `pathlib`, `glob`, `os`, `zipfile`

## Testing

Test job folder: `\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\Testing` (per memory).

```bash
python .claude/skills/reference-finder/scripts/reffinder.py --job J16567
# or for testing:
python .claude/skills/reference-finder/scripts/reffinder.py --job J16482 \
    --jobs-root "//NPSVR05/FOXFAB_REDIRECT$/lbadong/Desktop/AGENT ARMY/Testing"
```
