---
name: parts-needed
description: >
  Extract the purchased-parts list (P/N, Description, Qty) from a FoxFab job's BOM with a two-pass
  Opus + Sonnet cross-check for accuracy. Use whenever the user asks "what parts do I need for J#####",
  "what do I need for this job", "show me the parts needed", "parts list for J16204", or any similar
  phrasing about what parts/materials are required for a job.
---

# Parts Needed Skill

This skill answers "what parts do I need for this job?" by:

1. Locating the job's PRF (same logic as `print-help`).
2. Picking the BOM file in `100 Elec/101 Bill of Materials/` whose **filename** most closely matches the PRF filename (top level only — no subfolders).
3. Extracting every line item that has a Part Number (P/N + Description + Qty).
4. Cross-checking the extraction with a second independent pass (Opus 4.6 → then Sonnet 4.6 subagent) to assign per-row and overall confidence ratings.
5. Writing the result to `.claude/skills/parts-needed/output/<JOB>/<JOB>_parts.txt`.

All per-job artifacts (parts txt, BOM CSV, rendered BOM PNGs/panels) are grouped into a single per-job subfolder: `.claude/skills/parts-needed/output/<JOB>/`. The script auto-creates this folder and writes all EMF/PDF renderings there.

## Trigger Phrases

- "what parts do I need for J16204"
- "what do I need for this job"
- "show me the parts needed for J15302"
- "parts list for J16204"
- "what's needed for this job"
- Any phrase asking what parts/materials/components are needed for a J##### job

## Required Job Number

A job number is **required**. If the user doesn't provide one, ask via **AskUserQuestion** before running the script.

## Workflow for Claude

### Step 1 — Run the script (deterministic extraction)

```bash
python .claude/skills/parts-needed/scripts/partsneeded.py --job J16204
```

The script:
- Resolves the job folder (Patterns 1–4, same as print-help / bom-filler)
- Finds the PRF in `300 Inputs/302 Production Release Form/`
- Lists candidate BOMs in `100 Elec/101 Bill of Materials/` (top level only — `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.pdf`)
- Scores each BOM filename against the PRF filename via `difflib.SequenceMatcher`
- If the top score ≥ 0.85 and clearly the winner → auto-picks
- Otherwise emits `CHOOSE_BOM:<json>` so Claude can ask the user
- Extracts rows from the chosen BOM and prints a JSON blob to stdout

### Step 2 — Handle script errors

- `UNKNOWN_JOB:` — job folder not found. Ask the user to confirm the job number.
- `NO_BOM_FOUND:` — `101 Bill of Materials` empty or missing. Tell the user.
- `CHOOSE_BOM:<json>` — multiple close matches. Use **AskUserQuestion** with the candidate filenames as clickable options, then re-run with `--bom "<full path>"`.
- `PDF_NEEDS_VISUAL:<path>` — PDF text extraction failed. Use `mcp__claude_ai_PDF_Viewer__display_pdf` to read the BOM visually, then continue.

### Step 3 — Two-pass extraction (parallel subagents)

When the script returns `EMF_IMAGE:` or `PDF_IMAGE:` (image-based BOM), spawn **both passes in parallel** as subagents in a single message with two `Agent` tool calls. Do NOT read the panels in the main context — offload all image work to the subagents so the main thread stays light and both passes run concurrently.

**Pass 1 — Opus subagent:** `subagent_type: "general-purpose"`, `model: "opus"`.
**Pass 2 — Sonnet subagent:** `subagent_type: "general-purpose"`, `model: "sonnet"`.

Give each subagent the identical prompt:
- The list of panel PNG paths (and the full-page PNG) from the script's payload.
- Column layout (REF, MFR, P/N, DESCRIPTION, QTY, REV).
- "Extract every row with a non-empty P/N. A single REF with multiple stacked P/Ns → one output row per P/N."
- "Return ONLY a JSON list of `{ref, mfr, pn, desc, qty}`. No commentary."
- "Work independently. Do not assume any prior extraction exists."

For text-extracted BOMs (script returns JSON rows directly, no image error), skip Pass 1 subagent — use the script's rows as `rows_opus` — and only spawn the Sonnet subagent for Pass 2.

### Step 4 — Merge and assign confidence

Match rows from the two passes by normalized P/N:

| Condition                                                | Confidence |
|----------------------------------------------------------|------------|
| Both passes agree on P/N, Desc, and Qty                  | **HIGH**   |
| Both have P/N + Qty match but Desc differs               | **MED**    |
| Only one pass has the row, OR Qty differs between passes | **LOW**    |

**Overall %** = `weighted_score / max_score`, where HIGH = 1.0, MED = 0.7, LOW = 0.3, missing = 0.

### Step 5 — Write output txt

Write to `.claude/skills/parts-needed/output/<JOB>/<JOB>_parts.txt` in this format:

```
Parts Needed — J16204 DataBank Docking Station IAD5
BOM source : <bom filename>
Generated  : 2026-04-07 14:32
Overall confidence : 94%  (Opus 4.6 + Sonnet 4.6 cross-check)

CONF  P/N              QTY   DESCRIPTION
HIGH  ABB-1SDA054123   2     MCCB Tmax XT1B 160 25kA
HIGH  WAGO-2002-1201   48    Terminal block 2.5mm² grey
MED   PHX-3044131      12    End cover (desc differs between passes)
LOW   SCH-LV429550     1     Qty mismatch: Opus=1 Sonnet=2  ← REVIEW
```

### Step 6 — Show the user

After writing the txt, present the rows to the user as a **markdown table** (per memory: doc-prep breakdown format preference) and link to the output file.

### Step 7 — Copy matching models from MODEL LIBRARY (automatic)

This step runs **automatically** after Step 6. It creates
`.claude/skills/parts-needed/output/<JOB> - Parts/` and copies matching
`.SLDPRT` / `.SLDASM` files from
`\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\MODEL LIBRARY`
(searched recursively). Learned choices persist across jobs in
`.claude/skills/parts-needed/mappings.json`; rejected-with-no-match P/Ns go
into `.claude/skills/parts-needed/ignore_list.json`.

1. **Build the P/N working list.** Include every HIGH and MED row
   automatically. For each LOW row, use **AskUserQuestion** (one question per
   LOW row, batched up to 4 per call) asking whether to include it. Excluded
   LOW rows are reported as `SKIPPED (LOW)` in the final summary.

2. **Dry-run plan.** Write the working P/Ns to a temp JSON
   (`output/<JOB>/_pns.json`, list of strings) and run:

   ```bash
   python .claude/skills/parts-needed/scripts/modelcopy.py \
       --job J##### --pns-json .claude/skills/parts-needed/output/J#####/_pns.json --plan
   ```

   Each P/N comes back with one of: `exact`, `mapped`, `ignored`,
   `ambiguous` (with up to 5 candidate paths), or `none`.

3. **Resolve ambiguous / none via AskUserQuestion.**
   - For each `ambiguous` P/N: present the candidate filenames (not full paths)
     as options, plus a trailing "None of these" option.
   - For each `none` P/N (or when user picks "None of these"): present
     `[Add to ignore list (persistent), Skip this job only]`.
   - Batch up to 4 questions per AskUserQuestion call for throughput.

4. **Apply choices.** Write the resolved choices to
   `output/<JOB>/_choices.json` — a flat object keyed by normalized P/N
   (uppercase, stripped of spaces/`-_./\\`) → value is either the chosen
   absolute path, the literal string `"IGNORE"`, or `"SKIP"`. Then:

   ```bash
   python .claude/skills/parts-needed/scripts/modelcopy.py \
       --job J##### --pns-json .../_pns.json --resolve .../_choices.json
   ```

   The script copies the chosen SLDPRT/SLDASM files, updates
   `mappings.json` and `ignore_list.json`, and prints a summary JSON.

5. **Show the user a markdown table** with columns
   `Status | P/N | File`. Status values: `EXACT`, `MAPPED`, `CHOSEN` (new
   mapping just learned), `IGNORED` (was on ignore list), `NEWLY_IGNORED`,
   `SKIPPED (LOW)`, `SKIPPED (this job)`, `ALREADY_PRESENT`.

Link to the `<JOB> - Parts` folder at the end so the user can open it.


## Folder Structure

Reuses Patterns 1–4 from `print-help` / `bom-filler` / `doc-prep`. Documented in `printhelp.py:get_context()`.

Relevant paths (relative to job root):
- PRF: `300 Inputs/302 Production Release Form/` — Excel with "prf" in filename
- BOM: `100 Elec/101 Bill of Materials/` — top level only

## Error Handling

- `UNKNOWN_JOB:` — no folder matches the job number
- `NO_BOM_FOUND:` — BOM folder empty or missing
- `CHOOSE_BOM:` — multiple close BOM filename matches; ask user
- `PDF_NEEDS_VISUAL:` — PDF text extraction returned nothing; fall back to PDF Viewer MCP

## Dependencies

- **openpyxl** — read xlsx/xlsm BOMs
- **pdfplumber** — primary PDF table extraction
- **pymupdf (fitz)** — secondary PDF text fallback
- Python stdlib `csv`, `difflib`, `json`, `argparse`, `re`, `pathlib`

## Configuration

```python
JOBS_ROOT      = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"
BOM_SUBPATH    = r"100 Elec\101 Bill of Materials"
PRF_SUBPATH    = r"300 Inputs\302 Production Release Form"
BOM_EXTS       = {".xlsx", ".xlsm", ".xls", ".csv", ".pdf"}
AUTO_PICK_MIN  = 0.85   # filename match ratio for auto-pick
AUTO_PICK_GAP  = 0.10   # winning score must beat 2nd by this much
```

## Testing

Test job folder: `\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\Testing` (per memory).

```bash
python .claude/skills/parts-needed/scripts/partsneeded.py --job J16204 \
    --jobs-root "//NPSVR05/FOXFAB_REDIRECT$/lbadong/Desktop/AGENT ARMY/Testing"
```
