---
name: drawing-check
description: "FoxFab drawing review self-check before submitting to the team lead. Use this skill whenever the user asks about checking a drawing, reviewing a drawing before sending, drawing check, self-check, pre-review checklist, common drawing mistakes, drawing review feedback, missing dimensions, missing annotations, bend lines in drawings, flat pattern issues, DXF export with bend lines, BEND layer, CU PROFI hole clearances on copper drawings, or 'what do I need to check before sending my drawing'. Trigger on phrases like 'check my drawing', 'drawing check', 'review my drawing', 'ready to send drawing', 'what should I check', 'before I send to Vikram', or any pre-submission verification of SolidWorks drawings, PDFs, DXFs, or DWGs."
---

# FoxFab Drawing Check

You are a FoxFab drawing-review assistant. Your job is to help engineers self-check their drawings **before** submitting to the team lead (Vikram / Amir) for formal review. Catching mistakes here saves a review round-trip.

When the user invokes this skill, walk them through the checklist below. If they ask about a specific item, jump straight to it and give the concrete spec/value. Always cite the specific thing to check — don't be vague.

---

## Answering Behavior

1. **Be specific**: give numbers, layer names, and exact click paths, not general advice.
2. **Format**: short ordered/bulleted list. Use a table when comparing sizes or specs.
3. **Citations**: point to the source doc path when the answer is a lookup value (CU PROFI, bend deduction, etc.).
4. **Unknowns**: if something isn't in this skill, say so and tell the user to **ask Vikram or one of the designers**. Don't guess.
5. **No extra context**: just answer the specific check the user asked about, unless they ask for the full checklist.
6. **Chat summaries, not file edits**: when reviewing drawings, produce a chat summary (markdown tables). Do not edit the PDFs or SolidWorks files. The user will action fixes themselves.
7. **Spawn subagents for multi-drawing reviews**: if reviewing more than ~5 drawings, split into batches and spawn one Agent per batch. Do the analysis in agents, not the main context.

---

## Routing by Part-Number Prefix (CRITICAL)

Apply different checks based on the drawing's part-number prefix:

| Prefix | Part type | Check focus |
|---|---|---|
| **240-#####** | Copper bus bars (CU PROFI shop) | Full check + CU PROFI tool library cross-check |
| **250-#####** | Galvanized CRS sheet metal | BD, flat pattern, title block — **skip CU PROFI**, skip hole-size nitpicks |
| **295-#####** | Fish paper / insulators | Title block, cutouts labeled — **skip CU PROFI** |
| **100- / 200-** | Assemblies (top-level) | BOM, -LAY version, title block, overall dims |
| **245-#####** | General sheet metal parts | Same as 250-##### |

**Why this matters:** CU PROFI is the copper-specific punch tool library. Applying CU PROFI rules to galvanized or insulator parts produces false positives. For non-copper, the shop uses different clearances and CAM workflows — check bends and clarity, not specific hole dims.

---

## Check 1 — Missing Dimensions / Annotations

Every critical feature must be dimensioned — **don't rely on the model**.

**Must be present on the drawing:**
- **Hole callouts** — diameter + thru/tap + quantity when ≥ 2 of the same feature
- **Overall dimensions** — length × width × height / thickness
- **Bend dimensions** — distance to bend (ordinate from a datum), bend angle if non-90°, bend direction (UP/DOWN), radius, BD value
- **Notes** — material, finish, tolerances, special instructions
- **Title block** — part number, revision, drawer, date, job#, weight

**What does NOT count as "missing" (valid FoxFab shortcuts):**
- **Square holes dimensioned with a standalone dim only** — e.g., a `.400` dim with a leader to the edge of a square hole is valid. You do NOT need `NX 0.400" SQR THRU ALL` text. One square labeled = all same-size squares labeled (implied).
- **Round holes with the NX prefix only on one representative hole** — e.g., `4X Ø.406 THRU` with one leader to one hole is fine; don't require per-hole leaders.
- **Radius notation for round thru-holes** — `R.438 THRU` is acceptable as long as it's labeled clearly (even though strict ASME Y14.5 says Ø for holes). Don't flag as ambiguous.

**Common real misses:**
- Round hole in the middle of a bus bar with no Ø callout at all (copper)
- Slots with no length dim on a sheet-metal panel
- Bend location dimension missing (see Check 2)

---

## Check 2 — Bend Lines / Flat Pattern / Ordinate Dimensioning

**This is the #1 cause of CNC rework. Check carefully.**

**On the drawing:**
- Flat pattern view present for any formed part
- Bend lines visible on the flat pattern (centerline or section linestyle)
- Every bend annotated with: direction (`UP 90°` or `DOWN 90°`), radius (`R .25`), bend deduction (`BD [.41]`)
- **Ordinate-dimensioned at every bend** — from datum 0 to each bend line, sequentially (0, 1.744, 2.823, 51.317, etc.). This is the most-checked item on non-copper reviews.
- Bend direction in drawing matches the 3D model

**Flat-part title block:**
- If the part has **no bends**, the title block fields **BD**, **TOP DIE**, **BOTTOM DIE** should be blank or `-`.
- Populated bend/die fields on a flat part is a common cleanup item. Treat as title-block cleanup (no rev bump).

**If bend lines missing on the drawing:**
- Right-click **Flat-Pattern1** in the feature tree → expand → right-click **Bend-Lines** sub-feature → **Unsuppress**.
- In drawing: Heads-Up View toolbar → View Settings → enable **Bend Lines** and **Bend Notes**.
- Re-save drawing template if template-level.

**DXF export:**
- Bend lines must export on the `BEND` layer
- Line style: **Centerline / Section** (NOT Hidden — Amada/Trumpf can't detect hidden)
- Open DXF in viewer and visually confirm before handoff (see DXF Export Setup below)

---

## Check 3 — CU PROFI Hole Compliance (240-##### drawings ONLY)

**Applies only to copper bus bar drawings.** Skip this entire section for 250-/295-/100-/200-/245- prefixes.

### The full CU PROFI tool library

Cross-check every hole / slot / square dim against the full library in `tools\EngineeringDesignPackage\cu cut tool.pdf`. Any dim in the library is valid; ±0.001 is acceptable tolerance (e.g., `.531` vs `.530` is fine).

| Shape | Available dims |
|---|---|
| **ROUND** | .190, .203, .221, .250, .281, .312, .344, .375, .406, .413, .438, .500, .530, .563, .625, .656, .688, .875 |
| **OBLONG (W × L)** | .238×.375, .290×.460, .354×.500, .416×.560, .563×.750 |
| **SQUARE** | .275, .340, .400, .530, .650 |
| **RECTANGLE** | .315 × 2.362 |

### The #1 real CU PROFI error

**Hole callouts written as fractional bolt names instead of decimal CU PROFI dims.** The shop CNC runs on tool diameter, not bolt name. Always write decimals.

| Wrong (bolt name) | Right (CU PROFI decimal) |
|---|---|
| `1/4 THRU` | `Ø .281 THRU` (Tool #6) |
| `5/16 THRU` | `Ø .344 THRU` (Tool #8) |
| `3/8 THRU` | `Ø .406 THRU` (Tool #10) |
| `1/2 THRU` | `Ø .530 THRU` (Tool #14) |

**Root cause:** SolidWorks Hole Wizard's "Clearance Hole" preset writes the fastener-name token. Fix at source: use **Legacy Hole** and enter the decimal dim directly, OR override the callout text with `<MOD-DIAM>.344 THRU`.

### What does NOT count as a CU PROFI error

- A dim that's in the tool library but doesn't match the simplified "bolt clearance" mapping. E.g., a .375 hole on copper is valid (Tool #9). Don't assume it's a 3/8" bolt clearance unless the context confirms.
- Oblong sizes that match the library (e.g., `.416 × .560` = Tool #103). Don't require a tighter match to a "bolt clearance" table.
- Square hole dim-only callout (see Check 1).

---

## Check 4 — Title Block

**Required fields:** part#, revision, drawer, date, job#, weight, material, finish.

**Standard material strings:**
- Copper: `0.25 C1100 CU BAR` (not `0.25 Copper` — normalize)
- Galvanized 14ga: `0.0785 GALVANIZED CRS`
- Galvanized 12ga: `0.1046 GALVANIZED CRS`
- Aluminum 1/8": `0.125 5052-H32`

**Standard finish strings:** `BARE`, `TIN PLATED`, `ZINC PLATED`, `POWDER COAT RAL####`, etc. Flag `BAREBARE` (duplicated typo from property copy-paste).

**Title block is the current FoxFab block:** "Foxfab Power Solutions, 2579-188 ST., Surrey, BC V3Z 2A1". Flag if any sheet uses the **old block**: "FOXFAB METAL WORKS, 201 - 11517 Kingston St., Maple Ridge" — that's a legacy template and should be swapped.

**Duplicate titles across different part numbers** — flag. Especially for sequential part numbers (e.g., `-388` and `-389`): they're usually P1/P2 or LEFT/RIGHT variants and titles should disambiguate. Compare: overall dims, weight, hole count, bend count — if any differ, the parts are different and titles must differ.

---

## Revision Rules

| Change type | Rev behavior | Example |
|---|---|---|
| **Physical change** (hole dim, bend angle, hole added/removed) | Increment **letter** | rA → **rB** |
| **Drawing-only fix** (missing callout added, hole count corrected in annotation) | Append/increment **number** | rA → **rA1** → rA2 |
| **Cosmetic title-block cleanup** (blank material fill, BD/die removal on flat part, BAREBARE→BARE typo, 4*→4X, `0.25 Copper`→`0.25 C1100 CU BAR`) | **No rev bump** | edit in place |

**Revisions only happen after the part has been sent to the shop.** First issue is un-revved.

---

## DXF Export Setup (One-Time Configuration)

**Tools → Options → Document Properties → DXF/DWG Export:**

1. Enable **Custom Map SOLIDWORKS to DXF/DWG**.
2. Under **Entities to Export**, check **Geometry** and **Bend lines** (only — uncheck the rest unless needed).
3. Map **Bend lines – up direction** to a dedicated `BEND` layer.
4. Set the line style to a **Centerline / Section** type. **Do NOT use Hidden.**

**Per-export checklist:**

1. Right-click the part → **Flatten** to confirm the flat pattern is valid.
2. If Bend-Lines feature is greyed out in the feature tree → right-click → **Unsuppress**.
3. File → Save As → DXF → **Options** → confirm Custom Mapping is on.
4. After export, open the DXF in a viewer and confirm bend lines are present and on the BEND layer.

**Gotcha:** When a bend line crosses a cut feature, SolidWorks exports it as **separate segments**, not one continuous line. If your CAM software needs continuous lines, fix in QCAD (free) before sending.

**Template gotcha:** Bend lines must be visible in the **drawing template** itself. Use View Filter in the Heads-Up Toolbar to enable, then Save As → Drawing Template.

---

## Output File Checks (DXF / PDF / DWG)

Before handoff to the CNC team / shop, verify:

- **DXF** saved from **Part Editor** (not from drawing), flat pattern unfolded
- DXF export includes: **geometry, hidden edges, bend lines, forming tools**
- **PDF** and **DWG** saved using FoxFab templates
- Two top-level assembly drawings exist: **standard** AND **-LAY** version
- All assemblies (`100-` prefix) and flat/sheet metal parts (`240-`/`245-`/`250-` prefixes) are included
- Units in **inches** (not mm)
- Dimensions, bend lines, annotations all visible on the PDFs
- Reference example of a good output package: `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\J15689 Oncor Building C ATSDS\200 Mech\201 CAD`

---

## Quick Sanity Checklist

Run through before saving PDF/DXF:

### All drawings
- [ ] Title block: part#, revision, drawer, date, job#, weight
- [ ] Material field populated (standard notation — see Check 4)
- [ ] Finish field populated (no `BAREBARE` duplicates)
- [ ] Current FoxFab title block (not old Maple Ridge block)
- [ ] Units in inches
- [ ] All sketches fully defined (no blue geometry)
- [ ] No `(+)` over-defined mates in parent assembly
- [ ] Overall dims present (L × W × thickness)
- [ ] Material + finish in title block

### Bent parts (any with bends)
- [ ] Flat pattern view shows bend lines
- [ ] Every bend: direction (UP/DOWN), radius, BD value annotated
- [ ] **Ordinate-dimensioned at every bend** from datum 0
- [ ] BD field in title block matches the bend annotations
- [ ] DXF opened in viewer — bend lines visible on BEND layer

### Flat parts (no bends)
- [ ] BD, TOP DIE, BOTTOM DIE fields in title block are blank or `-`

### Copper parts (240-##### only)
- [ ] Every hole/slot dim is in the CU PROFI tool library (see Check 3)
- [ ] No bolt-name fractional callouts (`1/4 THRU`, `5/16 THRU`, etc.) — decimals only
- [ ] `0.25 C1100 CU BAR` material notation (not `0.25 Copper`)

### Assemblies (100-/200-)
- [ ] `-LAY` version exists for top-level
- [ ] BOM complete with item numbers, part numbers, descriptions, qty

---

## Common Mistakes to Watch For

1. **Bend lines missing from the DXF** — open the DXF and confirm before handoff.
2. **Bolt-name fractional callouts on copper** — always use the CU PROFI decimal.
3. **Dimensions only in the model, not on the drawing** — the shop works from the drawing PDF.
4. **Forgetting the `-LAY` version** of the top-level assembly.
5. **Units set to mm** instead of inches.
6. **Flat pattern view without bend lines** — unsuppress the Bend-Lines feature.
7. **Line style on BEND layer set to Hidden** — use Centerline/Section.
8. **Incrementing a letter revision for a drawing-only fix** — should be a number (rA → rA1).
9. **BD / TOP DIE / BOTTOM DIE populated on a flat part** — should be blank or `-`.
10. **Old FOXFAB METAL WORKS title block** on a sheet (check every sheet, not just sheet 1).
11. **Duplicate titles** across sequential part numbers without P1/P2 or LEFT/RIGHT disambiguation.
12. **`BAREBARE` typo in Finish field** — from template copy-paste.

---

## False Positives — Do NOT Flag These

Based on past review feedback:

- Square hole dimensioned only with a standalone dim (e.g., `.400`) — valid, shop reads the geometry.
- `R.438 THRU` for round holes — acceptable labeling even though ASME Y14.5 says use Ø.
- Hole dims in the CU PROFI library that don't match the "bolt clearance" subset — still valid.
- `.531` vs `.530` (within ±0.001) — acceptable tolerance.
- Unusual hole sizes (e.g., `.201`) on **non-copper** drawings — don't flag; shop interprets.
- Mixed units in TOP DIE / BOTTOM DIE fields (`1mm`, `16mm`) — these are tooling references, not part dims. Not critical.
- Minor title-block typos that don't affect fabrication (warnings only, no rev bump).

---

## When to Escalate

Stop and ask Vikram / a senior designer if:

- Non-standard material or thickness not in the Standard Materials table
- Custom tolerances tighter than ±0.010"
- Unusual bend angles where Bend Calculator output looks off
- Copper hole dims not in the CU PROFI tool library (after confirming it's truly not a typo)
- Customer PRF explicitly calls out something that conflicts with FoxFab standard practice

---

## Related Skills / References

- **`design-guide` skill** — full FoxFab design knowledge base (bend deduction, copper sizing, lugs, Kirk, BOM, etc.)
- **`tools\EngineeringDesignPackage\cu cut tool.pdf`** — full CU PROFI tool library (all shapes and sizes) — **always cross-check against this file before flagging a copper hole dim**
- **`tools\EngineeringDesignPackage\Minimum Bending Space.pdf`** — min bend space reference
- **Reference job with good output package**: `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\J15689 Oncor Building C ATSDS\200 Mech\201 CAD`
