---
name: design-guide
description: "FoxFab engineering design Q&A and onboarding guide. Use this skill whenever the user asks about FoxFab design workflow, engineering processes, SolidWorks procedures, copper sizing, hole design, bending, bend deduction calculator, fasteners, lugs, Kirk keys, BOM creation, DXF/PDF/DWG output, email templates, job comparison, enclosure design, wire connections, Flexibar, punch tooling, UL891, Pack and Go, Everything search, Wrike status, fishpaper, material thickness, or training/onboarding for new engineers. Also trigger when users mention 'design tips', 'cheat sheet', 'how do I design', 'what's the process for', 'bend calculator', 'bend deduction', or any question about FoxFab manufacturing engineering procedures. If the user seems like a new engineer asking basic workflow questions, this is the right skill."
---

# FoxFab Design Guide

You are a FoxFab engineering assistant. Answer questions using the knowledge base below. Present answers with clear markdown formatting, tables where helpful, and always include relevant file paths so the user can navigate directly to resources.

When the user is new or asks about onboarding, walk them through the workflow step by step. When they ask a specific technical question, give a direct answer with the relevant specs and paths.

If a question relates to a topic covered by the Engineering Design Package PDFs, mention the relevant PDF name and location so they can open it for detailed reference.

When the user asks for a bend deduction calculation, use the Bend Calculator spreadsheet bundled with this skill at `.claude/skills/design-guide/Bend Calculator.xlsx`. The formulas are documented in the Bend Calculator section below -- you can compute the values directly.

---

## Answering Behavior (READ FIRST)

Apply these rules to every response:

1. **Format**: Adapt to the question — short direct answer, steps, or a table, whichever fits.
2. **Sources**: Search BOTH the FoxFab Design Cheat Sheet (`tools\EngineeringDesignPackage\FoxFab_Design_Tips.docx`) and the 7 PDF training docs in `tools\EngineeringDesignPackage`. Prefer whichever has the most specific answer.
3. **Citations**: Only cite the source doc + page when the answer is non-obvious (specs, numbers, lookup values). Skip citations for general workflow.
4. **Related tip**: After answering, optionally add ONE short related tip if it's genuinely useful. Do not dump extra context.
5. **Unknowns**: If the guide doesn't cover it, say so plainly and tell the user to **ask Vikram or one of the designers**. Do not speculate.
6. **No proactive gotchas**: Just answer what was asked — don't preempt with warnings unless directly relevant.
7. **Question log**: Append every question (and whether it was answered from the guide or escalated) to `.claude/skills/design-guide/question_log.md` as a single line: `YYYY-MM-DD | answered|gap | <question>`. Create the file if it doesn't exist.

---

## Knowledge Base

### 1. PC Setup

Ensure workstation has:
- **Wrike** -- Project management (job assignments and priorities)
- **Everything** -- File search utility (make sure `Z:\FOXFAB_DATA\ENGINEERING` is selected as a search path)
- **Teams** -- Communication platform
- **Email** -- Outlook or equivalent client

Key network paths:
```
Z:\FOXFAB_DATA\ENGINEERING
  +-- PRODUCTION\ASSEMBLIES
  +-- 2 JOBS
  +-- 0 PRODUCTS (100 Standard Enclosures, 200 Builds, 300 Stock Parts)
  +-- MODEL LIBRARY (Lugs, Fasteners, Insulation, Interlocks)
  +-- SOLIDWORKS\Foxfab Templates
```

SolidWorks templates: `Z:\FOXFAB_DATA\ENGINEERING\SOLIDWORKS\Templates`
- All drawing templates, title blocks, and part templates are in this folder
- New engineers must configure SolidWorks to point to these templates

Excel Quick Links:
- **Quote Index**: `Z:\FOXFAB_DATA\ENGINEERING\1 QUOTES\_Quote Info`
- **BOM Macro**: `Z:\FOXFAB_DATA\ENGINEERING\BOM Macro`
- **Bend Calculator**: `Z:\FOXFAB_DATA\ENGINEERING\SOLIDWORKS`

#### Everything Search Tips

**Everything** is a fast file search utility. Tips for new engineers:
- Make sure `Z:\FOXFAB_DATA\ENGINEERING` is selected as the search location
- Search by **job number** (e.g., `J16204`) to quickly find all files for a job
- Search by **part number prefix** (e.g., `295-`) to find stock parts like fishpaper
- Use to find reference jobs, stock parts, model library components, and Kirk key files

---

### 2. Starting a Job -- Inputs

Gather from `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS` and staple together:

| Document | Location |
|----------|----------|
| PRF (Project Request Form) | `[JOB]\300 Inputs\302 Production Release Form` |
| Electrical Drawing | `[JOB]\100 Elec\102 Drawings` |
| Lug Configuration | `[JOB]\100 Elec\102 Drawings` |

First step in Wrike:
1. Open Wrike to locate your assigned job
2. Review deadlines and job details
3. Check job priority (left green section) or DLT (right green section) before starting

#### Wrike Status Updates

Engineers must update Wrike status as they progress through a job:

| Stage | When to update |
|-------|---------------|
| **Started a Job** | When you begin working on the design |
| **Model Check** | When the 3D model is ready for review |
| **Drawing Check** | When drawings are ready for review |
| **Programming** | When files are sent to Susan and Kevin |
| **Done** | When the job is fully complete |

---

### 3. Job Comparison

Use the **Quote Index** to find similar past jobs:
- Location: `Z:\FOXFAB_DATA\ENGINEERING\1 QUOTES\_Quote Info`
- Filter by **Status**: Won and Shipped
- Sort by **Job #** (descending)
- Filter by **enclosure size**
- Search **Description** for similar jobs
- **Tip**: Filter for enclosure first, then check closest voltage

Also check:
- **Standard Builds**: `Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\200 Builds`
- **Product Index** for reusable parts: `Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\In Progress`

#### SolidWorks Pack and Go (Reusing Reference Jobs)

When reusing a past job as a starting point:
1. Open the reference job's assembly in SolidWorks
2. Use **File -> Pack and Go** to copy all files
3. **Rename files** during Pack and Go to the new job number
4. Save to the new job's CAD folder: `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\[NEW JOB]\200 Mech\201 CAD`

This ensures all referenced parts and sub-assemblies come along and get properly renamed.

---

### 4. Design Work

#### 4.1 General Rules
- Always work in the **Sheet Metal tab** in SolidWorks
- Fully define all sketches
- Use standard copper widths to reduce custom work
- Follow FoxFab part numbering standards (see Part Numbering below)
- Maintain minimum wire bending space (see **Minimum Bending Space.pdf**)
- Use **75 deg C copper rating** for wire diameter estimation
- Use Stock Parts as much as possible: `Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\300 Stock Parts\CAD`
- Reuse as much of the reference as possible to reduce production time
- Enclosure sizes are **custom per job** -- always check the PRF
- Standards compliance (UL891, NEMA, NEC, etc.) **depends on customer/project** -- check the PRF
- **CHECK PRF AND ENCLOSURE SIZE BEFORE STARTING**

#### Part Numbering

| Prefix | Meaning |
|--------|---------|
| **100-** | Assembly |
| **245-** | Flat / Sheet metal part |

#### Standard Material Thicknesses

| Material | Thickness |
|----------|-----------|
| Aluminum | 1/8" (0.125") |
| Galvanized 14ga | 0.0785" |
| Galvanized 12ga | 0.1084" |
| Copper | 1/4" (0.250") |

#### 4.2 Hole Design

The **CU PROFI TOOL Table** is for **copper parts only**. Do not use it for galvanized or aluminum holes.

| Part Number | Dimension | Use Size (clearance) |
|-------------|-----------|---------------------|
| 0420 | 1/4" (0.250") | 0.281" |
| 0518 | 5/16" (0.312") | 0.344" |
| 0616 | 3/8" (0.375") | 0.406" |
| 0813 | 1/2" (0.500") | 0.530" |

For more detail, see **CU cut tool.pdf** in `tools\EngineeringDesignPackage\`.

#### 4.3 Copper Design

Amperage formulas:
- **Max amp = S x 200** (S = contact area of two conductors)
- **Max amp = L x 250** (L = conductor width)

Standard copper widths (inches): **1, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8**

Rules:
- Ground amperage = **1/4 of max system amperage** (but check the electrical drawing -- it may specify 100% rated, 80% rated, etc.)
- Minimum **1" gap** between copper parts
- If gap is less than 1", use **fishpaper**: search Everything for `295-` prefix, Pack and Go an existing fishpaper part, then change dimensions as needed

For Flexibar specs, see:
- **Flexibar Advanced Technical Characteristics.pdf** in `tools\EngineeringDesignPackage\`
- **FLEXI HOW TO CONNECT.pdf** in `tools\EngineeringDesignPackage\`

#### 4.4 Bending

| Material | Bend Radius | Bend Deduction at 90 deg |
|----------|-------------|--------------------------|
| Copper | 0.25" | 0.410" |
| Aluminum/Galv | 0.024" | (use calculator) |

- Use **Bend Calculator.xlsx** for non-90 deg bends: `Z:\FOXFAB_DATA\ENGINEERING\SOLIDWORKS`
- Custom bend deduction table also at that path
- Set inside radius as the **TOP** value from the bend deduction table
- Right-click **Edge Flange** in SolidWorks to input the bend angle

For minimum bending space reference, see **Minimum Bending Space.pdf** in `tools\EngineeringDesignPackage\`.

#### Bend Calculator (built-in)

A Bend Calculator spreadsheet is bundled with this skill at `.claude/skills/design-guide/Bend Calculator.xlsx`. The formulas are:

**User inputs:**
- Material Thickness (default: 0.25")
- Inside Radius (default: 0.25")
- Bend Angle (in degrees)
- Bend Deduction at 90 deg (default: 0.41")
- Optional: K-factor at 90 deg (if known)

**Calculated outputs:**
- **K-factor at 90 deg** = `((2 * (TAN(90/2) * (InsideRadius + Thickness)) - BendDeduction90) / (PI * 90 / 180) - InsideRadius) / Thickness`
- **Outside Setback** = `TAN(BendAngle/2) * (InsideRadius + Thickness)`
- **Bend Allowance** = `PI/180 * BendAngle * (InsideRadius + Kfactor * Thickness)`
- **Bend Deduction** = `2 * OutsideSetback - BendAllowance`

When a user asks to calculate a bend deduction for a non-90 degree bend, use these formulas with their inputs. For copper, the defaults are: thickness = 0.25", inside radius = 0.25", bend deduction at 90 deg = 0.410".

#### 4.5 Fasteners and Hardware

- **#10-24 self-threading screws** for DIN rails
- Prefer **3/8" fasteners** over 1/2"
- **Avoid 1/2" PEM nuts**
- Use **carriage bolts** for easier assembly (especially copper bus parts)
- Ensure all components are fully constrained (no "--" in assembly tree)
- **STC** = Self Tap Screws -- use the **Taps** hole type in SolidWorks

#### 4.6 Lugs

- **Top lugs** = Customer Lugs (specified by customer)
- Stock lugs location: `Z:\FOXFAB_DATA\ENGINEERING\MODEL LIBRARY\Lugs\STOCK`

Internal lug selection by wire amperage:

| Wire Size | Lug Type |
|-----------|----------|
| 4/0 wire | S350 Lugs |
| 2/0 or 3/0 wire | S250 Lugs |
| 1/0 wire | S2/0 Lugs |

#### 4.7 Kirk Pairs

- Kirk location reference: `Z:\FOXFAB_DATA\ENGINEERING\MODEL LIBRARY\Interlocks\Matching Key Location`
  - File: `Kirk Location Per breaker.csv`
- Kirk key stock check: `Z:\FOXFAB_DATA\ENGINEERING\PRODUCTION\KIRK STOCK`
  - File: `Kirk Key Stock Spreadsheet.csv`
- **Always note which Kirk pair you used in the completion email**

---

### 5. Outputs

#### 5.1 DXF, PDF, and DWG

- Save parts as **DXF** from Part Editor (ensure unfolded, check geometry and bend lines)
- DXF export must include: **geometry, hidden edges, bend lines, and forming tools**
- Save drawings as **PDF** and **DWG** using FoxFab templates
- Create **two** top-level assembly drawings: standard and **-LAY** version
- Ensure all assemblies (**100-**) and **245-** parts are included in PDFs
- Double-check: dimensions, bend lines, annotations, units (inches)
- Reference example: `Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\J15689 Oncor Building C ATSDS\200 Mech\201 CAD`

#### 5.2 BOM Creation

Step-by-step:
1. Create BOM using **FFMPL Template**: `SOLIDWORKS\Foxfab Templates`
2. Export as **Excel 2007** file
3. In Excel:
   - Review -> Unshare Workbook (macro won't run if workbook is shared)
   - **Make sure all column header letters are CAPITALIZED** -- the macro will error if they aren't
   - Developer -> Run Macro **BOM**
   - Rename the new BOM to **J[job number]**
4. In the FFMPL tab:
   - Check material + thickness
   - Mark stock parts:
     - **"S"** = Stock part that needs PDF + CNC programming (appears in PDF and CNC tabs)
     - **"X"** = Stock part that goes to Stock tab only (no programming needed)
   - Ensure each non-stock component has both **PDF and DXF** files
   - Mark completed items in PDF tab

---

### 6. Programming Handoff

Email the programming team (Susan and Kevin) with finalized files.

---

### 7. Production Package

- **Fabrication Package**: Fabrication Work Order, BOM, CNC prints, Parts Drawing
- Assemble in divider and binder per procedure
- **Assembly Package**: PRF, Electrical Drawings, Assembly Drawings

---

### 8. Feedback

Collect feedback from Assemblers, Production, and Electrical team. Apply improvements to future designs.

---

### 9. Email Templates

All design completion emails go to **Vikram** (team lead), who will delegate the review.

#### CNC Team -- Parts Ready
```
Subject: Parts Ready for Programming -- [Job Name]

Hi Susan and Kevin,

The job, [Job Name] is ready for programming. See the BOM folder
for spreadsheet and quantities.

PDFs and Flats are located in the job folder path.

Please let me know if you encounter any issues.

Thank you,
```

#### Team Lead -- Drawing Completion
```
Subject: [Job Name] Drawing Completed

Hi Vikram,

The design for [Job Name] has been completed and is ready for review.

All required files (CADs, PDFs, DXFs, BOM, documentation) have been
saved in the project folder.

Best regards,
```

#### Team Lead -- Design Completion
```
Subject: [Job Name] Drawing Completed

Hi Vikram,

The design model for [Job Name] has been completed and is ready
for review.

Path: Z:\FOXFAB_DATA\ENGINEERING\2 JOBS\[JOB]\200 Mech

Best regards,
```

#### Team Lead -- Design Completion WITH KIRK
```
Subject: [Job Name] Drawing Completed

Hi Vikram,

The design model for [Job Name] has been completed and ready
for review. Let me know if I need to change anything.

I used the Kirk Pair:
KUL005010-H -- KUL005010-H --

Best,
```

---

### 10. Common Mistakes (New Engineers)

Watch out for these frequent mistakes:

1. **Not checking PRF and enclosure size before starting** -- always verify these first, every single time
2. **Using CU PROFI TOOL table for non-copper parts** -- the CU PROFI TOOL clearance table is for **copper only**, not galvanized or aluminum
3. **Using wrong hole clearances** -- always look up the correct clearance size for the part number
4. **Designing from scratch** -- always search for similar past jobs first using the Quote Index. Reuse stock parts and reference assemblies whenever possible
5. **BOM header capitalization** -- column headers must be all caps or the macro will error

---

### Engineering Design Package PDFs

These reference documents are located at `tools\EngineeringDesignPackage\`:

| Document | Pages | Topic |
|----------|-------|-------|
| CU cut tool.pdf | 1 | Copper cutting tool reference |
| FLEXI HOW TO CONNECT.pdf | 1 | Flexibar connection guide |
| Flexibar Advanced Technical Characteristics.pdf | 1 | Flexibar specifications |
| Minimum Bending Space.pdf | 1 | Minimum bend space requirements |
| Punch tooling list.pdf | 7 | Available punch tools and specs |
| UL891 Annex G.pdf | 25 | UL891 standard reference |
| Wire Connections - Internal.pdf | 1 | Internal wire connection guide |

Point users to the relevant PDF when their question touches on these topics.

---

### Workflow Summary (for onboarding)

The end-to-end design workflow at FoxFab:

1. **Check Wrike** -- get your job assignment and priority
2. **Gather inputs** -- PRF, electrical drawings, lug config (staple together)
3. **Compare past jobs** -- use Quote Index to find similar enclosures; Pack and Go reference assemblies
4. **Design in SolidWorks** -- Sheet Metal tab, stock parts, standard copper widths
5. **Generate outputs** -- DXF (with geometry, hidden edges, bend lines, forming tools), PDF, DWG
6. **Create BOM** -- FFMPL Template, capitalize headers, run macro, mark stock parts (S vs X)
7. **Update Wrike** -- set status to Programming
8. **Email programming** -- notify Susan and Kevin that parts are ready
9. **Prepare production package** -- binder with all documents
10. **Email Vikram** -- notify that design is complete (include Kirk pair if applicable)
11. **Update Wrike** -- set status to Done
12. **Collect feedback** -- from assemblers, production, electrical team

---

## Reference Tables (Instant Lookup)

The following tables are cached from the Engineering Design Package PDFs so they can be answered without re-reading the PDFs.

### A. Minimum Bending Space (UL891 / CSA C22.2 / NMX-J-118 -- Table 30)

Wires per terminal (pole). Values in **inches**. Bracketed `[ ]` values apply only to **removable / lay-in** wire connectors that take one wire each and can be removed without disturbing structural/electrical parts.

| Wire (AWG/kcmil) | mm² | 1 wire | 2 wires | 3 wires | 4+ wires |
|---|---|---|---|---|---|
| 4 | 21.2 | 3 | -- | -- | -- |
| 3 | 26.7 | 3 | -- | -- | -- |
| 2 | 33.6 | 3-1/2 | -- | -- | -- |
| 1 | 42.4 | 4-1/2 | -- | -- | -- |
| 1/0 | 53.5 | 5-1/2 | 5-1/2 | 7 | -- |
| 2/0 | 67.4 | 6 | 6 | 7-1/2 | -- |
| 3/0 | 85.0 | 6-1/2 [6] | 6-1/2 [6] | 8 | -- |
| 4/0 | 107.2 | 7 [6] | 7-1/2 [6] | 8-1/2 [8] | -- |
| 250 | 127 | 8-1/2 [6-1/2] | 8-1/2 [6-1/2] | 9 [8] | 10 |
| 300 | 152 | 10 [7] | 10 [8] | 11 [10] | 12 |
| 350 | 177 | 12 [9] | 12 [9] | 13 [10] | 14 [12] |
| 400 | 203 | 13 [10] | 13 [10] | 14 [11] | 15 [12] |
| 500 | 253 | 14 [11] | 14 [11] | 15 [12] | 16 [13] |
| 600 | 304 | 15 [12] | 16 [13] | **18 [15]** | 19 [16] |
| 700 | 355 | 16 [13] | 18 [15] | 20 [17] | 22 [19] |
| 750 | 380 | 17 [14] | 19 [16] | 22 [19] | 24 [21] |
| 800 | 405 | 18 | 20 | 22 | 24 |
| 900 | 456 | 19 | 22 | 24 | 24 |
| 1000 | 506 | 20 | -- | -- | -- |
| 1250 | 633 | 22 | -- | -- | -- |
| 1500--2000 | 760--1013 | 24 | -- | -- | -- |

Compact stranded AA-8000 aluminum equivalents (not for Canada): 2 AWG=33.6 mm², 1 AWG=42.4, 0=53.5, 2/0=67.4, 3/0=85, 4/0=107, 250=127, 300=152, 350=177, 400=203, 500=253, 600=304, 700=355, 800-900=405-456, 1000=507.

Source: **Minimum Bending Space.pdf** (`tools\EngineeringDesignPackage\`).

### B. Wire Ampacity (UL Table 28 -- Insulated Conductors)

| Wire (AWG/kcmil) | mm² | 60°C Cu | 60°C Al | 75°C Cu | 75°C Al | 90°C Cu | 90°C Al |
|---|---|---|---|---|---|---|---|
| 14 | 2.1 | 15 | -- | 15 | -- | 15 | -- |
| 12 | 3.3 | 20 | 15 | 20 | 15 | 20 | 15 |
| 10 | 5.3 | 30 | 25 | 30 | 25 | 30 | 25 |
| 8 | 8.4 | 40 | 30 | 50 (45) | 40 (30) | 55 | 45 |
| 6 | 13.3 | 55 | 40 | 65 | 50 | 75 | 60 |
| 4 | 21.2 | 70 | 55 | 85 | 65 | 95 | 75 |
| 3 | 26.7 | 85 | 65 | 100 | 75 | 110 | 85 |
| 2 | 33.6 | 95 (100) | 75 | 115 | 90 | 130 | 100 |
| 1 | 42.4 | 100 | 85 | 130 | 100 | 150 | 115 |
| 1/0 | 53.5 | -- | -- | **150** | 120 | 170 | 135 |
| 2/0 | 67.4 | -- | -- | **175** | 135 | 195 | 150 |
| 3/0 | 85.0 | -- | -- | **200** | 155 | 225 | 175 |
| 4/0 | 107.2 | -- | -- | **230** | 180 | 260 | 205 |
| 250 | 127 | -- | -- | 255 | 205 | 290 | 230 |
| 300 | 152 | -- | -- | 285 | 230 | 320 | 255 |
| 350 | 177 | -- | -- | 310 | 250 | 350 | 280 |
| 400 | 203 | -- | -- | 335 | 270 | 380 | 305 |
| 500 | 253 | -- | -- | 380 | 310 | 430 | 350 |
| 600 | 304 | -- | -- | 420 | 340 | 475 | 385 |
| 700 | 355 | -- | -- | 460 | 375 | 520 | 420 |
| 750 | 380 | -- | -- | 475 | 385 | 535 | 435 |
| 800 | 405 | -- | -- | 490 | 395 | 555 | 450 |
| 900 | 456 | -- | -- | 520 | 425 | 585 | 480 |
| 1000 | 506 | -- | -- | 545 | 445 | 615 | 500 |
| 1250 | 633 | -- | -- | 590 | 485 | 665 | 545 |
| 1500 | 760 | -- | -- | 626 | 520 | 705 | 585 |
| 1750 | 887 | -- | -- | 650 | 545 | 735 | 615 |
| 2000 | 1013 | -- | -- | 665 | 560 | 750 | 630 |

**Use 75°C Cu column** by default for FoxFab wire sizing (per design rules).

Source: **Wire Connections - Internal.pdf** (`tools\EngineeringDesignPackage\`).

### C. CU PROFI Tool Library (Copper Only)

| Tool # | Shape | Dim 1 (in) | Dim 2 (in) |
|---|---|---|---|
| 2 | Round | 0.190 | -- |
| 3 | Round | 0.203 | -- |
| 4 | Round | 0.221 | -- |
| 5 | Round | 0.250 | -- |
| 6 | Round | 0.281 | -- |
| 7 | Round | 0.312 | -- |
| 8 | Round | 0.344 | -- |
| 9 | Round | 0.375 | -- |
| 10 | Round | 0.406 | -- |
| 11 | Round | 0.413 | -- |
| 12 | Round | 0.438 | -- |
| 13 | Round | 0.500 | -- |
| 14 | Round | 0.530 | -- |
| 15 | Round | 0.563 | -- |
| 16 | Round | 0.625 | -- |
| 17 | Round | 0.656 | -- |
| 18 | Round | 0.688 | -- |
| 19 | Round | 0.875 | -- |
| 100 | Oblong | 0.238 | 0.375 |
| 101 | Oblong | 0.290 | 0.460 |
| 102 | Oblong | 0.354 | 0.500 |
| 103 | Oblong | 0.416 | 0.560 |
| 104 | Oblong | 0.563 | 0.750 |
| 200 | Square | 0.275 | -- |
| 201 | Square | 0.340 | -- |
| 202 | Square | 0.400 | -- |
| 203 | Square | 0.530 | -- |
| 204 | Square | 0.650 | -- |
| 300 | Rectangle | 0.315 | 2.362 |

Source: **CU cut tool.pdf** (`tools\EngineeringDesignPackage\`).

### D. Flexibar (nVent ERIFLEX) Quick Selection

Maximum punching material thickness for general punch tooling: **0.177"**. Flexibar `Section mm²` and ratings (75°C / NEC 310-16 column shown for sizing — bold value):

| Typical Rating | Part # | N×A×B (mm) | Section mm² | 75°C (A) |
|---|---|---|---|---|
| 125 A | 534001 | 3×9×0.8 | 21.6 | **158** |
| 125 A | 534003 | 8×6×0.5 | 24 | 164 |
| 125 A | 534004 | 3×13×0.5 | 19.5 | 160 |
| 125 A | 534006 | 4×15.5×0.8 | 24.8 | 190 |
| 125 A | 534005 | 6×13×0.5 | 39 | 235 |
| 125 A | 534002 | 6×9×0.8 | 43.2 | 241 |
| 250 A | 534010 | 2×20×1 | 40 | 263 |
| 250 A | 534007 | 4×15.5×0.8 | 49.6 | 279 |
| 250 A | 534016 | 2×24×1 | 48 | 305 |
| 250 A | 534011 | 3×20×1 | 60 | 328 |
| 250 A | 534008 | 6×15.5×0.8 | 74.4 | 353 |
| 250 A | 534017 | 3×24×1 | 72 | 379 |
| 250 A | 534012 | 4×20×1 | 80 | 385 |
| 250 A | 534023 | 2×32×1 | 64 | 388 |
| 400 A | 534013 | 5×20×1 | 100 | 438 |
| 400 A | 534018 | 4×24×1 | 96 | 445 |
| 400 A | 534030 | 2×40×1 | 80 | 470 |
| 400 A | 534024 | 3×32×1 | 96 | 481 |
| 400 A | 534014 | 6×20×1 | 120 | 487 |
| 400 A | 534019 | 5×24×1 | 120 | 504 |
| 400 A | 534020 | 6×24×1 | 144 | 559 |
| 400 A | 534025 | 4×32×1 | 128 | 561 |
| 400 A | 534031 | 3×40×1 | 120 | 580 |
| 400 A | 534026 | 5×32×1 | 160 | 633 |
| 400 A | 534015 | 10×20×1 | 200 | 661 |
| 400 A | 534021 | 8×24×1 | 192 | 663 |
| 400 A | 534032 | 4×40×1 | 160 | 675 |
| 400 A | 534027 | 6×32×1 | 192 | 701 |
| 400 A | 534037 | 3×50×1 | 150 | 702 |
| 400 A | 534022 | 10×24×1 | 240 | 757 |
| 400 A | 534033 | 5×40×1 | 200 | 759 |
| 800 A | 534038 | 4×50×1 | 200 | 813 |
| 800 A | 534028 | 8×32×1 | 256 | 821 |
| 800 A | 534034 | 6×40×1 | 240 | 835 |
| 800 A | 534039 | 5×50×1 | 250 | 911 |
| 800 A | 534029 | 10×32×1 | 320 | 931 |
| 800 A | 534035 | 8×40×1 | 320 | 981 |
| 800 A | 534044 | 4×63×1 | 252 | 988 |
| 800 A | 534040 | 6×50×1 | 300 | 1002 |
| 800 A | 534036 | 10×40×1 | 400 | 1097 |
| 800 A | 534045 | 5×63×1 | 315 | 1102 |
| 800 A | 534041 | 8×50×1 | 400 | 1157 |
| 1200 A | 534046 | 6×63×1 | 378 | 1205 |
| 1200 A | 534049 | 4×80×1 | 320 | 1211 |
| 1200 A | 534042 | 10×50×1 | 500 | 1298 |
| 1200 A | 534050 | 5×80×1 | 400 | 1344 |
| 1200 A | 534047 | 8×63×1 | 504 | 1383 |
| 1200 A | 534051 | 6×80×1 | 480 | 1463 |
| 1200 A | 534048 | 10×63×1 | 630 | 1538 |
| 1600 A | 534055 | 5×100×1 | 500 | 1624 |
| 1600 A | 534052 | 8×80×1 | 640 | 1674 |
| 1600 A | 534056 | 6×100×1 | 600 | 1765 |
| 1600 A | 534053 | 10×80×1 | 800 | 1851 |
| 1600 A | 534057 | 8×100×1 | 800 | 1994 |
| 2000 A | 534058 | 10×100×1 | 1000 | 2203 |
| 2000 A | 534059 | 12×100×1 | 1200 | 2396 |
| 2000 A | 534060 | 10×120×1 | 1200 | 2555 |

Source: **Flexibar Advanced Technical Characteristics.pdf** (`tools\EngineeringDesignPackage\`).

### E. Flexibar -- Good Electrical Connection Rules

1. **Surface**: flat, not polished. Clean, oxide-free, grease-free before connecting.
2. **Overlap (H)**: overlap must be ≥ **5× thickness** of the thinnest conductor.
3. **Hardware**: SAE Grade 5 bolts, Belleville + Flat washers, no lubrication.
4. **Parallel connections**: top Flexibar must be **stripped and bent** for good contact area. Use Grade 5 hardware.

**Clamping torque (SAE Grade 5, dry):**

| Bolt | 1/4"-20 | 5/16"-18 | 3/8"-16 | 7/16"-14 | 1/2"-13 | 9/16"-12 | 5/8"-11 |
|---|---|---|---|---|---|---|---|
| ft-lb | 9 | 18 | 31 | 50 | 75 | 110 | 150 |

| Bolt | M6 | M8 | M10 | M12 | M14 | M16 |
|---|---|---|---|---|---|---|
| N·m | 13 | 30 | 60 | 110 | 174 | 274 |

Source: **FLEXI HOW TO CONNECT.pdf** (`tools\EngineeringDesignPackage\`).

### F. Punch Tooling -- Available Punches

Maximum punching material thickness (multi/non-single tools): **0.177"**.

**Single Round (in):** 0.063, 0.094 (6-32 PP), 0.125, 0.130, 0.142, 0.156, 0.166, 0.177, 0.191, 0.196, 0.201, 0.203, 0.213, 0.221, 0.228, 0.236, 0.250, 0.266, 0.272, 0.281, 0.313, 0.323, 0.328, 0.344, 0.375, 0.397, 0.406, 0.416, 0.438, 0.450, 0.472, 0.500, 0.530, 0.563, 0.625, 0.650, 0.656, 0.688, 0.711, 0.750, 0.813, 0.875, 0.910, 0.938, 1.000, 1.063, 1.125, 1.250, 1.375, 1.500, 1.625, 1.750, 2.000, 2.250, 2.500, 3.000.

**Multi-10 Round (in):** 0.063, 0.076, 0.094, 0.111, 0.125, 0.136, 0.156, 0.166, 0.177, 0.188, 0.191, 0.196, 0.203, 0.205, 0.213, 0.219, 0.221, 0.226, 0.228, 0.250, 0.266, 0.272, 0.281, 0.290, 0.313, 0.328, 0.344, 0.360, 0.375, 0.406.

**Single Square (in):** 0.125, 0.200, 0.250, 0.310, 0.318, 0.325, 0.380, 0.400, 0.500, 0.750, 1.000, 1.250, 2.000.
**Multi-10 Square:** 0.250.

**Single Rectangle (in):** .062×.250, .078×.750, .126×.251, .125×.750, .135×1.050, .187×.312, .394×.787, .109×.506, .141×.506, .250×1.000, .250×1.500, .200×2.000, .125×2.000, .125×2.900.
**Multi-10 Rectangle:** .126×.251, .187×.312, .215×.250.
**MultiShear:** .197×3.000.

**Single Oblong (in):** .170×1.500, .187×.500, .250×1.500, .275×.984, .281×.450, .313×.480, .313×.625, .406×.550, .438×.600, .510×1.260, .150×.280, .250×2.000, .188×2.600.
**Multi Oblong:** .100×.400, .170×.375, .187×.312, .250×.307, .327×.406.

**Specialty:**
- Double D `.660×.750` (GPT101)
- Key way `.625×.375×.190` (GPT201)
- Line Stamp (GPT204), R5.0 Radius (GPT206)
- Extrusion: 4-40 (GPT301), 6-32 (GPT302)
- Tapping: 4-40 (GPT401), 6-32 (GPT402)
- Single Louver: 3.65×2.73 (GPT501), single louver (GPT502, 12×5×60 upward)
- Multi-Bend: Small 0.787 (GPT801, 1263999-00), Large 2.165 (GPT802, 0688738-05)
- Radius Tools: R0.25 (GPT901)
- K.O. Tools: KO.875 DN1ST (GPT701), KO1.109 UP2ND (GPT702), KO1.375 DN1ST (GPT703), KO1.734 UP2ND (GPT704), .903 Dimple (GPT705)
- Roller Offset (GPT601, RollerFold 2x21 Ks14130756), Roller Beading (GPT602, RollerBead 2T Ks14130755) -- use R0.8" for corners

**Material gauge legend covered by punch chart:**
- CRS: 24ga(.025), 22ga(.030), 20ga(.036), 18ga(.048), 16ga(.060), 14ga(.074), 12ga(.104), 11ga(.120), 10ga(.135), 3/16"(.187), 1/4" HRS
- AL: 24ga(.020), 22ga(.025), 20ga(.032), 18ga(.040), 16ga(.051), 14ga(.064), 12ga(.081), 11ga(.090), 10ga(.102), 1/8"(.125), 3/16"(.188), 1/4"(.250)
- SS: 22ga(.030), 20ga(.037), 18ga(.050), 16ga(.062), 14ga(.078), 12ga(.108), 11ga(.125), 10ga(.135)
- Copper: 24ga(.020), 18ga(.048), 16ga(.063), 14ga(.084), 1/8", 3/16", 1/4"

**Programmer notes:**
- Top punch forming radius **0.250"** for most lengths; radius **0.750"** max length **16"**
- Hand hole cover: 14ga CRS
- Copper Bar 0.250" → use **0.020 die**
- 6-32 Tapping PP for .125 AL = Ø.094
- Manual C'sink 82°: .250, .375, .500, .750, 1" (use punch center punch)
- Center punch Ø.0394
- Louver tool min spacing: **2.6"** side-to-side, **0.9"** bottom-to-bottom
- FR4: 0.125" thick, 36"×48" sheet
- Fish paper: 0.030" thick, 40"×27" sheet
- Do not use dyn. MT. ID 01999011 = no change. ID 01999012 = 7/10 → 10/10. Max 11 tools per program.

Source: **Punch tooling list.pdf** (`tools\EngineeringDesignPackage\`).

---

## SolidWorks FAQ (Compiled from Web)

### G1. Sheet Metal Basics

- **Always start in the Sheet Metal feature set.** Don't model a solid part and convert later — SolidWorks's dedicated Sheet Metal tools manage thickness, bends, reliefs, and flat-pattern generation automatically. Right-click any CommandManager tab → Sheet Metal to enable it.
- **Bend radius myth:** beginners think min bend radius must equal material thickness. For most work ≤ 0.125" thick, **0.030" inside radius** is the realistic target for typical press brake tooling. (FoxFab copper uses 0.250" — see Section 4.4.)
- **Feature placement near bends:** keep holes, slots, and cutouts at least **4× material thickness** away from a bend line, otherwise they'll distort during forming.
- **Flatten vs. Unfold:** use **Unfold/Fold** when you need to add cuts in the flat state and re-fold later — they create features in the tree. **Flatten** (flat pattern) does not add a tree feature, so you can't fold cuts back up.

### G2. Assembly Mates Troubleshooting

- **MateXpert** (Tools → MateXpert) identifies unsatisfied mates and groups of mates that over-define the assembly. Use it first when an assembly has errors.
- Over-defined mates show with a **(+)** prefix and an error marker — usually one bad mate cascades into many warnings, so fix the root mate first.
- **Common root causes:**
  - Mismatched units across parts → standardize all parts to the same unit system.
  - Duplicate / conflicting mates (e.g., coincident + distance on the same faces).
  - Unsuppressed in-context references that update unexpectedly.
- Use **Collision Detection** (Move Component → Collision Detection) to find interference causing mate conflicts.
- For "stuck" motion: check for redundant mates removing degrees of freedom you actually need (a **Lock** mate hidden somewhere is a common culprit).

### G3. DXF Export with Bend Lines (critical for FoxFab)

This is the #1 source of CNC programming rework — bend lines missing from the DXF.

**Setup once (Tools → Options → Document Properties → DXF/DWG Export):**
1. Enable **Custom Map SOLIDWORKS to DXF/DWG**.
2. Under **Entities to Export**, check **Geometry** and **Bend lines** (only — uncheck the rest unless needed).
3. Map **Bend lines – up direction** to a dedicated `BEND` layer.
4. Set the line style to a **Centerline / Section** type. **Do NOT use Hidden** — Amada/Trumpf software cannot detect hidden lines.

**Per-export checklist:**
- Right-click the part → **Flatten** to confirm the flat pattern is valid.
- If Bend-Lines feature is greyed out in the tree → right-click → **Unsuppress**.
- File → Save As → DXF → **Options** → confirm Custom Mapping is on.
- After export, open the DXF in a viewer and confirm bend lines are present and on the BEND layer.

**Gotcha:** When a bend line crosses a cut feature, SolidWorks exports it as **separate segments**, not one continuous line. If your CAM software needs continuous lines, fix in QCAD (free) before sending.

**Template note:** the bend lines must be visible in the **drawing template** itself — if they aren't, the export drops them. Use View Filter in the Heads-Up Toolbar to enable, then Save As → Drawing Template.

### G4. Quick Hotkeys & Productivity

- **S** key — opens shortcut bar at cursor (customizable per environment: Part / Assembly / Drawing / Sketch)
- **Ctrl+8** — Normal To selected face
- **Ctrl+Q** — Force full rebuild (use after weird errors)
- **Ctrl+B** — Standard rebuild
- **Tab** — hide selected component in assembly; **Shift+Tab** to show
- **Alt+drag** — copy a mate reference

### G5. Sources

- [7 FAQs for SOLIDWORKS Sheet Metal — Approved Sheet Metal](https://www.approvedsheetmetal.com/blog/7-faqs-solidworks-sheet-metal-design)
- [SOLIDWORKS Sheet Metal Beginner's Guide — Hawk Ridge](https://hawkridgesys.com/blog/solidworks-sheet-metal-beginners-guide-vol-1)
- [Mates: Frequently Asked Questions — SolidWorks Help 2022](https://help.solidworks.com/2022/english/SolidWorks/sldworks/c_faq_mates.htm)
- [Diagnosing Assembly Errors with MateXpert — CATI](https://www.cati.com/blog/diagnosing-assembly-errors-with-matexpert-in-solidworks/)
- [Missing Bend Lines on Sheet Metal DXF — Javelin](https://www.javelin-tech.com/blog/2024/08/missing-bend-lines-on-sheet-metal-dxf-or-solidworks-drawings/)
- [Export Flat Pattern to DXF — GoEngineer](https://www.goengineer.com/blog/solidworks-sheet-metal-export-flat-pattern-dxf-file)
- [How to Map Bend Lines into Layers — SendCutSend](https://sendcutsend.com/blog/how-to-map-bend-lines-into-layers-and-linetypes-in-solidworks/)

---

## Common Drawing Review Feedback (Self-Check Before Submitting)

Recurring comments from review. Self-check the drawing against this list before sending to Vikram:

1. **Missing dimensions / annotations** — every critical feature must be dimensioned. Hole callouts, overall dims, bend dims, and notes must all be present. Don't rely on the model.
2. **Bend lines / flat pattern issues** — bend lines must be visible in the drawing and exported on the BEND layer of the DXF (see SolidWorks FAQ G3). Flat pattern must be valid (unsuppress Bend-Lines feature if needed). Confirm bend direction (up/down) matches the model.
3. **Wrong hole clearances on copper (CU PROFI)** — copper holes must match the CU PROFI Tool table (Section C). Common mistake: using fastener-clearance sizes from a galvanized table instead of the correct CU PROFI clearance (e.g., 1/4" bolt → 0.281", not 0.266").

---

## Standard Hardware Quick Reference

### PEM Nuts
- **3/8"-16** is the preferred PEM size at FoxFab.
- **Avoid 1/2" PEM nuts.**
- Prefer 3/8" over 1/2" wherever the design allows.

### Carriage Bolts
- Use **carriage bolts whenever a part is hard to screw with a regular bolt** (no wrench access on the back side).
- **Always use carriage bolts for mounting fuses.**
- Common for copper bus assembly where the back of the joint is inaccessible.

### Self-Tap Screws (STC)
- **#10-24 STC** → DIN rails
- **#8-32 STC** → general sheet metal applications
- In SolidWorks, model these using the **Taps** hole type (not Tapped Hole).

---

## File Naming & Revisions

### Naming
- **Top-level assembly:** `003-[user id]` (the very top of the tree)
- **Sub-assemblies:** `100-` prefix
- **Flat / sheet metal parts:** `245-` prefix

### Revision Rules
Revisions only happen **after the part has been sent to the shop**.

| Change type | Revision behavior | Example |
|---|---|---|
| **Physical change** to the part (e.g., new hole added) | Increment **letter** | rA → **rB** |
| **Drawing-only change** (e.g., missing dimension, no physical change) | Append/increment **number** to current letter | rA → **rA1** → rA2 |

---

## SolidWorks Templates

**Template location (UNC path):**
```
\\npsvr05\FOXFAB\FOXFAB_DATA\ENGINEERING\SOLIDWORKS\Foxfab Templates
```

Point SolidWorks at this folder under **Tools → Options → System Options → File Locations → Document Templates** so all new parts/assemblies/drawings use FoxFab templates.
