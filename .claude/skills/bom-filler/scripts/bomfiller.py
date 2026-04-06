"""
BOM Filler — Headless CLI Script for Claude Code Skill
-------------------------------------------------------
Two-pass BOM processor:
  Pass 1: Check each part number against the Stock Parts folder (via Everything CLI)
           and mark stock parts in the FFMPL sheet.
  Pass 2: For all NON-stock parts, find the highest-revision PDF and DXF and copy
           them into the target folder. Highlight missing parts with red fill.

Sheet markings:
  Stock parts:     B -> "X",  G -> "S",  H -> "S"
  Non-stock PDFs:  G -> "X"   (if PDF or DXF found & copied or already exists)
  Flexibar parts:  H -> "N/A" (245- prefix; no DXF exists for these parts)
  Missing parts:   Row highlighted red (both PDF and DXF not found)

Usage:
    python bomfiller.py --job J16204
    python bomfiller.py --bom "Z:\\path\\to\\BOM.xlsm"
    python bomfiller.py --bom "Z:\\path\\to\\BOM.xlsm" --target "Z:\\path\\to\\202 PDFs_Flats"
"""

import argparse
import subprocess
import sys
import os
import re
import shutil
from pathlib import Path

import xlwings as xw


# ── CONFIGURATION ──────────────────────────────────────────────────
JOBS_ROOT = r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS"
STOCK_PARTS_FOLDER = r"Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\300 Stock Parts\PDFs & Flats"

# Part number prefixes that never have DXF files (flexibars, etc.)
FLEXIBAR_PREFIXES = ("245-",)

SHEET_NAME = "FFMPL"

# Column layout (1-based)
COL_PART_NUMBER = 1   # A
COL_STOCK_PART  = 2   # B
COL_PDF         = 7   # G
COL_CNC         = 8   # H

HEADER_ROW = 5
DATA_START_ROW = 6

# Red fill RGB for missing parts
MISSING_RED = (255, 200, 200)
# ───────────────────────────────────────────────────────────────────


# ── ES.EXE LOCATION ───────────────────────────────────────────────

def _find_es_exe() -> str:
    """Locate es.exe from known project paths or system PATH."""
    # Check common project locations
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parents[3]  # scripts -> bom-filler -> skills -> .claude -> project root

    candidates = [
        project_root / "tools" / "BomFiller" / "es.exe",
        project_root / "dist" / "_internal" / "es.exe",
        project_root / "dist" / "EngineeringToolHub" / "_internal" / "es.exe",
        script_dir / "es.exe",
    ]

    # PyInstaller frozen bundle
    if getattr(sys, "frozen", False):
        bundled = Path(sys._MEIPASS) / "es.exe"
        candidates.insert(0, bundled)

    for p in candidates:
        if p.is_file():
            return str(p)

    return "es.exe"  # fall back to system PATH


ES_EXE_PATH = _find_es_exe()


# ── JOB FOLDER RESOLUTION ─────────────────────────────────────────

def resolve_job_folder(job_number: str) -> Path:
    """Find the job folder in JOBS_ROOT matching a job number prefix."""
    job_number = job_number.upper().strip()
    jobs_root = Path(JOBS_ROOT)

    if not jobs_root.is_dir():
        print(f"[ERROR] Jobs root not accessible: {JOBS_ROOT}")
        sys.exit(1)

    matches = [
        d for d in jobs_root.iterdir()
        if d.is_dir() and d.name.upper().startswith(job_number)
    ]

    if not matches:
        print(f"[ERROR] No job folder found for '{job_number}' in {JOBS_ROOT}")
        sys.exit(1)

    if len(matches) == 1:
        print(f"Job folder: {matches[0].name}")
        return matches[0]

    # Multiple matches — pick exact prefix match if possible
    exact = [d for d in matches if d.name.upper().startswith(job_number + " ") or d.name.upper() == job_number]
    if len(exact) == 1:
        print(f"Job folder: {exact[0].name}")
        return exact[0]

    print(f"\nMultiple job folders match '{job_number}':")
    for i, d in enumerate(matches, 1):
        print(f"  [{i}] {d.name}")
    while True:
        choice = input(f"Select folder (1-{len(matches)}): ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(matches):
                return matches[idx]
        except ValueError:
            pass
        print("  Invalid selection, try again.")


def find_bom_paths(job_folder: Path) -> list[tuple[Path, Path]]:
    """
    Walk folder structure patterns to find all 204 BOM directories.
    Returns list of (bom_dir, target_dir) tuples.

    Patterns checked (same as doc-prep skill):
      1. Variant-selected: job_folder IS the variant (204 BOM/ directly inside)
      2. Standard 200 Mech: 200 Mech/204 BOM/
      3. Numbered variants: 200 Mech/J#####-01/204 BOM/
      4. Named subfolders: 200 Mech/Internal/204 BOM/ or 200 Mech/J#####-01/Internal/204 BOM/
    """
    results: list[tuple[Path, Path]] = []

    # Pattern 1: 204 BOM directly in job folder (variant-selected)
    bom_dir = job_folder / "204 BOM"
    if bom_dir.is_dir():
        target = job_folder / "202 PDFs_Flats"
        results.append((bom_dir, target))
        return results

    mech = job_folder / "200 Mech"
    if not mech.is_dir():
        print(f"[ERROR] No '200 Mech' folder found in {job_folder}")
        sys.exit(1)

    # Pattern 2: Standard — 200 Mech/204 BOM/
    bom_dir = mech / "204 BOM"
    if bom_dir.is_dir():
        target = mech / "202 PDFs_Flats"
        results.append((bom_dir, target))
        return results

    # Scan subdirectories of 200 Mech
    for sub in sorted(mech.iterdir()):
        if not sub.is_dir():
            continue

        # Pattern 3: Numbered variants — 200 Mech/J#####-01/204 BOM/
        bom_dir = sub / "204 BOM"
        if bom_dir.is_dir():
            target = sub / "202 PDFs_Flats"
            results.append((bom_dir, target))
            continue

        # Pattern 4: Named subfolders — 200 Mech/J#####-01/Internal/204 BOM/
        #            or 200 Mech/Internal/204 BOM/
        for named in sorted(sub.iterdir()):
            if not named.is_dir():
                continue
            bom_dir = named / "204 BOM"
            if bom_dir.is_dir():
                target = named / "202 PDFs_Flats"
                results.append((bom_dir, target))

    return results


def pick_workbook(bom_dir: Path) -> Path:
    """Find the BOM workbook in a 204 BOM directory. Pick highest revision if multiple."""
    candidates = [
        f for f in bom_dir.iterdir()
        if f.suffix.lower() in (".xlsx", ".xlsm")
        and not f.name.startswith("~$")
    ]

    if not candidates:
        print(f"[ERROR] No .xlsx/.xlsm files in {bom_dir}")
        sys.exit(1)

    if len(candidates) == 1:
        return candidates[0]

    # Sort by revision (highest last) and pick the highest
    def rev_key(f: Path) -> tuple[str, int]:
        stem = f.stem
        _, rev = _split_stem_rev(stem, allow_subrev=True)
        return _rev_sort_key(rev)

    candidates.sort(key=rev_key)
    chosen = candidates[-1]
    print(f"  Multiple BOMs found — selected highest revision: {chosen.name}")
    return chosen


# ── EVERYTHING CLI HELPERS ─────────────────────────────────────────

def _run_es(args: list[str]) -> str | None:
    """Run es.exe with the given args. Returns stdout or None on failure."""
    try:
        result = subprocess.run(
            [ES_EXE_PATH] + args,
            capture_output=True, text=True, timeout=10,
        )
        return result.stdout.strip()
    except FileNotFoundError:
        print(f"  [ERROR] es.exe not found at '{ES_EXE_PATH}'.")
        sys.exit(1)
    except subprocess.TimeoutExpired:
        return None


# ── COMBINED PART NUMBER HELPERS ───────────────────────────────────

def _decode_combined_stem(base_stem: str) -> list[str]:
    """
    Decode a combined-part file stem into all covered part numbers.

    E.g. '240-90123_124_125' -> ['240-90123', '240-90124', '240-90125']
    Algorithm: each _<suffix> segment replaces the last len(suffix) characters
    of the previous part number.
    Returns [base_stem] unchanged if it doesn't match the combined-file pattern.
    """
    parts = base_stem.split('_')
    if len(parts) < 2:
        return [base_stem]
    base = parts[0]
    suffixes = parts[1:]
    if not all(s.isdigit() for s in suffixes):
        return [base_stem]  # not a combined file (e.g. revision or other suffix)
    covered = [base]
    current = base
    for suffix in suffixes:
        n = len(suffix)
        if len(current) < n:
            return [base_stem]  # malformed — bail out
        current = current[:-n] + suffix
        covered.append(current)
    return covered


def _split_stem_rev(stem: str, allow_subrev: bool = False) -> tuple[str, str]:
    """
    Split a file stem into (base_without_revision, revision_or_empty).
    Handles ' rB', '-rB', '_rB' separators.
    When *allow_subrev* is True, also captures numeric sub-revisions
    like rA1, rA2, rB12.
    E.g. '240-90123_124 rB'  -> ('240-90123_124', 'rB')
         '250-31025-rA2'     -> ('250-31025', 'rA2')  (when allow_subrev=True)
    """
    pat = r'[-_\s]r([A-Za-z]\d*)$' if allow_subrev else r'[-_\s]r([A-Za-z])$'
    m = re.search(pat, stem, re.IGNORECASE)
    if m:
        raw = m.group(1)
        normalised = raw[0].upper() + raw[1:]
        return stem[:m.start()].rstrip(), f"r{normalised}"
    return stem, ""


def _rev_sort_key(rev: str) -> tuple[str, int]:
    """Return a sortable key for revision strings like 'rA', 'rA1', 'rB12'.
    No-rev ('') sorts lowest."""
    if not rev:
        return ("", -1)
    letter = rev[1].upper()
    num = int(rev[2:]) if len(rev) > 2 else -1
    return (letter, num)


def _strip_bom_rev(raw_pn: str) -> tuple[str, str]:
    """Strip a revision suffix from a BOM part-number cell value.

    BOM cells sometimes carry the revision inline:
        '250-20120 RB'  -> ('250-20120', 'rB')
        '240-70964 RA'  -> ('240-70964', 'rA')
        '250-40239 RC'  -> ('250-40239', 'rC')
        '250-20120 rB'  -> ('250-20120', 'rB')
        '250-20120 RevB' -> ('250-20120', 'rB')

    Only strips when the separator is a SPACE (to avoid false positives on
    dash/underscore suffixes like '250-30834_837' or '240-30589-001').
    Returns (base_part_number, normalised_revision_or_empty).
    """
    # Pattern: space, optional "r"/"rev", then one letter + optional digits at end
    m = re.search(r'\s+(?:rev|r)?([A-Za-z]\d*)$', raw_pn, re.IGNORECASE)
    if m:
        rev_str = m.group(1)
        if len(rev_str) >= 1 and rev_str[0].isalpha():
            base = raw_pn[:m.start()].strip()
            normalised = f"r{rev_str[0].upper()}{rev_str[1:]}"
            return base, normalised
    return raw_pn, ""


def _strip_config_suffix(part_number: str) -> str:
    """Strip a config/variant suffix from a part number for fallback search.

    Config suffixes follow the base ###-##### pattern:
        '240-30589-001'  -> '240-30589'
        '250-40036-002'  -> '250-40036'
        '250-20058-001'  -> '250-20058'

    Also handles underscore config suffixes (already in stock check):
        '250-30834_001'  -> '250-30834'

    Returns the original part number unchanged if no config suffix detected.
    """
    # Dash config: ###-#####-### (3-segment)
    base = re.sub(r'^(\d{3}-\d{5})-\d+$', r'\1', part_number)
    if base != part_number:
        return base
    # Underscore config: ###-#####_###
    base = re.sub(r'_\d+$', '', part_number)
    if base != part_number:
        return base
    return part_number


# ── STOCK CHECK ────────────────────────────────────────────────────

def check_part_in_stock_folder(part_number: str) -> bool:
    """
    Check if a part number exists anywhere in the stock parts folder.
    Also handles config-variant parts (e.g. BOM has 250-30834_001, stock folder
    has 250-30834_002): strips the _### suffix and retries with the base number.
    """
    output = _run_es(["-path", STOCK_PARTS_FOLDER, part_number])
    if output is None:
        print(f"  [WARNING] Search timed out for '{part_number}'. Skipping.")
        return False
    if any(line.strip() for line in output.splitlines()):
        return True

    # If the part has a config suffix (_digits), also try the base part number
    base = re.sub(r'_\d+$', '', part_number)
    if base != part_number:
        output = _run_es(["-path", STOCK_PARTS_FOLDER, base])
        if output is None:
            return False
        if any(line.strip() for line in output.splitlines()):
            return True

    return False


# ── REVISION DETECTION ─────────────────────────────────────────────

def find_highest_revision(part_number: str, ext: str = "") -> str:
    """
    Search for all files matching a part number and return the highest
    revision string (e.g. "rC", "rA2"), or "" if none found.
    Handles combined-part filenames like '240-90123_124 rB.pdf'.
    When *ext* is given (e.g. "pdf", "dxf") only files of that type are considered.
    For PDFs, numeric sub-revisions (rA1, rA2, rB12) are recognised.
    """
    output = _run_es([part_number]) if not ext else _run_es([f"ext:{ext}", part_number])
    if not output:
        return ""

    subrev = (ext.lower() == "pdf")
    rev_pat = r'^[-_\s]?r([A-Za-z]\d*)$' if subrev else r'^[-_\s]?r([A-Za-z])$'
    highest = ""
    highest_key: tuple[str, int] = ("", -1)
    for line in output.splitlines():
        line = line.strip()
        if ext and not line.upper().endswith(f".{ext.upper()}"):
            continue
        filename = Path(line).stem
        if not filename.upper().startswith(part_number.upper()):
            continue
        suffix = filename[len(part_number):].strip()
        # Strip combined-part segments (e.g. _124_125) before checking revision
        suffix = re.sub(r'^(_\d+)+', '', suffix).strip()
        match = re.match(rev_pat, suffix)
        if match:
            raw = match.group(1)
            normalised = f"r{raw[0].upper()}{raw[1:]}"
            key = _rev_sort_key(normalised)
            if key > highest_key:
                highest_key = key
                highest = normalised

    return highest


# ── FILE FIND & COPY ───────────────────────────────────────────────

def find_and_copy_file(part_number: str, revision: str, target_path: Path,
                       ext: str,
                       standalone_only: bool = False) -> tuple[bool | None, list[str]]:
    """
    Search for the file matching part_number (with optional revision) and copy
    it to the target folder. Also detects combined-part filenames like
    '240-90123_124.pdf' which cover multiple part numbers.

    When *standalone_only* is True, only files that cover exactly one part
    (the requested part_number) are considered — combined-part files are skipped.

    Returns:
        (result, covered_parts) where:
          result        = True (copied), None (already exists), False (not found)
          covered_parts = list of part numbers covered by the matched file
    """
    ext_upper = f".{ext.upper()}"

    # ── Collect candidate file paths ───────────────────────────────
    candidates: set[str] = set()
    out = _run_es([f"ext:{ext}", part_number])
    if out:
        candidates.update(line.strip() for line in out.splitlines() if line.strip())

    # Broader search: part_number might appear as a DERIVED part in a combined
    # file (e.g. 240-90124 won't match 240-90123_124.pdf directly).
    if len(part_number) > 5:
        for drop in (2, 3):
            prefix = part_number[:-drop]
            extra = _run_es([f"ext:{ext}", prefix])
            if extra:
                candidates.update(
                    line.strip() for line in extra.splitlines() if line.strip()
                )

    if not candidates:
        return False, []

    # ── Build filtered list, then sort so highest revision comes first.
    viable: list[tuple[Path, str, list[str]]] = []
    for line in candidates:
        p = Path(line)
        if p.suffix.upper() != ext_upper:
            continue
        if "PUNCH PROGRAM" in str(p).upper():
            continue

        subrev = (ext.lower() == "pdf")
        clean_stem, file_rev = _split_stem_rev(p.stem, allow_subrev=subrev)
        covered_parts = _decode_combined_stem(clean_stem)
        covered_upper = [c.upper() for c in covered_parts]

        # Check if this file covers our part number
        if part_number.upper() not in covered_upper:
            continue

        if standalone_only and len(covered_parts) > 1:
            continue

        # Revision compatibility: if a revision was found, require it to match
        if revision and file_rev and file_rev.upper() != revision.upper():
            continue

        viable.append((p, file_rev, covered_parts))

    # Sort: highest revision first (using letter + sub-number), no-rev last
    viable.sort(key=lambda t: _rev_sort_key(t[1]), reverse=True)

    for p, _fr, covered_parts in viable:
        if not p.exists():
            continue
        dest = target_path / p.name
        if dest.exists():
            return None, covered_parts
        shutil.copy2(str(p), str(dest))
        return True, covered_parts

    return False, []


# ── MISSING ROW HIGHLIGHT ──────────────────────────────────────────

def highlight_missing_row(ws, row: int, last_col: int = 8):
    """Apply a light-red background fill to columns A through last_col for a row."""
    r, g, b = MISSING_RED
    for col in range(1, last_col + 1):
        ws.range((row, col)).color = (r, g, b)


# ── MAIN WORKFLOW ──────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="BOM Filler — stock check + PDF/DXF copy")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--job", help="Job number (e.g. J16204) — auto-locates BOM in job folder")
    group.add_argument("--bom", help="Direct path to BOM workbook (.xlsx/.xlsm) or its directory")
    parser.add_argument("--target", help="Target folder for copied PDFs/DXFs (default: auto-detect ../202 PDFs_Flats)")
    args = parser.parse_args()

    # ── Resolve BOM workbook and target folder ─────────────────────
    if args.job:
        job_folder = resolve_job_folder(args.job)
        bom_pairs = find_bom_paths(job_folder)

        if not bom_pairs:
            print(f"[ERROR] No '204 BOM' folders found in job {args.job}")
            sys.exit(1)

        if len(bom_pairs) == 1:
            bom_dir, target_dir = bom_pairs[0]
        else:
            print(f"\nMultiple BOM locations found:")
            for i, (bd, td) in enumerate(bom_pairs, 1):
                # Show relative path from job folder for clarity
                try:
                    rel = bd.relative_to(job_folder)
                except ValueError:
                    rel = bd
                print(f"  [{i}] {rel}")
            while True:
                choice = input(f"Select BOM (1-{len(bom_pairs)}): ").strip()
                try:
                    idx = int(choice) - 1
                    if 0 <= idx < len(bom_pairs):
                        bom_dir, target_dir = bom_pairs[idx]
                        break
                except ValueError:
                    pass
                print("  Invalid selection, try again.")

        wb_path = pick_workbook(bom_dir)
        if args.target:
            target_path = Path(args.target)
        else:
            target_path = target_dir

    else:
        # --bom path provided
        raw = args.bom.strip().strip('"')
        if os.path.isdir(raw):
            bom_dir = Path(raw)
            wb_path = pick_workbook(bom_dir)
        elif os.path.isfile(raw):
            wb_path = Path(raw)
            bom_dir = wb_path.parent
        else:
            # Try appending extensions
            for ext in (".xlsx", ".xlsm"):
                if os.path.isfile(raw + ext):
                    wb_path = Path(raw + ext)
                    bom_dir = wb_path.parent
                    break
            else:
                print(f"[ERROR] File not found: {raw}")
                sys.exit(1)

        if args.target:
            target_path = Path(args.target)
        else:
            target_path = Path(os.path.normpath(bom_dir / ".." / "202 PDFs_Flats"))

    # Ensure target folder exists
    target_path.mkdir(parents=True, exist_ok=True)

    print(f"\nBOM workbook: {wb_path}")
    print(f"Target folder: {target_path}")

    # ── Open workbook via COM ──────────────────────────────────────
    print(f"\nOpening workbook in Excel...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(str(wb_path))
    except Exception as e:
        print(f"[ERROR] Could not open workbook: {e}")
        app.quit()
        sys.exit(1)

    sheet_names = [s.name for s in wb.sheets]
    if SHEET_NAME not in sheet_names:
        print(f"[ERROR] Sheet '{SHEET_NAME}' not found. Available: {sheet_names}")
        wb.close()
        app.quit()
        sys.exit(1)

    ws = wb.sheets[SHEET_NAME]

    # ── Read data in one shot ──────────────────────────────────────
    last_row = ws.range(f"A{DATA_START_ROW}").end("down").row
    if last_row > 1_000_000:
        last_row = DATA_START_ROW

    part_numbers = ws.range(f"A{DATA_START_ROW}:A{last_row}").value
    if part_numbers is None:
        part_numbers = []
    elif not isinstance(part_numbers, list):
        part_numbers = [part_numbers]

    total_parts = sum(1 for pn in part_numbers if pn and str(pn).strip())
    print(f"Found {total_parts} part numbers (rows {DATA_START_ROW}-{last_row})")

    # ══════════════════════════════════════════════════════════════
    #  PASS 1: Stock parts check
    # ══════════════════════════════════════════════════════════════
    print(f"\n{'=' * 60}")
    print(f"  PASS 1 — Stock Parts Check")
    print(f"  Searching in: {STOCK_PARTS_FOLDER}")
    print(f"{'=' * 60}")

    stock_found = 0
    stock_checked = 0
    stock_flags = []            # track which rows are stock for Pass 2

    for i, pn in enumerate(part_numbers):
        row = DATA_START_ROW + i

        if pn is None or str(pn).strip() == "":
            stock_flags.append(False)
            print(f"  Row {row}: (blank — skipped)")
            continue

        raw_pn = str(pn).strip()
        base_pn, embedded_rev = _strip_bom_rev(raw_pn)
        part_number = base_pn          # use base PN for all searches
        stock_checked += 1
        label = f"{raw_pn}" if not embedded_rev else f"{raw_pn}  [base: {base_pn}]"
        print(f"  Checking: {label:<40}", end="", flush=True)

        if check_part_in_stock_folder(part_number):
            ws.range(f"B{row}").value = "X"
            ws.range(f"G{row}").value = "S"
            ws.range(f"H{row}").value = "S"
            print("  FOUND — marked X / S / S")
            stock_found += 1
            stock_flags.append(True)
        else:
            print("  not found")
            stock_flags.append(False)

    print(f"\nPass 1 results: {stock_found} of {stock_checked} parts found in stock folder.")

    # ══════════════════════════════════════════════════════════════
    #  PASS 2: Non-stock PDF + DXF copy
    # ══════════════════════════════════════════════════════════════
    print(f"\n{'=' * 60}")
    print(f"  PASS 2 — Non-Stock PDF + DXF Copy")
    print(f"  Target folder: {target_path}")
    print(f"{'=' * 60}")

    pdf_copied = 0
    pdf_not_found = 0
    pdf_skipped_existing = 0
    dxf_copied = 0
    dxf_not_found = 0
    dxf_skipped_existing = 0
    dxf_na = 0
    non_stock_checked = 0
    missing_parts: list[tuple[str, int, str]] = []  # (part_number, row, what_missing)

    # Build part-number → row index for cross-marking combined parts
    # Map BOTH the raw cell value AND the stripped base PN so combined-part
    # cross-marking works regardless of whether the BOM cell had a rev suffix.
    pn_to_row: dict[str, int] = {}
    for i, pn in enumerate(part_numbers):
        if pn and str(pn).strip():
            raw = str(pn).strip().upper()
            pn_to_row[raw] = DATA_START_ROW + i
            base, _ = _strip_bom_rev(raw)
            if base.upper() != raw:
                pn_to_row[base.upper()] = DATA_START_ROW + i

    for i, pn in enumerate(part_numbers):
        row = DATA_START_ROW + i

        if pn is None or str(pn).strip() == "":
            print(f"  Row {row}: (blank — skipped)")
            continue

        # Skip anything Pass 1 already tagged as stock
        if stock_flags[i]:
            continue

        raw_pn = str(pn).strip()
        base_pn, embedded_rev = _strip_bom_rev(raw_pn)
        part_number = base_pn          # use base PN for all searches
        is_flexibar = part_number.startswith(FLEXIBAR_PREFIXES)
        non_stock_checked += 1

        if embedded_rev:
            print(f"  {raw_pn:<25} [base:{base_pn}]", end="", flush=True)
        else:
            print(f"  {part_number:<25}", end="", flush=True)

        # ── Detect revisions ──────────────────────────────────
        pdf_rev = find_highest_revision(part_number, "pdf")
        dxf_rev = find_highest_revision(part_number, "dxf") if not is_flexibar else ""

        # If BOM cell had an embedded revision, prefer it when it matches
        # what exists; otherwise fall back to latest found revision.
        if embedded_rev:
            if pdf_rev and _rev_sort_key(embedded_rev) <= _rev_sort_key(pdf_rev):
                pdf_rev = embedded_rev
            if dxf_rev and _rev_sort_key(embedded_rev) <= _rev_sort_key(dxf_rev):
                dxf_rev = embedded_rev

        rev_lbl = f" [PDF:{pdf_rev or '–'} DXF:{dxf_rev or '–'}]"
        print(f"{rev_lbl:<20}", end="", flush=True)

        # ── PDF ────────────────────────────────────────────────
        # Prefer standalone PDF; fall back to combined.
        pdf_result, pdf_covered = find_and_copy_file(part_number, pdf_rev, target_path, "pdf",
                                                     standalone_only=True)
        if pdf_result is False:
            pdf_result, pdf_covered = find_and_copy_file(part_number, pdf_rev, target_path, "pdf")

        # Config-suffix fallback: try base PN if exact match failed
        if pdf_result is False:
            cfg_base = _strip_config_suffix(part_number)
            if cfg_base != part_number:
                cfg_pdf_rev = find_highest_revision(cfg_base, "pdf")
                pdf_result, pdf_covered = find_and_copy_file(cfg_base, cfg_pdf_rev, target_path, "pdf",
                                                              standalone_only=True)
                if pdf_result is False:
                    pdf_result, pdf_covered = find_and_copy_file(cfg_base, cfg_pdf_rev, target_path, "pdf")
                if pdf_result is not False:
                    print(f"  [cfg->{cfg_base}]", end="")

        if pdf_result is True:
            print("  PDF:COPIED", end="")
            pdf_copied += 1
        elif pdf_result is None:
            print("  PDF:EXISTS", end="")
            pdf_skipped_existing += 1
        else:
            print("  PDF:NOT FOUND", end="")
            pdf_not_found += 1

        # ── DXF ────────────────────────────────────────────────
        if is_flexibar:
            dxf_result, dxf_covered = False, []
            print("  DXF:N/A (flexibar)", end="")
            dxf_na += 1
        else:
            dxf_result, dxf_covered = find_and_copy_file(part_number, dxf_rev, target_path, "dxf")

            # Config-suffix fallback for DXF
            if dxf_result is False:
                cfg_base = _strip_config_suffix(part_number)
                if cfg_base != part_number:
                    cfg_dxf_rev = find_highest_revision(cfg_base, "dxf")
                    dxf_result, dxf_covered = find_and_copy_file(cfg_base, cfg_dxf_rev, target_path, "dxf")
                    if dxf_result is not False:
                        print(f"  [cfg->{cfg_base}]", end="")

            if dxf_result is True:
                print("  DXF:COPIED", end="")
                dxf_copied += 1
            elif dxf_result is None:
                print("  DXF:EXISTS", end="")
                dxf_skipped_existing += 1
            else:
                print("  DXF:NOT FOUND", end="")
                dxf_not_found += 1

        # ── Mark column G if either PDF or DXF was found/copied ──
        if pdf_result is not False or dxf_result is not False:
            ws.range(f"G{row}").value = "X"

        # ── Mark column H as N/A for flexibars with a found PDF ──
        if is_flexibar and pdf_result is not False:
            ws.range(f"H{row}").value = "N/A"

        # ── Track and highlight missing parts ──────────────────
        pdf_missing = pdf_result is False
        dxf_missing = (dxf_result is False) and not is_flexibar

        if pdf_missing and dxf_missing:
            missing_parts.append((part_number, row, "PDF + DXF"))
            highlight_missing_row(ws, row)
            print("  [MISSING]", end="")
        elif pdf_missing:
            missing_parts.append((part_number, row, "PDF"))
            highlight_missing_row(ws, row)
            print("  [MISSING PDF]", end="")
        elif dxf_missing:
            missing_parts.append((part_number, row, "DXF"))
            highlight_missing_row(ws, row)
            print("  [MISSING DXF]", end="")

        print()  # newline after each part

        # ── Cross-mark other BOM rows covered by combined files ──
        for covered_list in (pdf_covered, dxf_covered):
            for other_pn in covered_list:
                if other_pn.upper() == part_number.upper():
                    continue
                if other_pn.upper() in pn_to_row:
                    other_row = pn_to_row[other_pn.upper()]
                    ws.range(f"G{other_row}").value = "X"
                    print(f"    Combined: also marked {other_pn} (row {other_row})")
                else:
                    print(f"    [WARN] Combined file covers {other_pn} — not in BOM")

    # ── Summary ────────────────────────────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"  SUMMARY")
    print(f"{'=' * 60}")
    print(f"  Pass 1 (Stock):    {stock_found} of {stock_checked} parts found")
    print(f"  Pass 2 (Non-stock): {non_stock_checked} parts checked")
    print(f"  PDFs  — copied: {pdf_copied},  existed: {pdf_skipped_existing},  not found: {pdf_not_found}")
    dxf_summary = f"  DXFs  — copied: {dxf_copied},  existed: {dxf_skipped_existing},  not found: {dxf_not_found}"
    if dxf_na:
        dxf_summary += f",  N/A (flexibar): {dxf_na}"
    print(dxf_summary)

    if missing_parts:
        print(f"\n  [MISSING] {len(missing_parts)} part(s) need attention (highlighted red in BOM):")
        for pn, row, what in missing_parts:
            print(f"    Row {row}: {pn} — {what}")
    else:
        print(f"\n  All parts accounted for!")

    # ── Save & cleanup ─────────────────────────────────────────────
    changes = stock_found + pdf_copied + pdf_skipped_existing + dxf_copied + dxf_skipped_existing + len(missing_parts)
    if changes > 0:
        wb.save()
        print(f"\nWorkbook saved: {wb_path}")
    else:
        print("\nNo changes made.")

    wb.close()
    app.quit()
    print("Done.")


if __name__ == "__main__":
    main()
