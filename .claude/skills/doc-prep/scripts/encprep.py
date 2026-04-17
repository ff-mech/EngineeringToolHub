r"""
Enclosure Package Prep — mech-docs-only doc prep for standard enclosure product folders.

Used for folders under Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\100 Standard Enclosures
(or any folder that contains 204 BOM / 205 CNC / 202 PDFs_Flats / 203 Assemblies directly).

Skips FWO, PRF, and Electrical Pack since those only exist inside job folders.

Usage:
    python encprep.py --path "Z:\...\WM-363016-SS-N3R"
    python encprep.py --path "Z:\...\WM-363016-SS-N3R" --print
"""

import argparse
import shutil
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docprep import (
    PREFERRED_PRINTER,
    _has_mech_subs,
    _missing,
    error,
    generate_documents,
    log,
    make_breakdown,
    match_assemblies,
    match_cnc,
    match_excel,
    match_flats,
    print_documents,
    safe_name,
    warn,
)


FWO_TEMPLATE = SCRIPT_DIR.parent / "Fabrication Work Order - Standard v1.0.pdf"


def build_product_plan(product_folder: Path) -> dict:
    if not product_folder.is_dir():
        raise RuntimeError(f"Product folder not found: {product_folder}")
    if not _has_mech_subs(product_folder):
        raise RuntimeError(
            "Product folder missing one of: 204 BOM, 205 CNC, 202 PDFs_Flats, 203 Assemblies"
        )

    # FWO: use the template bundled with the skill; fill only Date + Enclosure (= folder name)
    fwo_path = FWO_TEMPLATE if FWO_TEMPLATE.is_file() else None
    if fwo_path is None:
        warn(f"FWO template not found at {FWO_TEMPLATE}; skipping FWO.")
    prf_data = {
        "job_no": "",
        "job_name": "",
        "enclosure": product_folder.name,
        "qty": "",
        "model_no": "",
    } if fwo_path else None

    plan = {
        "job_folder": str(product_folder),
        "ref_folder": None,
        "base": str(product_folder),
        "main_base": str(product_folder),
        "mech_roots": [str(product_folder)],
        "variant_only": False,
        "sources": {"fwo": str(FWO_TEMPLATE.parent) if fwo_path else ""},
        "fwo": fwo_path,
        "prf": None,
        "pack": None,
        "prf_data": prf_data,
    }

    plan["bom"] = match_excel(product_folder / "204 BOM", "BOM", "BOM")
    plan["sources"]["bom"] = str(product_folder)

    cnc = match_cnc(product_folder / "205 CNC")
    flats = match_flats(product_folder / "202 PDFs_Flats")
    assemblies, excluded_lay = match_assemblies(product_folder / "203 Assemblies")

    plan["cnc"] = sorted(cnc, key=lambda p: p.name.lower())
    plan["flats"] = sorted(flats, key=lambda p: p.name.lower())
    plan["assemblies"] = sorted(assemblies, key=lambda p: p.name.lower())
    plan["excluded_lay"] = sorted(excluded_lay, key=lambda p: p.name.lower())
    plan["sources"]["cnc"] = str(product_folder)
    plan["sources"]["flats"] = str(product_folder)
    plan["sources"]["assemblies"] = str(product_folder)

    cnc_folder = product_folder / "205 CNC"
    plan["cnc_mark_folder"] = cnc_folder if cnc_folder.is_dir() else None
    return plan


def run(product_path: str, do_print: bool = False):
    if _missing:
        print("\n[WARNING] Missing dependencies:")
        for dep in _missing:
            print(f"  - {dep}")

    folder = Path(product_path)
    log(f"Enclosure package folder: {folder}")

    plan = build_product_plan(folder)

    name = safe_name(folder.name)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_tmp = Path(tempfile.gettempdir()) / "FoxFab_DocPrep"
    base_tmp.mkdir(parents=True, exist_ok=True)

    # Cleanup old outputs (>7 days)
    try:
        cutoff = time.time() - (7 * 86400)
        for old in base_tmp.iterdir():
            if old.is_dir() and old.name.startswith(("DocPrep_Output_", "EncPrep_Output_")):
                try:
                    if old.stat().st_mtime < cutoff:
                        shutil.rmtree(old)
                except Exception:
                    pass
    except Exception:
        pass

    out_dir = base_tmp / f"EncPrep_Output_{name}_{ts}"

    log("Generating documents...")
    generated = generate_documents(plan, out_dir)

    if not generated:
        error("No documents were generated.")
        return out_dir, "No documents generated."

    breakdown = make_breakdown(plan, generated, out_dir)
    print("\n" + breakdown)
    (out_dir / "BREAKDOWN.txt").write_text(breakdown, encoding="utf-8")

    if do_print:
        log(f"Sending {len(generated)} documents to printer: {PREFERRED_PRINTER}")
        print_documents(generated, PREFERRED_PRINTER)
    else:
        log(f"Simulation complete. {len(generated)} PDFs saved to: {out_dir}")

    return out_dir, breakdown


def main():
    p = argparse.ArgumentParser(description="FoxFab Enclosure Package Prep (mech docs only)")
    p.add_argument("--path", required=True, help="Path to the product folder containing 204 BOM / 205 CNC / 202 PDFs_Flats / 203 Assemblies")
    p.add_argument("--print", dest="do_print", action="store_true", help="Send to printer after generating")
    args = p.parse_args()

    try:
        run(args.path, args.do_print)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
