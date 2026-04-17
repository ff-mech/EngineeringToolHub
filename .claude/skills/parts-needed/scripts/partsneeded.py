"""
Parts Needed — Extract P/N + Description + Qty from a FoxFab job's BOM.

Usage:
    python partsneeded.py --job J16204
    python partsneeded.py --job J16204 --bom "Z:\\path\\to\\specific BOM.xlsx"
    python partsneeded.py --job J16204 --jobs-root "//server/share/Testing"

Output: prints a JSON document to stdout describing the chosen BOM and its rows.
Errors of the form 'CHOOSE_BOM:', 'NO_BOM_FOUND:', 'UNKNOWN_JOB:', 'PDF_NEEDS_VISUAL:'
are written to stderr and exit non-zero so Claude can react.
"""

import argparse
import csv
import difflib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

# ── Configuration ──────────────────────────────────────────────────────
JOBS_ROOT      = Path(r"Z:\FOXFAB_DATA\ENGINEERING\2 JOBS")
BOM_SUBPATH    = Path("100 Elec") / "101 Bill of Materials"
PRF_SUBPATH    = Path("300 Inputs") / "302 Production Release Form"
BOM_EXTS       = {".xlsx", ".xlsm", ".xls", ".csv", ".pdf"}
EXCEL_EXTS     = {".xlsx", ".xlsm", ".xls"}
AUTO_PICK_MIN  = 0.85
AUTO_PICK_GAP  = 0.10

PN_PAT   = re.compile(r"^\s*(p\.?\s*/?\s*n\.?|part\s*(no\.?|number|#)|mfr\s*p/?n)\s*$", re.I)
DESC_PAT = re.compile(r"^\s*(description|desc|item description|part description)\s*$", re.I)
QTY_PAT  = re.compile(r"^\s*(qty|quantity|qnty|qty\.?)\s*$", re.I)


# ── Errors ─────────────────────────────────────────────────────────────
def die(code: str, msg: str = "", payload=None):
    if payload is not None:
        sys.stderr.write(f"{code}:{json.dumps(payload)}\n")
    elif msg:
        sys.stderr.write(f"{code}: {msg}\n")
    else:
        sys.stderr.write(f"{code}\n")
    sys.exit(1)


# ── Job folder resolution ──────────────────────────────────────────────
def find_job_folder(job_number: str) -> Path:
    if not JOBS_ROOT.is_dir():
        die("UNKNOWN_JOB", f"Jobs root not found: {JOBS_ROOT}")
    job_up = job_number.upper()
    matches = [d for d in JOBS_ROOT.iterdir()
               if d.is_dir() and d.name.upper().startswith(job_up)]
    if not matches:
        die("UNKNOWN_JOB", f"No folder matches '{job_number}' in {JOBS_ROOT}")
    if len(matches) > 1:
        die("CHOOSE_FOLDER", payload=[m.name for m in matches])
    return matches[0]


# ── PRF discovery ──────────────────────────────────────────────────────
def find_prf(job_folder: Path) -> Path | None:
    prf_folder = job_folder / PRF_SUBPATH
    if not prf_folder.is_dir():
        return None
    prfs = [f for f in prf_folder.iterdir()
            if f.is_file()
            and f.suffix.lower() in EXCEL_EXTS
            and "prf" in f.name.lower()]
    if not prfs:
        return None
    return sorted(prfs, key=lambda p: p.name.lower())[0]


# ── BOM scoring ────────────────────────────────────────────────────────
def list_boms(job_folder: Path) -> list[Path]:
    bom_folder = job_folder / BOM_SUBPATH
    if not bom_folder.is_dir():
        die("NO_BOM_FOUND", f"Folder missing: {bom_folder}")
    boms = [f for f in bom_folder.iterdir()
            if f.is_file()
            and f.suffix.lower() in BOM_EXTS
            and not f.name.startswith("~$")]
    if not boms:
        die("NO_BOM_FOUND", f"No BOM files in {bom_folder}")
    return boms


def score_bom(prf_stem: str, bom_stem: str) -> float:
    return difflib.SequenceMatcher(
        None, prf_stem.lower(), bom_stem.lower()).ratio()


def pick_bom(prf: Path | None, boms: list[Path]) -> Path:
    if len(boms) == 1:
        return boms[0]
    if prf is None:
        die("CHOOSE_BOM", payload=[{"name": b.name, "path": str(b), "score": 0.0}
                                    for b in boms])
    scored = sorted(
        [(score_bom(prf.stem, b.stem), b) for b in boms],
        key=lambda t: t[0], reverse=True)
    top_score, top_bom = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    if top_score >= AUTO_PICK_MIN and (top_score - second_score) >= AUTO_PICK_GAP:
        return top_bom
    die("CHOOSE_BOM", payload=[
        {"name": b.name, "path": str(b), "score": round(s, 3)}
        for s, b in scored[:5]])


# ── Header detection helpers ───────────────────────────────────────────
def detect_header_idx(rows: list[list[str]]) -> tuple[int, int, int, int] | None:
    """Return (header_row_index, pn_col, desc_col, qty_col) or None."""
    for i, row in enumerate(rows[:30]):
        cells = [str(c).strip() if c is not None else "" for c in row]
        pn = desc = qty = -1
        for j, c in enumerate(cells):
            if pn < 0 and PN_PAT.match(c):
                pn = j
            elif desc < 0 and DESC_PAT.match(c):
                desc = j
            elif qty < 0 and QTY_PAT.match(c):
                qty = j
        if pn >= 0 and desc >= 0 and qty >= 0:
            return i, pn, desc, qty
    return None


def rows_from_grid(grid: list[list], pn_col: int, desc_col: int, qty_col: int,
                   start: int) -> list[dict]:
    out = []
    for row in grid[start:]:
        if not row:
            continue
        def cell(idx):
            if idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()
        pn = cell(pn_col)
        if not pn:
            continue
        out.append({"pn": pn, "desc": cell(desc_col), "qty": cell(qty_col)})
    return out


# ── DWG-embedded xlsx detection + EMF preview extraction ──────────────
def xlsx_has_embedded_dwg(path: Path) -> bool:
    """True if the xlsx's sheet1 is empty but it contains an OLE embedding
    whose payload begins with 'AC10' (AutoCAD DWG magic)."""
    import zipfile, re as _re
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            sheet = "xl/worksheets/sheet1.xml"
            if sheet not in names:
                return False
            xml = z.read(sheet).decode("utf-8", errors="replace")
            if _re.search(r"<row[\s>]", xml):
                return False  # real data present
            embeds = [n for n in names if n.startswith("xl/embeddings/") and n.endswith(".bin")]
            if not embeds:
                return False
            import olefile, io as _io
            blob = z.read(embeds[0])
            try:
                ole = olefile.OleFileIO(_io.BytesIO(blob))
            except Exception:
                return False
            for stream in ole.listdir():
                head = ole.openstream(stream).read(16)
                if head.startswith(b"AC10"):
                    return True
    except Exception:
        pass
    return False


def extract_emf_as_png(xlsx_path: Path, out_png: Path, dpi: int = 300) -> bool:
    """Pull xl/media/image1.emf from the xlsx and render it to PNG at the given DPI.
    Returns True on success."""
    import zipfile
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            emfs = [n for n in z.namelist()
                    if n.startswith("xl/media/") and n.lower().endswith(".emf")]
            if not emfs:
                return False
            emf_bytes = z.read(emfs[0])
    except Exception:
        return False

    out_png.parent.mkdir(parents=True, exist_ok=True)
    tmp_emf = out_png.with_suffix(".emf")
    tmp_emf.write_bytes(emf_bytes)

    # Try Pillow first — its WMF/EMF plugin handles EMF on Windows via GDI.
    try:
        from PIL import Image
        img = Image.open(str(tmp_emf))
        try:
            img.load(dpi=dpi)
        except TypeError:
            img.load()
        img.save(str(out_png), "PNG")
        tmp_emf.unlink(missing_ok=True)
        return True
    except Exception:
        pass

    # Fallback: pywin32 GDI+ — play the metafile into a high-res bitmap.
    try:
        import win32ui, win32con, win32gui
        from ctypes import windll
        hemf = windll.gdi32.GetEnhMetaFileW(str(tmp_emf))
        if not hemf:
            return False
        # Query header for source size in device units
        import struct
        header_size = 88
        buf = (windll.gdi32.GetEnhMetaFileHeader(hemf, header_size, None))
        # Just render at a generous fixed size; precise DPI handling via Pillow
        # is usually enough and this fallback is rare.
        w = int(11 * dpi)  # assume ~11 inch wide
        h = int(8.5 * dpi)
        hdc_screen = win32gui.GetDC(0)
        mem_dc = win32ui.CreateDCFromHandle(hdc_screen).CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(win32ui.CreateDCFromHandle(hdc_screen), w, h)
        mem_dc.SelectObject(bmp)
        mem_dc.FillSolidRect((0, 0, w, h), 0xFFFFFF)
        rect = (0, 0, w, h)
        windll.gdi32.PlayEnhMetaFile(mem_dc.GetHandleOutput(), hemf, rect)
        bmp.SaveBitmapFile(mem_dc, str(out_png.with_suffix(".bmp")))
        from PIL import Image
        Image.open(str(out_png.with_suffix(".bmp"))).save(str(out_png), "PNG")
        out_png.with_suffix(".bmp").unlink(missing_ok=True)
        tmp_emf.unlink(missing_ok=True)
        windll.gdi32.DeleteEnhMetaFile(hemf)
        return True
    except Exception:
        return False


# ── Excel reader ───────────────────────────────────────────────────────
def read_excel(path: Path) -> list[dict]:
    try:
        import openpyxl
    except ImportError:
        die("MISSING_DEP", "openpyxl not installed (pip install openpyxl)")
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    for ws in wb.worksheets:
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        if not grid:
            continue
        hdr = detect_header_idx([[str(c) if c is not None else "" for c in row]
                                 for row in grid])
        if hdr is None:
            continue
        i, pn, desc, qty = hdr
        rows = rows_from_grid(grid, pn, desc, qty, i + 1)
        if rows:
            return rows
    return []


# ── CSV reader ─────────────────────────────────────────────────────────
def read_csv(path: Path) -> list[dict]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        grid = [row for row in csv.reader(f)]
    hdr = detect_header_idx(grid)
    if hdr is None:
        return []
    i, pn, desc, qty = hdr
    return rows_from_grid(grid, pn, desc, qty, i + 1)


# ── PDF reader (text-first, visual-fallback signal) ────────────────────
def read_pdf(path: Path) -> list[dict]:
    rows: list[dict] = []
    # Try pdfplumber tables
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            grid: list[list] = []
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    grid.extend(table)
            if grid:
                hdr = detect_header_idx(grid)
                if hdr is not None:
                    i, pn, desc, qty = hdr
                    rows = rows_from_grid(grid, pn, desc, qty, i + 1)
                    if rows:
                        return rows
    except ImportError:
        pass
    except Exception:
        pass

    # Fallback: pymupdf raw text — try line-based parsing
    try:
        import fitz
        text_lines = []
        with fitz.open(str(path)) as doc:
            for page in doc:
                text_lines.extend(page.get_text().splitlines())
        # Best-effort: look for lines that start with a P/N-ish token
        for line in text_lines:
            m = re.match(r"^\s*([A-Z0-9][A-Z0-9\-./]{3,})\s+(.+?)\s+(\d+(?:\.\d+)?)\s*$", line)
            if m:
                rows.append({"pn": m.group(1), "desc": m.group(2).strip(),
                             "qty": m.group(3)})
        if rows:
            return rows
    except ImportError:
        pass
    except Exception:
        pass

    return rows


VISION_MAX_SIDE = 1500   # Stay under Claude's ~1568 vision input limit
OVERLAP_FRAC    = 0.12   # Tile overlap so split rows still appear in full
HEADER_FRAC     = 0.10   # Top band stamped onto every non-top-row tile


def panelize(src: Path, out_dir: Path, prefix: str) -> list[Path]:
    """If `src` already fits within VISION_MAX_SIDE on both axes, return [src]
    unchanged so the model receives the full BOM at native resolution.
    Otherwise split into overlapping panels (each ≤ VISION_MAX_SIDE) and stamp
    the column-header band onto every non-top-row panel so column context is
    preserved. Returns the list of panel paths in row-major order.
    Guarantees: every row of the BOM appears, in full, in at least one panel.
    """
    from PIL import Image
    img = Image.open(str(src)).convert("RGB")
    W, H = img.size
    if W <= VISION_MAX_SIDE and H <= VISION_MAX_SIDE:
        return [src]

    cols = max(1, -(-W // VISION_MAX_SIDE))
    rows = max(1, -(-H // VISION_MAX_SIDE))
    base_tw, base_th = -(-W // cols), -(-H // rows)
    ov_w = int(base_tw * OVERLAP_FRAC)
    ov_h = int(base_th * OVERLAP_FRAC)
    header_h = max(1, int(H * HEADER_FRAC))
    header_band = img.crop((0, 0, W, header_h))

    out: list[Path] = []
    for r in range(rows):
        for c in range(cols):
            L = max(0, c * base_tw - ov_w)
            T = max(0, r * base_th - ov_h)
            R = min(W, (c + 1) * base_tw + ov_w)
            B = min(H, (r + 1) * base_th + ov_h)
            tile = img.crop((L, T, R, B))
            if r > 0:
                hdr = header_band.crop((L, 0, R, header_h))
                stamped = Image.new("RGB", (tile.size[0], hdr.size[1] + tile.size[1]),
                                    (255, 255, 255))
                stamped.paste(hdr, (0, 0))
                stamped.paste(tile, (0, hdr.size[1]))
                tile = stamped
            p = out_dir / f"{prefix}_p{r}{c}.png"
            tile.save(str(p), "PNG")
            out.append(p)
    return out


def trim_whitespace(src: Path, padding: int = 20) -> None:
    """Trim outer whitespace from a PNG in place. Removes the title-block
    margin around a BOM so the table itself fills the frame. No tiling."""
    from PIL import Image, ImageChops
    img = Image.open(str(src)).convert("RGB")
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if not bbox:
        return
    L, T, R, B = bbox
    L = max(0, L - padding)
    T = max(0, T - padding)
    R = min(img.size[0], R + padding)
    B = min(img.size[1], B + padding)
    img.crop((L, T, R, B)).save(str(src), "PNG")


def extract_rows(bom: Path, job_number: str) -> list[dict]:
    ext = bom.suffix.lower()
    if ext in EXCEL_EXTS:
        # FoxFab pattern: xlsx is an Excel container wrapping an embedded
        # AutoCAD DWG. The cells are empty; the BOM is the drawing.
        if xlsx_has_embedded_dwg(bom):
            out_dir = Path(__file__).resolve().parent.parent / "output" / job_number
            out_dir.mkdir(parents=True, exist_ok=True)
            png = out_dir / f"{job_number}_bom.png"
            if extract_emf_as_png(bom, png, dpi=300):
                trim_whitespace(png)
                panels = panelize(png, out_dir, f"{job_number}_bom")
                die("EMF_IMAGE", payload={"full": str(png),
                                           "panels": [str(p) for p in panels]})
            die("EMF_IMAGE_FAILED", str(bom))
        return read_excel(bom)
    if ext == ".csv":
        return read_csv(bom)
    if ext == ".pdf":
        # Render every page to PNG via pymupdf, then tile (auto-cropped).
        try:
            import fitz
        except ImportError:
            die("MISSING_DEP", "pymupdf not installed (pip install pymupdf)")
        out_dir = Path(__file__).resolve().parent.parent / "output" / job_number
        out_dir.mkdir(parents=True, exist_ok=True)
        full_pages: list[Path] = []
        all_panels: list[Path] = []
        with fitz.open(str(bom)) as doc:
            for i, page in enumerate(doc):
                pix = page.get_pixmap(dpi=300)
                page_png = out_dir / f"{job_number}_bom_pg{i+1}.png"
                pix.save(str(page_png))
                trim_whitespace(page_png)
                full_pages.append(page_png)
                all_panels.extend(panelize(page_png, out_dir, f"{job_number}_bom_pg{i+1}"))
        die("PDF_IMAGE", payload={"pages": [str(p) for p in full_pages],
                                   "panels": [str(p) for p in all_panels]})
    die("NO_BOM_FOUND", f"Unsupported BOM extension: {ext}")


# ── Main ───────────────────────────────────────────────────────────────
def run(job_number: str, bom_override: str | None = None,
        jobs_root: str | None = None):
    global JOBS_ROOT
    if jobs_root:
        JOBS_ROOT = Path(jobs_root)

    job_folder = find_job_folder(job_number)
    prf = find_prf(job_folder)

    if bom_override:
        bom = Path(bom_override)
        if not bom.is_file():
            die("NO_BOM_FOUND", f"Override BOM not found: {bom}")
    else:
        boms = list_boms(job_folder)
        bom = pick_bom(prf, boms)

    rows = extract_rows(bom, job_number)

    out = {
        "job_number": job_number,
        "job_folder": job_folder.name,
        "prf": prf.name if prf else None,
        "bom": {"name": bom.name, "path": str(bom)},
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "row_count": len(rows),
        "rows": rows,
    }
    sys.stdout.write(json.dumps(out, indent=2, ensure_ascii=False))
    sys.stdout.write("\n")


def main():
    p = argparse.ArgumentParser(description="Parts Needed — extract BOM rows for a job")
    p.add_argument("--job", required=True, help="Job number, e.g. J16204")
    p.add_argument("--bom", default=None, help="Force a specific BOM path (after CHOOSE_BOM)")
    p.add_argument("--jobs-root", default=None, help="Override jobs root (for testing)")
    args = p.parse_args()
    run(args.job, bom_override=args.bom, jobs_root=args.jobs_root)


if __name__ == "__main__":
    main()
