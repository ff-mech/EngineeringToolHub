"""
Model Copy — match BOM P/Ns against the MODEL LIBRARY and copy matching
SLDPRT/SLDASM files into a per-job folder. Learns user choices via a
persistent mappings.json and supports an ignore_list.json.

Standalone CLI:
    python modelcopy.py --job J16204 --pns-json rows.json --plan
    python modelcopy.py --job J16204 --pns-json rows.json --resolve choices.json
"""

import argparse
import difflib
import json
import re
import shutil
import sys
from pathlib import Path

# ── Configuration ──────────────────────────────────────────────────────
MODEL_LIBRARY_ROOT = Path(r"\\NPSVR05\FOXFAB_REDIRECT$\lbadong\Desktop\AGENT ARMY\MODEL LIBRARY")
MODEL_EXTS         = {".sldprt", ".sldasm"}
FUZZY_CUTOFF       = 0.55
FUZZY_TOP_N        = 5

# Description-based rules (applied before P/N matching)
DESC_IGNORE_PATTERNS = [
    re.compile(r"rail\s*mount(ed)?\s*breaker", re.I),
    re.compile(r"phase\s*sequence\s*relay", re.I),
    re.compile(r"fuse\s*holder", re.I),
    re.compile(r"supplied\s*by\s*customer", re.I),
]
# Description-based "write to txt" rules: (pattern, txt_filename)
# Parts matching these get listed in <filename>.txt in the job's Parts folder
# instead of having a 3D model copied.
DESC_TXT_RULES = [
    (re.compile(r"kirk\s*key", re.I),      "Kirk Key.txt"),
    (re.compile(r"bridging\s*bar", re.I),  "Bridging Bar.txt"),
]

# P/N-based auto-ignore patterns (applied before P/N matching, after desc rules)
PN_IGNORE_PATTERNS = [
    # ABB XT-series molded-case breakers — library has no exact matches
    re.compile(r"^XT[2-9][A-Z]", re.I),
    # ABB breaker accessories (aux contacts, shunts, lug kits, etc.)
    re.compile(r"^(KXTA|KXT2|ZE1)", re.I),
]

SKILL_ROOT    = Path(__file__).resolve().parent.parent
OUTPUT_ROOT   = SKILL_ROOT / "output"
MAPPINGS_FILE = SKILL_ROOT / "mappings.json"
IGNORE_FILE   = SKILL_ROOT / "ignore_list.json"


# ── Normalization ──────────────────────────────────────────────────────
_NORM_STRIP = re.compile(r"[\s\-_./\\]+")


def norm(s: str) -> str:
    return _NORM_STRIP.sub("", (s or "").upper())


# ── State ──────────────────────────────────────────────────────────────
def load_state():
    mappings = {}
    ignore = []
    if MAPPINGS_FILE.is_file():
        try:
            mappings = json.loads(MAPPINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            mappings = {}
    if IGNORE_FILE.is_file():
        try:
            ignore = json.loads(IGNORE_FILE.read_text(encoding="utf-8"))
        except Exception:
            ignore = []
    return mappings, set(ignore)


def save_state(mappings: dict, ignore: set):
    MAPPINGS_FILE.write_text(
        json.dumps(mappings, indent=2, ensure_ascii=False), encoding="utf-8")
    IGNORE_FILE.write_text(
        json.dumps(sorted(ignore), indent=2, ensure_ascii=False), encoding="utf-8")


# ── Index the model library ────────────────────────────────────────────
def build_index(root: Path) -> list[tuple[str, str, Path]]:
    """Return [(normalized_stem, original_filename, full_path), ...]
    restricted to MODEL_EXTS."""
    out: list[tuple[str, str, Path]] = []
    if not root.is_dir():
        return out
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in MODEL_EXTS:
            continue
        if p.name.startswith("~$"):
            continue  # Office lockfile
        out.append((norm(p.stem), p.name, p))
    # De-duplicate by full path
    seen = set()
    deduped = []
    for t in out:
        key = str(t[2]).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(t)
    return deduped


# ── Matching ───────────────────────────────────────────────────────────
def match_part(pn: str, index: list, mappings: dict, ignore: set):
    """Returns (status, data):
        status ∈ {'ignored', 'mapped', 'exact', 'ambiguous', 'none'}
        data: list[Path] for ignored/none it is [].
    """
    npn = norm(pn)
    if not npn:
        return "none", []

    if npn in ignore:
        return "ignored", []

    if npn in mappings:
        mapped = Path(mappings[npn])
        if mapped.is_file():
            return "mapped", [mapped]
        # stale — fall through and remove later when a new choice is made
        mappings.pop(npn, None)

    # Exact token/substring match on normalized stems
    exact_hits = [p for (nstem, _name, p) in index if npn in nstem]
    if len(exact_hits) == 1:
        return "exact", exact_hits
    if len(exact_hits) > 1:
        # too many — present as ambiguous, but cap at FUZZY_TOP_N
        return "ambiguous", exact_hits[:FUZZY_TOP_N]

    # Fuzzy match — use difflib against normalized stems
    stems = [t[0] for t in index]
    close = difflib.get_close_matches(npn, stems, n=FUZZY_TOP_N, cutoff=FUZZY_CUTOFF)
    if not close:
        return "none", []
    # map back to Paths (first index per stem)
    by_stem: dict[str, Path] = {}
    for nstem, _name, p in index:
        by_stem.setdefault(nstem, p)
    return "ambiguous", [by_stem[s] for s in close if s in by_stem]


# ── Copying ────────────────────────────────────────────────────────────
def copy_files(paths: list[Path], dest: Path) -> list[dict]:
    dest.mkdir(parents=True, exist_ok=True)
    results = []
    for src in paths:
        target = dest / src.name
        if target.exists():
            results.append({"src": str(src), "dst": str(target), "state": "already_present"})
            continue
        try:
            shutil.copy2(src, target)
            results.append({"src": str(src), "dst": str(target), "state": "copied"})
        except Exception as e:
            results.append({"src": str(src), "dst": str(target), "state": f"error: {e}"})
    return results


# ── Description rules ─────────────────────────────────────────────────
def check_desc_rule(desc: str):
    """Return ('desc_ignored', None), ('txt', filename), or (None, None)."""
    if not desc:
        return None, None
    for pat in DESC_IGNORE_PATTERNS:
        if pat.search(desc):
            return "desc_ignored", None
    for pat, fname in DESC_TXT_RULES:
        if pat.search(desc):
            return "txt", fname
    return None, None


def check_pn_rule(pn: str) -> bool:
    """Return True if P/N matches any auto-ignore regex."""
    if not pn:
        return False
    for pat in PN_IGNORE_PATTERNS:
        if pat.search(pn):
            return True
    return False


# ── Planning + resolving ───────────────────────────────────────────────
def plan(job_number: str, rows: list[dict]) -> dict:
    mappings, ignore = load_state()
    index = build_index(MODEL_LIBRARY_ROOT)
    report = []
    for row in rows:
        pn = row.get("pn", "")
        desc = row.get("desc", "")
        rule, txt_name = check_desc_rule(desc)
        if rule == "desc_ignored":
            report.append({"pn": pn, "desc": desc, "npn": norm(pn),
                           "status": "desc_ignored", "candidates": []})
            continue
        if rule == "txt":
            report.append({"pn": pn, "desc": desc, "npn": norm(pn),
                           "status": f"txt:{txt_name}", "candidates": []})
            continue
        if check_pn_rule(pn):
            report.append({"pn": pn, "desc": desc, "npn": norm(pn),
                           "status": "pn_pattern_ignored", "candidates": []})
            continue
        status, paths = match_part(pn, index, mappings, ignore)
        report.append({
            "pn": pn,
            "desc": desc,
            "npn": norm(pn),
            "status": status,
            "candidates": [str(p) for p in paths],
        })
    return {
        "job": job_number,
        "model_library": str(MODEL_LIBRARY_ROOT),
        "index_size": len(index),
        "report": report,
    }


def resolve(job_number: str, rows: list[dict], choices: dict) -> dict:
    """`choices` maps normalized P/N → absolute path | "IGNORE" | "SKIP".
    Auto-resolvable (exact/mapped) items don't need to be in `choices`.
    """
    mappings, ignore = load_state()
    index = build_index(MODEL_LIBRARY_ROOT)

    dest = OUTPUT_ROOT / f"{job_number} - Parts"
    summary = []
    txt_buckets: dict[str, list] = {}  # filename -> list of {pn, desc}

    for row in rows:
        pn = row.get("pn", "")
        desc = row.get("desc", "")
        npn = norm(pn)

        rule, txt_name = check_desc_rule(desc)
        if rule == "desc_ignored":
            summary.append({"pn": pn, "desc": desc, "status": "desc_ignored", "files": []})
            continue
        if rule == "txt":
            txt_buckets.setdefault(txt_name, []).append({"pn": pn, "desc": desc})
            summary.append({"pn": pn, "desc": desc,
                            "status": f"txt:{txt_name}", "files": []})
            continue
        if check_pn_rule(pn):
            summary.append({"pn": pn, "desc": desc,
                            "status": "pn_pattern_ignored", "files": []})
            continue

        status, paths = match_part(pn, index, mappings, ignore)

        if status == "ignored":
            summary.append({"pn": pn, "status": "ignored", "files": []})
            continue

        if status in ("mapped", "exact"):
            results = copy_files(paths, dest)
            summary.append({"pn": pn, "status": status, "files": results})
            continue

        # ambiguous / none → look for a user choice
        choice = choices.get(npn) or choices.get(pn)
        if choice is None:
            summary.append({"pn": pn, "status": f"unresolved_{status}", "files": []})
            continue

        if choice == "IGNORE":
            ignore.add(npn)
            summary.append({"pn": pn, "status": "newly_ignored", "files": []})
            continue

        if choice == "SKIP":
            summary.append({"pn": pn, "status": "skipped_this_job", "files": []})
            continue

        # concrete path chosen
        chosen = Path(choice)
        if not chosen.is_file():
            summary.append({"pn": pn, "status": "choice_missing",
                            "files": [], "choice": str(chosen)})
            continue

        mappings[npn] = str(chosen)
        results = copy_files([chosen], dest)
        summary.append({"pn": pn, "status": "chosen", "files": results})

    save_state(mappings, ignore)

    # Write one .txt per DESC_TXT_RULES bucket that had hits
    for fname, entries in txt_buckets.items():
        dest.mkdir(parents=True, exist_ok=True)
        label = fname.rsplit(".", 1)[0]
        lines = [f"{label} parts — handled manually, no 3D model copied",
                 "=" * 60, ""]
        for e in entries:
            lines.append(f"{e['pn']}  —  {e['desc']}")
        (dest / fname).write_text("\n".join(lines) + "\n", encoding="utf-8")

    return {
        "job": job_number,
        "dest": str(dest),
        "summary": summary,
    }


# ── CLI ────────────────────────────────────────────────────────────────
def _load_rows(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise SystemExit("pns-json must be a JSON list")
    out = []
    for x in data:
        if isinstance(x, dict):
            if x.get("pn"):
                out.append({"pn": x.get("pn", ""), "desc": x.get("desc", "")})
        elif x:
            out.append({"pn": str(x), "desc": ""})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--job", required=True)
    ap.add_argument("--pns-json", required=True, type=Path,
                    help="JSON file: list of rows or list of P/N strings")
    ap.add_argument("--plan", action="store_true", help="Dry-run report only")
    ap.add_argument("--resolve", type=Path, default=None,
                    help="JSON file: {normalized_pn: path|IGNORE|SKIP}")
    args = ap.parse_args()

    rows = _load_rows(args.pns_json)

    if args.plan:
        sys.stdout.write(json.dumps(plan(args.job, rows), indent=2, ensure_ascii=False))
        sys.stdout.write("\n")
        return

    choices = {}
    if args.resolve and args.resolve.is_file():
        choices = json.loads(args.resolve.read_text(encoding="utf-8"))
    sys.stdout.write(json.dumps(resolve(args.job, rows, choices),
                                indent=2, ensure_ascii=False))
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
