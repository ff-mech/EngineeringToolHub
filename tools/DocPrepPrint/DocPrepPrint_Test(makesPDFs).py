import os
import re
import sys
import time
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_TITLE = "Doc Prep Print - SIMULATION MODE"
PREFERRED_PRINTER = "FoxFab (Konica Bizhub C360i) on NPSVR05"
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
PDF_EXTENSIONS = {".pdf", ".PDF"}

try:
    import win32print
    import win32com.client
except Exception:
    win32print = None
    win32com = None

try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    PdfReader = None
    PdfWriter = None

LOG_LINES = []
LOG_FILE = None
RUN_HAD_ERRORS = False

UI_BG = "#F6F8FB"
UI_PANEL = "#FFFFFF"
UI_BORDER = "#D8DEE9"
UI_ACCENT = "#1F6FEB"
UI_TEXT = "#1F2937"
UI_SUBTLE = "#6B7280"


def configure_styles(root):
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background=UI_BG)
        style.configure("Card.TFrame", background=UI_PANEL, relief="solid", borderwidth=1)
        style.configure("Header.TLabel", background=UI_BG, foreground=UI_TEXT, font=("Segoe UI", 15, "bold"))
        style.configure("Subtle.TLabel", background=UI_BG, foreground=UI_SUBTLE, font=("Segoe UI", 10))
        style.configure("Body.TLabel", background=UI_PANEL, foreground=UI_TEXT, font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Secondary.TButton", font=("Segoe UI", 10))
    except Exception:
        pass



def log(msg: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LOG_LINES.append(f"[{stamp}] {msg}")


def safe_name(text: str) -> str:
    if not text:
        return "UnknownJob"
    return re.sub(r'[\\/:*?"<>|]+', "_", text).strip() or "UnknownJob"


def init_log(job_folder: str):
    global LOG_FILE
    script_dir = Path(__file__).resolve().parent
    job_name = safe_name(Path(job_folder).name)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    LOG_FILE = script_dir / f"DocPrepPrint_{job_name}_{stamp}.log"
    log(f"Log initialized: {LOG_FILE}")



SIMULATION_OUTPUT_DIR = None
SELECTED_CONTEXT_CACHE = {}


def init_simulation_output_dir(job_folder: str):
    global SIMULATION_OUTPUT_DIR
    script_dir = Path(__file__).resolve().parent
    job_name = safe_name(Path(job_folder).name)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    SIMULATION_OUTPUT_DIR = script_dir / f"Simulated_Print_Output_{job_name}_{stamp}"
    SIMULATION_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Simulation output dir initialized: {SIMULATION_OUTPUT_DIR}")


def get_sim_output_path(order_num: int, label: str):
    if SIMULATION_OUTPUT_DIR is None:
        raise RuntimeError("Simulation output directory has not been initialized.")
    safe_label = safe_name(label)
    return SIMULATION_OUTPUT_DIR / f"{order_num:02d}_{safe_label}.pdf"


def copy_pdf_to_output(src_pdf: Path, dest_pdf: Path):
    import shutil
    shutil.copy2(src_pdf, dest_pdf)
    log(f"Saved simulated PDF output: {dest_pdf} (copied from {src_pdf})")


def export_excel_to_pdf(file_path: Path, dest_pdf: Path, first_sheet_only: bool, excel_app=None):
    ensure_com()
    log(f"Simulating Excel output to PDF: {file_path} -> {dest_pdf} | first_sheet_only={first_sheet_only}")
    created_local_excel = excel_app is None
    excel = excel_app
    wb = None
    xlTypePDF = 0
    try:
        if excel is None:
            excel = get_excel_app()
        wb = excel.Workbooks.Open(str(file_path))
        if first_sheet_only:
            ws = wb.Worksheets(1)
            if not ws.Visible:
                raise RuntimeError(f"First sheet is hidden in {file_path.name}.")
            ws.ExportAsFixedFormat(xlTypePDF, str(dest_pdf))
        else:
            active = excel.ActiveWindow.ActiveSheet if getattr(excel, "ActiveWindow", None) else wb.ActiveSheet
            if not active.Visible:
                raise RuntimeError(f"Active sheet is hidden in {file_path.name}.")
            page_setup = active.PageSetup
            page_setup.Orientation = 2
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = False
            active.ExportAsFixedFormat(xlTypePDF, str(dest_pdf))
        log(f"Saved simulated Excel PDF output: {dest_pdf}")
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception as e:
            log(f"Excel workbook close warning for simulated export {file_path}: {e}")
        if created_local_excel and excel is not None:
            try:
                excel.Quit()
            except Exception as e:
                log(f"Excel quit warning for simulated export {file_path}: {e}")


def save_log():
    if not LOG_FILE:
        return
    try:
        LOG_FILE.write_text("\n".join(LOG_LINES) + "\n", encoding="utf-8")
    except Exception:
        pass


def save_crash_log(exc: BaseException):
    global RUN_HAD_ERRORS
    RUN_HAD_ERRORS = True
    script_dir = Path(__file__).resolve().parent
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    crash_path = script_dir / f"DocPrepPrint_CRASH_{stamp}.log"
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    content = f"Unhandled exception at {datetime.now()}\n\n{tb}\n"
    try:
        crash_path.write_text(content, encoding="utf-8")
    except Exception:
        pass
    log("UNHANDLED EXCEPTION:")
    for line in tb.splitlines():
        log(line)
    save_log()


def info(msg: str):
    messagebox.showinfo(APP_TITLE, msg)


def warn(msg: str):
    messagebox.showwarning(APP_TITLE, msg)


def ask_yes_no(msg: str) -> bool:
    return messagebox.askyesno(APP_TITLE, msg)


def ask_retry_cancel(msg: str) -> str:
    return "continue" if messagebox.askyesno(APP_TITLE, msg + "\n\nYes = Continue\nNo = Cancel") else "cancel"


def choose_folder():
    return filedialog.askdirectory(title="Select the main job folder or a specific numbered mech variant", mustexist=True)



def choose_from_list(title: str, prompt: str, items):
    items = list(items)
    if not items:
        return None

    top = tk.Toplevel()
    top.title(title)
    top.geometry("760x560")
    top.minsize(680, 480)
    top.configure(bg=UI_BG)
    top.grab_set()

    result = {"value": None}

    outer = ttk.Frame(top, style="App.TFrame", padding=18)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=title, style="Header.TLabel").pack(anchor="w")
    ttk.Label(outer, text=prompt, style="Subtle.TLabel").pack(anchor="w", pady=(4, 14))

    card = ttk.Frame(outer, style="Card.TFrame", padding=12)
    card.pack(fill="both", expand=True)

    list_frame = ttk.Frame(card, style="Card.TFrame")
    list_frame.pack(fill="both", expand=True)

    scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(
        list_frame,
        yscrollcommand=scrollbar.set,
        font=("Segoe UI", 10),
        activestyle="none",
        borderwidth=0,
        highlightthickness=0,
        selectmode="browse",
    )
    for item in items:
        listbox.insert("end", str(item))
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    button_bar = ttk.Frame(outer, style="App.TFrame")
    button_bar.pack(fill="x", pady=(14, 0))

    def on_ok():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning(title, "Please select an item.")
            return
        result["value"] = items[sel[0]]
        top.destroy()

    def on_cancel():
        top.destroy()

    ttk.Button(button_bar, text="OK", command=on_ok, style="Primary.TButton").pack(side="left")
    ttk.Button(button_bar, text="Cancel", command=on_cancel, style="Secondary.TButton").pack(side="left", padx=(8, 0))

    top.wait_window()
    return result["value"]


def choose_mech_variant(variant_paths):
    variant_paths = list(variant_paths)
    if not variant_paths:
        return None
    names = [Path(p).name for p in variant_paths]
    selected_name = choose_from_list(
        "Select Mechanical Variant",
        "Multiple numbered mechanical variants were found. Choose which one to use for this run:",
        names,
    )
    if not selected_name:
        raise RuntimeError("Mechanical variant selection cancelled.")
    for p in variant_paths:
        if Path(p).name == selected_name:
            log(f"Selected mechanical variant: {p}")
            return p
    raise RuntimeError("Selected mechanical variant was not found.")


def summary_dialog(summary_text: str):
    top = tk.Toplevel()
    top.title("Pre-Print Summary")
    top.geometry("960x680")
    top.minsize(820, 560)
    top.configure(bg=UI_BG)
    top.grab_set()

    result = {"value": "cancel"}

    outer = ttk.Frame(top, style="App.TFrame", padding=18)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text="Pre-Print Summary", style="Header.TLabel").pack(anchor="w")
    ttk.Label(
        outer,
        text="Review the print packet, then continue or adjust the folder/printer.",
        style="Subtle.TLabel"
    ).pack(anchor="w", pady=(4, 14))

    card = ttk.Frame(outer, style="Card.TFrame", padding=12)
    card.pack(fill="both", expand=True)

    text_frame = ttk.Frame(card, style="Card.TFrame")
    text_frame.pack(fill="both", expand=True)

    scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")

    text_widget = tk.Text(
        text_frame,
        wrap="word",
        yscrollcommand=scrollbar.set,
        font=("Consolas", 10),
        borderwidth=0,
        highlightthickness=0,
        background=UI_PANEL,
        foreground=UI_TEXT,
        padx=8,
        pady=8,
    )
    text_widget.insert("1.0", summary_text)
    text_widget.configure(state="disabled")
    text_widget.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=text_widget.yview)

    button_bar = ttk.Frame(outer, style="App.TFrame")
    button_bar.pack(fill="x", pady=(14, 0))

    def choose(value):
        result["value"] = value
        top.destroy()

    ttk.Button(button_bar, text="Print", command=lambda: choose("print"), style="Primary.TButton").pack(side="left")
    ttk.Button(button_bar, text="Change Printer", command=lambda: choose("printer"), style="Secondary.TButton").pack(side="left", padx=(8, 0))
    ttk.Button(button_bar, text="Change Folder", command=lambda: choose("folder"), style="Secondary.TButton").pack(side="left", padx=(8, 0))
    ttk.Button(button_bar, text="Cancel", command=lambda: choose("cancel"), style="Secondary.TButton").pack(side="right")

    top.wait_window()
    return result["value"]


def get_installed_printers():
    if not win32print:
        raise RuntimeError("pywin32 is required: win32print not available.")
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = sorted({p[2] for p in win32print.EnumPrinters(flags)})
    log(f"Detected printers: {printers}")
    return printers


def choose_printer():
    printers = get_installed_printers()
    preferred_present = PREFERRED_PRINTER in printers
    if preferred_present and ask_yes_no(f"Use this printer?\n\n{PREFERRED_PRINTER}"):
        log(f"Selected preferred printer: {PREFERRED_PRINTER}")
        return PREFERRED_PRINTER
    selected = choose_from_list("Select Printer", "Choose the printer for this run:", printers)
    if not selected:
        raise RuntimeError("Printer selection cancelled.")
    log(f"Selected printer: {selected}")
    return selected


def set_default_printer(printer_name: str):
    if not win32print:
        raise RuntimeError("pywin32 is required: win32print not available.")
    win32print.SetDefaultPrinter(printer_name)
    log(f"Set default printer: {printer_name}")



def get_queue_jobs(printer_name: str):
    if not win32print:
        return []
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        info = win32print.GetPrinter(hprinter, 2)
        total_jobs = info.get("cJobs", 0)
        if total_jobs <= 0:
            return []
        jobs = win32print.EnumJobs(hprinter, 0, total_jobs, 1)
        normalized = []
        for job in jobs:
            normalized.append({
                "JobId": job.get("JobId"),
                "pDocument": str(job.get("pDocument") or ""),
                "Status": job.get("Status"),
            })
        return normalized
    finally:
        win32print.ClosePrinter(hprinter)


def queue_snapshot(printer_name: str):
    jobs = get_queue_jobs(printer_name)
    return {str(job["JobId"]): job.get("pDocument", "") for job in jobs}


def wait_for_new_queue_job(printer_name: str, before_snapshot, expected_name=None, timeout=45, poll_interval=0.5):
    start = time.time()
    expected_lc = (expected_name or "").lower()
    while time.time() - start < timeout:
        current = queue_snapshot(printer_name)
        new_ids = [job_id for job_id in current if job_id not in before_snapshot]
        if new_ids:
            log(f"Queue change detected on {printer_name}: new job ids {new_ids}")
            return current
        if expected_lc:
            for name in current.values():
                if expected_lc and expected_lc in name.lower():
                    log(f"Queue document match detected on {printer_name}: {name}")
                    return current
        time.sleep(poll_interval)
    raise RuntimeError(f"Timed out waiting for print job to appear in queue for {printer_name}.")


def wait_for_queue_settle(printer_name: str, settle_seconds=2.5, timeout=20, poll_interval=0.5):
    start = time.time()
    stable_since = None
    previous = tuple(sorted(queue_snapshot(printer_name).items()))
    while time.time() - start < timeout:
        time.sleep(poll_interval)
        current = tuple(sorted(queue_snapshot(printer_name).items()))
        if current == previous:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= settle_seconds:
                log(f"Queue settled for {printer_name}.")
                return
        else:
            previous = current
            stable_since = None
    log(f"Queue settle timeout reached for {printer_name}; continuing.")


def wait_for_spool_sequence(printer_name: str, before_snapshot, expected_name=None):
    try:
        current = wait_for_new_queue_job(printer_name, before_snapshot, expected_name=expected_name, timeout=20, poll_interval=0.35)
        time.sleep(0.6)
        return current
    except Exception as e:
        log(f"Queue wait warning for {printer_name}: {e}")
        time.sleep(1.5)


def wait_for_section_boundary(printer_name: str):
    try:
        wait_for_queue_settle(printer_name, stable_seconds=1.5, timeout=20, poll_interval=0.5)
    except Exception as e:
        log(f"Section boundary queue wait warning for {printer_name}: {e}")
        time.sleep(2)


def is_mech_variant_folder(path: Path):
    return path.is_dir() and re.search(r"-\d{2}$", path.name) is not None


def get_selected_context(job_folder: Path):
    cache_key = str(Path(job_folder).resolve())
    if cache_key in SELECTED_CONTEXT_CACHE:
        return SELECTED_CONTEXT_CACHE[cache_key]

    # Case 1: user selected a numbered mechanical variant directly, e.g. ...\200 Mech\J15302-01
    if is_mech_variant_folder(job_folder):
        parent = job_folder.parent
        mech_parent_ok = parent.name == "200 Mech"
        required = [
            job_folder / "204 BOM",
            job_folder / "205 CNC",
            job_folder / "202 PDFs_Flats",
            job_folder / "203 Assemblies",
        ]
        if mech_parent_ok and all(p.is_dir() for p in required):
            base_job = job_folder.parent.parent
            log(f"Using selected mechanical variant only: {job_folder}")
            ctx = {
                "job_root": base_job,
                "mech_roots": [job_folder],
                "selected_variant_only": True,
            }
            SELECTED_CONTEXT_CACHE[cache_key] = ctx
            return ctx

    # Case 2: user selected the normal top-level job folder with direct 200 Mech structure
    mech_root = job_folder / "200 Mech"
    if not mech_root.is_dir():
        raise RuntimeError("Missing required folder: 200 Mech")

    direct_required = [
        mech_root / "204 BOM",
        mech_root / "205 CNC",
        mech_root / "202 PDFs_Flats",
        mech_root / "203 Assemblies",
    ]
    if all(p.is_dir() for p in direct_required):
        log(f"Using direct 200 Mech structure: {mech_root}")
        ctx = {
            "job_root": job_folder,
            "mech_roots": [mech_root],
            "selected_variant_only": False,
        }
        SELECTED_CONTEXT_CACHE[cache_key] = ctx
        return ctx

    # Case 3: user selected the main job folder and 200 Mech contains numbered variants
    variant_roots = []
    for child in sorted(mech_root.iterdir(), key=lambda p: p.name.lower()):
        if not child.is_dir():
            continue
        if re.search(r"-\d{2}$", child.name):
            req = [
                child / "204 BOM",
                child / "205 CNC",
                child / "202 PDFs_Flats",
                child / "203 Assemblies",
            ]
            if all(p.is_dir() for p in req):
                variant_roots.append(child)

    if variant_roots:
        selected_variant = choose_mech_variant(variant_roots)
        log(f"Using numbered 200 Mech variant from main job folder: {selected_variant}")
        ctx = {
            "job_root": job_folder,
            "mech_roots": [selected_variant],
            "selected_variant_only": True,
        }
        SELECTED_CONTEXT_CACHE[cache_key] = ctx
        return ctx

    raise RuntimeError(
        "Could not find a usable mechanical folder structure. Expected either "
        "'200 Mech\\204 BOM' etc. directly, or numbered folders inside 200 Mech like '*-01', '*-02' "
        "that contain 202/203/204/205 subfolders."
    )


def validate_required_folders(job_folder: Path):
    missing = []

    try:
        ctx = get_selected_context(job_folder)
    except Exception as e:
        log(str(e))
        return [str(e)]

    base_job = Path(ctx["job_root"])

    required = [
        base_job / "300 Inputs",
        base_job / "300 Inputs" / "302 Production Release Form",
        base_job / "100 Elec" / "102 Drawings",
    ]
    for full in required:
        if not full.is_dir():
            missing.append(str(full.relative_to(base_job)))
            log(f"Missing required folder: {full}")

    for mech_root in ctx["mech_roots"]:
        req = [
            mech_root / "204 BOM",
            mech_root / "205 CNC",
            mech_root / "202 PDFs_Flats",
            mech_root / "203 Assemblies",
        ]
        for full in req:
            if not full.is_dir():
                missing.append(str(full))
                log(f"Missing required folder: {full}")

    return missing


def list_direct_files(folder: Path):
    return sorted([p for p in folder.iterdir() if p.is_file()], key=lambda p: p.name.lower())


def choose_one_file(files, title, prompt):
    names = [p.name for p in sorted(files, key=lambda p: p.name.lower())]
    selected_name = choose_from_list(title, prompt, names)
    if not selected_name:
        raise RuntimeError(f"{title} selection cancelled.")
    selected = next(p for p in files if p.name == selected_name)
    log(f"Selected from multiple matches in {title}: {selected}")
    return selected


def match_fwo(folder: Path):
    files = list_direct_files(folder)
    matches = []
    skipped = []
    for f in files:
        if f.suffix not in PDF_EXTENSIONS:
            skipped.append(f)
            continue
        if f.stem == "Fabrication Work Order - Standard v1.0":
            matches.append(f)
        else:
            skipped.append(f)
    for f in skipped:
        log(f"Skipped FWO candidate/non-candidate: {f}")
    if len(matches) != 1:
        raise RuntimeError("Fabrication Work Order file was not found exactly once.")
    return matches[0]


def match_contains_excel(folder: Path, token: str, title: str):
    files = list_direct_files(folder)
    matches = []
    for f in files:
        if f.suffix.lower() not in {e.lower() for e in EXCEL_EXTENSIONS}:
            log(f"Skipped non-Excel in {title}: {f}")
            continue
        if token.lower() in f.name.lower():
            matches.append(f)
        else:
            log(f"Skipped Excel not matching {token} in {title}: {f}")
    if not matches:
        raise RuntimeError(f"No Excel file containing '{token}' found in {title}.")
    if len(matches) > 1:
        return choose_one_file(matches, title, f"Multiple matches found for {title}. Choose one:")
    return matches[0]


def match_pack_pdf(folder: Path):
    files = list_direct_files(folder)
    matches = []
    for f in files:
        if f.suffix not in PDF_EXTENSIONS:
            log(f"Skipped non-PDF in Electrical Pack folder: {f}")
            continue
        if "PACK" in f.name.upper():
            matches.append(f)
        else:
            log(f"Skipped non-PACK PDF in Electrical Pack folder: {f}")
    if not matches:
        raise RuntimeError("No PDF containing 'PACK' found in Electrical Drawings.")
    if len(matches) > 1:
        return choose_one_file(matches, "Electrical Drawing Pack", "Multiple PACK PDFs found. Choose one:")
    return matches[0]


def classify_cnc(pdf_path: Path) -> str:
    base = pdf_path.stem
    if re.match(r"^[Jj]", base):
        return "duplex"
    if re.match(r"^\d{3}-", base):
        return "duplex"
    return "simplex"


def match_cnc(folder: Path):
    pdfs = []
    for f in list_direct_files(folder):
        if f.suffix in PDF_EXTENSIONS:
            pdfs.append(f)
        else:
            log(f"Skipped non-PDF in CNC: {f}")
    if not pdfs:
        raise RuntimeError("No PDFs found in CNC.")
    pdfs = sorted(pdfs, key=lambda p: p.name.lower())
    return pdfs


def match_flats(folder: Path):
    pdfs = []
    for f in list_direct_files(folder):
        if f.suffix in PDF_EXTENSIONS:
            pdfs.append(f)
        else:
            log(f"Skipped non-PDF in PDFs_Flats: {f}")
    if not pdfs:
        raise RuntimeError("No PDFs found in PDFs_Flats.")
    return sorted(pdfs, key=lambda p: p.name.lower())


def match_assemblies(folder: Path):
    pdfs = []
    excluded_lay = []
    for f in list_direct_files(folder):
        if f.suffix not in PDF_EXTENSIONS:
            log(f"Skipped non-PDF in Assemblies: {f}")
            continue
        if f.stem.endswith("-LAY"):
            excluded_lay.append(f)
            log(f"Skipped -LAY in Assemblies: {f}")
        else:
            pdfs.append(f)
    if not pdfs:
        raise RuntimeError("No printable PDFs found in Assemblies.")
    return sorted(pdfs, key=lambda p: p.name.lower()), excluded_lay


def build_plan(job_folder: Path):
    plan = {}
    warnings = []

    missing = validate_required_folders(job_folder)
    if missing:
        raise RuntimeError("Missing required folders:\n" + "\n".join(missing))

    ctx = get_selected_context(job_folder)
    base_job = Path(ctx["job_root"])
    mech_roots = ctx["mech_roots"]

    plan["job_folder"] = str(job_folder)
    plan["base_job_folder"] = str(base_job)
    plan["selected_variant_only"] = ctx["selected_variant_only"]
    plan["mech_roots"] = [str(p) for p in mech_roots]

    plan["fwo"] = match_fwo(base_job / "300 Inputs")
    plan["bom"] = match_contains_excel(mech_roots[0] / "204 BOM", "BOM", "BOM")

    cnc = []
    for mech_root in mech_roots:
        cnc.extend(match_cnc(mech_root / "205 CNC"))
    plan["cnc"] = sorted(cnc, key=lambda p: p.name.lower())

    flats = []
    for mech_root in mech_roots:
        flats.extend(match_flats(mech_root / "202 PDFs_Flats"))
    plan["flats"] = sorted(flats, key=lambda p: p.name.lower())

    plan["prf"] = match_contains_excel(base_job / "300 Inputs" / "302 Production Release Form", "PRF", "Production Release Form")
    plan["pack"] = match_pack_pdf(base_job / "100 Elec" / "102 Drawings")

    assemblies = []
    excluded_lay = []
    for mech_root in mech_roots:
        a, ex = match_assemblies(mech_root / "203 Assemblies")
        assemblies.extend(a)
        excluded_lay.extend(ex)
    plan["assemblies"] = sorted(assemblies, key=lambda p: p.name.lower())
    plan["assemblies_excluded_lay"] = sorted(excluded_lay, key=lambda p: p.name.lower())

    if ctx["selected_variant_only"]:
        warnings.append(f"Variant-only run: printing only from {Path(mech_roots[0]).name}")

    plan["warnings"] = warnings
    return plan


def make_summary(plan, printer_name: str):
    cnc_duplex = sum(1 for p in plan["cnc"] if classify_cnc(p) == "duplex")
    cnc_simplex = len(plan["cnc"]) - cnc_duplex

    lines = [
        f"Simulation Target Name: {printer_name}",
        f"Job Folder: {plan['job_folder']}",
        "",
        f"Fabrication Work Order: {plan['fwo'].name}",
        f"BOM: {plan['bom'].name}",
        f"CNC: {len(plan['cnc'])} PDFs total ({cnc_duplex} duplex individual, {cnc_simplex} simplex merged)",
        f"PDFs_Flats: {len(plan['flats'])} PDFs (combined into 1 print job)",
        f"Production Release Form: {plan['prf'].name}",
        f"Electrical Drawing Pack: {plan['pack'].name} (pages 1-2)",
        f"Assemblies: {len(plan['assemblies'])} PDFs (combined into 1 print job, excluded {len(plan['assemblies_excluded_lay'])} -LAY files)",
    ]
    if plan["warnings"]:
        lines += ["", "WARNINGS:"]
        lines += [f"- {w}" for w in plan["warnings"]]
    return "\n".join(lines)


def ensure_com():
    if win32com is None:
        raise RuntimeError("pywin32 is required: win32com.client not available.")


def get_excel_app():
    ensure_com()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def print_excel_active_sheet(file_path: Path, printer_name: str, mode: str, first_sheet_only: bool, excel_app=None):
    ensure_com()
    log(f"Attempting Excel print: {file_path} | mode={mode} | first_sheet_only={first_sheet_only}")
    created_local_excel = excel_app is None
    excel = excel_app
    wb = None
    try:
        if excel is None:
            excel = get_excel_app()
        wb = excel.Workbooks.Open(str(file_path))
        if first_sheet_only:
            ws = wb.Worksheets(1)
            if not ws.Visible:
                raise RuntimeError(f"First sheet is hidden in {file_path.name}.")
            ws.PrintOut(ActivePrinter=printer_name)
        else:
            active = excel.ActiveWindow.ActiveSheet if getattr(excel, "ActiveWindow", None) else wb.ActiveSheet
            if not active.Visible:
                raise RuntimeError(f"Active sheet is hidden in {file_path.name}.")
            page_setup = active.PageSetup
            page_setup.Orientation = 2
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = False
            active.PrintOut(ActivePrinter=printer_name)
        log(f"Excel print queued successfully: {file_path}")
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception as e:
            log(f"Excel workbook close warning for {file_path}: {e}")
        if created_local_excel and excel is not None:
            try:
                excel.Quit()
            except Exception as e:
                log(f"Excel quit warning for {file_path}: {e}")


def _set_duplex_flag(printer_name: str, duplex_mode: str):
    if not win32print:
        raise RuntimeError("pywin32 is required: win32print not available.")
    # 1 = simplex, 2 = vertical duplex
    desired = 1 if duplex_mode == "simplex" else 2
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        props = win32print.GetPrinter(hprinter, 2)
        devmode = props["pDevMode"]
        if hasattr(devmode, "Duplex"):
            devmode.Duplex = desired
            props["pDevMode"] = devmode
            win32print.SetPrinter(hprinter, 2, props, 0)
            log(f"Set printer duplex flag for {printer_name} to {duplex_mode}")
        else:
            log(f"Printer devmode does not expose Duplex for {printer_name}")
    finally:
        win32print.ClosePrinter(hprinter)




def merge_pdfs_to_temp(file_paths, suffix_label: str):
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError(
            "Combining PDFs requires pypdf. Install it with: py -m pip install --user pypdf"
        )

    files = list(file_paths)
    if not files:
        raise RuntimeError(f"No PDF files supplied for merge: {suffix_label}")

    import tempfile

    writer = PdfWriter()
    total_pages = 0

    for pdf in files:
        try:
            reader = PdfReader(str(pdf))
            page_count = len(reader.pages)
            for page in reader.pages:
                writer.add_page(page)
            total_pages += page_count
            log(f"Added to merged PDF [{suffix_label}]: {pdf} ({page_count} pages)")
        except Exception as e:
            raise RuntimeError(f"Could not merge PDF {pdf.name} into {suffix_label}: {e}") from e

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{suffix_label}.pdf")
    tmp_path = Path(tmp.name)
    tmp.close()
    with open(tmp_path, "wb") as f:
        writer.write(f)

    log(f"Created merged PDF [{suffix_label}]: {tmp_path} | source_files={len(files)} | total_pages={total_pages}")
    return tmp_path


def save_pdf_simulation(file_path: Path, order_num: int, label: str, mode: str, pages=None):
    import tempfile

    print_target = file_path
    temp_to_cleanup = None
    dest_pdf = get_sim_output_path(order_num, label)

    log(f"Simulating PDF output: {file_path} | mode={mode} | pages={pages} | dest={dest_pdf}")

    if pages is not None:
        if PdfReader is None or PdfWriter is None:
            raise RuntimeError(
                "PACK pages 1-2 simulation requires pypdf. Install it with: py -m pip install --user pypdf"
            )
        try:
            reader = PdfReader(str(file_path))
            total_pages = len(reader.pages)
            start, end = pages
            if total_pages < (end + 1):
                raise RuntimeError(
                    f"{file_path.name} does not have enough pages. Needed pages {start + 1}-{end + 1}, found {total_pages}."
                )
            writer = PdfWriter()
            for idx in range(start, end + 1):
                writer.add_page(reader.pages[idx])
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix="_PACK_pages_1_2.pdf")
            tmp_path = Path(tmp.name)
            tmp.close()
            with open(tmp_path, "wb") as f:
                writer.write(f)
            temp_to_cleanup = tmp_path
            print_target = tmp_path
            log(f"Created temporary PDF for page-range simulation: {print_target}")
        except Exception as e:
            raise RuntimeError(f"Could not create temporary page-range PDF for {file_path.name}: {e}") from e

    try:
        copy_pdf_to_output(print_target, dest_pdf)
    finally:
        if temp_to_cleanup is not None:
            try:
                if temp_to_cleanup.exists():
                    temp_to_cleanup.unlink()
                    log(f"Deleted temporary page-range PDF: {temp_to_cleanup}")
            except Exception as e:
                log(f"Temporary page-range PDF cleanup warning for {temp_to_cleanup}: {e}")


def handle_print_error(title: str, err: Exception):
    global RUN_HAD_ERRORS
    RUN_HAD_ERRORS = True
    tb = "".join(traceback.format_exception(type(err), err, err.__traceback__))
    log(f"ERROR during {title}: {err}")
    for line in tb.splitlines():
        log(line)
    return ask_retry_cancel(f"{title} failed:\n\n{err}")



def save_combined_pdf_section(file_paths, order_num: int, label: str):
    merged_pdf = None
    dest_pdf = get_sim_output_path(order_num, label)
    try:
        merged_pdf = merge_pdfs_to_temp(file_paths, label)
        copy_pdf_to_output(merged_pdf, dest_pdf)
    finally:
        if merged_pdf is not None:
            try:
                if Path(merged_pdf).exists():
                    Path(merged_pdf).unlink()
                    log(f"Deleted merged PDF: {merged_pdf}")
            except Exception as e:
                log(f"Merged PDF cleanup warning for {merged_pdf}: {e}")


def print_cnc_merged_by_mode(file_paths, start_order_num: int):
    duplex = [p for p in file_paths if classify_cnc(p) == "duplex"]
    simplex = [p for p in file_paths if classify_cnc(p) == "simplex"]

    current = start_order_num

    for pdf in duplex:
        save_pdf_simulation(pdf, current, f"CNC_Duplex_{pdf.stem}", "duplex")
        current += 1

    if simplex:
        save_combined_pdf_section(simplex, current, "CNC_Simplex_Merged")
        current += 1

    return current


def run_prints(plan, printer_name: str):
    order_num = 1

    log("Starting simulation output sequence")

    save_pdf_simulation(plan["fwo"], order_num, "Fabrication_Work_Order", "simplex")
    order_num += 1

    excel = None
    try:
        excel = get_excel_app()
        export_excel_to_pdf(plan["bom"], get_sim_output_path(order_num, "BOM"), first_sheet_only=False, excel_app=excel)
        order_num += 1

        order_num = print_cnc_merged_by_mode(plan["cnc"], order_num)

        save_combined_pdf_section(plan["flats"], order_num, "PDFs_Flats_Merged")
        order_num += 1

        export_excel_to_pdf(plan["prf"], get_sim_output_path(order_num, "Production_Release_Form"), first_sheet_only=True, excel_app=excel)
        order_num += 1
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception as e:
                log(f"Excel quit warning after simulated BOM/PRF export: {e}")

    save_pdf_simulation(plan["pack"], order_num, "Electrical_Drawing_Pack_Pages_1_2", "simplex", pages=(0, 1))
    order_num += 1

    save_combined_pdf_section(plan["assemblies"], order_num, "Assemblies_Merged")
    order_num += 1

    log("Completed simulation output sequence")


def main():
    global RUN_HAD_ERRORS

    root = tk.Tk()
    root.withdraw()
    configure_styles(root)

    selected_printer = None
    plan = None
    job_folder = None

    while True:
        if not job_folder:
            folder = choose_folder()
            if not folder:
                log("Cancelled by user before folder selection.")
                return
            job_folder = folder
            init_log(job_folder)
            log(f"Selected job folder: {job_folder}")

        if not selected_printer:
            selected_printer = choose_printer()
            log(f"Printer confirmed: {selected_printer}")

        try:
            plan = build_plan(Path(job_folder))
        except Exception as e:
            RUN_HAD_ERRORS = True
            log(f"Plan build error: {e}")
            warn(str(e))
            return

        summary = make_summary(plan, selected_printer)
        action = summary_dialog(summary)
        log(f"Summary action: {action}")

        if action == "printer":
            selected_printer = None
            continue
        if action == "folder":
            job_folder = None
            SELECTED_CONTEXT_CACHE.clear()
            selected_printer = None
            continue
        if action != "print":
            log("Cancelled by user at summary.")
            return

        break

    init_simulation_output_dir(job_folder)
    run_prints(plan, selected_printer)

    if RUN_HAD_ERRORS:
        info(f"Simulation sequence completed with errors.\n\nOutput folder:\n{SIMULATION_OUTPUT_DIR}")
    else:
        info(f"Simulation sequence complete.\n\nOutput folder:\n{SIMULATION_OUTPUT_DIR}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        save_crash_log(exc)
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(APP_TITLE, f"A fatal error occurred.\n\nA crash log was written next to the script.\n\n{exc}")
        except Exception:
            pass
        raise
    finally:
        save_log()
