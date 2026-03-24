"""
Engineering Tool Hub  —  app.py
FoxFab internal engineering utilities, combined into one application.

Tools:
  • BOM Check          – marks stock parts, copies non-stock PDFs/DXFs
  • Doc Prep & Print   – builds and prints (or simulates) a manufacturing packet
  • SW Batch Update    – updates SolidWorks custom properties and exports DXFs
"""

from __future__ import annotations

import os
import re
import sys
import time
import shutil
import queue
import threading
import tempfile
import traceback
import subprocess
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── optional heavy imports ────────────────────────────────────────────
try:
    import win32print
    import win32com.client
    import pythoncom
except Exception:
    win32print = None
    win32com = None
    pythoncom = None

try:
    import xlwings as xw
except Exception:
    xw = None

try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    PdfReader = None
    PdfWriter = None


# ═════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═════════════════════════════════════════════════════════════════════

APP_TITLE   = "Engineering Tool Hub"
APP_VERSION = "1.0.0"

PREFERRED_PRINTER  = "FoxFab (Konica Bizhub C360i) on NPSVR05"
STOCK_PARTS_FOLDER = r"Z:\FOXFAB_DATA\ENGINEERING\0 PRODUCTS\300 Stock Parts\PDFs & Flats"
BOM_SHEET_NAME     = "FFMPL"
EXCEL_EXTENSIONS   = {".xlsx", ".xls", ".xlsm"}
PDF_EXTENSIONS     = {".pdf", ".PDF"}

# Sidebar
C_SIDEBAR   = "#1E2B40"
C_SIDEBAR_H = "#263550"
C_SIDEBAR_A = "#1F6FEB"
C_SIDEBAR_T = "#CBD5E1"     # normal text
C_SIDEBAR_M = "#64748B"     # muted text (version, status)
C_SIDEBAR_D = "#2D3E56"     # divider

# Content
C_BG        = "#F6F8FB"
C_PANEL     = "#FFFFFF"
C_BORDER    = "#D8DEE9"
C_ACCENT    = "#1F6FEB"
C_ACCENT_H  = "#1558D6"
C_TEXT      = "#1F2937"
C_SUBTLE    = "#6B7280"
C_SUCCESS   = "#16A34A"
C_ERROR     = "#DC2626"
C_WARN      = "#D97706"

# Terminal
C_TERM_BG   = "#0F1923"
C_TERM_FG   = "#E2E8F0"
C_TERM_HDR  = "#1A2535"
C_TERM_INFO = "#60A5FA"
C_TERM_OK   = "#4ADE80"
C_TERM_ERR  = "#F87171"
C_TERM_WARN = "#FCD34D"
C_TERM_MUTE = "#4B5563"

F_BODY  = ("Segoe UI", 10)
F_SMALL = ("Segoe UI", 9)
F_MONO  = ("Consolas", 10)


# ═════════════════════════════════════════════════════════════════════
#  UTILITIES
# ═════════════════════════════════════════════════════════════════════

def exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def safe_name(text: str) -> str:
    if not text:
        return "unknown"
    return re.sub(r'[\\/:*?"<>|]+', "_", text).strip() or "unknown"


def _find_es_exe() -> str:
    if getattr(sys, "frozen", False):
        bundled = os.path.join(sys._MEIPASS, "es.exe")
        if os.path.isfile(bundled):
            return bundled
    for candidate in [exe_dir() / "es.exe",
                      exe_dir() / "tools" / "BomCheck" / "es.exe"]:
        if candidate.exists():
            return str(candidate)
    return "es.exe"


ES_EXE = _find_es_exe()


# ═════════════════════════════════════════════════════════════════════
#  MASTER LOG
# ═════════════════════════════════════════════════════════════════════

class MasterLog:
    def __init__(self):
        self._lock = threading.Lock()
        log_dir = exe_dir() / "logs"
        log_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d")
        self._path = log_dir / f"ETH_master_{stamp}.log"

    def append_section(self, tool: str, lines: list[str]):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        block = (
            f"\n{'=' * 70}\n"
            f"  TOOL : {tool}\n"
            f"  START: {now}\n"
            f"{'=' * 70}\n"
            + "\n".join(lines)
            + "\n"
        )
        with self._lock:
            try:
                with open(self._path, "a", encoding="utf-8") as f:
                    f.write(block)
            except Exception:
                pass

    @property
    def path(self) -> Path:
        return self._path


_master_log = MasterLog()


# ═════════════════════════════════════════════════════════════════════
#  STOP MECHANISM
# ═════════════════════════════════════════════════════════════════════

_stop_event = threading.Event()


class StopRequested(Exception):
    pass


def check_stop():
    if _stop_event.is_set():
        raise StopRequested()


# ═════════════════════════════════════════════════════════════════════
#  APP CLASS
# ═════════════════════════════════════════════════════════════════════

class App:
    TOOLS = [
        ("bom", "BOM Check"),
        ("dpp", "Doc Prep & Print"),
        ("sw",  "SW Batch Update"),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1300x820")
        root.minsize(1100, 720)
        root.configure(bg=C_BG)

        self._log_queue: queue.Queue = queue.Queue()
        self._worker_thread: threading.Thread | None = None
        self._active_key = "bom"
        self._active_term: tk.Text | None = None
        self._panel_ui: dict[str, dict] = {}   # key -> {run, stop, progress, lbl}
        self._panel_terms: dict[str, tk.Text] = {}

        self._configure_styles()
        self._build_layout()
        self._switch_tool("bom")
        self._poll_log_queue()

    # ── styles ────────────────────────────────────────────────────────

    def _configure_styles(self):
        s = ttk.Style(self.root)
        try:
            s.theme_use("vista")
        except Exception:
            try:
                s.theme_use("clam")
            except Exception:
                pass
        s.configure("ETH.Horizontal.TProgressbar",
                    troughcolor=C_BORDER, background=C_ACCENT, thickness=7)

    # ── layout ────────────────────────────────────────────────────────

    def _build_layout(self):
        outer = tk.Frame(self.root, bg=C_BG)
        outer.pack(fill="both", expand=True)

        # sidebar
        self._sidebar = tk.Frame(outer, bg=C_SIDEBAR, width=215)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)
        self._build_sidebar()

        # 1px divider
        tk.Frame(outer, bg=C_BORDER, width=1).pack(side="left", fill="y")

        # content area
        self._content = tk.Frame(outer, bg=C_BG)
        self._content.pack(side="left", fill="both", expand=True)

        # build each panel (hidden until activated)
        self._panels: dict[str, tk.Frame] = {}
        builders = {
            "bom": self._build_bom_panel,
            "dpp": self._build_dpp_panel,
            "sw":  self._build_sw_panel,
        }
        for key, _ in self.TOOLS:
            frame = tk.Frame(self._content, bg=C_BG)
            self._panels[key] = frame
            builders[key](frame)

    def _build_sidebar(self):
        sb = self._sidebar

        # App title
        title_f = tk.Frame(sb, bg=C_SIDEBAR, pady=20)
        title_f.pack(fill="x")
        tk.Label(title_f, text="Engineering", bg=C_SIDEBAR,
                 fg="#FFFFFF", font=("Segoe UI", 13, "bold")).pack()
        tk.Label(title_f, text="Tool Hub", bg=C_SIDEBAR,
                 fg="#93C5FD", font=("Segoe UI", 11)).pack()
        tk.Label(title_f, text=f"v{APP_VERSION}  —  FoxFab",
                 bg=C_SIDEBAR, fg=C_SIDEBAR_M,
                 font=("Segoe UI", 8)).pack(pady=(3, 0))

        tk.Frame(sb, bg=C_SIDEBAR_D, height=1).pack(fill="x", padx=16, pady=(0, 6))

        # nav buttons
        self._nav_btns: dict[str, tk.Button] = {}
        for key, label in self.TOOLS:
            btn = tk.Button(
                sb, text=f"   {label}",
                anchor="w", bg=C_SIDEBAR, fg=C_SIDEBAR_T,
                activebackground=C_SIDEBAR_H, activeforeground="#FFFFFF",
                font=("Segoe UI", 10), bd=0, pady=10,
                cursor="hand2", relief="flat",
                command=lambda k=key: self._switch_tool(k),
            )
            btn.pack(fill="x")
            self._nav_btns[key] = btn

        # master log path label (bottom)
        tk.Frame(sb, bg=C_SIDEBAR_D, height=1).pack(
            fill="x", padx=16, side="bottom", pady=(0, 2))
        self._status_lbl = tk.Label(
            sb, text="Ready", bg=C_SIDEBAR, fg=C_SIDEBAR_M,
            font=("Segoe UI", 8), anchor="w", padx=14, wraplength=190,
            justify="left")
        self._status_lbl.pack(side="bottom", fill="x", pady=(0, 6))

    def _switch_tool(self, key: str):
        self._active_key = key
        for k, btn in self._nav_btns.items():
            if k == key:
                btn.configure(bg=C_SIDEBAR_A, fg="#FFFFFF",
                               font=("Segoe UI", 10, "bold"))
            else:
                btn.configure(bg=C_SIDEBAR, fg=C_SIDEBAR_T,
                               font=("Segoe UI", 10))
        for k, frame in self._panels.items():
            if k == key:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()
        if key in self._panel_terms:
            self._active_term = self._panel_terms[key]

    # ── shared UI builders ────────────────────────────────────────────

    def _section_header(self, parent, title: str, subtitle: str = ""):
        hdr = tk.Frame(parent, bg=C_BG)
        hdr.pack(fill="x", padx=26, pady=(20, 4))
        tk.Label(hdr, text=title, bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(hdr, text=subtitle, bg=C_BG, fg=C_SUBTLE,
                     font=F_SMALL).pack(anchor="w", pady=(2, 0))
        tk.Frame(parent, bg=C_ACCENT, height=2).pack(
            fill="x", padx=26, pady=(4, 0))

    def _card(self, parent, title: str = "", padx=26, pady=(8, 4)):
        wrap = tk.Frame(parent, bg=C_BG)
        wrap.pack(fill="x", padx=padx, pady=pady)
        border = tk.Frame(wrap, bg=C_BORDER, bd=0)
        border.pack(fill="x")
        inner = tk.Frame(border, bg=C_PANEL)
        inner.pack(fill="x", padx=1, pady=1)
        content = tk.Frame(inner, bg=C_PANEL)
        content.pack(fill="x", padx=14, pady=10)
        if title:
            tk.Label(content, text=title, bg=C_PANEL, fg=C_TEXT,
                     font=("Segoe UI", 10, "bold")).pack(
                         anchor="w", pady=(0, 6))
        return content

    def _field_row(self, parent, label: str, var: tk.StringVar,
                   mode="file", filetypes=None, width=16):
        row = tk.Frame(parent, bg=C_PANEL)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=C_PANEL, fg=C_SUBTLE,
                 font=F_SMALL, width=width, anchor="w").pack(side="left")
        entry = tk.Entry(row, textvariable=var, font=F_BODY,
                         bg="#F8FAFC", fg=C_TEXT,
                         relief="solid", bd=1,
                         highlightbackground=C_BORDER,
                         highlightthickness=1)
        entry.pack(side="left", fill="x", expand=True, padx=(4, 6))

        def _browse():
            if mode == "dir":
                p = filedialog.askdirectory(title=f"Select {label}",
                                             mustexist=True)
            else:
                p = filedialog.askopenfilename(
                    title=f"Select {label}",
                    filetypes=filetypes or [("All files", "*.*")])
            if p:
                var.set(p)

        tk.Button(row, text="Browse", command=_browse,
                  bg=C_BG, fg=C_ACCENT, font=F_SMALL,
                  relief="solid", bd=1, padx=8, pady=2,
                  cursor="hand2").pack(side="left")
        return entry

    def _action_bar(self, parent, key: str, run_cmd, stop_cmd,
                    run_label="Run", extras=None):
        bar = tk.Frame(parent, bg=C_BG)
        bar.pack(fill="x", padx=26, pady=(6, 2))

        run_btn = tk.Button(
            bar, text=run_label,
            bg=C_ACCENT, fg="#FFFFFF",
            activebackground=C_ACCENT_H, activeforeground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=20, pady=7, cursor="hand2",
            command=run_cmd,
        )
        run_btn.pack(side="left")

        stop_btn = tk.Button(
            bar, text="  Stop  ",
            bg="#EF4444", fg="#FFFFFF",
            activebackground="#DC2626", activeforeground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=14, pady=7, cursor="hand2",
            state="disabled",
            command=stop_cmd,
        )
        stop_btn.pack(side="left", padx=(8, 0))

        if extras:
            for lbl, cmd in extras:
                tk.Button(bar, text=lbl, command=cmd,
                          bg=C_BG, fg=C_TEXT, font=F_BODY,
                          relief="solid", bd=1, padx=12, pady=6,
                          cursor="hand2").pack(side="left", padx=(8, 0))

        progress = ttk.Progressbar(
            bar, style="ETH.Horizontal.TProgressbar",
            mode="determinate", value=0, maximum=100, length=200)
        progress.pack(side="left", padx=(20, 0), pady=2)

        prog_lbl = tk.Label(bar, text="", bg=C_BG, fg=C_SUBTLE, font=F_SMALL)
        prog_lbl.pack(side="left", padx=(8, 0))

        self._panel_ui[key] = dict(
            run=run_btn, stop=stop_btn,
            progress=progress, lbl=prog_lbl)
        return bar

    def _terminal(self, parent, key: str, rows=14):
        wrap = tk.Frame(parent, bg=C_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=(4, 18))

        # header
        hdr = tk.Frame(wrap, bg=C_TERM_HDR)
        hdr.pack(fill="x")
        tk.Label(hdr, text=" Output", bg=C_TERM_HDR, fg="#94A3B8",
                 font=("Segoe UI", 8, "bold"),
                 pady=4, padx=6).pack(side="left")

        def _clear():
            t.configure(state="normal")
            t.delete("1.0", "end")
            t.configure(state="disabled")

        tk.Button(hdr, text="Clear", bg=C_TERM_HDR, fg="#94A3B8",
                  font=("Segoe UI", 8), bd=0, padx=8, pady=3,
                  activebackground="#263550", activeforeground="#FFFFFF",
                  cursor="hand2", relief="flat",
                  command=_clear).pack(side="right")

        # body
        body = tk.Frame(wrap, bg=C_TERM_BG)
        body.pack(fill="both", expand=True)

        sb = tk.Scrollbar(body, orient="vertical", bg=C_TERM_BG,
                          troughcolor=C_TERM_BG)
        sb.pack(side="right", fill="y")

        t = tk.Text(
            body, bg=C_TERM_BG, fg=C_TERM_FG,
            font=F_MONO, wrap="word", state="disabled",
            yscrollcommand=sb.set, bd=0,
            padx=10, pady=8,
            selectbackground=C_ACCENT,
            height=rows,
        )
        t.pack(side="left", fill="both", expand=True)
        sb.config(command=t.yview)

        t.tag_configure("info",    foreground=C_TERM_INFO)
        t.tag_configure("ok",      foreground=C_TERM_OK)
        t.tag_configure("error",   foreground=C_TERM_ERR)
        t.tag_configure("warn",    foreground=C_TERM_WARN)
        t.tag_configure("heading", foreground="#FFFFFF",
                         font=("Consolas", 10, "bold"))
        t.tag_configure("muted",   foreground=C_TERM_MUTE)

        self._panel_terms[key] = t
        return t

    # ── terminal write ────────────────────────────────────────────────

    def _term_write(self, term: tk.Text, msg: str, tag: str = ""):
        term.configure(state="normal")
        if tag:
            term.insert("end", msg + "\n", tag)
        else:
            term.insert("end", msg + "\n")
        term.see("end")
        term.configure(state="disabled")

    # ── running state ─────────────────────────────────────────────────

    def _set_running(self, key: str, running: bool):
        ui = self._panel_ui.get(key)
        if not ui:
            return
        ui["run"].configure(state="disabled" if running else "normal")
        ui["stop"].configure(state="normal" if running else "disabled")
        if running:
            ui["progress"].configure(mode="indeterminate")
            ui["progress"].start(12)
        else:
            ui["progress"].stop()
            ui["progress"].configure(mode="determinate", value=0)
            ui["lbl"].configure(text="")

    def _set_progress(self, key: str, value: int, maximum: int, label: str = ""):
        ui = self._panel_ui.get(key)
        if not ui:
            return
        ui["progress"].stop()
        ui["progress"].configure(mode="determinate",
                                  maximum=max(maximum, 1), value=value)
        if label:
            ui["lbl"].configure(text=label)

    def _set_status(self, msg: str):
        try:
            self._status_lbl.configure(text=msg)
        except Exception:
            pass

    # ── log queue polling ─────────────────────────────────────────────

    def _poll_log_queue(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                tag, payload = item
                if tag == "__done__":
                    self._on_worker_done(*payload)
                elif tag == "__progress__":
                    key, v, mx, lbl = payload
                    self._set_progress(key, v, mx, lbl)
                elif tag == "__status__":
                    self._set_status(payload)
                else:
                    term = self._panel_terms.get(self._active_key)
                    if term:
                        self._term_write(term, payload, tag)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_log_queue)

    def _on_worker_done(self, key: str, success: bool,
                        tool_name: str, log_lines: list[str]):
        self._set_running(key, False)
        _master_log.append_section(tool_name, log_lines)
        status = f"{tool_name} — done" if success else f"{tool_name} — errors"
        self._set_status(status)
        self._worker_thread = None

    # ── stop handler ──────────────────────────────────────────────────

    def _on_stop(self):
        _stop_event.set()
        self._log_queue.put(
            ("warn", "[USER] Stop requested — will halt after current step..."))
        self._set_status("Stopping...")
        ui = self._panel_ui.get(self._active_key)
        if ui:
            ui["stop"].configure(state="disabled")


# ═════════════════════════════════════════════════════════════════════
#  BOM CHECK PANEL
# ═════════════════════════════════════════════════════════════════════

    def _build_bom_panel(self, parent):
        self._section_header(
            parent,
            "BOM Check",
            "Mark stock parts in the FFMPL sheet, then copy non-stock PDFs and DXFs to the target folder.")

        card = self._card(parent, "Configuration")

        self._bom_wb     = tk.StringVar()
        self._bom_target = tk.StringVar()

        self._field_row(card, "BOM Workbook", self._bom_wb,
                        mode="file",
                        filetypes=[("Excel", "*.xlsx *.xlsm *.xls"),
                                   ("All files", "*.*")])
        self._field_row(card, "Target Folder", self._bom_target, mode="dir")

        note_row = tk.Frame(card, bg=C_PANEL)
        note_row.pack(fill="x", pady=(8, 0))
        tk.Label(note_row,
                 text="The BOM workbook must be closed in Excel before running.",
                 bg=C_PANEL, fg=C_WARN, font=F_SMALL).pack(anchor="w")

        self._action_bar(parent, "bom", self._run_bom, self._on_stop,
                         run_label="  Run BOM Check  ")
        self._terminal(parent, "bom", rows=15)

    def _run_bom(self):
        wb = self._bom_wb.get().strip()
        tgt = self._bom_target.get().strip()

        if not wb:
            messagebox.showwarning(APP_TITLE, "Please select a BOM workbook.")
            return
        if not os.path.isfile(wb):
            messagebox.showwarning(APP_TITLE, f"File not found:\n{wb}")
            return
        if not tgt:
            messagebox.showwarning(APP_TITLE, "Please select a target folder.")
            return
        if not os.path.isdir(tgt):
            messagebox.showwarning(APP_TITLE, f"Target folder not found:\n{tgt}")
            return

        confirmed = messagebox.askyesno(
            APP_TITLE,
            "BOM Check requires the workbook to be CLOSED in Excel.\n\n"
            "Have you closed the workbook and are ready to proceed?",
        )
        if not confirmed:
            return

        self._active_key = "bom"
        _stop_event.clear()
        self._set_running("bom", True)
        self._set_status("BOM Check — running...")

        t = threading.Thread(
            target=self._bom_worker, args=(wb, tgt), daemon=True)
        self._worker_thread = t
        t.start()

    def _bom_worker(self, wb_path: str, target: str):
        q = self._log_queue
        log_lines: list[str] = []
        success = False

        def emit(msg: str, tag: str = "info"):
            ts = datetime.now().strftime("%H:%M:%S")
            line = f"[{ts}]  {msg}"
            log_lines.append(line)
            q.put((tag, line))

        try:
            if xw is None:
                raise RuntimeError(
                    "xlwings is not installed. Run:  pip install xlwings")

            emit(f"Opening workbook: {Path(wb_path).name}", "heading")
            xw_app = xw.App(visible=False, add_book=False)
            xw_app.display_alerts = False
            xw_app.screen_updating = False

            try:
                wb = xw_app.books.open(wb_path)
            except Exception as e:
                raise RuntimeError(f"Could not open workbook: {e}")

            sheet_names = [s.name for s in wb.sheets]
            if BOM_SHEET_NAME not in sheet_names:
                wb.close(); xw_app.quit()
                raise RuntimeError(
                    f"Sheet '{BOM_SHEET_NAME}' not found.\n"
                    f"Available sheets: {sheet_names}")

            ws = wb.sheets[BOM_SHEET_NAME]
            DATA_START = 6

            last_row = ws.range(f"A{DATA_START}").end("down").row
            if last_row > 1_000_000:
                last_row = DATA_START

            raw = ws.range(f"A{DATA_START}:A{last_row}").value
            if raw is None:
                part_numbers = []
            elif not isinstance(raw, list):
                part_numbers = [raw]
            else:
                part_numbers = raw

            total = sum(1 for p in part_numbers if p and str(p).strip())
            emit(f"Found {total} part numbers to process.", "info")

            # ── PASS 1 ─────────────────────────────────────────────────
            emit("Pass 1  —  Stock Parts Check", "heading")

            stock_found = 0
            stock_flags: list[bool] = []
            done = 0

            for i, pn in enumerate(part_numbers):
                check_stop()
                if pn is None or not str(pn).strip():
                    stock_flags.append(False)
                    continue
                pn_s = str(pn).strip()
                done += 1
                q.put(("__progress__",
                       ("bom", done, total, f"Pass 1 — {done}/{total}")))

                is_stock = _bom_check_stock(pn_s)
                row = DATA_START + i
                if is_stock:
                    ws.range(f"B{row}").value = "X"
                    ws.range(f"G{row}").value = "S"
                    ws.range(f"H{row}").value = "S"
                    emit(f"  {pn_s:<28}  STOCK", "ok")
                    stock_found += 1
                    stock_flags.append(True)
                else:
                    stock_flags.append(False)

            emit(f"  {stock_found} stock  /  {done - stock_found} non-stock", "info")

            # ── PASS 2 ─────────────────────────────────────────────────
            emit("Pass 2  —  Non-Stock PDF + DXF Copy", "heading")

            tgt_path = Path(target)
            pdf_cp = pdf_ex = pdf_nf = 0
            dxf_cp = dxf_ex = dxf_nf = 0
            ns_done = 0
            ns_total = sum(1 for i, p in enumerate(part_numbers)
                           if p and str(p).strip()
                           and i < len(stock_flags) and not stock_flags[i])

            for i, pn in enumerate(part_numbers):
                check_stop()
                if pn is None or not str(pn).strip():
                    continue
                if i < len(stock_flags) and stock_flags[i]:
                    continue
                pn_s = str(pn).strip()
                ns_done += 1
                q.put(("__progress__",
                       ("bom", ns_done, max(ns_total, 1),
                        f"Pass 2 — {ns_done}/{ns_total}")))

                rev = _bom_find_revision(pn_s)
                row = DATA_START + i

                pdf_r = _bom_find_copy(pn_s, rev, tgt_path, "pdf")
                dxf_r = _bom_find_copy(pn_s, rev, tgt_path, "dxf")

                # concise status labels
                plbl = ("copied" if pdf_r is True
                        else "exists" if pdf_r is None else "not found")
                dlbl = ("copied" if dxf_r is True
                        else "exists" if dxf_r is None else "not found")

                if pdf_r is True:   pdf_cp += 1
                elif pdf_r is None: pdf_ex += 1
                else:               pdf_nf += 1

                if dxf_r is True:   dxf_cp += 1
                elif dxf_r is None: dxf_ex += 1
                else:               dxf_nf += 1

                has_issue = pdf_r is False or dxf_r is False
                tag = "warn" if has_issue else "ok"
                rv  = f" [{rev}]" if rev else ""
                emit(f"  {pn_s:<28}{rv:<6}  PDF: {plbl:<12}  DXF: {dlbl}", tag)

                if pdf_r is not False or dxf_r is not False:
                    ws.range(f"G{row}").value = "X"

            emit(f"  PDFs  copied {pdf_cp}  existed {pdf_ex}  not found {pdf_nf}",
                 "info")
            emit(f"  DXFs  copied {dxf_cp}  existed {dxf_ex}  not found {dxf_nf}",
                 "info")

            changes = stock_found + pdf_cp + pdf_ex + dxf_cp + dxf_ex
            if changes > 0:
                wb.save()
                emit(f"Workbook saved.", "ok")
            else:
                emit("No changes made.", "warn")

            wb.close()
            xw_app.quit()
            success = True
            emit("Done.", "ok")

        except StopRequested:
            emit("Stopped by user.", "warn")
        except Exception as e:
            emit(f"ERROR: {e}", "error")
            for ln in traceback.format_exc().splitlines():
                emit(ln, "error")
        finally:
            q.put(("__done__", ("bom", success, "BOM Check", log_lines)))


# ═════════════════════════════════════════════════════════════════════
#  DOC PREP & PRINT PANEL
# ═════════════════════════════════════════════════════════════════════

    def _build_dpp_panel(self, parent):
        self._section_header(
            parent,
            "Doc Prep & Print",
            "Build a manufacturing print packet and send it to the printer (or save PDFs in Simulation Mode).")

        card = self._card(parent, "Configuration")

        self._dpp_folder  = tk.StringVar()
        self._dpp_printer = tk.StringVar(value=PREFERRED_PRINTER)
        self._dpp_sim     = tk.BooleanVar(value=False)
        self._dpp_plan: dict | None = None

        self._field_row(card, "Job Folder", self._dpp_folder, mode="dir")

        # printer row
        pr_row = tk.Frame(card, bg=C_PANEL)
        pr_row.pack(fill="x", pady=3)
        tk.Label(pr_row, text="Printer", bg=C_PANEL, fg=C_SUBTLE,
                 font=F_SMALL, width=16, anchor="w").pack(side="left")
        tk.Entry(pr_row, textvariable=self._dpp_printer,
                 font=F_BODY, bg="#F8FAFC", fg=C_TEXT,
                 relief="solid", bd=1,
                 highlightbackground=C_BORDER,
                 highlightthickness=1).pack(
                     side="left", fill="x", expand=True, padx=(4, 6))
        tk.Button(pr_row, text="Detect",
                  command=self._dpp_detect_printer,
                  bg=C_BG, fg=C_ACCENT, font=F_SMALL,
                  relief="solid", bd=1, padx=8, pady=2,
                  cursor="hand2").pack(side="left")

        # simulation toggle
        sim_row = tk.Frame(card, bg=C_PANEL)
        sim_row.pack(fill="x", pady=(10, 2))
        tk.Checkbutton(
            sim_row,
            text="Simulation Mode  —  saves PDFs to folder instead of printing",
            variable=self._dpp_sim,
            bg=C_PANEL, fg=C_TEXT, font=F_BODY,
            activebackground=C_PANEL, selectcolor=C_PANEL,
        ).pack(anchor="w")

        self._action_bar(
            parent, "dpp",
            self._run_dpp, self._on_stop,
            run_label="  Build Plan & Run  ",
            extras=[("Clear Plan", self._dpp_clear_plan)],
        )
        self._terminal(parent, "dpp", rows=15)

    def _dpp_detect_printer(self):
        if win32print is None:
            messagebox.showwarning(APP_TITLE,
                "pywin32 not available — cannot detect printers.")
            return
        try:
            flags = (win32print.PRINTER_ENUM_LOCAL
                     | win32print.PRINTER_ENUM_CONNECTIONS)
            printers = sorted({p[2] for p in win32print.EnumPrinters(flags)})
            if PREFERRED_PRINTER in printers:
                self._dpp_printer.set(PREFERRED_PRINTER)
                return
            choice = self._pick_from_list("Select Printer", printers)
            if choice:
                self._dpp_printer.set(choice)
        except Exception as e:
            messagebox.showwarning(APP_TITLE,
                f"Could not enumerate printers:\n{e}")

    def _pick_from_list(self, title: str, items: list[str]) -> str | None:
        top = tk.Toplevel(self.root)
        top.title(title)
        top.geometry("580x380")
        top.configure(bg=C_BG)
        top.grab_set()
        top.transient(self.root)
        result: dict[str, str | None] = {"v": None}

        tk.Label(top, text=title, bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 12, "bold")).pack(
                     pady=(16, 4), padx=16, anchor="w")

        body = tk.Frame(top, bg=C_BORDER)
        body.pack(fill="both", expand=True, padx=16, pady=4)
        inner = tk.Frame(body, bg=C_PANEL)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        sb2 = tk.Scrollbar(inner)
        sb2.pack(side="right", fill="y")
        lb = tk.Listbox(inner, yscrollcommand=sb2.set,
                         font=F_BODY, bd=0, highlightthickness=0,
                         selectbackground=C_ACCENT,
                         selectforeground="#FFFFFF")
        for item in items:
            lb.insert("end", item)
        lb.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        sb2.config(command=lb.yview)

        bar2 = tk.Frame(top, bg=C_BG)
        bar2.pack(fill="x", padx=16, pady=(4, 14))

        def _ok():
            sel = lb.curselection()
            if sel:
                result["v"] = items[sel[0]]
            top.destroy()

        tk.Button(bar2, text="OK",
                  bg=C_ACCENT, fg="#FFFFFF",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=16, pady=5,
                  cursor="hand2",
                  command=_ok).pack(side="left")
        tk.Button(bar2, text="Cancel",
                  bg=C_BG, fg=C_TEXT, font=F_BODY,
                  relief="solid", bd=1, padx=12, pady=5,
                  cursor="hand2",
                  command=top.destroy).pack(side="left", padx=(8, 0))

        top.wait_window()
        return result["v"]

    def _dpp_clear_plan(self):
        self._dpp_plan = None
        term = self._panel_terms.get("dpp")
        if term:
            self._term_write(term, "Plan cleared.", "muted")

    def _run_dpp(self):
        folder  = self._dpp_folder.get().strip()
        printer = self._dpp_printer.get().strip()
        sim     = self._dpp_sim.get()

        if not folder:
            messagebox.showwarning(APP_TITLE, "Please select a job folder.")
            return
        if not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, f"Folder not found:\n{folder}")
            return
        if not sim and not printer:
            messagebox.showwarning(APP_TITLE,
                "Please enter a printer name, or enable Simulation Mode.")
            return

        self._active_key = "dpp"
        term = self._panel_terms.get("dpp")

        # Build plan on main thread (dialogs may appear)
        try:
            if term:
                self._term_write(term, f"Building plan: {folder}", "info")
            plan = dpp_build_plan(Path(folder), self)
        except Exception as e:
            if term:
                self._term_write(term, f"Plan error: {e}", "error")
            return

        # Show summary + confirm
        summary = dpp_make_summary(plan, printer, sim)
        if term:
            self._term_write(term, summary, "muted")

        ok = messagebox.askyesno(
            APP_TITLE,
            summary + "\n\nProceed?",
        )
        if not ok:
            if term:
                self._term_write(term, "Cancelled by user.", "warn")
            return

        self._dpp_plan = plan
        _stop_event.clear()
        self._set_running("dpp", True)
        self._set_status("Doc Prep — running...")

        t = threading.Thread(
            target=self._dpp_worker,
            args=(plan, printer, sim),
            daemon=True,
        )
        self._worker_thread = t
        t.start()

    def _dpp_worker(self, plan: dict, printer_name: str, simulation: bool):
        q = self._log_queue
        log_lines: list[str] = []
        success = False

        def emit(msg: str, tag: str = "info"):
            ts = datetime.now().strftime("%H:%M:%S")
            line = f"[{ts}]  {msg}"
            log_lines.append(line)
            q.put((tag, line))

        if pythoncom:
            pythoncom.CoInitialize()

        sim_dir: Path | None = None
        try:
            if simulation:
                job_name = safe_name(Path(plan["job_folder"]).name)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                sim_dir = exe_dir() / f"Simulated_Print_Output_{job_name}_{ts}"
                sim_dir.mkdir(parents=True, exist_ok=True)
                emit(f"Simulation output: {sim_dir}", "info")

            sections = dpp_build_sections(plan, printer_name, simulation, sim_dir)
            total = len(sections)

            for idx, (title, func) in enumerate(sections):
                check_stop()
                q.put(("__progress__",
                       ("dpp", idx, total, f"{idx+1}/{total}: {title}")))
                emit(f"  ▶  {title}", "heading")
                try:
                    func()
                    emit(f"     ✓  {title}", "ok")
                except StopRequested:
                    raise
                except Exception as e:
                    emit(f"     ✗  {title}: {e}", "error")
                    emit("     Skipping to next section...", "warn")

            q.put(("__progress__", ("dpp", total, total, "Done")))
            success = True
            if simulation:
                emit(f"Simulation complete. Output: {sim_dir}", "ok")
            else:
                emit("Print sequence complete.", "ok")

        except StopRequested:
            emit("Stopped by user.", "warn")
        except Exception as e:
            emit(f"FATAL: {e}", "error")
            for ln in traceback.format_exc().splitlines():
                emit(ln, "error")
        finally:
            if pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            q.put(("__done__", ("dpp", success, "Doc Prep & Print", log_lines)))


# ═════════════════════════════════════════════════════════════════════
#  SW BATCH UPDATE PANEL
# ═════════════════════════════════════════════════════════════════════

    def _build_sw_panel(self, parent):
        self._section_header(
            parent,
            "SW Batch Update",
            "Run the SolidWorks macro directly to update custom properties and export DXFs.")

        macro_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "tools", "SolidworksBatchUpdate", "SoldworksBatchUpdate.swp")

        # ── How to Run card ───────────────────────────────────────
        card = self._card(parent, "How to Run")

        steps = [
            ("1", "Open SolidWorks 2023."),
            ("2", "Go to  Tools  →  Macro  →  Run..."),
            ("3", "Navigate to the macro file shown below and click Open."),
            ("4", "The macro dialog will appear — fill in your options and click Run."),
        ]
        for num, text in steps:
            row = tk.Frame(card, bg=C_PANEL)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=num, bg=C_ACCENT, fg="white",
                     font=("Segoe UI", 9, "bold"),
                     width=2, anchor="center").pack(side="left", padx=(0, 10))
            tk.Label(row, text=text, bg=C_PANEL, fg=C_TEXT,
                     font=F_BODY, anchor="w").pack(side="left")

        # ── Macro Location card ───────────────────────────────────
        loc_card = self._card(parent, "Macro File Location")

        path_row = tk.Frame(loc_card, bg=C_PANEL)
        path_row.pack(fill="x")

        path_var = tk.StringVar(value=macro_path)
        path_entry = tk.Entry(path_row, textvariable=path_var, font=F_BODY,
                              bg="#F8FAFC", fg=C_TEXT,
                              relief="solid", bd=1, state="readonly")
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def _open_folder():
            folder = os.path.dirname(macro_path)
            if os.path.isdir(folder):
                os.startfile(folder)

        tk.Button(path_row, text="Open Folder", font=F_BODY,
                  bg=C_PANEL, fg=C_ACCENT, relief="solid", bd=1,
                  cursor="hand2", command=_open_folder).pack(side="left")

        # ── What it does card ─────────────────────────────────────
        info_card = self._card(parent, "What the Macro Does")
        for bullet in [
            "Updates  DrawnBy  on all part/assembly files in a chosen folder.",
            "Updates  DwgDrawnBy  and re-stamps all standard drawing properties.",
            "Exports sheet metal flat-pattern DXFs to a chosen destination folder.",
            "Optionally skips files whose name starts with  003-.",
        ]:
            row = tk.Frame(info_card, bg=C_PANEL)
            row.pack(fill="x", pady=2)
            tk.Label(row, text="•", bg=C_PANEL, fg=C_ACCENT,
                     font=F_BODY).pack(side="left", padx=(0, 8))
            tk.Label(row, text=bullet, bg=C_PANEL, fg=C_TEXT,
                     font=F_BODY, anchor="w").pack(side="left")


# ═════════════════════════════════════════════════════════════════════
#  BOM CHECK  —  module-level helpers
# ═════════════════════════════════════════════════════════════════════

def _run_es(args: list[str]) -> str:
    try:
        r = subprocess.run([ES_EXE] + args,
                           capture_output=True, text=True, timeout=10)
        return r.stdout.strip()
    except Exception:
        return ""


def _bom_check_stock(pn: str) -> bool:
    out = _run_es(["-path", STOCK_PARTS_FOLDER, pn])
    return bool([l for l in out.splitlines() if l.strip()])


def _bom_find_revision(pn: str) -> str:
    out = _run_es([pn])
    highest = ""
    for line in out.splitlines():
        stem = Path(line.strip()).stem
        if not stem.upper().startswith(pn.upper()):
            continue
        suffix = stem[len(pn):].strip()
        m = re.match(r'^[-_\s]?r([A-Za-z])$', suffix)
        if m:
            r = m.group(1).upper()
            if r > highest:
                highest = r
    return f"r{highest}" if highest else ""


def _bom_find_copy(pn: str, rev: str, target: Path, ext: str):
    out = _run_es([f"ext:{ext}", pn])
    if not out:
        return False
    expected = ({f"{pn} {rev}".upper(), f"{pn}-{rev}".upper()}
                if rev else {pn.upper()})
    ext_up = f".{ext.upper()}"
    for line in out.splitlines():
        p = Path(line.strip())
        if p.stem.upper() in expected and p.suffix.upper() == ext_up:
            if p.exists():
                dest = target / p.name
                if dest.exists():
                    return None
                shutil.copy2(str(p), str(dest))
                return True
    return False


# ═════════════════════════════════════════════════════════════════════
#  DOC PREP & PRINT  —  module-level helpers
# ═════════════════════════════════════════════════════════════════════

def _dpp_get_excel():
    if win32com is None:
        raise RuntimeError("pywin32 required. pip install pywin32")
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    return app


def _dpp_classify_cnc(p: Path) -> str:
    base = p.stem
    if re.match(r"^[Jj]", base) or re.match(r"^\d{3}-", base):
        return "duplex"
    return "simplex"


def _dpp_list_files(folder: Path) -> list[Path]:
    return sorted([p for p in folder.iterdir() if p.is_file()],
                  key=lambda p: p.name.lower())


def _dpp_match_fwo(folder: Path) -> Path:
    for f in _dpp_list_files(folder):
        if (f.suffix in PDF_EXTENSIONS
                and f.stem == "Fabrication Work Order - Standard v1.0"):
            return f
    raise RuntimeError("Fabrication Work Order PDF not found in 300 Inputs.")


def _dpp_match_excel(folder: Path, token: str, title: str,
                     app: "App") -> Path:
    matches = [f for f in _dpp_list_files(folder)
               if f.suffix.lower() in {e.lower() for e in EXCEL_EXTENSIONS}
               and token.lower() in f.name.lower()]
    if not matches:
        raise RuntimeError(f"No Excel file containing '{token}' found in {title}.")
    if len(matches) == 1:
        return matches[0]
    choice = app._pick_from_list(
        f"Select {title}",
        [f.name for f in matches])
    if not choice:
        raise RuntimeError(f"{title} selection cancelled.")
    return next(f for f in matches if f.name == choice)


def _dpp_match_pack(folder: Path, app: "App") -> Path:
    matches = [f for f in _dpp_list_files(folder)
               if f.suffix in PDF_EXTENSIONS and "PACK" in f.name.upper()]
    if not matches:
        raise RuntimeError("No PDF with 'PACK' in name found in Electrical Drawings.")
    if len(matches) == 1:
        return matches[0]
    choice = app._pick_from_list("Select Electrical Drawing Pack",
                                  [f.name for f in matches])
    if not choice:
        raise RuntimeError("Pack PDF selection cancelled.")
    return next(f for f in matches if f.name == choice)


def _dpp_match_cnc(folder: Path) -> list[Path]:
    pdfs = [f for f in _dpp_list_files(folder) if f.suffix in PDF_EXTENSIONS]
    if not pdfs:
        raise RuntimeError("No PDFs found in CNC folder.")
    return sorted(pdfs, key=lambda p: p.name.lower())


def _dpp_match_flats(folder: Path) -> list[Path]:
    pdfs = [f for f in _dpp_list_files(folder) if f.suffix in PDF_EXTENSIONS]
    if not pdfs:
        raise RuntimeError("No PDFs found in PDFs_Flats folder.")
    return sorted(pdfs, key=lambda p: p.name.lower())


def _dpp_match_assemblies(folder: Path) -> tuple[list[Path], list[Path]]:
    pdfs, lay = [], []
    for f in _dpp_list_files(folder):
        if f.suffix not in PDF_EXTENSIONS:
            continue
        if f.stem.endswith("-LAY"):
            lay.append(f)
        else:
            pdfs.append(f)
    if not pdfs:
        raise RuntimeError("No printable PDFs found in Assemblies folder.")
    return (sorted(pdfs, key=lambda p: p.name.lower()),
            sorted(lay, key=lambda p: p.name.lower()))


def _dpp_get_context(job_folder: Path) -> dict:
    def is_variant(p: Path):
        return p.is_dir() and re.search(r"-\d{2}$", p.name)

    def has_mech_subs(p: Path):
        return all((p / s).is_dir() for s in
                   ["204 BOM", "205 CNC", "202 PDFs_Flats", "203 Assemblies"])

    if is_variant(job_folder) and has_mech_subs(job_folder):
        return {"job_root": job_folder.parent.parent,
                "mech_roots": [job_folder],
                "variant_only": True}

    mech = job_folder / "200 Mech"
    if not mech.is_dir():
        raise RuntimeError("Missing folder: 200 Mech")

    if has_mech_subs(mech):
        return {"job_root": job_folder,
                "mech_roots": [mech],
                "variant_only": False}

    variants = [c for c in sorted(mech.iterdir(), key=lambda p: p.name.lower())
                if is_variant(c) and has_mech_subs(c)]
    if variants:
        return {"job_root": job_folder,
                "mech_roots": variants,
                "variant_only": False}

    raise RuntimeError(
        "Could not find a usable mechanical folder structure. "
        "Expected '200 Mech\\204 BOM' etc. directly, or numbered "
        "variants like '*-01' inside 200 Mech.")


def dpp_build_plan(job_folder: Path, app: "App") -> dict:
    ctx = _dpp_get_context(job_folder)
    base = Path(ctx["job_root"])
    mechs = ctx["mech_roots"]

    plan: dict = {
        "job_folder": str(job_folder),
        "base": str(base),
        "mech_roots": [str(m) for m in mechs],
        "variant_only": ctx["variant_only"],
    }

    plan["fwo"]  = _dpp_match_fwo(base / "300 Inputs")
    plan["bom"]  = _dpp_match_excel(mechs[0] / "204 BOM", "BOM", "BOM", app)
    plan["prf"]  = _dpp_match_excel(
        base / "300 Inputs" / "302 Production Release Form",
        "PRF", "Production Release Form", app)
    plan["pack"] = _dpp_match_pack(base / "100 Elec" / "102 Drawings", app)

    cnc, flats, assemblies, excluded_lay = [], [], [], []
    for m in mechs:
        cnc.extend(_dpp_match_cnc(m / "205 CNC"))
        flats.extend(_dpp_match_flats(m / "202 PDFs_Flats"))
        a, ex = _dpp_match_assemblies(m / "203 Assemblies")
        assemblies.extend(a)
        excluded_lay.extend(ex)

    plan["cnc"]              = sorted(cnc, key=lambda p: p.name.lower())
    plan["flats"]            = sorted(flats, key=lambda p: p.name.lower())
    plan["assemblies"]       = sorted(assemblies, key=lambda p: p.name.lower())
    plan["assemblies_excl"]  = sorted(excluded_lay, key=lambda p: p.name.lower())
    return plan


def dpp_make_summary(plan: dict, printer: str, simulation: bool) -> str:
    cnc_d = sum(1 for p in plan["cnc"] if _dpp_classify_cnc(p) == "duplex")
    cnc_s = len(plan["cnc"]) - cnc_d
    mode  = "SIMULATION MODE (save PDFs)" if simulation else f"PRINT  »  {printer}"
    lines = [
        f"Mode: {mode}",
        f"Job Folder: {plan['job_folder']}",
        "",
        f"  Fabrication Work Order : {plan['fwo'].name}",
        f"  BOM                    : {plan['bom'].name}",
        f"  CNC files              : {len(plan['cnc'])} PDFs"
        f"  ({cnc_d} duplex individual, {cnc_s} simplex merged)",
        f"  PDFs_Flats             : {len(plan['flats'])} PDFs (merged)",
        f"  Production Release Form: {plan['prf'].name}",
        f"  Electrical Pack        : {plan['pack'].name} (pages 1-2)",
        f"  Assemblies             : {len(plan['assemblies'])} PDFs (merged,"
        f" excl. {len(plan['assemblies_excl'])} -LAY files)",
    ]
    return "\n".join(lines)


def _dpp_merge_pdfs(files: list[Path], label: str) -> Path:
    if PdfReader is None:
        raise RuntimeError("pypdf not installed. pip install pypdf")
    writer = PdfWriter()
    for pdf in files:
        reader = PdfReader(str(pdf))
        for page in reader.pages:
            writer.add_page(page)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{label}.pdf")
    tmp_path = Path(tmp.name)
    tmp.close()
    with open(tmp_path, "wb") as f:
        writer.write(f)
    return tmp_path


def _dpp_set_duplex(printer: str, mode: str):
    if win32print is None:
        return
    desired = 1 if mode == "simplex" else 2
    h = win32print.OpenPrinter(printer)
    try:
        props = win32print.GetPrinter(h, 2)
        dm = props.get("pDevMode")
        if dm and hasattr(dm, "Duplex"):
            dm.Duplex = desired
            props["pDevMode"] = dm
            win32print.SetPrinter(h, 2, props, 0)
    finally:
        win32print.ClosePrinter(h)


def _dpp_queue_snap(printer: str) -> dict:
    if win32print is None:
        return {}
    h = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(h, 2)
        total = info.get("cJobs", 0)
        if total <= 0:
            return {}
        jobs = win32print.EnumJobs(h, 0, total, 1)
        return {str(j.get("JobId")): str(j.get("pDocument") or "")
                for j in jobs}
    finally:
        win32print.ClosePrinter(h)


def _dpp_wait_settle(printer: str, settle=2.0, timeout=20):
    if win32print is None:
        time.sleep(1.5)
        return
    deadline = time.time() + timeout
    stable_since = None
    prev = tuple(sorted(_dpp_queue_snap(printer).items()))
    while time.time() < deadline:
        time.sleep(0.5)
        check_stop()
        cur = tuple(sorted(_dpp_queue_snap(printer).items()))
        if cur == prev:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= settle:
                return
        else:
            prev = cur
            stable_since = None


def _dpp_print_pdf(pdf: Path, printer: str, mode: str, pages=None):
    target = pdf
    tmp_cleanup = None
    if pages is not None:
        reader = PdfReader(str(pdf))
        writer = PdfWriter()
        for i in range(pages[0], pages[1] + 1):
            writer.add_page(reader.pages[i])
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix="_page_range.pdf")
        target = Path(tmp.name)
        tmp.close()
        with open(target, "wb") as f:
            writer.write(f)
        tmp_cleanup = target
    try:
        _dpp_set_duplex(printer, mode)
        before = _dpp_queue_snap(printer)
        os.startfile(str(target), "print")
        # wait a moment for spool
        deadline = time.time() + 25
        while time.time() < deadline:
            time.sleep(0.5)
            check_stop()
            cur = _dpp_queue_snap(printer)
            if any(k not in before for k in cur):
                break
        _dpp_wait_settle(printer)
    finally:
        if tmp_cleanup and tmp_cleanup.exists():
            try:
                tmp_cleanup.unlink()
            except Exception:
                pass


def _dpp_print_excel(file: Path, printer: str, first_sheet_only: bool,
                     excel=None):
    created = excel is None
    if created:
        excel = _dpp_get_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(file))
        if first_sheet_only:
            ws = wb.Worksheets(1)
            ws.PrintOut(ActivePrinter=printer)
        else:
            active = wb.ActiveSheet
            ps = active.PageSetup
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            active.PrintOut(ActivePrinter=printer)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if created and excel:
            try:
                excel.Quit()
            except Exception:
                pass


def _dpp_print_merged_pdfs(files: list[Path], printer: str,
                            mode: str, label: str):
    tmp = None
    try:
        tmp = _dpp_merge_pdfs(files, label)
        _dpp_print_pdf(tmp, printer, mode)
    finally:
        if tmp and tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass


def _dpp_sim_save_pdf(src: Path, dest: Path, pages=None):
    if pages is not None:
        reader = PdfReader(str(src))
        writer = PdfWriter()
        for i in range(pages[0], pages[1] + 1):
            writer.add_page(reader.pages[i])
        with open(dest, "wb") as f:
            writer.write(f)
    else:
        shutil.copy2(str(src), str(dest))


def _dpp_sim_excel_to_pdf(file: Path, dest: Path,
                           first_sheet_only: bool, excel=None):
    created = excel is None
    if created:
        excel = _dpp_get_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(file))
        if first_sheet_only:
            ws = wb.Worksheets(1)
            ws.ExportAsFixedFormat(0, str(dest))
        else:
            active = wb.ActiveSheet
            ps = active.PageSetup
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            active.ExportAsFixedFormat(0, str(dest))
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if created and excel:
            try:
                excel.Quit()
            except Exception:
                pass


def _dpp_sim_save_merged(files: list[Path], dest: Path):
    merged = _dpp_merge_pdfs(files, dest.stem)
    try:
        shutil.copy2(str(merged), str(dest))
    finally:
        if merged.exists():
            try:
                merged.unlink()
            except Exception:
                pass


def dpp_build_sections(plan: dict, printer: str,
                       simulation: bool, sim_dir: Path | None) -> list:
    sections = []

    if simulation:
        n = [0]

        def dest(label: str) -> Path:
            n[0] += 1
            return sim_dir / f"{n[0]:02d}_{safe_name(label)}.pdf"

        excel_holder = [None]

        def get_xl():
            if excel_holder[0] is None:
                excel_holder[0] = _dpp_get_excel()
            return excel_holder[0]

        def quit_xl():
            if excel_holder[0]:
                try:
                    excel_holder[0].Quit()
                except Exception:
                    pass
                excel_holder[0] = None

        sections.append(("Fabrication Work Order",
            lambda: _dpp_sim_save_pdf(plan["fwo"], dest("Fabrication_Work_Order"))))

        sections.append(("BOM",
            lambda: _dpp_sim_excel_to_pdf(
                plan["bom"], dest("BOM"),
                first_sheet_only=False, excel=get_xl())))

        for pdf in plan["cnc"]:
            p = pdf  # capture
            if _dpp_classify_cnc(p) == "duplex":
                sections.append((f"CNC (duplex): {p.name}",
                    lambda _p=p: _dpp_sim_save_pdf(_p, dest(f"CNC_{_p.stem}"))))

        cnc_simplex = [p for p in plan["cnc"]
                       if _dpp_classify_cnc(p) == "simplex"]
        if cnc_simplex:
            sections.append(("CNC Simplex (merged)",
                lambda: _dpp_sim_save_merged(cnc_simplex, dest("CNC_Simplex_Merged"))))

        sections.append(("PDFs_Flats (merged)",
            lambda: _dpp_sim_save_merged(plan["flats"], dest("PDFs_Flats_Merged"))))

        sections.append(("Production Release Form",
            lambda: _dpp_sim_excel_to_pdf(
                plan["prf"], dest("Production_Release_Form"),
                first_sheet_only=True, excel=get_xl())))

        sections.append(("Electrical Pack (pages 1-2)",
            lambda: _dpp_sim_save_pdf(
                plan["pack"], dest("Electrical_Pack_Pages_1_2"), pages=(0, 1))))

        sections.append(("Assemblies (merged)",
            lambda: _dpp_sim_save_merged(
                plan["assemblies"], dest("Assemblies_Merged"))))

        sections.append(("Cleanup Excel", quit_xl))

    else:
        excel_holder = [None]

        def get_xl():
            if excel_holder[0] is None:
                excel_holder[0] = _dpp_get_excel()
            return excel_holder[0]

        def quit_xl():
            if excel_holder[0]:
                try:
                    excel_holder[0].Quit()
                except Exception:
                    pass
                excel_holder[0] = None

        if win32print:
            win32print.SetDefaultPrinter(printer)

        sections.append(("Fabrication Work Order",
            lambda: _dpp_print_pdf(plan["fwo"], printer, "simplex")))

        sections.append(("BOM",
            lambda: _dpp_print_excel(
                plan["bom"], printer,
                first_sheet_only=False, excel=get_xl())))

        for pdf in plan["cnc"]:
            p = pdf
            if _dpp_classify_cnc(p) == "duplex":
                sections.append((f"CNC (duplex): {p.name}",
                    lambda _p=p: _dpp_print_pdf(_p, printer, "duplex")))

        cnc_simplex = [p for p in plan["cnc"]
                       if _dpp_classify_cnc(p) == "simplex"]
        if cnc_simplex:
            sections.append(("CNC Simplex (merged)",
                lambda: _dpp_print_merged_pdfs(
                    cnc_simplex, printer, "simplex", "CNC_Simplex")))

        sections.append(("PDFs_Flats (merged)",
            lambda: _dpp_print_merged_pdfs(
                plan["flats"], printer, "simplex", "PDFs_Flats")))

        sections.append(("Production Release Form",
            lambda: _dpp_print_excel(
                plan["prf"], printer,
                first_sheet_only=True, excel=get_xl())))

        sections.append(("Electrical Pack (pages 1-2)",
            lambda: _dpp_print_pdf(plan["pack"], printer, "simplex", pages=(0, 1))))

        sections.append(("Assemblies (merged)",
            lambda: _dpp_print_merged_pdfs(
                plan["assemblies"], printer, "simplex", "Assemblies")))

        sections.append(("Cleanup Excel", quit_xl))

    return sections


# ═════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════════════


def main():
    root = tk.Tk()
    root.withdraw()

    # DPI awareness on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root.deiconify()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
