#!/usr/bin/env python3
"""
Google Maps Scraper - Desktop Client
"""

import tkinter as tk
from tkinter import ttk
import threading
import json
import math
import requests
from datetime import datetime
from pathlib import Path
import openpyxl
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_API_URL   = "https://gmaps-api.fouiguira.com"
POLL_INTERVAL     = 3
OVERALL_TIMEOUT   = 20 * 60
CONFIG_FILE       = Path.home() / ".gmaps-client" / "config.json"
GRID_CELL_SIZE_KM = 3.0
GRID_ZOOM         = 14
GRID_MAX_PARALLEL = 5

# Light professional color palette
C_BG        = "#F1F5F9"   # slate-100 page background
C_PANEL     = "#FFFFFF"   # white cards
C_BORDER    = "#E2E8F0"   # slate-200 borders
C_BORDER_F  = "#CBD5E1"   # slate-300 focused borders
C_ACCENT    = "#2563EB"   # blue-600
C_ACCENT_H  = "#1D4ED8"   # blue-700 hover
C_GREEN     = "#059669"   # emerald-600
C_GREEN_BG  = "#ECFDF5"   # emerald-50
C_RED       = "#DC2626"   # red-600
C_RED_BG    = "#FEF2F2"   # red-50
C_ORANGE    = "#D97706"   # amber-600
C_ORANGE_BG = "#FFFBEB"   # amber-50
C_TEXT      = "#1E293B"   # slate-800 primary
C_MUTED     = "#64748B"   # slate-500 secondary
C_HINT      = "#94A3B8"   # slate-400 placeholder/dim
C_WHITE     = "#FFFFFF"
C_INPUT_BG  = "#F8FAFC"   # slate-50 inputs

KM_PER_DEG_LAT = 111.32

CITY_PRESETS = {
    "Custom...":  "",
    "Berlin":     "52.338,13.088,52.675,13.761",
    "Paris":      "48.815,2.224,48.902,2.470",
    "London":     "51.286,-0.510,51.692,0.334",
    "New York":   "40.477,-74.260,40.918,-73.700",
    "Madrid":     "40.312,-3.889,40.644,-3.524",
    "Rome":       "41.794,12.341,41.987,12.616",
    "Amsterdam":  "52.279,4.729,52.431,5.079",
    "Barcelona":  "41.320,2.069,41.469,2.228",
    "Dubai":      "25.002,55.000,25.357,55.566",
}

# ---------------------------------------------------------------------------
# Geo helpers
# ---------------------------------------------------------------------------

def generate_cells(min_lat, min_lon, max_lat, max_lon, cell_km=GRID_CELL_SIZE_KM):
    lat_step = cell_km / KM_PER_DEG_LAT
    mid_lat  = (min_lat + max_lat) / 2
    cos_mid  = math.cos(math.radians(mid_lat)) or 1e-6
    lon_step = cell_km / (KM_PER_DEG_LAT * cos_mid)
    cells = []
    lat = min_lat + lat_step / 2
    while lat < max_lat:
        lon = min_lon + lon_step / 2
        while lon < max_lon:
            cells.append((lat, lon))
            lon += lon_step
        lat += lat_step
    return cells

def parse_bbox(s):
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError("Bounding box must be: minLat,minLon,maxLat,maxLon")
    vals = [float(p) for p in parts]
    if vals[0] >= vals[2]:
        raise ValueError("minLat must be less than maxLat")
    if vals[1] >= vals[3]:
        raise ValueError("minLon must be less than maxLon")
    return vals

# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class GMapsClient:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Scraper")
        self.root.geometry("1060x840")
        self.root.minsize(860, 680)
        self.root.configure(bg=C_BG)

        self.config          = self.load_config()
        self.is_running      = False
        self._spinner_idx    = 0
        self._spinner_active = False
        self._start_time     = None

        self._build_styles()
        self._build_ui()

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

    def _build_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TNotebook",
                        background=C_PANEL, borderwidth=0,
                        tabmargins=[0, 0, 0, 0])
        style.configure("TNotebook.Tab",
                        background=C_BG, foreground=C_MUTED,
                        padding=[16, 8], font=("Segoe UI", 9),
                        borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", C_PANEL)],
                  foreground=[("selected", C_ACCENT)])

        style.configure("Light.TCombobox",
                        fieldbackground=C_INPUT_BG, background=C_INPUT_BG,
                        foreground=C_TEXT, selectbackground=C_ACCENT,
                        arrowcolor=C_MUTED, bordercolor=C_BORDER)
        style.map("Light.TCombobox",
                  fieldbackground=[("readonly", C_INPUT_BG)],
                  foreground=[("readonly", C_TEXT)])

        style.configure("Blue.Horizontal.TProgressbar",
                        troughcolor=C_BORDER, background=C_ACCENT,
                        thickness=10, borderwidth=0)

        style.configure("TSeparator", background=C_BORDER)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=C_PANEL,
                       highlightthickness=1, highlightbackground=C_BORDER,
                       height=64)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="Google Maps Scraper",
                 font=("Segoe UI", 15, "bold"),
                 bg=C_PANEL, fg=C_TEXT).pack(side=tk.LEFT, padx=24)

        self._settings_btn = tk.Button(
            hdr, text="Settings",
            font=("Segoe UI", 9), bg=C_PANEL, fg=C_MUTED,
            activebackground=C_BG, activeforeground=C_ACCENT,
            relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
            bd=0, highlightthickness=1, highlightbackground=C_BORDER,
            command=self.open_settings)
        self._settings_btn.pack(side=tk.RIGHT, padx=20, pady=14)

        # ── Page scroll canvas (keeps padding consistent) ───────────────
        page = tk.Frame(self.root, bg=C_BG)
        page.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ── Search card ─────────────────────────────────────────────────
        search_card = self._card(page)
        search_card.pack(fill=tk.X, pady=(0, 12))

        tk.Label(search_card, text="Search Query",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(anchor=tk.W, padx=20, pady=(18, 6))

        input_row = tk.Frame(search_card, bg=C_PANEL)
        input_row.pack(fill=tk.X, padx=20, pady=(0, 18))

        self.query_input = tk.Entry(
            input_row,
            font=("Segoe UI", 12), bg=C_INPUT_BG, fg=C_TEXT,
            insertbackground=C_TEXT, relief=tk.FLAT,
            highlightthickness=1, highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT)
        self.query_input.pack(side=tk.LEFT, fill=tk.X, expand=True,
                              ipady=10, padx=(0, 12))
        self.query_input.insert(0, "dentists in berlin")
        self.query_input.bind("<Return>", lambda _: self.start_run())

        self.run_btn = tk.Button(
            input_row, text="Run",
            font=("Segoe UI", 10, "bold"),
            bg=C_ACCENT, fg=C_WHITE,
            activebackground=C_ACCENT_H, activeforeground=C_WHITE,
            relief=tk.FLAT, padx=28, pady=10, cursor="hand2",
            bd=0)
        self.run_btn.config(command=self.start_run)
        self.run_btn.pack(side=tk.RIGHT)

        # ── Options card ────────────────────────────────────────────────
        opts_card = self._card(page)
        opts_card.pack(fill=tk.X, pady=(0, 12))

        self.notebook = ttk.Notebook(opts_card)
        self.notebook.pack(fill=tk.X, padx=0, pady=0)

        self._build_simple_tab()

        # ── Status card ─────────────────────────────────────────────────
        status_card = self._card(page)
        status_card.pack(fill=tk.X, pady=(0, 12))

        # Row 1: indicator dot + status text + elapsed
        top_row = tk.Frame(status_card, bg=C_PANEL)
        top_row.pack(fill=tk.X, padx=20, pady=(18, 10))

        self._status_dot = tk.Label(top_row, text="●",
                                    font=("Segoe UI", 16),
                                    bg=C_PANEL, fg=C_HINT)
        self._status_dot.pack(side=tk.LEFT)

        self._status_label = tk.Label(top_row,
                                      text="   Ready to scrape",
                                      font=("Segoe UI", 12, "bold"),
                                      bg=C_PANEL, fg=C_TEXT)
        self._status_label.pack(side=tk.LEFT)

        self._elapsed_label = tk.Label(top_row, text="",
                                       font=("Segoe UI", 9),
                                       bg=C_PANEL, fg=C_HINT)
        self._elapsed_label.pack(side=tk.RIGHT)

        # Row 2: phase stepper
        stepper_row = tk.Frame(status_card, bg=C_PANEL)
        stepper_row.pack(fill=tk.X, padx=20, pady=(0, 10))

        self._phase_frames = []
        _phase_names = ["Submit", "Scraping", "Export", "Done"]
        for i, name in enumerate(_phase_names):
            if i > 0:
                tk.Label(stepper_row, text="  →  ",
                         font=("Segoe UI", 9), bg=C_PANEL,
                         fg=C_HINT).pack(side=tk.LEFT)
            badge = tk.Label(stepper_row, text=f"  {name}  ",
                             font=("Segoe UI", 8, "bold"),
                             bg=C_BORDER, fg=C_MUTED,
                             padx=2, pady=4, relief=tk.FLAT)
            badge.pack(side=tk.LEFT)
            self._phase_frames.append(badge)

        # Row 3: result counter
        counter_row = tk.Frame(status_card, bg=C_PANEL)
        counter_row.pack(fill=tk.X, padx=20, pady=(0, 8))

        self._result_label = tk.Label(counter_row, text="",
                                      font=("Segoe UI", 10),
                                      bg=C_PANEL, fg=C_MUTED)
        self._result_label.pack(side=tk.LEFT)

        # Row 4: progress bar
        self._progress_var = tk.DoubleVar()
        self._progress_bar = ttk.Progressbar(
            status_card, variable=self._progress_var, maximum=100,
            style="Blue.Horizontal.TProgressbar")
        self._progress_bar.pack(fill=tk.X, padx=20, pady=(0, 18))

        # ── Log card ────────────────────────────────────────────────────
        log_card = self._card(page)
        log_card.pack(fill=tk.BOTH, expand=True)

        log_hdr = tk.Frame(log_card, bg=C_PANEL)
        log_hdr.pack(fill=tk.X, padx=20, pady=(14, 0))

        tk.Label(log_hdr, text="Activity Log",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT)

        self._clear_btn = tk.Button(
            log_hdr, text="Clear",
            font=("Segoe UI", 8), bg=C_PANEL, fg=C_HINT,
            activebackground=C_BG, activeforeground=C_MUTED,
            relief=tk.FLAT, padx=8, cursor="hand2",
            command=self._clear_logs)
        self._clear_btn.pack(side=tk.RIGHT)

        log_inner = tk.Frame(log_card, bg=C_INPUT_BG,
                             highlightthickness=1,
                             highlightbackground=C_BORDER)
        log_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=(8, 18))

        self.logs = tk.Text(
            log_inner,
            font=("Consolas", 9), bg=C_INPUT_BG, fg=C_TEXT,
            insertbackground=C_TEXT, relief=tk.FLAT,
            wrap=tk.WORD, state=tk.DISABLED,
            padx=12, pady=10)
        self.logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(log_inner, command=self.logs.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.logs.configure(yscrollcommand=sb.set)

        self.logs.tag_config("ts",    foreground=C_HINT)
        self.logs.tag_config("info",  foreground=C_TEXT)
        self.logs.tag_config("good",  foreground=C_GREEN)
        self.logs.tag_config("warn",  foreground=C_ORANGE)
        self.logs.tag_config("error", foreground=C_RED)
        self.logs.tag_config("dim",   foreground=C_HINT)

    def _card(self, parent):
        return tk.Frame(parent, bg=C_PANEL,
                        highlightthickness=1, highlightbackground=C_BORDER)

    def _build_simple_tab(self):
        f = tk.Frame(self.notebook, bg=C_PANEL, padx=20, pady=14)
        self.notebook.add(f, text="  Simple  ")

        row0 = tk.Frame(f, bg=C_PANEL)
        row0.pack(fill=tk.X, pady=(0, 10))
        tk.Label(row0, text="Max Depth",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT)
        self.depth_var = tk.StringVar(value="10")
        tk.Entry(row0, textvariable=self.depth_var, width=6,
                 font=("Segoe UI", 9), bg=C_INPUT_BG, fg=C_TEXT,
                 insertbackground=C_TEXT, relief=tk.FLAT,
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).pack(side=tk.LEFT, padx=10, ipady=4)
        tk.Label(row0, text="pages  ·  ~20 results each  ·  max 100",
                 font=("Segoe UI", 8), bg=C_PANEL, fg=C_HINT).pack(side=tk.LEFT)

        row1 = tk.Frame(f, bg=C_PANEL)
        row1.pack(fill=tk.X)
        tk.Label(row1, text="Fields",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT)
        self.fields_var = tk.StringVar(value="lite")
        self._fields_combo = ttk.Combobox(
            row1, textvariable=self.fields_var,
            values=["lite", "all",
                    "title,phone,web_site,address",
                    "title,phone,web_site,address,review_rating,review_count,latitude,longitude"],
            width=52, style="Light.TCombobox", font=("Segoe UI", 9))
        self._fields_combo.pack(side=tk.LEFT, padx=10, ipady=4)
        tk.Label(row1,
                 text="lite = category · name · phone · website",
                 font=("Segoe UI", 8), bg=C_PANEL, fg=C_HINT).pack(side=tk.LEFT)

    def _build_grid_tab(self):
        f = tk.Frame(self.notebook, bg=C_PANEL, padx=20, pady=14)
        self.notebook.add(f, text="  Grid — more results  ")

        row0 = tk.Frame(f, bg=C_PANEL)
        row0.pack(fill=tk.X, pady=(0, 10))

        tk.Label(row0, text="City Preset",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT)
        self.city_var = tk.StringVar(value="Custom...")
        city_cb = ttk.Combobox(row0, textvariable=self.city_var,
                               values=list(CITY_PRESETS.keys()),
                               state="readonly", width=16,
                               style="Light.TCombobox", font=("Segoe UI", 9))
        city_cb.pack(side=tk.LEFT, padx=10, ipady=4)
        city_cb.bind("<<ComboboxSelected>>", self._on_city_select)

        tk.Label(row0, text="Bounding Box",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT, padx=(20, 0))
        self.bbox_var = tk.StringVar()
        tk.Entry(row0, textvariable=self.bbox_var, width=34,
                 font=("Segoe UI", 9), bg=C_INPUT_BG, fg=C_TEXT,
                 insertbackground=C_TEXT, relief=tk.FLAT,
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).pack(side=tk.LEFT, padx=10, ipady=4)
        tk.Label(row0, text="minLat,minLon,maxLat,maxLon",
                 font=("Segoe UI", 8), bg=C_PANEL, fg=C_HINT).pack(side=tk.LEFT)

        row1 = tk.Frame(f, bg=C_PANEL)
        row1.pack(fill=tk.X)

        tk.Label(row1, text="Cell Size (km)",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT)
        self.cell_km_var = tk.StringVar(value="3")
        tk.Entry(row1, textvariable=self.cell_km_var, width=6,
                 font=("Segoe UI", 9), bg=C_INPUT_BG, fg=C_TEXT,
                 insertbackground=C_TEXT, relief=tk.FLAT,
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).pack(side=tk.LEFT, padx=10, ipady=4)

        tk.Label(row1, text="Zoom",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(side=tk.LEFT, padx=(20, 0))
        self.zoom_var = tk.StringVar(value=str(GRID_ZOOM))
        tk.Entry(row1, textvariable=self.zoom_var, width=5,
                 font=("Segoe UI", 9), bg=C_INPUT_BG, fg=C_TEXT,
                 insertbackground=C_TEXT, relief=tk.FLAT,
                 highlightthickness=1, highlightbackground=C_BORDER
                 ).pack(side=tk.LEFT, padx=10, ipady=4)

        self._cell_count_label = tk.Label(row1, text="",
                                          font=("Segoe UI", 9),
                                          bg=C_PANEL, fg=C_ACCENT)
        self._cell_count_label.pack(side=tk.LEFT, padx=16)

        self.bbox_var.trace_add("write", self._update_cell_count)
        self.cell_km_var.trace_add("write", self._update_cell_count)

    # ------------------------------------------------------------------
    # Status helpers
    # ------------------------------------------------------------------

    _SPINNER     = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
    _PHASE_NAMES = ["Submit", "Scraping", "Export", "Done"]

    def _set_phase(self, phase_idx):
        for i, badge in enumerate(self._phase_frames):
            if i < phase_idx:
                badge.config(bg=C_GREEN_BG, fg=C_GREEN, text=f"  ✓ {self._PHASE_NAMES[i]}  ")
            elif i == phase_idx:
                badge.config(bg=C_ACCENT, fg=C_WHITE, text=f"  {self._PHASE_NAMES[i]}  ")
            else:
                badge.config(bg=C_BORDER, fg=C_MUTED, text=f"  {self._PHASE_NAMES[i]}  ")

    def _reset_phases(self):
        for i, badge in enumerate(self._phase_frames):
            badge.config(bg=C_BORDER, fg=C_MUTED, text=f"  {self._PHASE_NAMES[i]}  ")

    def _set_state(self, state, msg=""):
        dot_colors = {"idle": C_HINT, "running": C_ACCENT,
                      "done": C_GREEN, "error": C_RED}
        labels = {
            "idle":    "   Ready to scrape",
            "running": f"   {msg}" if msg else "   Working…",
            "done":    f"   {msg}" if msg else "   Done",
            "error":   f"   {msg}" if msg else "   Failed",
        }
        self._status_dot.config(fg=dot_colors.get(state, C_HINT))
        self._status_label.config(
            text=labels.get(state, f"   {msg}"),
            fg=C_TEXT if state != "error" else C_RED)

        if state == "running":
            self._start_time = time.time()
            self._spinner_active = True
            self._tick_spinner()
            self._tick_elapsed()
            self._progress_bar.config(mode="indeterminate")
            self._progress_bar.start(40)
        else:
            self._spinner_active = False
            self._progress_bar.stop()
            self._progress_bar.config(mode="determinate")
            if state == "done":
                self._progress_var.set(100)
                self._set_phase(4)
            elif state == "idle":
                self._reset_phases()
                self._result_label.config(text="")
            self._elapsed_label.config(text="")

    def _tick_spinner(self):
        if not self._spinner_active:
            return
        self._spinner_idx = (self._spinner_idx + 1) % len(self._SPINNER)
        self._status_dot.config(text=self._SPINNER[self._spinner_idx])
        self.root.after(100, self._tick_spinner)

    def _tick_elapsed(self):
        if not self._spinner_active or self._start_time is None:
            return
        elapsed = int(time.time() - self._start_time)
        m, s = divmod(elapsed, 60)
        self._elapsed_label.config(text=f"{m:02d}:{s:02d}")
        self.root.after(1000, self._tick_elapsed)

    def _set_result_count(self, n):
        if n > 0:
            self._result_label.config(
                text=f"  {n:,} results found so far",
                fg=C_GREEN, font=("Segoe UI", 10, "bold"))
        else:
            self._result_label.config(text="")

    def _set_grid_progress(self, done, total):
        pct = done / total * 100 if total else 0
        self._progress_bar.stop()
        self._progress_bar.config(mode="determinate")
        self._progress_var.set(pct)
        self._status_label.config(
            text=f"   Scanning grid — cell {done} of {total}  ({pct:.0f}%)")

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def log(self, msg, level="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.logs.config(state=tk.NORMAL)
        self.logs.insert(tk.END, f"[{ts}]  ", "ts")
        self.logs.insert(tk.END, f"{msg}\n", level)
        self.logs.see(tk.END)
        self.logs.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _clear_logs(self):
        self.logs.config(state=tk.NORMAL)
        self.logs.delete(1.0, tk.END)
        self.logs.config(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Run orchestration
    # ------------------------------------------------------------------

    def start_run(self):
        query = self.query_input.get().strip()
        if not query:
            self._flash_error("Please enter a search query")
            return
        if self.is_running:
            return

        self.is_running = True
        self.run_btn.config(state=tk.DISABLED, text="Running…",
                            bg=C_HINT, cursor="arrow")
        self._clear_logs()
        self._set_result_count(0)
        self._reset_phases()
        self._set_state("running", "Submitting job…")

        threading.Thread(target=self._run_simple, args=(query,), daemon=True).start()

    def _finish(self, success, msg, result_count=0):
        self.is_running = False
        self.run_btn.config(state=tk.NORMAL, text="Run",
                            bg=C_ACCENT, cursor="hand2")
        if success:
            self._set_state("done", f"Done — {result_count:,} results exported")
            self._set_result_count(result_count)
        else:
            self._set_state("error", msg)

    # -- Simple --

    def _run_simple(self, query):
        try:
            depth  = int(self.depth_var.get()) if self.depth_var.get().isdigit() else 10
            fields = self._resolve_fields()
            self.log(f"Query: {query}   depth: {depth}   fields: {fields or 'all'}", "dim")
            self.log(f"API: {self.config['api_url']}", "dim")

            self._set_phase(0)
            self._set_state("running", "Submitting job to API…")
            job_id = self.submit_job(query, max_depth=depth, fields=fields)
            self.log(f"Job created  →  {job_id}", "dim")

            self._set_phase(1)
            self._set_state("running", "Scraping in progress — waiting for results…")
            result = self.wait_for_job(job_id, self._on_poll)
            rows   = self.parse_results(result.get("results"))

            if not rows:
                raise Exception("API returned 0 results")

            # Client-side field filter — ensures only requested columns
            # are written even if the server returns the full payload.
            rows = self._filter_rows(rows, fields)

            self._set_phase(2)
            self._set_state("running", f"Saving {len(rows):,} results to Excel…")
            out = self.export_xlsx(query, rows)
            self.log(f"Saved  →  {out}", "good")
            self._finish(True, "", len(rows))

        except Exception as e:
            self.log(f"Error: {e}", "error")
            self._finish(False, str(e)[:80])

    def _on_poll(self, status_obj):
        s = status_obj.get("status", "")
        n = status_obj.get("result_count", 0)
        if s == "running":
            self._set_state("running", "Scraping in progress — waiting for results…")
            if n:
                self._set_result_count(n)
        self.log(f"Poll  status={s}  results={n}", "dim")

    # -- Grid --

    def _run_grid(self, query):
        bbox_s = self.bbox_var.get().strip()
        if not bbox_s:
            self.log("Bounding box is required for Grid mode", "error")
            self._finish(False, "Bounding box required")
            return
        try:
            vals = parse_bbox(bbox_s)
        except ValueError as e:
            self.log(str(e), "error")
            self._finish(False, str(e))
            return

        try:
            cell_km = float(self.cell_km_var.get())
        except ValueError:
            cell_km = GRID_CELL_SIZE_KM
        try:
            zoom = int(self.zoom_var.get())
        except ValueError:
            zoom = GRID_ZOOM

        cells  = generate_cells(*vals, cell_km=cell_km)
        total  = len(cells)
        fields = self._resolve_fields()

        self.log(f"Grid: {total} cells × ~20 results   zoom={zoom}   fields={fields or 'all'}", "dim")
        self.log(f"API: {self.config['api_url']}", "dim")

        self._set_phase(1)

        all_rows  = []
        seen_ids  = set()
        completed = 0
        failed    = 0

        try:
            with ThreadPoolExecutor(max_workers=GRID_MAX_PARALLEL) as pool:
                futures = {
                    pool.submit(self._run_cell, query, lat, lon, zoom, fields): (lat, lon)
                    for lat, lon in cells
                }
                for future in as_completed(futures):
                    lat, lon = futures[future]
                    completed += 1
                    self._set_grid_progress(completed, total)
                    try:
                        rows = future.result()
                        new  = 0
                        for row in rows:
                            uid = (row.get("place_id")
                                   or f"{row.get('title','')}|{row.get('address','')}")
                            if uid not in seen_ids:
                                seen_ids.add(uid)
                                all_rows.append(row)
                                new += 1
                        if new:
                            self.log(
                                f"Cell ({lat:.3f}, {lon:.3f})  +{new} new"
                                f"  total {len(all_rows):,}", "info")
                            self._set_result_count(len(all_rows))
                    except Exception as e:
                        failed += 1
                        self.log(f"Cell ({lat:.3f}, {lon:.3f}) failed: {e}", "warn")

            if not all_rows:
                raise Exception("No results from any grid cell")

            self._set_phase(2)
            self._set_state("running", f"Saving {len(all_rows):,} results to Excel…")
            out = self.export_xlsx(query, all_rows)
            self.log(f"Saved  →  {out}", "good")
            if failed:
                self.log(f"{failed} cells failed (results still exported)", "warn")
            self._finish(True, "", len(all_rows))

        except Exception as e:
            self.log(f"Error: {e}", "error")
            self._finish(False, str(e)[:80])

    def _run_cell(self, query, lat, lon, zoom, fields=None):
        job_id = self.submit_job(
            query, max_depth=3,
            geo_coordinates=f"{lat},{lon}",
            zoom=zoom, fast_mode=True, fields=fields)
        result = self.wait_for_job(job_id, logger=None)
        return self.parse_results(result.get("results"))

    # ------------------------------------------------------------------
    # API
    # ------------------------------------------------------------------

    LITE_FIELDS = ["title", "category", "phone", "web_site"]

    def _resolve_fields(self):
        raw = self.fields_var.get().strip().lower()
        if not raw or raw == "all":
            return []
        if raw == "lite":
            return ["lite"]
        return [f.strip() for f in raw.split(",") if f.strip()]

    def _filter_rows(self, rows, fields):
        """Keep only requested columns in each row (client-side enforcement)."""
        if not fields:
            return rows
        allowed = set(self.LITE_FIELDS if fields == ["lite"] else fields)
        return [{k: v for k, v in row.items() if k in allowed} for row in rows]

    def submit_job(self, query, max_depth=10, geo_coordinates=None,
                   zoom=None, fast_mode=False, fields=None):
        payload = {"keyword": query, "lang": "en",
                   "max_depth": max_depth, "timeout": 300}
        if geo_coordinates:
            payload["geo_coordinates"] = geo_coordinates
        if zoom:
            payload["zoom"] = zoom
        if fast_mode:
            payload["fast_mode"] = True
        if fields:
            payload["fields"] = fields

        resp = requests.post(
            f"{self.config['api_url'].rstrip('/')}/api/v1/scrape",
            json=payload, timeout=30)
        if not (200 <= resp.status_code < 300):
            raise Exception(f"Submit failed: HTTP {resp.status_code}")
        data = resp.json()
        if not data.get("job_id"):
            raise Exception("Empty job_id from API")
        return data["job_id"]

    def wait_for_job(self, job_id, logger=None):
        deadline = time.time() + OVERALL_TIMEOUT
        url = f"{self.config['api_url'].rstrip('/')}/api/v1/jobs/{job_id}"
        while time.time() < deadline:
            try:
                resp = requests.get(url, timeout=30)
                if 200 <= resp.status_code < 300:
                    status = resp.json()
                    if logger:
                        logger(status)
                    s = status.get("status", "").lower()
                    if s == "completed":
                        return status
                    if s == "failed":
                        raise Exception(status.get("error") or "Job failed")
            except requests.RequestException as e:
                if logger:
                    self.log(f"Poll error: {e}", "warn")
            time.sleep(POLL_INTERVAL)
        raise Exception("Job timed out")

    def parse_results(self, raw):
        if not raw:
            return []
        if isinstance(raw, list):
            return raw
        if isinstance(raw, str):
            data = json.loads(raw)
            return data if isinstance(data, list) else []
        return []

    # ------------------------------------------------------------------
    # XLSX export
    # ------------------------------------------------------------------

    def export_xlsx(self, query, rows):
        if not rows:
            raise Exception("No rows to export")
        priority = ["title", "category", "address", "phone", "web_site",
                    "review_rating", "review_count", "latitude", "longitude"]
        all_keys = set()
        for row in rows:
            all_keys.update(row.keys())
        headers = [k for k in priority if k in all_keys]
        headers += sorted(all_keys - set(headers))

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(headers)
        for row in rows:
            ws.append([self._cell_val(row.get(h)) for h in headers])
        ws.freeze_panes = "A2"
        for col in ws.columns:
            mx = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(mx + 2, 50)

        safe = re.sub(r"[^a-z0-9_\-]", "", query.lower().replace(" ", "_"))[:60]
        out  = Path.home() / "Downloads" / f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out.parent.mkdir(exist_ok=True)
        wb.save(out)
        return str(out)

    def _cell_val(self, v):
        if v is None:
            return ""
        if isinstance(v, (str, int, float, bool)):
            return v
        return json.dumps(v, ensure_ascii=False)

    # ------------------------------------------------------------------
    # Settings dialog
    # ------------------------------------------------------------------

    def open_settings(self):
        d = tk.Toplevel(self.root)
        d.title("Settings")
        d.geometry("560x180")
        d.configure(bg=C_BG)
        d.transient(self.root)
        d.resizable(False, False)

        inner = tk.Frame(d, bg=C_PANEL,
                         highlightthickness=1, highlightbackground=C_BORDER)
        inner.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        tk.Label(inner, text="API URL",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_PANEL, fg=C_MUTED).pack(anchor=tk.W, padx=16, pady=(16, 6))

        entry = tk.Entry(inner, font=("Segoe UI", 11),
                         bg=C_INPUT_BG, fg=C_TEXT,
                         insertbackground=C_TEXT, relief=tk.FLAT,
                         highlightthickness=1, highlightbackground=C_BORDER,
                         highlightcolor=C_ACCENT)
        entry.pack(fill=tk.X, padx=16, ipady=8)
        entry.insert(0, self.config["api_url"])

        def save():
            url = entry.get().strip().rstrip("/")
            if not url.startswith(("http://", "https://")):
                self._flash_error("URL must start with http:// or https://")
                return
            self.config["api_url"] = url
            self.save_config()
            self.log(f"API URL saved: {url}", "dim")
            d.destroy()

        tk.Button(inner, text="Save", command=save,
                  font=("Segoe UI", 10, "bold"),
                  bg=C_ACCENT, fg=C_WHITE, relief=tk.FLAT,
                  padx=24, pady=8, cursor="hand2",
                  activebackground=C_ACCENT_H, activeforeground=C_WHITE
                  ).pack(pady=14)

    # ------------------------------------------------------------------
    # Misc helpers
    # ------------------------------------------------------------------

    def _on_city_select(self, _event=None):
        self.bbox_var.set(CITY_PRESETS.get(self.city_var.get(), ""))

    def _update_cell_count(self, *_):
        try:
            vals = parse_bbox(self.bbox_var.get().strip())
            n    = len(generate_cells(*vals, cell_km=float(self.cell_km_var.get())))
            self._cell_count_label.config(
                text=f"→  {n} cells  ≈  {n * 20:,}+ results")
        except Exception:
            self._cell_count_label.config(text="")

    def _flash_error(self, msg):
        self.log(msg, "error")
        self._set_state("error", msg[:60])

    def load_config(self):
        if CONFIG_FILE.exists():
            try:
                return json.loads(CONFIG_FILE.read_text())
            except Exception:
                pass
        return {"api_url": DEFAULT_API_URL}

    def save_config(self):
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(self.config, indent=2))


if __name__ == "__main__":
    root = tk.Tk()
    GMapsClient(root)
    root.mainloop()
