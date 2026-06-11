#!/usr/bin/env python3
"""
Google Maps Scraper - Desktop Client
Query Google Maps, download results as XLSX
Supports grid mode for large-area searches that break Google's per-search limit.
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
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

# Constants
DEFAULT_API_URL = "https://gmaps-api.fouiguira.com"
POLL_INTERVAL = 3
OVERALL_TIMEOUT = 20 * 60
CONFIG_FILE = Path.home() / ".gmaps-client" / "config.json"
GRID_CELL_SIZE_KM = 3.0
GRID_ZOOM = 14
GRID_MAX_PARALLEL = 5  # concurrent jobs against the API


# ---------------------------------------------------------------------------
# Geo grid helpers (mirrors the Go grid package)
# ---------------------------------------------------------------------------

KM_PER_DEG_LAT = 111.32


def generate_cells(min_lat, min_lon, max_lat, max_lon, cell_km=GRID_CELL_SIZE_KM):
    lat_step = cell_km / KM_PER_DEG_LAT
    mid_lat = (min_lat + max_lat) / 2
    cos_mid = math.cos(math.radians(mid_lat)) or 1e-6
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
    """Parse 'minLat,minLon,maxLat,maxLon' string."""
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError("Bounding box must be: minLat,minLon,maxLat,maxLon")
    vals = [float(p) for p in parts]
    if vals[0] >= vals[2]:
        raise ValueError("minLat must be less than maxLat")
    if vals[1] >= vals[3]:
        raise ValueError("minLon must be less than maxLon")
    return vals  # [min_lat, min_lon, max_lat, max_lon]


# ---------------------------------------------------------------------------
# Preset bounding boxes for common cities
# ---------------------------------------------------------------------------
CITY_PRESETS = {
    "Custom...": "",
    "Berlin": "52.338,13.088,52.675,13.761",
    "Paris": "48.815,2.224,48.902,2.470",
    "London": "51.286,-0.510,51.692,0.334",
    "New York": "40.477,-74.260,40.918,-73.700",
    "Madrid": "40.312,-3.889,40.644,-3.524",
    "Rome": "41.794,12.341,41.987,12.616",
    "Amsterdam": "52.279,4.729,52.431,5.079",
    "Barcelona": "41.320,2.069,41.469,2.228",
    "Dubai": "25.002,55.000,25.357,55.566",
}


class GMapsClient:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Scraper Client")
        self.root.geometry("960x720")
        self.root.resizable(True, True)

        self.config = self.load_config()
        self.is_running = False

        self.build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def build_ui(self):
        # Top bar
        top_frame = tk.Frame(self.root, bg="#f0f0f0")
        top_frame.pack(fill=tk.X, padx=10, pady=8)
        tk.Label(top_frame, text="Google Maps API Client", font=("Arial", 14, "bold"), bg="#f0f0f0").pack(side=tk.LEFT)
        tk.Button(top_frame, text="⚙ Settings", command=self.open_settings).pack(side=tk.RIGHT)

        # Query
        tk.Label(self.root, text="Search Query", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10)
        self.query_input = tk.Entry(self.root, font=("Arial", 11))
        self.query_input.pack(fill=tk.X, padx=10, pady=(0, 6))
        self.query_input.insert(0, "dentists in berlin")

        # Mode tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.X, padx=10, pady=(0, 6))

        # -- Simple mode tab --
        simple_frame = tk.Frame(self.notebook, padx=8, pady=6)
        self.notebook.add(simple_frame, text="  Simple  ")
        tk.Label(simple_frame, text="Max Depth  (pages, ~20 results each, max 100)", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W)
        self.depth_var = tk.StringVar(value="10")
        tk.Entry(simple_frame, textvariable=self.depth_var, width=8, font=("Arial", 10)).grid(row=0, column=1, padx=8)
        tk.Label(simple_frame, text="≈ up to 200 results", font=("Arial", 9), fg="#666").grid(row=0, column=2, sticky=tk.W)

        # Fields selector
        tk.Label(simple_frame, text="Fields:", font=("Arial", 9)).grid(row=1, column=0, sticky=tk.W, pady=(6, 0))
        self.fields_var = tk.StringVar(value="lite")
        fields_menu = ttk.Combobox(simple_frame, textvariable=self.fields_var,
                                   values=["lite", "all", "title,phone,web_site,address",
                                           "title,phone,web_site,address,review_rating,review_count,latitude,longitude"],
                                   width=52, font=("Arial", 9))
        fields_menu.grid(row=1, column=1, columnspan=3, padx=8, sticky=tk.W, pady=(6, 0))
        tk.Label(simple_frame, text="  lite = name, phone, website, address, rating  |  all = every field",
                 font=("Arial", 8), fg="#888").grid(row=2, column=0, columnspan=4, sticky=tk.W)

        # -- Grid mode tab --
        grid_frame = tk.Frame(self.notebook, padx=8, pady=6)
        self.notebook.add(grid_frame, text="  Grid (more results)  ")

        tk.Label(grid_frame, text="City preset:", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W)
        self.city_var = tk.StringVar(value="Custom...")
        city_menu = ttk.Combobox(grid_frame, textvariable=self.city_var, values=list(CITY_PRESETS.keys()), state="readonly", width=16)
        city_menu.grid(row=0, column=1, padx=6, sticky=tk.W)
        city_menu.bind("<<ComboboxSelected>>", self._on_city_select)

        tk.Label(grid_frame, text="Bounding box  (minLat,minLon,maxLat,maxLon):", font=("Arial", 9)).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        self.bbox_var = tk.StringVar()
        tk.Entry(grid_frame, textvariable=self.bbox_var, width=36, font=("Arial", 10)).grid(row=1, column=1, columnspan=3, padx=6, sticky=tk.W, pady=(4, 0))

        tk.Label(grid_frame, text="Cell size (km):", font=("Arial", 9)).grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        self.cell_km_var = tk.StringVar(value="3")
        tk.Entry(grid_frame, textvariable=self.cell_km_var, width=8, font=("Arial", 10)).grid(row=2, column=1, padx=6, sticky=tk.W, pady=(4, 0))
        self.cell_count_label = tk.Label(grid_frame, text="", font=("Arial", 9), fg="#666")
        self.cell_count_label.grid(row=2, column=2, sticky=tk.W)
        self.bbox_var.trace_add("write", self._update_cell_count)
        self.cell_km_var.trace_add("write", self._update_cell_count)

        tk.Label(grid_frame, text="Zoom (1-21):", font=("Arial", 9)).grid(row=3, column=0, sticky=tk.W, pady=(4, 0))
        self.zoom_var = tk.StringVar(value=str(GRID_ZOOM))
        tk.Entry(grid_frame, textvariable=self.zoom_var, width=8, font=("Arial", 10)).grid(row=3, column=1, padx=6, sticky=tk.W, pady=(4, 0))

        # Run button
        self.run_btn = tk.Button(self.root, text="Run Query And Export XLSX",
                                 command=self.start_run,
                                 font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", height=2)
        self.run_btn.pack(fill=tk.X, padx=10, pady=6)

        # Progress bar (grid mode)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=10)

        # Status
        self.status_label = tk.Label(self.root, text="Ready", font=("Arial", 10), fg="#333")
        self.status_label.pack(anchor=tk.W, padx=10, pady=4)

        # Separator + logs
        tk.Frame(self.root, height=1, bg="#ddd").pack(fill=tk.X, pady=2)
        tk.Label(self.root, text="Logs", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10)
        self.logs = scrolledtext.ScrolledText(self.root, height=16, font=("Courier", 9))
        self.logs.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))
        self.logs.config(state=tk.DISABLED)

    def _on_city_select(self, _event=None):
        city = self.city_var.get()
        bbox = CITY_PRESETS.get(city, "")
        self.bbox_var.set(bbox)

    def _update_cell_count(self, *_):
        bbox_s = self.bbox_var.get().strip()
        cell_s = self.cell_km_var.get().strip()
        if not bbox_s or not cell_s:
            self.cell_count_label.config(text="")
            return
        try:
            vals = parse_bbox(bbox_s)
            cell_km = float(cell_s)
            cells = generate_cells(*vals, cell_km=cell_km)
            self.cell_count_label.config(text=f"→ {len(cells)} grid cells, ~{len(cells) * 20}+ results")
        except Exception:
            self.cell_count_label.config(text="")

    # ------------------------------------------------------------------
    # Run logic
    # ------------------------------------------------------------------

    def start_run(self):
        query = self.query_input.get().strip()
        if not query:
            messagebox.showerror("Error", "Query is required")
            return

        self.is_running = True
        self.run_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.logs.config(state=tk.NORMAL)
        self.logs.delete(1.0, tk.END)
        self.logs.config(state=tk.DISABLED)

        tab = self.notebook.index(self.notebook.select())
        if tab == 1:  # grid mode
            thread = threading.Thread(target=self._run_grid, args=(query,), daemon=True)
        else:
            thread = threading.Thread(target=self._run_simple, args=(query,), daemon=True)
        thread.start()

    # -- Simple mode --

    def _run_simple(self, query):
        try:
            depth = int(self.depth_var.get())
        except ValueError:
            depth = 10

        try:
            self.log(f"Using API URL: {self.config['api_url']}")
            self.log(f"Query: {query}  |  max_depth: {depth}")
            self.set_status("Submitting job...")

            fields = self._resolve_fields()
            job_id = self.submit_job(query, max_depth=depth, fields=fields)
            self.log(f"Job submitted: {job_id}  fields={fields or 'all'}")
            self.set_status("Job running...")

            result = self.wait_for_job(job_id, lambda m: self.log(m))
            rows = self.parse_results(result.get("results"))

            if not rows:
                self.fail("No results returned")
                return

            out = self.export_xlsx(query, rows)
            self.log(f"XLSX saved: {out}")
            self.set_status(f"Done — {len(rows)} results")
            self.progress_var.set(100)
            messagebox.showinfo("Success", f"{len(rows)} results saved:\n{out}")

        except Exception as e:
            self.fail(str(e))
        finally:
            self.is_running = False
            self.run_btn.config(state=tk.NORMAL)

    # -- Grid mode --

    def _run_grid(self, query):
        bbox_s = self.bbox_var.get().strip()
        if not bbox_s:
            self.fail("Bounding box is required for grid mode")
            self.is_running = False
            self.run_btn.config(state=tk.NORMAL)
            return

        try:
            vals = parse_bbox(bbox_s)
        except ValueError as e:
            self.fail(str(e))
            self.is_running = False
            self.run_btn.config(state=tk.NORMAL)
            return

        try:
            cell_km = float(self.cell_km_var.get())
        except ValueError:
            cell_km = GRID_CELL_SIZE_KM

        try:
            zoom = int(self.zoom_var.get())
        except ValueError:
            zoom = GRID_ZOOM

        cells = generate_cells(*vals, cell_km=cell_km)
        total = len(cells)

        self.log(f"Grid mode: {total} cells × ~20 results each")
        self.log(f"Query: {query}  |  zoom: {zoom}  |  cell: {cell_km} km")
        self.set_status(f"Submitting {total} jobs...")

        all_rows = []
        seen_ids = set()
        completed = 0
        failed = 0

        fields = self._resolve_fields()
        self.log(f"Fields: {fields or 'all'}")

        try:
            with ThreadPoolExecutor(max_workers=GRID_MAX_PARALLEL) as pool:
                futures = {
                    pool.submit(self._run_cell, query, lat, lon, zoom, fields): (lat, lon)
                    for lat, lon in cells
                }

                for future in as_completed(futures):
                    lat, lon = futures[future]
                    completed += 1
                    pct = completed / total * 100
                    self.progress_var.set(pct)
                    self.set_status(f"Progress: {completed}/{total} cells")

                    try:
                        rows = future.result()
                        new = 0
                        for row in rows:
                            uid = row.get("place_id") or row.get("title", "") + "|" + row.get("address", "")
                            if uid not in seen_ids:
                                seen_ids.add(uid)
                                all_rows.append(row)
                                new += 1
                        if new:
                            self.log(f"Cell ({lat:.3f},{lon:.3f}): +{new} new (total {len(all_rows)})")
                    except Exception as e:
                        failed += 1
                        self.log(f"Cell ({lat:.3f},{lon:.3f}) failed: {e}")

            if not all_rows:
                self.fail("No results from any cell")
                return

            out = self.export_xlsx(query, all_rows)
            self.log(f"XLSX saved: {out}")
            self.set_status(f"Done — {len(all_rows)} unique results ({failed} cells failed)")
            messagebox.showinfo("Success", f"{len(all_rows)} unique results saved:\n{out}")

        except Exception as e:
            self.fail(str(e))
        finally:
            self.is_running = False
            self.run_btn.config(state=tk.NORMAL)

    def _run_cell(self, query, lat, lon, zoom, fields=None):
        """Submit + wait for a single grid cell. Runs in a thread pool."""
        job_id = self.submit_job(
            query,
            max_depth=3,
            geo_coordinates=f"{lat},{lon}",
            zoom=zoom,
            fast_mode=True,
            fields=fields,
        )
        result = self.wait_for_job(job_id, logger=None)
        return self.parse_results(result.get("results"))

    # ------------------------------------------------------------------
    # API helpers
    # ------------------------------------------------------------------

    def _resolve_fields(self):
        """Parse the fields input into a list to send to the API."""
        raw = self.fields_var.get().strip().lower()
        if not raw or raw == "all":
            return []  # empty = all fields
        if raw == "lite":
            return ["lite"]
        # comma-separated custom list
        return [f.strip() for f in raw.split(",") if f.strip()]

    def submit_job(self, query, max_depth=10, geo_coordinates=None, zoom=None, fast_mode=False, fields=None):
        payload = {
            "keyword": query,
            "lang": "en",
            "max_depth": max_depth,
            "timeout": 300,
        }
        if geo_coordinates:
            payload["geo_coordinates"] = geo_coordinates
        if zoom:
            payload["zoom"] = zoom
        if fast_mode:
            payload["fast_mode"] = True
        if fields:
            payload["fields"] = fields

        url = f"{self.config['api_url'].rstrip('/')}/api/v1/scrape"
        resp = requests.post(url, json=payload, timeout=30)
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
                if not (200 <= resp.status_code < 300):
                    time.sleep(POLL_INTERVAL)
                    continue
                status = resp.json()
                if logger:
                    logger(f"Status: {status['status']} | Results: {status.get('result_count', 0)}")
                s = status.get("status", "").lower()
                if s == "completed":
                    return status
                if s == "failed":
                    raise Exception(status.get("error") or "Job failed")
            except requests.RequestException as e:
                if logger:
                    logger(f"Poll error: {e}")
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

        # Collect all headers (sorted), prioritise common fields first
        priority = ["title", "category", "address", "phone", "website", "rating", "reviews_count"]
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

        # Auto column widths
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

        safe = re.sub(r"[^a-z0-9_\-]", "", query.lower().replace(" ", "_"))[:60]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        downloads = Path.home() / "Downloads"
        downloads.mkdir(exist_ok=True)
        out = downloads / f"{safe}_{ts}.xlsx"
        wb.save(out)
        return str(out)

    def _cell_val(self, v):
        if v is None:
            return ""
        if isinstance(v, (str, int, float, bool)):
            return v
        return json.dumps(v, ensure_ascii=False)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def open_settings(self):
        d = tk.Toplevel(self.root)
        d.title("Settings")
        d.geometry("520x150")
        d.transient(self.root)

        tk.Label(d, text="API URL:", font=("Arial", 10)).pack(anchor=tk.W, padx=10, pady=10)
        entry = tk.Entry(d, font=("Arial", 11))
        entry.pack(fill=tk.X, padx=10)
        entry.insert(0, self.config["api_url"])

        def save():
            url = entry.get().strip().rstrip("/")
            if not url.startswith(("http://", "https://")):
                messagebox.showerror("Error", "URL must start with http:// or https://")
                return
            self.config["api_url"] = url
            self.save_config()
            self.log(f"API URL updated: {url}")
            d.destroy()

        tk.Button(d, text="Save", command=save, bg="#4CAF50", fg="white", font=("Arial", 10)).pack(pady=10)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.logs.config(state=tk.NORMAL)
        self.logs.insert(tk.END, line)
        self.logs.see(tk.END)
        self.logs.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def set_status(self, s):
        self.status_label.config(text=s)
        self.root.update_idletasks()

    def fail(self, msg):
        self.log(f"Error: {msg}")
        self.set_status("Failed")
        messagebox.showerror("Error", msg)

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
