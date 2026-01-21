#!/usr/bin/env python3
"""
Listings Viewer (Excel or CSV)
- Pre-downloads and decodes ALL images before enabling navigation
  so Next/Previous have effectively zero delay.
"""

import io
import sys
import os
import webbrowser
from dataclasses import dataclass
from typing import Optional, Dict, List, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import pandas as pd
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Display size for images in the GUI (pixels)
MAX_W = 640
MAX_H = 640

# Parallelism & networking
MAX_WORKERS = 6
REQUEST_TIMEOUT = (5, 20)  # (connect, read) seconds


@dataclass
class ListingRow:
    title: str
    image_url: Optional[str]
    detail_url: Optional[str]


class ListingsViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Listings Viewer")
        self.geometry("800x800")
        self.minsize(700, 700)

        # Top bar: Pick file
        top = tk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)

        tk.Button(top, text="Open CSV/Excel...", command=self.load_table).pack(side=tk.LEFT)
        self.path_label = tk.Label(top, text="", anchor="w")
        self.path_label.pack(side=tk.LEFT, padx=10)

        # Image area
        self.image_panel = tk.Label(self, bd=1, relief=tk.SUNKEN, bg="#f4f4f4")
        self.image_panel.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(6, 4))

        # Title label
        self.title_label = tk.Label(self, text="", font=("Segoe UI", 12), wraplength=760, justify="left")
        self.title_label.pack(side=tk.TOP, anchor="w", padx=10, pady=(0, 8))

        # Row status
        self.status_label = tk.Label(self, text="No file loaded", anchor="w", fg="#666")
        self.status_label.pack(side=tk.TOP, anchor="w", padx=10, pady=(0, 8))

        # Controls
        ctrl = tk.Frame(self)
        ctrl.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        self.prev_btn = tk.Button(ctrl, text="◀ Previous", command=self.prev_row, state=tk.DISABLED)
        self.prev_btn.pack(side=tk.LEFT)

        self.open_btn = tk.Button(ctrl, text="Open Link", command=self.open_link, state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT, padx=8)

        self.next_btn = tk.Button(ctrl, text="Next ▶", command=self.next_row, state=tk.DISABLED)
        self.next_btn.pack(side=tk.LEFT)

        # Keyboard shortcuts
        self.bind("<Right>", lambda e: self.next_row())
        self.bind("<Left>", lambda e: self.prev_row())

        # Data
        self.rows: List[ListingRow] = []
        self.idx: int = -1

        # Cache: url -> PhotoImage (ALL images are loaded here during preload)
        self.image_cache: Dict[str, ImageTk.PhotoImage] = {}
        self.current_photo: Optional[ImageTk.PhotoImage] = None

        # Networking: pooled session with retries
        self.session = self._make_session()

    # ---------- Session with pooling/retries ----------
    def _make_session(self) -> requests.Session:
        s = requests.Session()
        retry = Retry(
            total=2,
            backoff_factor=0.3,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=("GET",)
        )
        adapter = HTTPAdapter(pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS, max_retries=retry)
        s.mount("http://", adapter)
        s.mount("https://", adapter)
        s.headers.update({
            "User-Agent": "ListingsViewer/1.0 (+GUI)",
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
        })
        return s

    # ---------- File loading ----------
    def load_table(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            df = self._read_any_table(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file:\n{e}")
            return

        # Normalize/locate columns (case-insensitive)
        cols_map = {c.strip().lower(): c for c in df.columns if isinstance(c, str)}
        def col(name): return cols_map.get(name)

        image_col = col("image_url") or col("image") or col("image link")
        title_en_col = col("title_en")
        title_col = col("title")
        detail_col = col("detail_url")

        if not image_col:
            messagebox.showerror("Missing column", "Could not find an 'image_url' column.")
            return

        rows: List[ListingRow] = []
        for _, r in df.iterrows():
            # Prefer title_en if it's plain text; else fallback to title.
            t_en = r.get(title_en_col) if title_en_col else None
            t = r.get(title_col) if title_col else None

            title = ""
            if isinstance(t_en, str) and t_en.strip() and not t_en.strip().startswith("="):
                title = t_en.strip()
            elif isinstance(t, str) and t.strip():
                title = t.strip()
            elif isinstance(t_en, (int, float)):
                title = str(t_en)
            elif isinstance(t, (int, float)):
                title = str(t)

            image_url = r.get(image_col)
            if isinstance(image_url, str):
                image_url = image_url.strip().strip('"').strip("'")
            else:
                image_url = None

            detail_url = r.get(detail_col) if detail_col else None
            if isinstance(detail_url, str):
                detail_url = detail_url.strip()
            else:
                detail_url = None

            rows.append(ListingRow(title=title or "(no title)", image_url=image_url, detail_url=detail_url))

        self.rows = rows
        self.idx = 0 if self.rows else -1
        self.path_label.config(text=path)
        self.status_label.config(text=f"Loaded {len(self.rows)} rows")

        # Disable controls while we preload all images
        self.prev_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.DISABLED)

        # Start full preload (modal progress). When done, show current row.
        self.after(0, self._preload_all_images)

    def _read_any_table(self, path: str) -> pd.DataFrame:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            try:
                return pd.read_csv(path, engine="python", sep=None, encoding="utf-8-sig", on_bad_lines="skip")
            except Exception:
                return pd.read_csv(path, encoding="utf-8", on_bad_lines="skip")
        else:
            return pd.read_excel(path, engine="openpyxl")

    # ---------- Preload-all with progress (fixed) ----------
    def _preload_all_images(self):
        """Download + resize + decode ALL images; create PhotoImages and cache them.
           Shows a modal progress dialog and enables navigation when finished.
        """
        urls = [r.image_url for r in self.rows if r.image_url]
        urls = list(dict.fromkeys(urls))  # de-duplicate, preserve order
        total = len(urls)

        if total == 0:
            # Nothing to preload
            self._enable_controls()
            self.show_current()
            return

        # Progress dialog (modal)
        dlg = tk.Toplevel(self)
        dlg.title("Preloading images…")
        dlg.transient(self)
        dlg.grab_set()  # modal
        ttk.Label(dlg, text=f"Downloading {total} images…").pack(padx=16, pady=(16, 6))
        pb = ttk.Progressbar(dlg, orient="horizontal", mode="determinate", length=320, maximum=total)
        pb.pack(padx=16, pady=6)
        msg = tk.Label(dlg, text="", fg="#666")
        msg.pack(padx=16, pady=(0, 16))
        dlg.update_idletasks()
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 60, self.winfo_rooty() + 60))

        # Download/resize in threads; create PhotoImages on main thread
        def download_resize(u: str) -> Tuple[str, Optional[bytes], Optional[str]]:
            try:
                resp = self.session.get(u, timeout=REQUEST_TIMEOUT)
                resp.raise_for_status()
                img = Image.open(io.BytesIO(resp.content)).convert("RGB")
                img.thumbnail((MAX_W, MAX_H), Image.LANCZOS)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                return u, buf.getvalue(), None
            except Exception as e:
                return u, None, str(e)

        successes = 0
        failures = 0

        # Use a thread pool to fetch/resize quickly
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {ex.submit(download_resize, u): u for u in urls}
            index = 0
            for fut in as_completed(futures):
                u, data, err = fut.result()
                index += 1

                # Create PhotoImage in main thread
                def make_photo(url=u, blob=data, error=err, i=index):
                    nonlocal successes, failures
                    pb["value"] = i
                    if blob is not None:
                        img = Image.open(io.BytesIO(blob))
                        tk_img = ImageTk.PhotoImage(img)
                        self.image_cache[url] = tk_img
                        successes += 1
                        msg.config(text=f"{successes}/{total} downloaded")
                    else:
                        failures += 1
                        msg.config(text=f"{successes}/{total} downloaded • {failures} failed")
                    dlg.update_idletasks()

                self.after(0, make_photo)

        # Close dialog & enable UI (fixed: no bogus wait)
        dlg.grab_release()
        dlg.destroy()

        self._enable_controls()
        self.show_current()

    def _enable_controls(self):
        have = self.idx >= 0
        self.prev_btn.config(state=(tk.NORMAL if have and self.idx > 0 else tk.DISABLED))
        self.next_btn.config(state=(tk.NORMAL if have and self.idx < len(self.rows) - 1 else tk.DISABLED))
        self.open_btn.config(state=(tk.NORMAL if have and self.rows[self.idx].detail_url else tk.DISABLED))

    # ---------- Navigation ----------
    def prev_row(self):
        if self.idx > 0:
            self.idx -= 1
            self.show_current()

    def next_row(self):
        if self.idx < len(self.rows) - 1:
            self.idx += 1
            self.show_current()

    # ---------- Rendering ----------
    def show_current(self):
        if self.idx < 0 or self.idx >= len(self.rows):
            self.image_panel.config(image="", text="No data loaded")
            self.title_label.config(text="")
            self.status_label.config(text="No data")
            self._enable_controls()
            return

        row = self.rows[self.idx]
        self.title_label.config(text=row.title)

        url = row.image_url
        if not url:
            self.current_photo = None
            self.image_panel.config(image="", text="(no image)")
        else:
            photo = self.image_cache.get(url)
            if photo is None:
                # Shouldn't happen if we preloaded, but handle gracefully
                self.image_panel.config(image="", text="(not preloaded)")
                self.current_photo = None
            else:
                self.current_photo = photo
                self.image_panel.config(image=photo, text="")

        self.status_label.config(text=f"Row {self.idx + 1} of {len(self.rows)}")
        self._enable_controls()

    # ---------- Actions ----------
    def open_link(self):
        if self.idx < 0:
            return
        url = self.rows[self.idx].detail_url
        if not url:
            messagebox.showinfo("No URL", "This row has no detail_url.")
            return
        webbrowser.open(url)


def main():
    app = ListingsViewer()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        sys.stderr.write(
            f"Missing dependency: {e}\n"
            "Install with: pip install pillow pandas openpyxl requests\n"
        )
        raise
