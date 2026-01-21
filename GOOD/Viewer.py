#!/usr/bin/env python3
"""
Listings Viewer (Excel or CSV)
- Open an .xlsx/.xlsm OR .csv exported by your scraper.
- Shows one row at a time: image (from image_url) + title (title_en or title).
- Controls: Previous, Next, Open Link (detail_url).
"""

import io
import sys
import os
import webbrowser
from dataclasses import dataclass
from typing import Optional, Dict

import requests
import pandas as pd
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox

# Display size for images in the GUI (pixels)
MAX_W = 640
MAX_H = 640


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
        self.rows: list[ListingRow] = []
        self.idx: int = -1
        self.image_cache: Dict[str, ImageTk.PhotoImage] = {}
        self.current_photo: Optional[ImageTk.PhotoImage] = None

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

        rows: list[ListingRow] = []
        for _, r in df.iterrows():
            # Prefer title_en if it is plain text (not a formula that starts with '=')
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
        self.update_controls()
        self.show_current()

    def _read_any_table(self, path: str) -> pd.DataFrame:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            # Try to auto-detect delimiter; BOM-safe; skip bad lines if any.
            try:
                return pd.read_csv(path, engine="python", sep=None, encoding="utf-8-sig", on_bad_lines="skip")
            except Exception:
                return pd.read_csv(path, encoding="utf-8", on_bad_lines="skip")
        else:
            # Excel
            return pd.read_excel(path, engine="openpyxl")

    # ---------- Navigation ----------
    def prev_row(self):
        if self.idx > 0:
            self.idx -= 1
            self.show_current()

    def next_row(self):
        if self.idx < len(self.rows) - 1:
            self.idx += 1
            self.show_current()

    def update_controls(self):
        have = self.idx >= 0
        self.prev_btn.config(state=(tk.NORMAL if have and self.idx > 0 else tk.DISABLED))
        self.next_btn.config(state=(tk.NORMAL if have and self.idx < len(self.rows) - 1 else tk.DISABLED))
        self.open_btn.config(state=(tk.NORMAL if have and self.rows[self.idx].detail_url else tk.DISABLED))

    # ---------- Rendering ----------
    def show_current(self):
        if self.idx < 0 or self.idx >= len(self.rows):
            self.image_panel.config(image="", text="No data loaded")
            self.title_label.config(text="")
            self.status_label.config(text="No data")
            self.update_controls()
            return

        row = self.rows[self.idx]
        self.title_label.config(text=row.title)

        # Fetch/resize image
        photo = self.fetch_image(row.image_url)
        if photo is None:
            self.image_panel.config(image="", text="(no image)", font=("Segoe UI", 12))
        else:
            self.image_panel.config(image=photo, text="")
        self.current_photo = photo  # keep reference to prevent GC

        self.status_label.config(text=f"Row {self.idx + 1} of {len(self.rows)}")
        self.update_controls()

    def fetch_image(self, url: Optional[str]) -> Optional[ImageTk.PhotoImage]:
        if not url:
            return None
        if url in self.image_cache:
            return self.image_cache[url]
        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            img = Image.open(io.BytesIO(resp.content)).convert("RGB")
        except Exception as e:
            self.image_panel.config(text=f"(image load failed)\n{e}", image="")
            return None

        # Fit within MAX_W x MAX_H while preserving aspect
        img.thumbnail((MAX_W, MAX_H), Image.LANCZOS)
        tk_img = ImageTk.PhotoImage(img)
        self.image_cache[url] = tk_img
        return tk_img

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
