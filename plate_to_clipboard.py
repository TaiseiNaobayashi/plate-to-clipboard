"""Find the 96-well values in a plate reader (Gen5) xlsx export and copy them
to the clipboard as a single column of "A1<TAB>value" lines.

Usage:
  - Drag & drop an xlsx onto the window, or click it to choose a file
  - Dropping an xlsx onto the exe / .py also opens the window and copies it
"""
import ctypes
import ctypes.wintypes as wt
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog

import openpyxl

ROWS = "ABCDEFGH"
COLS = list(range(1, 13))


def find_plates(path):
    """Return every block on any sheet with column headers 1..12 and row labels A..H."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    plates = []
    for ws in wb.worksheets:
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        for r, row in enumerate(grid):
            for c in range(len(row) - 11):
                if [_num(v) for v in row[c:c + 12]] != COLS:
                    continue
                # Row labels A..H must sit below the header, one column to its left
                lc = c - 1
                if lc < 0 or r + 8 >= len(grid):
                    continue
                labels = [_cell(grid, r + 1 + i, lc) for i in range(8)]
                if [str(v).strip().upper() if v is not None else "" for v in labels] != list(ROWS):
                    continue
                values = {}
                for i, rl in enumerate(ROWS):
                    for j, cl in enumerate(COLS):
                        values[f"{rl}{cl}"] = _cell(grid, r + 1 + i, c + j)
                plates.append((ws.title, values))
    wb.close()
    return plates


def _cell(grid, r, c):
    row = grid[r]
    return row[c] if c < len(row) else None


def _num(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else None
    except (TypeError, ValueError):
        return None


def to_tsv(plates):
    """One plate: "A1<TAB>value". Several plates: one column per plate, with a header row."""
    lines = []
    if len(plates) > 1:
        lines.append("\t".join(["Well"] + [title for title, _ in plates]))
    for well in plates[0][1]:
        vals = ["" if v.get(well) is None else str(v[well]) for _, v in plates]
        lines.append("\t".join([well] + vals))
    return "\r\n".join(lines) + "\r\n"


def set_clipboard(text):
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    k32 = ctypes.windll.kernel32
    u32 = ctypes.windll.user32
    k32.GlobalAlloc.restype = wt.HGLOBAL
    k32.GlobalAlloc.argtypes = [wt.UINT, ctypes.c_size_t]
    k32.GlobalLock.restype = wt.LPVOID
    k32.GlobalLock.argtypes = [wt.HGLOBAL]
    k32.GlobalUnlock.argtypes = [wt.HGLOBAL]
    u32.SetClipboardData.argtypes = [wt.UINT, wt.HANDLE]
    u32.SetClipboardData.restype = wt.HANDLE

    data = text.encode("utf-16-le") + b"\x00\x00"
    h = k32.GlobalAlloc(GMEM_MOVEABLE, len(data))
    p = k32.GlobalLock(h)
    ctypes.memmove(p, data, len(data))
    k32.GlobalUnlock(h)
    if not u32.OpenClipboard(None):
        raise RuntimeError("Could not open the clipboard")
    try:
        u32.EmptyClipboard()
        if not u32.SetClipboardData(CF_UNICODETEXT, h):
            raise RuntimeError("Failed to write to the clipboard")
    finally:
        u32.CloseClipboard()


class App:
    def __init__(self, root):
        self.root = root
        root.title("Plate → Clipboard")
        root.geometry("520x380")
        root.minsize(420, 300)

        self.drop = tk.Label(
            root, text="Drop an xlsx file here\n(or click to choose)",
            relief="groove", bd=2, bg="#f4f6fa", fg="#444",
            font=("Yu Gothic UI", 13), cursor="hand2", height=4,
        )
        self.drop.pack(fill="x", padx=12, pady=(12, 8))
        self.drop.bind("<Button-1>", lambda e: self.choose())
        try:
            from tkinterdnd2 import DND_FILES
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)
        except Exception:
            self.drop.config(text="Click to choose an xlsx file")

        tk.Label(root, text="On the clipboard:", anchor="w",
                 font=("Yu Gothic UI", 9)).pack(fill="x", padx=12)
        self.current = tk.Label(
            root, text="(nothing copied yet)", anchor="w", justify="left",
            font=("Yu Gothic UI", 12, "bold"), fg="#888", wraplength=490,
        )
        self.current.pack(fill="x", padx=12)
        self.detail = tk.Label(root, text="", anchor="w", justify="left",
                               font=("Yu Gothic UI", 9), fg="#666", wraplength=490)
        self.detail.pack(fill="x", padx=12, pady=(0, 8))

        tk.Label(root, text="History (double-click to copy again):", anchor="w",
                 font=("Yu Gothic UI", 9)).pack(fill="x", padx=12)
        self.history = tk.Listbox(root, font=("Yu Gothic UI", 9), activestyle="none")
        self.history.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.history.bind("<Double-Button-1>", self.on_history)
        self.paths = []

    def choose(self):
        path = filedialog.askopenfilename(
            title="Choose a plate reader xlsx",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.load(path)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.load(files[0])

    def on_history(self, _event):
        sel = self.history.curselection()
        if sel:
            self.load(self.paths[sel[0]])

    def load(self, path):
        path = str(Path(path))
        name = Path(path).name
        try:
            plates = find_plates(path)
            if not plates:
                raise ValueError("No 96-well block found")
            set_clipboard(to_tsv(plates))
        except Exception as e:
            self.current.config(text=f"✖ {name}", fg="#c0392b")
            self.detail.config(text=f"Error: {e}\n(clipboard not changed)")
            return
        now = datetime.now().strftime("%H:%M:%S")
        self.current.config(text=f"📋 {name}", fg="#1a7f37")
        self.detail.config(
            text=f"{Path(path).parent}\n{len(plates)} plate(s) × 96 wells  ·  copied at {now}")
        self.root.title(f"Plate → Clipboard  [{name}]")
        if path in self.paths:
            idx = self.paths.index(path)
            self.paths.pop(idx)
            self.history.delete(idx)
        self.paths.insert(0, path)
        self.history.insert(0, f"{now}  {name}")


def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    app = App(root)
    if len(sys.argv) > 1:
        root.after(100, app.load, sys.argv[1])
    root.mainloop()


if __name__ == "__main__":
    main()
