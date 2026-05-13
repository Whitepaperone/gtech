from datetime import datetime, date
from tkinter import filedialog, ttk
from typing import Optional
import tkinter as tk

import pandas as pd


def select_excel_file(title, parent=None):
    return filedialog.askopenfilename(
        parent=parent,
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )


def create_progress_window(root, title="진행 중"):
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("420x120")
    win.resizable(False, False)
    win.attributes("-topmost", True)

    label = tk.Label(win, text="준비 중...", anchor="w")
    label.pack(fill="x", padx=20, pady=(20, 5))

    bar = ttk.Progressbar(win, orient="horizontal", length=360, mode="determinate")
    bar.pack(padx=20, pady=10)

    percent_label = tk.Label(win, text="0%")
    percent_label.pack()

    win.update()

    def update_progress(percent, text=None):
        percent = max(0, min(100, int(percent)))
        bar["value"] = percent
        percent_label.config(text=f"{percent}%")
        if text:
            label.config(text=text)
        win.update_idletasks()
        win.update()

    return win, update_progress


def normalize_process_token(v) -> str:
    return (
        str(v or "")
        .strip()
        .upper()
        .replace(" ", "")
        .replace("\n", "")
        .replace("악세서리", "액세서리")
        .replace("액세사리", "액세서리")
    )


def is_date_value(v) -> bool:
    return isinstance(v, (datetime, date))


def safe_float(v) -> float:
    n = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(n) else float(n)


def normalize_part_no(v) -> str:
    t = normalize_process_token(v).upper()
    return t.replace(" ", "").replace("_", "")


def normalize_kind(kind_text: str) -> Optional[str]:
    t = normalize_process_token(kind_text).replace("\n", "")
    if t in ("미달", "계획", "실적"):
        return t
    return None


def header_text(ws, row: int, col: int) -> str:
    a = normalize_process_token(ws.cell(row, col).value)
    b = normalize_process_token(ws.cell(row + 1, col).value)
    return f"{a} {b}".strip()


def row_join_text(ws, r: int, upto_col: Optional[int] = None) -> str:
    if upto_col is None:
        upto_col = ws.max_column
    vals = [normalize_process_token(ws.cell(r, c).value) for c in range(1, upto_col + 1)]
    vals = [v for v in vals if v]
    return " ".join(vals)


def get_merged_value(ws, row: int, col: int):
    cell = ws.cell(row, col)
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            return ws.cell(merged.min_row, merged.min_col).value
    return cell.value


def build_merged_map(ws):
    merged_map = {}
    for merged in ws.merged_cells.ranges:
        top_value = ws.cell(merged.min_row, merged.min_col).value
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                merged_map[(row, col)] = top_value
    return merged_map


def filter_by_period(df: pd.DataFrame, start_date=None, end_date=None) -> pd.DataFrame:
    if df.empty or "날짜" not in df.columns:
        return df

    out = df.copy()
    out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.normalize()

    if start_date is not None:
        out = out[out["날짜"] >= start_date]

    if end_date is not None:
        out = out[out["날짜"] <= end_date]

    return out.copy()
