from datetime import datetime, date
from pathlib import Path
from tkinter import filedialog, ttk, messagebox
from typing import Optional
import tkinter as tk
from tkcalendar import Calendar

import pandas as pd


def select_excel_file(title, parent=None, initial_path=None):
    initialdir = None
    initialfile = None

    if initial_path:
        path = Path(initial_path).resolve()
        initialdir = str(path.parent)
        initialfile = path.name

    return filedialog.askopenfilename(
        parent=parent,
        title=title,
        initialdir=initialdir,
        initialfile=initialfile,
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


def normalize_header_token(v) -> str:
    return str(v or "").strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")


def find_column_by_keywords(df: pd.DataFrame, keywords):
    normalized_keywords = [normalize_header_token(k) for k in keywords]

    for col in df.columns:
        text = normalize_header_token(col)
        if all(k in text for k in normalized_keywords):
            return col

    return None


def load_actual_quantities_by_part(
    result_file: str,
    part_keywords=("품번",),
    qty_keywords=("실적", "수량"),
) -> dict:
    df = pd.read_excel(result_file)
    part_col = find_column_by_keywords(df, part_keywords)
    qty_col = find_column_by_keywords(df, qty_keywords)

    if part_col is None or qty_col is None:
        raise RuntimeError("실적 파일에 품번/실적수량 컬럼이 없습니다.")

    df = df[[part_col, qty_col]].copy()
    df[part_col] = df[part_col].apply(normalize_part_no)
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df = df[df[part_col] != ""]

    return df.groupby(part_col)[qty_col].sum().to_dict()


def apply_quantities_by_part_left_to_right(
    df: pd.DataFrame,
    quantities_by_part: dict,
    part_col,
    value_start_idx: int,
) -> pd.DataFrame:
    if df.empty or not quantities_by_part:
        return df

    out = df.copy()
    part_col_name = out.columns[part_col] if isinstance(part_col, int) else part_col

    value_col_positions = range(value_start_idx, len(out.columns))

    for idx, row in out.iterrows():
        part_no = normalize_part_no(row[part_col_name])
        remain = safe_float(quantities_by_part.get(part_no, 0))
        if remain <= 0:
            continue

        for col_pos in value_col_positions:
            if remain <= 0:
                break

            qty = safe_float(out.iat[idx, col_pos])
            if qty <= 0:
                continue

            used = min(qty, remain)
            out.iat[idx, col_pos] = qty - used
            remain -= used

    value_values = out.iloc[:, value_start_idx:]
    return out[(value_values != 0).any(axis=1)].reset_index(drop=True)


def timestamped_filename(source_file: str, suffix: str, ext: str = ".xlsx") -> str:
    source_path = Path(source_file)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{source_path.stem}_{suffix}_{timestamp}{ext}"


def select_excel_save_file(title: str, source_file: str, suffix: str, parent=None):
    source_path = Path(source_file)
    return filedialog.asksaveasfilename(
        parent=parent,
        title=title,
        initialdir=str(source_path.parent),
        initialfile=timestamped_filename(source_file, suffix),
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
    )


def select_date_range(parent):
    win = tk.Toplevel(parent)
    win.title("기간 선택")
    win.geometry("620x330")
    win.resizable(False, False)

    selected = {"start": None, "end": None}

    tk.Label(win, text="시작일").grid(row=0, column=0, padx=10, pady=(10, 0))
    tk.Label(win, text="종료일").grid(row=0, column=1, padx=10, pady=(10, 0))

    cal_start = Calendar(win, selectmode="day", date_pattern="yyyy-mm-dd")
    cal_start.grid(row=1, column=0, padx=10, pady=10)

    cal_end = Calendar(win, selectmode="day", date_pattern="yyyy-mm-dd")
    cal_end.grid(row=1, column=1, padx=10, pady=10)

    def confirm():
        start = pd.to_datetime(cal_start.get_date()).normalize()
        end = pd.to_datetime(cal_end.get_date()).normalize()

        if start > end:
            messagebox.showerror("오류", "시작일이 종료일보다 늦습니다.", parent=win)
            return

        selected["start"] = start
        selected["end"] = end
        win.destroy()

    def no_filter():
        selected["start"] = None
        selected["end"] = None
        win.destroy()

    tk.Button(win, text="확인", width=15, command=confirm).grid(row=2, column=0, pady=10)
    tk.Button(win, text="기간 필터 없이 진행", width=20, command=no_filter).grid(row=2, column=1, pady=10)

    win.protocol("WM_DELETE_WINDOW", no_filter)
    win.update_idletasks()
    win.lift()
    win.attributes("-topmost", True)
    win.after(300, lambda: win.attributes("-topmost", False))
    win.focus_force()
    win.grab_set()

    parent.wait_window(win)
    return selected["start"], selected["end"]
