#!/usr/bin/env python3
"""Extract header and tube requirements from CORE sheet and export Excel summary.

Usage:
  python extract_plan_tubes.py --input core_sheet.xlsx --date YYYY-MM-DD

Config is read from `extract_tubes_config.yaml` if present. Defaults are sensible.
"""
import argparse
import sys
from pathlib import Path
import datetime
import pandas as pd
from AppConstants import (
    CORE_TUBE_OUTPUT_FILE,
    CORE_TUBE_HEADER_COL,
    CORE_TUBE_TUBE_COL,
    CORE_TUBE_COUNT_COL,
)
from CommonUtils import select_excel_file, select_date_range
from CommonUtils import create_progress_window
import tkinter as tk
from tkinter import messagebox, simpledialog
from ExtractPlan import export_core_tube_requirements


def load_config(path: Path = None):
    # Use AppConstants defaults. YAML configs are intentionally not used.
    default = {
        'sheet_name': 'CORE공정(실적)',
        'header_col': CORE_TUBE_HEADER_COL or 'H',
        'tube_col': CORE_TUBE_TUBE_COL or 'I',
        'tube_count_col': CORE_TUBE_COUNT_COL or 'J',
        'date_column': 'auto',
        'qty_column': 'auto',
        'model_column': 'auto',
        'output': CORE_TUBE_OUTPUT_FILE or 'plan_header_tube.xlsx'
    }
    return default


def col_letter_to_index(letter: str) -> int:
    letter = letter.upper().strip()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def detect_columns(df: pd.DataFrame, cfg):
    # Determine header, tube, tube_count by letter -> index fallback
    try:
        h_idx = col_letter_to_index(cfg['header_col'])
        t_idx = col_letter_to_index(cfg['tube_col'])
        tc_idx = col_letter_to_index(cfg['tube_count_col'])
    except Exception:
        h_idx = t_idx = tc_idx = None

    def get_col_by_idx(idx):
        if idx is None:
            return None
        if idx < len(df.columns):
            return df.columns[idx]
        return None

    header_col = get_col_by_idx(h_idx) or cfg.get('header_col')
    tube_col = get_col_by_idx(t_idx) or cfg.get('tube_col')
    tube_count_col = get_col_by_idx(tc_idx) or cfg.get('tube_count_col')

    # Date column detection
    date_col = None
    if cfg.get('date_column') and cfg.get('date_column') != 'auto':
        if cfg['date_column'] in df.columns:
            date_col = cfg['date_column']
    else:
        for candidate in ['날짜', '일자', 'Date', 'date']:
            if candidate in df.columns:
                date_col = candidate
                break
        if date_col is None:
            # try to find datetime dtype
            for c in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[c]):
                    date_col = c
                    break

    qty_col = None
    if cfg.get('qty_column') and cfg.get('qty_column') != 'auto':
        qty_col = cfg['qty_column'] if cfg['qty_column'] in df.columns else None
    else:
        for candidate in ['수량', 'Qty', '수', '소요량', 'quantity', 'Quantity']:
            if candidate in df.columns:
                qty_col = candidate
                break
        if qty_col is None:
            # pick first numeric column
            for c in df.columns:
                if pd.api.types.is_numeric_dtype(df[c]):
                    qty_col = c
                    break

    model_col = None
    if cfg.get('model_column') and cfg.get('model_column') != 'auto':
        model_col = cfg['model_column'] if cfg['model_column'] in df.columns else None
    else:
        for candidate in ['모델', 'Model', '모델명', '품번']:
            if candidate in df.columns:
                model_col = candidate
                break

    return header_col, tube_col, tube_count_col, date_col, qty_col, model_col


def prepare_df(df: pd.DataFrame, cols):
    header_col, tube_col, tube_count_col, date_col, qty_col, model_col = cols
    working = df.copy()
    # If header/tube are positional like 'H' because name not found, try iloc fallback
    if header_col not in working.columns and isinstance(header_col, int):
        header_col = working.columns[header_col]
    if tube_col not in working.columns and isinstance(tube_col, int):
        tube_col = working.columns[tube_col]

    # Normalize date
    if date_col is None:
        raise SystemExit('Date column not found. Please set `date_column` in config.')
    working[date_col] = pd.to_datetime(working[date_col], errors='coerce')
    working = working.dropna(subset=[date_col])

    if qty_col is None:
        raise SystemExit('Quantity column not found. Please set `qty_column` in config.')

    # ensure numeric
    working[qty_col] = pd.to_numeric(working[qty_col], errors='coerce').fillna(0)
    # tube count numeric
    if tube_count_col not in working.columns:
        # try index fallback
        try:
            tc_idx = col_letter_to_index(tube_count_col)
            tube_count_col = working.columns[tc_idx]
        except Exception:
            pass
    working[tube_count_col] = pd.to_numeric(working[tube_count_col], errors='coerce').fillna(0)

    return working, header_col, tube_col, tube_count_col, date_col, qty_col, model_col


def build_summary(working, header_col, tube_col, tube_count_col, date_col, qty_col, model_col, asof_date: datetime.date):
    # Compute core quantities grouped by header/tube and date
    working['date_only'] = working[date_col].dt.date

    # backlog mask
    backlog_mask = working['date_only'] < asof_date
    future_mask = working['date_only'] >= asof_date

    # Header summary: required = core_qty * 2
    header_grp = working.groupby([header_col, 'date_only'])[qty_col].sum().reset_index()
    header_pivot = header_grp.pivot_table(index=header_col, columns='date_only', values=qty_col, aggfunc='sum', fill_value=0)
    # multiply by 2
    header_pivot = header_pivot * 2

    # Model column: pick first model per header
    model_map = {}
    if model_col and model_col in working.columns:
        first_models = working.sort_values(date_col).groupby(header_col)[model_col].first()
        model_map = first_models.to_dict()

    # add backlog column and format date columns as strings
    dates = sorted([c for c in header_pivot.columns if isinstance(c, (datetime.date, datetime.datetime, pd.Timestamp))])
    backlog_col_name = '미달(이전)'
    # compute past sum; if negative (pre-production) subtract from future dates
    past_sum = header_pivot[[d for d in dates if d < asof_date]].sum(axis=1) if dates else pd.Series(0, index=header_pivot.index)
    # build continuous display dates including weekends
    display_start = asof_date
    display_end = max(dates) if dates else display_start
    display_dates = pd.to_datetime(pd.date_range(start=display_start, end=display_end, freq='D')).normalize().tolist()
    display_df = header_pivot.reindex(columns=display_dates).fillna(0)
    adj_df = pd.DataFrame(0, index=header_pivot.index, columns=display_dates)
    for idx in header_pivot.index:
        carry = past_sum.loc[idx] if isinstance(past_sum, pd.Series) else past_sum
        for d in display_dates:
            orig = display_df.at[idx, d] if d in display_df.columns else 0
            adj = orig + carry
            if adj < 0:
                adj_df.at[idx, d] = 0
                carry = adj
            else:
                adj_df.at[idx, d] = adj
                carry = 0
    header_pivot[backlog_col_name] = past_sum.clip(lower=0)
    for d in display_dates:
        header_pivot[d] = adj_df[d]
    keep_dates = [backlog_col_name] + display_dates
    header_out = header_pivot.reindex(columns=keep_dates).fillna(0)
    header_out['총소요량'] = header_out.sum(axis=1)
    header_out.insert(0, '헤더품번', header_out.index)
    header_out.insert(0, '모델명', header_out['헤더품번'].map(model_map))
    # convert date column headers to YYYY/MM/DD strings for Excel
    new_cols = []
    for c in header_out.columns:
        if isinstance(c, (pd.Timestamp, datetime.datetime, datetime.date)):
            new_cols.append(c.strftime('%Y/%m/%d'))
        else:
            new_cols.append(c)
    header_out.columns = new_cols
    # reorder: 모델명, 헤더품번, 총소요량, then date columns
    date_cols_str = [c for c in header_out.columns if c not in ('모델명', '헤더품번', '총소요량')]
    header_out = header_out[['모델명', '헤더품번', '총소요량'] + date_cols_str]

    # Tube summary: required = core_qty * tube_count
    working['tube_req'] = working[qty_col] * working[tube_count_col]
    tube_grp = working.groupby([tube_col, 'date_only'])['tube_req'].sum().reset_index()
    tube_pivot = tube_grp.pivot_table(index=tube_col, columns='date_only', values='tube_req', aggfunc='sum', fill_value=0)
    # model per tube: pick first model where tube appears
    tube_model_map = {}
    if model_col and model_col in working.columns:
        tm = working.sort_values(date_col).groupby(tube_col)[model_col].first()
        tube_model_map = tm.to_dict()

    dates_t = sorted([c for c in tube_pivot.columns if isinstance(c, (datetime.date, datetime.datetime, pd.Timestamp))])
    past_sum_t = tube_pivot[[d for d in dates_t if d < asof_date]].sum(axis=1) if dates_t else pd.Series(0, index=tube_pivot.index)
    # include full daily date range between asof and last date (keep weekends)
    display_start_t = asof_date
    display_end_t = max(dates_t) if dates_t else display_start_t
    display_dates_t = pd.to_datetime(pd.date_range(start=display_start_t, end=display_end_t, freq='D')).normalize().tolist()
    display_df_t = tube_pivot.reindex(columns=display_dates_t).fillna(0)
    adj_df_t = pd.DataFrame(0, index=tube_pivot.index, columns=display_dates_t)
    for idx in tube_pivot.index:
        carry = past_sum_t.loc[idx] if isinstance(past_sum_t, pd.Series) else past_sum_t
        for d in display_dates_t:
            orig = display_df_t.at[idx, d] if d in display_df_t.columns else 0
            adj = orig + carry
            if adj < 0:
                adj_df_t.at[idx, d] = 0
                carry = adj
            else:
                adj_df_t.at[idx, d] = adj
                carry = 0
    tube_pivot[backlog_col_name] = past_sum_t.clip(lower=0)
    for d in display_dates_t:
        tube_pivot[d] = adj_df_t[d]
    keep_dates_t = [backlog_col_name] + display_dates_t
    tube_out = tube_pivot.reindex(columns=keep_dates_t).fillna(0)
    tube_out['총소요량'] = tube_out.sum(axis=1)
    tube_out.insert(0, '튜브플레이트품번', tube_out.index)
    tube_out.insert(0, '모델명', tube_out['튜브플레이트품번'].map(tube_model_map))
    # convert date headers to strings and reorder: 모델명, 튜브품번, 총소요량, dates...
    new_cols_t = []
    for c in tube_out.columns:
        if isinstance(c, (pd.Timestamp, datetime.datetime, datetime.date)):
            new_cols_t.append(c.strftime('%Y/%m/%d'))
        else:
            new_cols_t.append(c)
    tube_out.columns = new_cols_t
    date_cols_t_str = [c for c in tube_out.columns if c not in ('모델명', '튜브플레이트품번', '총소요량')]
    tube_out = tube_out[['모델명', '튜브플레이트품번', '총소요량'] + date_cols_t_str]

    return header_out.reset_index(drop=True), tube_out.reset_index(drop=True)


def save_excel(header_df: pd.DataFrame, tube_df: pd.DataFrame, outpath: Path):
    def _append_totals(df: pd.DataFrame, key_cols: list):
        if df.empty:
            return df
        cols = list(df.columns)
        # date columns are those after the key columns
        date_cols = [c for c in cols if c not in key_cols]
        # compute totals for numeric columns
        totals = {}
        for c in cols:
            if c in date_cols:
                totals[c] = df[c].sum()
            else:
                totals[c] = ""

        # place label
        totals[key_cols[0]] = "합계"
        # For 총소요량, ensure it's present and numeric sum
        if '총소요량' in df.columns:
            totals['총소요량'] = df['총소요량'].sum()

        # append
        return pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    hdr = _append_totals(header_df, ['모델명', '헤더품번'])
    tub = _append_totals(tube_df, ['모델명', '튜브플레이트품번'])

    with pd.ExcelWriter(outpath, engine='openpyxl') as writer:
        hdr.to_excel(writer, sheet_name='헤더소요량', index=False)
        tub.to_excel(writer, sheet_name='튜브소요량', index=False)


def main():
    root = tk.Tk()
    root.withdraw()

    plan_file = select_excel_file("생산계획 엑셀 선택", parent=root)
    if not plan_file:
        messagebox.showinfo("취소", "생산계획 엑셀을 선택하지 않았습니다.")
        return
    START_DATE, END_DATE = select_date_range(root)
    if START_DATE is None and END_DATE is None:
        # user chose no filter -> ask for as-of date for backlog behavior
        date_str = simpledialog.askstring("기준일", "미달(이전) 기준일을 입력하세요 (YYYY-MM-DD):", parent=root)
        if not date_str:
            messagebox.showinfo("취소", "기준일이 입력되지 않았습니다.")
            return
        try:
            asof_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            START_DATE = asof_date
            END_DATE = None
        except Exception:
            messagebox.showerror("오류", "날짜 형식이 잘못되었습니다. YYYY-MM-DD 형식으로 입력하세요.")
            return
    else:
        # normalize to date
        START_DATE = START_DATE.date() if isinstance(START_DATE, pd.Timestamp) or isinstance(START_DATE, datetime.datetime) else START_DATE
        END_DATE = END_DATE.date() if isinstance(END_DATE, pd.Timestamp) or isinstance(END_DATE, datetime.datetime) else END_DATE

    cfg = load_config()

    infile = Path(plan_file)
    # create progress window and pass updater into extractor
    win, update_progress = create_progress_window(root, title="CORE 추출 진행")
    try:
        outpath = export_core_tube_requirements(
            str(infile),
            asof_date_str=None,
            start_date_str=START_DATE.strftime('%Y-%m-%d') if START_DATE else None,
            end_date_str=END_DATE.strftime('%Y-%m-%d') if END_DATE else None,
            output_file=cfg.get('output'),
            progress=update_progress,
        )
    except Exception as e:
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{e}")
        try:
            win.destroy()
        except Exception:
            pass
        return
    finally:
        try:
            win.destroy()
        except Exception:
            pass

    messagebox.showinfo("완료", f"저장 완료: {outpath}")


if __name__ == '__main__':
    main()
