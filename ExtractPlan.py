from typing import Optional, List, Dict
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
import pandas as pd
from openpyxl import load_workbook
import datetime

from AppConstants import (
    DATA_START_ROW,
    DATE_HEADER_GAP_BREAK,
    DATE_SCAN_MIN_COL,
    HEADER_DATE_ROW,
    HEADER_KIND_ROW,
    MES_SHEET_CONFIG,
    PLAN_OUTPUT_FILE,
    PLAN_SHEET_NAMES,
    WORKSHOP_TO_PROCESS,
    EXTRACT_STOCK,
    STOCK_COL,
    CORE_TUBE_OUTPUT_FILE,
    CORE_TUBE_HEADER_COL,
    CORE_TUBE_TUBE_COL,
    CORE_TUBE_COUNT_COL,
)
from CommonUtils import (
    build_merged_map,
    header_text,
    is_date_value,
    normalize_kind,
    normalize_part_no,
    normalize_process_token,
    safe_float,
)

# =========================
# 설정
# =========================
SHEET_NAME = PLAN_SHEET_NAMES
OUTPUT_FILE = PLAN_OUTPUT_FILE


# =========================
# 열 찾기
# =========================
def find_date_start_col(ws) -> Optional[int]:
    for c in range(DATE_SCAN_MIN_COL, ws.max_column + 1):
        if is_date_value(ws.cell(HEADER_DATE_ROW, c).value):
            return c
    return None


def find_col_by_keywords(ws, keywords, required: bool = False, date_start_col: Optional[int] = None) -> Optional[int]:
    keys = [k.upper().replace(" ", "") for k in keywords]
    date_start_col = date_start_col or find_date_start_col(ws) or (ws.max_column + 1)

    for c in range(1, date_start_col):
        txt = header_text(ws, HEADER_DATE_ROW, c)
        txt = txt.upper().replace(" ", "").replace("\n", "")

        if all(k in txt for k in keys):
            return c

    if required:
        raise RuntimeError(f"{ws.title}: 헤더를 찾지 못했습니다. keywords={keywords}")

    return None


def build_header_lookup(ws) -> Dict[str, Optional[int]]:
    date_start_col = find_date_start_col(ws) or (ws.max_column + 1)
    return {
        "model_col": find_col_by_keywords(ws, ["모델"], required=True, date_start_col=date_start_col),
        "customer_part_col": find_col_by_keywords(ws, ["고객사", "품번"], date_start_col=date_start_col),
        "core_part_col": find_col_by_keywords(ws, ["CORE", "품번"], date_start_col=date_start_col),
        "tank_part_col": find_col_by_keywords(ws, ["TANK", "품번"], date_start_col=date_start_col),
        "finish_part_col": find_col_by_keywords(ws, ["완성", "품번"], date_start_col=date_start_col),
        "accessory_part_col": find_col_by_keywords(ws, ["액세", "품번"], date_start_col=date_start_col),
        "process_print_col": find_col_by_keywords(ws, ["공정", "인쇄"], date_start_col=date_start_col),
    }


def build_plan_date_columns(ws) -> List[Dict]:
    """
    날짜 열 중에서 구분행이 '계획'인 열만 추출.
    미달 / 실적은 스킵.
    """
    cols = []
    started = False
    invalid_streak = 0

    start_col = find_date_start_col(ws)
    if start_col is None:
        raise RuntimeError(f"{ws.title}: 날짜 시작 열을 찾지 못했습니다.")

    for c in range(start_col, ws.max_column + 1):
        date_val = ws.cell(HEADER_DATE_ROW, c).value
        kind_val = normalize_process_token(ws.cell(HEADER_KIND_ROW, c).value).replace("\n", "")

        if is_date_value(date_val):
            started = True
            invalid_streak = 0

            # 핵심: 미달 / 실적은 스킵, 계획만 읽음
            if kind_val != "계획":
                continue

            cols.append({
                "col": c,
                "date": pd.to_datetime(date_val).normalize(),
                "kind": kind_val,
            })
            continue

        if started:
            invalid_streak += 1
            if invalid_streak >= DATE_HEADER_GAP_BREAK:
                break

    return cols


def get_sheet_columns(ws) -> Dict[str, Optional[int]]:
    cols = build_header_lookup(ws)
    if not cols["finish_part_col"]:
        raise RuntimeError(f"{ws.title}: 헤더를 찾지 못했습니다. keywords={['완성', '품번']}")
    return cols


# =========================
# 스킵 조건
# =========================


def is_skip_summary_row(*values) -> bool:
    text = " ".join([normalize_process_token(v) for v in values if normalize_process_token(v)])

    if not text:
        return True

    skip_keywords = ["합계", "소계", "TOTAL", "누계"]

    return any(k in text.upper() for k in skip_keywords)


def should_skip_row(part_no, process_name, row_text="", sheet_title: str = "", product_key=None) -> bool:
    part_txt = normalize_process_token(part_no)
    product_txt = normalize_process_token(product_key if product_key is not None else part_no)
    process_txt = normalize_process_token(process_name)
    row_txt = normalize_process_token(row_text)

    if part_txt.startswith("HH") or product_txt.startswith("HH"):
        return True

    if part_txt == "2":
        return True

    if sheet_title == "TANK공정(실적)" and "클린칭C" in process_txt:
        return True

    if "개발,기타" in process_txt or "개발,기타" in row_txt:
        return True

    return False


def starts_with_alpha_part_no(part_no) -> bool:
    part_txt = normalize_process_token(part_no)
    return bool(part_txt) and part_txt[0].isalpha() and part_txt[0].isascii()


# =========================
# 생산계획 추출
# =========================
def extract_finish_plan_sheet(ws) -> pd.DataFrame:
    return _extract_plan_sheet(ws, mode="finish")


def build_mes_date_columns(ws) -> List[Dict]:
    cols = []
    started = False
    invalid_streak = 0
    start_col = find_date_start_col(ws)

    if start_col is None:
        return []

    for c in range(start_col, ws.max_column + 1):
        date_val = ws.cell(HEADER_DATE_ROW, c).value
        kind_val = ws.cell(HEADER_KIND_ROW, c).value

        if is_date_value(date_val):
            started = True
            invalid_streak = 0
            kind = normalize_kind(kind_val) or "계획"
            cols.append({
                "col": c,
                "date": pd.to_datetime(date_val).normalize(),
                "kind": kind,
            })
            continue

        if started:
            invalid_streak += 1
            if invalid_streak >= DATE_HEADER_GAP_BREAK:
                break

    return cols


def get_mes_sheet_columns(ws) -> Dict[str, Optional[int]]:
    return build_header_lookup(ws)


def is_red_font_cell(ws, row: int, col: int) -> bool:
    cell = ws.cell(row, col)
    color = getattr(cell.font, "color", None)
    if color is None:
        return False

    rgb = getattr(color, "rgb", None)
    if isinstance(rgb, str):
        rgb = rgb.upper()
        if rgb in {"FFFF0000", "00FF0000", "FF0000"}:
            return True

    return False


def map_workcenter_and_team(process_name):
    p_raw = normalize_process_token(process_name)
    p = normalize_process_token(p_raw)

    if "AL" == p or "AL" in p:
        return "CORE조립", "CORE조립-AL"
    if "CU" == p or "CU" in p:
        return "CORE조립", "CORE조립-CU"
    if "수동" in p_raw:
        return "용접", "TANK용접 - 수동"
    if "로봇" in p:
        return "용접", "TANK용접 - ROBOT"
    if "클린칭" in p:
        return "CLINCHING", "CLINCHING"
    if "TANK조립" in p_raw:
        return "TANK조립&Leak Test", "TANK조립&Leak Test"
    if "한국" in p_raw:
        return "용접", "TANK용접 - 한국RAD"
    if "일반" in p_raw:
        return "완성조립", "완성조립공정 - 일반"
    if "G2" in p:
        return "완성조립", "완성조립공정 - G2"
    if "클라크" in p_raw:
        return "완성조립", "완성조립공정 - 클라크"
    if "특수품" in p_raw:
        return "완성조립", "완성조립공정 - 특수품"
    if "액세" in p_raw or "HEATSCREEN" in p:
        return "출하-액세사리", "출하-액세사리"
    return None, None


def map_process_from_workshop(name: str) -> Optional[str]:
    return WORKSHOP_TO_PROCESS.get(normalize_process_token(name))


def pick_plan_part_no(workshop_name: str, row_values: Dict[str, str]) -> str:
    w = normalize_process_token(workshop_name)

    if w == "CORE조립":
        return row_values.get("core_part_no", "")
    if w in ("용접", "CLINCHING", "TANK조립&LEAKTEST"):
        return row_values.get("tank_part_no", "")
    if w == "완성조립":
        return row_values.get("finish_part_no", "")
    if "액세" in w:
        return row_values.get("accessory_part_no", "") or row_values.get("finish_part_no", "")

    return (
        row_values.get("finish_part_no", "")
        or row_values.get("tank_part_no", "")
        or row_values.get("core_part_no", "")
    )


def make_compare_key(process: str, workshop: str, team: str, part_no: str, day) -> str:
    process = normalize_process_token(process)
    workshop = normalize_process_token(workshop)
    team = normalize_process_token(team)
    part_no = normalize_part_no(part_no)
    day = pd.to_datetime(day, errors="coerce").normalize()

    if process == "완성" and "액세" in workshop:
        return f"{process}|{part_no}|{day}"

    return f"{process}|{workshop}|{team}|{part_no}|{day}"


def extract_plan_sheet(ws, config: Dict, check_red_font: bool = False, progress=None) -> pd.DataFrame:
    return _extract_plan_sheet(ws, mode="mes", config=config, check_red_font=check_red_font, progress=progress)


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


def filter_work_order_period(df: pd.DataFrame, start_date=None, end_date=None) -> pd.DataFrame:
    if df.empty:
        return df

    date_col = df.columns[8]
    out = df.copy()
    out[date_col] = pd.to_datetime(out[date_col], errors="coerce").dt.normalize()

    if start_date is not None:
        out = out[out[date_col] >= start_date]
    if end_date is not None:
        out = out[out[date_col] <= end_date]

    return out.copy()


def extract_work_order_file(plan_file: str, progress=None) -> pd.DataFrame:
    if progress:
        progress(8, "생산계획 통합 문서를 여는 중...")

    wb = load_workbook(plan_file, data_only=True)
    frames = []
    sheet_items = [
        (sheet_name, config)
        for sheet_name, config in MES_SHEET_CONFIG.items()
        if sheet_name in wb.sheetnames
    ]
    total = max(len(sheet_items), 1)

    for i, (sheet_name, config) in enumerate(sheet_items, start=1):
        if progress:
            start = 10 + int((i - 1) / total * 65)
            progress(start, f"{sheet_name} 시트 추출 중...")

        def sheet_progress(percent, text=None, i=i, sheet_name=sheet_name):
            if progress:
                base = 10 + ((i - 1) / total * 65)
                span = 65 / total
                progress(base + (percent / 100 * span), text or f"{sheet_name} 시트 추출 중...")

        df_sheet = extract_plan_sheet(
            wb[sheet_name],
            config,
            check_red_font=True,
            progress=sheet_progress,
        )
        if not df_sheet.empty:
            frames.append(df_sheet)

        if progress:
            done = 10 + int(i / total * 65)
            progress(done, f"{sheet_name} 시트 추출 완료")

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def build_work_order_upload_df(plan_df: pd.DataFrame) -> pd.DataFrame:
    output_columns = ["작업장명", "작업반명", "모델", "품번", "공정 인쇄"]

    if plan_df.empty:
        return pd.DataFrame(columns=output_columns)

    cols = list(plan_df.columns)
    workcenter_col = cols[3]
    team_col = cols[4]
    model_col = cols[5]
    part_col = cols[6]
    date_col = cols[8]
    kind_col = cols[9]
    qty_col = cols[10]
    process_print_col = cols[11]

    df = plan_df.copy()
    plan_kind = normalize_kind("계획") or "계획"
    df = df[df[kind_col].apply(normalize_kind).fillna(df[kind_col]) == plan_kind].copy()
    if df.empty:
        return pd.DataFrame(columns=output_columns)

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.normalize()
    df[part_col] = df[part_col].apply(normalize_part_no)
    df = df[~df[part_col].isin(["", "NAN", "NONE"])].copy()
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df = df[df[qty_col] != 0].copy()

    def first_non_empty(series):
        for value in series:
            if pd.notna(value) and str(value).strip():
                return value
        return ""

    grouped = (
        df.groupby(
            [part_col, date_col],
            as_index=False,
            dropna=False,
        )
        .agg(
            **{
                workcenter_col: (workcenter_col, first_non_empty),
                team_col: (team_col, first_non_empty),
                model_col: (model_col, first_non_empty),
                process_print_col: (process_print_col, first_non_empty),
                qty_col: (qty_col, "sum"),
            }
        )
    )

    wide = grouped.pivot_table(
        index=[part_col],
        columns=date_col,
        values=qty_col,
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    metadata = (
        grouped.groupby(part_col, as_index=False, dropna=False)
        .agg(
            **{
                workcenter_col: (workcenter_col, first_non_empty),
                team_col: (team_col, first_non_empty),
                model_col: (model_col, first_non_empty),
                process_print_col: (process_print_col, first_non_empty),
            }
        )
    )
    wide = metadata.merge(wide, on=part_col, how="left")

    date_columns = sorted([c for c in wide.columns if isinstance(c, pd.Timestamp)])
    wide = wide[[workcenter_col, team_col, model_col, part_col, process_print_col] + date_columns]
    wide = wide.rename(
        columns={
            workcenter_col: "작업장명",
            team_col: "작업반명",
            model_col: "모델",
            part_col: "품번",
            process_print_col: "공정 인쇄",
        }
    )
    wide = wide.rename(columns={c: c.strftime("%Y/%m/%d") for c in date_columns})

    qty_cols = [c for c in wide.columns if c not in output_columns]
    wide = wide[(wide[qty_cols] != 0).any(axis=1)].copy()
    return wide.reset_index(drop=True)


def _extract_plan_sheet(ws, mode: str, config: Optional[Dict] = None, check_red_font: bool = False, progress=None) -> pd.DataFrame:
    is_mes = mode == "mes"
    config = config or {}
    merged_map = build_merged_map(ws)
    date_cols = build_mes_date_columns(ws) if is_mes else build_plan_date_columns(ws)

    if not date_cols:
        raise RuntimeError(f"{ws.title}: 날짜/구분 컬럼을 찾지 못했습니다.")

    cols = get_mes_sheet_columns(ws) if is_mes else get_sheet_columns(ws)
    records = []
    current_model = ""
    accessory_mode = False
    max_col = ws.max_column
    model_col = cols["model_col"]
    customer_part_col = cols["customer_part_col"]
    core_part_col = cols["core_part_col"] if is_mes else None
    tank_part_col = cols["tank_part_col"] if is_mes else None
    finish_part_col = cols["finish_part_col"]
    accessory_part_col = cols["accessory_part_col"] if is_mes else None
    process_print_col = cols["process_print_col"]
    stock_col = STOCK_COL if EXTRACT_STOCK else None

    for row_cells in ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row, max_col=max_col):
        r = row_cells[0].row
        if is_mes and (r - DATA_START_ROW) % 200 == 0:
            percent = (r - DATA_START_ROW) / max(ws.max_row - DATA_START_ROW, 1) * 100
            if progress:
                progress(percent, f"{ws.title} 계획 추출 중... {percent:.1f}%")

        values = [cell.value for cell in row_cells]

        model_val = normalize_process_token(values[model_col - 1])
        if model_val:
            current_model = model_val
        model = current_model
        row_text = " ".join(
            text
            for text in (normalize_process_token(cell.value) for cell in row_cells)
            if text
        )

        row_values = {
            "customer_part_no": normalize_process_token(values[customer_part_col - 1]) if customer_part_col else "",
            "core_part_no": normalize_process_token(values[core_part_col - 1]) if core_part_col else "",
            "tank_part_no": normalize_process_token(values[tank_part_col - 1]) if tank_part_col else "",
            "finish_part_no": normalize_process_token(values[finish_part_col - 1]) if finish_part_col else "",
            "accessory_part_no": normalize_process_token(values[accessory_part_col - 1]) if accessory_part_col else "",
        }

        first_col_raw = merged_map.get((r, 1), values[0])
        first_col_text = normalize_process_token(first_col_raw)

        if ws.title == "완성공정(실적)":
            finish_part_no = row_values.get("finish_part_no", "") or ""
            if "용접C/M" in first_col_text or "코어C/M" in first_col_text or "선발주-용접C/M" in first_col_text:
                if "-반" not in finish_part_no:
                    continue

        if ws.title == "TANK공정(실적)" and first_col_text.startswith("코어C/M"):
            continue

        row_text_norm = normalize_process_token(row_text)
        if "액세" in row_text_norm and "HEATSCREEN" in row_text_norm:
            accessory_mode = True
            continue

        if accessory_mode:
            stop_prefix = "단품" if is_mes else "제외"
            if first_col_text.startswith(stop_prefix):
                break

        if "제외" in first_col_text:
            continue

        process_name_raw = ""
        if accessory_mode and ws.title == "완성공정(실적)":
            process_name_raw = "액세사리 & HEAT SCREEN"
        elif process_print_col:
            process_name_raw = normalize_process_token(values[process_print_col - 1])

        if is_mes:
            if not process_name_raw:
                process_name_raw = config.get("default_process", "")

            workshop_name, team_name = map_workcenter_and_team(process_name_raw)

            if workshop_name and "액세" in workshop_name:
                workshop_name = "출하-액세사리"
                team_name = "출하-액세사리"

            if not workshop_name:
                if config.get("default_process") == "CORE":
                    workshop_name, team_name = "CORE조립", ""
                elif config.get("default_process") == "TANK":
                    workshop_name, team_name = "용접", ""
                elif config.get("default_process") == "완성":
                    workshop_name, team_name = "완성조립", ""

            part_no = pick_plan_part_no(workshop_name, row_values)
            product_key = row_values.get("customer_part_no") or part_no
        else:
            part_no = row_values["finish_part_no"]
            product_key = row_values.get("customer_part_no") or part_no

        if should_skip_row(part_no, process_name_raw, row_text, sheet_title=ws.title, product_key=product_key):
            continue

        if is_skip_summary_row(model, part_no, row_text):
            continue

        if not model or not part_no:
            continue

        if not starts_with_alpha_part_no(part_no):
            continue

        for dc in date_cols:
            if is_mes and check_red_font and is_red_font_cell(ws, r, dc["col"]):
                continue

            qty = values[dc["col"] - 1]
            if qty is None or qty == "":
                continue

            qty_num = safe_float(qty)
            if qty_num == 0:
                continue

            # 재고 읽기 (행 단위)
            stock_val = None
            if stock_col and 1 <= stock_col <= len(values):
                stock_raw = values[stock_col - 1]
                stock_val = safe_float(stock_raw)

            if is_mes:
                process = map_process_from_workshop(workshop_name)
                workshop_raw = workshop_name or ""
                workshop = normalize_process_token(workshop_name)
                team = team_name or ""
                part_norm = normalize_part_no(part_no)

                records.append({
                    "비교키": make_compare_key(process, workshop, team, part_norm, dc["date"]),
                    "시트명": ws.title,
                    "공정": process,
                    "작업장명": workshop_raw,
                    "작업반명": team,
                    "모델": model,
                    "품번": part_norm,
                    "고객사품번": row_values.get("customer_part_no", ""),
                    "날짜": dc["date"],
                    "구분": dc["kind"],
                    "수량": qty_num,
                    "공정인쇄": process_name_raw,
                    "원본행": r,
                    "원본열": dc["col"],
                    "재고": stock_val,
                })
            else:
                records.append({
                    "시트명": ws.title,
                    "모델": model,
                    "품번": normalize_part_no(part_no),
                    "고객사품번": row_values.get("customer_part_no", ""),
                    "날짜": dc["date"],
                    "구분": dc["kind"],
                    "계획수량": qty_num,
                    "공정인쇄": process_name_raw,
                    "재고": stock_val,
                })

    df = pd.DataFrame(records)
    if df.empty or not is_mes:
        return df

    return df[df["공정"].notna()].copy()


def extract_plan_file(plan_file: str, progress=None) -> pd.DataFrame:
    wb = load_workbook(plan_file, data_only=True)

    frames = []

    total = len(SHEET_NAME)
    for i, sheet_name in enumerate(SHEET_NAME, start=1):
        if sheet_name not in wb.sheetnames:
            if progress:
                progress(10 + int(i / total * 70), f"{sheet_name} 시트 없음, 건너뜀")
            continue

        if progress:
            progress(10 + int((i - 1) / total * 70), f"{sheet_name} 계획 추출 중...")

        ws = wb[sheet_name]
        df = extract_finish_plan_sheet(ws)

        if progress:
            progress(10 + int(i / total * 70), f"{sheet_name} 계획 추출 완료: {len(df)}건")
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def _col_letter_to_index(letter: str) -> int:
    letter = (letter or "").upper().strip()
    idx = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'):
            continue
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx


def extract_core_tube_from_ws(ws, asof_date=None, start_date=None, end_date=None, progress=None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Extract header and tube requirement summaries from a CORE sheet.

    Returns (header_df, tube_df)
    """
    start_col = find_date_start_col(ws)
    if start_col is None:
        raise RuntimeError(f"{ws.title}: 날짜 시작 열을 찾지 못했습니다.")

    # 계획 열과 사용자가 선택한 날짜의 미달 열을 함께 읽는다.
    # 미달은 계획 수량과 섞지 않고 아래에서 별도로 집계한다.
    asof_calc = asof_date or (start_date if start_date is not None else None)
    selected_date = pd.to_datetime(asof_calc).normalize() if asof_calc is not None else None
    date_cols = [
        dc for dc in build_mes_date_columns(ws)
        if dc["kind"] == "계획"
        or (
            selected_date is not None
            and dc["kind"] == "미달"
            and dc["date"] == selected_date
        )
    ]
    if not date_cols:
        raise RuntimeError(f"{ws.title}: 날짜 컬럼을 찾지 못했습니다.")

    # column indexes: AppConstants give letters
    h_idx = _col_letter_to_index(CORE_TUBE_HEADER_COL)
    t_idx = _col_letter_to_index(CORE_TUBE_TUBE_COL)
    tc_idx = _col_letter_to_index(CORE_TUBE_COUNT_COL)

    # note: _col_letter_to_index returns 1-based index (A->1)
    records = []
    current_model = ""
    model_col = build_header_lookup(ws).get("model_col")

    total_rows = max(1, ws.max_row - DATA_START_ROW + 1)
    min_col = 1
    max_col = ws.max_column
    nf = normalize_process_token
    sf = safe_float
    date_normalize = lambda d: pd.to_datetime(d).normalize().date()

    rows_iter = ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row, min_col=min_col, max_col=max_col, values_only=True)
    for i, row_vals in enumerate(rows_iter, start=1):
        r = DATA_START_ROW + i - 1
        # report progress occasionally
        if progress and (i % 500 == 0 or i == 1):
            pct = int(i / total_rows * 70)
            progress(pct, f"행 {r} 처리중...")

        try:
            if model_col and row_vals[model_col - 1]:
                mval = nf(row_vals[model_col - 1])
                if mval:
                    current_model = mval
        except Exception:
            pass

        header_part = nf(row_vals[h_idx - 1]) if h_idx and h_idx <= max_col and row_vals[h_idx - 1] is not None else ""
        tube_part = nf(row_vals[t_idx - 1]) if t_idx and t_idx <= max_col and row_vals[t_idx - 1] is not None else ""
        try:
            tube_count = sf(row_vals[tc_idx - 1]) if tc_idx and tc_idx <= max_col and row_vals[tc_idx - 1] is not None else 0
        except Exception:
            tube_count = 0

        if not header_part and not tube_part:
            continue

        for dc in date_cols:
            col_idx = dc["col"] - 1
            if col_idx < 0 or col_idx >= len(row_vals):
                continue
            qty_raw = row_vals[col_idx]
            qty = sf(qty_raw)
            if qty == 0:
                continue

            records.append({
                "모델": current_model,
                "헤더품번": header_part,
                "튜브플레이트품번": tube_part,
                "튜브개수": tube_count,
                "날짜": date_normalize(dc["date"]),
                "구분": dc["kind"],
                "수량": qty,
            })

    if progress:
        progress(75, "요약 집계 중...")

    if not records:
        return pd.DataFrame(), pd.DataFrame()

    df = pd.DataFrame(records)

    plan_df = df[df["구분"] == "계획"].copy()
    shortage_df = df[df["구분"] == "미달"].copy()

    # Header: 계획 수량*2. 선택 날짜의 미달은 별도 집계한다.
    df_header = plan_df.groupby(["헤더품번", "모델", "날짜"], dropna=False, as_index=False)["수량"].sum()
    df_header["수량"] = df_header["수량"] * 2
    header_wide = df_header.pivot_table(index=["헤더품번"], columns="날짜", values="수량", aggfunc="sum", fill_value=0)

    # model map: first model per header
    model_map = df.groupby("헤더품번")["모델"].first().to_dict()

    dates = sorted([c for c in header_wide.columns if isinstance(c, (pd.Timestamp, datetime.datetime, datetime.date))])
    backlog_name = "미달(이전)"
    if asof_calc is not None:
        header_shortage = shortage_df.groupby("헤더품번")["수량"].sum().mul(2)
        asof_day_value = header_shortage.reindex(header_wide.index, fill_value=0)
        # determine display date range (include weekends by generating full daily range)
        display_start = start_date if start_date is not None else asof_calc
        display_end = end_date if end_date is not None else (max(dates) if dates else display_start)
        display_dates = pd.to_datetime(pd.date_range(start=display_start, end=display_end, freq='D')).normalize().tolist()
        # build adjusted columns: apply asof_date value's carry starting from asof_date forward
        display_df = header_wide.reindex(columns=display_dates).fillna(0)
        adj_df = pd.DataFrame(0, index=header_wide.index, columns=display_dates)
        for idx in header_wide.index:
            # 음수 미달(선생산)만 이후 계획에서 날짜순으로 차감한다.
            carry = min(asof_day_value.loc[idx], 0)
            for d in display_dates:
                orig = display_df.at[idx, d] if d in display_df.columns else 0
                adj = orig + carry
                if adj < 0:
                    adj_df.at[idx, d] = 0
                    carry = adj
                else:
                    adj_df.at[idx, d] = adj
                    carry = 0
        # backlog: show asof_date day's value as positive (미달) only
        header_wide[backlog_name] = asof_day_value.clip(lower=0)
        # replace display date columns with adjusted values
        for d in display_dates:
            header_wide[d] = adj_df[d]
    else:
        header_wide[backlog_name] = 0
        # include full continuous date range between min/max dates to keep weekends
        if dates:
            display_dates = pd.to_datetime(pd.date_range(start=min(dates), end=max(dates), freq='D')).normalize().tolist()
        else:
            display_dates = []

    keep_cols = [backlog_name] + display_dates
    header_out = header_wide.reindex(columns=keep_cols).fillna(0)
    header_out["총소요량"] = header_out.sum(axis=1)
    header_out.insert(0, "헤더품번", header_out.index)
    header_out.insert(0, "모델명", header_out["헤더품번"].map(model_map))
    # reorder so '총소요량' is placed before the date columns
    date_cols_after = [c for c in header_out.columns if c not in ("모델명", "헤더품번", "총소요량")]
    header_out = header_out[["모델명", "헤더품번", "총소요량"] + date_cols_after]
    header_out = header_out.reset_index(drop=True)

    # Tube: sum 수량 * 튜브개수 (rename to 튜브플레이트)
    plan_df["튜브플레이트요청"] = plan_df["수량"] * plan_df["튜브개수"]
    shortage_df["튜브플레이트요청"] = shortage_df["수량"] * shortage_df["튜브개수"]
    df_tube = plan_df.groupby(["튜브플레이트품번", "모델", "날짜"], dropna=False, as_index=False)["튜브플레이트요청"].sum()
    tube_wide = df_tube.pivot_table(index=["튜브플레이트품번"], columns="날짜", values="튜브플레이트요청", aggfunc="sum", fill_value=0)
    tube_model_map = df.groupby("튜브플레이트품번")["모델"].first().to_dict()

    dates_t = sorted([c for c in tube_wide.columns if isinstance(c, (pd.Timestamp, datetime.datetime, datetime.date))])
    if asof_calc is not None:
        tube_shortage = shortage_df.groupby("튜브플레이트품번")["튜브플레이트요청"].sum()
        asof_day_value_t = tube_shortage.reindex(tube_wide.index, fill_value=0)
        display_start_t = start_date if start_date is not None else asof_calc
        display_end_t = end_date if end_date is not None else (max(dates_t) if dates_t else display_start_t)
        display_dates_t = pd.to_datetime(pd.date_range(start=display_start_t, end=display_end_t, freq='D')).normalize().tolist()
        display_df_t = tube_wide.reindex(columns=display_dates_t).fillna(0)
        adj_df_t = pd.DataFrame(0, index=tube_wide.index, columns=display_dates_t)
        for idx in tube_wide.index:
            # 음수 미달(선생산)만 이후 계획에서 날짜순으로 차감한다.
            carry = min(asof_day_value_t.loc[idx], 0)
            for d in display_dates_t:
                orig = display_df_t.at[idx, d] if d in display_df_t.columns else 0
                adj = orig + carry
                if adj < 0:
                    adj_df_t.at[idx, d] = 0
                    carry = adj
                else:
                    adj_df_t.at[idx, d] = adj
                    carry = 0
        tube_wide[backlog_name] = asof_day_value_t.clip(lower=0)
        for d in display_dates_t:
            tube_wide[d] = adj_df_t[d]
    else:
        tube_wide[backlog_name] = 0
        if dates_t:
            display_dates_t = pd.to_datetime(pd.date_range(start=min(dates_t), end=max(dates_t), freq='D')).normalize().tolist()
        else:
            display_dates_t = []

    keep_t = [backlog_name] + display_dates_t
    tube_out = tube_wide.reindex(columns=keep_t).fillna(0)
    tube_out["총소요량"] = tube_out.sum(axis=1)
    tube_out.insert(0, "튜브플레이트품번", tube_out.index)
    tube_out.insert(0, "모델명", tube_out["튜브플레이트품번"].map(tube_model_map))
    # reorder so '총소요량' is placed before the date columns
    date_cols_after_t = [c for c in tube_out.columns if c not in ("모델명", "튜브플레이트품번", "총소요량")]
    tube_out = tube_out[["모델명", "튜브플레이트품번", "총소요량"] + date_cols_after_t]
    tube_out = tube_out.reset_index(drop=True)

    # rename date columns to string YYYY/MM/DD for consistency
    def _rename_dates(df_obj):
        new_cols = []
        for c in df_obj.columns:
            if isinstance(c, (pd.Timestamp, datetime.datetime)):
                new_cols.append(c.strftime("%Y/%m/%d"))
            elif isinstance(c, datetime.date):
                new_cols.append(c.strftime("%Y/%m/%d"))
            else:
                new_cols.append(c)
        df_obj.columns = new_cols

    if progress:
        progress(90, "날짜 열 포맷 변환 중...")
    _rename_dates(header_out)
    _rename_dates(tube_out)

    if progress:
        progress(100, "완료")

    return header_out, tube_out


def export_core_tube_requirements(plan_file: str, asof_date_str: str = None, start_date_str: str = None, end_date_str: str = None, output_file: str = None, progress=None):
    """High-level helper to export core/tube requirements from a plan file.

    asof_date_str: YYYY-MM-DD
    """
    asof_date = pd.to_datetime(asof_date_str).date() if asof_date_str else None
    start_date = pd.to_datetime(start_date_str).normalize().date() if start_date_str else None
    end_date = pd.to_datetime(end_date_str).normalize().date() if end_date_str else None
    wb = load_workbook(plan_file, data_only=True)
    sheet_name = "CORE공정(실적)"
    if sheet_name not in wb.sheetnames:
        raise RuntimeError(f"{sheet_name} 시트가 없습니다: {plan_file}")

    header_df, tube_df = extract_core_tube_from_ws(wb[sheet_name], asof_date=asof_date, start_date=start_date, end_date=end_date, progress=progress)
    out = output_file or CORE_TUBE_OUTPUT_FILE
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        header_df.to_excel(writer, sheet_name="헤더소요량", index=False)
        tube_df.to_excel(writer, sheet_name="튜브플레이트소요량", index=False)

    return out


# =========================
# 실행
# =========================
def main():
    root = tk.Tk()
    root.withdraw()

    plan_file = filedialog.askopenfilename(
        title="생산계획 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not plan_file:
        raise SystemExit("생산계획 엑셀을 선택하지 않았습니다.")

    plan_df = extract_plan_file(plan_file)

    if plan_df.empty:
        messagebox.showinfo("알림", "추출된 계획 수량이 없습니다.")
        return

    plan_df.to_excel(OUTPUT_FILE, index=False)
    messagebox.showinfo("완료", f"저장 완료: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
