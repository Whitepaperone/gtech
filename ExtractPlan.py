from datetime import datetime, date
from typing import Optional, List, Dict
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook


# =========================
# 설정
# =========================
SHEET_NAME = ["완성공정(실적)","TANK공정(실적)","CORE공정(실적)"]
OUTPUT_FILE = "./생산계획_계획수량_추출.xlsx"

HEADER_DATE_ROW = 33   # 날짜 행
HEADER_KIND_ROW = 34   # 구분 행: 미달 / 계획 / 실적
DATA_START_ROW = 35

DATE_SCAN_MIN_COL = 1
DATE_HEADER_GAP_BREAK = 3


# =========================
# 공통 함수
# =========================
def normalize_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def normalize_compact(v) -> str:
    return normalize_text(v).upper().replace(" ", "").replace("\n", "")


def is_date_value(v) -> bool:
    return isinstance(v, (datetime, date))


def safe_float(v) -> float:
    n = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(n) else float(n)


def normalize_part_no(v) -> str:
    t = normalize_text(v).upper()
    t = t.replace(" ", "").replace("_", "")
    return t


def header_text(ws, row: int, col: int) -> str:
    a = normalize_text(ws.cell(row, col).value)
    b = normalize_text(ws.cell(row + 1, col).value)
    return f"{a} {b}".strip()


def row_join_text(ws, r: int) -> str:
    vals = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
    vals = [v for v in vals if v]
    return " ".join(vals)


def get_merged_value(ws, row: int, col: int):
    cell = ws.cell(row, col)
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            return ws.cell(merged.min_row, merged.min_col).value
    return cell.value


# =========================
# 열 찾기
# =========================
def find_date_start_col(ws) -> Optional[int]:
    for c in range(DATE_SCAN_MIN_COL, ws.max_column + 1):
        if is_date_value(ws.cell(HEADER_DATE_ROW, c).value):
            return c
    return None


def find_col_by_keywords(ws, keywords, required: bool = False) -> Optional[int]:
    keys = [k.upper().replace(" ", "") for k in keywords]
    date_start_col = find_date_start_col(ws) or (ws.max_column + 1)

    for c in range(1, date_start_col):
        txt = header_text(ws, HEADER_DATE_ROW, c)
        txt = txt.upper().replace(" ", "").replace("\n", "")

        if all(k in txt for k in keys):
            return c

    if required:
        raise RuntimeError(f"{ws.title}: 헤더를 찾지 못했습니다. keywords={keywords}")

    return None


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
        kind_val = normalize_text(ws.cell(HEADER_KIND_ROW, c).value).replace("\n", "")

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
    return {
        "model_col": find_col_by_keywords(ws, ["모델"], required=True),
        "customer_part_col": find_col_by_keywords(ws, ["고객사", "품번"]),
        "finish_part_col": find_col_by_keywords(ws, ["완성", "품번"], required=True),
        "process_print_col": find_col_by_keywords(ws, ["공정", "인쇄"]),
    }


# =========================
# 스킵 조건
# =========================


def is_skip_summary_row(*values) -> bool:
    text = " ".join([normalize_text(v) for v in values if normalize_text(v)])

    if not text:
        return True

    skip_keywords = ["합계", "소계", "TOTAL", "누계"]

    return any(k in text.upper() for k in skip_keywords)


def should_skip_row(part_no, process_name, row_text) -> bool:
    part_txt = normalize_text(part_no).upper()
    process_txt = normalize_text(process_name)
    row_txt = normalize_text(row_text)

    if part_txt.startswith("HH"):
        return True

    if part_txt == "2":
        return True

    if "개발,기타" in process_txt or "개발,기타" in row_txt:
        return True

    return False


# =========================
# 생산계획 추출
# =========================
def extract_finish_plan_sheet(ws) -> pd.DataFrame:
    date_cols = build_plan_date_columns(ws)
    cols = get_sheet_columns(ws)

    records = []
    current_model = ""
    accessory_mode = False

    for r in range(DATA_START_ROW, ws.max_row + 1):
        model_val = normalize_text(ws.cell(r, cols["model_col"]).value)

        if model_val:
            current_model = model_val

        model = current_model
        row_text = row_join_text(ws, r)

        first_col_raw = get_merged_value(ws, r, 1)
        first_col_text = normalize_compact(first_col_raw)

        # 완성공정 시트 내부의 다른 구역 스킵
        if (
            "용접C/M" in first_col_text
            or "코어C/M" in first_col_text
            or "선발주-용접C/M" in first_col_text
        ):
            continue

        compact_row = normalize_compact(row_text)

        # 액세서리 영역 처리
        if "액세서리&HEATSCREEN" in compact_row:
            accessory_mode = True
            continue

        if accessory_mode:
            if first_col_text.startswith("단품"):
                break

        finish_part_no = normalize_text(ws.cell(r, cols["finish_part_col"]).value)
        customer_part_no = (
            normalize_text(ws.cell(r, cols["customer_part_col"]).value)
            if cols["customer_part_col"]
            else ""
        )

        process_print = (
            normalize_text(ws.cell(r, cols["process_print_col"]).value)
            if cols["process_print_col"]
            else ""
        )

        if accessory_mode:
            process_print = "액세서리 & HEAT SCREEN"

        part_no = finish_part_no

        if is_skip_summary_row(model, part_no, row_text):
            continue

        if should_skip_row(part_no, process_print, row_text):
            continue

        if not model or not part_no:
            continue

        for dc in date_cols:
            qty = ws.cell(r, dc["col"]).value

            if qty is None or qty == "":
                continue

            qty_num = safe_float(qty)

            if qty_num == 0:
                continue

            records.append({
                "시트명": ws.title,
                "모델": model,
                "품번": normalize_part_no(part_no),
                "고객사품번": customer_part_no,
                "날짜": dc["date"],
                "구분": dc["kind"],
                "계획수량": qty_num,
                "공정인쇄": process_print,
            })

    return pd.DataFrame(records)


def extract_plan_file(plan_file: str) -> pd.DataFrame:
    wb = load_workbook(plan_file, data_only=True)

    frames = []

    for sheet_name in SHEET_NAME:
        if sheet_name not in wb.sheetnames:
            print(f"[건너뜀] 시트 없음: {sheet_name}")
            continue

        ws = wb[sheet_name]
        df = extract_finish_plan_sheet(ws)

        print(f"[계획 추출 완료] {sheet_name}: {len(df)}건")
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


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
        print("추출된 계획 수량이 없습니다.")
        return

    plan_df.to_excel(OUTPUT_FILE, index=False)
    print(f"[저장 완료] {OUTPUT_FILE}")


if __name__ == "__main__":
    main()