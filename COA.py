from datetime import datetime, date
from typing import Optional, List, Dict
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook


# =========================
# 설정
# =========================
SHEET_NAME = ["완성공정(실적)"]
OUTPUT_FILE = "./생산계획_COA수주_비교.xlsx"

HEADER_DATE_ROW = 33   # 날짜 행
HEADER_KIND_ROW = 34   # 구분 행: 미달 / 계획 / 실적
DATA_START_ROW = 35

DATE_SCAN_MIN_COL = 1
DATE_HEADER_GAP_BREAK = 3

EXCLUDE_CUSTOMERS = [
    "HD건설기계㈜ 인천"
]


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
# 수주 파일 추출
# =========================
def extract_order_file(order_file: str) -> pd.DataFrame:
    df = pd.read_excel(order_file, header=0)

    required_cols = ["수주번호", "날짜", "납기", "출고(고객)사", "품번", "품명", "규격", "수량", "할당수량", "잔여량", "비고"]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise RuntimeError(f"수주 파일에서 컬럼을 찾지 못했습니다: {missing}")

    df = df[required_cols].copy()
    df = df[
        ~df["출고(고객)사"].apply(lambda x: normalize_text(x) in EXCLUDE_CUSTOMERS)
    ].copy()

    df["품번"] = df["품번"].apply(normalize_part_no)
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce").dt.normalize()
    df["납기"] = pd.to_datetime(df["납기"], errors="coerce").dt.normalize()

    for col in ["수량", "할당수량", "잔여량"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .replace("nan", "0")
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df = df[df["잔여량"] > 0].copy()
    df = df[df["품번"] != ""].copy()
    df = df[df["납기"].notna()].copy()

    return df

# =========================
# 품번 매핑
# =========================
def apply_part_mapping(order_df: pd.DataFrame, mapping_file: str = None) -> pd.DataFrame:
    order_df = order_df.copy()

    # 수주 원본 품번 보존
    order_df["수주품번"] = order_df["품번"]
    order_df["비교품번"] = order_df["품번"]
    order_df["품번매핑여부"] = "N"

    if not mapping_file:
        return order_df

    map_df = pd.read_excel(mapping_file, header=0)
    map_df.columns = [normalize_text(c).replace(" ", "") for c in map_df.columns]

    # 매핑표 헤더 표준화
    rename_map = {
        "수주품번": "수주품번",
        "변경품번": "비교품번",
        "계획품번": "비교품번",
    }

    map_df = map_df.rename(columns=rename_map)

    if "수주품번" not in map_df.columns or "비교품번" not in map_df.columns:
        raise RuntimeError(
            f"품번매핑 파일에 '수주 품번'과 '변경 품번' 컬럼이 필요합니다. 현재 컬럼: {list(map_df.columns)}"
        )

    map_df = map_df[["수주품번", "비교품번"]].copy()

    map_df["수주품번"] = map_df["수주품번"].apply(normalize_part_no)
    map_df["비교품번"] = map_df["비교품번"].apply(normalize_part_no)

    map_df = map_df[
        (map_df["수주품번"] != "") &
        (map_df["비교품번"] != "")
    ].copy()

    map_df = map_df.drop_duplicates(subset=["수주품번"], keep="last")

    mapping_dict = dict(zip(map_df["수주품번"], map_df["비교품번"]))

    order_df["비교품번"] = order_df["수주품번"].map(mapping_dict).fillna(order_df["수주품번"])

    order_df["품번매핑여부"] = order_df.apply(
        lambda r: "Y" if r["비교품번"] != r["수주품번"] else "N",
        axis=1
    )

    print(f"[품번매핑 적용] {order_df['품번매핑여부'].eq('Y').sum()}건 변경")

    return order_df

# =========================
# 수주잔량 vs 생산계획 비교
# =========================
def compare_order_balance_vs_plan(plan_df: pd.DataFrame, order_df: pd.DataFrame, horizon_weeks: int = 6):
    if plan_df.empty:
        raise RuntimeError("생산계획 데이터가 없습니다.")

    if order_df.empty:
        raise RuntimeError("잔여량이 있는 수주 데이터가 없습니다.")

    plan_df = plan_df.copy()
    order_df = order_df.copy()

    today = pd.Timestamp.today().normalize()
    horizon_end = today + pd.Timedelta(weeks=horizon_weeks)

    plan_df["품번"] = plan_df["품번"].apply(normalize_part_no)
    plan_df["날짜"] = pd.to_datetime(plan_df["날짜"], errors="coerce").dt.normalize()
    plan_df["계획수량"] = pd.to_numeric(plan_df["계획수량"], errors="coerce").fillna(0)

    plan_df = plan_df[
        (plan_df["품번"] != "") &
        (plan_df["날짜"].notna()) &
        (plan_df["계획수량"] > 0)
    ].copy()

    plan_start_date = plan_df["날짜"].min()

    # 1) 과거 납기 잔량
    past_order_df = order_df[
        order_df["납기"] < plan_start_date
    ].copy()

    past_order_df["판정"] = "계획시작일 이전 납기 / 재고 또는 미출고 확인 필요"
    past_order_df["오늘날짜"] = today
    past_order_df["계획시작일"] = plan_start_date

    # 2) 장기 납기 잔량
    future_order_df = order_df[
        order_df["납기"] > horizon_end
    ].copy()

    future_order_df["판정"] = f"{horizon_weeks}주 이후 납기 / 확인 제외"
    future_order_df["오늘날짜"] = today
    future_order_df["확인종료일"] = horizon_end

    # 3) 실제 비교 대상: 오늘 ~ 오늘+6주
    target_order_df = order_df[
        (order_df["납기"] >= plan_start_date) &
        (order_df["납기"] <= horizon_end)
    ].copy()

    plan_group = (
        plan_df
        .groupby(["품번", "날짜"], as_index=False)["계획수량"]
        .sum()
        .sort_values(["품번", "날짜"])
    )

    plan_group["계획잔량"] = plan_group["계획수량"]

    results = []

    target_order_df = target_order_df.sort_values(["품번", "납기", "수주번호"])

    for _, order in target_order_df.iterrows():
        part_no = order["비교품번"]
        due_date = order["납기"]
        need_qty = float(order["잔여량"])
        remain_need = need_qty

        matched_plan_qty = 0.0
        matched_plan_dates = []

        candidate_idx = plan_group[
            (plan_group["품번"] == part_no) &
            (plan_group["날짜"] <= due_date) &
            (plan_group["계획잔량"] > 0)
        ].index

        for idx in candidate_idx:
            available = float(plan_group.at[idx, "계획잔량"])
            use_qty = min(available, remain_need)

            plan_group.at[idx, "계획잔량"] = available - use_qty
            remain_need -= use_qty
            matched_plan_qty += use_qty

            matched_plan_dates.append(
                f"{plan_group.at[idx, '날짜'].strftime('%Y-%m-%d')}({use_qty:g})"
            )

            if remain_need <= 0:
                break

        shortage_qty = max(remain_need, 0)

        if matched_plan_qty == 0:
            judgment = "계획 없음"
        elif shortage_qty > 0:
            judgment = "계획 부족"
        else:
            judgment = "반영 완료"

        row = order.to_dict()
        row.update({
            "확인종료일": horizon_end,
            "납기전_반영계획수량": matched_plan_qty,
            "부족수량": shortage_qty,
            "반영계획일자": ", ".join(matched_plan_dates),
            "판정": judgment,
        })

        results.append(row)

    compare_df = pd.DataFrame(results)
    remain_plan_df = plan_group[plan_group["계획잔량"] > 0].copy()

    return compare_df, past_order_df, future_order_df, remain_plan_df


# =========================
# 실행부
# =========================
def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    plan_file = filedialog.askopenfilename(
        title="생산계획 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )

    if not plan_file:
        print("생산계획 파일을 선택하지 않았습니다.")
        return

    order_file = filedialog.askopenfilename(
        title="수주 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )

    if not order_file:
        print("수주 파일을 선택하지 않았습니다.")
        return

    mapping_file = filedialog.askopenfilename(title="품번매핑 파일 선택")
    if not mapping_file:
        print("품번매핑 파일을 선택하지 않았습니다. 매핑 없이 진행합니다.")
        mapping_file = None

    order_df = extract_order_file(order_file)
    if mapping_file:
        order_df = apply_part_mapping(order_df, mapping_file)

    plan_df = extract_plan_file(plan_file)

    compare_df, past_order_df, future_order_df, remain_plan_df = compare_order_balance_vs_plan(
    plan_df,
    order_df,
    horizon_weeks=6
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        compare_df.to_excel(writer, sheet_name="6주내_수주잔량_계획확인", index=False)
        past_order_df.to_excel(writer, sheet_name="과거납기_미출고잔량", index=False)
        future_order_df.to_excel(writer, sheet_name="6주이후_확인제외", index=False)
        remain_plan_df.to_excel(writer, sheet_name="계획잔량_참고", index=False)
        plan_df.to_excel(writer, sheet_name="추출된_생산계획", index=False)
        order_df.to_excel(writer, sheet_name="추출된_수주잔량", index=False)

    print(f"[완료] 결과 저장: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()