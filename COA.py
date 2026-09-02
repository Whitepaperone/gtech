# -*- coding: utf-8 -*-
from tkinter import messagebox
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from openpyxl.utils import column_index_from_string

from ApiClient import fetch_orders
from AppConstants import (
    COA_EXCLUDE_CUSTOMERS,
    DATA_START_ROW,
    FINISH_PLAN_SHEET_NAMES,
    PLAN_SHEET_CONFIG,
)
from CommonUtils import (
    create_progress_window,
    normalize_part_no,
    normalize_process_token,
    select_excel_save_file,
)
from ExtractPlan import (
    extract_finish_plan_sheet as extract_finish_plan_sheet_from_module,
    extract_today_shortage_sheet,
    load_calamine_worksheets,
)


# =========================
# 설정
# =========================
SHEET_NAME = FINISH_PLAN_SHEET_NAMES
EXCLUDE_CUSTOMERS = COA_EXCLUDE_CUSTOMERS
DEFAULT_MAPPING_FILE = Path(
    r"\\Data-server\01. 공유문서\08.생산계획\생산계획\MES 생산계획 업로딩용"
    r"\프로그램\4. 품번매핑.xlsx"
)


def is_due_date_ignored_order(row) -> bool:
    """지정 고객 조건에 따라 납기 제한 없이 계획을 비교할지 판별한다."""
    customer = normalize_process_token(row.get("출고(고객)사", ""))
    remark = normalize_process_token(row.get("비고", ""))
    if customer.upper() == "MESA":
        return True
    return customer == "뉴라텍" and remark.startswith(("현대에버다임", "하나티이씨"))


def extract_plan_file(plan_file: str, progress=None, extract_stock: bool = False) -> pd.DataFrame:
    if progress:
        progress(0, "생산계획 파일 열기 중...")
    
    worksheets = load_calamine_worksheets(plan_file, SHEET_NAME)
    if progress:
            progress(35, "생산계획 파일 열기 완료")
    frames = []
    stock_records = []
    total = len(SHEET_NAME)

    for i, sheet_name in enumerate(SHEET_NAME, start=1):
        if sheet_name not in worksheets:
            if progress:
                progress(35 + int(i / total * 50), f"{sheet_name} 시트 없음, 건너뜀")
            continue
        if progress:
            progress(35 + int((i - 1) / total * 50), f"{sheet_name} 추출 중...")

        ws = worksheets[sheet_name]
        sheet_config = PLAN_SHEET_CONFIG[sheet_name]
        part_col = column_index_from_string(str(sheet_config["part_col"]).upper())
        stock_col = column_index_from_string(str(sheet_config["stock_col"]).upper())
        config = {"extract_stock": True, "stock_col": stock_col} if extract_stock else None
        df = extract_finish_plan_sheet_from_module(ws, config=config)
        today_shortage_df = extract_today_shortage_sheet(ws, config=config)
        if not today_shortage_df.empty:
            # 음수 미달은 선생산되어 AD열 재고에 이미 반영된 수량이므로
            # COA의 가용 생산계획에는 다시 넣지 않는다.
            today_shortage_df = today_shortage_df[
                pd.to_numeric(today_shortage_df["계획수량"], errors="coerce").fillna(0) > 0
            ].copy()
            df = pd.concat([df, today_shortage_df], ignore_index=True, sort=False)

        # 일반 계획은 오늘부터 마지막 계획일까지 모두 사용한다. 오늘 미달분은
        # 별도 이월 계획이므로 위에서 선별한 양수 행만 그대로 유지한다.
        if not df.empty:
            plan_dates = pd.to_datetime(df["날짜"], errors="coerce").dt.normalize()
            kinds = df["구분"].apply(normalize_process_token)
            today = pd.Timestamp.today().normalize()
            df = df[(kinds != "계획") | (plan_dates >= today)].copy()

        # 계획이 0인 품번도 재고로 수주를 충당할 수 있으므로 F/AD열은 별도로 읽는다.
        if extract_stock:
            for stock_row in ws.iter_rows(
                min_row=DATA_START_ROW,
                min_col=min(part_col, stock_col),
                max_col=max(part_col, stock_col),
                values_only=True,
            ):
                first_col = min(part_col, stock_col)
                part_value = stock_row[part_col - first_col]
                stock_value = stock_row[stock_col - first_col]
                part_no = normalize_part_no(part_value)
                stock_number = pd.to_numeric(stock_value, errors="coerce")
                if part_no and pd.notna(stock_number):
                    stock_records.append({"품번": part_no, "재고": float(stock_number)})

        frames.append(df)

        if progress:
            progress(35 + int(i / total * 50), f"{sheet_name} 추출 완료: {len(df)}건")

    if progress:
        progress(90, "생산계획 데이터 정리 중...")
    if not frames:
        return pd.DataFrame()

    result = pd.concat(frames, ignore_index=True)
    if extract_stock and stock_records:
        stock_df = pd.DataFrame(stock_records).groupby("품번", as_index=False)["재고"].max()
        result = result.drop(columns=["재고"], errors="ignore").merge(stock_df, on="품번", how="left")

        # 생산계획이 전혀 없는 재고 보유 품번도 비교 대상 재고표에 남긴다.
        missing_stock = stock_df[~stock_df["품번"].isin(result["품번"])].copy()
        if not missing_stock.empty:
            missing_stock["날짜"] = pd.NaT
            missing_stock["계획수량"] = 0.0
            result = pd.concat([result, missing_stock], ignore_index=True, sort=False)
    return result


# =========================
# 수주 파일 추출
# =========================
def extract_order_file(order_file: str) -> pd.DataFrame:
    required_cols = ["수주번호", "날짜", "납기", "출고(고객)사", "품번", "품명", "규격", "수량", "할당수량", "잔여량", "비고"]

    # COA 내보내기는 보통 A열이 비고 B열부터 헤더가 시작한다. 파일에 제목행이
    # 추가되는 경우도 있으므로 첫 20행에서 필수 헤더가 가장 많이 있는 행을 찾는다.
    preview = pd.read_excel(order_file, header=None, nrows=20, engine="calamine")
    best_header_row = 0
    best_match_count = -1
    for row_idx, row in preview.iterrows():
        values = {normalize_process_token(value) for value in row.tolist()}
        match_count = sum(col in values for col in required_cols)
        if match_count > best_match_count:
            best_header_row = row_idx
            best_match_count = match_count

    df = pd.read_excel(order_file, header=best_header_row, engine="calamine")
    df.columns = [normalize_process_token(col) for col in df.columns]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise RuntimeError(f"수주 파일에서 컬럼을 찾지 못했습니다: {missing}")

    df = df[required_cols].copy()
    df = df[
        ~df["출고(고객)사"].apply(lambda x: normalize_process_token(x) in EXCLUDE_CUSTOMERS)
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
    ignore_due_date = df.apply(is_due_date_ignored_order, axis=1)
    df = df[df["납기"].notna() | ignore_due_date].copy()

    return df


def fetch_orders_from_api(base_url: str = None, timeout: float = None) -> pd.DataFrame:
    """수주 API 응답을 기존 COA 엑셀과 동일한 컬럼 구조로 변환한다."""
    records = fetch_orders(base_url=base_url, timeout=timeout)

    api_columns = {
        "SO_CODE": "수주번호",
        "ORDER_DATE": "날짜",
        "DUE_DATE": "납기",
        "CUSTOMER_NAME": "출고(고객)사",
        "ITEM_CODE": "품번",
        "ITEM_NAME": "품명",
        "ITEM_DESCRIPTION": "규격",
        "ORDER_QTY": "수량",
        "PROCESSED_QTY": "할당수량",
        "REMAIN_QTY": "잔여량",
        "REMARK": "비고",
    }
    df = pd.DataFrame(records).reindex(columns=api_columns).rename(columns=api_columns)
    for col in ["수주번호", "출고(고객)사", "품명", "규격", "비고"]:
        df[col] = df[col].fillna("")

    df = df[
        ~df["출고(고객)사"].apply(lambda x: normalize_process_token(x) in EXCLUDE_CUSTOMERS)
    ].copy()
    df["품번"] = df["품번"].apply(normalize_part_no)
    for col in ["날짜", "납기"]:
        # API의 UTC 날짜(예: 전날 15:00Z)를 한국 현지 날짜로 복원한다.
        df[col] = (
            pd.to_datetime(df[col], errors="coerce", utc=True)
            .dt.tz_convert("Asia/Seoul")
            .dt.tz_localize(None)
            .dt.normalize()
        )
    for col in ["수량", "할당수량", "잔여량"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df = df[(df["잔여량"] > 0) & (df["품번"] != "")].copy()
    ignore_due_date = df.apply(is_due_date_ignored_order, axis=1)
    return df[df["납기"].notna() | ignore_due_date].copy()


# =========================
# 품번 매핑
# =========================
def apply_part_mapping(order_df: pd.DataFrame, mapping_file: str = None) -> pd.DataFrame:
    order_df = order_df.copy()

    # 수주 원본 품번 보존
    order_df["수주품번"] = order_df["품번"]
    # 수주 품번 끝의 -AS는 매핑/계획 품번 비교에서만 제외한다.
    mapping_keys = order_df["수주품번"].str.replace(r"-AS$", "", regex=True, case=False)
    order_df["비교품번"] = mapping_keys
    order_df["품번매핑여부"] = "N"

    if not mapping_file:
        return order_df

    map_df = pd.read_excel(mapping_file, header=0, engine="calamine")
    map_df.columns = [normalize_process_token(c) for c in map_df.columns]

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

    mapped_parts = mapping_keys.map(mapping_dict)
    order_df["비교품번"] = mapped_parts.fillna(mapping_keys)

    order_df["품번매핑여부"] = mapped_parts.notna().map({True: "Y", False: "N"})

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

    # AD열 재고는 같은 품번의 날짜별 계획 행마다 반복되므로 합산하지 않는다.
    # 품번별 최댓값을 현재고로 보고, 납기가 빠른 수주부터 한 번씩 차감한다.
    if "재고" in plan_df.columns:
        stock_source = plan_df[["품번", "재고"]].copy()
        stock_source["재고"] = pd.to_numeric(stock_source["재고"], errors="coerce")
        stock_source = stock_source[stock_source["재고"].notna()]
        stock_by_part = stock_source.groupby("품번")["재고"].max().clip(lower=0).to_dict()
    else:
        stock_by_part = {}
    stock_remaining = dict(stock_by_part)

    plan_df = plan_df[
        (plan_df["품번"] != "") &
        (plan_df["날짜"].notna()) &
        (plan_df["계획수량"] > 0)
    ].copy()

    plan_start_date = plan_df["날짜"].min()
    order_df["납기무관비교"] = order_df.apply(is_due_date_ignored_order, axis=1)

    # 1) 과거 납기 잔량
    past_order_df = order_df[
        (~order_df["납기무관비교"]) & (order_df["납기"] < plan_start_date)
    ].copy()

    past_order_df = past_order_df.sort_values(["비교품번", "납기", "수주번호"], kind="stable")
    past_rows = []
    for _, order in past_order_df.iterrows():
        part_no = order["비교품번"]
        need_qty = float(order["잔여량"])
        available_stock = float(stock_remaining.get(part_no, 0))
        used_stock = min(available_stock, need_qty)
        stock_remaining[part_no] = available_stock - used_stock

        row = order.to_dict()
        row.update({
            "재고": float(stock_by_part.get(part_no, 0)),
            "반영재고수량": used_stock,
            "재고잔량": stock_remaining[part_no],
            "부족수량": max(need_qty - used_stock, 0),
            "판정": "재고 반영 완료" if used_stock >= need_qty else "계획시작일 이전 납기 / 재고 부족 또는 미출고 확인 필요",
        })
        past_rows.append(row)
    past_order_df = pd.DataFrame(past_rows, columns=[
        *past_order_df.columns, "재고", "반영재고수량", "재고잔량", "부족수량", "판정"
    ])
    past_order_df["오늘날짜"] = today
    past_order_df["계획시작일"] = plan_start_date

    # 2) 장기 납기 잔량
    future_order_df = order_df[
        (~order_df["납기무관비교"]) & (order_df["납기"] > horizon_end)
    ].copy()

    future_order_df["판정"] = f"{horizon_weeks}주 이후 납기 / 확인 제외"
    future_order_df["오늘날짜"] = today
    future_order_df["확인종료일"] = horizon_end

    # 3) 실제 비교 대상: 오늘 ~ 오늘+6주
    target_order_df = order_df[
        order_df["납기무관비교"] |
        (
            (order_df["납기"] >= plan_start_date) &
            (order_df["납기"] <= horizon_end)
        )
    ].copy()

    plan_group = (
        plan_df
        .groupby(["품번", "날짜"], as_index=False)["계획수량"]
        .sum()
        .sort_values(["품번", "날짜"])
    )

    plan_group["계획잔량"] = plan_group["계획수량"]

    results = []

    target_order_df = target_order_df.sort_values(["비교품번", "납기", "수주번호"], kind="stable")

    for _, order in target_order_df.iterrows():
        part_no = order["비교품번"]
        due_date = order["납기"]
        ignore_due_date = bool(order["납기무관비교"])
        need_qty = float(order["잔여량"])
        available_stock = float(stock_remaining.get(part_no, 0))
        used_stock = min(available_stock, need_qty)
        stock_remaining[part_no] = available_stock - used_stock
        remain_need = need_qty - used_stock

        matched_plan_qty = 0.0
        matched_plan_dates = []

        plan_date_mask = pd.Series(True, index=plan_group.index) if ignore_due_date else (
            plan_group["날짜"] <= due_date
        )
        candidate_idx = plan_group[
            (plan_group["품번"] == part_no) &
            plan_date_mask &
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

        if shortage_qty == 0 and used_stock >= need_qty:
            judgment = "재고 반영 완료"
        elif shortage_qty == 0 and used_stock > 0:
            judgment = "재고+계획 반영 완료"
        elif matched_plan_qty == 0 and used_stock == 0:
            judgment = "계획 없음"
        elif shortage_qty > 0:
            judgment = "재고+계획 부족" if used_stock > 0 else "계획 부족"
        else:
            judgment = "반영 완료"

        row = order.to_dict()
        row.update({
            "확인종료일": horizon_end,
            "재고": float(stock_by_part.get(part_no, 0)),
            "반영재고수량": used_stock,
            "재고잔량": stock_remaining[part_no],
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
        messagebox.showinfo("알림", "생산계획 파일을 선택하지 않았습니다.")
        return

    if DEFAULT_MAPPING_FILE.is_file():
        mapping_file = str(DEFAULT_MAPPING_FILE)
    else:
        messagebox.showinfo(
            "알림",
            "공용서버 품번매핑 파일에 접근할 수 없습니다.\n"
            "대신 사용할 품번매핑 파일을 선택해 주세요.",
        )
        mapping_file = filedialog.askopenfilename(
            title="품번매핑 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if not mapping_file:
            messagebox.showinfo("알림", "품번매핑 파일을 선택하지 않아 매핑 없이 진행합니다.")
            mapping_file = None

    save_file = select_excel_save_file("결과 파일 저장", plan_file, "COA수주_비교", root)
    if not save_file:
        messagebox.showinfo("알림", "저장할 파일을 선택하지 않았습니다. 결과 저장 없이 종료합니다.")
        return

   

    progress_win, progress = create_progress_window(root, "COA 수주잔량 비교 진행")

    try:
        progress(10, "서버에서 수주 데이터 받는 중...")
        order_df = fetch_orders_from_api()

        progress(20, "품번 매핑 적용 중...")
        order_df = apply_part_mapping(order_df, mapping_file)
        mapped_count = order_df["품번매핑여부"].eq("Y").sum()
        progress(30, f"품번 매핑 적용 완료: {mapped_count}건 변경")

        plan_df = extract_plan_file(plan_file, progress=progress, extract_stock=True)

        progress(92, "수주잔량과 생산계획 비교 중...")
        compare_df, past_order_df, future_order_df, remain_plan_df = compare_order_balance_vs_plan(
            plan_df, order_df, horizon_weeks=6
        )

        progress(96, "엑셀 저장 중...")

        with pd.ExcelWriter(save_file, engine="openpyxl") as writer:
            compare_df.to_excel(writer, sheet_name="6주내_수주잔량_계획확인", index=False)
            past_order_df.to_excel(writer, sheet_name="과거납기_미출고잔량", index=False)
            future_order_df.to_excel(writer, sheet_name="6주이후_확인제외", index=False)
            remain_plan_df.to_excel(writer, sheet_name="계획잔량_참고", index=False)
            plan_df.to_excel(writer, sheet_name="추출된_생산계획", index=False)
            order_df.to_excel(writer, sheet_name="추출된_수주잔량", index=False)

        progress(100, "완료")
        progress_win.destroy()
        messagebox.showinfo("알림", f"결과 저장 완료: {save_file}")

    except Exception as e:
        progress_win.destroy()
        messagebox.showerror("오류", str(e))

if __name__ == "__main__":
    main()
