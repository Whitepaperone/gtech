from typing import Optional
import tkinter as tk
from tkinter import messagebox

import pandas as pd

from AppConstants import MES_OUTPUT_FILE, MES_SHEET_CONFIG, WORKSHOP_TO_PROCESS
from CommonUtils import (
    create_progress_window, filter_by_period, normalize_part_no,
    normalize_process_token, select_date_range, select_excel_file,
)
from ExtractPlan import extract_plan_sheet as extract_plan_sheet_from_module
from ExtractPlan import load_calamine_worksheets


OUTPUT_FILE = MES_OUTPUT_FILE
SHEET_CONFIG = MES_SHEET_CONFIG
EXCLUDED_MES_WORKSHOPS = {"CORE부품"}


def canonical_workshop_name(name: str) -> str:
    return (normalize_process_token(name)
            .replace("악세서리", "액세서리")
            .replace("액세사리", "액세서리"))


def map_process_from_workshop(name: str) -> Optional[str]:
    return WORKSHOP_TO_PROCESS.get(normalize_process_token(name))


def extract_plan_all(plan_file: str, progress=None) -> pd.DataFrame:
    worksheets = load_calamine_worksheets(plan_file, SHEET_CONFIG.keys())

    frames = []
    total = max(len(SHEET_CONFIG), 1)
    for i, (sheet_name, config) in enumerate(SHEET_CONFIG.items(), start=1):
        if sheet_name not in worksheets:
            if progress:
                progress(20 + int(i / total * 25), f"{sheet_name} 시트 없음, 건너뜀")
            continue
        if progress:
            progress(20 + int((i - 1) / total * 25), f"{sheet_name} 데이터 추출 중...")
        frame = extract_plan_sheet_from_module(
            worksheets[sheet_name], config, check_red_font=False,
            progress=lambda pct, text, i=i: progress(
                20 + int(((i - 1) + pct / 100) / total * 25), text
            ) if progress else None,
        )
        if not frame.empty:
            frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def extract_mes(mes_file: str) -> pd.DataFrame:
    """MES 작업지시 원본을 흐름 계산 없이 정규화한다."""
    df = pd.read_excel(mes_file, sheet_name="Sheet")
    needed = ["계획일", "작업장명", "작업반명", "작업지시상태", "품번", "지시량", "실적량"]
    missing = [column for column in needed if column not in df.columns]
    if missing:
        raise RuntimeError(f"MES 파일에 필요한 컬럼이 없습니다: {missing}")

    df = df[needed].copy()
    df["작업장명"] = df["작업장명"].fillna("").astype(str).str.strip()
    df["작업반명"] = df["작업반명"].fillna("").astype(str).str.strip()
    df = df[~df["작업장명"].isin(EXCLUDED_MES_WORKSHOPS)].copy()
    df["공정"] = df["작업장명"].map(canonical_workshop_name).map(map_process_from_workshop)
    df = df[df["공정"].notna() & (df["작업지시상태"] != "종료")].copy()
    df["날짜"] = pd.to_datetime(df["계획일"], errors="coerce").dt.normalize()
    df["지시량"] = pd.to_numeric(df["지시량"], errors="coerce").fillna(0)
    df["실적량"] = pd.to_numeric(df["실적량"], errors="coerce").fillna(0)
    df["품번"] = df["품번"].apply(normalize_part_no)
    return (df.groupby(["공정", "작업장명", "작업반명", "품번", "날짜"], as_index=False)
              .agg({"지시량": "sum", "실적량": "sum"})
              .sort_values(["공정", "작업장명", "작업반명", "품번", "날짜"])
              .reset_index(drop=True))


def compare_shortage_actual_flow(
    plan_df: pd.DataFrame,
    mes_df: pd.DataFrame,
    today=None,
    target_keys=None,
) -> pd.DataFrame:
    """오늘 미달과 MES 작업지시/실적 흐름을 품번별로 비교한다.

    양수 미달은 오늘까지 남아 있는 MES 작업지시 잔량과 비교한다.
    음수 미달은 오늘 이후 계획을 가까운 날짜부터 선실적으로 소진한 결과가
    MES에 동일하게 반영되었는지 검사한다.
    """
    result_columns = [
        "공정", "작업장명", "작업반명", "품번", "기준일", "미달수량",
        "비교수량", "판정", "상세",
    ]
    if plan_df.empty:
        return pd.DataFrame(columns=result_columns)

    base_date = (
        pd.Timestamp(today).normalize()
        if today is not None else pd.Timestamp.today().normalize()
    )

    plan = plan_df.copy()
    plan["날짜"] = pd.to_datetime(plan["날짜"], errors="coerce").dt.normalize()
    plan["수량"] = pd.to_numeric(plan["수량"], errors="coerce").fillna(0)
    plan["품번"] = plan["품번"].apply(normalize_part_no)
    plan["구분"] = plan["구분"].fillna("").astype(str).str.strip()

    mes = mes_df.copy()
    if not mes.empty:
        mes["날짜"] = pd.to_datetime(mes["날짜"], errors="coerce").dt.normalize()
        mes["지시량"] = pd.to_numeric(mes["지시량"], errors="coerce").fillna(0)
        mes["실적량"] = pd.to_numeric(mes["실적량"], errors="coerce").fillna(0)
        mes["품번"] = mes["품번"].apply(normalize_part_no)

    shortage = plan[(plan["날짜"] == base_date) & (plan["구분"] == "미달")].copy()
    if target_keys is not None:
        shortage = shortage[
            shortage["품번"].isin(target_keys)
        ]
    shortage = (
        shortage.groupby(["품번"], as_index=False, dropna=False)
        .agg(
            공정=("공정", "first"),
            작업장명=("작업장명", "first"),
            작업반명=("작업반명", "first"),
            미달수량=("수량", "sum"),
        )
    )

    results = []
    for row in shortage.itertuples(index=False):
        process = row.공정
        part_no = row.품번
        shortage_qty = float(row.미달수량)
        same_plan = plan[plan["품번"] == part_no]
        same_mes = mes[mes["품번"] == part_no]

        if shortage_qty > 0:
            past_mes = same_mes[same_mes["날짜"] <= base_date]
            outstanding = float((past_mes["지시량"] - past_mes["실적량"]).sum())
            matched = abs(outstanding - shortage_qty) < 1e-9
            checks = []
            remaining = shortage_qty
            past_plan = same_plan[
                (same_plan["날짜"] < base_date) & (same_plan["구분"] == "계획")
            ]
            past_plan = (
                past_plan.groupby("날짜", as_index=False)["수량"].sum()
                .sort_values("날짜", ascending=False)
            )
            for plan_row in past_plan.itertuples(index=False):
                day = plan_row.날짜
                plan_qty = float(plan_row.수량)
                day_mes = same_mes[same_mes["날짜"] == day]
                work_qty = float(day_mes["지시량"].sum())
                actual_qty = float(day_mes["실적량"].sum())

                if remaining <= 0:
                    if work_qty != 0 or actual_qty != 0:
                        matched = False
                        checks.append(f"{day:%Y-%m-%d} 소멸 대상 작업지시가 남음")
                    continue

                expected_work = plan_qty
                expected_actual = max(0.0, plan_qty - remaining)
                if (
                    abs(work_qty - expected_work) >= 1e-9
                    or abs(actual_qty - expected_actual) >= 1e-9
                ):
                    matched = False
                    checks.append(
                        f"{day:%Y-%m-%d} 기대 지시/실적 "
                        f"{expected_work:g}/{expected_actual:g}, "
                        f"MES {work_qty:g}/{actual_qty:g}"
                    )
                remaining = max(0.0, remaining - plan_qty)

            if remaining > 0:
                matched = False
                checks.append(f"과거 계획보다 큰 미달 잔량 {remaining:g}")
            detail = (
                "; ".join(checks)
                if checks else f"최신 계획부터 미달 {shortage_qty:g} 정상 배분"
            )
            compare_qty = outstanding
        elif shortage_qty < 0:
            prebuilt = abs(shortage_qty)
            future_plan = same_plan[
                (same_plan["날짜"] >= base_date) & (same_plan["구분"] == "계획")
            ]
            future_plan = (
                future_plan.groupby("날짜", as_index=False)["수량"].sum()
                .sort_values("날짜")
            )

            matched = True
            remaining = prebuilt
            checks = []
            compare_qty = 0.0
            for plan_row in future_plan.itertuples(index=False):
                day = plan_row.날짜
                plan_qty = float(plan_row.수량)
                day_mes = same_mes[same_mes["날짜"] == day]
                work_qty = float(day_mes["지시량"].sum())
                actual_qty = float(day_mes["실적량"].sum())

                if remaining >= plan_qty:
                    if work_qty != 0 or actual_qty != 0:
                        matched = False
                        checks.append(
                            f"{day:%Y-%m-%d} 소진 계획 {plan_qty:g}인데 MES 작업지시가 남음"
                        )
                    remaining -= plan_qty
                    if remaining == 0:
                        break
                    continue

                expected_actual = remaining
                compare_qty = actual_qty
                if work_qty == 0 or abs(actual_qty - expected_actual) >= 1e-9:
                    matched = False
                    checks.append(
                        f"{day:%Y-%m-%d} 기대 실적 {expected_actual:g}, "
                        f"MES 지시/실적 {work_qty:g}/{actual_qty:g}"
                    )
                remaining = 0
                break

            if remaining > 0:
                matched = False
                checks.append(f"오늘 이후 계획 부족: 미반영 선실적 {remaining:g}")
            detail = "; ".join(checks) if checks else f"선실적 {prebuilt:g} 정상 반영"
        else:
            matched = True
            compare_qty = 0.0
            detail = "미달수량 0"

        results.append({
            "공정": process,
            "작업장명": row.작업장명,
            "작업반명": row.작업반명,
            "품번": part_no,
            "기준일": base_date.strftime("%Y-%m-%d"),
            "미달수량": shortage_qty,
            "비교수량": compare_qty,
            "판정": "실적흐름일치" if matched else "실적흐름불일치",
            "상세": detail,
        })

    return pd.DataFrame(results, columns=result_columns)


def compare_work_orders(
    plan_df: pd.DataFrame,
    mes_df: pd.DataFrame,
    today=None,
) -> pd.DataFrame:
    """계획과 MES 작업지시 수량을 비교한다.

    오늘 음수 미달에 의해 완전히 소진되는 오늘 이후 계획은 작업지시가 없는
    것이 정상이다. 실적 흐름과 중복되는 총실적 초과/계획 소진 후 잔존 판정은
    여기서 수행하지 않는다.
    """
    columns = [
        "공정", "작업장명", "작업반명", "품번", "날짜",
        "계획수량", "미달수량", "MES실적수량", "작업지시수량",
        "1차판정", "판정", "상세",
    ]
    if plan_df.empty and mes_df.empty:
        return pd.DataFrame(columns=columns)

    base_date = (
        pd.Timestamp(today).normalize()
        if today is not None else pd.Timestamp.today().normalize()
    )
    key_columns = ["품번", "날짜"]

    plan = plan_df.copy()
    if not plan.empty:
        plan["날짜"] = pd.to_datetime(plan["날짜"], errors="coerce").dt.normalize()
        plan["수량"] = pd.to_numeric(plan["수량"], errors="coerce").fillna(0)
        plan["품번"] = plan["품번"].apply(normalize_part_no)
        plan["구분"] = plan["구분"].fillna("").astype(str).str.strip()

    mes = mes_df.copy()
    if not mes.empty:
        mes["날짜"] = pd.to_datetime(mes["날짜"], errors="coerce").dt.normalize()
        mes["지시량"] = pd.to_numeric(mes["지시량"], errors="coerce").fillna(0)
        mes["실적량"] = pd.to_numeric(mes["실적량"], errors="coerce").fillna(0)
        mes["품번"] = mes["품번"].apply(normalize_part_no)

    plan_rows = plan[plan["구분"] == "계획"].copy()
    plan_grouped = (
        plan_rows.groupby(key_columns, as_index=False, dropna=False)
        .agg(
            공정=("공정", "first"),
            작업장명=("작업장명", "first"),
            작업반명=("작업반명", "first"),
            계획수량=("수량", "sum"),
        )
    )
    mes_grouped = (
        mes.groupby(key_columns, as_index=False, dropna=False)
        .agg(
            MES공정=("공정", "first"),
            MES작업장명=("작업장명", "first"),
            MES작업반명=("작업반명", "first"),
            작업지시수량=("지시량", "sum"),
            MES실적수량=("실적량", "sum"),
        )
        if not mes.empty else pd.DataFrame(
            columns=key_columns + ["MES공정", "MES작업장명", "MES작업반명", "작업지시수량", "MES실적수량"]
        )
    )

    # 오늘 미달은 오늘 계획 반영 전 값이다. 오늘 이후 계획은 전량 소진 여부를 따로 확인한다.
    current_shortages = plan[
        (plan["구분"] == "미달") & (plan["날짜"] == base_date)
    ].groupby(["품번"], as_index=False)["수량"].sum()
    shortage_map = {
        row.품번: float(row.수량)
        for row in current_shortages.itertuples(index=False)
    }

    # 전월부터 어제까지의 MES 작업지시를 오늘 미달과 비교한다.
    # 오늘 미달은 오늘 계획이 반영되기 전의 누적값이므로 오늘 날짜의
    # 작업지시/실적은 여기 합계에 넣지 않는다.
    # MES 실적에는 오늘 뒤늦게 등록한 값이 섞일 수 있으므로 실적이 있으면
    # 수량에서 차감해 자동 판정하지 않고 수동 확인 대상으로 남긴다.
    # 전월 작업지시와 이번 달 신규 작업지시가 함께 남아 있는 경우도
    # 현재 미달 한 번으로 판단할 수 있다.
    outstanding_totals = {}
    if shortage_map and not mes.empty:
        this_month = base_date.replace(day=1)
        previous_month = this_month - pd.offsets.MonthBegin(1)
        through_yesterday = mes[
            (mes["날짜"] >= previous_month) & (mes["날짜"] < base_date)
        ]
        totals = through_yesterday.groupby("품번", as_index=False).agg(
            지시량=("지시량", "sum"), 실적량=("실적량", "sum")
        )
        outstanding_totals = {
            row.품번: (float(row.지시량), float(row.실적량))
            for row in totals.itertuples(index=False)
        }

    # 양수 미달은 오늘 이전 계획을 최신 날짜부터 역순으로 남긴다.
    positive_shortage_absorbed_keys = set()
    positive_shortages = current_shortages[current_shortages["수량"] > 0]
    for shortage in positive_shortages.itertuples(index=False):
        remaining = float(shortage.수량)
        past = plan_grouped[
            (plan_grouped["품번"] == shortage.품번)
            & (plan_grouped["날짜"] < base_date)
        ].sort_values("날짜", ascending=False)
        for past_row in past.itertuples(index=False):
            if remaining > 0:
                remaining = max(0.0, remaining - float(past_row.계획수량))
            else:
                positive_shortage_absorbed_keys.add(
                    (shortage.품번, past_row.날짜)
                )

    # 음수 미달로 전량 소진되어 MES 작업지시가 없어야 하는 오늘 이후 계획 날짜를 계산한다.
    absorbed_keys = set()
    shortages = current_shortages[current_shortages["수량"] < 0]
    for shortage in shortages.itertuples(index=False):
        remaining = abs(float(shortage.수량))
        future = plan_grouped[
            (plan_grouped["품번"] == shortage.품번)
            & (plan_grouped["날짜"] >= base_date)
        ].sort_values("날짜")
        for future_row in future.itertuples(index=False):
            if remaining < float(future_row.계획수량):
                break
            absorbed_keys.add((shortage.품번, future_row.날짜))
            remaining -= float(future_row.계획수량)
            if remaining == 0:
                break

    merged = plan_grouped.merge(mes_grouped, on=key_columns, how="outer")
    for name in ["계획수량", "작업지시수량", "MES실적수량"]:
        merged[name] = pd.to_numeric(merged[name], errors="coerce").fillna(0)
    merged["공정"] = merged["공정"].fillna(merged["MES공정"])
    merged["작업장명"] = merged["작업장명"].fillna(merged["MES작업장명"])
    merged["작업반명"] = merged["작업반명"].fillna(merged["MES작업반명"])
    merged["미달수량"] = [
        shortage_map.get(part_no, 0.0)
        for part_no in merged["품번"]
    ]

    first_judgments = []
    judgments = []
    details = []
    for row in merged.itertuples(index=False):
        plan_qty = float(row.계획수량)
        work_qty = float(row.작업지시수량)
        actual_qty = float(row.MES실적수량)
        key = (row.품번, row.날짜)
        has_current_shortage = row.품번 in shortage_map
        shortage_qty = shortage_map.get(row.품번, 0.0)
        should_be_gone = (
            key in absorbed_keys
            or key in positive_shortage_absorbed_keys
            or (
                has_current_shortage
                and shortage_qty <= 0
                and row.날짜 < base_date
            )
        )
        outstanding = outstanding_totals.get(row.품번)
        is_past_mes_only = (
            row.날짜 < base_date
            and plan_qty == 0
            and work_qty != 0
            and has_current_shortage
        )
        has_mes_actual_to_review = (
            is_past_mes_only
            and outstanding is not None
            and abs(outstanding[1]) >= 1e-9
        )
        today_shortage_matches = (
            is_past_mes_only
            and outstanding is not None
            and not has_mes_actual_to_review
            and abs(outstanding[0] - shortage_qty) < 1e-9
        )

        # 1차는 미달/실적을 보지 않고 계획과 MES 작업지시만 비교한다.
        if has_mes_actual_to_review:
            first_judge = "MES실적있음_수동확인"
        elif today_shortage_matches:
            first_judge = "작업지시일치"
        elif plan_qty == work_qty:
            first_judge = "작업지시일치"
        elif plan_qty == 0:
            first_judge = "해당날짜_계획없는데작업지시있음"
        elif work_qty == 0:
            first_judge = "작업지시없음"
        elif work_qty < plan_qty:
            first_judge = "작업지시부족"
        else:
            first_judge = "작업지시과다"

        # 1차 불일치인 경우에만 미달과 MES 실적을 적용한다.
        if has_mes_actual_to_review:
            judge = "MES실적있음_수동확인"
            detail = "MES 실적 등록시점 수동 확인 필요"
        elif today_shortage_matches:
            judge = "작업지시일치"
            detail = "전월~어제 작업지시 합계가 오늘 미달과 일치"
        elif first_judge == "작업지시일치" and not should_be_gone:
            judge = first_judge
            detail = ""
        elif should_be_gone and work_qty > 0:
            judge = "작업지시잔존"
            detail = (
                f"오늘 미달 {shortage_qty:g} 기준 소멸 대상인데 "
                f"작업지시 {work_qty:g} 잔존"
            )
        elif should_be_gone and work_qty == 0:
            judge = "작업지시일치"
            detail = "미달 기준 소멸 대상 작업지시 정상 소멸"
        elif key in absorbed_keys:
            if work_qty == 0:
                judge = "작업지시일치_선생산소진"
                detail = f"계획 {plan_qty:g} 전량을 오늘 음수 미달로 소진"
            else:
                judge = "작업지시과다"
                detail = f"선생산 소진 계획인데 작업지시 {work_qty:g} 남음"
        elif key in positive_shortage_absorbed_keys and work_qty == 0:
            judge = "작업지시일치"
            detail = f"오늘 양수 미달 {row.미달수량:g}보다 오래된 계획 정상 소멸"
        elif plan_qty > 0 and work_qty == 0:
            if actual_qty == plan_qty:
                judge = "작업지시일치_MES실적대체"
                detail = f"계획 {plan_qty:g}을 MES실적 {actual_qty:g}으로 충족"
            else:
                judge = "작업지시없음"
                detail = ""
        elif first_judge == "작업지시부족":
            judge = first_judge
            detail = ""
        elif first_judge == "작업지시과다":
            judge = first_judge
            detail = ""
        else:
            judge = first_judge
            detail = ""
        first_judgments.append(first_judge)
        judgments.append(judge)
        details.append(detail)

    merged["1차판정"] = first_judgments
    merged["판정"] = judgments
    merged["상세"] = details
    merged["품번최초날짜"] = merged.groupby("품번")["날짜"].transform("min")
    merged = merged.sort_values(
        ["품번최초날짜", "품번", "날짜"], ascending=[True, True, True]
    )
    return merged[columns].reset_index(drop=True)


def save_normalized_results(
    plan_df: pd.DataFrame,
    mes_df: pd.DataFrame,
    shortage_flow_df: pd.DataFrame,
    work_order_df: pd.DataFrame,
    output_file: str,
):
    """새 알고리즘의 입력이 될 정규화 데이터만 저장한다."""
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        plan_df.to_excel(writer, sheet_name="1_계획정규화", index=False)
        mes_df.to_excel(writer, sheet_name="2_MES정규화", index=False)
        shortage_flow_df.to_excel(writer, sheet_name="3_미달실적흐름", index=False)
        work_order_output = work_order_df[[
            "품번", "작업장명", "작업반명", "날짜", "계획수량",
            "미달수량", "MES실적수량", "작업지시수량", "판정", "상세",
        ]].copy()
        work_order_output["날짜"] = pd.to_datetime(
            work_order_output["날짜"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        work_order_output.to_excel(writer, sheet_name="4_작업지시판정", index=False)


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    plan_file = select_excel_file("생산계획 엑셀 선택", root)
    if not plan_file:
        messagebox.showinfo("알림", "생산계획 엑셀을 선택하지 않았습니다.")
        return
    mes_file = select_excel_file("MES 작업지시 엑셀 선택", root)
    if not mes_file:
        messagebox.showinfo("알림", "MES 작업지시 엑셀을 선택하지 않았습니다.")
        return

    start_date, end_date = select_date_range(root)
    progress_win, progress = create_progress_window(root, "MES 데이터 정규화 진행")
    try:
        progress(20, "생산계획 정규화 중...")
        all_plan_df = extract_plan_all(plan_file, progress=progress)
        plan_df = filter_by_period(all_plan_df, start_date, end_date)
        progress(60, "MES 작업지시 정규화 중...")
        all_mes_df = extract_mes(mes_file)
        mes_df = filter_by_period(all_mes_df, start_date, end_date)
        progress(75, "계획과 MES 작업지시 1차 비교 중...")
        work_order_df = compare_work_orders(all_plan_df, all_mes_df)
        mismatch_rows = work_order_df[work_order_df["1차판정"] != "작업지시일치"]
        mismatch_keys = set(mismatch_rows["품번"])
        progress(82, "불일치 품번의 오늘 미달 실적 흐름 비교 중...")
        shortage_flow_df = compare_shortage_actual_flow(
            all_plan_df, all_mes_df, target_keys=mismatch_keys
        )
        work_order_df = filter_by_period(work_order_df, start_date, end_date)
        progress(90, "엑셀 저장 중...")
        save_normalized_results(
            plan_df, mes_df, shortage_flow_df, work_order_df, OUTPUT_FILE
        )
        progress(100, "완료")
        progress_win.destroy()
        messagebox.showinfo("완료", f"정규화 데이터 저장 완료: {OUTPUT_FILE}")
    except Exception as error:
        progress_win.destroy()
        messagebox.showerror("오류", str(error))


if __name__ == "__main__":
    main()
