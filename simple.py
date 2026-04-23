from tkinter import Tk, filedialog
import pandas as pd

OUTPUT_FILE = "./비교결과.xlsx"

def pick_file(title: str) -> str:
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not path:
        raise SystemExit(f"{title} 파일을 선택하지 않았습니다.")
    return path


def main():
    root = Tk()
    root.withdraw()

    # 1) 파일 선택
    plan_file = pick_file("정규화된 계획 엑셀 선택")
    mes_file = pick_file("MES 작업지시 엑셀 선택")

    output_file = OUTPUT_FILE

    # 2) 계획 파일 읽기
    # 정규화 파일 시트명: "정규화데이터" 가정
    plan_df = pd.read_excel(plan_file, sheet_name="정규화데이터")

    required_plan_cols = ["날짜", "품번", "구분", "수량"]
    missing_plan = [c for c in required_plan_cols if c not in plan_df.columns]
    if missing_plan:
        raise RuntimeError(f"계획 파일에 필요한 컬럼이 없습니다: {missing_plan}")

    # 계획만 사용
    plan_df = plan_df[plan_df["구분"] == "계획"].copy()
    plan_df["날짜"] = pd.to_datetime(plan_df["날짜"]).dt.date
    plan_df["품번"] = plan_df["품번"].astype(str).str.strip()
    plan_df["수량"] = pd.to_numeric(plan_df["수량"], errors="coerce").fillna(0)

    plan_grouped = (
        plan_df.groupby(["날짜", "품번"], as_index=False)
               .agg(계획수량=("수량", "sum"))
    )

    # 3) MES 파일 읽기
    # 시트명은 상황에 따라 바꾸세요. 보통 첫 시트면 0도 가능
    mes_df = pd.read_excel(mes_file)

    required_mes_cols = ["계획일", "품번", "지시량"]
    missing_mes = [c for c in required_mes_cols if c not in mes_df.columns]
    if missing_mes:
        raise RuntimeError(f"MES 파일에 필요한 컬럼이 없습니다: {missing_mes}")

    # 필요하면 아래 한 줄 활성화해서 '종료' 제외
    # if "작업지시상태" in mes_df.columns:
    #     mes_df = mes_df[mes_df["작업지시상태"] != "종료"].copy()

    mes_df["날짜"] = pd.to_datetime(mes_df["계획일"]).dt.date
    mes_df["품번"] = mes_df["품번"].astype(str).str.strip()
    mes_df["지시량"] = pd.to_numeric(mes_df["지시량"], errors="coerce").fillna(0)

    mes_grouped = (
        mes_df.groupby(["날짜", "품번"], as_index=False)
              .agg(작업지시수량=("지시량", "sum"))
    )

    # 4) 비교
    result = plan_grouped.merge(
        mes_grouped,
        on=["날짜", "품번"],
        how="outer"
    )

    result["계획수량"] = result["계획수량"].fillna(0)
    result["작업지시수량"] = result["작업지시수량"].fillna(0)
    result["차이"] = result["계획수량"] - result["작업지시수량"]

    def judge(row):
        plan_qty = row["계획수량"]
        mes_qty = row["작업지시수량"]

        if plan_qty > 0 and mes_qty == 0:
            return "작업지시없음"
        elif plan_qty == 0 and mes_qty > 0:
            return "계획없는데 작업지시있음"
        elif plan_qty == mes_qty:
            return "일치"
        elif plan_qty > mes_qty:
            return "수량부족"
        else:
            return "작업지시과다"

    result["판정"] = result.apply(judge, axis=1)

    result = result.sort_values(["날짜", "품번"]).reset_index(drop=True)

    # 5) 저장
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        plan_grouped.to_excel(writer, sheet_name="계획집계", index=False)
        mes_grouped.to_excel(writer, sheet_name="MES집계", index=False)
        result.to_excel(writer, sheet_name="비교결과", index=False)

        summary = (
            result.groupby("판정", as_index=False)
                  .size()
                  .rename(columns={"size": "건수"})
        )
        summary.to_excel(writer, sheet_name="요약", index=False)

    print("저장 완료:", output_file)


if __name__ == "__main__":
    main()