from tkinter import Tk, filedialog, messagebox

from CommonUtils import (
    apply_quantities_by_part_left_to_right,
    load_actual_quantities_by_part,
    select_excel_file,
    select_excel_save_file,
)
from ExtractPlan import (
    build_work_order_upload_df,
    extract_work_order_file,
    filter_work_order_period,
    select_date_range,
)


OUTPUT_FILE = "./output.xlsx"


def main():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.lift()
    root.focus_force()

    file_path = filedialog.askopenfilename(
        title="생산계획 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx")],
    )
    if not file_path:
        return

    try:
        start_date, end_date = select_date_range(root)
        plan_df = extract_work_order_file(file_path)
        plan_df = filter_work_order_period(plan_df, start_date=start_date, end_date=end_date)

        if plan_df.empty:
            raise Exception("조건에 맞는 계획 데이터가 없습니다.")

        final_df = build_work_order_upload_df(plan_df)
        if final_df.empty:
            raise Exception("작업지시로 생성할 계획 수량이 없습니다.")

        if messagebox.askyesno(
            "실적 적용 여부",
            "./output.xlsx 같은 오늘 실적 파일을 불러와서\n작업지시 수량에 반영하시겠습니까?",
        ):
            result_file = select_excel_file("오늘 실적 파일 선택", root, initial_path=OUTPUT_FILE)
            if result_file:
                today_quantities = load_actual_quantities_by_part(result_file)
                final_df = apply_quantities_by_part_left_to_right(
                    final_df,
                    today_quantities,
                    part_col=3,
                    value_start_idx=5,
                )

                if final_df.empty:
                    raise Exception("오늘 실적 적용 후 남은 작업지시 수량이 없습니다.")

        save_file = select_excel_save_file("작업지시 파일 저장", file_path, "MES작업지시", root)
        if not save_file:
            messagebox.showinfo("알림", "저장할 파일을 선택하지 않았습니다. 결과 저장 없이 종료합니다.")
            return

        final_df.to_excel(save_file, index=False)
        messagebox.showinfo("완료", f"저장 완료:\n{save_file}")

    except Exception as e:
        messagebox.showerror("오류", str(e))


if __name__ == "__main__":
    main()
