from tkinter import Tk, filedialog, messagebox

from CommonUtils import (
    create_progress_window,
    select_excel_save_file,
)
from ExtractPlan import (
    build_work_order_upload_df,
    extract_work_order_file,
    filter_work_order_period,
    select_date_range,
)


def main():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.lift()
    root.focus_force()

    file_path = filedialog.askopenfilename(
        title="생산계획 파일 선택",
        filetypes=[("Excel files", "*.xlsx")],
    )
    if not file_path:
        return

    save_file = select_excel_save_file("작업지시 파일 저장", file_path, "MES작업지시", root)
    if not save_file:
        messagebox.showinfo("알림", "저장할 파일을 선택하지 않았습니다. 결과 저장 없이 종료합니다.")
        return

    progress_win = None

    try:
        start_date, end_date = select_date_range(root)

        progress_win, progress = create_progress_window(root, "작업지시 생성 중")
        progress(5, "생산계획 파일을 읽는 중...")

        plan_df = extract_work_order_file(file_path, progress=progress)

        progress(78, "선택한 기간의 계획을 정리하는 중...")
        plan_df = filter_work_order_period(plan_df, start_date=start_date, end_date=end_date)

        if plan_df.empty:
            raise Exception("조건에 맞는 계획 데이터가 없습니다.")

        progress(88, "작업지시 업로드 양식을 만드는 중...")
        final_df = build_work_order_upload_df(plan_df)
        if final_df.empty:
            raise Exception("작업지시로 생성할 계획 수량이 없습니다.")

        progress(96, "엑셀 파일로 저장하는 중...")
        final_df.to_excel(save_file, index=False)
        progress(100, "완료")

        progress_win.destroy()
        progress_win = None

        messagebox.showinfo("완료", f"저장 완료:\n{save_file}")

    except Exception as e:
        if progress_win:
            progress_win.destroy()
        messagebox.showerror("오류", str(e))


if __name__ == "__main__":
    main()
