from __future__ import annotations

import argparse
import tkinter as tk
from tkinter import filedialog, messagebox

from CommonUtils import create_progress_window
from ERPPartsDb import DEFAULT_ERP_DB, add_erp_parts_from_excel, refresh_erp_db_from_excel


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="ERP 품목상품 엑셀 파일로 ERP 품목 DB를 갱신합니다.")
    parser.add_argument("erp_file", nargs="?", help="ERP 품목상품 엑셀 파일 경로")
    parser.add_argument("--erp-db", default=str(DEFAULT_ERP_DB), help="저장할 ERP 품목 DB 경로")
    parser.add_argument("--add", action="store_true", help="기존 DB를 지우지 않고 선택한 엑셀의 품번만 추가/덮어쓰기합니다.")
    return parser


def select_erp_file(parent) -> str:
    return filedialog.askopenfilename(
        parent=parent,
        title="ERP 품목상품 엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )


def run_cli(erp_file: str, db_path: str, add_only: bool = False) -> None:
    count = add_erp_parts_from_excel(erp_file, db_path) if add_only else refresh_erp_db_from_excel(erp_file, db_path)
    mode_text = "ADD 완료" if add_only else "전체 갱신 완료"
    print(f"ERP DB {mode_text}: {db_path}")
    print(f"저장 품번 수: {count}")


def run_gui() -> None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        erp_file = select_erp_file(root)
        if not erp_file:
            return

        add_only = messagebox.askyesno(
            "갱신 방식 선택",
            "추가 상품만 ADD로 반영할까요?\n\n"
            "예: 기존 DB 유지, 선택한 엑셀의 품번만 추가/덮어쓰기\n"
            "아니오: 기존 DB 전체 삭제 후 선택한 엑셀로 재생성",
            parent=root,
        )

        progress_win, update_progress = create_progress_window(root, "ERP DB 갱신")
        try:
            if add_only:
                count = add_erp_parts_from_excel(erp_file, DEFAULT_ERP_DB, progress=update_progress)
            else:
                count = refresh_erp_db_from_excel(erp_file, DEFAULT_ERP_DB, progress=update_progress)
            update_progress(100, "완료")
        finally:
            progress_win.destroy()

        mode_text = "ADD 완료" if add_only else "전체 갱신 완료"
        messagebox.showinfo(
            "완료",
            f"ERP DB {mode_text}\n\n"
            f"DB 파일: {DEFAULT_ERP_DB}\n"
            f"저장 품번 수: {count}",
            parent=root,
        )
    except Exception as exc:
        messagebox.showerror("오류", str(exc), parent=root)
        raise
    finally:
        root.destroy()


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.erp_file:
        run_cli(args.erp_file, args.erp_db, add_only=args.add)
        return

    run_gui()


if __name__ == "__main__":
    main()
