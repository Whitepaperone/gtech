from tkinter import filedialog, messagebox
import tkinter as tk
import re
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd
from openpyxl.styles import PatternFill

from AppConstants import COA_STOCK_CONFIG
from COA import apply_part_mapping, extract_order_file
from CommonUtils import (
    create_progress_window,
    normalize_part_no,
    normalize_process_token,
    select_excel_save_file,
)


REPORT_COLUMNS = [
    "수주번호", "날짜", "납기", "출고기준일", "중요표시", "출고(고객)사",
    "품번", "품명", "규격",
    "수량", "할당수량", "잔여량", "재고", "비고",
]
HIGHLIGHT_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
OVERSEAS_FILL = PatternFill(fill_type="solid", fgColor="F4B183")
URGENT_FILL = PatternFill(fill_type="solid", fgColor="F8696B")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _read_shared_strings(book: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in book.namelist():
        return []

    strings = []
    with book.open(path) as source:
        for event, elem in ET.iterparse(source, events=("end",)):
            if elem.tag == f"{{{MAIN_NS}}}si":
                strings.append("".join(node.text or "" for node in elem.iter(f"{{{MAIN_NS}}}t")))
                elem.clear()
    return strings


def _worksheet_paths(book: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(book.read("xl/workbook.xml"))
    relationships = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
    }

    result = {}
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        rel_id = sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
        target = targets.get(rel_id, "").replace("\\", "/")
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = f"xl/{target}"
        result[sheet.attrib["name"]] = target
    return result


def _cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))

    value_node = cell.find(f"{{{MAIN_NS}}}v")
    if value_node is None or value_node.text is None:
        return None
    if cell_type == "s":
        try:
            return shared_strings[int(value_node.text)]
        except (ValueError, IndexError):
            return None
    return value_node.text


def extract_stock_file_fast(plan_file: str, progress=None) -> pd.DataFrame:
    """openpyxl 없이 XLSX에서 설정된 품번/재고 두 열만 스트리밍한다."""
    path = Path(plan_file)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise RuntimeError("빠른 재고 추출은 .xlsx 또는 .xlsm 생산계획 파일만 지원합니다.")

    sheet_names = COA_STOCK_CONFIG.get("sheet_names", ["완성공정(실적)"])
    part_col = str(COA_STOCK_CONFIG.get("part_col", "F")).upper()
    stock_col = str(COA_STOCK_CONFIG.get("stock_col", "AD")).upper()
    start_row = int(COA_STOCK_CONFIG.get("data_start_row", 35))
    cell_ref_pattern = re.compile(r"([A-Z]+)\d+")
    records = []

    if progress:
        progress(35, "생산계획에서 품번/재고 열만 빠르게 읽는 중...")

    with zipfile.ZipFile(path) as book:
        shared_strings = _read_shared_strings(book)
        worksheet_paths = _worksheet_paths(book)
        found_sheet = False

        for sheet_name in sheet_names:
            sheet_path = worksheet_paths.get(sheet_name)
            if not sheet_path:
                continue
            found_sheet = True
            with book.open(sheet_path) as source:
                for event, row in ET.iterparse(source, events=("end",)):
                    if row.tag != f"{{{MAIN_NS}}}row":
                        continue
                    row_number = int(row.attrib.get("r", "0"))
                    if row_number >= start_row:
                        part_value = None
                        stock_value = None
                        for cell in row.findall(f"{{{MAIN_NS}}}c"):
                            match = cell_ref_pattern.match(cell.attrib.get("r", ""))
                            column = match.group(1) if match else ""
                            if column == part_col:
                                part_value = _cell_value(cell, shared_strings)
                            elif column == stock_col:
                                stock_value = _cell_value(cell, shared_strings)

                        part_no = normalize_part_no(part_value)
                        if part_no:
                            stock_number = pd.to_numeric(
                                str(stock_value or "").replace(",", ""), errors="coerce"
                            )
                            records.append({
                                "품번": part_no,
                                "재고": None if pd.isna(stock_number) else float(stock_number),
                            })
                    row.clear()

        if not found_sheet:
            raise RuntimeError(f"생산계획에서 재고 시트를 찾지 못했습니다: {sheet_names}")

    if progress:
        progress(90, f"재고 추출 완료: {len(records)}건")
    return pd.DataFrame(records, columns=["품번", "재고"])


def parse_lcs_date(remark, fallback_year) -> pd.Timestamp:
    """비고의 'LCS 날짜' 뒤에서 연/월/일 또는 월/일 형식을 읽는다."""
    text = str(remark or "")
    pattern = re.compile(
        r"LCS\s*(?:날짜)?\s*[:：]?\s*"
        r"(?:(\d{2,4})\s*[./\-년]\s*)?"
        r"(\d{1,2})\s*[./\-월]\s*"
        r"(\d{1,2})\s*일?",
        re.IGNORECASE,
    )
    match = pattern.search(text)
    if not match:
        return pd.NaT

    year_text, month_text, day_text = match.groups()
    year = int(year_text) if year_text else int(fallback_year)
    if year < 100:
        year += 2000
    try:
        return pd.Timestamp(year=year, month=int(month_text), day=int(day_text)).normalize()
    except ValueError:
        return pd.NaT


def add_overseas_shipping_priority(report: pd.DataFrame) -> pd.DataFrame:
    """클라크 해외수주에 출고기준일과 중요 표시를 추가한다."""
    report = report.copy()
    report["출고기준일"] = pd.NaT
    report["중요표시"] = ""

    for idx, row in report.iterrows():
        customer = normalize_process_token(row.get("출고(고객)사", ""))
        if not customer.startswith("㈜클라크"):
            continue

        due_date = pd.to_datetime(row.get("납기"), errors="coerce")
        if pd.isna(due_date):
            continue

        lcs_date = parse_lcs_date(row.get("비고", ""), due_date.year)
        if pd.notna(lcs_date):
            report.at[idx, "출고기준일"] = lcs_date - pd.Timedelta(days=1)
            report.at[idx, "중요표시"] = "긴급: LCS 날짜 전일 출고"
        else:
            report.at[idx, "출고기준일"] = due_date.normalize() - pd.Timedelta(days=1)
            report.at[idx, "중요표시"] = "중요: 해외수주 납기 전일 출고"

    return report


def build_stock_by_part(plan_df: pd.DataFrame) -> pd.DataFrame:
    """생산계획에서 비교품번별 현재고를 만든다."""
    if plan_df.empty or "재고" not in plan_df.columns:
        return pd.DataFrame(columns=["비교품번", "재고"])

    stock_df = plan_df[["품번", "재고"]].copy()
    stock_df["비교품번"] = stock_df["품번"].apply(normalize_part_no)
    stock_df["재고"] = pd.to_numeric(stock_df["재고"], errors="coerce")
    stock_df = stock_df[(stock_df["비교품번"] != "") & stock_df["재고"].notna()].copy()
    if stock_df.empty:
        return pd.DataFrame(columns=["비교품번", "재고"])

    # 날짜별 계획 레코드에 같은 현재고가 반복되므로 합계 대신 최댓값을 사용한다.
    return stock_df.groupby("비교품번", as_index=False)["재고"].max()


def build_order_balance_report(order_df: pd.DataFrame, plan_df: pd.DataFrame) -> pd.DataFrame:
    """매핑된 품번으로 수주잔량과 현재고를 연결한다."""
    report = order_df.copy()
    report = report.merge(build_stock_by_part(plan_df), on="비교품번", how="left")
    report["재고"] = pd.to_numeric(report["재고"], errors="coerce").fillna(0)

    # 결과에는 COA 원래 품번을 표시하고 비교에만 매핑 품번을 사용한다.
    report["품번"] = report["수주품번"]
    report = add_overseas_shipping_priority(report)
    report = report[REPORT_COLUMNS].copy()
    return report.sort_values(["납기", "수주번호", "품번"], kind="stable").reset_index(drop=True)


def format_order_balance_sheet(ws, report_df: pd.DataFrame, today=None) -> int:
    """재고 충당 가능 행과 클라크 중요 출고 행을 색상으로 강조한다."""
    today = pd.Timestamp(today).normalize() if today is not None else pd.Timestamp.today().normalize()
    highlighted_rows = set()

    header_columns = {cell.value: cell.column for cell in ws[1]}
    for column_name in ("날짜", "납기", "출고기준일"):
        column = header_columns.get(column_name)
        if column:
            for row_number in range(2, ws.max_row + 1):
                ws.cell(row_number, column).number_format = "yyyy-mm-dd"

    for row_idx, row in report_df.reset_index(drop=True).iterrows():
        due_date = pd.to_datetime(row["납기"], errors="coerce")
        balance = pd.to_numeric(row["잔여량"], errors="coerce")
        stock = pd.to_numeric(row["재고"], errors="coerce")
        if (
            pd.notna(due_date)
            and due_date.normalize() < today
            and pd.notna(balance)
            and pd.notna(stock)
            and balance <= stock
        ):
            for cell in ws[row_idx + 2]:
                cell.fill = HIGHLIGHT_FILL
            highlighted_rows.add(row_idx)

        shipping_date = pd.to_datetime(row.get("출고기준일"), errors="coerce")
        if row.get("중요표시"):
            priority_fill = URGENT_FILL if pd.notna(shipping_date) and shipping_date <= today else OVERSEAS_FILL
            for cell in ws[row_idx + 2]:
                cell.fill = priority_fill
            highlighted_rows.add(row_idx)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return len(highlighted_rows)


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    plan_file = filedialog.askopenfilename(
        title="생산계획 엑셀 선택 (재고 추출용)",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
    )
    if not plan_file:
        messagebox.showinfo("알림", "생산계획 파일을 선택하지 않았습니다.")
        return

    order_file = filedialog.askopenfilename(
        title="수주 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )
    if not order_file:
        messagebox.showinfo("알림", "수주 파일을 선택하지 않았습니다.")
        return

    mapping_file = filedialog.askopenfilename(
        title="품번매핑 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )
    if not mapping_file:
        messagebox.showinfo("알림", "품번매핑 파일을 선택하지 않았습니다. 매핑 없이 진행합니다.")
        mapping_file = None

    save_file = select_excel_save_file("결과 파일 저장", order_file, "수주잔량_재고현황", root)
    if not save_file:
        return

    progress_win, progress = create_progress_window(root, "수주잔량과 재고 비교")
    try:
        progress(10, "수주 파일 읽는 중...")
        order_df = extract_order_file(order_file)

        progress(20, "품번 매핑 적용 중...")
        order_df = apply_part_mapping(order_df, mapping_file)
        mapped_count = int(order_df["품번매핑여부"].eq("Y").sum())

        progress(30, f"생산계획 재고 읽는 중... (매핑 {mapped_count}건)")
        plan_df = extract_stock_file_fast(plan_file, progress=progress)

        progress(92, "수주잔량과 재고 연결 중...")
        report_df = build_order_balance_report(order_df, plan_df)

        progress(96, "결과 저장 중...")
        with pd.ExcelWriter(save_file, engine="openpyxl") as writer:
            report_df.to_excel(writer, sheet_name="수주잔량_재고현황", index=False)
            highlighted_count = format_order_balance_sheet(
                writer.book["수주잔량_재고현황"], report_df
            )

        progress(100, "완료")
        progress_win.destroy()
        messagebox.showinfo(
            "알림",
            f"결과 저장 완료: {save_file}\n"
            f"품번매핑: {mapped_count}건\n색칠된 행: {highlighted_count}건",
        )
    except Exception as exc:
        progress_win.destroy()
        messagebox.showerror("오류", str(exc))


if __name__ == "__main__":
    main()
