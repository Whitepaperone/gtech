# -*- coding: utf-8 -*-
"""출고 변경 API 데이터를 기존 출고·폐기 엑셀에 추가한다."""

from copy import copy
from datetime import date, datetime, timedelta
import getpass
from pathlib import Path
import socket
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill, Side

from ApiClient import acknowledge_shipment_changes, fetch_shipment_changes
from CommonUtils import create_progress_window, normalize_part_no, normalize_process_token


TARGET_SHEET = "출고, 폐기"
REQUIRED_HEADERS = [
    "날짜", "출고번호", "분류", "품번", "품명", "규격", "수량", "반품/선매출고",
    "출고량", "단위", "단가", "단가세액", "공급가", "세액", "소계", "비용", "이익",
    "통화", "담당자", "업체코드", "출고(고객)사", "전화", "주소", "인도장소", "창고",
    "비고", "입력방법", "근거번호",
]
API_FIELD_BY_HEADER = {
    "날짜": "DDAY",
    "출고번호": "SALESCODE",
    "분류": "CATEGORY",
    "품번": "ITEMCODE",
    "품명": "ITEMNAME",
    "규격": "DESCRIPTION",
    "수량": "QTY",
    "출고량": "RSQTY",
    "단위": "UNIT",
    "단가": "UNITPRICE",
    "단가세액": "UNITTAX",
    "공급가": "SALESPRICE",
    "세액": "STAX",
    "소계": "SUBTOTAL",
    "비용": "COST",
    "이익": "PROFIT",
    "통화": "CURRENCY",
    "담당자": "STAFFNAME",
    "업체코드": "CUSTOMERCODE",
    "출고(고객)사": "CUSTOMERNAME",
    "전화": "PHONE",
    "주소": "ADDRESS",
    "인도장소": "OUTLOCATION",
    "창고": "LOCATIONNAME",
    "비고": "REMARK",
    "입력방법": "BASETYPE",
    "근거번호": "SALESCODE",
}
CHANGE_FILLS = {
    "MODIFIED": PatternFill("solid", fgColor="FFF2CC"),
    "PAST_NEW": PatternFill("solid", fgColor="DDEBF7"),
    "DELETED": PatternFill("solid", fgColor="F4CCCC"),
}
PRESALE_COLOR = "FCE4D6"  # Excel: 주황, 강조 6, 80% 더 밝게
PRESALE_FILL = PatternFill("solid", fgColor=PRESALE_COLOR)
EXCLUDED_UNIT_CATEGORIES = {"소모품", "상품", "제관품", "기타", "부재료"}
DEFAULT_MAPPING_FILE = Path(
    r"\\Data-server\01. 공유문서\08.생산계획\생산계획\MES 생산계획 업로딩용"
    r"\프로그램\4. 품번매핑.xlsx"
)


def find_header_row(
    ws, max_scan_rows: int = 100, max_scan_columns: int = 64
) -> tuple[int, dict[str, int]]:
    """필수 헤더가 있는 행과 헤더별 열 번호를 찾는다."""
    best = (0, {})
    for row_number in range(1, min(ws.max_row, max_scan_rows) + 1):
        columns = {}
        # 일부 ERP 엑셀은 빈 서식 때문에 사용 범위가 XFD열 가까이 확장된다.
        # 실제 출고 표 영역만 검색해 수백만 개의 빈 셀 생성을 막는다.
        cells = next(
            ws.iter_rows(
                min_row=row_number,
                max_row=row_number,
                min_col=1,
                max_col=max_scan_columns,
            )
        )
        for cell in cells:
            header = normalize_process_token(cell.value)
            if header in REQUIRED_HEADERS:
                columns[header] = cell.column
        if len(columns) > len(best[1]):
            best = (row_number, columns)

    missing = [header for header in REQUIRED_HEADERS if header not in best[1]]
    if missing:
        raise RuntimeError(f"'{ws.title}' 시트에서 헤더를 찾지 못했습니다: {missing}")
    return best


def find_last_data_row(ws, header_row: int, header_columns: dict[str, int]) -> int:
    """업무 데이터가 입력된 마지막 행을 찾는다."""
    columns = list(header_columns.values())
    for row_number in range(ws.max_row, header_row, -1):
        if any(ws.cell(row_number, column).value not in (None, "") for column in columns):
            return row_number
    return header_row


def load_part_mapping(mapping_file: str | Path | None) -> dict[str, str]:
    if not mapping_file:
        return {}
    workbook = load_workbook(mapping_file, read_only=True, data_only=True)
    try:
        ws = workbook.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows, ())
        normalized = [normalize_process_token(value) for value in header]
        source_names = ("수주품번", "수주 품번")
        target_names = ("변경품번", "변경 품번", "계획품번", "계획 품번")
        source_index = next((normalized.index(name) for name in source_names if name in normalized), None)
        target_index = next((normalized.index(name) for name in target_names if name in normalized), None)
        if source_index is None or target_index is None:
            raise RuntimeError("품번매핑 파일에 수주품번과 변경품번 컬럼이 필요합니다.")

        mapping = {}
        for row in rows:
            source = normalize_part_no(row[source_index] if source_index < len(row) else None)
            target = normalize_part_no(row[target_index] if target_index < len(row) else None)
            if source and target:
                mapping[source] = target
        return mapping
    finally:
        workbook.close()


def _find_formula_source(workbook, target_ws, header_row: int):
    """대상 시트를 우선하고, 없으면 다른 시트에서 AD·AE 수식 행을 찾는다."""
    worksheets = [target_ws, *(ws for ws in workbook.worksheets if ws is not target_ws)]
    for ws in worksheets:
        minimum_row = header_row + 1 if ws is target_ws else 1
        for row_number in range(ws.max_row, minimum_row - 1, -1):
            if any(
                isinstance(ws.cell(row_number, column).value, str)
                and ws.cell(row_number, column).value.startswith("=")
                for column in (30, 31)
            ):
                return ws, row_number
    return target_ws, 0


def _copy_row_style_and_formulas(
    target_ws, source_ws, source_row: int, target_row: int
) -> None:
    """기존 행 서식과 AD·AE 수식을 대상 행 기준으로 복사한다."""
    if source_row <= 0:
        return
    target_ws.row_dimensions[target_row].height = source_ws.row_dimensions[source_row].height
    for column in range(1, 32):  # A:AE
        source = source_ws.cell(source_row, column)
        target = target_ws.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format

    for column in (30, 31):  # AD, AE
        source = source_ws.cell(source_row, column)
        if isinstance(source.value, str) and source.value.startswith("="):
            target = target_ws.cell(target_row, column)
            target.value = Translator(source.value, origin=source.coordinate).translate_formula(
                target.coordinate
            )


def _remove_bottom_border(ws, row_number: int) -> None:
    """연속 추가 행의 중간 행에서는 아래쪽 테두리를 제거한다."""
    for column in range(1, 32):
        cell = ws.cell(row_number, column)
        border = copy(cell.border)
        border.bottom = Side()
        cell.border = border


def _comparable(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if hasattr(value, "isoformat") and not isinstance(value, str):
        try:
            return value.isoformat()
        except Exception:
            pass
    return normalize_process_token(value)




def _is_presale(record: dict) -> bool:
    price = record.get("UNITPRICE")
    try:
        zero_price = price not in (None, "") and float(price) == 0
    except (TypeError, ValueError):
        zero_price = False
    return "선매출" in str(record.get("REMARK") or "") or zero_price


def _write_excel_values(ws, first_row, batch_values, header_columns, progress=None):
    """혼합형 배열 전달 대신 명시적인 셀 주소에 스칼라 값을 기록한다."""
    for offset, values in enumerate(batch_values):
        if progress:
            progress(65 + int(15 * offset / len(batch_values)),
                     f"출고 데이터 입력 중: {offset + 1}/{len(batch_values)}건")
        for column in header_columns.values():
            cell = ws.Cells(first_row + offset, column)
            value = values[column - 1]
            # ERP 문자열을 날짜/수식으로 자동 변환하지 않도록 보존한다.
            if isinstance(value, str):
                cell.NumberFormat = "@"
            cell.Value2 = value
        if progress and (offset % 20 == 0 or offset == len(batch_values) - 1):
            progress(65 + int(15 * (offset + 1) / len(batch_values)),
                     f"출고 데이터 입력 중: {offset + 1}/{len(batch_values)}건")


def _verify_excel_values(ws, first_row, batch_values, header_columns):
    """입력값 불일치는 경고로 반환하고 파일 저장은 허용한다."""
    warnings = []
    for header, column in header_columns.items():
        actual_rows = ws.Range(
            ws.Cells(first_row, column),
            ws.Cells(first_row + len(batch_values) - 1, column),
        ).Value2
        if len(batch_values) == 1:
            actual_rows = ((actual_rows,),)
        for offset, (actual,) in enumerate(actual_rows):
            expected = batch_values[offset][column - 1]
            if actual == expected or (actual in (None, "") and expected in (None, "")):
                continue
            warnings.append(
                f"{first_row + offset}행 '{header}' "
                f"(예상 {expected!r}, 실제 {actual!r})"
            )
    return warnings


def find_existing_change_row(
    ws,
    header_row: int,
    last_data_row: int,
    header_columns: dict[str, int],
    record: dict,
    part_mapping: dict[str, str],
) -> tuple[int | None, bool]:
    """수정 전 값을 포함한 주요 키로 기존 엑셀 행을 찾는다."""
    old_values = record.get("OLD_VALUES")
    old_values = old_values if isinstance(old_values, dict) else {}

    def original(field):
        return old_values.get(field, record.get(field))

    original_part = normalize_part_no(original("ITEMCODE"))
    expected = {
        "출고번호": _comparable(original("SALESCODE")),
        "근거번호": _comparable(original("SALESCODE")),
        "품번": _comparable(part_mapping.get(original_part, original_part)),
        "날짜": _comparable(original("DDAY")),
        "수량": _comparable(original("QTY")),
    }
    match_orders = [
        ("출고번호", "근거번호", "품번"),
        ("출고번호", "품번", "날짜", "수량"),
        ("근거번호", "품번", "날짜", "수량"),
        ("출고번호", "근거번호"),
    ]
    for keys in match_orders:
        if any(not expected[key] for key in keys):
            continue
        matches = []
        for row_number in range(header_row + 1, last_data_row + 1):
            if all(
                _comparable(ws.cell(row_number, header_columns[key]).value) == expected[key]
                for key in keys
            ):
                matches.append(row_number)
        if len(matches) == 1:
            return matches[0], False
        if len(matches) > 1:
            return None, True
    return None, False


def append_shipment_changes(
    input_file: str,
    output_file: str,
    records: list[dict],
    mapping_file: str | Path | None = None,
    target_mode: str = "production_plan",
) -> dict:
    """API 레코드를 대상 시트 마지막 데이터 다음 행에 추가하고 저장한다."""
    keep_vba = Path(input_file).suffix.lower() == ".xlsm"
    workbook = load_workbook(input_file, keep_vba=keep_vba)
    try:
        if target_mode == "shipment_file":
            ws = workbook.worksheets[0]
        else:
            if TARGET_SHEET not in workbook.sheetnames:
                raise RuntimeError(f"엑셀에서 '{TARGET_SHEET}' 시트를 찾지 못했습니다.")
            ws = workbook[TARGET_SHEET]
        header_row, header_columns = find_header_row(ws)
        last_data_row = find_last_data_row(ws, header_row, header_columns)
        formula_source_ws, formula_source_row = _find_formula_source(
            workbook, ws, header_row
        )
        part_mapping = load_part_mapping(mapping_file)
        mapped_count = 0
        presale_count = 0
        marked_modified_count = 0
        marked_deleted_count = 0
        not_found_count = 0
        ambiguous_count = 0
        append_records = []

        for record in records:
            change_type = str(record.get("CHANGE_TYPE") or "NEW").upper()
            if change_type not in {"MODIFIED", "DELETED"}:
                append_records.append(record)
                continue

            existing_row, ambiguous = find_existing_change_row(
                ws, header_row, last_data_row, header_columns, record, part_mapping
            )
            if existing_row is None:
                ambiguous_count += int(ambiguous)
                not_found_count += int(not ambiguous)
                continue
            fill = CHANGE_FILLS[change_type]
            for column in range(1, 32):
                ws.cell(existing_row, column).fill = copy(fill)
            if change_type == "MODIFIED":
                marked_modified_count += 1
            else:
                marked_deleted_count += 1

        for offset, record in enumerate(append_records, start=1):
            target_row = last_data_row + offset
            _copy_row_style_and_formulas(
                ws, formula_source_ws, formula_source_row, target_row
            )
            # 직전 행의 배경색은 새 데이터로 이어받지 않는다.
            for column in range(1, 32):
                ws.cell(target_row, column).fill = PatternFill()
            remark = str(record.get("REMARK") or "")
            presale = _is_presale(record)

            for header, api_field in API_FIELD_BY_HEADER.items():
                value = record.get(api_field)
                if header == "날짜" and value:
                    try:
                        value = datetime.fromisoformat(str(value).replace("Z", "+00:00")).date()
                    except ValueError:
                        pass
                elif header == "품번":
                    original_part = normalize_part_no(value)
                    value = part_mapping.get(original_part, original_part)
                    mapped_count += int(bool(original_part) and value != original_part)
                elif header == "단위" and presale:
                    value = "선매출"
                elif header == "단위" and normalize_process_token(
                    record.get("CATEGORY")
                ) in EXCLUDED_UNIT_CATEGORIES:
                    value = "제외"
                elif header == "출고량" and presale:
                    value = 0
                ws.cell(target_row, header_columns[header], value)

            return_value = record.get("RSQTY") if presale else record.get("RETURNQTY")
            return_cell = ws.cell(target_row, header_columns["반품/선매출고"], return_value)

            fill = CHANGE_FILLS.get(str(record.get("CHANGE_TYPE") or "").upper())
            if fill:
                for column in range(1, 32):
                    ws.cell(target_row, column).fill = copy(fill)
            if presale:
                for column in range(1, 32):
                    ws.cell(target_row, column).fill = copy(PRESALE_FILL)
                presale_count += 1

        # 표의 아래쪽 마감선은 새로 추가된 마지막 행에만 남긴다.
        for row_number in range(last_data_row + 1, last_data_row + len(append_records)):
            _remove_bottom_border(ws, row_number)

        workbook.save(output_file)
        return {
            "added_count": len(append_records),
            "mapped_count": mapped_count,
            "presale_count": presale_count,
            "marked_modified_count": marked_modified_count,
            "marked_deleted_count": marked_deleted_count,
            "not_found_count": not_found_count,
            "ambiguous_count": ambiguous_count,
            "formula_source_found": formula_source_row > 0,
            "target_sheet": ws.title,
            "first_row": last_data_row + 1 if append_records else None,
            "last_row": last_data_row + len(append_records) if append_records else None,
        }
    finally:
        workbook.close()


def _excel_color(hex_color: str) -> int:
    """RGB 16진 색상을 Excel COM 색상 정수로 변환한다."""
    value = hex_color.lstrip("#")
    red, green, blue = (int(value[index:index + 2], 16) for index in (0, 2, 4))
    return red + (green << 8) + (blue << 16)


def _excel_date_value(value):
    """시간대 변환 없이 YYYY-MM-DD를 Excel 날짜 일련번호로 변환한다."""
    if not value:
        return value
    try:
        parsed = date.fromisoformat(str(value)[:10])
    except ValueError:
        return value
    return (parsed - date(1899, 12, 30)).days


def append_shipment_changes_fast(
    input_file: str,
    output_file: str,
    records: list[dict],
    mapping_file: str | Path | None = None,
    target_mode: str = "production_plan",
    progress=None,
) -> dict:
    """설치된 Excel을 백그라운드에서 사용해 출고 변경을 빠르게 반영한다."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("빠른 엑셀 처리를 위해 Microsoft Excel과 pywin32가 필요합니다.") from exc

    excel = None
    workbook = None
    original_calculation = None
    try:
        if progress:
            progress(20, "Excel 실행 준비 중...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        if progress:
            progress(30, "선택한 Excel 파일 여는 중...")
        workbook = excel.Workbooks.Open(str(Path(input_file).resolve()), UpdateLinks=0)
        # 생산계획의 SUMIF/VLOOKUP 등이 셀을 쓸 때마다 재계산되지 않게 한다.
        # 이 작업 전용 Excel 인스턴스이며 저장 전에 기존 모드로 복원한다.
        original_calculation = excel.Calculation
        excel.Calculation = -4135  # xlCalculationManual
        if target_mode == "shipment_file":
            ws = workbook.Worksheets(1)
        else:
            try:
                ws = workbook.Worksheets(TARGET_SHEET)
            except Exception as exc:
                raise RuntimeError(f"엑셀에서 '{TARGET_SHEET}' 시트를 찾지 못했습니다.") from exc

        # 앞 100행/64열만 한 번에 읽어 헤더를 찾는다.
        if progress:
            progress(45, f"'{ws.Name}' 시트 헤더 확인 중...")
        values = ws.Range(ws.Cells(1, 1), ws.Cells(100, 64)).Value
        best_row, header_columns = 0, {}
        for row_number, row in enumerate(values, start=1):
            columns = {}
            for column, value in enumerate(row, start=1):
                header = normalize_process_token(value)
                if header in REQUIRED_HEADERS:
                    columns[header] = column
            if len(columns) > len(header_columns):
                best_row, header_columns = row_number, columns
        missing = [header for header in REQUIRED_HEADERS if header not in header_columns]
        if missing:
            raise RuntimeError(f"'{ws.Name}' 시트에서 헤더를 찾지 못했습니다: {missing}")

        key_columns = list(header_columns.values())
        last_data_row = best_row
        for column in key_columns:
            row = ws.Cells(ws.Rows.Count, column).End(-4162).Row  # xlUp
            last_data_row = max(last_data_row, int(row))

        formula_ws, formula_row = ws, 0
        for candidate_ws in [ws, *[sheet for sheet in workbook.Worksheets if sheet.Name != ws.Name]]:
            for column in (30, 31):
                row = candidate_ws.Cells(candidate_ws.Rows.Count, column).End(-4162).Row
                formula = candidate_ws.Cells(row, column).Formula
                if isinstance(formula, str) and formula.startswith("="):
                    formula_ws, formula_row = candidate_ws, int(row)
                    break
            if formula_row:
                break

        part_mapping = load_part_mapping(mapping_file)
        if progress:
            progress(55, "기존 데이터와 변경 내역 비교 중...")
        existing_values = ws.Range(
            ws.Cells(best_row + 1, 1), ws.Cells(max(last_data_row, best_row + 1), 64)
        ).Value

        def find_com_row(record):
            old = record.get("OLD_VALUES")
            old = old if isinstance(old, dict) else {}
            original = lambda field: old.get(field, record.get(field))
            part = normalize_part_no(original("ITEMCODE"))
            expected = {
                "출고번호": _comparable(original("SALESCODE")),
                "근거번호": _comparable(original("SALESCODE")),
                "품번": _comparable(part_mapping.get(part, part)),
                "날짜": _comparable(original("DDAY")),
                "수량": _comparable(original("QTY")),
            }
            orders = [
                ("출고번호", "근거번호", "품번"),
                ("출고번호", "품번", "날짜", "수량"),
                ("근거번호", "품번", "날짜", "수량"),
                ("출고번호", "근거번호"),
            ]
            for keys in orders:
                if any(not expected[key] for key in keys):
                    continue
                matches = []
                for offset, row in enumerate(existing_values, start=best_row + 1):
                    if all(_comparable(row[header_columns[key] - 1]) == expected[key] for key in keys):
                        matches.append(offset)
                if len(matches) == 1:
                    return matches[0], False
                if len(matches) > 1:
                    return None, True
            return None, False

        append_records = []
        acknowledged_change_ids = []
        marked_modified = marked_deleted = not_found = ambiguous_count = 0
        colors = {"MODIFIED": "FFF2CC", "DELETED": "F4CCCC"}
        for record in records:
            change_type = str(record.get("CHANGE_TYPE") or "NEW").upper()
            if change_type not in colors:
                append_records.append(record)
                if record.get("CHANGE_ID"):
                    acknowledged_change_ids.append(record["CHANGE_ID"])
                continue
            row_number, ambiguous = find_com_row(record)
            if row_number is None:
                ambiguous_count += int(ambiguous)
                not_found += int(not ambiguous)
                continue
            ws.Range(ws.Cells(row_number, 1), ws.Cells(row_number, 31)).Interior.Color = _excel_color(colors[change_type])
            marked_modified += int(change_type == "MODIFIED")
            marked_deleted += int(change_type == "DELETED")
            if record.get("CHANGE_ID"):
                acknowledged_change_ids.append(record["CHANGE_ID"])

        mapped_count = presale_count = 0
        batch_values = []
        row_markers = []
        data_last_column = max(header_columns.values())
        for offset, record in enumerate(append_records, start=1):
            target_row = last_data_row + offset
            row_values = [None] * data_last_column
            remark = str(record.get("REMARK") or "")
            presale = _is_presale(record)
            for header, api_field in API_FIELD_BY_HEADER.items():
                value = record.get(api_field)
                if header == "날짜" and value:
                    value = _excel_date_value(value)
                elif header == "품번":
                    original_part = normalize_part_no(value)
                    value = part_mapping.get(original_part, original_part)
                    mapped_count += int(bool(original_part) and value != original_part)
                elif header == "단위" and presale:
                    value = "선매출"
                elif header == "단위" and normalize_process_token(record.get("CATEGORY")) in EXCLUDED_UNIT_CATEGORIES:
                    value = "제외"
                elif header == "출고량" and presale:
                    value = 0
                row_values[header_columns[header] - 1] = value
            row_values[header_columns["반품/선매출고"] - 1] = (
                record.get("RSQTY") if presale else record.get("RETURNQTY")
            )
            change_type = str(record.get("CHANGE_TYPE") or "").upper()
            batch_values.append(tuple(row_values))
            row_markers.append((target_row, change_type, presale))
            presale_count += int(presale)

        validation_warnings = []
        if append_records:
            first_target_row = last_data_row + 1
            last_target_row = last_data_row + len(append_records)
            if progress:
                progress(62, f"출고 데이터 {len(append_records)}건 서식·수식 준비 중...")

            if formula_row:
                formula_ws.Range(
                    formula_ws.Cells(formula_row, 1), formula_ws.Cells(formula_row, 31)
                ).Copy()
                ws.Range(
                    ws.Cells(first_target_row, 1), ws.Cells(last_target_row, 31)
                ).PasteSpecial(-4122)  # xlPasteFormats
                for column in (30, 31):
                    formula = formula_ws.Cells(formula_row, column).FormulaR1C1
                    if isinstance(formula, str) and formula.startswith("="):
                        ws.Range(
                            ws.Cells(first_target_row, column),
                            ws.Cells(last_target_row, column),
                        ).FormulaR1C1 = formula

            _write_excel_values(ws, first_target_row, batch_values, header_columns, progress)

            all_new_rows = ws.Range(
                ws.Cells(first_target_row, 1), ws.Cells(last_target_row, 31)
            )
            all_new_rows.Interior.Pattern = -4142
            if last_target_row > first_target_row:
                # 일괄 복사된 행별 아래선을 지우고 범위 최하단의 바깥선만 남긴다.
                all_new_rows.Borders(12).LineStyle = -4142  # xlInsideHorizontal / xlNone

            # 변경 유형과 선매출처럼 행별로 다른 색만 개별 적용한다.
            for target_row, change_type, presale in row_markers:
                if change_type == "PAST_NEW":
                    ws.Range(ws.Cells(target_row, 1), ws.Cells(target_row, 31)).Interior.Color = _excel_color("DDEBF7")
                if presale:
                    ws.Range(ws.Cells(target_row, 1), ws.Cells(target_row, 31)).Interior.Color = _excel_color(PRESALE_COLOR)
            try:
                validation_warnings = _verify_excel_values(
                    ws, first_target_row, batch_values, header_columns
                )
            except Exception as exc:
                validation_warnings = [f"입력값 확인 중 오류: {exc}"]
            if progress:
                progress(85, f"출고 데이터 추가 완료: {len(append_records)}건")

        output_path = str(Path(output_file).resolve())
        file_format = 52 if Path(output_file).suffix.lower() == ".xlsm" else 51
        if progress:
            progress(88, "입력 완료: 생산계획 수식 계산 및 계산 모드 복원 중...")
        # 자동 계산 모드 복원 시 누적된 변경을 한 번에 계산한다.
        excel.Calculation = original_calculation
        if progress:
            progress(90, "결과 Excel 파일 저장 중...")
        workbook.SaveAs(output_path, FileFormat=file_format)
        return {
            "added_count": len(append_records), "mapped_count": mapped_count,
            "presale_count": presale_count, "marked_modified_count": marked_modified,
            "marked_deleted_count": marked_deleted, "not_found_count": not_found,
            "ambiguous_count": ambiguous_count, "formula_source_found": formula_row > 0,
            "target_sheet": ws.Name,
            "acknowledged_change_ids": acknowledged_change_ids,
            "validation_warnings": validation_warnings,
        }
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    input_file = filedialog.askopenfilename(
        parent=root,
        title="출고·폐기 엑셀 선택",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
    )
    if not input_file:
        return

    production_plan = messagebox.askyesnocancel(
        "파일 종류 선택",
        "선택한 파일의 종류를 지정해 주세요.\n\n"
        "예: 생산계획 파일 ('출고, 폐기' 시트 사용)\n"
        "아니요: 별도 출고관리 파일 (첫 번째 시트 사용)\n"
        "취소: 작업 취소",
        parent=root,
    )
    if production_plan is None:
        return
    target_mode = "production_plan" if production_plan else "shipment_file"
    consumer = "PRODUCTION_PLAN" if production_plan else "SHIPMENT_DISPOSAL"

    if DEFAULT_MAPPING_FILE.is_file():
        mapping_file = DEFAULT_MAPPING_FILE
        mapping_source = "공용서버 자동 연결"
    else:
        mapping_file = filedialog.askopenfilename(
            parent=root,
            title="품번매핑 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        mapping_source = "직접 선택" if mapping_file else "사용 안 함"

    source = Path(input_file)
    target_date = datetime.now() - timedelta(days=1)
    output_file = filedialog.asksaveasfilename(
        parent=root,
        title="결과 파일 저장",
        initialdir=str(source.parent),
        initialfile=f"{source.stem}_{target_date:%Y%m%d}{source.suffix}",
        defaultextension=source.suffix,
        filetypes=[("Excel files", f"*{source.suffix}")],
    )
    if not output_file:
        return

    progress_win, progress = create_progress_window(root, "출고 변경사항 반영")
    try:
        progress(5, "서버에서 출고 변경 데이터 받는 중...")
        payload = fetch_shipment_changes(consumer)
        progress(15, f"API 데이터 수신 완료: {len(payload['data'])}건")
        result = append_shipment_changes_fast(
            input_file,
            output_file,
            payload["data"],
            mapping_file=mapping_file,
            target_mode=target_mode,
            progress=progress,
        )
        acknowledgement_warning = ""
        try:
            acknowledged_count = acknowledge_shipment_changes(
                consumer,
                result["acknowledged_change_ids"],
                metadata={
                    "operator": getpass.getuser(),
                    "computer": socket.gethostname(),
                    "sourceFile": Path(input_file).name,
                    "outputFile": Path(output_file).name,
                },
            )
        except Exception as exc:
            acknowledged_count = 0
            acknowledgement_warning = (
                "\n\n주의: 엑셀은 저장됐지만 서버 처리 이력 기록에 실패했습니다.\n"
                f"{exc}"
            )
        progress(100, "완료")
        progress_win.destroy()
        summary = payload.get("summary") or {}
        validation_warnings = result.get("validation_warnings", [])
        validation_message = ""
        if validation_warnings:
            validation_message = (
                f"\n\n주의: 입력값 확인 경고 {len(validation_warnings)}건이 있으나 파일은 저장했습니다.\n"
                + "\n".join(validation_warnings[:5])
                + ("\n(나머지 경고 생략)" if len(validation_warnings) > 5 else "")
            )
        messagebox.showinfo(
            "완료",
            f"결과 저장 완료: {output_file}\n"
            f"입력 시트: {result['target_sheet']}\n"
            f"추가: {result['added_count']}건 "
            f"(NEW {summary.get('newCount', 0)}, MODIFIED {summary.get('modifiedCount', 0)}, "
            f"PAST_NEW {summary.get('pastNewCount', 0)}, DELETED {summary.get('deletedCount', 0)})\n"
            f"품번매핑: {result['mapped_count']}건 ({mapping_source})\n"
            f"기존 행 표시: 수정 {result['marked_modified_count']}건, "
            f"삭제 {result['marked_deleted_count']}건\n"
            f"기존 행 미발견: {result['not_found_count']}건, "
            f"중복 후보: {result['ambiguous_count']}건\n"
            f"선매출 표시: {result['presale_count']}건"
            f"\n서버 처리 완료 기록: {acknowledged_count}건"
            f"{'\n주의: 통합문서에서 AD·AE 원본 수식을 찾지 못했습니다.' if not result['formula_source_found'] else ''}"
            f"{acknowledgement_warning}"
            f"{validation_message}",
            parent=root,
        )
    except Exception as exc:
        if progress_win.winfo_exists():
            progress_win.destroy()
        messagebox.showerror("오류", str(exc), parent=root)
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
