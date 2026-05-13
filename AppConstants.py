import json
import sys
from pathlib import Path


DEFAULT_CONFIG = {
    "plan_sheet_names": ["완성공정(실적)", "TANK공정(실적)", "CORE공정(실적)"],
    "finish_plan_sheet_names": ["완성공정(실적)"],
    "plan_output_file": "./생산계획_계획수량_추출.xlsx",
    "mes_output_file": "./계획수량_작업지시_비교결과.xlsx",
    "header_date_row": 33,
    "header_kind_row": 34,
    "data_start_row": 35,
    "date_scan_min_col": 1,
    "date_header_gap_break": 3,
    "mes_sheet_config": {
        "완성공정(실적)": {"default_process": "완성"},
        "TANK공정(실적)": {"default_process": "TANK"},
        "CORE공정(실적)": {"default_process": "CORE"},
    },
    "workshop_to_process": {
        "용접": "TANK",
        "CLINCHING": "TANK",
        "TANK조립&LEAKTEST": "TANK",
        "CORE조립": "CORE",
        "출하-액세서리": "완성",
        "완성조립": "완성",
    },
    "final_compare_columns": [
        "품번", "날짜", "공정", "작업장명", "작업반명", "공정인쇄",
        "계획수량", "미달수량",
        "계획표실적수량", "MES실적수량", "전일마감기준실적수량", "계획대비실적수량",
        "작업지시수량", "판정",
    ],
    "coa_exclude_customers": ["HD건설기계㈜인천"],
}


def _candidate_config_paths():
    if getattr(sys, "frozen", False):
        yield Path(sys.executable).resolve().with_name("app_config.json")

    yield Path(__file__).resolve().with_name("app_config.json")


def _load_config():
    for path in _candidate_config_paths():
        if path.exists():
            with path.open("r", encoding="utf-8") as f:
                loaded = json.load(f)
            return {**DEFAULT_CONFIG, **loaded}

    return DEFAULT_CONFIG


CONFIG = _load_config()

PLAN_SHEET_NAMES = CONFIG["plan_sheet_names"]
FINISH_PLAN_SHEET_NAMES = CONFIG["finish_plan_sheet_names"]

PLAN_OUTPUT_FILE = CONFIG["plan_output_file"]
MES_OUTPUT_FILE = CONFIG["mes_output_file"]

HEADER_DATE_ROW = CONFIG["header_date_row"]
HEADER_KIND_ROW = CONFIG["header_kind_row"]
DATA_START_ROW = CONFIG["data_start_row"]
DATE_SCAN_MIN_COL = CONFIG["date_scan_min_col"]
DATE_HEADER_GAP_BREAK = CONFIG["date_header_gap_break"]

MES_SHEET_CONFIG = CONFIG["mes_sheet_config"]
WORKSHOP_TO_PROCESS = CONFIG["workshop_to_process"]
FINAL_COMPARE_COLUMNS = CONFIG["final_compare_columns"]
COA_EXCLUDE_CUSTOMERS = CONFIG["coa_exclude_customers"]
