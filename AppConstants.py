PLAN_SHEET_NAMES = ["완성공정(실적)", "TANK공정(실적)", "CORE공정(실적)"]
FINISH_PLAN_SHEET_NAMES = ["완성공정(실적)"]

PLAN_OUTPUT_FILE = "./생산계획_계획수량_추출.xlsx"
MES_OUTPUT_FILE = "./계획수량_작업지시_비교결과.xlsx"

HEADER_DATE_ROW = 33
HEADER_KIND_ROW = 34
DATA_START_ROW = 35
DATE_SCAN_MIN_COL = 1
DATE_HEADER_GAP_BREAK = 3

MES_SHEET_CONFIG = {
    "완성공정(실적)": {"default_process": "완성"},
    "TANK공정(실적)": {"default_process": "TANK"},
    "CORE공정(실적)": {"default_process": "CORE"},
}

WORKSHOP_TO_PROCESS = {
    "용접": "TANK",
    "CLINCHING": "TANK",
    "TANK조립&LEAKTEST": "TANK",
    "CORE조립": "CORE",
    "출하-액세서리": "완성",
    "완성조립": "완성",
}

FINAL_COMPARE_COLUMNS = [
    "품번", "날짜", "공정", "작업장명", "작업반명", "공정인쇄",
    "계획수량", "미달수량",
    "계획표실적수량", "MES실적수량", "전일마감기준실적수량", "계획대비실적수량",
    "작업지시수량", "판정",
]

COA_EXCLUDE_CUSTOMERS = [
    "HD건설기계㈜인천",
]
