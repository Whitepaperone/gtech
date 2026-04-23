from openpyxl import load_workbook
import pandas as pd
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

OUTPUT_FILE = "./output.xlsx"


def find_date_columns(df):
    """
    날짜 열 자동 탐지 (datetime 기반)
    """
    for row_idx in range(0, 10):
        row = df.iloc[row_idx]
        date_cols = []

        for i, val in enumerate(row):
            if isinstance(val, (datetime, pd.Timestamp)):
                date_cols.append(i)

        if len(date_cols) > 5:
            return row_idx, min(date_cols), max(date_cols) + 1

    raise Exception("날짜 열을 찾을 수 없습니다.")


def find_data_start_row(df, start_col, end_col):
    """
    실제 데이터 시작 행 찾기
    """
    for i in range(0, len(df)):
        row = df.iloc[i, start_col:end_col]

        # 숫자 포함된 행 찾기
        numeric_count = sum(
            1 for x in row
            if isinstance(x, (int, float)) and not pd.isna(x)
        )

        if numeric_count >= 1:
            return i

    raise Exception("데이터 시작 행을 찾을 수 없습니다.")


def adjust_row(values):
    """
    재고 → 미래 계획 차감
    """
    stock = 0
    result = values.copy()

    # 첫날 음수 → 선생산량
    if result[0] < 0:
        stock = abs(result[0])
        result[0] = 0

    for i in range(1, len(result)):
        if stock <= 0:
            break

        val = result[i]

        if pd.isna(val):
            continue

        if val > 0:
            if val >= stock:
                result[i] -= stock
                stock = 0
            else:
                stock -= val
                result[i] = 0

    return result
def apply_today_result(row, today_dict):    
    product = row["품번"]

    if product not in today_dict:
        return row

    remain = today_dict[product]

    values = row[3:].tolist()

    for i in range(len(values)):
        if remain <= 0:
            break

        if values[i] > 0:
            if values[i] >= remain:
                values[i] -= remain
                remain = 0
            else:
                remain -= values[i]
                values[i] = 0

    row[3:] = values
    return row

def main():
    Tk().withdraw()

    file_path = filedialog.askopenfilename(
        title="엑셀 선택",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, header=None)

        # 1️⃣ 날짜 열 찾기
        date_row_idx, start_col, end_col = find_date_columns(df)

        # 2️⃣ 데이터 시작 행 찾기
        start_row = find_data_start_row(df, start_col, end_col)

            
        today_df = pd.read_excel(OUTPUT_FILE)  # 시트 이름 또는 index

        # 품번 / 수량 컬럼 맞게 수정 필요
        today_df = today_df[["품번", "실적수량"]]

        today_df["품번"] = today_df["품번"].astype(str).str.strip()

        today_grouped = today_df.groupby("품번")["실적수량"].sum().to_dict()


        product_col = start_col - 1
        customer_code_col = start_col - 2
        model_col = start_col - 3


        date_row = df.iloc[date_row_idx, start_col:end_col].tolist()
        date_row = [ pd.to_datetime(d).strftime("%Y/%m/%d") if not pd.isna(d) else "" for d in date_row]

        processed_rows = []

        # 3️⃣ 처리
        
        for i in range(start_row, len(df)):
            row_full = df.iloc[i]


            product = row_full[product_col]
            model= row_full[model_col]
            customer_code = row_full[customer_code_col]

            if pd.isna(product):
                continue

            row = row_full[start_col:end_col].tolist()
            row = [pd.to_numeric(x, errors='coerce') for x in row]
            row = [0 if pd.isna(x) else x for x in row]

            adjusted = adjust_row(row)

            processed_rows.append([model, customer_code, product] + adjusted)


        # 저장
        # 🔥 DataFrame으로 변환
        result_df = pd.DataFrame(processed_rows)

        result_df.columns = ["모델", "고객사 품번", "품번"] + date_row

        result_df.iloc[:, 3:] = result_df.iloc[:, 3:].apply(pd.to_numeric, errors='coerce').fillna(0)

        agg_dict = {"모델": "first", "고객사 품번": lambda x: x.dropna().iloc[0] if not x.dropna().empty else ""}


        for col in result_df.columns[3:]:
            agg_dict[col] = "sum"

        # 🔥 품번 기준 그룹화 + 합계
        result_df["품번"] = result_df["품번"].astype(str).str.strip()
        result_df = result_df.groupby("품번", as_index=False).agg(agg_dict)
        result_df = result_df.apply(lambda row: apply_today_result(row, today_grouped), axis=1)
        result_df = result_df[(result_df.iloc[:, 3:] != 0).any(axis=1)]

        output_path = file_path.replace(".xlsx", "_MRP.xlsx")

        result_df.to_excel(output_path, index=False)


        wb = load_workbook(output_path)
        ws = wb.active

        for col in range(3, len(result_df.columns) + 1):
            for row in range(1, len(result_df) + 1):  # ⭐ 헤더 포함
                cell = ws.cell(row=row, column=col)
                
                if isinstance(cell.value, datetime):
                    cell.number_format = 'yyyy-mm-dd'

        wb.save(output_path)

        messagebox.showinfo("완료", f"저장 완료:\n{output_path}")

    except Exception as e:
        messagebox.showerror("오류", str(e))


if __name__ == "__main__":
    main()