import pdfplumber
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
import sys

def check_table_structure(table):
    if "MARKS & NUMBERS" in table[0]:
        return True
    else:
        return False

def extract_and_process_data(pdf_file_path, columns_to_keep):
    final_df = pd.DataFrame(columns=columns_to_keep)
    
    with pdfplumber.open(pdf_file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if check_table_structure(table):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df = df[columns_to_keep]
                    final_df = pd.concat([final_df, df], ignore_index=True)                
    return final_df

def check_extract_and_process_data(pdf_file_path, columns_to_keep):
    all_tables = []
    with pdfplumber.open(pdf_file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            all_tables.extend(tables)

    dataframes = []

    for table in all_tables:
        if not check_table_structure(table):
            df = convert_to_dataframe(table, columns_to_keep)
            dataframes.append(df)

    final_df = pd.concat(dataframes, ignore_index=True)
    return final_df

def convert_to_dataframe(table, columns_to_keep):
    rows = {}
    for row in table:
        for item in row:
            if item is not None:
                lines = item.split('\n')
                for line in lines:
                    if ":" in line:
                        key, value = line.split(":", 1)
                        key_stripped = key.strip()
                        value_stripped = value.strip()
                        if key_stripped in columns_to_keep:
                            if key_stripped not in rows:
                                rows[key_stripped] = [value_stripped]
                            else:
                                rows[key_stripped].append(value_stripped)
    df = pd.DataFrame(rows)
    return df

def export_to_excel(df, output_excel):
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Auto fit columns
    for cell in ws[1]:
        ws.column_dimensions[cell.column_letter].width = max(len(str(cell.value)) + 2, 15)

    wb.save(output_excel)

pdf_file_path = sys.argv[1]
columns_to_keep_1 = ["HB/L No", "MB/L No", "Container(s)"]
columns_to_keep_2 = ["MARKS & NUMBERS", "WEIGHT", "VOLUME", "PKGS"]
output_excel = sys.argv[2]

# Xử lý dữ liệu từ PDF và lưu vào DataFrame
data_df = extract_and_process_data(pdf_file_path, columns_to_keep_2)

# Kiểm tra và xử lý dữ liệu từ PDF
checked_df = check_extract_and_process_data(pdf_file_path, columns_to_keep_1)
combined_data = {**checked_df, **data_df}    
df = pd.DataFrame(combined_data)

# Export DataFrame thành file Excel và tự động fit kích thước cột
export_to_excel(df, output_excel)
