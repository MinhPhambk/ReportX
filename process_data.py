from openpyxl import load_workbook
from datetime import date
import pandas as pd
import os
from copy import copy

def copy_cell_format(src_cell, tgt_cell):
    if src_cell.has_style:
        if src_cell.font: tgt_cell.font = copy(src_cell.font)
        if src_cell.border: tgt_cell.border = copy(src_cell.border)
        if src_cell.fill: tgt_cell.fill = copy(src_cell.fill)
        if src_cell.number_format: tgt_cell.number_format = src_cell.number_format
        if src_cell.protection: tgt_cell.protection = copy(src_cell.protection)
        if src_cell.alignment: tgt_cell.alignment = copy(src_cell.alignment)

def update_report_file():
    file_path = "INPUT DATA/INPUT DATA/output.report sample.xlsx"
    today = pd.Timestamp(date.today())
    today_str = today.strftime("%Y-%m-%d")

    # Tạo tên file mới
    dir_name, base_name = os.path.split(file_path)
    name_part, ext = os.path.splitext(base_name)
    new_file = os.path.join(dir_name, f"{name_part}_{today_str}{ext}")

    # Đọc dữ liệu kiểm tra ngày
    df = pd.read_excel(file_path, header=None)
    weekdays_row = df.iloc[0]
    dates_row = pd.to_datetime(df.iloc[1], errors='coerce')

    valid_cols = [i for i in df.columns if pd.notna(weekdays_row[i]) and pd.notna(dates_row[i])]
    if not valid_cols:
        print("⚠️ Không tìm thấy cột hợp lệ.")
        return

    last_col = valid_cols[-1]
    last_date = dates_row[last_col].normalize()

    if last_date == today:
        print("✅ Ngày hôm nay đã có trong bảng.")
        return

    insert_at = last_col + 1
    weekday = today.strftime("%a").upper()

    # Load workbook
    wb = load_workbook(file_path)
    ws = wb.active

    ws.insert_cols(insert_at + 1)

    # Copy format từ cột trước đó
    for row in range(1, ws.max_row + 1):
        src_cell = ws.cell(row=row, column=insert_at)
        tgt_cell = ws.cell(row=row, column=insert_at + 1)
        copy_cell_format(src_cell, tgt_cell)

    # Ghi dữ liệu ngày và thứ
    ws.cell(row=1, column=insert_at + 1).value = weekday
    ws.cell(row=2, column=insert_at + 1).value = today

    wb.save(new_file)
    print(f"✅ Đã lưu file mới: {new_file}")

def update_DL_to_HN_summary(file_path_input, file_path_report):
    pass

update_report_file()
