import pandas as pd
from datetime import date
import numpy as np

# Đọc file Excel
file_path = "INPUT_DATA/INPUT DATA/output.report sample.xlsx"
df = pd.read_excel(file_path, header=None)

# Lấy dòng 0 (thứ) và dòng 1 (ngày)
weekdays_row = df.iloc[0]
dates_row = pd.to_datetime(df.iloc[1], errors='coerce')  # ép ngày

# Xác định các cột mà:
# - dòng 0 là chữ (không rỗng), và
# - dòng 1 là ngày hợp lệ
valid_cols = [i for i in df.columns if pd.notna(weekdays_row[i]) and pd.notna(dates_row[i])]

if valid_cols:
    last_col = valid_cols[-1]
    last_date = dates_row[last_col].normalize()
    today = pd.Timestamp(date.today())

    if last_date == today:
        print("✅ Ngày hôm nay đã có trong bảng.")
    elif last_date < today:
        insert_at = last_col + 1
        weekday = today.strftime("%a").upper()
        df.insert(loc=insert_at, column=f"col_{insert_at}", value=np.nan)
        df.iloc[0, insert_at] = weekday
        df.iloc[1, insert_at] = today
        print(f"🆕 Đã chèn cột ngày hôm nay ({today.strftime('%Y-%m-%d')}) sau cột {last_col}.")
        # df.to_excel("updated_report.xlsx", index=False, header=False)
else:
    print("⚠️ Không có cột nào vừa có thứ và ngày hợp lệ.")
