import pandas as pd
from pathlib import Path

folder_path = Path("D:\學姊出帳系統\要出帳的Excel")  # e.g. r"C:/data/excels"

# One DataFrame per file, stored in a dict
all_dfs = {}
for file in folder_path.glob("*.xls*"):      # use '*.xls*' if you also have .xls
    df = pd.read_excel(file, sheet_name = "主持人－計畫簡稱（明細）")  
    df.columns = df.iloc[2]
    df = df.iloc[3:].reset_index(drop=True) 
    df = df.iloc[1:, :10]          
    all_dfs[file.stem] = df

combined_df = pd.concat(all_dfs, ignore_index=True)

def search_date_range(df, col, start, end):
    col_str = df[col].astype(str).str.strip()
    return df[(col_str >= str(start)) & (col_str <= str(end))]

# usage
matches = search_date_range(combined_df, "請購日期", "1150101", "1151231")
output_path = Path("D:/學姊出帳系統/results.xlsx")  # adjust path
matches.to_excel(output_path, index=False)