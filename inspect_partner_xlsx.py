import pandas as pd
import sys

# Windows terminal encoding support
sys.stdout.reconfigure(encoding='utf-8')

try:
    excel_file = '知床財団在庫表260708.xlsx'
    xls = pd.ExcelFile(excel_file)
    print("=== シート一覧 ===")
    print(xls.sheet_names)
    
    for name in xls.sheet_names:
        print(f"\n=== シート: {name} (データサンプル) ===")
        df = pd.read_excel(excel_file, sheet_name=name)
        # Drop rows/columns that are completely empty to make it readable
        df_cleaned = df.dropna(how='all').dropna(axis=1, how='all')
        print(df_cleaned.head(20).to_string())
        
except Exception as e:
    print("Error:", e)
