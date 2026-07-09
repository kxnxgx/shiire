import pandas as pd
import json

try:
    xls = pd.ExcelFile('知床財団在庫表260708.xlsx')
    
    for sheet_name in xls.sheet_names:
        print(f"\n==========================================")
        print(f"Sheet Name: {sheet_name}")
        df = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name=sheet_name)
        print("Columns:", df.columns.tolist())
        print("First 5 rows:")
        print(df.head(5).to_string())
        
except Exception as e:
    print("Error:", e)
