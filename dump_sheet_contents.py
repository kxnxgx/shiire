import pandas as pd
import sys

# Set output encoding to utf-8
sys.stdout.reconfigure(encoding='utf-8')

try:
    xls = pd.ExcelFile('知床財団在庫表260708.xlsx')
    
    for idx, sheet_name in enumerate(xls.sheet_names):
        print(f"\n==========================================")
        print(f"Sheet Name: {sheet_name}")
        df = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name=sheet_name)
        
        # Output first 30 rows and 10 columns
        print(df.iloc[:40, :10].to_string())
        
except Exception as e:
    print("Error:", e)
