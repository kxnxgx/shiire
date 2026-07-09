import openpyxl
import pandas as pd

try:
    xls = pd.ExcelFile('知床財団在庫表260708.xlsx')
    print("Sheets in partner inventory:")
    print(xls.sheet_names)
    
    # Read first sheet
    df = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name=xls.sheet_names[0])
    print("\nColumns in first sheet:")
    print(df.columns.tolist())
    print("\nFirst 10 rows:")
    print(df.head(10))
    
except Exception as e:
    print("Error:", e)
