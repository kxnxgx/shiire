import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    df = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name='知床財団在庫表')
    print("=== 知床財団在庫表 (All Rows) ===")
    print(df.to_string())
except Exception as e:
    print("Error:", e)
