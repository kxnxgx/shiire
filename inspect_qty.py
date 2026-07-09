import pandas as pd
import json

try:
    df_qty = pd.read_excel('シレトコ分析.xlsx', sheet_name='数量実績')
    print("数量実績 columns:")
    print(df_qty.columns.tolist())
    print("\n数量実績 head:")
    print(df_qty.head(10).to_string())
except Exception as e:
    print(f"Error: {e}")
