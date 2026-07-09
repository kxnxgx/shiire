import pandas as pd

try:
    df = pd.read_excel('シレトコ分析.xlsx', sheet_name='全店実績')
    print(df.to_string())
except Exception as e:
    print(f"Error: {e}")
