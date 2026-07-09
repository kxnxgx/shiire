import pandas as pd

try:
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    for sheet in xls.sheet_names:
        print(f"=== Sheet: {sheet} ===")
        df = pd.read_excel('シレトコ分析.xlsx', sheet_name=sheet)
        print(f"Shape: {df.shape}")
        print("Columns:", df.columns.tolist()[:10])
        # Print a sample of non-null values
        print(df.dropna(how='all').head(5).to_string())
        print("\n")
except Exception as e:
    print(f"Error: {e}")
