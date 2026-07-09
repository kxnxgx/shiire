import pandas as pd

try:
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    for name in xls.sheet_names:
        df = pd.read_excel('シレトコ分析.xlsx', sheet_name=name)
        print(f"=== Sheet: {name} ===")
        # Search for any cell containing '2024' or check the shape and row sums
        # Let's check the sum of the last column if it is numeric
        last_col = df.columns[-1]
        try:
            total_sum = pd.to_numeric(df.iloc[1:, -1], errors='coerce').sum()
            print(f"Sum of last column: {total_sum}")
        except:
            pass
except Exception as e:
    print("Error:", e)
