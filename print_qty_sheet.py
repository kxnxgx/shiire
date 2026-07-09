import pandas as pd

try:
    df_qty = pd.read_excel('シレトコ分析.xlsx', sheet_name='数量実績')
    # Print columns
    print("All Columns in 数量実績:")
    print(df_qty.columns.tolist())
    # Print the row 0 (which contains headers)
    print("\nRow 0 (headers):")
    print(df_qty.iloc[0].to_dict())
    
    # Print row 1-5
    print("\nRow 1-5:")
    print(df_qty.iloc[1:6].to_string())
    
except Exception as e:
    print(f"Error: {e}")
