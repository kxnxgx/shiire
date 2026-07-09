import pandas as pd

try:
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
        
    df_raw = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_qty = pd.read_excel('シレトコ分析.xlsx', sheet_name='数量実績')
    
    # Let's inspect the sum of quantities by month in raw data
    df_raw_s = df_raw[df_raw['表記部門名1'] == 'Shiretoko'].copy()
    print("Raw data monthly sum for Shiretoko:")
    print(df_raw_s.groupby(['year', 'month'])['数量'].sum())
    
    print("\n数量実績 sheet sum by column:")
    # columns are like 'month' (which is 202701.0), Unnamed: 4 (202702.0), etc.
    # The header is at row 0.
    df_qty_header = df_qty.iloc[0]
    df_qty_data = df_qty.iloc[1:]
    
    # Print columns and sum of values (excluding the 'v' (total) column)
    for idx, col in enumerate(df_qty.columns):
        if idx >= 3:
            col_name = df_qty_header[col]
            col_sum = pd.to_numeric(df_qty_data[col], errors='coerce').sum()
            print(f"Col {col} ({col_name}): {col_sum}")
            
except Exception as e:
    print(f"Error: {e}")
