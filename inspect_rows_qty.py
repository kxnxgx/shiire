import pandas as pd

try:
    df_qty = pd.read_excel('シレトコ分析.xlsx', sheet_name='数量実績')
    # Row 0 contains headers
    header = df_qty.iloc[0]
    data = df_qty.iloc[1:].copy()
    data.columns = header
    
    # Check for duplicates of (店舗名称, 3rd Item No.)
    dupes = data[data.duplicated(subset=['店舗名称', '3rd Item No.'], keep=False)]
    print("Duplicates in 数量実績 of (店舗名称, 3rd Item No.):")
    print(len(dupes))
    if len(dupes) > 0:
        print(dupes.head(10).to_string())
        
    print("\nUnique 3rd Item No. in 数量実績:", data['3rd Item No.'].nunique())
    print("Unique 3rd Item No. in 営業日付別売上分析(旧) (Shiretoko only):")
    
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_raw = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_raw_s = df_raw[df_raw['表記部門名1'] == 'Shiretoko']
    print(df_raw_s['3rd Item No.'].nunique())
    
    # Are the item numbers in 数量実績 starting with SNF?
    print("\nDoes 数量実績 contain non-SNF items?")
    print(data[~data['3rd Item No.'].str.startswith('SNF', na=False)]['3rd Item No.'].unique())
    
except Exception as e:
    print(f"Error: {e}")
