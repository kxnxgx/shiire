import pandas as pd

try:
    # 1. Excel (Sales data this year)
    df_this = pd.read_excel('シレトコ分析.xlsx', sheet_name='営業日付別売上分析(旧)')
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko']
    print("=== This Year (Excel) Unique Stores ===")
    print(df_this_s['店舗名称'].dropna().str.strip().unique())
    
    # 2. CSV (Sales data last year)
    df_last = pd.read_csv('営業日付別売上分析(旧) (1).csv', encoding='cp932')
    df_last_s = df_last[df_last['表記部門名1'] == 'Shiretoko']
    print("\n=== Last Year (CSV) Unique Stores ===")
    print(df_last_s['店舗名称'].dropna().str.strip().unique())
    
    # 3. Inventory (CSV)
    df_inv = pd.read_csv('在庫一覧 .csv', encoding='cp932')
    df_inv_s = df_inv[df_inv['BrandName'] == 'Shiretoko']
    print("\n=== Inventory CSV Unique Stores ===")
    print(df_inv_s['拠点名'].dropna().str.strip().unique())
    
except Exception as e:
    print("Error:", e)
