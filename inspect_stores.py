import pandas as pd

try:
    # 1. Excel (Sales data)
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_this = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko']
    sales_stores_this = df_this_s['店舗名称'].dropna().str.strip().unique()
    
    # 昨年度 (CSV)
    df_last = pd.read_csv('営業日付別売上分析(旧) (1).csv', encoding='cp932')
    df_last_s = df_last[df_last['表記部門名1'] == 'Shiretoko']
    sales_stores_last = df_last_s['店舗名称'].dropna().str.strip().unique()
    
    # Combined sales stores
    sales_stores = list(set(sales_stores_this).union(set(sales_stores_last)))
    
    # 2. Inventory (CSV)
    df_inv = pd.read_csv('在庫一覧 .csv', encoding='cp932')
    df_inv_s = df_inv[df_inv['BrandName'] == 'Shiretoko']
    inv_stores = df_inv_s['拠点名'].dropna().str.strip().unique()
    
    print("=== Sales Data Stores (Shiretoko) ===")
    for s in sorted(sales_stores):
        print(f"- '{s}'")
        
    print("\n=== Inventory CSV Stores (Shiretoko) ===")
    for s in sorted(inv_stores):
        print(f"- '{s}'")
        
except Exception as e:
    print("Error:", e)
