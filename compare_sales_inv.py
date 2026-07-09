import pandas as pd

try:
    # 1. 営業日付別売上分析(旧)（売上実績）からShiretokoのユニーク品番数を取得
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_raw = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_sales = df_raw[df_raw['表記部門名1'] == 'Shiretoko']
    sales_skus = set(df_sales['3rd Item No.'].dropna().unique())
    
    # 2. 在庫一覧 .csv からShiretokoのユニーク品番数を取得
    df_inv = pd.read_csv('在庫一覧 .csv', encoding='cp932')
    df_inv_s = df_inv[df_inv['BrandName'] == 'Shiretoko']
    inv_skus = set(df_inv_s['商品コード'].dropna().unique())
    
    print("--- 比較結果 ---")
    print(f"売上実績（営業日付別売上分析）の品番数: {len(sales_skus)}")
    print(f"在庫一覧の品番数: {len(inv_skus)}")
    
    # 重なり
    common = sales_skus.intersection(inv_skus)
    print(f"共通する品番数: {len(common)}")
    print(f"売上実績にのみ存在する品番数: {len(sales_skus - inv_skus)}")
    print(f"在庫一覧にのみ存在する品番数: {len(inv_skus - sales_skus)}")
    
    # 売上のみ存在のリスト
    if sales_skus - inv_skus:
        print("\n【売上実績のみにあり、在庫一覧にない品番】")
        for sku in list(sales_skus - inv_skus)[:10]:
            name = df_sales[df_sales['3rd Item No.'] == sku]['商品名'].iloc[0]
            print(f"- {sku}: {name}")
            
    # 在庫のみ存在のリスト
    if inv_skus - sales_skus:
        print("\n【在庫一覧のみにあり、売上実績にない品番】")
        for sku in list(inv_skus - sales_skus)[:10]:
            name = df_inv_s[df_inv_s['商品コード'] == sku]['商品名'].iloc[0]
            print(f"- {sku}: {name}")
            
except Exception as e:
    print("Error:", e)
