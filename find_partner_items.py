import pandas as pd

try:
    # 1. Partner inventory
    df_partner = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name='知床財団在庫表')
    
    # 2. Sales data (This year)
    df_this = pd.read_excel('シレトコ分析.xlsx', sheet_name='営業日付別売上分析(旧)')
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko']
    this_skus = df_this_s[['3rd Item No.', '商品名']].drop_duplicates()
    
    # 3. Inventory CSV
    df_inv = pd.read_csv('在庫一覧 .csv', encoding='cp932')
    df_inv_s = df_inv[df_inv['BrandName'] == 'Shiretoko']
    inv_skus = df_inv_s[['商品コード', '商品名']].drop_duplicates()
    
    print("=== Sales items in our database ===")
    print(this_skus.to_string())
    
    print("\n=== Inventory CSV items in our database ===")
    print(inv_skus.to_string())
    
except Exception as e:
    print("Error:", e)
