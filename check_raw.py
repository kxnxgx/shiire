import pandas as pd

try:
    df = pd.read_excel('シレトコ分析.xlsx', sheet_name='営業日付別売上分析(旧) (1)')
    print("Columns:", df.columns.tolist())
    
    # Check if there is a department name for Shiretoko
    print("\nDistinct 表記部門名1:")
    print(df['表記部門名1'].value_counts())
    
    # Check if any item contains '知床' in 商品名
    shiretoko_items = df[df['商品名'].str.contains('知床', na=False)]
    print(f"\nItems with '知床' in name: {len(shiretoko_items)}")
    if len(shiretoko_items) > 0:
        print(shiretoko_items[['店舗名称', '3rd Item No.', '商品名', '表記部門名1']].head())

except Exception as e:
    print(f"Error: {e}")
