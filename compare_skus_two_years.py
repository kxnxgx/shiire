import pandas as pd

try:
    # 1. 昨年度の売上データ（CSV）
    df_last = pd.read_csv('営業日付別売上分析(旧) (1).csv', encoding='cp932')
    df_last_s = df_last[df_last['表記部門名1'] == 'Shiretoko']
    last_skus = set(df_last_s['3rd Item No.'].dropna().unique())
    
    # 2. 今年度の売上データ（Excel）
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_this = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko']
    this_skus = set(df_this_s['3rd Item No.'].dropna().unique())
    
    print("--- 昨年度 vs 今年度 品番比較分析 ---")
    print(f"昨年度 (2024/7〜2025/6) の売上品番数: {len(last_skus)}")
    print(f"今年度 (2025/7〜2026/6) の売上品番数: {len(this_skus)}")
    
    common = last_skus.intersection(this_skus)
    dropped = last_skus - this_skus
    added = this_skus - last_skus
    
    print(f"共通する売上品番数: {len(common)}")
    print(f"昨年度のみ売上があり、今年度売上がない品番数 (販売終了など): {len(dropped)}")
    print(f"今年度のみ新しく売上が発生した品番数 (新商品など): {len(added)}")
    
    if dropped:
        print("\n【昨年度のみ売上ありの品番】")
        for sku in sorted(dropped):
            # Find product name
            name = df_last_s[df_last_s['3rd Item No.'] == sku]['商品名'].iloc[0]
            print(f"- {sku}: {name}")
            
    if added:
        print("\n【今年度のみ新しく売上ありの品番】")
        for sku in sorted(added):
            name = df_this_s[df_this_s['3rd Item No.'] == sku]['商品名'].iloc[0]
            print(f"- {sku}: {name}")
            
except Exception as e:
    print("Error:", e)
