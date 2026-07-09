import pandas as pd
import numpy as np
import os

excel_file = 'シレトコ分析.xlsx'
last_year_csv = '営業日付別売上分析(旧) (1).csv'
inventory_csv = '在庫一覧 .csv'

print("データを読み込んでいます...")

try:
    # 1. 昨年度の売上データ（CSV）
    if not os.path.exists(last_year_csv):
        print(f"Error: {last_year_csv} が見つかりません。")
        exit(1)
    df_last = pd.read_csv(last_year_csv, encoding='cp932')
    df_last_s = df_last[df_last['表記部門名1'] == 'Shiretoko'].copy()
    
    # 2. 今年度の売上データ（Excel）
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} が見つかりません。")
        exit(1)
        
    xls = pd.ExcelFile(excel_file)
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_this = pd.read_excel(excel_file, sheet_name=target_sheet)
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko'].copy()
    
    # 3. 在庫データ（CSV）
    if not os.path.exists(inventory_csv):
        print(f"Error: {inventory_csv} が見つかりません。")
        exit(1)
    df_inv = pd.read_csv(inventory_csv, encoding='cp932')
    df_inv_s = df_inv[df_inv['BrandName'] == 'Shiretoko'].copy()
    
    # 店舗名統一関数
    def normalize_store_name(name):
        if not isinstance(name, str):
            return name
        name = name.strip()
        if 'SAPPORO' in name:
            return 'FJALLRAVEN by 3NITY SAPPORO HUTTE'
        if name == 'TOKYO' or '3NITY TOKYO' in name:
            return 'FJALLRAVEN by 3NITY TOKYO'
        if 'ルクア大阪' in name:
            return 'FJALLRAVEN by 3NITY ルクア大阪'
        if '大丸心斎橋' in name or '心斎橋パルコ' in name:
            return 'FJALLRAVEN by 3NITY 大丸心斎橋'
        if '玉川高島屋' in name:
            return 'FJALLRAVEN by 3NITY玉川高島屋S・C'
        if '小田急町田' in name:
            return 'FJALLRAVEN by 3NITY 小田急町田'
        return name
        
    df_last_s['店舗名称'] = df_last_s['店舗名称'].apply(normalize_store_name)
    df_this_s['店舗名称'] = df_this_s['店舗名称'].apply(normalize_store_name)
    df_inv_s['拠点名'] = df_inv_s['拠点名'].apply(normalize_store_name)
    
    # --- SKUレベルの集計 ---
    sku_last = df_last_s.groupby('3rd Item No.').agg(
        商品名=('商品名', 'first'),
        昨年度数量=('数量', 'sum')
    ).reset_index()
    
    sku_this = df_this_s.groupby('3rd Item No.').agg(
        今年度数量=('数量', 'sum')
    ).reset_index()
    
    sku_inv = df_inv_s.groupby('商品コード').agg(
        現在庫数=('現在数量', 'sum')
    ).reset_index().rename(columns={'商品コード': '3rd Item No.'})
    
    sku_summary = pd.merge(sku_last, sku_this, on='3rd Item No.', how='outer')
    sku_summary = pd.merge(sku_summary, sku_inv, on='3rd Item No.', how='outer')
    sku_summary['昨年度数量'] = sku_summary['昨年度数量'].fillna(0).astype(int)
    sku_summary['今年度数量'] = sku_summary['今年度数量'].fillna(0).astype(int)
    sku_summary['現在庫数'] = sku_summary['現在庫数'].fillna(0).astype(int)
    
    # 商品名の名寄せ補完
    inv_name_map = df_inv_s.set_index('商品コード')['商品名'].to_dict()
    sku_summary['商品名'] = sku_summary['商品名'].fillna(sku_summary['3rd Item No.'].map(inv_name_map))
    
    # 1. 売上トレンド分析 (全体)
    print("=== 全体販売実績トレンド ===")
    total_last = sku_summary['昨年度数量'].sum()
    total_this = sku_summary['今年度数量'].sum()
    total_inv = sku_summary['現在庫数'].sum()
    print(f"昨年度総販売数: {total_last}点")
    print(f"今年度総販売数: {total_this}点 (前年比: {total_this/total_last*100:.1f}%)")
    print(f"現在庫総数: {total_inv}点 (今年度販売数の {total_inv/total_this*12:.1f}ヶ月分)")
    
    # 2. 過剰在庫（デッドストック）の検出
    sku_summary['在庫月数'] = np.where(sku_summary['今年度数量'] > 0, 
                                       sku_summary['現在庫数'] / sku_summary['今年度数量'] * 12, 
                                       np.where(sku_summary['現在庫数'] > 0, 999, 0))
                                       
    excess_stock = sku_summary[(sku_summary['在庫月数'] >= 18) & (sku_summary['現在庫数'] >= 5)].sort_values(by='現在庫数', ascending=False)
    print("\n=== 過剰在庫（デッドストック）の懸念がある品番 ===")
    for _, row in excess_stock.iterrows():
        status = f"{row['在庫月数']:.1f}ヶ月分" if row['在庫月数'] < 999 else "今年度販売実績なし"
        print(f"- {row['3rd Item No.']} | {row['商品名']} | 在庫: {row['現在庫数']}個 | 今年度販売: {row['今年度数量']}個 | 期間: {status}")

    # 3. 欠品・在庫不足（即発注が必要）の検出
    shortage_stock = sku_summary[(sku_summary['今年度数量'] >= 10) & (sku_summary['在庫月数'] < 1.5)].sort_values(by='今年度数量', ascending=False)
    print("\n=== 在庫不足・欠品（早急な発注が必要）の品番 ===")
    for _, row in shortage_stock.iterrows():
        print(f"- {row['3rd Item No.']} | {row['商品名']} | 在庫: {row['現在庫数']}個 | 今年度販売: {row['今年度数量']}個 | 在庫月数: {row['在庫月数']:.1f}ヶ月分")

    # 4. 店舗間移動（店舗デポ）の可能性の検出
    print("\n=== 店舗間移動（店舗間在庫調整）で解決できるミスマッチ ===")
    
    store_this = df_this_s.groupby(['店舗名称', '3rd Item No.'])['数量'].sum().reset_index().rename(columns={'数量': '今年度売上'})
    store_inv = df_inv_s.groupby(['拠点名', '商品コード'])['現在数量'].sum().reset_index().rename(columns={'拠点名': '店舗名称', '商品コード': '3rd Item No.', '現在数量': '現在在庫'})
    
    store_m = pd.merge(store_this, store_inv, on=['店舗名称', '3rd Item No.'], how='outer').fillna(0)
    store_m['今年度売上'] = store_m['今年度売上'].astype(int)
    store_m['現在在庫'] = store_m['現在在庫'].astype(int)
    
    items = store_m['3rd Item No.'].unique()
    
    for item in items:
        df_item = store_m[store_m['3rd Item No.'] == item]
        # 在庫余り店舗: 在庫が3個以上あり、売上が今年度0個か1個
        surplus_stores = df_item[(df_item['現在在庫'] >= 3) & (df_item['今年度売上'] <= 1)]
        # 在庫不足店舗: 在庫が0個で、今年度売上が3個以上
        deficit_stores = df_item[(df_item['現在在庫'] == 0) & (df_item['今年度売上'] >= 3)]
        
        if not surplus_stores.empty and not deficit_stores.empty:
            # 商品名を探す
            name_find = df_this_s[df_this_s['3rd Item No.'] == item]['商品名'].unique()
            item_name = name_find[0] if len(name_find) > 0 else item
            
            print(f"\n商品: {item} ({item_name})")
            print("  ▼ 在庫が余っている店舗 (他店へ移動可能):")
            for _, r in surplus_stores.iterrows():
                print(f"    - {r['店舗名称']}: 在庫 {r['現在在庫']}個 / 年間売上 {r['今年度売上']}個")
            print("  ▲ 在庫が切れている店舗 (補充が必要):")
            for _, r in deficit_stores.iterrows():
                print(f"    - {r['店舗名称']}: 在庫 {r['現在在庫']}個 / 年間売上 {r['今年度売上']}個")

except Exception as e:
    print("Error:", e)
