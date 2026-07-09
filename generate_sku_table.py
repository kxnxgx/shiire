import pandas as pd
import os

try:
    # 1. 昨年度の売上データ（CSV）
    df_last = pd.read_csv('営業日付別売上分析(旧) (1).csv', encoding='cp932')
    df_last_s = df_last[df_last['表記部門名1'] == 'Shiretoko'].copy()
    
    # 2. 今年度の売上データ（Excel）
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
    df_this = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_this_s = df_this[df_this['表記部門名1'] == 'Shiretoko'].copy()
    
    # 品番ごとの売上数量を集計
    last_qty = df_last_s.groupby('3rd Item No.').agg(
        商品名=('商品名', 'first'),
        昨年度数量=('数量', 'sum')
    ).reset_index()
    
    this_qty = df_this_s.groupby('3rd Item No.').agg(
        商品名=('商品名', 'first'),
        今年度数量=('数量', 'sum')
    ).reset_index()
    
    # マージして一覧表を作成
    m = pd.merge(last_qty, this_qty, on='3rd Item No.', how='outer', suffixes=('_last', '_this'))
    
    # 区分を決定（商品名の有無で判定）
    def get_category(row):
        has_last = pd.notna(row['商品名_last'])
        has_this = pd.notna(row['商品名_this'])
        if has_last and has_this:
            return '共通取り扱い'
        elif has_last:
            return '昨年度のみ取り扱い'
        else:
            return '今年度のみ取り扱い'
            
    m['区分'] = m.apply(get_category, axis=1)
    
    # 商品名を統合
    m['商品名'] = m['商品名_this'].fillna(m['商品名_last'])
    m['昨年度数量'] = m['昨年度数量'].fillna(0).astype(int)
    m['今年度数量'] = m['今年度数量'].fillna(0).astype(int)
    
    # ソート順を整理（区分、商品コード順）
    m = m.sort_values(by=['区分', '3rd Item No.'])
    
    # マークダウンテーブルの構築
    md_content = """# 知床財団（Shiretoko）商品コード・商品名 年度別取り扱い一覧表

過去2年間の売上実績があるすべての知床財団（Shiretoko）製品について、商品コード（品番）、商品名、および各年度の販売数量実績をまとめました。

## 取り扱い商品一覧

| 商品コード (品番) | 商品名 | 昨年度数量<br>(2024/7〜2025/6) | 今年度数量<br>(2025/7〜2026/6) | 区分 |
| :--- | :--- | :---: | :---: | :---: |
"""
    
    for _, row in m.iterrows():
        md_content += f"| `{row['3rd Item No.']}` | {row['商品名']} | {row['昨年度数量']:,} | {row['今年度数量']:,} | {row['区分']} |\n"
        
    # 保存先アーティファクトパス
    artifact_path = 'C:/Users/kesuzuki/.gemini/antigravity/brain/9dd811da-a282-444f-9d1d-2f91822b5bf5/sku_comparison_table.md'
    
    os.makedirs(os.path.dirname(artifact_path), exist_ok=True)
    with open(artifact_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
        
    print(f"成功: {artifact_path} にマークダウンテーブルを保存しました。")
    
except Exception as e:
    print("Error:", e)
