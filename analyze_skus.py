import pandas as pd

def analyze_skus():
    try:
        # Excelファイルの読み込み
        xls = pd.ExcelFile('シレトコ分析.xlsx')
        target_sheet = '営業日付別売上分析(旧)'
        if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
            target_sheet = '営業日付別売上分析(旧) (1)'
            
        df_raw = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
        
        # Shiretokoブランドに絞り込み
        df = df_raw[df_raw['表記部門名1'] == 'Shiretoko'].copy()
        
        # 昨年度と今年度のフラグを作成
        def get_fiscal_year(y, m):
            if m >= 7:
                return f"{int(y)}07-{int(y)+1}06"
            else:
                return f"{int(y)-1}07-{int(y)}06"
                
        df['FiscalYear'] = df.apply(lambda row: get_fiscal_year(row['year'], row['month']), axis=1)
        
        col_last_year = '202407-202506'
        col_this_year = '202507-202606'
        
        # 昨年度の売上実績がある品番 (数量 > 0)
        df_last = df[(df['FiscalYear'] == col_last_year) & (df['数量'] > 0)]
        skus_last = set(df_last['3rd Item No.'].unique())
        
        # 今年度の売上実績がある品番 (数量 > 0)
        df_this = df[(df['FiscalYear'] == col_this_year) & (df['数量'] > 0)]
        skus_this = set(df_this['3rd Item No.'].unique())
        
        print(f"--- Shiretokoブランド 品番数比較分析 ---")
        print(f"昨年度 ({col_last_year}) の取り扱い品番数: {len(skus_last)} 品番")
        print(f"今年度 ({col_this_year}) の取り扱い品番数: {len(skus_this)} 品番")
        
        # 共通の品番
        common = skus_last.intersection(skus_this)
        print(f"両年度で共通して売上があった品番数: {len(common)} 品番")
        
        # 昨年度のみ売上があった品番
        only_last = skus_last - skus_this
        print(f"昨年度のみ売上があり、今年度は売れなかった（または取り扱いがなくなった）品番数: {len(only_last)} 品番")
        if only_last:
            print("\n【昨年度のみの品番例】")
            for sku in list(only_last)[:10]:
                item_name = df_last[df_last['3rd Item No.'] == sku]['商品名'].iloc[0]
                print(f"- {sku}: {item_name}")
                
        # 今年度のみ売上があった品番
        only_this = skus_this - skus_last
        print(f"\n今年度から新しく売上があった品番数: {len(only_this)} 品番")
        if only_this:
            print("【今年度のみの品番例】")
            for sku in list(only_this)[:10]:
                item_name = df_this[df_this['3rd Item No.'] == sku]['商品名'].iloc[0]
                print(f"- {sku}: {item_name}")
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    analyze_skus()
