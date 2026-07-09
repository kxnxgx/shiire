import pandas as pd

try:
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
        
    df_raw = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    df_qty = pd.read_excel('シレトコ分析.xlsx', sheet_name='数量実績')
    
    # Let's filter both for a specific store and item
    # e.g., Store: "大丸心斎橋", Item: "SNF002012X"
    df_raw_sel = df_raw[(df_raw['店舗名称'].str.strip() == '大丸心斎橋') & (df_raw['3rd Item No.'] == 'SNF002012X')]
    print("Raw data transactions for 大丸心斎橋 / SNF002012X:")
    print(df_raw_sel[['year', 'month', '数量']])
    
    # Look at 数量実績
    # The columns are ['店舗名称', '3rd Item No.', '商品名', ...]
    # Let's find the row in df_qty
    df_qty_sel = df_qty[(df_qty.iloc[:, 0].str.strip() == '大丸心斎橋') & (df_qty.iloc[:, 1] == 'SNF002012X')]
    print("\n数量実績 row for 大丸心斎橋 / SNF002012X:")
    print(df_qty_sel.to_string())
    
except Exception as e:
    print(f"Error: {e}")
