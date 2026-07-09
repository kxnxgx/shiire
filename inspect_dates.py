import pandas as pd

try:
    xls = pd.ExcelFile('シレトコ分析.xlsx')
    target_sheet = '営業日付別売上分析(旧)'
    if target_sheet not in xls.sheet_names and '営業日付別売上分析(旧) (1)' in xls.sheet_names:
        target_sheet = '営業日付別売上分析(旧) (1)'
        
    df = pd.read_excel('シレトコ分析.xlsx', sheet_name=target_sheet)
    print("Year value counts:")
    print(df['year'].value_counts())
    print("\nMin/Max date:")
    print(df['営業日付'].min(), df['営業日付'].max())
    
    # Check Shiretoko brand years
    df_s = df[df['表記部門名1'] == 'Shiretoko']
    print("\nShiretoko Brand Years:")
    print(df_s['year'].value_counts())
    print("Shiretoko Brand Month value counts for each year:")
    print(df_s.groupby('year')['month'].value_counts())
    
except Exception as e:
    print(f"Error: {e}")
