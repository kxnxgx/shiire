import pandas as pd

try:
    df = pd.read_csv('営業日付別売上分析(旧) (1).csv', encoding='cp932')
    
    # Parse 営業日付 from Excel serial number
    df['営業日付'] = pd.to_datetime(df['営業日付'], unit='D', origin='1899-12-30')
    df['year'] = df['営業日付'].dt.year
    df['month'] = df['営業日付'].dt.month
    
    print("CSV Min/Max date:")
    print(df['営業日付'].min(), df['営業日付'].max())
    
    # Check Shiretoko brand
    df_s = df[df['表記部門名1'] == 'Shiretoko']
    print("\nShiretoko Brand Years in CSV:")
    print(df_s['year'].value_counts())
    print("Shiretoko Brand monthly quantities:")
    print(df_s.groupby(['year', 'month'])['数量'].sum())
    print("\nUnique Shiretoko SKUs in CSV:")
    print(df_s['3rd Item No.'].nunique())
    
except Exception as e:
    print(f"Error: {e}")
