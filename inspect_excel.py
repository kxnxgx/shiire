import pandas as pd
import json

try:
    xl = pd.ExcelFile('シレトコ分析.xlsx')
    info = {}
    info['sheets'] = xl.sheet_names
    info['data'] = {}
    
    for sheet in xl.sheet_names:
        df = pd.read_excel('シレトコ分析.xlsx', sheet_name=sheet, nrows=10)
        info['data'][sheet] = df.to_json(orient='records', force_ascii=False)
        
    with open('excel_info.json', 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
except Exception as e:
    with open('excel_info.json', 'w', encoding='utf-8') as f:
        json.dump({"error": str(e)}, f)
