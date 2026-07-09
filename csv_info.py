import pandas as pd
import json

try:
    df = pd.read_csv('在庫一覧 .csv', encoding='cp932')
    info = {
        'columns': df.columns.tolist(),
        'sample': json.loads(df.head(5).to_json(orient='records', force_ascii=False))
    }
    with open('csv_info.json', 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
except Exception as e:
    with open('csv_info.json', 'w', encoding='utf-8') as f:
        json.dump({"error": str(e)}, f)
