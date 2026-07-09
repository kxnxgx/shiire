import pandas as pd
import numpy as np

try:
    # 1. Partner inventory
    df_partner = pd.read_excel('知床財団在庫表260708.xlsx', sheet_name='知床財団在庫表')
    
    # 2. Sales and internal inventory (using our prepare_order logic)
    # We will just write the store mapping and match names
    partner_name_map = {
        '知床財団ｵﾘｼﾞﾅﾙｴｺﾊﾞｯｸﾞ2018（森）': 'SNF005X',
        '知床財団ｵﾘｼﾞﾅﾙｴｺﾊﾞｯｸﾞ（海）': 'SNF011X',
        '知床財団ﾛｺﾞｽﾃｯｶｰ': 'SNF014X',
        '知床財団ｴｺﾎﾞﾄﾙ（0.5Lｸﾞﾚｰ）': 'SNF002012X',
        '知床財団ｴｺﾎﾞﾄﾙ（0.5Lﾌﾞﾙｰ）': 'SNF002032X',
        '知床財団ｴｺﾎﾞﾄﾙ（0.5Lｸﾘｱ）': 'SNF002010X',
        '知床財団ｴｺﾎﾞﾄﾙ（0.5Lｼｰﾌｫｰﾑ）': 'SNF002031X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｸﾞﾚｰ）': 'SNF001012X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｸﾘｱ）': 'SNF001010X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lﾁｪﾘｰﾌﾞﾛｯｻﾑ）': 'SNF001030X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lﾊﾞﾀｰ）': 'SNF001029X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｺｯﾄﾝ）': 'SNF001032X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lﾃﾞﾆﾑ）': 'SNF001033X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｼﾞㄟﾄﾞ）': 'SNF001034X',  # handles potential encoding issues
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｼﾞｪｲﾄﾞ）': 'SNF001034X',
        '知床財団ｴｺﾎﾞﾄﾙ（1.0Lｼｰﾌｫｰﾑ）': 'SNF001011X',
        '知床財団ｴｺ手袋（ﾌｷとｸﾏ）': 'SNF008031X',
        '知床財団ｴｺ軍手（ｳﾆとｺﾝﾌﾞ）': 'SNF008032X',
        '知床財団ロゴ刺繍今治ｶﾞｰｾﾞﾊﾝｶﾁ（ﾈｲﾋﾞｰ)': 'SNF004020X',
        '知床財団ﾛｺﾞ刺繍今治ｶﾞｰｾﾞﾊﾝｶﾁ（ﾍﾞｰｼﾞｭ)': 'SNF004056X',
        '知床財団ﾛｺﾞ刺繍今治ｶﾞｰｾﾞﾊﾝｶﾁ（ｸﾞﾚｰﾍﾞｰｼﾞｭ)': 'SNF004057X',
        '知床財団ﾛｺﾞ刺繍今治ｶﾞｰｾﾞﾊﾝｶﾁ（ﾓﾈﾋﾟﾝｸ)': 'SNF004058X',
        '知床財団ﾛｺﾞ刺繍今治ｶﾞｰｾﾞﾊﾝｶﾁ（ﾀ-ｺｲｽﾞﾌﾞﾙｰ)': 'SNF004059X',
        'シレトコ野帳（ヒグマ）': 'SNF009035X',
        'シレトコ野帳（エゾシカ）': 'SNF009036X',
        'シレトコ野帳（エゾクロテン）': 'SNF009037X',
        'シレトコ野帳（エゾリス）': 'SNF009050X',
        'シレトコ野帳（キタキツネ）': 'SNF009051X',
        '知床かばん（フウロ）': 'SNF013002X',
        '知床かばん（リス）': 'SNF013003X',
        '知床かばん（りす）': 'SNF013003X',
    }
    
    # Map partner items
    df_partner['商品コード'] = df_partner['商品名'].map(partner_name_map)
    df_partner_mapped = df_partner.dropna(subset=['商品コード'])
    
    print("=== Mapped Partner Stock ===")
    print(df_partner_mapped[['商品コード', '商品名', '在庫数']].to_string())
    
except Exception as e:
    print("Error:", e)
