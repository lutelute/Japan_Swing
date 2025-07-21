import pandas as pd

# ❶ エリア名 (北→南)
areas = ['北海道', '東北', '東京', '北陸', '中部',
         '関西', '中国', '四国', '九州', '沖縄']

# ❷ Master シート列
cols = ['Area', 'Generator_Count', 'p_m', 'b', 'b_int',
        'epsilon', 'Connection_Coeff']

# ❸ 初期値を作成
rows = [
    {
        'Area'            : area,
        'Generator_Count' : 20,
        'p_m'             : 0.95,
        'b'               : 1,
        'b_int'           : 100,
        'epsilon'         : 0.1,
        'Connection_Coeff': 0.0 if area == '沖縄' else 0.1
    }
    for area in areas
]

master_df = pd.DataFrame(rows)

# ❹ Excel 書き出し
with pd.ExcelWriter('area_parameters_template.xlsx',
                    engine='xlsxwriter') as writer:
    master_df.to_excel(writer, sheet_name='Master', index=False)

    # 各エリアシートを追加
    for _, row in master_df.iterrows():
        df = pd.DataFrame({
            'Parameter': cols[1:],               # 行見出し
            'Value'   : [row[c] for c in cols[1:]]
        })
        df.to_excel(writer, sheet_name=row['Area'], index=False)

print('area_parameters_template.xlsx を生成しました')