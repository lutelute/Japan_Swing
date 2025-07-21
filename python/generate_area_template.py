#!/usr/bin/env python3
"""
generate_area_template.py
日本の10エリア電力系統パラメータのExcelテンプレート生成
"""

import pandas as pd
import os

def generate_template(filename='area_parameters_template.xlsx'):
    """
    エリアパラメータのExcelテンプレートを生成
    
    Args:
        filename (str): 出力するExcelファイル名
    """
    
    # エリア名 (北→南)
    areas = ['北海道', '東北', '東京', '北陸', '中部',
             '関西', '中国', '四国', '九州', '沖縄']
    
    # Master シート列定義
    columns = ['Area', 'Generator_Count', 'p_m', 'b', 'b_int',
              'epsilon', 'Connection_Coeff']
    
    # 初期パラメータ値を設定
    initial_data = []
    
    for area in areas:
        row_data = {
            'Area'            : area,
            'Generator_Count' : 20,                    # 発電機台数
            'p_m'             : 0.95,                  # 機械的入力パワー [p.u.]
            'b'               : 1.0,                   # 同期化力係数 [p.u.]
            'b_int'           : 100.0,                 # エリア内結合係数 [p.u.]
            'epsilon'         : 0.1,                   # エリア間結合強度 [p.u.]
            'Connection_Coeff': 0.0 if area == '沖縄' else 0.1  # 沖縄は本土から独立
        }
        initial_data.append(row_data)
    
    # DataFrameに変換
    master_df = pd.DataFrame(initial_data)
    
    # Excelファイルに書き出し
    try:
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # Masterシートを作成
            master_df.to_excel(writer, sheet_name='Master', index=False)
            
            # 各エリア用の個別シートを作成
            for _, row in master_df.iterrows():
                area_name = row['Area']
                
                # エリア用のDataFrameを作成（パラメータ詳細用）
                area_params_df = pd.DataFrame({
                    'Parameter': columns[1:],  # 'Area'を除く
                    'Value': [row[col] for col in columns[1:]],
                    'Description': [
                        '発電機台数',
                        '機械的入力パワー [p.u.]',
                        '同期化力係数 [p.u.]',
                        'エリア内結合係数 [p.u.]',
                        'エリア間結合強度 [p.u.]',
                        'エリア間接続係数 [p.u.]'
                    ]
                })
                
                area_params_df.to_excel(writer, sheet_name=area_name, index=False)
            
            # ワークブックとワークシートオブジェクトを取得してフォーマット設定
            workbook = writer.book
            
            # セルフォーマットを定義
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'border': 1,
                'valign': 'top'
            })
            
            # Masterシートのフォーマット
            master_worksheet = writer.sheets['Master']
            
            # ヘッダー行のフォーマット
            for col_num, column_name in enumerate(columns):
                master_worksheet.write(0, col_num, column_name, header_format)
            
            # データ行のフォーマット
            for row_num in range(1, len(master_df) + 1):
                for col_num in range(len(columns)):
                    master_worksheet.write(row_num, col_num, 
                                         master_df.iloc[row_num-1, col_num], 
                                         cell_format)
            
            # 列幅を自動調整
            master_worksheet.set_column('A:G', 15)
            
            # 各エリアシートのフォーマット
            for area in areas:
                if area in writer.sheets:
                    worksheet = writer.sheets[area]
                    worksheet.set_column('A:C', 25)
                    
                    # ヘッダー行のフォーマット
                    for col_num, header in enumerate(['Parameter', 'Value', 'Description']):
                        worksheet.write(0, col_num, header, header_format)
        
        print(f'✓ {filename} を生成しました')
        
        # ファイルサイズを表示
        file_size = os.path.getsize(filename)
        print(f'  ファイルサイズ: {file_size:,} bytes')
        
        # 生成されたシート情報を表示
        print(f'  生成シート数: {len(areas) + 1} (Master + 各エリア)')
        print('  シート一覧:')
        print('    - Master (全エリア概要)')
        for area in areas:
            print(f'    - {area} (詳細パラメータ)')
            
    except Exception as e:
        print(f'❌ Excelファイル生成エラー: {e}')
        return False
        
    return True

def validate_template(filename='area_parameters_template.xlsx'):
    """
    生成されたテンプレートの妥当性をチェック
    
    Args:
        filename (str): チェックするExcelファイル名
        
    Returns:
        bool: 妥当性チェック結果
    """
    
    if not os.path.exists(filename):
        print(f'❌ ファイル {filename} が存在しません')
        return False
    
    try:
        # Masterシートの読み込みテスト
        master_df = pd.read_excel(filename, sheet_name='Master')
        
        # 必要な列が存在するかチェック
        required_columns = ['Area', 'Generator_Count', 'p_m', 'b', 'b_int', 
                           'epsilon', 'Connection_Coeff']
        
        missing_columns = [col for col in required_columns if col not in master_df.columns]
        if missing_columns:
            print(f'❌ 必要な列が不足しています: {missing_columns}')
            return False
        
        # エリア数チェック
        if len(master_df) != 10:
            print(f'❌ エリア数が正しくありません (期待値: 10, 実際: {len(master_df)})')
            return False
        
        # 各エリアシートの存在チェック
        excel_file = pd.ExcelFile(filename)
        expected_sheets = ['Master'] + master_df['Area'].tolist()
        missing_sheets = [sheet for sheet in expected_sheets if sheet not in excel_file.sheet_names]
        
        if missing_sheets:
            print(f'❌ 不足しているシート: {missing_sheets}')
            return False
        
        print('✓ テンプレート妥当性チェック完了 - すべて正常です')
        return True
        
    except Exception as e:
        print(f'❌ テンプレート妥当性チェックエラー: {e}')
        return False

def main():
    """メイン関数 - スタンドアローン実行用"""
    print("=== 日本10エリア電力系統パラメータテンプレート生成 ===")
    
    filename = 'area_parameters_template.xlsx'
    
    # テンプレート生成
    if generate_template(filename):
        # 妥当性チェック
        validate_template(filename)
        print(f'\n📋 使用方法:')
        print(f'  1. {filename} をExcelで開く')
        print(f'  2. Masterシートで全体パラメータを調整')
        print(f'  3. 各エリアシートで詳細パラメータを調整')
        print(f'  4. simulate_area_network.py でシミュレーション実行')
    else:
        print('❌ テンプレート生成に失敗しました')

if __name__ == "__main__":
    main()