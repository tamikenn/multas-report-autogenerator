import pandas as pd
import os

# source_dataフォルダ内のExcelファイル一覧を取得
data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]

for file in files:
    file_path = os.path.join(data_dir, file)
    print(f'--- {file} ---')
    # Excelファイルの全シート名を取得
    xls = pd.ExcelFile(file_path)
    print('シート一覧:', xls.sheet_names)
    # 1つ目のシートを読み込んで先頭5行を表示
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    print('先頭5行:')
    print(df.head())
    print('カラム情報:')
    print(df.dtypes)
    print('\n')
