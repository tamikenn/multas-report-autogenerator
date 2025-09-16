import pandas as pd
import os
from collections import Counter

import matplotlib.pyplot as plt
import numpy as np

# source_dataフォルダ内のExcelファイル一覧を取得
data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]

for file in files:
    file_path = os.path.join(data_dir, file)
    xls = pd.ExcelFile(file_path)
    # overall以外の各ユーザーシートを処理
    for sheet in xls.sheet_names:
        if sheet == 'overall':
            continue
        df = pd.read_excel(xls, sheet_name=sheet)
        # API検証カラムから1-12の値をカウント
        api_col = df['API検証']
        counts = Counter(api_col)
        radar_counts = [counts.get(i, 0) for i in range(1, 13)]
        print(f'{sheet} のAPI分類値カウント: {radar_counts}')

        # レーダーチャート作成（12が0時、1-11が時計回りに並ぶように修正）
        labels = [f'{i}時' for i in [12] + list(range(1, 12+1))[:-1]]  # [12,1,2,...,11]
        # データも同じ順序に並べ替え
        values = [radar_counts[11]] + radar_counts[:11]  # [12,1,2,...,11]
        values += [values[0]]  # 閉じるため先頭を最後に
        angles = np.linspace(0, 2 * np.pi, 13)

        fig, ax = plt.subplots(subplot_kw={'polar': True})
        ax.set_theta_zero_location('N')  # 0度を上（北）に
        ax.set_theta_direction(-1)       # 時計回り
        ax.plot(angles, values, 'o-', linewidth=2)
        ax.fill(angles, values, alpha=0.25)
        ax.set_thetagrids(np.arange(0, 360, 30), labels)
        ax.set_title(f'{sheet}のAPI分類レーダーチャート')
        ax.set_rlim(0, max(values)+1)

        # 保存
        out_dir = os.path.join(os.path.dirname(__file__), 'output')
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f'{sheet}_radar.png')
        plt.savefig(out_path)
        plt.close()
        print(f'レーダーチャート画像を保存: {out_path}')
