import pandas as pd
import os
from collections import Counter
import matplotlib.pyplot as plt
import numpy as np
import glob

def create_radar_chart(counts, sheet_name, out_dir):
    """
    12時方向から時計回りにレーダーチャートを生成
    counts: Counter object with values for 1-12
    """
    # データを12時から時計回りに配置 [12,1,2,...,11]
    hour_order = [12] + list(range(1, 12))
    values = [counts.get(h, 0) for h in hour_order]
    
    # グラフを閉じるため最初の値を最後にも追加
    values.append(values[0])
    
    # 角度の設定（12時=0°から時計回り）
    angles = np.linspace(0, 2*np.pi, len(values), endpoint=True)
    
    # ラベルは12個（時計の文字盤）
    labels = [f'{h}時' for h in hour_order]
    
    # プロット
    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw={'polar': True})
    ax.set_theta_zero_location('N')    # 0°を北（12時）に
    ax.set_theta_direction(-1)         # 時計回り
    
    # データのプロットと塗りつぶし
    ax.plot(angles, values, 'o-', linewidth=2, label=sheet_name)
    ax.fill(angles, values, alpha=0.25)
    
    # 軸の設定
    ax.set_thetagrids(np.arange(0, 360, 30), labels)
    rmax = max(max(values), 10) + 1
    ax.set_rlim(5, rmax)
    rticks = list(range(5, rmax, 5))
    ax.set_yticks(rticks)
    
    # タイトル
    ax.set_title(f'{sheet_name}のAPI分類レーダーチャート')
    
    # 保存
    out_path = os.path.join(out_dir, f'{sheet_name}_radar.png')
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print(f'レーダーチャート画像を保存: {out_path}')

# 出力ディレクトリの準備
out_dir = os.path.join(os.path.dirname(__file__), 'output')
os.makedirs(out_dir, exist_ok=True)
for f in glob.glob(os.path.join(out_dir, '*_radar.png')):
    os.remove(f)

# データ取得と処理
data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
for excel_file in [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]:
    file_path = os.path.join(data_dir, excel_file)
    xls = pd.ExcelFile(file_path)
    
    for sheet in xls.sheet_names:
        if sheet == 'overall':
            continue
            
        # シートからデータ取得とカウント
        df = pd.read_excel(xls, sheet_name=sheet)
        api_counts = Counter(df['API検証'])
        print(f'{sheet} のAPI分類値カウント:', [api_counts.get(i, 0) for i in range(1, 13)])
        
        # レーダーチャート生成
        create_radar_chart(api_counts, sheet, out_dir)