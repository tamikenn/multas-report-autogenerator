import pandas as pd
import os
from collections import Counter
import matplotlib.pyplot as plt
import numpy as np
import glob
from matplotlib import font_manager

# 日本語フォントの設定
font_manager.fontManager.addfont('/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf')
plt.rcParams['font.family'] = 'IPAGothic'

# 定数定義
CHART_SIZE = (14, 14)         # グラフのサイズをさらに大きく
MARKER_STYLE = 'o-'          # マーカーと線のスタイル
LINE_WIDTH = 2               # 線の太さ
FILL_ALPHA = 0.25           # 塗りつぶしの透明度
MIN_RADIUS = 6              # 最小半径（データが少ない場合の見やすさ確保）
TICK_INTERVAL = 5           # 目盛りの間隔
BASE_VALUE = 1              # 基準値（0点に相当する値）
DPI = 300                   # 画像の解像度
LABEL_PADDING = 1.4         # ラベルの余白調整を増加

def prepare_plot_data(counts):
    """データを12時方向から時計回りに準備"""
    # [12, 1, 2, ..., 11] の順序で値を取得
    hours = [12] + list(range(1, 12))
    # 各値に基準値を加算
    values = [counts.get(h, 0) + BASE_VALUE for h in hours]
    # グラフを閉じるため最初の値を最後に追加
    values.append(values[0])
    
    # 分類名の定義
    category_names = {
        12: '社会医学',
        1: '医療倫理',
        2: '地域医療',
        3: '医学的知識',
        4: '診察・手技',
        5: '問題解決能力',
        6: '統合的臨床能力',
        7: '多職種連携',
        8: 'コミュニケーション',
        9: '一般教養',
        10: '保健・福祉',
        11: '行政'
    }
    
    return values, [category_names[h] for h in hours]

def setup_radar_chart():
    """レーダーチャートの基本設定"""
    fig, ax = plt.subplots(figsize=CHART_SIZE, subplot_kw={'polar': True})
    ax.set_theta_zero_location('N')    # 0°を北（12時）に
    ax.set_theta_direction(-1)         # 時計回り
    return fig, ax

def plot_data(ax, values, angles, title):
    """データのプロットと装飾"""
    # データのプロットと塗りつぶし
    ax.plot(angles, values, MARKER_STYLE, linewidth=LINE_WIDTH, label=title)
    ax.fill(angles, values, alpha=FILL_ALPHA)

def configure_axes(ax, labels, values):
    """軸と目盛りの設定"""
    # 角度軸の設定（30度ごとにラベル）
    angles = np.arange(0, 360, 30)
    ax.set_thetagrids(angles, labels)
    
    # ラベルのフォントサイズと位置の調整
    for label, angle in zip(ax.get_xticklabels(), angles):
        if angle < 180:
            label.set_verticalalignment('bottom')
        else:
            label.set_verticalalignment('top')
        label.set_fontsize(16)  # フォントサイズを2倍に
    
    # 半径軸の設定
    rmax = max(max(values), MIN_RADIUS) + 1
    ax.set_rlim(0, rmax * LABEL_PADDING)  # ラベルの余白を確保
    
    # 目盛りの設定
    rticks = list(range(BASE_VALUE, rmax, TICK_INTERVAL))
    ax.set_yticks(rticks)
    # 表示値を実際の値からBASE_VALUE分引く
    ax.set_yticklabels([str(int(x - BASE_VALUE)) for x in rticks])

def create_radar_chart(counts, sheet_name, out_dir):
    """
    API分類のレーダーチャートを生成
    
    Parameters:
        counts (Counter): 1-12の値を持つCounterオブジェクト
        sheet_name (str): シート名（タイトルとファイル名に使用）
        out_dir (str): 出力ディレクトリのパス
    """
    # データの準備
    values, labels = prepare_plot_data(counts)
    angles = np.linspace(0, 2*np.pi, len(values), endpoint=True)
    
    # グラフの基本設定
    fig, ax = setup_radar_chart()
    
    # データのプロット
    plot_data(ax, values, angles, sheet_name)
    
    # 軸と目盛りの設定
    configure_axes(ax, labels, values)
    
    # タイトルの設定（フォントサイズを大きく）
    ax.set_title(f'{sheet_name}のAPI分類レーダーチャート', fontsize=18, pad=20)
    
    # 画像の保存
    out_path = os.path.join(out_dir, f'{sheet_name}_radar.png')
    plt.savefig(out_path, bbox_inches='tight', dpi=DPI)
    plt.close()
    print(f'レーダーチャート画像を保存: {out_path}')

# メインの処理
if __name__ == '__main__':
    # 出力ディレクトリの準備
    out_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(out_dir, exist_ok=True)
    
    # 既存の画像ファイルを削除
    for f in glob.glob(os.path.join(out_dir, '*_radar.png')):
        os.remove(f)
    
    # Excelファイルの処理
    data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
    for excel_file in [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]:
        file_path = os.path.join(data_dir, excel_file)
        xls = pd.ExcelFile(file_path)
        
        # 各シートの処理
        for sheet in xls.sheet_names:
            if sheet == 'overall':
                continue
            
            # シートからデータ取得とカウント
            df = pd.read_excel(xls, sheet_name=sheet)
            api_counts = Counter(df['API検証'])
            print(f'{sheet} のAPI分類値カウント:', [api_counts.get(i, 0) for i in range(1, 13)])
            
            # レーダーチャート生成
            create_radar_chart(api_counts, sheet, out_dir)