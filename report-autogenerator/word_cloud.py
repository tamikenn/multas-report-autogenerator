import pandas as pd
import MeCab
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os

# 日本語フォントの設定
font_path = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
font_manager.fontManager.addfont(font_path)
plt.rcParams['font.family'] = 'IPAGothic'

def extract_actions(text, mecab):
    """テキストから動詞を抽出"""
    actions = []
    node = mecab.parseToNode(text)
    while node:
        # 品詞情報を分割
        features = node.feature.split(',')
        # 動詞の基本形を取得（動詞の原形）
        if features[0] == '動詞' and features[6] != '*':
            actions.append(features[6])  # 基本形を使用
        node = node.next
    return actions

def create_wordcloud(word_freq, title):
    """ワードクラウドの生成"""
    wordcloud = WordCloud(
        font_path=font_path,
        width=800,
        height=400,
        background_color='white',
        colormap='viridis',
        prefer_horizontal=0.7,
        min_font_size=12,
        max_font_size=80
    ).generate_from_frequencies(word_freq)
    
    plt.figure(figsize=(10, 6))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.title(title, fontsize=16, pad=20)
    
    return plt

def analyze_student_actions(excel_path):
    """学生ごとのテキスト分析とワードクラウド生成"""
    mecab = MeCab.Tagger()  # デフォルト設定を使用
    xls = pd.ExcelFile(excel_path)
    
    # 出力ディレクトリの準備
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    for sheet in xls.sheet_names:
        if sheet == 'overall':
            continue
            
        df = pd.read_excel(excel_path, sheet_name=sheet)
        
        # API分類ごとの分析
        for category in range(1, 13):
            category_texts = df[df['API検証'] == category]['入力内容']
            
            if len(category_texts) == 0:
                continue
            
            # テキストから動詞を抽出
            all_actions = []
            for text in category_texts:
                if isinstance(text, str):  # テキストが文字列の場合のみ処理
                    all_actions.extend(extract_actions(text, mecab))
            
            # 動詞の頻度をカウント
            action_freq = Counter(all_actions)
            
            if action_freq:  # 動詞が存在する場合のみワードクラウドを生成
                # ワードクラウドの生成と保存
                plt = create_wordcloud(
                    action_freq,
                    f'{sheet} - API分類{category}の行動パターン'
                )
                plt.savefig(
                    os.path.join(output_dir, f'{sheet}_category{category}_wordcloud.png'),
                    bbox_inches='tight',
                    dpi=300
                )
                plt.close()
                
                print(f'{sheet} - API分類{category}のワードクラウドを生成しました')

if __name__ == '__main__':
    # データファイルの処理
    data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
    for excel_file in [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]:
        file_path = os.path.join(data_dir, excel_file)
        analyze_student_actions(file_path)