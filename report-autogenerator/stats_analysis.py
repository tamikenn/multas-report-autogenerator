import pandas as pd
import os
from collections import Counter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image

# フォント設定
FONT_NAME = 'IPAGothic'
FONT_PATH = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))

# API分類の名称定義
CATEGORY_NAMES = {
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
    11: '行政',
    12: '社会医学'
}

def create_stats_report(df, student_name, output_path):
    """統計レポートを生成"""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=25*mm,
        leftMargin=25*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )
    
    # スタイル設定
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='StatsTitle',
        fontName=FONT_NAME,
        fontSize=14,
        leading=16,
        spaceAfter=10,
        alignment=1  # 中央揃え
    ))
    styles.add(ParagraphStyle(
        name='StatsHeading',
        fontName=FONT_NAME,
        fontSize=12,
        leading=14,
        spaceBefore=8,
        spaceAfter=4
    ))
    
    # レーダーチャート画像のパス
    radar_path = os.path.join(os.path.dirname(output_path), f"{os.path.splitext(os.path.basename(output_path))[0].replace('_stats', '')}_radar.png")
    
    story = []
    
    # 1ページ目: レーダーチャートと基本統計
    title = Paragraph(f"{student_name}の統計データ", styles['StatsTitle'])
    story.append(title)
    
    # レーダーチャートの読み込みとリサイズ
    if os.path.exists(radar_path):
        radar_img = Image(radar_path)
        radar_width = 300
        aspect_ratio = radar_img.imageHeight / radar_img.imageWidth
        radar_img.drawWidth = radar_width
        radar_img.drawHeight = radar_width * aspect_ratio
        story.append(radar_img)
    story.append(Spacer(1, 10))
    
    # ①日別投稿数の集計
    story.append(Paragraph('①日別投稿数の集計', styles['StatsHeading']))
    daily_counts = df.groupby('DAY').size()
    total_posts = len(df)
    
    daily_data = [['日付', '投稿数']]
    for day in range(1, 6):
        count = daily_counts.get(day, 0)
        daily_data.append([f'Day {day}', str(count)])
    daily_data.append(['合計', str(total_posts)])
    
    daily_table = Table(daily_data, colWidths=[80, 80])
    daily_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), FONT_NAME),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ]))
    story.append(daily_table)
    story.append(Spacer(1, 15))
    
    # ②分類別記録数のランキング
    story.append(Paragraph('②分類別記録数のランキング', styles['StatsHeading']))
    category_counts = df['API検証'].value_counts().reindex(range(1, 13), fill_value=0)
    ranking_data = [['順位', '分類', '記録数', '割合']]
    
    for rank, (category, count) in enumerate(category_counts.items(), 1):
        percentage = (count / total_posts * 100) if total_posts > 0 else 0
        ranking_data.append([
            str(rank),
            f'{CATEGORY_NAMES[category]}',
            str(count),
            f'{percentage:.1f}%'
        ])

    ranking_table = Table(ranking_data, colWidths=[40, 280, 60, 60])
    ranking_table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), FONT_NAME),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ]))

    # レーダーチャートと統計情報を配置
    if os.path.exists(radar_path):
        # レーダーチャートと日別投稿数を横に配置
        top_row_data = [[radar_img, daily_table]]
        top_row = Table(top_row_data, colWidths=[300, 200])
        top_row.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(top_row)
    else:
        story.append(daily_table)
    
    story.append(Spacer(1, 20))
    story.append(ranking_table)
    story.append(PageBreak())
    
    category_counts = df['API検証'].value_counts().reindex(range(1, 13), fill_value=0)
    ranking_data = [['順位', '分類', '記録数', '割合']]
    
    for rank, (category, count) in enumerate(category_counts.items(), 1):
        percentage = (count / total_posts * 100) if total_posts > 0 else 0
        ranking_data.append([
            str(rank),
            f'{CATEGORY_NAMES[category]}',
            str(count),
            f'{percentage:.1f}%'
        ])

    t = Table(ranking_data, colWidths=[30, 200, 50, 50])
    t.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), FONT_NAME),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ]))
    story.append(t)
    story.append(Spacer(1, 15))

    # ③日別・分類別記録数の集計表
    story.append(Paragraph('③日別・分類別記録数', styles['StatsHeading']))
    story.append(Spacer(1, 5))
    
    # ピボットテーブルを作成
    pivot_data = pd.pivot_table(
        df,
        values='入力内容',
        index='API検証',
        columns='DAY',
        aggfunc='count',
        fill_value=0
    ).reindex(range(1, 13), fill_value=0)
    
    # テーブルデータの作成
    matrix_data = [['分類 \\ Day', 'Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5']]
    for category in range(1, 13):
        row = [CATEGORY_NAMES[category]]
        for day in range(1, 6):
            count = pivot_data.get(day, pd.Series())[category] if day in pivot_data else 0
            row.append(str(count))
        matrix_data.append(row)

    table_matrix = Table(matrix_data, colWidths=[250] + [50]*5)
    table_matrix.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), FONT_NAME),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTSIZE', (0, 0), (-1, -1), 9),  # フォントサイズを小さくして1ページに収まるように
    ]))
    story.append(table_matrix)

    # PDFの生成
    doc.build(story)

def generate_stats():
    """全学生の統計レポートを生成"""
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
    for excel_file in [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]:
        file_path = os.path.join(data_dir, excel_file)
        xls = pd.ExcelFile(file_path)
        
        for sheet in xls.sheet_names:
            if sheet == 'overall':
                continue
            
            print(f"{sheet}の統計レポートを生成中...")
            df = pd.read_excel(file_path, sheet_name=sheet)
            output_path = os.path.join(output_dir, f"{sheet}_stats.pdf")
            create_stats_report(df, sheet, output_path)
            print(f"統計レポートを保存しました: {output_path}")

if __name__ == '__main__':
    generate_stats()