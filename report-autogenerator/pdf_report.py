import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from datetime import datetime

# 日本語フォント設定
from reportlab.pdfbase.ttfonts import TTFont

# フォント定数
FONT_NAME = 'IPAGothic'
FONT_PATH = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'

# IPAフォントを登録
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

def create_pdf_report(df, student_name, output_path):
    """学生ごとのPDFレポートを生成"""
    # PDFドキュメントの設定
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=25*mm,
        leftMargin=25*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )
    
    # スタイル定数
    BLUE_COLOR = colors.HexColor('#2F5496')
    
    # 基本スタイルの設定
    def create_base_style(name, font_size, space_before=6, space_after=6, color=colors.black):
        return ParagraphStyle(
            name=name,
            fontName=FONT_NAME,
            fontSize=font_size,
            leading=font_size + 4,  # 行間は文字サイズ+4が標準的
            spaceBefore=space_before,
            spaceAfter=space_after,
            textColor=color
        )
    
    # スタイルの設定
    styles = getSampleStyleSheet()
    styles.add(create_base_style('Japanese', 10))
    styles.add(create_base_style('JapaneseTitle', 14, space_before=0, space_after=20))
    styles.add(create_base_style('CategoryHeader', 12, space_before=15, space_after=8, color=BLUE_COLOR))
    styles.add(create_base_style('DayHeader', 11, space_before=12, space_after=6, color=BLUE_COLOR))
    
    # ドキュメントの構築
    story = []
    
    # タイトル
    title = Paragraph(
        f"{student_name}の臨床実習記録まとめ",
        styles['JapaneseTitle']
    )
    story.append(title)
    
    # 日程一覧
    schedule_text = Paragraph(
        "実習日程:<br/>" +
        "Day1: 2025/7/28<br/>" +
        "Day2: 2025/7/29<br/>" +
        "Day3: 2025/7/30<br/>" +
        "Day4: 2025/7/31<br/>" +
        "Day5: 2025/8/1",
        styles['Japanese']
    )
    story.append(schedule_text)
    story.append(Spacer(1, 20))
    
    # API分類ごとの記録を処理
    for category in range(1, 13):
        category_entries = df[df['API検証'] == category]
        if len(category_entries) > 0:
            # カテゴリヘッダー（余白調整）
            story.append(Spacer(1, 5))  # カテゴリー前に少し余白を追加
            header = Paragraph(
                f"■ {CATEGORY_NAMES[category]}（API分類{category}）",
                styles['CategoryHeader']
            )
            story.append(header)
            
            # 記録を日付でグループ化
            def group_entries_by_day(entries):
                grouped = {}
                for _, entry in entries.iterrows():
                    day = entry['DAY']
                    content = str(entry['入力内容']).replace('\n', '<br/>')
                    grouped.setdefault(day, []).append(content)
                return grouped
            
            # テーブルデータの作成
            def create_table_data(grouped_data, styles):
                data = []
                for day in sorted(grouped_data.keys()):
                    # Day表記を追加
                    data.append([
                        Paragraph(f'Day {day}', styles['DayHeader']),
                        Paragraph('', styles['Japanese'])
                    ])
                    
                    # その日の記録を追加
                    for content in grouped_data[day]:
                        data.append([
                            Paragraph(content, styles['Japanese']),
                            Paragraph('', styles['Japanese'])
                        ])
                return data
            
            # テーブルデータの生成
            grouped_data = group_entries_by_day(category_entries)
            data = create_table_data(grouped_data, styles)
            
            if data:
                # テーブルスタイルの設定
                GREY_LINE = colors.Color(0.8, 0.8, 0.8)  # 薄いグレー
                CELL_PADDING = 3  # 基本的なパディング
                VERTICAL_PADDING = 8  # 上下のパディング
                
                table_style = TableStyle([
                    # フォントと配置
                    ('FONT', (0, 0), (-1, -1), FONT_NAME),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    
                    # 背景と文字色
                    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    
                    # パディング設定
                    ('LEFTPADDING', (0, 0), (-1, -1), CELL_PADDING),
                    ('RIGHTPADDING', (0, 0), (-1, -1), CELL_PADDING),
                    ('TOPPADDING', (0, 0), (-1, -1), VERTICAL_PADDING),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), VERTICAL_PADDING),
                    
                    # 区切り線（破線）
                    ('LINEBELOW', (0, 0), (-1, -1), 0.5, GREY_LINE, 1, (3, 2))
                ])
                
                # テーブルの作成（内容を1列目に配置し、フルワイドで表示）
                table = Table(data, colWidths=[doc.width * 0.95, doc.width * 0.05])
                table.setStyle(table_style)
                story.append(table)
            
            story.append(Spacer(1, 10))
    
    # PDFの生成
    doc.build(story)

def generate_reports():
    """全学生のレポートを生成"""
    # 出力ディレクトリの準備
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    # データファイルの処理
    data_dir = os.path.join(os.path.dirname(__file__), 'source_data')
    for excel_file in [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]:
        file_path = os.path.join(data_dir, excel_file)
        xls = pd.ExcelFile(file_path)
        
        for sheet in xls.sheet_names:
            if sheet == 'overall':
                continue
                
            print(f"{sheet}のレポートを生成中...")
            df = pd.read_excel(file_path, sheet_name=sheet)
            output_path = os.path.join(output_dir, f"{sheet}_report.pdf")
            create_pdf_report(df, sheet, output_path)
            print(f"レポートを保存しました: {output_path}")

if __name__ == '__main__':
    generate_reports()