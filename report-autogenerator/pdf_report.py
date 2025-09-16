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

# 日本語フォントの設定
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
# IPAフォントを登録
pdfmetrics.registerFont(TTFont('IPAGothic', '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'))

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
    
    # スタイルの設定
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='Japanese',
        fontName='IPAGothic',
        fontSize=10,
        leading=14,  # 行間
        spaceBefore=6,  # 段落前のスペース
        spaceAfter=6   # 段落後のスペース
    ))
    styles.add(ParagraphStyle(
        name='JapaneseTitle',
        fontName='IPAGothic',
        fontSize=14,
        leading=16,
        spaceAfter=20
    ))
    styles.add(ParagraphStyle(
        name='CategoryHeader',
        fontName='IPAGothic',
        fontSize=12,
        leading=14,
        spaceBefore=15,  # カテゴリー前の余白を増やす
        spaceAfter=8,
        textColor=colors.HexColor('#2F5496')
    ))
    styles.add(ParagraphStyle(
        name='DayHeader',
        fontName='IPAGothic',
        fontSize=11,
        leading=14,
        spaceBefore=12,  # Day前の余白を増やす
        spaceAfter=6,
        textColor=colors.HexColor('#2F5496')
    ))
    
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
            # カテゴリヘッダー（次の内容と一緒に表示）
            story.append(Spacer(1, 5))  # カテゴリー前に少し余白を追加
            
            # 記録を日付でグループ化
            grouped_data = {}
            for _, entry in category_entries.iterrows():
                day = entry['DAY']
                content = str(entry['入力内容']).replace('\n', ' ')  # 改行を空白に置換
                if day not in grouped_data:
                    grouped_data[day] = []
                grouped_data[day].append(content)
            
            # 記録の一覧を表形式で表示
            data = []
            for day in sorted(grouped_data.keys()):
                first_content = True
                first_day = len(data) == 0  # このカテゴリーの最初のDayかどうか
                for content in grouped_data[day]:
                    if first_content:
                        # 最初の記録はDay表示と同じ行に配置
                        header_text = (f"■ {CATEGORY_NAMES[category]}（API分類{category}）  Day {day}  {content}"
                                     if first_day else f"Day {day}  {content}")
                        data.append([
                            Paragraph(header_text, styles['CategoryHeader'] if first_day else styles['DayHeader']),
                            Paragraph('', styles['Japanese'])
                        ])
                        first_content = False
                    else:
                        # 2つ目以降の記録は新しい行に配置
                        data.append([
                            Paragraph(content, styles['Japanese']),
                            Paragraph('', styles['Japanese'])
                        ])
            
            if data:
                # テーブルスタイルの設定
                grey_line = colors.Color(0.8, 0.8, 0.8)  # 薄いグレー
                table_style = TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'IPAGothic'),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),  # 上部の余白を増やす
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),  # 下部の余白を増やす
                    # Day表記の行のスタイル
                    ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#2F5496')),
                    # 各エントリー間に破線を追加（Day表記の行は除く）
                    ('LINEBELOW', (0, 0), (-1, -1), 0.5, grey_line, 1, (3, 2)),  # 破線の横罫線
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