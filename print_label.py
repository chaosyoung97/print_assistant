import pandas as pd
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate
import tkinter as tk
from tkinter import filedialog
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.lib.styles import ParagraphStyle
import datetime

# 固定仓库名
WAREHOUSE_NAME = "河北顾家家居仓"

# 注册Noto Sans SC字体（需将NotoSansSC-VariableFont_wght.ttf放在脚本目录）
pdfmetrics.registerFont(TTFont('NotoSansSC', 'NotoSansSC-VariableFont_wght.ttf'))
addMapping('NotoSansSC', 0, 0, 'NotoSansSC')

def excel_date_to_str(excel_date):
    # Excel起始日期为1899-12-30
    try:
        date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(excel_date))
        return date.strftime('%Y-%m-%d')
    except Exception:
        return str(excel_date)

# 生成面单PDF函数
def generate_label(data, output_path):
    width, height = 76 * mm, 130 * mm
    left_margin = right_margin = 1 * mm
    top_margin = bottom_margin = 1 * mm
    table_width = width - left_margin - right_margin
    # 列宽分配，尽量填充整个面单
    col_widths = [22*mm, 22*mm, 14*mm, 16*mm]
    row_heights = [16*mm, 10*mm, 10*mm, 30*mm, 20*mm, 10*mm]

    # 格式化日期
    date_str = str(data['订单日期'])
    try:
        # 优先尝试Excel数字日期
        if date_str.isdigit() or (date_str.replace('.', '', 1).isdigit() and '.' in date_str):
            date_fmt = excel_date_to_str(date_str)
        else:
            date_obj = pd.to_datetime(date_str)
            date_fmt = date_obj.strftime('%Y-%m-%d')
    except Exception:
        date_fmt = date_str

    # 样式
    style = ParagraphStyle(
        name='NotoSansSC', fontName='NotoSansSC', fontSize=10, leading=12, alignment=1,
    )
    style_left = ParagraphStyle(
        name='NotoSansSCLeft', fontName='NotoSansSC', fontSize=10, leading=12, alignment=0,
    )
    style_bold = ParagraphStyle(
        name='NotoSansSCBold', fontName='NotoSansSC', fontSize=13, leading=15, alignment=1,
    )

    # 表格数据，第一行为标题，合并4列
    table_data = [
        [Paragraph(WAREHOUSE_NAME, ParagraphStyle(name='Title', fontName='NotoSansSC', fontSize=16, alignment=1, leading=20)),'','',''],
        [Paragraph("日期", style), Paragraph(date_fmt, style), Paragraph("件数", style), Paragraph(str(data['包装件数']), style)],
        [Paragraph("装载号", style), Paragraph(str(data['装载号']), style_left), '', ''],
        [Paragraph("收货信息", style), Paragraph(str(data['买家姓名']), style_left), Paragraph(str(data['地址']), style_left), ''],
        [Paragraph("品名", style), Paragraph(str(data['品名']), style_left), '', ''],
        [Paragraph("顾客", style), Paragraph(str(data['顾客名字']), style_left), '', '']
    ]

    doc = SimpleDocTemplate(output_path, pagesize=(width, height), leftMargin=left_margin, rightMargin=right_margin, topMargin=top_margin, bottomMargin=bottom_margin)
    elements = []
    # 表格
    table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
    table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'NotoSansSC'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),  # 标题行居中
        ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
        ('SPAN', (0,0), (3,0)),  # 标题合并4列
        ('FONTSIZE', (0,0), (0,0), 16),      # 标题字体大
        ('FONTNAME', (0,0), (0,0), 'NotoSansSC'),
        ('FONTNAME', (0,1), (-1,-1), 'NotoSansSC'),
        ('ALIGN', (0,1), (0,-1), 'CENTER'),
        ('ALIGN', (1,1), (-1,-1), 'LEFT'),
        ('ALIGN', (2,1), (2,1), 'CENTER'),
        ('ALIGN', (3,1), (3,1), 'CENTER'),
        ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,1), (-1,1), colors.whitesmoke),
        ('SPAN', (1,2), (3,2)),
        # 收货信息行：地址合并为两格（第3、4格）
        ('SPAN', (2,3), (3,3)),
        ('SPAN', (1,4), (3,4)),
        ('SPAN', (1,5), (3,5)),
    ]))
    elements.append(table)
    doc.build(elements)

# 主程序
def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel Files', '*.xlsx')])
    if not file_path:
        print('未选择文件')
        return
    df = pd.read_excel(file_path, engine='openpyxl')
    if df.empty:
        print('Excel内容为空')
        return
    data = df.iloc[0]
    output_path = filedialog.asksaveasfilename(title='保存PDF', defaultextension='.pdf', filetypes=[('PDF Files', '*.pdf')])
    if not output_path:
        print('未选择保存路径')
        return
    generate_label(data, output_path)
    print(f'面单已生成：{output_path}')

if __name__ == '__main__':
    main() 