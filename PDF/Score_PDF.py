import os
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.fonts import addMapping

def generate_pdf_and_open_with_edge():
    # Load a Vietnamese font
    pdfmetrics.registerFont(TTFont('DejaVuSans', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\dejavu-sans\\ttf\\DejaVuSans.ttf'))
    addMapping('DejaVuSans', 0, 0, 'DejaVuSans')  # Regular

    # Create a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
        file_path = temp_file.name

    # Create a PDF file in the temporary location
    pdf = SimpleDocTemplate(file_path, pagesize=A4)

    # Styles
    styles = getSampleStyleSheet()
    styleN = ParagraphStyle(name='Normal', fontName='DejaVuSans', fontSize=12)
    styleH = ParagraphStyle(name='Heading', fontName='DejaVuSans', fontSize=16, alignment=TA_CENTER, textColor=colors.darkblue)

    # Header Table
    header_data = [
        ['LOGO', '', Paragraph('CBS TRUNG TÂM NGOẠI NGỮ CON ONG', styleH), '', '18 ĐƯỜNG SỐ 2, P5, Q8 TPHCM'],
        ['', '', Paragraph('THÔNG MINH', styleH), '', '077 900 7218 - 0938 893 479'],
        ['', '', Paragraph('KẾT QUẢ KIỂM TRA CUỐI KHÓA', styleH), '', ''],
    ]

    header_table = Table(header_data, colWidths=[4 * cm, 1 * cm, 8 * cm, 1 * cm, 5 * cm])
    header_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (2, 0), (2, 1)),
        ('SPAN', (4, 0), (4, 1)),
        ('SPAN', (2, 2), (4, 2)),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (2, 2), (2, 2), colors.lightblue),
        ('TEXTCOLOR', (2, 2), (2, 2), colors.darkred),
        ('FONTSIZE', (2, 2), (2, 2), 16),
        ('FONTNAME', (2, 2), (2, 2), 'DejaVuSans'),
        ('TOPPADDING', (2, 2), (2, 2), 10),
        ('BOTTOMPADDING', (2, 2), (2, 2), 10),
    ]))

    # Information Table
    info_data = [
        ['Cấp độ thi', '', 'Địa điểm thi', '', 'Ngày thi', ''],
        ['Giai đoạn', '', 'Bài thi', '', 'Giờ thi', ''],
        ['Ngày học', '', 'GV chính', '', 'GV gác thi', ''],
        ['Giờ học', '', '', '', '', ''],
    ]

    info_table = Table(info_data, colWidths=[3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm])
    info_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
    ]))

    # Section 1 Table
    section1_data = [
        ['Họ và tên HS', '', 'Ngày sinh', '', 'Lớp', ''],
    ]

    section1_table = Table(section1_data, colWidths=[3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm])
    section1_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    # Scores Table Template for Each Stage
    def create_scores_table(stage):
        data = [
            [f'Giai đoạn {stage}:'],
            ['Điểm nghe', '', '', 'Điểm đọc viết', '', '', 'Điểm nói', '', ''],
            ['Điểm tổng cộng', '', '', '', '', '', 'Phần trăm điểm đạt:', ''],
            ['Nhận xét:', '', '', '', '', '', '', '', '']
        ]

        table = Table(data, colWidths=[3 * cm, 1 * cm, 1 * cm, 3 * cm, 1 * cm, 1 * cm, 3 * cm, 1 * cm, 1 * cm])
        table.setStyle(TableStyle([
            ('GRID', (0, 1), (-1, -1), 1, colors.black),
            ('SPAN', (0, 0), (-1, 0)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('BACKGROUND', (0, 1), (-1, 1), colors.whitesmoke),
            ('BACKGROUND', (0, 2), (-1, 2), colors.lightgrey),
        ]))
        return table

    # Combine All Elements
    elements = [header_table, Spacer(1, 12), info_table, Spacer(1, 12), section1_table]

    # Add all stages
    for i in range(1, 6):
        elements.append(Spacer(1, 12))
        elements.append(create_scores_table(i))

    # Footer Table
    footer_data = [
        ['Họ và tên HS', '', 'Ngày sinh', '', 'Lớp', ''],
    ]

    footer_table = Table(footer_data, colWidths=[3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm, 3 * cm])
    footer_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    elements.append(Spacer(1, 12))
    elements.append(footer_table)

    # Build the PDF
    pdf.build(elements)

    # Open the PDF with Microsoft Edge
    os.system(f'start msedge "{file_path}"')

# Example usage
generate_pdf_and_open_with_edge()
