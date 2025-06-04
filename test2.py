from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
import os

# โหลดฟอนต์
font_dir = os.path.join(os.path.dirname(__file__), 'Font')
pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(font_dir, 'THSarabunNew.ttf')))
pdfmetrics.registerFont(TTFont('THSarabun-Bold', os.path.join(font_dir, 'THSarabunNew Bold.ttf')))

# ตั้งค่า PDF
doc = SimpleDocTemplate("ตารางจำลอง.pdf", pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)

# สไตล์ข้อความ
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='ThaiNormal', fontName='THSarabun', fontSize=16, leading=22))
styles.add(ParagraphStyle(name='ThaiBold', fontName='THSarabun-Bold', fontSize=16, leading=22, alignment=1))

# ฟังก์ชันสร้าง Paragraph
def thai_paragraph(text, bold=False):
    style = styles['ThaiBold'] if bold else styles['ThaiNormal']
    return Paragraph(text.replace('\n', '<br/>'), style)

# ข้อมูลตาราง
table_data = [
    [
        thai_paragraph('ลำดับ', bold=True),
        thai_paragraph('รายการขอซื้อจ้าง\n[ขั้นตอนตามระเบียบฯข้อ 22(2)]', bold=True),
        thai_paragraph('รายการแยกวัสดุตามระบบบัญชี\n3 มิติ', bold=True),
        thai_paragraph('จำนวนหน่วย', bold=True),
        thai_paragraph('กำหนดเวลาที่ต้องการใช้พัสดุ\n[ตามระเบียบข้อ 22 (5)]', bold=True)
    ],
    [
        thai_paragraph('1'),
        thai_paragraph('โปสเตอร์ ขนาด 80×120 ซม.'),
        thai_paragraph('ว.โฆษณาและเผยแพร่'),
        thai_paragraph('1 อัน'),
        thai_paragraph('มิ.ย.68')
    ]
]

# กำหนดขนาดคอลัมน์และแถวให้พอดีแบบ Word
col_widths = [2.0*cm, 7.2*cm, 5.2*cm, 3.0*cm, 3.0*cm]
row_heights = [2.4*cm, 1.6*cm]

# สร้างตาราง
table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)

# สไตล์ตาราง
table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),
    ('FONTSIZE', (0, 0), (-1, -1), 16),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('GRID', (0, 0), (-1, -1), 0.7, colors.black),
]))

# สร้าง PDF
elements = [table]
doc.build(elements)
print("✅ ตารางสร้างเสร็จสมบูรณ์: ตารางจำลอง.pdf")
