from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# โหลดฟอนต์ THSarabunNew
font_dir = os.path.join(os.path.dirname(__file__), 'Font')
pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(font_dir, 'THSarabunNew.ttf')))
pdfmetrics.registerFont(TTFont('THSarabun-Bold', os.path.join(font_dir, 'THSarabunNew Bold.ttf')))

# กำหนดไฟล์ PDF
file_name = "หน้าแรกขออนุมัติ.pdf"
doc = SimpleDocTemplate(
    file_name,
    pagesize=A4,
    leftMargin=2*cm,
    rightMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2*cm
)

# สไตล์
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='ThaiNormal', fontName='THSarabun', fontSize=16, leading=20))
styles.add(ParagraphStyle(name='ThaiBold', fontName='THSarabun-Bold', fontSize=16, leading=20))
style = styles['ThaiNormal']
bold = styles['ThaiBold']

elements = []

# ส่วนหัวเอกสาร
elements.append(Paragraph("บันทึกข้อความ", bold))
elements.append(Spacer(1, 0.3 * cm))
elements.append(Paragraph("ส่วนราชการ ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ฯ โทร. 2749", style))
elements.append(Paragraph(
    "ที่ อว 0603.07.04/1734" + 
    "&nbsp;"*60 + 
    "วันที่ 29 พฤษภาคม 2568", style))
elements.append(Spacer(1, 0.3 * cm))
elements.append(Paragraph("เรื่อง ขออนุมัติจ้างพิมพ์โปสเตอร์", bold))
elements.append(Paragraph("เรียน คณบดีคณะเกษตรศาสตร์ฯ", bold))
elements.append(Spacer(1, 0.4 * cm))

# เนื้อหา
body = (
    "ด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจ้างพิมพ์โปสเตอร์ จำนวน 1 รายการ "
    "เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568 "
    "ในโครงการพัฒนาผลิตภัณฑ์ นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร "
    "โดยของบประมาณรายได้ 2568 กองทุนเพื่อการศึกษา หมวดค่าวัสดุโฆษณาและเผยแพร่ "
    "เป็นเงิน 1,000 บาท (หนึ่งพันบาทถ้วน) โดยมีผู้รับผิดชอบคือ รศ.ดร.ทิพวรรณ ทองสุข "
    "และแต่งตั้งกรรมการตรวจรับคือ ผศ.ดร.ปริตา ธนสุกาญจน์"
)
elements.append(Paragraph(body, style))
elements.append(Spacer(1, 0.5 * cm))

# ตาราง
table_data = [
    ['ลำดับ', 'รายการขอซื้อจ้าง\n[ขั้นตอนตามระเบียบข้อ 22(2)]', 'รายการแยกวัสดุตาม\nระบบบัญชี\n3มิติ', 'จำนวนหน่วย', 'กำหนดเวลาที่\nต้องการใช้พัสดุ\n[ตามระเบียบข้อ\n22 (5)]'],
    ['1', 'โปสเตอร์ ขนาด 80x120 ซม.', 'ว.โฆษณาและเผยแพร่', '1 อัน', 'มิ.ย. 68'],
]

table = Table(table_data, colWidths=[2*cm, 7*cm, 5*cm, 3*cm, 3*cm])
table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),
    ('FONTSIZE', (0, 0), (-1, -1), 16),
    ('BACKGROUND', (0, 0), (-1, 0), colors.white),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('GRID', (0, 0), (-1, -1), 0.7, colors.black),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
     ('LEFTPADDING', (0, 0), (-1, -1), 0.19*cm),
    ('RIGHTPADDING', (0, 0), (-1, -1), 0.19*cm),
]))
elements.append(table)

# ปิดท้าย
elements.append(Spacer(1, 1.5 * cm))
elements.append(Paragraph("ลงชื่อ ..........................................................", style))
elements.append(Paragraph("(รศ.ดร.ทิพวรรณ ทองสุข)", style))

# สร้าง PDF
doc.build(elements)
print("✅ สร้างไฟล์ PDF เสร็จสมบูรณ์ →", file_name)
