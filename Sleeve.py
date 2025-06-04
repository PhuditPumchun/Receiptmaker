from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

import Backend
import os

Datum = Backend.Data()
Datum.appendlist("a","b","c","d")
def Sleeve1(Data):
    # โหลดฟอนต์ภาษาไทย
    font_dir = os.path.join(os.path.dirname(__file__), 'Font')
    pdfmetrics.registerFont(TTFont('THSarabun', os.path.join(font_dir, 'THSarabunNew.ttf')))
    pdfmetrics.registerFont(TTFont('THSarabun-Bold', os.path.join(font_dir, 'THSarabunNew Bold.ttf')))

    # ตั้งค่าไฟล์ PDF
    file_name = "test.pdf"
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
    styles.add(ParagraphStyle(name='ThaiTitle', fontName='THSarabun-Bold', fontSize=26, leading=26, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='ThaiNormal', fontName='THSarabun', fontSize=16, leading=20))
    styles.add(ParagraphStyle(name='Line', fontName='THSarabun', fontSize=15, leading=20))
    styles.add(ParagraphStyle(name='ThaiBold', fontName='THSarabun-Bold', fontSize=16, leading=20))
    # Removed ThaiHeaderRight style as it's no longer used in this header setup

    style = styles['ThaiNormal']
    bold = styles['ThaiBold']
    elements = []

    # โลโก้และหัวเรื่อง
    logo_path = os.path.join(os.path.dirname(__file__), "logo.jpg")
    logo = Image(logo_path, width=2.0*cm, height=2.0*cm)
    title = Paragraph("บันทึกข้อความ", styles['ThaiTitle'])

    # Create the main header table to include the logo and "บันทึกข้อความ"
    header_table = Table(
        [[logo, title]], # Only two elements now: logo and title
        # Adjust colWidths to push logo left and center title
        # 2.5*cm for logo, (A4 width - 2*leftMargin - 2*rightMargin - 2.5*cm) for title
        # A4 width is 21 cm. Margins are 2cm each. So usable width is 21 - 4 = 17cm.
        # Remaining width for title: 17cm - 2.5cm = 14.5cm
    colWidths=[1.0*cm, 16.75*cm]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),   # Align logo to the left
        ('ALIGN', (1, 0), (1, 0), 'CENTER'), # Align 'บันทึกข้อความ' to center
        # Removed ('ALIGN', (2, 0), (2, 0), 'RIGHT') as there's no third column
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    #    ('GRID', (0, 0), (-1, -1), 0.7, colors.black), # Keep grid for visualization if needed
    ]))
    elements.append(header_table)


    # ข้อความส่วนราชการและวัน
    elements.append(Spacer(1, 0.3 * cm))
    elements.append(Paragraph('<b>ส่วนราชการ</b> ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ฯ ทรัพยากรธรรมชาติและสิ่งแวดล้อม โทร. 2749', style))
    elements.append(Paragraph("ที่ อว 0603.07.04/1734" + "&nbsp;"*60 + Data.day, style))
    elements.append(Spacer(1, 0.3 * cm))
    elements.append(Paragraph("เรื่อง ขออนุมัติจ้างพิมพ์โปสเตอร์", bold))
    elements.append(Paragraph("-------------------------------------------------------------------------------------------------------------------------------------------", styles['Line']))
    elements.append(Paragraph("เรียน คณบดีคณะเกษตรศาสตร์ฯ", bold))
    elements.append(Spacer(1, 0.4 * cm))

    # เนื้อหา
    body = (
        "        ด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจ้างพิมพ์โปสเตอร์ จำนวน 1 รายการ เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568 ในโครงการพัฒนาผลิตภัณฑ์ นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร และต้องการใช้สิ่งของดังกล่าว ประมาณ(เดือน/ปี) มิถุนายน 2568 และเบิกจ่ายจากเงินงบประมาณงบประมาณรายได้ 2568 กองทุนเพื่อการศึกษา แผนงานจัดการศึกษาอุดมศึกษา งานจัดการศึกษาสาขาเกษตรศาสตร์ หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดเงินอุดหนุน โครงการพัฒนากระบวนการจัดการเรียนการสอน (โครงการพัฒนาผลิตภัณฑ์นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร) หมวดค่าวัสดุโฆษณาและเผยแพร่ เป็นเงิน 1,000 บาท (หนึ่งพันบาทถ้วน)"
    )
    elements.append(Paragraph(body, style))
    elements.append(Spacer(1, 0.5 * cm))

    # ตารางข้อมูลจาก Backend
    table_data = Data.list
    table = Table(table_data, colWidths=[2*cm, 5*cm, 4*cm, 3*cm, 3*cm])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),
        ('FONTSIZE', (0, 0), (-1, -1), 16),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.7, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 0.19*cm),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.19*cm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.19*cm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.19*cm),
    ]))
    elements.append(table)

    # ลายเซ็น
    elements.append(Spacer(1, 1.5 * cm))
    elements.append(Paragraph("ลงชื่อ ..........................................................", style))
    elements.append(Paragraph("(รศ.ดร.ทิพวรรณ ทองสุข)", style))

    # สร้าง PDF
    doc.build(elements)
    print("✅ สร้างไฟล์ PDF เสร็จสมบูรณ์ →", file_name)

if __name__ == "__main__":
    Sleeve1(Datum)