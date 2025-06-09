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

Data = Backend.Data()
Data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "5 มิ.ย. 2568")
Data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "10 รีม", "6 มิ.ย. 2568")
Data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "7 มิ.ย. 2568")
Data.appendlist("โต๊ะทำงาน", "วัสดุสำนักงาน", "1 ตัว", "8 มิ.ย. 2568")
Data.appendlist("แฟ้มเอกสาร", "วัสดุสำนักงาน", "10 เล่ม", "9 มิ.ย. 2568")
Data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "5 รีม", "10 มิ.ย. 2568")

font_dir = os.path.join(os.path.dirname(__file__), 'Font')
# Ensure the Font directory exists for these paths to be valid
if not os.path.exists(font_dir):
    os.makedirs(font_dir) # Create the directory if it doesn't exist

# Create dummy font files if they don't exist, for demonstration purposes
# In a real scenario, you would place your actual font files here.
dummy_font_path_regular = os.path.join(font_dir, 'THSarabunNew.ttf')
dummy_font_path_bold = os.path.join(font_dir, 'THSarabunNew Bold.ttf')

if not os.path.exists(dummy_font_path_regular):
    with open(dummy_font_path_regular, 'w') as f:
        f.write('') 
if not os.path.exists(dummy_font_path_bold):
    with open(dummy_font_path_bold, 'w') as f:
        f.write('')

pdfmetrics.registerFont(TTFont('THSarabun', dummy_font_path_regular))
pdfmetrics.registerFont(TTFont('THSarabun-Bold', dummy_font_path_bold))

def Sleeve1(Data,bodyText):
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
    
    # สไตล์ใหม่สำหรับข้อความในตาราง
    styles.add(ParagraphStyle(name='TableCell', fontName='THSarabun', fontSize=14, leading=18, alignment=TA_CENTER)) # Increased leading
    styles.add(ParagraphStyle(name='TableHeader', fontName='THSarabun-Bold', fontSize=14, leading=18, alignment=TA_CENTER)) # Increased leading

    style = styles['ThaiNormal']
    bold = styles['ThaiBold']
    elements = []

    # โลโก้และหัวเรื่อง
    logo_path = os.path.join(os.path.dirname(__file__), "logo.jpg")
    if not os.path.exists(logo_path):
        from PIL import Image as PILImage
        dummy_img = PILImage.new('RGB', (100, 100), color = 'red')
        dummy_img.save(logo_path)

    logo = Image(logo_path, width=2.0*cm, height=2.0*cm)
    title = Paragraph("บันทึกข้อความ", styles['ThaiTitle'])

    # Create the main header table to include the logo and "บันทึกข้อความ"
    header_table = Table(
        [[logo, title]], # Only two elements now: logo and title
    colWidths=[1.0*cm, 16.75*cm]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),   # Align logo to the left
        ('ALIGN', (1, 0), (1, 0), 'CENTER'), # Align 'บันทึกข้อความ' to center
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
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
       bodyText
    )
    elements.append(Paragraph(body, style))
    elements.append(Spacer(1, 0.5 * cm))

    # ตารางข้อมูลจาก Backend
    table_data_formatted = []
    # Header row
    header_row = []
    for h_text in Data.list[0]:
        header_row.append(Paragraph(h_text, styles['TableHeader']))
    table_data_formatted.append(header_row)

    # Data rows
    for row in Data.list[1:]:
        formatted_row = []
        for cell_text in row:
            formatted_row.append(Paragraph(str(cell_text), styles['TableCell']))
        table_data_formatted.append(formatted_row)

    # กำหนด colWidths ให้ตารางหลัก
    main_table_col_widths = [2*cm, 5*cm, 4*cm, 3*cm, 3*cm]
    table = Table(table_data_formatted, colWidths=main_table_col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.85, 0.85, 0.85)), # Light grey background for header
        ('GRID', (0, 0), (-1, -1), 0.7, colors.black),
        
        ('TOPPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.1*cm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.1*cm),
        
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table)

    # ลายเซ็น
    elements.append(Spacer(1, 1.5 * cm))
    elements.append(Paragraph("ลงชื่อ ..........................................................", style))
    elements.append(Paragraph("(รศ.ดร.ทิพวรรณ ทองสุข)", style))

    # สร้าง PDF
    doc.build(elements)
    print("✅ สร้างไฟล์ PDF เสร็จสมบูรณ์ →", file_name)
    return 1

def summarySleeve(data):
    file_name = "summary.pdf"
    doc = SimpleDocTemplate(file_name, pagesize=A4)
    elements = []
    
    # ฟอนต์ไทย
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='ThaiTitle', fontName='THSarabun-Bold', fontSize=26, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='ThaiNormal', fontName='THSarabun', fontSize=16))
    styles.add(ParagraphStyle(name='ThaiBold', fontName='THSarabun-Bold', fontSize=16))
    
    # สไตล์ใหม่สำหรับข้อความในตารางสรุป
    styles.add(ParagraphStyle(name='SummaryTableCell', fontName='THSarabun', fontSize=16, leading=20, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='SummaryTableHeader', fontName='THSarabun-Bold', fontSize=16, leading=20, alignment=TA_CENTER))
    
    # หัวเรื่อง
    elements.append(Paragraph("รายงานการขอซื้อจ้าง", styles['ThaiTitle']))
    elements.append(Paragraph(data.day, styles['ThaiNormal']))
    elements.append(Spacer(1, 0.5 * cm))

    all_items_table_col_widths = [2*cm, 5*cm, 4*cm, 3*cm, 3*cm] 

    table_data = [['ลำดับ', 'รายการ', 'หมวดหมู่', 'จำนวน', 'วันที่']]
    # แปลงข้อมูลใน Data.list ให้เป็น Paragraph objects เพื่อควบคุมสไตล์ได้
    summary_table_data_formatted = []
    # Header row
    summary_header_row = []
    for h_text in table_data[0]:
        summary_header_row.append(Paragraph(h_text, styles['SummaryTableHeader']))
    summary_table_data_formatted.append(summary_header_row)

    # Data rows
    for i, row in enumerate(data.list[1:], start=1):
        formatted_row = []
        # Add sequence number
        formatted_row.append(Paragraph(str(i), styles['SummaryTableCell']))
        # Add other data from row[1:]
        for cell_text in row[1:]:
            formatted_row.append(Paragraph(str(cell_text), styles['SummaryTableCell']))
        summary_table_data_formatted.append(formatted_row)


    table = Table(summary_table_data_formatted, repeatRows=1, colWidths=all_items_table_col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        
        ('TOPPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.1*cm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.1*cm),
        
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 1 * cm))

    # ตารางสรุป
    elements.append(Paragraph("สรุปรายการรวม", styles['ThaiBold']))
    elements.append(Spacer(1, 1 * cm))
    
    # ตารางสรุปมี 4 คอลัมน์, ต้องปรับ colWidths ให้รวมกันได้ 17 cm (ความกว้างรวมของตารางแรก)
    # เช่น [2*cm, 7*cm, 5*cm, 3*cm] รวม 17cm - อันนี้เท่ากับของเดิมพอดี
    # หรือปรับให้เหมาะสมกับเนื้อหา
    summary_table_col_widths = [2*cm, 7*cm, 5*cm, 3*cm] # รวม 17cm

    summary_data_rows = [['ลำดับ', 'รายการ', 'หมวดหมู่', 'รวมจำนวน']]
    final_summary_table_data_formatted = []
    # Header row for final summary
    final_summary_header_row = []
    for h_text in summary_data_rows[0]:
        final_summary_header_row.append(Paragraph(h_text, styles['SummaryTableHeader']))
    final_summary_table_data_formatted.append(final_summary_header_row)

    for i, row in enumerate(data.summary(), start=1):
        formatted_row = []
        formatted_row.append(Paragraph(str(i), styles['SummaryTableCell']))
        for cell_text in row[1:]:
            formatted_row.append(Paragraph(str(cell_text), styles['SummaryTableCell']))
        final_summary_table_data_formatted.append(formatted_row)

    summary_table = Table(final_summary_table_data_formatted, repeatRows=1, colWidths=summary_table_col_widths)
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        
        ('TOPPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.15*cm),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.1*cm),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.1*cm),
        
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(summary_table)

    doc.build(elements)
    print("✅ สร้างไฟล์ PDF สรุปเสร็จสมบูรณ์ →", file_name)
    return 1


if __name__ == "__main__":
    Sleeve1(Data, "        ด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจ้างพิมพ์โปสเตอร์ จำนวน 1 รายการ เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568 ในโครงการพัฒนาผลิตภัณฑ์ นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร และต้องการใช้สิ่งของดังกล่าว ประมาณ(เดือน/ปี) มิถุนายน 2568 และเบิกจ่ายจากเงินงบประมาณงบประมาณรายได้ 2568 กองทุนเพื่อการศึกษา แผนงานจัดการศึกษาอุดมศึกษา งานจัดการศึกษาสาขาเกษตรศาสตร์ หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดเงินอุดหนุน โครงการพัฒนากระบวนการจัดการเรียนการสอน (โครงการพัฒนาผลิตภัณฑ์นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร) หมวดค่าวัสดุโฆษณาและเผยแพร่ เป็นเงิน 1,000 บาท (หนึ่งพันบาทถ้วน)")
    summarySleeve(Data)