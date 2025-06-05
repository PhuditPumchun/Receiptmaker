from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import Backend

Data = Backend.Data()
Data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "5 มิ.ย. 2568")
Data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "10 รีม", "6 มิ.ย. 2568")
Data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "7 มิ.ย. 2568")
Data.appendlist("โต๊ะทำงาน", "วัสดุสำนักงาน", "1 ตัว", "8 มิ.ย. 2568")
Data.appendlist("แฟ้มเอกสาร", "วัสดุสำนักงาน", "10 เล่ม", "9 มิ.ย. 2568")
Data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "5 รีม", "10 มิ.ย. 2568")

def set_font_thai(run):
    run.font.name = 'TH Sarabun New'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), 'TH Sarabun New')
    run.font.size = Pt(16)

def Sleeve1(Data):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # หัวเรื่อง
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("บันทึกข้อความ")
    set_font_thai(run)
    run.bold = True
    run.font.size = Pt(22)

    doc.add_paragraph().add_run("")  # spacer

    doc.add_paragraph("ส่วนราชการ ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ฯ ทรัพยากรธรรมชาติและสิ่งแวดล้อม โทร. 2749")
    doc.add_paragraph("ที่ อว 0603.07.04/1734" + " " * 50 + Data.day)
    doc.add_paragraph("เรื่อง ขออนุมัติจ้างพิมพ์โปสเตอร์")
    doc.add_paragraph("-------------------------------------------------------------------------------------------------------------------------------------------")
    doc.add_paragraph("เรียน คณบดีคณะเกษตรศาสตร์ฯ")

    body = (
        "\tด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจ้างพิมพ์โปสเตอร์ จำนวน 1 รายการ "
        "เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568 ในโครงการพัฒนาผลิตภัณฑ์ นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร "
        "และต้องการใช้สิ่งของดังกล่าว ประมาณ(เดือน/ปี) มิถุนายน 2568 และเบิกจ่ายจากเงินงบประมาณงบประมาณรายได้ 2568 กองทุนเพื่อการศึกษา "
        "แผนงานจัดการศึกษาอุดมศึกษา งานจัดการศึกษาสาขาเกษตรศาสตร์ หลักสูตรวิทยาศาสตรบัณฑิต สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดเงินอุดหนุน "
        "โครงการพัฒนากระบวนการจัดการเรียนการสอน (โครงการพัฒนาผลิตภัณฑ์นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร) หมวดค่าวัสดุโฆษณาและเผยแพร่ "
        "เป็นเงิน 1,000 บาท (หนึ่งพันบาทถ้วน)"
    )
    para = doc.add_paragraph()
    run = para.add_run(body)
    set_font_thai(run)
    run.font.size = Pt(15)

    # ตารางรายการ
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    headers = Data.list[0]
    for i in range(5):
        run = hdr_cells[i].paragraphs[0].add_run(headers[i])
        set_font_thai(run)
        run.bold = True

    for row in Data.list[1:]:
        cells = table.add_row().cells
        for i in range(5):
            run = cells[i].paragraphs[0].add_run(str(row[i]))
            set_font_thai(run)

    doc.add_paragraph("\n\nลงชื่อ ..........................................................")
    doc.add_paragraph("(รศ.ดร.ทิพวรรณ ทองสุข)")

    doc.save("Sleeve1.docx")
    print("✅ Sleeve1.docx created")

def summarySleeve(Data):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # หัวเรื่อง
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("รายงานการขอซื้อจ้าง")
    set_font_thai(run)
    run.bold = True
    run.font.size = Pt(26)

    doc.add_paragraph(Data.day)

    # ตารางรายการทั้งหมด
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ['ลำดับ', 'รายการ', 'หมวดหมู่', 'จำนวน', 'วันที่']
    for i in range(5):
        run = table.rows[0].cells[i].paragraphs[0].add_run(headers[i])
        set_font_thai(run)
        run.bold = True

    for idx, row in enumerate(Data.list[1:], 1):
        cells = table.add_row().cells
        values = [idx] + row[1:]
        for i in range(5):
            run = cells[i].paragraphs[0].add_run(str(values[i]))
            set_font_thai(run)

    doc.add_paragraph("\nสรุปรายการรวม")

    # ตารางสรุป
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    headers = ['ลำดับ', 'รายการ', 'หมวดหมู่', 'รวมจำนวน']
    for i in range(4):
        run = table.rows[0].cells[i].paragraphs[0].add_run(headers[i])
        set_font_thai(run)
        run.bold = True

    for idx, row in enumerate(Data.summary(), 1):
        cells = table.add_row().cells
        values = [idx] + row[1:]
        for i in range(4):
            run = cells[i].paragraphs[0].add_run(str(values[i]))
            set_font_thai(run)

    doc.save("summarySleeve.docx")
    print("✅ summarySleeve.docx created")

if __name__ == "__main__":
    Sleeve1(Data)
    summarySleeve(Data)