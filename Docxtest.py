from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches
import os
import time
import platform # เพิ่ม import นี้เข้ามา

from Backend import Data

# ตั้งค่า font ภาษาไทย
def set_font_thai(run, size_pt=16, bold=False):
    run.font.name = 'TH Sarabun New'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH Sarabun New')
    run.font.size = Pt(size_pt)
    run.font.bold = bold

# แปลงข้อความจาก textarea
def prepare_body_paragraphs(doc, raw_text):
    lines = raw_text.split('\n')
    for line in lines:
        # แก้ไขบรรทัดนี้: ลบการแทนที่แท็บ 6 ช่องว่าง เพราะเราจะใช้ paragraph_format.first_line_indent แทน
        clean_line = line.strip() # ใช้ strip() เพื่อลบช่องว่างหัวท้าย
        if clean_line:
            para = doc.add_paragraph()
            # เพิ่ม indent สำหรับย่อหน้าแรกของแต่ละบรรทัด (ถ้าจำเป็น)
            # ถ้าต้องการให้ทุกย่อหน้ามี indent ให้ใส่ใน clean_line ก่อน add_run หรือใช้ style
            para.paragraph_format.first_line_indent = Cm(1.27)
            run = para.add_run(clean_line)
            set_font_thai(run, size_pt=16)

# 🔁 บันทึกไฟล์พร้อม retry และ kill Word อัตโนมัติ
def save_doc_with_retry(doc, filename="Sleeve1_Output.docx", max_retries=3):
    for attempt in range(max_retries):
        try:
            doc.save(filename)
            print(f"✅ {filename} created successfully!")
            # ✅ เพิ่มส่วนนี้เพื่อเปิดไฟล์อัตโนมัติ
            if platform.system() == "Windows":
                os.startfile(filename)
            return True
        except PermissionError:
            print(f"⚠️ ไม่สามารถบันทึกไฟล์ {filename} ได้ อาจยังเปิดอยู่ใน Word")
            print("🛑 กำลังพยายามปิด Microsoft Word อัตโนมัติ...")
            os.system("taskkill /f /im WINWORD.EXE")
            time.sleep(2)  # รอให้ Word ปิด
    print("❌ ไม่สามารถบันทึกไฟล์ได้ กรุณาปิดไฟล์ด้วยตนเองแล้วลองใหม่อีกครั้ง")
    return False

Datum = Data()
Datum.appendlist("ถั่วเขียวเราะเปลือก", "ว.งานบ้านงานครัว", "1 ถุง", "มิ.ย.68")
Datum.appendlist("ถั่วแดงหลวง", "ว.งานบ้านงานครัว", "8 ถุง", "")
Datum.appendlist("ใบชา", "ว.งานบ้านงานครัว", "2 กล่อง", "")
Datum.appendlist("ถุงใส ขนาด 20x30 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "")
Datum.appendlist("ถุงตัดตรง LLDPE ขนาด 16x26 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "")

def Sleeve1(Data, title, runNumber, bodyText1):
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'TH Sarabun New'
    font.size = Pt(16)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH Sarabun New')

    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("บันทึกข้อความ")
    set_font_thai(run, size_pt=22, bold=True)

    p_gov_section = doc.add_paragraph("ส่วนราชการ ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ฯ ทรัพยากรธรรมชาติและสิ่งแวดล้อม โทร. 2749")
    p_gov_section.paragraph_format.space_after = Pt(0)

    p_ref_date = doc.add_paragraph()
    p_ref_date.add_run(f"ที่ {runNumber}")
    p_ref_date.paragraph_format.space_after = Pt(0)
    p_ref_date.paragraph_format.tab_stops.add_tab_stop(Inches(5.5), WD_PARAGRAPH_ALIGNMENT.RIGHT)
    p_ref_date.add_run("\t")
    date_run = p_ref_date.add_run(f"วันที่ {Data.day}")
    set_font_thai(date_run, size_pt=16)

    p_subject = doc.add_paragraph(f"เรื่อง {title}")
    set_font_thai(p_subject.runs[0], size_pt=16)
    p_subject.paragraph_format.space_after = Pt(0)

    p_line = doc.add_paragraph()
    run_line = p_line.add_run("-" * 110)
    set_font_thai(run_line, size_pt=16)
    p_line.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_line.paragraph_format.space_after = Pt(0)
    p_line.paragraph_format.space_before = Pt(0)

    p_dean = doc.add_paragraph("เรียน คณบดีคณะเกษตรศาสตร์ฯ")
    set_font_thai(p_dean.runs[0], size_pt=16)
    p_dean.paragraph_format.space_before = Pt(0)
    p_dean.paragraph_format.space_after = Pt(12)

    # ✅ เนื้อความ
    prepare_body_paragraphs(doc, bodyText1)

    # ✅ ตาราง
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells

    headers = [
        "ลำดับ",
        "รายการขอซื้อจ้าง\n[ขั้นตอนตามระเบียบฯ ข้อ 22(2)]",
        "รายการแยกวัสดุ\nตามระบบบัญชี\n3 มิติ",
        "จำนวนหน่วย",
        "กำหนดเวลาที่\nต้องการใช้พัสดุ\n[ตามระเบียบข้อ\n22 (5)]"
    ]

    for i, header_text in enumerate(headers):
        run = hdr_cells[i].paragraphs[0].add_run(header_text)
        set_font_thai(run, bold=True, size_pt=14)
        hdr_cells[i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for idx, row_data in enumerate(Data.list, start=1):
        cells = table.add_row().cells

        item_name = row_data[0] if len(row_data) > 0 else ""
        category = row_data[1] if len(row_data) > 1 else ""
        quantity = row_data[2] if len(row_data) > 2 else ""
        date_needed = row_data[3] if len(row_data) > 3 else ""

        run_idx = cells[0].paragraphs[0].add_run(str(idx))
        set_font_thai(run_idx, size_pt=16)
        cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        run_item = cells[1].paragraphs[0].add_run(item_name)
        set_font_thai(run_item, size_pt=16)
        cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        cells[1].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        run_category = cells[2].paragraphs[0].add_run(category)
        set_font_thai(run_category, size_pt=16)
        cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[2].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        run_quantity = cells[3].paragraphs[0].add_run(quantity)
        set_font_thai(run_quantity, size_pt=16)
        cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[3].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        run_date_needed = cells[4].paragraphs[0].add_run(date_needed)
        set_font_thai(run_date_needed, size_pt=16)
        cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[4].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for _ in range(3):
        doc.add_paragraph()

    p_signature = doc.add_paragraph()
    p_signature.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_sig = p_signature.add_run("ลงชื่อ ..........................................................")
    set_font_thai(run_sig, size_pt=16)

    p_name = doc.add_paragraph()
    p_name.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_name = p_name.add_run("(รศ.ดร.ทิพวรรณ ทองสุข)")
    set_font_thai(run_name, size_pt=16)

    return save_doc_with_retry(doc)

# ✅ ทดสอบ
if __name__ == '__main__':
    title = "ขออนุมัติABC"
    run = "อว 0603.07.04/"
    example_text = (
        "        ด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจัดซื้อวัสดุงานบ้านงานครัว จำนวน 11 รายการ\n"
        "        เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568\n"
        "        และต้องการใช้สิ่งของดังกล่าว ประมาณ มิถุนายน 2568\n"
        "        โดยขอแต่งตั้งกรรมการตรวจรับ คือ ผศ.ดร.ปริตา ธนสุกาญจน์"
    )
    Sleeve1(Datum, title, run, example_text)