from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches 
from Backend import Data
# Helper function to set Thai font properties
def set_font_thai(run, size_pt=16, bold=False):
    """
    Sets the font for a run object to 'TH Sarabun New' with specified size and bold status.
    """
    run.font.name = 'TH Sarabun New'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH Sarabun New')
    run.font.size = Pt(size_pt)
    run.font.bold = bold

Datum = Data()
Datum.appendlist("ถั่วเขียวเราะเปลือก", "ว.งานบ้านงานครัว", "1 ถุง", "มิ.ย.68")
Datum.appendlist("ถั่วแดงหลวง", "ว.งานบ้านงานครัว", "8 ถุง", "")
Datum.appendlist("ใบชา", "ว.งานบ้านงานครัว", "2 กล่อง", "")
Datum.appendlist("ถุงใส ขนาด 20x30 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "")
Datum.appendlist("ถุงตัดตรง LLDPE ขนาด 16x26 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "")


def Sleeve1(Data):
    doc = Document()

    # --- Initial Font Settings ---
    style = doc.styles['Normal']
    font = style.font
    font.name = 'TH Sarabun New'
    font.size = Pt(16)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH Sarabun New')
    # --------------------------------

    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    # Header - "บันทึกข้อความ"
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("บันทึกข้อความ")
    set_font_thai(run, size_pt=22, bold=True)

    # Empty paragraph for spacing
    #doc.add_paragraph()

    # Header Section Details
    # Adjusted space after for the paragraph containing "ส่วนราชการ"
    p_gov_section = doc.add_paragraph("ส่วนราชการ ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ฯ ทรัพยากรธรรมชาติและสิ่งแวดล้อม โทร. 2749")
    p_gov_section.paragraph_format.space_after = Pt(0) # Reduce space after

    # "ที่ อว..." line with right-aligned date
    p_ref_date = doc.add_paragraph()
    p_ref_date.add_run("ที่ อว 0603.07.04/")
    p_ref_date.paragraph_format.space_after = Pt(0) # Reduce space after
    p_ref_date.paragraph_format.tab_stops.add_tab_stop(Inches(5.5), WD_PARAGRAPH_ALIGNMENT.RIGHT) # Adjusted tab stop using Inches for more precision
    p_ref_date.add_run("\t")
    date_run = p_ref_date.add_run(f"วันที่ {Data.day}")
    set_font_thai(date_run, size_pt=16)

    # Subject
    p_subject = doc.add_paragraph("เรื่อง ขออนุมัติจัดซื้อวัสดุ")
    set_font_thai(p_subject.runs[0], size_pt=16)
    # Changed space_after to 0 to make the line appear immediately below
    p_subject.paragraph_format.space_after = Pt(0) 

    # Horizontal Line - Changed to hyphens and removed border logic
    p_line = doc.add_paragraph()
    run_line = p_line.add_run("-" * 110) # Use hyphens to form the line
    set_font_thai(run_line, size_pt=16) # Apply font to hyphens
    p_line.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT # Align left to match image
    p_line.paragraph_format.space_after = Pt(0) # No space after
    p_line.paragraph_format.space_before = Pt(0) # No space before
    
    # Removed border creation code as it's no longer used
    # p_line_element = p_line._element
    # pPr = p_line_element.get_or_add_pPr()
    # pBdr = OxmlElement('w:pBdr')
    # pPr.append(pBdr)
    # bottom_border_element = OxmlElement('w:bottom')
    # bottom_border_element.set(qn('w:val'), 'single')
    # bottom_border_element.set(qn('w:sz'), '6')
    # bottom_border_element.set(qn('w:space'), '1')
    # pBdr.append(bottom_border_element)


    p_dean = doc.add_paragraph("เรียน คณบดีคณะเกษตรศาสตร์ฯ")
    set_font_thai(p_dean.runs[0], size_pt=16)
    # Changed space_before to 0 to make the text appear immediately below the line
    p_dean.paragraph_format.space_before = Pt(0) 
    p_dean.paragraph_format.space_after = Pt(12) 

    # Main Body Content
    body_text = (
        "\tด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจัดซื้อวัสดุงานบ้านงานครัว จำนวน "
        "11 รายการ เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568 "
        "ในโครงการพัฒนาผลิตภัณฑ์ นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร และต้องการใช้สิ่งของดังกล่าว "
        "ประมาณ(เดือน/ปี) มิถุนายน 2568 และเบิกจ่ายจากเงินงบประมาณงบประมาณรายได้ 2568 กองทุนเพื่อการศึกษา "
        "แผนงานจัดการศึกษาอุดมศึกษา งานจัดการศึกษาสาขาเกษตรศาสตร์ หลักสูตรวิทยาศาสตรบัณฑิต "
        "สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดเงินอุดหนุน โครงการพัฒนากระบวนการจัดการเรียนการสอน (โครงการพัฒนาผลิตภัณฑ์นิสิตสาขาวิทยาศาสตร์และเทคโนโลยีการ"
        "อาหาร) หมวดค่าวัสดุงานบ้านงานครัว เป็นเงิน 4,000 บาท (สี่พันบาทถ้วน)"
    )       
    para_body = doc.add_paragraph()
    run_body = para_body.add_run(body_text)
    set_font_thai(run_body, size_pt=16)
    # Ensure first line indent, if needed, you can explicitly set it:
    para_body.paragraph_format.first_line_indent = Cm(1.27) # Typical first line indent

    # Add the "โดยมีเห็นควรมอบหมาย..." paragraph
    body_text_2 = (
        "โดยมีเห็นควรมอบหมายผู้รับผิดชอบในการจัดทำรายละเอียดคุณลักษณะของพัสดุ ตามระเบียบฯ "
        "ข้อ 21 ดังนี้ รศ.ดร.ทิพวรรณ ทองสุข และขอแต่งตั้งกรรมการตรวจรับ คือ ผศ.ดร.ปริตา ธนสุกาญจน์"
    )
    para_body_2 = doc.add_paragraph()
    run_body_2 = para_body_2.add_run(body_text_2)
    set_font_thai(run_body_2, size_pt=16)
    para_body_2.paragraph_format.space_after = Pt(24) # Add more space after this paragraph before the table

    # --- Table Section ---
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells

    # Set column widths (fine-tuned to match image)
    table.columns[0].width = Cm(1.5)  # ลำดับ
    table.columns[1].width = Cm(7)    # รายการขอซื้อจ้าง
    table.columns[2].width = Cm(5)    # รายการแยกวัสดุ
    table.columns[3].width = Cm(2)    # จำนวนหน่วย
    table.columns[4].width = Cm(3)    # กำหนดเวลาที่ต้องการใช้พัสดุ

    # Headers for the table, matching the image exactly
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

    # Populate table with data from Data.list
    for idx, row_data in enumerate(Data.list, start=1):
        cells = table.add_row().cells

        # Safely get data, provide empty string if index out of range
        item_name = row_data[0] if len(row_data) > 0 else ""
        category = row_data[1] if len(row_data) > 1 else ""
        quantity = row_data[2] if len(row_data) > 2 else ""
        date_needed = row_data[3] if len(row_data) > 3 else ""

        # ลำดับ (Sequence Number)
        run_idx = cells[0].paragraphs[0].add_run(str(idx))
        set_font_thai(run_idx, size_pt=16)
        cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # รายการขอซื้อจ้าง
        run_item = cells[1].paragraphs[0].add_run(item_name)
        set_font_thai(run_item, size_pt=16)
        cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        cells[1].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # รายการแยกวัสดุตามระบบบัญชี 3 มิติ
        run_category = cells[2].paragraphs[0].add_run(category)
        set_font_thai(run_category, size_pt=16)
        cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[2].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # จำนวนหน่วย
        run_quantity = cells[3].paragraphs[0].add_run(quantity)
        set_font_thai(run_quantity, size_pt=16)
        cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[3].vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # กำหนดเวลาที่ต้องการใช้พัสดุ
        run_date_needed = cells[4].paragraphs[0].add_run(date_needed)
        set_font_thai(run_date_needed, size_pt=16)
        cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cells[4].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Add extra empty lines after the table for spacing before signature
    for _ in range(3):
        doc.add_paragraph()

    # Signature Section
    p_signature = doc.add_paragraph()
    p_signature.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_sig = p_signature.add_run("ลงชื่อ ..........................................................")
    set_font_thai(run_sig, size_pt=16)

    p_name = doc.add_paragraph()
    p_name.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_name = p_name.add_run("(รศ.ดร.ทิพวรรณ ทองสุข)")
    set_font_thai(run_name, size_pt=16)

    doc.save("Sleeve1_Output.docx")
    print("✅ Sleeve1_Output.docx created successfully!")
    return 1

# Example usage with dummy data
if __name__ == '__main__':
    Sleeve1(Datum)
