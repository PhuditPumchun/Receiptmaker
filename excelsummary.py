# File: excelsummary.py

import os
import platform
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

# นำเข้าคลาส Data จาก backend.py
# สมมติว่ามีไฟล์ backend.py ที่มีคลาส Data และเมธอด parse_amount และ format_thai_date
from Backend import Data

def create_excel_summary(data_list, transaction_info, filename="Summary_Output.xlsx"):
    """
    สร้างไฟล์ Excel สรุปยอดโดยใช้ข้อมูลที่รับมาและใส่สูตรคำนวณ
    - data_list: รายการพัสดุทั้งหมด
    - transaction_info: dict ข้อมูลสรุป (วันที่รับ, จ่ายให้)
    """
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "สรุปยอด"

        # --- การตั้งค่า Font และ Border ---
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        thai_font = Font(name='TH Sarabun New', size=11)
        bold_thai_font = Font(name='TH Sarabun New', size=11, bold=True)
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')
        right_align = Alignment(horizontal='right', vertical='center')
        top_center_align = Alignment(horizontal='center', vertical='top')

        # --- ส่วนหัวตาราง ---
        # ปรับการ Merge Cells ให้ครอบคลุมแค่ A ถึง J
        ws.merge_cells('A1:J1')
        ws['A1'] = "เล่มที่.......... เลขที่..........."
        ws['A1'].font = bold_thai_font
        ws['A1'].alignment = right_align

        ws.merge_cells('A2:J2')
        ws['A2'] = "ชื่อ คณะเกษตรศาสตร์"
        ws['A2'].font = bold_thai_font
        ws['A2'].alignment = center_align

        ws.merge_cells('A3:J3')
        ws['A3'] = "งบประมาณรายได้ปี 2568 ภาควิชาอุตสาหกรรมการเกษตร คณะเกษตรศาสตร์ฯ"
        ws['A3'].font = bold_thai_font
        ws['A3'].alignment = center_align
        
        # --- Headers หลักและย่อย (ปรับตามที่ต้องการ) ---
        ws.merge_cells('A4:A5') # วันที่
        ws.merge_cells('B4:B5') # รายการ
        ws.merge_cells('C4:E4') # รับ
        ws.merge_cells('F4:H4') # จ่าย
        ws.merge_cells('I4:J4') # คงเหลือ (ปรับให้ครอบคลุม J4 เท่านั้น)
        
        # Main headers text
        ws['A4'] = 'วันที่'
        ws['B4'] = 'รายการ'
        ws['C4'] = 'รับ'
        ws['F4'] = 'จ่าย'
        ws['I4'] = 'คงเหลือ'

        # Sub-headers for 'รับ'
        ws['C5'] = 'ใบรับที่'
        ws['D5'] = 'จำนวน'
        ws['E5'] = 'บาท'

        # Sub-headers for 'จ่าย'
        ws['F5'] = 'ใบรับที่' # จะใช้เป็น "รวมจ่าย" แทน
        ws['G5'] = 'จำนวน'
        ws['H5'] = 'บาท'

        # Sub-headers for 'คงเหลือ'
        ws['I5'] = 'จำนวน'
        ws['J5'] = 'บาท'

        # Apply styles to headers (A4 to J5)
        for row_idx in range(4, 6):
            for col_idx in range(1, 11): # Columns A to J
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = bold_thai_font
                cell.alignment = center_align
                cell.border = thin_border
        
        # --- การตั้งค่าความกว้างคอลัมน์ (ปรับตามคอลัมน์ที่เหลือ) ---
        col_widths = {
            'A': 15, # วันที่
            'B': 45, # รายการ
            'C': 15, 'D': 10, 'E': 12, # รับ(ใบรับที่, จำนวน, บาท)
            'F': 15, 'G': 10, 'H': 12, # จ่าย(ใบรับที่, จำนวน, บาท)
            'I': 10, 'J': 12, # คงเหลือ (จำนวน, บาท)
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        current_row = 6 # เริ่มต้นที่แถว 6 สำหรับข้อมูล

        # --- จัดกลุ่มข้อมูลตาม 'received_from' ---
        grouped_data = defaultdict(list)
        data_handler_instance = Data() # สร้าง instance ของ Data เพื่อใช้ parse_amount และ format_thai_date
        for item in data_list:
            received_from_company = item[5] # Index 5 คือ received_from
            grouped_data[received_from_company].append(item)

        # ตัวแปรสำหรับเก็บยอดคงเหลือสะสม
        cumulative_balance_qty = 0
        cumulative_balance_baht = 0.0

        for company_name, items_for_company in grouped_data.items():
            # แถว "รับจาก"
            # A: วันที่ (ใช้ purchase_date แทน date_needed)
            first_item_date = items_for_company[0][7] if items_for_company else "" # เปลี่ยนเป็น index 7 สำหรับ purchase_date
            ws.cell(row=current_row, column=1, value=first_item_date).font = bold_thai_font
            ws.cell(row=current_row, column=1).alignment = center_align

            # B: รับจากบริษัท
            ws.cell(row=current_row, column=2, value=f"รับจาก {company_name}").font = bold_thai_font
            ws.cell(row=current_row, column=2).alignment = left_align
            
            # C: ใบรับที่ (ของรายการแรกในกลุ่ม)
            first_invoice_no = items_for_company[0][6] if items_for_company else "" # invoice_no คือ index 6
            ws.cell(row=current_row, column=3, value=first_invoice_no).font = bold_thai_font
            ws.cell(row=current_row, column=3).alignment = center_align

            # D: จำนวน (จำนวนรายการที่มี) - เพิ่ม " รก."
            item_count_for_company = len(items_for_company) # จำนวนรายการของบริษัทนี้
            ws.cell(row=current_row, column=4, value=f"{item_count_for_company} รก.").font = bold_thai_font
            ws.cell(row=current_row, column=4).alignment = center_align

            # E: บาท (รวมเงินของรายการในกลุ่มนี้) - คำนวณ จำนวน * ราคา
            total_price_for_company_E = 0
            for item in items_for_company:
                quantity_numeric = data_handler_instance.parse_amount(item[2]) # item[2] คือ amount (เช่น "10 ด้าม")
                price_per_item = item[4] # item[4] คือ price (เช่น 15.00)
                total_price_for_company_E += (quantity_numeric * price_per_item)

            cell_company_total_price_E = ws.cell(row=current_row, column=5, value=total_price_for_company_E)
            cell_company_total_price_E.font = bold_thai_font
            cell_company_total_price_E.number_format = '#,##0.00'
            cell_company_total_price_E.alignment = right_align

            # F, G, H (จ่าย) - ใส่เครื่องหมาย "-"
            ws.cell(row=current_row, column=6, value="-").font = thai_font
            ws.cell(row=current_row, column=6).alignment = center_align
            ws.cell(row=current_row, column=7, value="-").font = thai_font
            ws.cell(row=current_row, column=7).alignment = center_align
            ws.cell(row=current_row, column=8, value="-").font = thai_font
            ws.cell(row=current_row, column=8).alignment = center_align


            # I: จำนวน (คงเหลือ) - ยอดสะสมจำนวน
            cumulative_balance_qty += item_count_for_company
            ws.cell(row=current_row, column=9, value=f"{cumulative_balance_qty} รก.").font = bold_thai_font
            ws.cell(row=current_row, column=9).alignment = center_align

            # J: บาท (คงเหลือ) - ยอดสะสมบาท
            cumulative_balance_baht += total_price_for_company_E
            cell_company_total_price_J = ws.cell(row=current_row, column=10, value=cumulative_balance_baht) 
            cell_company_total_price_J.font = bold_thai_font
            cell_company_total_price_J.number_format = '#,##0.00'
            cell_company_total_price_J.alignment = right_align

            # Apply border to the entire row
            for col_idx in range(1, 11): # Columns A to J
                ws.cell(row=current_row, column=col_idx).border = thin_border
            current_row += 1
            
            for item in items_for_company:
                # [0:name, 1:category, 2:amount, 3:date_needed, 4:price, 5:received_from, 6:invoice_no, 7: purchase_date]
                item_name = item[0]
                amount_str = item[2] # เช่น "10 ด้าม"
                price_per_item = item[4] # เช่น 15.00
                
                quantity_numeric = data_handler_instance.parse_amount(amount_str)
                unit = ''.join(filter(str.isalpha, amount_str)) # ดึงส่วนที่เป็นตัวอักษรของหน่วยนับ

                total_item_price = quantity_numeric * price_per_item

                # สร้างข้อความตาม format ที่ต้องการ
                if quantity_numeric > 1:
                    item_display_text = f"-{item_name} {quantity_numeric} {unit}@{price_per_item:.2f}.- = {total_item_price:.2f}.-"
                else:
                    item_display_text = f"-{item_name} {quantity_numeric} {unit}@{price_per_item:.2f}.-"


                # A: วันที่ - เว้นว่าง
                ws.cell(row=current_row, column=1, value="").font = thai_font
                ws.cell(row=current_row, column=1).alignment = center_align

                # B: รายการ (แสดงในรูปแบบใหม่)
                ws.cell(row=current_row, column=2, value=item_display_text).font = thai_font
                ws.cell(row=current_row, column=2).alignment = left_align

                # C, D, E (รับ) - เว้นว่าง
                ws.cell(row=current_row, column=3, value="").font = thai_font
                ws.cell(row=current_row, column=4, value="").font = thai_font
                ws.cell(row=current_row, column=5, value="").font = thai_font
                
                # F, G, H (จ่าย) - เว้นว่างสำหรับกรอกข้อมูล
                ws.cell(row=current_row, column=6, value="").font = thai_font
                ws.cell(row=current_row, column=7, value="").font = thai_font
                ws.cell(row=current_row, column=8, value="").font = thai_font

                # I, J (คงเหลือ) - เว้นว่าง
                ws.cell(row=current_row, column=9, value="").font = thai_font
                ws.cell(row=current_row, column=10, value="").font = thai_font

                for col_idx in range(1, 11): # Apply border to all cells in the row (A to J)
                    ws.cell(row=current_row, column=col_idx).border = thin_border
                
                current_row += 1

        # --- แถวสรุป "จ่ายให้" และ "รวมจ่าย" ---
        current_row += 1 # เว้น 1 บรรทัด
        
        # A: วันที่ (เว้นว่าง)
        ws.cell(row=current_row, column=1, value="").font = thai_font
        ws.cell(row=current_row, column=1).alignment = center_align

        # B: รายการ (จ่ายให้...)
        ws.cell(row=current_row, column=2, value=f"จ่ายให้ {transaction_info['paid_to']}").font = bold_thai_font
        ws.cell(row=current_row, column=2).alignment = left_align

        # C, D, E (รับ) - เว้นว่าง
        ws.cell(row=current_row, column=3, value="").font = thai_font
        ws.cell(row=current_row, column=4, value="").font = thai_font
        ws.cell(row=current_row, column=5, value="").font = thai_font

        # # F: ใบรับที่ (เปลี่ยนเป็น "รวมจ่าย")
        # ws.cell(row=current_row, column=6, value="รวมจ่าย").font = bold_thai_font
        # ws.cell(row=current_row, column=6).alignment = center_align

        # G: จำนวน (ใช้ค่าจาก cumulative_balance_qty ล่าสุด และเพิ่ม " รก." เข้าไปด้วย)
        ws.cell(row=current_row, column=7, value=f"{cumulative_balance_qty} รก.").font = bold_thai_font
        ws.cell(row=current_row, column=7).alignment = center_align

        # H: บาท (ใช้ค่าจาก cumulative_balance_baht ล่าสุด)
        ws.cell(row=current_row, column=8, value=cumulative_balance_baht).font = bold_thai_font
        ws.cell(row=current_row, column=8).number_format = '#,##0.00'
        ws.cell(row=current_row, column=8).alignment = right_align

        # I, J (คงเหลือ) - เว้นว่าง
        ws.cell(row=current_row, column=9, value="").font = thai_font
        ws.cell(row=current_row, column=10, value="").font = thai_font

        for col_idx in range(1, 11): # A to J
            ws.cell(row=current_row, column=col_idx).border = thin_border
        
        # --- Save the workbook ---
        wb.save(filename)
        print(f"✅ Excel file '{filename}' created successfully!")
        if platform.system() == "Windows":
            os.startfile(filename)
        return True
    except Exception as e:
        print(f"❌ Error creating Excel file: {e}")
        return False

def main():
    """
    ฟังก์ชันหลักสำหรับสร้างข้อมูลตัวอย่างและเรียกใช้ create_excel_summary
    โดยมีข้อมูลจาก 3 บริษัท ใน 3 วันที่แตกต่างกัน
    """
    print("🚀 กำลังสร้างข้อมูลตัวอย่างและไฟล์ Excel สรุปยอด...")

    # สร้าง instance ของคลาส Data
    data_handler = Data()

    # --- ข้อมูลตัวอย่างสำหรับ 3 วันที่แตกต่างกัน และ 3 บริษัท ---

    # วันที่ 10/มิ.ย./68
    date_1 = "10/มิ.ย./68"
    company_a = "บริษัท สยามพัฒนา จำกัด"
    company_b = "ร้านค้าส่งอุปกรณ์"
    company_c = "บริษัท เครื่องเขียนไทย"

    data_handler.appendlist("ปากกาเคมี", "วัสดุสำนักงาน", "5 ด้าม", date_1, 20.00, company_a, "INV_A_001",date_1)
    data_handler.appendlist("กระดาษโน้ต", "วัสดุสำนักงาน", "3 เล่ม", date_1, 15.00, company_a, "INV_A_001",date_1)
    data_handler.appendlist("เทปใส", "วัสดุสำนักงาน", "2 ม้วน", date_1, 10.00, company_b, "INV_B_001",date_1)
    data_handler.appendlist("กรรไกร", "วัสดุสำนักงาน", "1 อัน", date_1, 35.00, company_c, "INV_C_001",date_1)

    # วันที่ 15/มิ.ย./68
    date_2 = "15/มิ.ย./68"
    company_d = "บริษัท วัสดุก่อสร้าง"
    company_e = "ร้านอุปกรณ์ไฟฟ้า"
    company_f = "บริษัท เฟอร์นิเจอร์ดี"

    data_handler.appendlist("หลอดไฟ LED", "วัสดุไฟฟ้า", "10 หลอด", date_2, 45.00, company_d, "INV_D_001",date_2)
    data_handler.appendlist("สายไฟ", "วัสดุไฟฟ้า", "1 ม้วน", date_2, 200.00, company_d, "INV_D_001",date_2)
    data_handler.appendlist("ไขควงชุด", "วัสดุช่าง", "1 ชุด", date_2, 150.00, company_e, "INV_E_001",date_2)
    data_handler.appendlist("เก้าอี้สำนักงาน", "ครุภัณฑ์", "1 ตัว", date_2, 1200.00, company_f, "INV_F_001",date_2)

    # วันที่ 20/มิ.ย./68
    date_3 = "20/มิ.ย./68"
    company_g = "บริษัท อาหารสด"
    company_h = "ร้านขายน้ำดื่ม"
    company_i = "บริษัท อุปกรณ์กีฬา"

    data_handler.appendlist("นมสด", "วัสดุบริโภค", "6 กล่อง", date_3, 30.00, company_g, "INV_G_001",date_3)
    data_handler.appendlist("น้ำดื่ม", "วัสดุบริโภค", "10 แพ็ค", date_3, 60.00, company_h, "INV_H_001",date_3)
    data_handler.appendlist("ลูกฟุตบอล", "วัสดุการศึกษา", "1 ลูก", date_3, 500.50, company_i, "INV_I_001",date_3)

    # กำหนดข้อมูล transaction_info (ใช้วันที่ปัจจุบันสำหรับการแสดงหัวบันทึกข้อความ)
    transaction_info = {
        "paid_to": "นาย ก. (ผู้รับของ)" # ข้อมูลนี้จะถูกนำไปใช้ในคอลัมน์ "รายการ"
    }

    # เรียกใช้ฟังก์ชันสร้าง Excel
    create_excel_summary(data_handler.list, transaction_info)
    print("✅ การสร้างไฟล์ Excel เสร็จสมบูรณ์")

if __name__ == "__main__":
    main()