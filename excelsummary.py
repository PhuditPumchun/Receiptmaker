# File: excelsummary.py

import os
import platform
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

def create_excel_summary(data_list, transaction_info, filename="Summary_Output.xlsx"):
    """
    สร้างไฟล์ Excel สรุปยอดโดยใช้ข้อมูลที่รับมาและใส่สูตรคำนวณ
    - data_list: รายการพัสดุทั้งหมด
    - transaction_info: dict ข้อมูลสรุป (วันที่รับ, รับจาก, จ่ายให้)
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
        # "บันทึกข้อความ"
        ws.merge_cells('A1:L1')
        ws['A1'] = "บันทึกข้อความ"
        ws['A1'].font = bold_thai_font
        ws['A1'].alignment = center_align

        # "ภาควิชาอุตสาหกรรมเกษตร"
        ws.merge_cells('A2:L2')
        ws['A2'] = "ภาควิชาอุตสาหกรรมเกษตร"
        ws['A2'].font = bold_thai_font
        ws['A2'].alignment = center_align

        # "สรุปยอดพัสดุ"
        ws.merge_cells('A3:L3')
        ws['A3'] = "สรุปยอดพัสดุ"
        ws['A3'].font = bold_thai_font
        ws['A3'].alignment = center_align
        
        # --- Headers หลักและย่อย ---
        # Merge for main headers
        ws.merge_cells('A4:A5') # รายการ
        ws.merge_cells('B4:D4') # รับ
        ws.merge_cells('E4:G4') # จ่าย
        ws.merge_cells('H4:L4') # คงเหลือ

        # Main headers text
        ws['A4'] = 'รายการ'
        ws['B4'] = 'รับ'
        ws['E4'] = 'จ่าย'
        ws['H4'] = 'คงเหลือ'

        # Sub-headers for 'รับ'
        ws['B5'] = 'ใบรับที่'
        ws['C5'] = 'จำนวน'
        ws['D5'] = 'บาท'

        # Sub-headers for 'จ่าย'
        ws['E5'] = 'ใบรับที่'
        ws['F5'] = 'จำนวน'
        ws['G5'] = 'บาท'

        # Sub-headers for 'คงเหลือ'
        ws['H5'] = 'จำนวน'
        ws['I5'] = 'บาท'
        ws['J5'] = 'จำนวน' # ช่องที่ 3 ในคงเหลือ
        ws['K5'] = 'บาท'   # ช่องที่ 4 ในคงเหลือ
        ws['L5'] = 'แหล่งที่มา' # คอลัมน์ L ในคงเหลือ

        # Apply styles to headers (B4 to L5)
        for row_idx in range(4, 6):
            for col_idx in range(1, 13): # Columns A to L
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = bold_thai_font
                cell.alignment = center_align
                cell.border = thin_border
        
        # --- การตั้งค่าความกว้างคอลัมน์ ---
        col_widths = {
            'A': 30, 'B': 15, 'C': 10, 'D': 12, # รายการ, รับ(ใบรับที่, จำนวน, บาท)
            'E': 15, 'F': 10, 'G': 12,          # จ่าย(ใบรับที่, จำนวน, บาท)
            'H': 10, 'I': 12, 'J': 10, 'K': 12, 'L': 15 # คงเหลือ (จำนวน, บาท, จำนวน, บาท, แหล่งที่มา)
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # --- ส่วนของข้อมูล (Data Section) ---
        current_row = 6
        
        # แถว "รับจาก" - ใช้ข้อมูลจาก transaction_info
        # A6: รับจาก...
        ws.cell(row=current_row, column=1, value=f"{transaction_info['receipt_date']} รับจาก {transaction_info['received_from']}").font = bold_thai_font
        ws.cell(row=current_row, column=1).alignment = left_align
        
        # Merge cells A6:L6 for the "รับจาก" row
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=12)
        
        for col_idx in range(1, 13): # Apply border to the merged cell
            ws.cell(row=current_row, column=col_idx).border = thin_border

        current_row += 1
        
        start_data_row = current_row

        # วนลูปเพื่อใส่ข้อมูลแต่ละรายการ
        for item in data_list:
            # [0:name, 1:category, 2:amount, 3:date, 4:price, 5:received_from, 6:invoice_no]
            item_name, _, quantity_text, _, price, _, invoice_no = item
            
            # A: รายการ (item_name)
            ws.cell(row=current_row, column=1, value=f"-{item_name}").font = thai_font
            ws.cell(row=current_row, column=1).alignment = left_align

            # B: ใบรับที่ (รับ)
            ws.cell(row=current_row, column=2, value=invoice_no).font = thai_font
            ws.cell(row=current_row, column=2).alignment = center_align

            # C: จำนวน (รับ)
            ws.cell(row=current_row, column=3, value=quantity_text).font = thai_font
            ws.cell(row=current_row, column=3).alignment = center_align

            # D: บาท (รับ)
            cell_price_received = ws.cell(row=current_row, column=4, value=price)
            cell_price_received.font = thai_font
            cell_price_received.number_format = '#,##0.00'
            cell_price_received.alignment = right_align
            
            # E, F, G (จ่าย) - เว้นว่างสำหรับกรอกข้อมูล
            ws.cell(row=current_row, column=5, value="").font = thai_font
            ws.cell(row=current_row, column=5).alignment = center_align
            ws.cell(row=current_row, column=6, value="").font = thai_font
            ws.cell(row=current_row, column=6).alignment = center_align
            cell_paid_amount = ws.cell(row=current_row, column=7, value="")
            cell_paid_amount.font = thai_font
            cell_paid_amount.number_format = '#,##0.00'
            cell_paid_amount.alignment = right_align

            # H, I, J, K (คงเหลือ) - ใส่สูตร Excel
            # จำนวนคงเหลือ (H) - ยังไม่มีสูตรที่ชัดเจนจากรูป ให้ผู้ใช้กรอกเอง หรือคำนวณง่ายๆ (รับ - จ่าย)
            ws.cell(row=current_row, column=8, value=f"=C{current_row}-F{current_row}").font = thai_font
            ws.cell(row=current_row, column=8).alignment = center_align

            # บาทคงเหลือ (I) - สูตร: บาทรับ - บาทจ่าย
            cell_balance_price = ws.cell(row=current_row, column=9, value=f"=D{current_row}-G{current_row}")
            cell_balance_price.font = thai_font
            cell_balance_price.number_format = '#,##0.00'
            cell_balance_price.alignment = right_align
            
            # ช่องว่าง (J, K) - ว่างไว้ตามรูป
            ws.cell(row=current_row, column=10, value="").font = thai_font
            ws.cell(row=current_row, column=10).alignment = center_align
            ws.cell(row=current_row, column=11, value="").font = thai_font
            ws.cell(row=current_row, column=11).alignment = right_align

            # L: แหล่งที่มา - ในรูปแรกไม่มี แต่ในรูป Page 1 มี อาจจะเพิ่ม หรือเว้นว่าง
            ws.cell(row=current_row, column=12, value="").font = thai_font
            ws.cell(row=current_row, column=12).alignment = left_align


            for col_idx in range(1, 13): # Apply border to all cells in the row
                ws.cell(row=current_row, column=col_idx).border = thin_border
            
            current_row += 1

        end_data_row = current_row - 1
        
        # --- แถวสรุปรวม (Total) ---
        ws.cell(row=current_row, column=1, value="รวม").font = bold_thai_font
        ws.cell(row=current_row, column=1).alignment = right_align
        
        # สูตร SUM สำหรับคอลัมน์ "บาท (รับ)" (D)
        total_received_cell_d = ws.cell(row=current_row, column=4)
        total_received_cell_d.font = bold_thai_font
        total_received_cell_d.value = f"=SUM(D{start_data_row}:D{end_data_row})"
        total_received_cell_d.number_format = '#,##0.00'
        total_received_cell_d.alignment = right_align

        # สูตร SUM สำหรับคอลัมน์ "บาท (จ่าย)" (G)
        total_paid_cell_g = ws.cell(row=current_row, column=7)
        total_paid_cell_g.font = bold_thai_font
        total_paid_cell_g.value = f"=SUM(G{start_data_row}:G{end_data_row})"
        total_paid_cell_g.number_format = '#,##0.00'
        total_paid_cell_g.alignment = right_align

        # สูตร SUM สำหรับคอลัมน์ "บาท (คงเหลือ)" (I)
        total_balance_cell_i = ws.cell(row=current_row, column=9)
        total_balance_cell_i.font = bold_thai_font
        total_balance_cell_i.value = f"=SUM(I{start_data_row}:I{end_data_row})" 
        total_balance_cell_i.number_format = '#,##0.00'
        total_balance_cell_i.alignment = right_align
        
        # Merge cells for "รวม" (A to C)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)

        # Apply border to the "รวม" row
        for col_idx in range(1, 13):
            ws.cell(row=current_row, column=col_idx).border = thin_border
        current_row += 1 # เพิ่มบรรทัดว่างหลังรวม

        # --- แถวสรุป "จ่ายไป" ---
        current_row += 1 # เว้น 1 บรรทัดตามรูป
        # A: จ่ายไป...
        ws.cell(row=current_row, column=1, value=f"จ่ายไป {transaction_info['paid_to']}").font = bold_thai_font
        ws.cell(row=current_row, column=1).alignment = left_align
        
        # E: "1 รก."
        ws.cell(row=current_row, column=5, value="1 รก.").font = bold_thai_font
        ws.cell(row=current_row, column=5).alignment = center_align

        # G: ค่าเงิน
        paid_amount_cell_g = ws.cell(row=current_row, column=7)
        paid_amount_cell_g.font = bold_thai_font
        paid_amount_cell_g.value = f"=G{end_data_row + 1}" # อ้างอิงจากช่องรวมยอดจ่าย (G)
        paid_amount_cell_g.number_format = '#,##0.00'
        paid_amount_cell_g.alignment = right_align

        # Merge cells A to D for "จ่ายไป" text
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)

        # Merge cells for "1 รก." (E to F)
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=6)

        # Apply border to the "จ่ายไป" row
        for col_idx in range(1, 13):
            ws.cell(row=current_row, column=col_idx).border = thin_border
        
        # --- Save the workbook ---
        wb.save(filename)
        print(f"✅ Excel file '{filename}' created successfully with formulas!")
        if platform.system() == "Windows":
            os.startfile(filename)
        return True
    except Exception as e:
        print(f"❌ Error creating Excel file: {e}")
        return False