import os
import platform
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

def create_excel_summary(data_list, filename="Summary_Output.xlsx"):
    """
    สร้างไฟล์ Excel สรุปยอดจากรายการข้อมูลที่ได้รับมา

    Args:
        data_list (list): รายการข้อมูลที่ได้จาก Data.list แต่ละรายการเป็น tuple/list
                          (ชื่อพัสดุ, หน่วยงาน, จำนวน, วันที่ต้องการใช้)
        filename (str): ชื่อไฟล์ Excel ที่จะสร้าง
    """
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "สรุปยอด"

        # Define thin border for cells
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        # Define Thai font styles
        thai_font = Font(name='TH Sarabun New', size=11)
        bold_thai_font = Font(name='TH Sarabun New', size=11, bold=True)

        # --- Table Header Section (based on the image) ---
        # Merge cells and set main titles
        ws.merge_cells('B1:L1')
        ws['B1'] = "บันทึกข้อความ" 
        ws['B1'].font = bold_thai_font
        ws['B1'].alignment = Alignment(horizontal='center', vertical='center')

        ws.merge_cells('B2:L2')
        ws['B2'] = "ภาควิชาอุตสาหกรรมเกษตร"
        ws['B2'].font = bold_thai_font
        ws['B2'].alignment = Alignment(horizontal='center', vertical='center')

        ws.merge_cells('B3:L3')
        ws['B3'] = "สรุปยอดปี 2567 ภาควิชาอุตสาหกรรมเกษตร"
        ws['B3'].font = bold_thai_font
        ws['B3'].alignment = Alignment(horizontal='center', vertical='center')

        # Headers for main table structure (Row 4)
        ws['B4'] = "รายการ"
        ws.merge_cells('B4:B5') # Merge 'รายการ' across 2 rows
        ws['B4'].font = bold_thai_font # Apply font to the merged cell's top-left cell
        ws['B4'].alignment = Alignment(horizontal='center', vertical='center') # Center alignment

        ws.merge_cells('C4:E4')
        ws['C4'] = "รับ"
        ws['C4'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C4'].font = bold_thai_font

        ws.merge_cells('F4:H4') # Corrected: 'จ่าย' should span 3 columns
        ws['F4'] = "จ่าย"
        ws['F4'].alignment = Alignment(horizontal='center', vertical='center')
        ws['F4'].font = bold_thai_font
        
        ws.merge_cells('I4:L4') # Corrected: 'คงเหลือ' should span 4 columns
        ws['I4'] = "คงเหลือ"
        ws['I4'].alignment = Alignment(horizontal='center', vertical='center')
        ws['I4'].font = bold_thai_font

        # Headers for Row 5 (sub-headers)
        # Manually set headers for row 5 based on column mapping for clarity
        ws['C5'].value = "ใบรับที่"
        ws['D5'].value = "จำนวน"
        ws['E5'].value = "บาท"
        ws['F5'].value = "ใบรับที่"
        ws['G5'].value = "จำนวน"
        ws['H5'].value = "บาท"
        ws['I5'].value = "จำนวน"
        ws['J5'].value = "บาท"
        ws['K5'].value = "จำนวน" # Second 'จำนวน' under 'คงเหลือ'
        ws['L5'].value = "บาท"   # Second 'บาท' under 'คงเหลือ'

        # Apply font and border to all cells in row 5 headers
        for col_letter in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            cell = ws[f'{col_letter}5']
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.font = thai_font
            cell.border = thin_border

        # Apply bold font and border to all header cells (rows 4 and 5)
        for row_idx in range(4, 6):
            for col_idx in range(2, 13): # Columns B to L (index 2 to 12)
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = bold_thai_font
                cell.border = thin_border

        # Set column widths
        ws.column_dimensions['A'].width = 3 # Margin column
        ws.column_dimensions['B'].width = 25 # รายการ
        ws.column_dimensions['C'].width = 15 # ใบรับที่ (รับ)
        ws.column_dimensions['D'].width = 10 # จำนวน (รับ)
        ws.column_dimensions['E'].width = 12 # บาท (รับ)
        ws.column_dimensions['F'].width = 15 # ใบรับที่ (จ่าย)
        ws.column_dimensions['G'].width = 10 # จำนวน (จ่าย)
        ws.column_dimensions['H'].width = 12 # บาท (จ่าย)
        ws.column_dimensions['I'].width = 10 # จำนวน (คงเหลือ) - First pair
        ws.column_dimensions['J'].width = 12 # บาท (คงเหลือ) - First pair
        ws.column_dimensions['K'].width = 10 # จำนวน (คงเหลือ) - Second pair
        ws.column_dimensions['L'].width = 12 # บาท (คงเหลือ) - Second pair

        # --- Data Section (mimicking the image structure) ---
        current_row = 6 # Start data from row 6

        # Fixed header row for 'รับจาก บจก.เอสซีเครื่องคลัง'
        ws.cell(row=current_row, column=2, value="26-มี.ค.-67 รับจาก บจก.เอสซีเครื่องคลัง").font = thai_font
        for col_idx in range(2, 13): # Apply border across the row
            ws.cell(row=current_row, column=col_idx).border = thin_border
        current_row += 1

        # Populate data from data_list
        for idx, item in enumerate(data_list):
            item_name = item[0] # Product Name
            quantity = item[2]  # Quantity (e.g., "1 ถุง")

            ws.cell(row=current_row, column=2, value=f"-{item_name}").font = thai_font
            ws.cell(row=current_row, column=3, value="INV67000267").font = thai_font # Placeholder for ใบรับที่ (รับ)
            ws.cell(row=current_row, column=4, value=quantity).font = thai_font # Quantity for รับ
            ws.cell(row=current_row, column=5, value="3,668.40").font = thai_font # Placeholder for บาท (รับ)

            ws.cell(row=current_row, column=6, value="-").font = thai_font # ใบรับที่ (จ่าย)
            ws.cell(row=current_row, column=7, value="-").font = thai_font # จำนวน (จ่าย)
            ws.cell(row=current_row, column=8, value="-").font = thai_font # บาท (จ่าย)

            ws.cell(row=current_row, column=9, value=quantity).font = thai_font # จำนวน (คงเหลือ) - First pair
            ws.cell(row=current_row, column=10, value="3,668.40").font = thai_font # บาท (คงเหลือ) - First pair
            ws.cell(row=current_row, column=11, value="").font = thai_font # Second คงเหลือ quantity
            ws.cell(row=current_row, column=12, value="").font = thai_font # Second คงเหลือ baht

            # Apply border to data cells
            for col_idx in range(2, 13):
                ws.cell(row=current_row, column=col_idx).border = thin_border
            
            current_row += 1

        # Additional summary row: "จ่ายไป ผศ.ดร.ศิริมา จิราราชะ"
        # This row appears after all items are listed
        ws.cell(row=current_row + 1, column=2, value="จ่ายไป ผศ.ดร.ศิริมา จิราราชะ").font = thai_font
        ws.cell(row=current_row + 1, column=3, value="-").font = thai_font
        ws.cell(row=current_row + 1, column=4, value="-").font = thai_font
        ws.cell(row=current_row + 1, column=5, value="-").font = thai_font
        ws.cell(row=current_row + 1, column=6, value="1 รก.").font = thai_font # Placeholder for 'จ่าย'
        ws.cell(row=current_row + 1, column=7, value="3,668.40").font = thai_font # Placeholder for 'จ่าย'
        ws.cell(row=current_row + 1, column=8, value="").font = thai_font
        ws.cell(row=current_row + 1, column=9, value="").font = thai_font
        ws.cell(row=current_row + 1, column=10, value="").font = thai_font
        ws.cell(row=current_row + 1, column=11, value="").font = thai_font
        ws.cell(row=current_row + 1, column=12, value="").font = thai_font

        # Apply border to this final summary row
        for col_idx in range(2, 13):
            ws.cell(row=current_row + 1, column=col_idx).border = thin_border
        
        # Save the workbook
        wb.save(filename)
        print(f"✅ Excel file '{filename}' created successfully!")
        if platform.system() == "Windows":
            os.startfile(filename)
        return True
    except Exception as e:
        print(f"❌ Error creating Excel file: {e}")
        # Return False to indicate an error to the UI
        return False

# Example usage (for testing this file directly)
if __name__ == '__main__':
    # This block allows you to run excelsummary.py directly for testing.
    # It creates a mock Data class if Backend/Data.py is not in the same directory.
    try:
        from Backend import Data
    except ImportError:
        print("Backend/Data.py not found. Using a mock Data class for testing.")
        class Data:
            def __init__(self):
                self.list = []
            def appendlist(self, name, list3d, amount, date):
                self.list.append((name, list3d, amount, date))
            def sorted(self): # Added sorted method for completeness
                pass # Not needed for this Excel function, but good for mock

    # Create dummy data (can be 1 or multiple items)
    mock_data_instance = Data()
    mock_data_instance.appendlist("ถั่วเขียวเราะเปลือก", "ว.งานบ้านงานครัว", "1 ถุง", "มิ.ย.68")
    mock_data_instance.appendlist("ถั่วแดงหลวง", "ว.งานบ้านงานครัว", "8 ถุง", "")
    mock_data_instance.appendlist("ใบชา", "ว.งานบ้านงานครัว", "2 กล่อง", "")
    # You can comment out lines above to test with only 1 item
    # mock_data_instance.appendlist("สินค้าทดสอบเดียว", "หน่วยงานทดสอบ", "10 ชิ้น", "ก.ค.68")

    create_excel_summary(mock_data_instance.list)
