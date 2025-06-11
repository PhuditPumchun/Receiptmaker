# File: Backend.py

from datetime import datetime
from collections import defaultdict
import re

class Data:
    def __init__(self):
        self.list = []
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")

    def appendlist(self, name, category, amount, date_needed, price, received_from, invoice_no, purchase_date):
        """
        ปรับปรุง: เพิ่ม received_from, invoice_no, และ purchase_date เข้าไปใน list
        โครงสร้างข้อมูลใน self.list จะเป็น:
        [0: name, 1: category, 2: amount, 3: date_needed, 4: price, 5: received_from, 6: invoice_no, 7: purchase_date]
        """
        numeric_price = self.parse_price(price)
        self.list.append([name, category, amount, date_needed, numeric_price, received_from, invoice_no, purchase_date])

    def format_thai_date(self, date_obj):
        thai_months = [
            "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ]
        day = date_obj.day
        month = thai_months[date_obj.month]
        year = date_obj.year + 543
        return f"{day}/{month}/{year}"

    def parse_price(self, price_str):
        try:
            return float(price_str)
        except ValueError:
            return 0.0

    def parse_amount(self, amount_str):
        # แก้ไข: ใช้ Regular Expression เพื่อดึงตัวเลขทศนิยมหรือจำนวนเต็ม
        try:
            # ค้นหาเลขชุดแรกในสตริง (รองรับทศนิยม)
            match = re.search(r'(\d+\.?\d*|\.\d+)', amount_str)
            if match:
                return float(match.group(0))
            return 0.0
        except (ValueError, TypeError):
            return 0.0

    def sorted(self):
        # เรียงตามชื่อพัสดุ (item[0])
        self.list.sort(key=lambda x: x[0])

    def remove_item_by_index(self, index):
        if 0 <= index < len(self.list):
            del self.list[index]
            return True
        return False

    def generate_purchase_request(self, purpose, month_year_needed, budget_year, people1, people2):
        # จัดกลุ่มรายการตามหมวดหมู่
        categorized_items = defaultdict(lambda: {'count': 0, 'total_cost': 0.0})
        for item in self.list:
            category = item[1] # item[1] คือ category
            amount_numeric = self.parse_amount(item[2]) # item[2] คือ amount
            price = item[4] # item[4] คือ price
            
            categorized_items[category]['count'] += 1
            categorized_items[category]['total_cost'] += (amount_numeric * price)

        output_string = (
            f"ด้วย ภาควิชาอุตสาหกรรมเกษตร คณะเกษตรศาสตร์ ทรัพยากรธรรมชาติและสิ่งแวดล้อม "
            f"มีความจำเป็นจะต้องขออนุมัติดำเนินการจัดซื้อ "
        )
        category_phrases = []
        for category_key, data in categorized_items.items():
            display_category_name = category_key.replace("ว.", "วัสดุ") # เปลี่ยน "ว." เป็น "วัสดุ" เพื่อการแสดงผล
            category_phrases.append(f"{display_category_name} จำนวน {data['count']} รายการ")
        output_string += ", ".join(category_phrases)
        output_string += (
            f" เพื่อ {purpose} และต้องการใช้สิ่งของดังกล่าวประมาณ(เดือน/ปี) {month_year_needed} "
            f"และเบิกจ่ายจากงบประมาณรายได้ปี {budget_year} กองทุนเพื่อการศึกษา แผนงานจัดการศึกษาอุดมศึกษา "
            f"งานจัดการศึกษาสาขาเกษตรศาสตร์ หลักสูตรวิทยาศาสตร์บันฑิต สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร "
        )
        cost_phrases = []
        for category_key, data in categorized_items.items():
            display_category_name = category_key.replace("ว.", "วัสดุ")
            cost_phrases.append(f"หมวดค่า {display_category_name} เป็นเงิน {data['total_cost']:.2f} บาท")
        output_string += ", ".join(cost_phrases)
        output_string += (
            f" โดยวิธีเฉพาะเจาะจง และขอแต่งตั้งผู้กำหนดคุณลักษณะเฉพาะ {people1} และผู้ตรวจพัสดุ {people2}"
        )
        return output_string