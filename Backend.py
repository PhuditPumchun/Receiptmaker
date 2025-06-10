# File: backend.py

from datetime import datetime

class Data:
    def __init__(self):
        self.list = []
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")

    def appendlist(self, name, category, amount, date_needed, price, received_from, invoice_no):
        """
        ปรับปรุง: เพิ่ม received_from และ invoice_no เข้าไปใน list
        โครงสร้างข้อมูลใน self.list จะเป็น:
        [0: name, 1: category, 2: amount, 3: date_needed, 4: price, 5: received_from, 6: invoice_no]
        """
        numeric_price = self.parse_price(price)
        self.list.append([name, category, amount, date_needed, numeric_price, received_from, invoice_no])

    def format_thai_date(self, date_obj):
        thai_months = [
            "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ]
        day = date_obj.day
        month = thai_months[date_obj.month]
        year = date_obj.year + 543
        return f"{day} {month} {year}"

    def parse_amount(self, text):
        parts = str(text).split()
        for part in parts:
            if part.isdigit():
                return int(part)
        return 0

    def parse_price(self, text):
        text = str(text)
        for word in text.split():
            try:
                return float(word)
            except (ValueError, TypeError):
                continue
        return 0.0

    def sorted(self):
        self.list.sort(key=lambda row: self.parse_amount(row[2]))

    def generate_purchase_request(self, purpose="", month_year_needed="", budget_year="", people1="", people2=""):
        categorized_items = {}
        
        for item in self.list:
            category_key = item[1]
            item_name = item[0]
            amount_numeric = self.parse_amount(item[2])
            # index ของราคาคือ 4 และเป็นตัวเลขแล้ว
            item_price = item[4] 

            if category_key not in categorized_items:
                categorized_items[category_key] = {
                    "count": 0,
                    "total_amount": 0,
                    "total_cost": 0,
                    "items": {}
                }
            
            categorized_items[category_key]["count"] += 1
            categorized_items[category_key]["total_amount"] += amount_numeric
            categorized_items[category_key]["total_cost"] += item_price 
            
            if item_name not in categorized_items[category_key]["items"]:
                categorized_items[category_key]["items"][item_name] = amount_numeric
            else:
                categorized_items[category_key]["items"][item_name] += amount_numeric

        output_string = (
            f"\t\t\tด้วย ภาควิชาอุตสาหกรรมการเกษตร คณะเกษตรศาสตร์ ทรัพยากรธรรมชาติและสิ่งแวดล้อม "
            f"มีความจำเป็นจะต้องขออนุมัติดำเนินการจัดซื้อ "
        )
        category_phrases = []
        for category_key, data in categorized_items.items():
            display_category_name = category_key.replace("ว.", "วัสดุ")
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
            f" โดยวิธีเฉพาะเจาะจง และขอแต่งตั้งผู้กำหนดคุณลักษณะเฉพาะ คือ {people1} "
            f"โดยขอให้แต่งตั้งผู้ตรวจพัสดุ คือ {people2}"
        )
        return output_string

    def remove_item_by_index(self, index):
        if 0 <= index < len(self.list):
            self.list.pop(index)