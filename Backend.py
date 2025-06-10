from datetime import datetime

class Data:
    def __init__(self):
        self.list = []
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")

    def appendlist(self, name, category, amount, date_needed, price):
        # เพิ่ม price เข้าไปใน list item
        self.list.append([name, category, amount, date_needed, price])

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
        for word in text.split():
            if word.isdigit():
                return int(word)
        return 0

    def parse_price(self, text):
        # Function to parse price (assuming it's a string like "15 บาท")
        for word in text.split():
            if word.isdigit():
                return int(word)
        return 0

    def sorted(self):
        # เรียงลำดับตาม amount (index 2) เหมือนเดิม หรือจะเปลี่ยนเป็น price ก็ได้
        self.list.sort(key=lambda row: self.parse_amount(row[2]))

    def summary(self):
        summary_list = []
        for row in self.list:
            name = row[0]
            category = row[1]
            amount_text = row[2]
            price = self.parse_price(row[4]) # ดึงราคามาใช้
            
            amount_parts = amount_text.split()
            if amount_parts and amount_parts[0].isdigit():
                amount = int(amount_parts[0])
                unit = amount_parts[1] if len(amount_parts) > 1 else ""
                
                found = False
                for item in summary_list:
                    if item[1] == category and unit in item[2]:
                        old_amount = int(item[2].split()[0])
                        item[2] = f"{old_amount + amount} {unit}"
                        item[3] += 1
                        item[4] += price # รวมราคา
                        found = True
                        break
                if not found:
                    summary_list.append([name, category, f"{amount} {unit}", 1, price]) # เพิ่มราคาใน summary
        return summary_list

    def remove_item_by_index(self, index):
        if 0 <= index < len(self.list):
            self.list.pop(index)

    def generate_purchase_request(self, purpose="", month_year_needed="", budget_year="",people1="",people2=""):
        categorized_items = {}
        
        for item in self.list:
            category_key = item[1]
            item_name = item[0]
            amount_numeric = self.parse_amount(item[2])
            item_price = self.parse_price(item[4]) # ดึงราคาของแต่ละรายการมาใช้

            if category_key not in categorized_items:
                categorized_items[category_key] = {
                    "count": 0,
                    "total_amount": 0,
                    "total_cost": 0, # เพิ่ม total_cost
                    "items": {}
                }
            
            categorized_items[category_key]["count"] += 1
            categorized_items[category_key]["total_amount"] += amount_numeric
            categorized_items[category_key]["total_cost"] += item_price # รวมราคา
            
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
            cost_phrases.append(f"หมวดค่า {display_category_name} เป็นเงิน {data['total_cost']} บาท") # ใช้ total_cost
        
        output_string += ", ".join(cost_phrases)

        output_string += (
            f" โดยวิธีเฉพาะเจาะจง และขอแต่งตั้งผู้กำหนดคุณลักษณะเฉพาะ คือ {people1} "
            f"โดยขอให้แต่งตั้งผู้ตรวจพัสดุ คือ {people2}"
        )
        return output_string

# ตัวอย่างการใช้งาน:
Datum = Data()
# เพิ่มราคาเป็น parameter สุดท้าย
Datum.appendlist("ถั่วเขียวเราะเปลือก", "ว.งานบ้านงานครัว", "1 ถุง", "มิ.ย.68", "15 บาท")
Datum.appendlist("ถั่วแดงหลวง", "ว.งานบ้านงานครัว", "8 ถุง", "", "120 บาท") # สมมติ 8 ถุง ราคา 120 บาท
Datum.appendlist("ใบชา", "ว.งานบ้านงานครัว", "2 กล่อง", "", "50 บาท")
Datum.appendlist("ถุงใส ขนาด 20x30 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "", "30 บาท")
Datum.appendlist("ถุงตัดตรง LLDPE ขนาด 15x26 นิ้ว", "ว.งานบ้านงานครัว", "2 แพ็ค", "", "40 บาท")
Datum.appendlist("สารเคมี A", "ว.วิทยาศาสตร์หรือการแพทย์", "2 ขวด", "ก.ค.68", "200 บาท")
Datum.appendlist("หลอดทดลอง", "ว.วิทยาศาสตร์หรือการแพทย์", "10 ชิ้น", "", "150 บาท")
Datum.appendlist("ปากกา", "ว.สำนักงาน", "5 ด้าม", "ก.ย.68", "25 บาท")
Datum.appendlist("กระดาษ A4", "ว.สำนักงาน", "2 รีม", "", "200 บาท")

generated_text = Datum.generate_purchase_request(
    purpose="ใช้ในการเรียนการสอนวิชาจุลชีววิทยาและงานธุรการ",
    month_year_needed="กรกฎาคม 2568",
    budget_year="2568",people1="เก",people2="เก"
)

if __name__ == "__main__":
    print(generated_text)