from datetime import datetime

class Data:
    def __init__(self):
        self.list = [
            ['ลำดับ', 'รายการขอซื้อจ้าง\n[ขั้นตอนตามระเบียบข้อ 22(2)]',
             'รายการแยกวัสดุตาม\nระบบบัญชี\n3มิติ',
             'จำนวนหน่วย',
             'กำหนดเวลาที่\nต้องการใช้พัสดุ\n[ตามระเบียบข้อ\n22 (5)]']
        ]
        
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")
        self.n = 0

    def appendlist(self, name, list3d, amount, date):
        self.n += 1
        self.list.append([str(self.n), name, list3d, amount, date])
        # self.sorted()

    def format_thai_date(self, date_obj):
        thai_months = [
            "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ]
        day = date_obj.day
        month = thai_months[date_obj.month]
        year = date_obj.year + 543
        return f"วันที่ {day} {month} {year}"
    
    def parse_amount(self,text):
            for word in text.split():
                if word.isdigit():
                    return int(word)
            return 0

    def sorted(self):
        data_only = self.list[1:]
        # เรียงตามจำนวนหน่วย
        sorted_data = sorted(data_only, key=lambda row: self.parse_amount(row[3]))

        # เขียนใหม่ทั้งหมด
        for i in range(len(sorted_data)):
            self.list[i + 1] = [str(i + 1)] + sorted_data[i][1:]

    def summary(self):
        self.data_only = self.list[1:]
        self.data_summary = []

        for row in self.data_only:
            found = False
            name = row[1]
            amount_text = row[3]

            # แยกตัวเลขและหน่วย เช่น "10 รีม"
            amount_parts = amount_text.split()
            if len(amount_parts) >= 1 and amount_parts[0].isdigit():
                amount = int(amount_parts[0])
                unit = amount_parts[1] if len(amount_parts) > 1 else ""

                # ตรวจว่ามีอยู่แล้วใน data_summary หรือยัง
                for item in self.data_summary:
                    if item[1] == name and unit in item[3]:
                        # ถ้ามีแล้วให้บวกจำนวน
                        old_amount = int(item[3].split()[0])
                        item[3] = f"{old_amount + amount} {unit}"
                        found = True
                        break

                # ถ้ายังไม่เคยเจอชื่อซ้ำ ให้เพิ่มใหม่
                if not found:
                    self.data_summary.append(
                        [str(len(self.data_summary) + 1), name, row[2], f"{amount} {unit}"]
                    )

        # แสดงผลสรุป
        # print("\nสรุปรายการรวม:")
        # for item in self.data_summary:
        #     print(item)
        return self.data_summary







# ทดสอบการใช้งาน
data = Data()

data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "5 มิ.ย. 2568")
data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "10 รีม", "6 มิ.ย. 2568")
data.appendlist("เครื่องพิมพ์", "วัสดุสำนักงาน", "2 เครื่อง", "7 มิ.ย. 2568")
data.appendlist("โต๊ะทำงาน", "วัสดุสำนักงาน", "1 ตัว", "8 มิ.ย. 2568")
data.appendlist("แฟ้มเอกสาร", "วัสดุสำนักงาน", "10 เล่ม", "9 มิ.ย. 2568")
data.appendlist("กระดาษ A4", "วัสดุสำนักงาน", "5 รีม", "10 มิ.ย. 2568")

print("ก่อนเรียง:")
for row in data.list:
    print(row)

data.sorted()

print("\nหลังเรียง:")
for row in data.list:
    print(row)

print("\n สรุป")
for row in data.summary():
    print(row)
