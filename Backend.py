from datetime import datetime

class Data:
    def __init__(self):
        self.list = []
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")

    def appendlist(self, name, category, amount, date_needed):
        self.list.append([name, category, amount, date_needed])

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

    def sorted(self):
        self.list.sort(key=lambda row: self.parse_amount(row[2]))  # ใช้ amount ที่ index 2

    def summary(self):
        summary_list = []
        for row in self.list:
            name = row[0]
            category = row[1]
            amount_text = row[2]
            amount_parts = amount_text.split()
            if amount_parts and amount_parts[0].isdigit():
                amount = int(amount_parts[0])
                unit = amount_parts[1] if len(amount_parts) > 1 else ""
                found = False
                for item in summary_list:
                    if item[0] == name and unit in item[2]:
                        old_amount = int(item[2].split()[0])
                        item[2] = f"{old_amount + amount} {unit}"
                        found = True
                        break
                if not found:
                    summary_list.append([name, category, f"{amount} {unit}"])
        return summary_list

    def remove_item_by_index(self, index):
        if 0 <= index < len(self.list):
            self.list.pop(index)
