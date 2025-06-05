from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from datetime import datetime

class Data:
    def __init__(self):
        self.list = [
        ]
        
        self.day = self.format_thai_date(datetime.today())
        self.time = datetime.now().strftime("%H:%M")
        self.n = 0

    def appendlist(self, name, list3d, amount, date):
        self.n += 1
        self.list.append([str(self.n), name, list3d, amount, date])

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
        sorted_data = sorted(data_only, key=lambda row: self.parse_amount(row[3]))

        for i in range(len(sorted_data)):
            self.list[i + 1] = [str(i + 1)] + sorted_data[i][1:]

    def summary(self):
        self.data_only = self.list[1:]
        self.data_summary = []

        for row in self.data_only:
            found = False
            name = row[1]
            amount_text = row[3]
            amount_parts = amount_text.split()
            if len(amount_parts) >= 1 and amount_parts[0].isdigit():
                amount = int(amount_parts[0])
                unit = amount_parts[1] if len(amount_parts) > 1 else ""
                for item in self.data_summary:
                    if item[1] == name and unit in item[3]:
                        old_amount = int(item[3].split()[0])
                        item[3] = f"{old_amount + amount} {unit}"
                        found = True
                        break
                if not found:
                    self.data_summary.append(
                        [str(len(self.data_summary) + 1), name, row[2], f"{amount} {unit}"]
                    )
        return self.data_summary
    
    def remove_item_by_index(self, index):
        if 1 <= index < len(self.list):
            self.list.pop(index)
            for i, row in enumerate(self.list[1:], start=1):
                row[0] = i
