# File: ui.py (Updated)

import tkinter as tk
from tkinter import ttk, messagebox
from Backend import Data 
from Docxtest import Sleeve1 # สมมติว่ามีไฟล์นี้อยู่
from excelsummary import create_excel_summary

data = Data()

def add_item():
    # ดึงข้อมูลจากช่องกรอกทั้งหมด
    name = entry_name.get()
    category = entry_category.get()
    amount = entry_amount.get()
    date = entry_date.get()
    price = entry_price.get()
    received_from = entry_received_from.get()
    invoice_no = entry_invoice_no.get()

    # ตรวจสอบว่ากรอกข้อมูลสำคัญครบหรือไม่
    if not all([name, category, amount, price, received_from, invoice_no]):
        messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบถ้วน\n(ชื่อ, หมวดหมู่, จำนวน, ราคา, รับจาก, ใบรับที่)")
        return

    # เรียกใช้ appendlist ที่มี parameters ครบ
    data.appendlist(name, category, amount, date, price, received_from, invoice_no)
    refresh_table()
    clear_fields()

def refresh_table():
    for i in tree.get_children():
        tree.delete(i)
    for row in data.list:
        # จัดรูปแบบราคาให้มีทศนิยม 2 ตำแหน่งตอนแสดงผล
        row_display = list(row)
        if isinstance(row_display[4], (int, float)):
             row_display[4] = f"{row_display[4]:.2f}"
        tree.insert('', 'end', values=tuple(row_display))


def clear_fields():
    entry_name.delete(0, tk.END)
    entry_category.delete(0, tk.END)
    entry_amount.delete(0, tk.END)
    entry_date.delete(0, tk.END)
    entry_price.delete(0, tk.END)
    entry_received_from.delete(0, tk.END)
    entry_invoice_no.delete(0, tk.END)

def sort_data():
    data.sorted()
    refresh_table()

def delete_selected_item():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("ไม่มีรายการที่เลือก", "กรุณาเลือกรายการที่ต้องการลบ")
        return
    item_index = tree.index(selected[0])
    data.remove_item_by_index(item_index)
    refresh_table()

def open_create_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("กรอกข้อมูลบันทึกข้อความ")
    dialog.geometry("700x700")

    tk.Label(dialog, text="เรื่อง (Title):").pack(anchor="w", padx=10, pady=(10,0))
    entry_title = tk.Entry(dialog, width=80)
    entry_title.pack(padx=10, pady=5)

    tk.Label(dialog, text="ที่ (Run Number):").pack(anchor="w", padx=10, pady=(10,0))
    entry_runnumber = tk.Entry(dialog, width=80)
    entry_runnumber.pack(padx=10, pady=5)

    tk.Label(dialog, text="วัตถุประสงค์ (Purpose):").pack(anchor="w", padx=10, pady=(10,0))
    entry_purpose = tk.Entry(dialog, width=80)
    entry_purpose.pack(padx=10, pady=5)

    tk.Label(dialog, text="ประมาณเดือน/ปี ที่ต้องการใช้ (Month/Year Needed):").pack(anchor="w", padx=10, pady=(10,0))
    entry_month_year = tk.Entry(dialog, width=80)
    entry_month_year.pack(padx=10, pady=5)

    tk.Label(dialog, text="งบประมาณปี (Budget Year):").pack(anchor="w", padx=10, pady=(10,0))
    entry_budget_year = tk.Entry(dialog, width=80)
    entry_budget_year.pack(padx=10, pady=5)

    tk.Label(dialog, text="ผู้กำหนดคุณลักษณะเฉพาะ (Person 1):").pack(anchor="w", padx=10, pady=(10,0))
    entry_people1 = tk.Entry(dialog, width=80)
    entry_people1.insert(0, "รศ. ดร. ทิพวรรณ ทองสุข")
    entry_people1.pack(padx=10, pady=5)

    tk.Label(dialog, text="ผู้ตรวจพัสดุ (Person 2):").pack(anchor="w", padx=10, pady=(10,0))
    entry_people2 = tk.Entry(dialog, width=80)
    entry_people2.insert(0, "รศ.กมลวรรณ โรจน์สุนทรกิตติ")
    entry_people2.pack(padx=10, pady=5)

    tk.Label(dialog, text="ตัวอย่างเนื้อความ (Generated Body Text - Review Only):").pack(anchor="w", padx=10, pady=(10,0))
    text_body_display = tk.Text(dialog, width=80, height=10, state=tk.DISABLED, font=('TH Sarabun New', 12))
    text_body_display.pack(padx=10, pady=5)

    def update_generated_text_preview():
        purpose = entry_purpose.get().strip()
        month_year_needed = entry_month_year.get().strip()
        budget_year = entry_budget_year.get().strip()
        people1 = entry_people1.get().strip()
        people2 = entry_people2.get().strip()
        
        generated_text = data.generate_purchase_request(
            purpose=purpose,
            month_year_needed=month_year_needed,
            budget_year=budget_year,
            people1=people1, 
            people2=people2  
        )
        
        text_body_display.config(state=tk.NORMAL)
        text_body_display.delete("1.0", tk.END)
        text_body_display.insert("1.0", generated_text)
        text_body_display.config(state=tk.DISABLED)

    button_dialog_frame = tk.Frame(dialog)
    button_dialog_frame.pack(pady=10)

    btn_preview = tk.Button(button_dialog_frame, text="แสดงตัวอย่างเนื้อความ", command=update_generated_text_preview)
    btn_preview.pack(side=tk.LEFT, padx=5)

    def on_create():
        title = entry_title.get().strip()
        runnum = entry_runnumber.get().strip()
        purpose = entry_purpose.get().strip()
        month_year_needed = entry_month_year.get().strip()
        budget_year = entry_budget_year.get().strip()
        people1 = entry_people1.get().strip()
        people2 = entry_people2.get().strip()

        if not all([title, runnum, purpose, month_year_needed, budget_year, people1, people2]):
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลบันทึกข้อความให้ครบทุกช่อง")
            return
        
        body_text = data.generate_purchase_request(
            purpose=purpose,
            month_year_needed=month_year_needed,
            budget_year=budget_year,
            people1=people1, 
            people2=people2  
        )

        success = Sleeve1(data, title, runnum, body_text)
        if success:
            messagebox.showinfo("สำเร็จ", "สร้างบันทึกข้อความเรียบร้อยแล้ว")
            dialog.destroy()
        else:
            messagebox.showerror("ไม่สำเร็จ", "สร้างบันทึกข้อความไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

    btn_create = tk.Button(button_dialog_frame, text="สร้างบันทึกข้อความ", command=on_create)
    btn_create.pack(side=tk.LEFT, padx=5)

def create_excel_directly():
    """
    *** ฟังก์ชันใหม่: สร้างไฟล์ Excel ทันทีโดยไม่มีหน้าต่างถามข้อมูล ***
    """
    if not data.list:
        messagebox.showwarning("ไม่มีข้อมูล", "กรุณาเพิ่มรายการพัสดุในตารางก่อนสร้างไฟล์ Excel")
        return

    try:
        # สร้างข้อมูลสรุปโดยอัตโนมัติ
        transaction_info = {
            "receipt_date": data.day,              # ใช้วันที่ปัจจุบัน
            "received_from": data.list[0][5],      # ใช้ข้อมูล "รับจาก" ของรายการแรก
            "paid_to": "ผศ.ดร.ศิริมา จิราราชะ"       # ใช้ค่าเริ่มต้น
        }
    except IndexError:
        messagebox.showerror("เกิดข้อผิดพลาด", "ไม่สามารถดึงข้อมูล 'รับจาก' ได้\nกรุณาตรวจสอบว่ามีรายการในตารางอย่างน้อย 1 รายการ")
        return

    # เรียกใช้ฟังก์ชันสร้าง Excel
    success = create_excel_summary(data.list, transaction_info)
    if success:
        messagebox.showinfo("สำเร็จ", "สร้างไฟล์ Excel สรุปยอดเรียบร้อยแล้ว")
    else:
        messagebox.showerror("ไม่สำเร็จ", "สร้างไฟล์ Excel สรุปยอดไม่สำเร็จ กรุณาตรวจสอบข้อผิดพลาด")


def clear_all():
    if messagebox.askyesno("ยืนยัน", "ต้องการล้างข้อมูลทั้งหมดหรือไม่?"):
        data.__init__()
        refresh_table()

# --- GUI Setup ---
root = tk.Tk()
root.title("แบบฟอร์มบันทึกพัสดุ")
root.geometry("1200x700")

form_frame = tk.Frame(root)
form_frame.pack(pady=10)

# --- Input Fields ---
tk.Label(form_frame, text="ชื่อพัสดุ").grid(row=0, column=0, sticky="w", padx=5, pady=2)
entry_name = tk.Entry(form_frame, width=40)
entry_name.grid(row=0, column=1, padx=5, pady=2)

tk.Label(form_frame, text="หมวดหมู่").grid(row=1, column=0, sticky="w", padx=5, pady=2)
entry_category = tk.Entry(form_frame, width=40)
entry_category.grid(row=1, column=1, padx=5, pady=2)

tk.Label(form_frame, text="จำนวน (ใส่หน่วยด้วย)").grid(row=2, column=0, sticky="w", padx=5, pady=2)
entry_amount = tk.Entry(form_frame, width=40)
entry_amount.grid(row=2, column=1, padx=5, pady=2)

tk.Label(form_frame, text="วันที่ต้องการใช้ (ถ้ามี)").grid(row=0, column=2, sticky="w", padx=5, pady=2)
entry_date = tk.Entry(form_frame, width=40)
entry_date.grid(row=0, column=3, padx=5, pady=2)

tk.Label(form_frame, text="ราคา (บาท)").grid(row=1, column=2, sticky="w", padx=5, pady=2)
entry_price = tk.Entry(form_frame, width=40)
entry_price.grid(row=1, column=3, padx=5, pady=2)

tk.Label(form_frame, text="รับจากใคร").grid(row=2, column=2, sticky="w", padx=5, pady=2)
entry_received_from = tk.Entry(form_frame, width=40)
entry_received_from.grid(row=2, column=3, padx=5, pady=2)

tk.Label(form_frame, text="ใบรับที่ (Invoice No.)").grid(row=3, column=2, sticky="w", padx=5, pady=2)
entry_invoice_no = tk.Entry(form_frame, width=40)
entry_invoice_no.grid(row=3, column=3, padx=5, pady=2)

# --- Buttons ---
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="เพิ่มรายการ", width=20, command=add_item).grid(row=0, column=0, padx=5, pady=5)
tk.Button(button_frame, text="เรียงตามจำนวน", width=20, command=sort_data).grid(row=0, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ลบรายการที่เลือก", width=20, command=delete_selected_item).grid(row=0, column=2, padx=5, pady=5)

tk.Button(button_frame, text="สร้างบันทึกข้อความ", width=20, command=open_create_dialog).grid(row=1, column=0, padx=5, pady=5)
# *** แก้ไข command ของปุ่มนี้ ให้เรียกใช้ฟังก์ชันใหม่ ***
tk.Button(button_frame, text="สร้าง Excel สรุปยอด", width=20, command=create_excel_directly).grid(row=1, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ล้างข้อมูลทั้งหมด", width=20, fg="red", command=clear_all).grid(row=1, column=2, pady=5)

# --- Treeview (Table Display) ---
columns = ("ชื่อพัสดุ", "หมวดหมู่", "จำนวน", "วันที่ต้องการใช้", "ราคา (บาท)", "รับจาก", "ใบรับที่")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="w", width=150)

tree.column("ชื่อพัสดุ", width=250, anchor="w")
tree.column("จำนวน", width=100, anchor="center")
tree.column("ราคา (บาท)", width=100, anchor="e") # e = align right
tree.column("หมวดหมู่", width=180, anchor="w")
tree.column("ใบรับที่", width=120, anchor="center")

tree.pack(fill="both", expand=True, padx=10, pady=10)

refresh_table()
root.mainloop()