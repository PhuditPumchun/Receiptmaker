# File: ui.py

import tkinter as tk
from tkinter import ttk, messagebox
from Backend import Data
from Docxtest import Sleeve1 # สมมติว่ามีไฟล์นี้อยู่
from excelsummary import create_excel_summary
from datetime import datetime
import json
import os

DATA_FILE = "saved_data.json" # ชื่อไฟล์สำหรับเก็บข้อมูล

def save_data_to_file():
    """บันทึกข้อมูลใน data.list ลงในไฟล์ JSON"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data.list, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error saving data: {e}")

def load_data_from_file():
    """อ่านข้อมูลจากไฟล์ JSON ถ้ามี"""
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading data: {e}")
    return [] # ถ้าไฟล์ไม่มีหรือมีปัญหา ให้คืนค่าเป็นลิสต์ว่าง

# --- เพิ่มฟังก์ชันสำหรับปุ่ม Save/Load Manual ---
def manual_save_data():
    """Manual save button command."""
    save_data_to_file()
    messagebox.showinfo("บันทึกข้อมูล", "บันทึกข้อมูลสำเร็จแล้ว")

def manual_load_data():
    """Manual load button command."""
    loaded_data = load_data_from_file()
    if loaded_data is not None:
        data.list = loaded_data
        refresh_table()
        messagebox.showinfo("โหลดข้อมูล", "โหลดข้อมูลสำเร็จแล้ว")
    else:
        messagebox.showwarning("โหลดข้อมูล", "ไม่พบข้อมูลที่บันทึกไว้ หรือเกิดข้อผิดพลาดในการโหลด")

# --- ส่วนของฟังก์ชันจัดการข้อมูลและ UI ---

def add_item():
    # ดึงข้อมูลจากช่องกรอกทั้งหมด
    name = entry_name.get()
    category = entry_category.get()
    amount = entry_amount.get()
    date_needed = entry_date.get()
    price = entry_price.get()
    received_from = entry_received_from.get()
    invoice_no = entry_invoice_no.get()
    purchase_date = entry_purchase_date.get()

    # ตรวจสอบว่ากรอกข้อมูลสำคัญครบหรือไม่
    if not all([name, category, amount, price, received_from, invoice_no, purchase_date]):
        messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบถ้วน\n(ชื่อ, หมวดหมู่, จำนวน, ราคา, รับจาก, ใบรับที่, วันที่ซื้อ)")
        return

    # เรียกใช้ appendlist ที่มี parameters ครบ
    data.appendlist(name, category, amount, date_needed, price, received_from, invoice_no, purchase_date)
    refresh_table()
    clear_fields()
    save_data_to_file() # <-- บันทึกข้อมูลทุกครั้งที่เพิ่มรายการใหม่

def refresh_table():
    for i in tree.get_children():
        tree.delete(i)
    for row in data.list:
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
    entry_purchase_date.delete(0, tk.END)

def sort_data():
    data.sorted()
    refresh_table()
    # ไม่ต้อง save ที่นี่ เพราะการเรียงลำดับไม่ได้เปลี่ยนเนื้อหาข้อมูล

def delete_selected_item():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("ไม่มีรายการที่เลือก", "กรุณาเลือกรายการที่ต้องการลบ")
        return
    item_index = tree.index(selected[0])
    data.remove_item_by_index(item_index)
    refresh_table()
    save_data_to_file() # <-- บันทึกข้อมูลทุกครั้งที่ลบรายการ

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


def open_excel_dialog():
    if not data.list:
        messagebox.showwarning("ไม่มีข้อมูล", "กรุณาเพิ่มรายการพัสดุในตารางก่อนสร้างไฟล์ Excel")
        return

    excel_dialog = tk.Toplevel(root)
    excel_dialog.title("กรอกข้อมูลสำหรับ Excel สรุปยอด")
    excel_dialog.geometry("400x250")

    tk.Label(excel_dialog, text="จ่ายให้:").pack(anchor="w", padx=10, pady=(10,0))
    entry_paid_to = tk.Entry(excel_dialog, width=50)
    entry_paid_to.insert(0, "ผศ.ดร.ศิริมา จิราราชะ")
    entry_paid_to.pack(padx=10, pady=5)

    def on_create_excel():
        paid_to = entry_paid_to.get().strip()

        if not all([paid_to]):
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลสำหรับ Excel ให้ครบทุกช่อง")
            return

        transaction_info = {
            "paid_to": paid_to
        }

        success = create_excel_summary(data.list, transaction_info)
        if success:
            messagebox.showinfo("สำเร็จ", "สร้างไฟล์ Excel สรุปยอดเรียบร้อยแล้ว")
            excel_dialog.destroy()
        else:
            messagebox.showerror("ไม่สำเร็จ", "สร้างไฟล์ Excel สรุปยอดไม่สำเร็จ กรุณาตรวจสอบข้อผิดพลาด")

    btn_create_excel = tk.Button(excel_dialog, text="สร้าง Excel", command=on_create_excel)
    btn_create_excel.pack(pady=10)


def clear_all():
    if messagebox.askyesno("ยืนยัน", "ต้องการล้างข้อมูลทั้งหมดหรือไม่?"):
        data.__init__()
        refresh_table()
        save_data_to_file() # <-- บันทึกข้อมูล (ซึ่งตอนนี้เป็นลิสต์ว่าง)

# --- GUI Setup ---
root = tk.Tk()
root.title("แบบฟอร์มบันทึกพัสดุ")
root.geometry("1400x700")

# --- โหลดข้อมูลตอนเริ่มต้นโปรแกรม ---
data = Data() # สร้าง instance ของ Data ก่อน
data.list = load_data_from_file() # โหลดข้อมูลจากไฟล์มาใส่ใน list ของ data

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

tk.Label(form_frame, text="ใบรับที่ (Invoice No.)").grid(row=3, column=0, sticky="w", padx=5, pady=2)
entry_invoice_no = tk.Entry(form_frame, width=40)
entry_invoice_no.grid(row=3, column=1, padx=5, pady=2)

tk.Label(form_frame, text="วันที่ซื้อ (Purchase Date)").grid(row=3, column=2, sticky="w", padx=5, pady=2)
entry_purchase_date = tk.Entry(form_frame, width=40)
entry_purchase_date.grid(row=3, column=3, padx=5, pady=2)

# --- Buttons ---
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="เพิ่มรายการ", width=20, command=add_item).grid(row=0, column=0, padx=5, pady=5)
tk.Button(button_frame, text="เรียงตามชื่อ", width=20, command=sort_data).grid(row=0, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ลบรายการที่เลือก", width=20, command=delete_selected_item).grid(row=0, column=2, padx=5, pady=5)

tk.Button(button_frame, text="สร้างบันทึกข้อความ", width=20, command=open_create_dialog).grid(row=1, column=0, padx=5, pady=5)
tk.Button(button_frame, text="สร้าง Excel สรุปยอด", width=20, command=open_excel_dialog).grid(row=1, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ล้างข้อมูลทั้งหมด", width=20, fg="red", command=clear_all).grid(row=1, column=2, pady=5)

# Add new Save and Load buttons
tk.Button(button_frame, text="บันทึกข้อมูล (Manual)", width=20, command=manual_save_data).grid(row=2, column=0, padx=5, pady=5)
tk.Button(button_frame, text="โหลดข้อมูล (Manual)", width=20, command=manual_load_data).grid(row=2, column=1, padx=5, pady=5)

# --- Treeview (Table Display) ---
columns = ("ชื่อพัสดุ", "หมวดหมู่", "จำนวน", "วันที่ต้องการใช้", "ราคา (บาท)", "รับจาก", "ใบรับที่", "วันที่ซื้อ")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="w", width=120)

tree.column("ชื่อพัสดุ", width=200, anchor="w")
tree.column("จำนวน", width=80, anchor="center")
tree.column("ราคา (บาท)", width=100, anchor="e")
tree.column("หมวดหมู่", width=150, anchor="w")
tree.column("ใบรับที่", width=100, anchor="center")
tree.column("วันที่ซื้อ", width=100, anchor="center")

tree.pack(fill="both", expand=True, padx=10, pady=10)

refresh_table() # <-- แสดงข้อมูลที่โหลดมาทันทีเมื่อเปิดโปรแกรม
root.mainloop()