import tkinter as tk
from tkinter import ttk, messagebox
from Backend import Data 
from Sleeve import Sleeve1, summarySleeve

data = Data()

def add_item():
    name = entry_name.get()
    list3d = entry_list3d.get()
    amount = entry_amount.get()
    date = entry_date.get()

    if not name or not list3d or not amount or not date:
        messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบทุกช่อง")
        return

    data.appendlist(name, list3d, amount, date)
    refresh_table()
    clear_fields()

def refresh_table():
    for i in tree.get_children():
        tree.delete(i)
    for row in data.list[1:]:
        tree.insert('', 'end', values=row)

def clear_fields():
    entry_name.delete(0, tk.END)
    entry_list3d.delete(0, tk.END)
    entry_amount.delete(0, tk.END)
    entry_date.delete(0, tk.END)

def sort_data():
    data.sorted()
    refresh_table()

def run_sleeve1():
    if Sleeve1(data) == 1:
        messagebox.showinfo("สำเร็จ", "สร้างบันทึกข้อความเรียบร้อยแล้ว")
    else:
        messagebox.showinfo("ไม่สำเร็จ", "สร้างบันทึกข้อความไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

def run_summary():
    if summarySleeve(data) == 1:
        messagebox.showinfo("สำเร็จ", "สร้างสรุปเรียบร้อยแล้ว")
    else:
        messagebox.showinfo("ไม่สำเร็จ", "สร้างสรุปไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

def clear_all():
    confirm = messagebox.askyesno("ยืนยัน", "ต้องการล้างข้อมูลทั้งหมดหรือไม่?")
    if confirm:
        data.__init__()  # รีเซ็ตข้อมูลใหม่
        refresh_table()

def delete_selected_item():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("ไม่มีรายการที่เลือก", "กรุณาเลือกรายการที่ต้องการลบ")
        return

    item = tree.item(selected[0])
    index_to_delete = int(item["values"][0])  # ลำดับเป็นค่าแรก
    data.remove_item_by_index(index_to_delete)
    refresh_table()

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("แบบฟอร์มบันทึกพัสดุ")
root.geometry("850x650")

# ===== ฟอร์มกรอกข้อมูล =====
form_frame = tk.Frame(root)
form_frame.pack(pady=10)

tk.Label(form_frame, text="ชื่อพัสดุ").grid(row=0, column=0, sticky="w")
entry_name = tk.Entry(form_frame, width=40)
entry_name.grid(row=0, column=1)

tk.Label(form_frame, text="หน่วยงาน").grid(row=1, column=0, sticky="w")
entry_list3d = tk.Entry(form_frame, width=40)
entry_list3d.grid(row=1, column=1)

tk.Label(form_frame, text="จำนวน (ใส่หน่วยด้วย)").grid(row=2, column=0, sticky="w")
entry_amount = tk.Entry(form_frame, width=40)
entry_amount.grid(row=2, column=1)

tk.Label(form_frame, text="วันที่ต้องการใช้ (เช่น 5 มิ.ย. 2568)").grid(row=3, column=0, sticky="w")
entry_date = tk.Entry(form_frame, width=40)
entry_date.grid(row=3, column=1)

# ===== ปุ่มการทำงาน =====
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="เพิ่มรายการ", width=20, command=add_item).grid(row=0, column=0, padx=5, pady=5)
tk.Button(button_frame, text="เรียงตามจำนวน", width=20, command=sort_data).grid(row=0, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ลบรายการที่เลือก", width=20, command=delete_selected_item).grid(row=0, column=2, padx=5, pady=5)

tk.Button(button_frame, text="สร้างบันทึกข้อความ", width=20, command=run_sleeve1).grid(row=1, column=0, padx=5, pady=5)
tk.Button(button_frame, text="สร้างสรุป", width=20, command=run_summary).grid(row=1, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ล้างข้อมูลทั้งหมด", width=20, fg="red", command=clear_all).grid(row=1, column=2, pady=5)

# ===== ตารางแสดงรายการ =====
columns = ("ลำดับ", "ชื่อพัสดุ", "บัญชี 3 มิติ", "จำนวน", "วันที่ต้องการใช้")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)

tree.pack(fill="both", expand=True, padx=10, pady=10)

root.mainloop()
