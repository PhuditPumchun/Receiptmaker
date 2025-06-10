import tkinter as tk
from tkinter import ttk, messagebox
from Backend import Data 
from Docxtest import Sleeve1
from excelsummary import create_excel_summary

data = Data()

def add_item():
    name = entry_name.get()
    category = entry_category.get()
    amount = entry_amount.get()
    date = entry_date.get()
    price = entry_price.get()

    if not name or not category or not amount or not date or not price:
        messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบทุกช่อง")
        return

    data.appendlist(name, category, amount, date, price)
    refresh_table()
    clear_fields()

def refresh_table():
    for i in tree.get_children():
        tree.delete(i)
    for row in data.list:
        tree.insert('', 'end', values=row)

def clear_fields():
    entry_name.delete(0, tk.END)
    entry_category.delete(0, tk.END)
    entry_amount.delete(0, tk.END)
    entry_date.delete(0, tk.END)
    entry_price.delete(0, tk.END)

def sort_data():
    data.sorted()
    refresh_table()

def open_create_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("กรอกข้อมูลบันทึกข้อความ")
    dialog.geometry("700x550")

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

    tk.Label(dialog, text="ตัวอย่างเนื้อความ (Generated Body Text - Review Only):").pack(anchor="w", padx=10, pady=(10,0))
    text_body_display = tk.Text(dialog, width=80, height=10, state=tk.DISABLED)
    text_body_display.pack(padx=10, pady=5)

    def update_generated_text_preview():
        purpose = entry_purpose.get().strip()
        month_year_needed = entry_month_year.get().strip()
        budget_year = entry_budget_year.get().strip()
        
        generated_text = data.generate_purchase_request(
            purpose=purpose,
            month_year_needed=month_year_needed,
            budget_year=budget_year
        )
        
        text_body_display.config(state=tk.NORMAL)
        text_body_display.delete("1.0", tk.END)
        text_body_display.insert("1.0", generated_text)
        text_body_display.config(state=tk.DISABLED)

    btn_preview = tk.Button(dialog, text="แสดงตัวอย่างเนื้อความ", command=update_generated_text_preview)
    btn_preview.pack(pady=5)

    def on_create():
        title = entry_title.get().strip()
        runnum = entry_runnumber.get().strip()
        
        purpose = entry_purpose.get().strip()
        month_year_needed = entry_month_year.get().strip()
        budget_year = entry_budget_year.get().strip()

        if not title or not runnum or not purpose or not month_year_needed or not budget_year:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลบันทึกข้อความให้ครบทุกช่อง")
            return
        
        # Generate the body text using the Data class method
        body_text = data.generate_purchase_request(
            purpose=purpose,
            month_year_needed=month_year_needed,
            budget_year=budget_year
        )

        success = Sleeve1(data, title, runnum, body_text)
        if success == 1:
            messagebox.showinfo("สำเร็จ", "สร้างบันทึกข้อความเรียบร้อยแล้ว")
            dialog.destroy()
        else:
            messagebox.showerror("ไม่สำเร็จ", "สร้างบันทึกข้อความไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

    btn_create = tk.Button(dialog, text="สร้างบันทึกข้อความ", command=on_create) # ปุ่มนี้จะเรียก on_create
    btn_create.pack(pady=10)

def create_excel():
    if not data.list:
        messagebox.showwarning("ไม่มีข้อมูล", "กรุณาเพิ่มรายการพัสดุในตารางก่อนสร้างไฟล์ Excel")
        return
    
    success = create_excel_summary(data.list)
    if success:
        messagebox.showinfo("สำเร็จ", "สร้างไฟล์ Excel สรุปยอดเรียบร้อยแล้ว")
    else:
        messagebox.showerror("ไม่สำเร็จ", "สร้างไฟล์ Excel สรุปยอดไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

def clear_all():
    confirm = messagebox.askyesno("ยืนยัน", "ต้องการล้างข้อมูลทั้งหมดหรือไม่?")
    if confirm:
        data.__init__()
        refresh_table()

def delete_selected_item():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("ไม่มีรายการที่เลือก", "กรุณาเลือกรายการที่ต้องการลบ")
        return

    item_index = tree.index(selected[0])
    data.remove_item_by_index(item_index)
    refresh_table()

root = tk.Tk()
root.title("แบบฟอร์มบันทึกพัสดุ")
root.geometry("1000x700")

form_frame = tk.Frame(root)
form_frame.pack(pady=10)

tk.Label(form_frame, text="ชื่อพัสดุ").grid(row=0, column=0, sticky="w", padx=5, pady=2)
entry_name = tk.Entry(form_frame, width=40)
entry_name.grid(row=0, column=1, padx=5, pady=2)

tk.Label(form_frame, text="หมวดหมู่").grid(row=1, column=0, sticky="w", padx=5, pady=2)
entry_category = tk.Entry(form_frame, width=40)
entry_category.grid(row=1, column=1, padx=5, pady=2)

tk.Label(form_frame, text="จำนวน (ใส่หน่วยด้วย)").grid(row=2, column=0, sticky="w", padx=5, pady=2)
entry_amount = tk.Entry(form_frame, width=40)
entry_amount.grid(row=2, column=1, padx=5, pady=2)

tk.Label(form_frame, text="วันที่ต้องการใช้ (เช่น 5 มิ.ย. 2568)").grid(row=3, column=0, sticky="w", padx=5, pady=2)
entry_date = tk.Entry(form_frame, width=40)
entry_date.grid(row=3, column=1, padx=5, pady=2)

tk.Label(form_frame, text="ราคา (เช่น 15 บาท)").grid(row=4, column=0, sticky="w", padx=5, pady=2)
entry_price = tk.Entry(form_frame, width=40)
entry_price.grid(row=4, column=1, padx=5, pady=2)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="เพิ่มรายการ", width=20, command=add_item).grid(row=0, column=0, padx=5, pady=5)
tk.Button(button_frame, text="เรียงตามจำนวน", width=20, command=sort_data).grid(row=0, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ลบรายการที่เลือก", width=20, command=delete_selected_item).grid(row=0, column=2, padx=5, pady=5)

tk.Button(button_frame, text="สร้างบันทึกข้อความ", width=20, command=open_create_dialog).grid(row=1, column=0, padx=5, pady=5)
tk.Button(button_frame, text="สร้าง Excel สรุปยอด", width=20, command=create_excel).grid(row=1, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ล้างข้อมูลทั้งหมด", width=20, fg="red", command=clear_all).grid(row=1, column=2, pady=5)

columns = ("ชื่อพัสดุ", "หมวดหมู่", "จำนวน", "วันที่ต้องการใช้", "ราคา")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)

tree.pack(fill="both", expand=True, padx=10, pady=10)

refresh_table()

root.mainloop()