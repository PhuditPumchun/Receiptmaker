import tkinter as tk
from tkinter import ttk, messagebox
from Backend import Data 
from Docxtest import Sleeve1
from excelsummary import create_excel_summary # Import the new function

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
    for row in data.list:
        tree.insert('', 'end', values=row)

def clear_fields():
    entry_name.delete(0, tk.END)
    entry_list3d.delete(0, tk.END)
    entry_amount.delete(0, tk.END)
    entry_date.delete(0, tk.END)

def sort_data():
    data.sorted()
    refresh_table()

def open_create_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("กรอกข้อมูลบันทึกข้อความ")
    dialog.geometry("500x400")

    tk.Label(dialog, text="เรื่อง (Title):").pack(anchor="w", padx=10, pady=(10,0))
    entry_title = tk.Entry(dialog, width=60)
    entry_title.pack(padx=10, pady=5)

    tk.Label(dialog, text="ที่ (Run Number):").pack(anchor="w", padx=10, pady=(10,0))
    entry_runnumber = tk.Entry(dialog, width=60)
    entry_runnumber.pack(padx=10, pady=5)

    tk.Label(dialog, text="เนื้อความ (Body Text):").pack(anchor="w", padx=10, pady=(10,0))
    text_body = tk.Text(dialog, width=60, height=10)
    text_body.pack(padx=10, pady=5)

    example_text = (
        "              ด้วยข้าพเจ้า รศ.ดร.ทิพวรรณ ทองสุข มีความจำเป็นที่จะขออนุมัติจัดซื้อวัสดุงานบ้านงานครัว จำนวน 11 รายการ\n"
        "เพื่อใช้ในการทดลองทำผลิตภัณฑ์ สำหรับเข้าแข่งขันประกวดนวัตกรรมผลิตภัณฑ์อาหาร ปี 2568\n"
        "และต้องการใช้สิ่งของดังกล่าว ประมาณ มิถุนายน 2568.\n"
        "\n" # บรรทัดว่างเปล่า เพื่อสร้างย่อหน้าที่ว่างเปล่าใน Word
        "\tโดยขอแต่งตั้งกรรมการตรวจรับ คือ ผศ.ดร.ปริตา ธนสุกาญจน์ และเห็นสมควรอนุมัติผู้ขอพร้อมทั้งขอแต่งตั้งกรรมการตรวจรับ\n"
        "ไม่โครงการที่มาเพิ่มในลิ้งค์ มีสายวิทยาศาสตร์และเทคโนโลยีการอาหาร และต้องการใช้สิ่งของดังกล่าว "
        "ประมาณเดือนมิถุนายน 2568 และเบิกจ่ายเงินจากงบประมาณของประมาณการงบ 2568 กองทุนเพื่อการศึกษา "
        "แผนงานบริหารการศึกษาอุดมศึกษา งานบริหารการศึกษาเกษตรฯ หลักสูตรวิทยาศาสตร์บัณฑิต "
        "สาขาวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดเงินอุดหนุน โครงการพัฒนาการเรียนการสอนเพื่อการอบรม "
        "โครงการพัฒนาผลิตภัณฑ์เพื่อเพิ่มสายวิทยาศาสตร์และเทคโนโลยีการอาหาร หมวดค่าใช้สอยเฉพาะหมวดอื่นๆ "
        "เป็นเงิน 1,000 บาท (หนึ่งพันบาทถ้วน)\n"
    )
    text_body.insert("1.0", example_text)

    def on_create():
        title = entry_title.get().strip()
        runnum = entry_runnumber.get().strip()
        body = text_body.get("1.0", "end-1c").strip() 

        if not title or not runnum or not body:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกข้อมูลให้ครบทุกช่อง")
            return

        success = Sleeve1(data, title, runnum, body)
        if success == 1:
            messagebox.showinfo("สำเร็จ", "สร้างบันทึกข้อความเรียบร้อยแล้ว")
            dialog.destroy()
        else:
            messagebox.showerror("ไม่สำเร็จ", "สร้างบันทึกข้อความไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")

    btn_create = tk.Button(dialog, text="สร้างบันทึกข้อความ", command=on_create)
    btn_create.pack(pady=10)

def create_excel():
    """
    ฟังก์ชันสำหรับเรียกสร้างไฟล์ Excel สรุปยอด
    """
    if not data.list:
        messagebox.showwarning("ไม่มีข้อมูล", "กรุณาเพิ่มรายการพัสดุในตารางก่อนสร้างไฟล์ Excel")
        return
    
    # เรียกใช้ฟังก์ชันสร้าง Excel จาก excelsummary.py
    success = create_excel_summary(data.list)
    if success:
        messagebox.showinfo("สำเร็จ", "สร้างไฟล์ Excel สรุปยอดเรียบร้อยแล้ว")
    else:
        messagebox.showerror("ไม่สำเร็จ", "สร้างไฟล์ Excel สรุปยอดไม่สำเร็จ กรุณาลองใหม่อีกครั้ง")


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
    values = item["values"]

    if values in data.list:
        data.list.remove(values)
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

tk.Button(button_frame, text="สร้างบันทึกข้อความ", width=20, command=open_create_dialog).grid(row=1, column=0, padx=5, pady=5)
# เพิ่มปุ่มสำหรับสร้าง Excel สรุปยอด
tk.Button(button_frame, text="สร้าง Excel สรุปยอด", width=20, command=create_excel).grid(row=1, column=1, padx=5, pady=5)
tk.Button(button_frame, text="ล้างข้อมูลทั้งหมด", width=20, fg="red", command=clear_all).grid(row=1, column=2, pady=5)

# ===== ตารางแสดงรายการ =====
columns = ("ชื่อพัสดุ", "หน่วยงาน", "จำนวน", "วันที่ต้องการใช้")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)

tree.pack(fill="both", expand=True, padx=10, pady=10)

root.mainloop()
