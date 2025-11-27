# (ไฟล์: daily_timesheet.py)
# (เวอร์ชัน V5.4 - เพิ่มชื่อพนักงานใน Header ของ Popup)

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
import calendar
import hr_database
from custom_widgets import DateDropdown

class DrivingDetailsPopup(tk.Toplevel):
    """หน้าต่างย่อยสำหรับกรอกรายละเอียดเที่ยวรถใน 1 วัน"""
    def __init__(self, parent, emp_id, date_obj, current_details, on_save_callback):
        super().__init__(parent)
        self.emp_id = emp_id
        self.date_obj = date_obj
        self.on_save = on_save_callback
        
        # --- (1) ดึงชื่อพนักงานมาแสดง ---
        self.emp_info = hr_database.load_single_employee(emp_id)
        if self.emp_info:
            self.emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}"
        else:
            self.emp_name = emp_id
            
        self.title(f"📝 รายละเอียดการวิ่งรถ - {self.emp_name} ({date_obj.strftime('%d/%m/%Y')})")
        self.geometry("900x650") # ขยายสูงขึ้นนิดนึงเผื่อ Header
        
        self.details_data = current_details if current_details else []
        
        self._build_ui()
        self._refresh_table()

    def _build_ui(self):
        # --- (2) ส่วนหัวแสดงชื่อพนักงาน (เพิ่มใหม่) ---
        header_frame = ttk.Frame(self, padding=(15, 10))
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text=f"👤 พนักงาน: {self.emp_name}", font=("", 12, "bold"), foreground="#2980b9").pack(side="left")
        ttk.Label(header_frame, text=f"📅 วันที่: {self.date_obj.strftime('%d/%m/%Y')}", font=("", 11)).pack(side="right")
        
        # 1. Form Input
        input_frame = ttk.LabelFrame(self, text="เพิ่มเที่ยวรถ", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        grid_opts = {'padx': 5, 'pady': 5, 'sticky': 'w'}
        
        # -- Row 0 --
        ttk.Label(input_frame, text="ประเภทรถ:").grid(row=0, column=0, **grid_opts)
        self.cb_car = ttk.Combobox(input_frame, values=["กระบะ", "เฮี๊ยบ"], state="readonly", width=10)
        self.cb_car.grid(row=0, column=1, **grid_opts)
        self.cb_car.set("กระบะ")
        
        ttk.Label(input_frame, text="ทะเบียนรถ:").grid(row=0, column=2, **grid_opts)
        self.ent_plate = ttk.Entry(input_frame, width=15)
        self.ent_plate.grid(row=0, column=3, **grid_opts)

        # -- Row 1 --
        ttk.Label(input_frame, text="ชื่อคนขับ:").grid(row=1, column=0, **grid_opts)
        self.ent_driver = ttk.Entry(input_frame, width=20)
        self.ent_driver.grid(row=1, column=1, **grid_opts)
        
        ttk.Label(input_frame, text="วันที่ส่งงาน:").grid(row=1, column=2, **grid_opts)
        self.ent_send_date = DateDropdown(input_frame)
        self.ent_send_date.set_date(self.date_obj)
        self.ent_send_date.grid(row=1, column=3, **grid_opts)

        # -- Row 2 (เอกสาร) --
        ttk.Label(input_frame, text="เอกสารอ้างอิง:").grid(row=2, column=0, **grid_opts)
        doc_frame = ttk.Frame(input_frame)
        doc_frame.grid(row=2, column=1, columnspan=3, **grid_opts)
        
        self.cb_doc_type = ttk.Combobox(doc_frame, values=["SO", "DS", "PX"], state="readonly", width=5)
        self.cb_doc_type.pack(side="left")
        self.cb_doc_type.set("SO")
        
        self.ent_doc_id = ttk.Entry(doc_frame, width=20)
        self.ent_doc_id.pack(side="left", padx=5)
        
        # -- Row 3 (Checkbox ฟรี/พ่วง) --
        self.is_free_var = tk.BooleanVar(value=False)
        self.cb_free = ttk.Checkbutton(input_frame, text="เที่ยวพ่วง / ฟรี (ไม่คิดเงิน)", variable=self.is_free_var)
        self.cb_free.grid(row=3, column=1, columnspan=2, **grid_opts)
        
        # ปุ่มเพิ่ม (วางไว้ Row 3 ขวาสุด)
        ttk.Button(input_frame, text="➕ เพิ่มรายการ", command=self._add_item, style="Success.TButton").grid(row=3, column=3, padx=10, sticky="e")

        # -- Row 4: หมายเหตุ --
        note_text = "* หมายเหตุ: หากเป็นเที่ยวปกติที่ต้องการเบิกเงิน ห้ามติ๊กช่องนี้เด็ดขาด"
        ttk.Label(input_frame, text=note_text, font=("Segoe UI", 9), foreground="#c0392b").grid(row=4, column=1, columnspan=3, sticky='w', padx=5)

        # 2. Table List
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=10)
        
        cols = ("car", "plate", "doc_ref", "driver", "send_date", "cost", "status")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        
        self.tree.heading("car", text="ประเภท")
        self.tree.heading("plate", text="ทะเบียน")
        self.tree.heading("doc_ref", text="เอกสารอ้างอิง") 
        self.tree.heading("driver", text="คนขับ")
        self.tree.heading("send_date", text="วันที่ส่ง")
        self.tree.heading("cost", text="ค่าเที่ยว")
        self.tree.heading("status", text="สถานะ")
        
        self.tree.column("car", width=60, anchor="center")
        self.tree.column("plate", width=80)
        self.tree.column("doc_ref", width=120) 
        self.tree.column("driver", width=100)
        self.tree.column("send_date", width=80, anchor="center")
        self.tree.column("cost", width=60, anchor="e")
        self.tree.column("status", width=60, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        # 3. Footer
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="❌ ลบรายการที่เลือก", command=self._delete_item).pack(side="left")
        self.lbl_total = ttk.Label(btn_frame, text="รวมเงิน: 0.00 บาท", font=("", 12, "bold"), foreground="blue")
        self.lbl_total.pack(side="left", padx=20)
        ttk.Button(btn_frame, text="💾 บันทึกและปิด", command=self._confirm_save, style="Primary.TButton").pack(side="right")

    def _add_item(self):
        car = self.cb_car.get()
        is_free = self.is_free_var.get()
        
        if is_free:
            cost = 0.0
        else:
            cost = 50.0 if car == "กระบะ" else 100.0
        
        prefix = self.cb_doc_type.get()
        number = self.ent_doc_id.get().strip()
        combined_doc_id = f"{prefix}{number}" if number else ""

        item = {
            "car_type": car,
            "license": self.ent_plate.get().strip(),
            "driver": self.ent_driver.get().strip(),
            "send_date": self.ent_send_date.get_date(),
            "cost": cost,
            "doc_type": prefix,
            "doc_id": combined_doc_id,
            "is_free": is_free
        }
        self.details_data.append(item)
        self._refresh_table()
        
        # Reset fields
        self.ent_doc_id.delete(0, tk.END)
        self.is_free_var.set(False) 
        
    def _delete_item(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        del self.details_data[idx]
        self._refresh_table()

    def _refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        total = 0
        for item in self.details_data:
            d_str = item['send_date'].strftime('%d/%m/%Y') if item['send_date'] else "-"
            doc_show = item.get('doc_id', '')
            status_text = "ฟรี" if item.get('is_free') else "คิดเงิน"
            
            self.tree.insert("", "end", values=(
                item.get('car_type'), item.get('license'), 
                doc_show, item.get('driver'), d_str, 
                f"{item.get('cost'):.2f}",
                status_text
            ))
            total += item.get('cost', 0)
        self.lbl_total.config(text=f"รวมเงิน: {total:,.2f} บาท")

    def _confirm_save(self):
        self.on_save(self.date_obj, self.details_data)
        self.destroy()


class DailyTimesheetWindow(tk.Toplevel):
    """หน้าต่างหลัก: ปฏิทินรายวัน (แสดงยอดรวม)"""
    def __init__(self, parent, emp_id, month, year):
        super().__init__(parent)
        self.emp_id = emp_id
        self.month = month
        self.year = year
        
        self.emp_info = hr_database.load_single_employee(emp_id)
        if not self.emp_info: self.destroy(); return

        self.emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}"
        self.title(f"🚚 ตารางค่าเที่ยวรถ - {self.emp_name} ({month}/{year})")
        self.geometry("600x750")
        
        self.row_widgets = {}
        self._build_ui()
        
    def _build_ui(self):
        header = ttk.Frame(self, padding=15)
        header.pack(fill="x")
        ttk.Label(header, text=f"พนักงาน: {self.emp_name}", font=("", 12, "bold")).pack(side="left")
        
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        headers = ["วันที่", "ยอดเงิน (บาท)", "จัดการ"]
        h_frame = ttk.Frame(self.scroll_frame)
        h_frame.pack(fill="x", pady=5)
        widths = [15, 20, 15]
        for i, h in enumerate(headers):
            ttk.Label(h_frame, text=h, width=widths[i], anchor="center", font=("", 10, "bold")).pack(side="left", padx=10)
            
        num_days = calendar.monthrange(self.year, self.month)[1]
        for day in range(1, num_days + 1):
            current_date = date(self.year, self.month, day)
            self._create_row(day, current_date)
            
    def _create_row(self, day, current_date):
        # ดึงข้อมูลจาก DB
        details = hr_database.get_driving_details(self.emp_id, current_date)
        
        # คำนวณยอดรวม (trip_cost)
        total_amt = sum(d.get('trip_cost', 0) or 0 for d in details)
        
        row = ttk.Frame(self.scroll_frame)
        row.pack(fill="x", pady=2)
        
        date_str = f"{day}/{self.month}/{self.year+543}"
        ttk.Label(row, text=date_str, width=15, anchor="center").pack(side="left", padx=10)
        
        lbl_amt = ttk.Label(row, text=f"{total_amt:,.2f}", width=20, anchor="e", font=("", 10, "bold"), foreground="blue")
        lbl_amt.pack(side="left", padx=10)
        
        btn = ttk.Button(row, text="📝 รายละเอียด", command=lambda d=current_date: self._open_details(d))
        btn.pack(side="left", padx=10)
        
        # เก็บข้อมูลดิบจาก DB ไว้ใน cache
        self.row_widgets[current_date] = {
            "label": lbl_amt,
            "data": details 
        }

    def _open_details(self, date_obj):
        # 1. ดึงข้อมูลดิบจาก Cache (ซึ่งมาจาก DB)
        raw_data = self.row_widgets[date_obj]["data"]
        
        # 2. แปลงข้อมูลให้เป็น Format ที่หน้าจอ Popup เข้าใจ (Mapping)
        formatted_data = []
        for item in raw_data:
            formatted_data.append({
                'car_type': item.get('car_type'),
                'license': item.get('license_plate', ''),
                'driver': item.get('driver_name', ''),
                'send_date': item.get('delivery_date'), # แปลง delivery_date -> send_date
                'cost': float(item.get('trip_cost', 0) or 0),
                'doc_type': item.get('ref_doc_type', ''),
                'doc_id': item.get('ref_doc_id', ''),
                'is_free': bool(item.get('is_free', False))
            })
        
        # 3. ฟังก์ชัน Callback เมื่อบันทึกเสร็จ
        def on_popup_save(target_date, new_details):
            # บันทึกลง DB
            hr_database.save_driving_details_list(self.emp_id, target_date, new_details)
            
            # คำนวณยอดรวมใหม่เพื่อโชว์
            new_total = sum(x['cost'] for x in new_details)
            self.row_widgets[target_date]["label"].config(text=f"{new_total:,.2f}")
            
            # โหลดข้อมูลจริงจาก DB มาเก็บใส่ Cache อีกรอบ (เพื่อให้ได้ Key ที่ถูกต้องเสมอ)
            updated_db_data = hr_database.get_driving_details(self.emp_id, target_date)
            self.row_widgets[target_date]["data"] = updated_db_data
            
        # 4. เปิดหน้าต่าง
        DrivingDetailsPopup(self, self.emp_id, date_obj, formatted_data, on_popup_save)