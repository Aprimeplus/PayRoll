# (ไฟล์: daily_timesheet.py)
# (เวอร์ชัน V7.2 - เพิ่มยอดรวมที่ Footer ตามคำขอ)

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
import calendar
import hr_database
from custom_widgets import DateDropdown

# ==========================================
#  ส่วนที่ 1: Popup บันทึกเที่ยวรถ (Driving) (เหมือนเดิม)
# ==========================================
class DrivingDetailsPopup(tk.Toplevel):
    """หน้าต่างย่อยสำหรับกรอกรายละเอียดเที่ยวรถ"""
    def __init__(self, parent, emp_id, date_obj, current_details, on_save_callback):
        super().__init__(parent)
        self.emp_info = hr_database.load_single_employee(emp_id)
        self.emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}" if self.emp_info else emp_id 
        self.title(f"🚚 รายละเอียดการวิ่งรถ - {self.emp_name} ({date_obj.strftime('%d/%m/%Y')})")
        self.geometry("950x650")
        self.emp_id = emp_id
        self.date_obj = date_obj
        self.on_save = on_save_callback
        self.details_data = current_details if current_details else []
        self._build_ui()
        self._refresh_table()

    def _build_ui(self):
        header_frame = ttk.Frame(self, padding=(15, 10))
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text=f"วันที่: {self.date_obj.strftime('%d/%m/%Y')}", font=("", 12, "bold"), foreground="#2980b9").pack(side="left")
        
        # 1. Form Input
        input_frame = ttk.LabelFrame(self, text="เพิ่มเที่ยวรถ", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        grid_opts = {'padx': 5, 'pady': 5, 'sticky': 'w'}
        
        ttk.Label(input_frame, text="ประเภทรถ:").grid(row=0, column=0, **grid_opts)
        self.cb_car = ttk.Combobox(input_frame, values=["กระบะ", "เฮี๊ยบ"], state="readonly", width=10)
        self.cb_car.grid(row=0, column=1, **grid_opts)
        self.cb_car.set("กระบะ")
        
        ttk.Label(input_frame, text="ทะเบียนรถ:").grid(row=0, column=2, **grid_opts)
        self.ent_plate = ttk.Entry(input_frame, width=15)
        self.ent_plate.grid(row=0, column=3, **grid_opts)

        ttk.Label(input_frame, text="ชื่อคนขับ:").grid(row=1, column=0, **grid_opts)
        self.ent_driver = ttk.Entry(input_frame, width=20)
        self.ent_driver.grid(row=1, column=1, **grid_opts)
        
        ttk.Label(input_frame, text="วันที่ส่งงาน:").grid(row=1, column=2, **grid_opts)
        self.ent_send_date = DateDropdown(input_frame)
        self.ent_send_date.set_date(self.date_obj)
        self.ent_send_date.grid(row=1, column=3, **grid_opts)

        ttk.Label(input_frame, text="เอกสารอ้างอิง:").grid(row=2, column=0, **grid_opts)
        doc_frame = ttk.Frame(input_frame)
        doc_frame.grid(row=2, column=1, columnspan=3, **grid_opts)
        
        self.cb_doc_type = ttk.Combobox(doc_frame, values=["SO", "DS", "PX"], state="readonly", width=5)
        self.cb_doc_type.pack(side="left")
        self.cb_doc_type.set("SO")
        
        self.ent_doc_id = ttk.Entry(doc_frame, width=20)
        self.ent_doc_id.pack(side="left", padx=5)

        opt_frame = ttk.Frame(input_frame)
        opt_frame.grid(row=3, column=0, columnspan=5, sticky='w', pady=5)
        
        self.is_free_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="เที่ยวพ่วง / ฟรี (0 บาท)", variable=self.is_free_var).pack(side="left", padx=(5, 20))
        
        self.is_service_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="บริการรวมลง", variable=self.is_service_var, command=self._toggle_service_fee).pack(side="left", padx=5)
        
        ttk.Label(opt_frame, text="จำนวนเงิน:").pack(side="left", padx=5)
        self.ent_service_fee = ttk.Entry(opt_frame, width=10, state="disabled")
        self.ent_service_fee.pack(side="left")
        ttk.Label(opt_frame, text="บาท").pack(side="left")
        
        ttk.Button(input_frame, text="➕ เพิ่มรายการ", command=self._add_item, style="Success.TButton").grid(row=3, column=4, padx=10, sticky="e")

        # 2. Table
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=10)
        
        cols = ("car", "plate", "doc_ref", "driver", "send_date", "cost", "service", "total")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        
        self.tree.heading("car", text="ประเภท")
        self.tree.heading("plate", text="ทะเบียน")
        self.tree.heading("doc_ref", text="เอกสาร") 
        self.tree.heading("driver", text="คนขับ")
        self.tree.heading("send_date", text="วันที่ส่ง")
        self.tree.heading("cost", text="ค่าเที่ยว")
        self.tree.heading("service", text="ค่าบริการ")
        self.tree.heading("total", text="รวม")
        
        self.tree.column("car", width=60, anchor="center")
        self.tree.column("plate", width=80)
        self.tree.column("doc_ref", width=110) 
        self.tree.column("driver", width=100)
        self.tree.column("send_date", width=80, anchor="center")
        self.tree.column("cost", width=60, anchor="e")
        self.tree.column("service", width=60, anchor="e")
        self.tree.column("total", width=70, anchor="e")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        # 3. Footer
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="❌ ลบรายการ", command=self._delete_item).pack(side="left")
        self.lbl_total = ttk.Label(btn_frame, text="รวมเงินสุทธิ: 0.00 บาท", font=("", 12, "bold"), foreground="blue")
        self.lbl_total.pack(side="left", padx=20)
        ttk.Button(btn_frame, text="💾 บันทึกและปิด", command=self._confirm_save, style="Primary.TButton").pack(side="right")

    def _toggle_service_fee(self):
        if self.is_service_var.get():
            self.ent_service_fee.config(state="normal")
            self.ent_service_fee.focus()
        else:
            self.ent_service_fee.delete(0, tk.END)
            self.ent_service_fee.config(state="disabled")

    def _add_item(self):
        car = self.cb_car.get()
        is_free = self.is_free_var.get()
        is_service = self.is_service_var.get()
        
        cost = 0.0
        if not is_free:
            cost = 50.0 if car == "กระบะ" else 100.0
            
        service_fee = 0.0
        if is_service:
            try: service_fee = float(self.ent_service_fee.get())
            except: return

        prefix = self.cb_doc_type.get()
        number = self.ent_doc_id.get().strip()
        combined_doc_id = f"{prefix}{number}" if number else ""

        item = {
            "car_type": car, "license": self.ent_plate.get().strip(),
            "driver": self.ent_driver.get().strip(), "send_date": self.ent_send_date.get_date(),
            "cost": cost, "doc_type": prefix, "doc_id": combined_doc_id,
            "is_free": is_free, "is_service": is_service, "service_fee": service_fee
        }
        self.details_data.append(item)
        self._refresh_table()
        self.ent_doc_id.delete(0, tk.END)
        
    def _delete_item(self):
        sel = self.tree.selection()
        if not sel: return
        del self.details_data[self.tree.index(sel[0])]
        self._refresh_table()

    def _refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        grand_total = 0
        for item in self.details_data:
            d_str = item['send_date'].strftime('%d/%m/%Y') if item['send_date'] else "-"
            doc_show = item.get('doc_id', '')
            cost = item.get('cost', 0)
            svc = item.get('service_fee', 0)
            total = cost + svc
            grand_total += total
            self.tree.insert("", "end", values=(
                item.get('car_type'), item.get('license'), doc_show, item.get('driver'), 
                d_str, f"{cost:.2f}", f"{svc:.2f}", f"{total:.2f}"
            ))
        self.lbl_total.config(text=f"รวมเงินสุทธิ: {grand_total:,.2f} บาท")

    def _confirm_save(self):
        self.on_save(self.date_obj, self.details_data)
        self.destroy()


# ==========================================
#  ส่วนที่ 2: หน้าจอหลัก (All-in-One)
# ==========================================
class DailyTimesheetWindow(tk.Toplevel):
    """หน้าต่างหลัก: รวมการจัดการเที่ยวรถและ OT ในหน้าเดียว"""
    def __init__(self, parent, emp_id, month, year):
        super().__init__(parent)
        self.emp_id = emp_id
        self.month = month
        self.year = year
        
        self.emp_info = hr_database.load_single_employee(emp_id)
        if not self.emp_info: self.destroy(); return

        self.emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}"
        self.title(f"📅 บันทึกงานรายวัน - {self.emp_name} ({month}/{year})")
        self.geometry("800x750")
        
        self.row_widgets = {}
        
        # ตัวแปรเก็บยอดรวม (เพื่ออัปเดตง่ายๆ)
        self.total_drive_var = tk.DoubleVar(value=0.0)
        self.total_ot_var = tk.DoubleVar(value=0.0)
        
        self._build_ui()
        self._calculate_initial_totals() # คำนวณยอดเริ่มต้นทันทีที่เปิด
        
    def _build_ui(self):
        header = ttk.Frame(self, padding=15)
        header.pack(fill="x")
        
        ttk.Label(header, text=f"พนักงาน: {self.emp_name}", font=("", 14, "bold"), foreground="#2c3e50").pack(side="left")
        ttk.Label(header, text=f"ประจำเดือน: {self.month}/{self.year}", font=("", 12)).pack(side="right")
        
        legend_frame = ttk.Frame(self)
        legend_frame.pack(fill="x", padx=20, pady=(0, 5))
        tk.Label(legend_frame, text="⬜ วันทำงาน", bg="white", relief="solid", borderwidth=1, padx=5).pack(side="left", padx=5)
        tk.Label(legend_frame, text="⬛ วันอาทิตย์ (หยุด)", bg="#f0f0f0", relief="solid", borderwidth=1, padx=5).pack(side="left", padx=5)

        h_frame = ttk.Frame(self)
        h_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        lbl_style = {"font": ("", 10, "bold"), "anchor": "center"}
        ttk.Label(h_frame, text="วันที่", width=15, **lbl_style).pack(side="left")
        ttk.Label(h_frame, text="💰 ค่าเที่ยว (บาท)", width=20, **lbl_style).pack(side="left", padx=10)
        ttk.Label(h_frame, text="", width=15).pack(side="left") 
        ttk.Label(h_frame, text="⏱️ OT (ชม.)", width=15, **lbl_style).pack(side="left", padx=10)

        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)
        
        num_days = calendar.monthrange(self.year, self.month)[1]
        for day in range(1, num_days + 1):
            current_date = date(self.year, self.month, day)
            self._create_row(day, current_date)
            
        # --- [NEW] Footer แสดงยอดรวม ---
        footer_frame = ttk.Frame(self, padding=10)
        footer_frame.pack(fill="x", side="bottom")
        
        # Label ยอดรวมค่าเที่ยว
        self.lbl_total_drive = ttk.Label(footer_frame, text="รวมค่าเที่ยว: 0.00 บาท", font=("", 11, "bold"), foreground="blue")
        self.lbl_total_drive.pack(side="left", padx=20)
        
        # Label ยอดรวม OT
        self.lbl_total_ot = ttk.Label(footer_frame, text="รวม OT: 0.00 ชม.", font=("", 11, "bold"), foreground="green")
        self.lbl_total_ot.pack(side="left", padx=20)
        
        # ปุ่มปิด (หรือปุ่มอื่นๆ ถ้ามี)
        ttk.Button(footer_frame, text="ปิด", command=self.destroy).pack(side="right", padx=10)

    def _create_row(self, day, current_date):
        row_frame = ttk.Frame(self.scroll_frame)
        row_frame.pack(fill="x", pady=2)
        
        date_str = f"{day:02d}/{self.month:02d}"
        bg_color = "#f0f0f0" if current_date.weekday() == 6 else "white"
        
        lbl_date = tk.Label(row_frame, text=date_str, width=15, anchor="center", bg=bg_color, relief="flat")
        lbl_date.pack(side="left")

        driving_details = hr_database.get_driving_details(self.emp_id, current_date)
        drive_amt = sum(d.get('trip_cost', 0) + d.get('service_fee', 0) for d in driving_details)
        
        lbl_drive = ttk.Label(row_frame, text=f"{drive_amt:,.2f}", width=20, anchor="e", foreground="blue" if drive_amt > 0 else "black")
        lbl_drive.pack(side="left", padx=10)
        
        btn_drive = ttk.Button(row_frame, text="🚚 เที่ยวรถ", width=12,
                               command=lambda d=current_date: self._open_driving_popup(d))
        btn_drive.pack(side="left", padx=2)

        ot_details = hr_database.get_ot_details_list(self.emp_id, current_date)
        ot_hrs = sum(d['period_hours'] for d in ot_details)
        
        lbl_ot = ttk.Label(row_frame, text=f"{ot_hrs:.2f}", width=15, anchor="center", foreground="red" if ot_hrs > 0 else "black")
        lbl_ot.pack(side="left", padx=10)
        
        self.row_widgets[current_date] = {
            "lbl_drive": lbl_drive,
            "lbl_ot": lbl_ot,
            "drive_data": driving_details,
            "ot_data": ot_details,
            "drive_val": drive_amt, # เก็บค่าตัวเลขไว้คำนวณรวม
            "ot_val": ot_hrs        # เก็บค่าตัวเลขไว้คำนวณรวม
        }

    def _calculate_initial_totals(self):
        """คำนวณยอดรวมทั้งหมดตอนเปิดหน้าต่าง"""
        total_drive = sum(w["drive_val"] for w in self.row_widgets.values())
        total_ot = sum(w["ot_val"] for w in self.row_widgets.values())
        self._update_footer_labels(total_drive, total_ot)

    def _update_footer_labels(self, drive_sum, ot_sum):
        """อัปเดตข้อความใน Footer"""
        self.lbl_total_drive.config(text=f"รวมค่าเที่ยว: {drive_sum:,.2f} บาท")
        self.lbl_total_ot.config(text=f"รวม OT: {ot_sum:.2f} ชม.")

    def _open_driving_popup(self, date_obj):
        raw_data = self.row_widgets[date_obj]["drive_data"]
        formatted = []
        for item in raw_data:
            formatted.append({
                'car_type': item.get('car_type'), 'license': item.get('license_plate', ''),
                'driver': item.get('driver_name', ''), 'send_date': item.get('delivery_date'),
                'cost': float(item.get('trip_cost', 0) or 0),
                'doc_type': item.get('ref_doc_type', ''), 'doc_id': item.get('ref_doc_id', ''),
                'is_free': bool(item.get('is_free', False)),
                'is_service': bool(item.get('is_service', False)),
                'service_fee': float(item.get('service_fee', 0) or 0)
            })
            
        def on_save(target_date, new_details):
            hr_database.save_driving_details_list(self.emp_id, target_date, new_details)
            new_total = sum(x['cost'] + x.get('service_fee', 0) for x in new_details)
            
            # อัปเดต Label แถวนั้น
            self.row_widgets[target_date]["lbl_drive"].config(text=f"{new_total:,.2f}", foreground="blue" if new_total > 0 else "black")
            
            # อัปเดตข้อมูลใน Memory
            self.row_widgets[target_date]["drive_data"] = hr_database.get_driving_details(self.emp_id, target_date)
            self.row_widgets[target_date]["drive_val"] = new_total # อัปเดตค่าตัวเลข
            
            # คำนวณยอดรวมใหม่ทั้งตาราง
            self._calculate_initial_totals()

        DrivingDetailsPopup(self, self.emp_id, date_obj, formatted, on_save)