import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date, timedelta
import calendar
import hr_database
from custom_widgets import DateDropdown

class OTDetailsPopup(tk.Toplevel):
    """หน้าต่างย่อยสำหรับเพิ่มช่วงเวลา OT"""
    def __init__(self, parent, emp_id, date_obj, current_details, on_save_callback):
        super().__init__(parent)
        self.emp_id = emp_id
        self.date_obj = date_obj
        self.on_save = on_save_callback
        
        # ดึงชื่อพนักงาน
        self.emp_info = hr_database.load_single_employee(emp_id)
        emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}" if self.emp_info else emp_id
        
        self.title(f"⏱️ บันทึก OT (รายช่วง) - {emp_name} ({date_obj.strftime('%d/%m/%Y')})")
        self.geometry("800x500")
        
        self.details_data = current_details if current_details else []
        
        self.time_options = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]
        
        self._build_ui()
        self._refresh_table()

    def _build_ui(self):
        # --- Input Frame ---
        input_frame = ttk.LabelFrame(self, text="เพิ่มช่วงเวลาทำงานล่วงเวลา", padding=10)
        input_frame.pack(fill="x", padx=10, pady=10)
        
        grid_opts = {'padx': 5, 'pady': 5, 'sticky': 'w'}
        
        # Start Time
        ttk.Label(input_frame, text="เริ่มเวลา:").grid(row=0, column=0, **grid_opts)
        self.cb_start = ttk.Combobox(input_frame, values=self.time_options, width=8)
        self.cb_start.grid(row=0, column=1, **grid_opts)
        
        # End Time
        ttk.Label(input_frame, text="ถึงเวลา:").grid(row=0, column=2, **grid_opts)
        self.cb_end = ttk.Combobox(input_frame, values=self.time_options, width=8)
        self.cb_end.grid(row=0, column=3, **grid_opts)
        
        # Auto Calc Button
        ttk.Button(input_frame, text="➡️ คำนวณ", command=self._calc_hours, width=8).grid(row=0, column=4, padx=5)

        # Hours (Manual Edit Allowed)
        ttk.Label(input_frame, text="จำนวน ชม.:").grid(row=0, column=5, **grid_opts)
        self.ent_hours = ttk.Entry(input_frame, width=8)
        self.ent_hours.grid(row=0, column=6, **grid_opts)
        
        # Description
        ttk.Label(input_frame, text="รายละเอียดงาน:").grid(row=1, column=0, **grid_opts)
        self.ent_desc = ttk.Entry(input_frame, width=40)
        self.ent_desc.grid(row=1, column=1, columnspan=4, **grid_opts)
        
        # Add Button
        ttk.Button(input_frame, text="➕ เพิ่มรายการ", command=self._add_item, style="Success.TButton").grid(row=1, column=6, padx=5)
        
        # --- Table List ---
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=10)
        
        cols = ("start", "end", "hours", "desc")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings")
        self.tree.heading("start", text="เริ่ม")
        self.tree.heading("end", text="ถึง")
        self.tree.heading("hours", text="จำนวน (ชม.)")
        self.tree.heading("desc", text="รายละเอียด")
        
        self.tree.column("start", width=80, anchor="center")
        self.tree.column("end", width=80, anchor="center")
        self.tree.column("hours", width=100, anchor="center")
        self.tree.column("desc", width=300, anchor="w")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        # --- Footer ---
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="❌ ลบรายการ", command=self._delete_item).pack(side="left")
        self.lbl_total = ttk.Label(btn_frame, text="รวม OT: 0.00 ชม.", font=("", 12, "bold"), foreground="blue")
        self.lbl_total.pack(side="left", padx=20)
        
        ttk.Button(btn_frame, text="💾 บันทึกและปิด", command=self._confirm_save, style="Primary.TButton").pack(side="right")

    def _calc_hours(self):
        """คำนวณเวลา (รองรับข้ามวัน)"""
        s_str = self.cb_start.get()
        e_str = self.cb_end.get()
        if not s_str or not e_str: return
        
        try:
            fmt = "%H:%M"
            t_start = datetime.strptime(s_str, fmt)
            t_end = datetime.strptime(e_str, fmt)
            
            # ถ้าจบเวลาข้ามวัน (เช่น 23:00 -> 02:00) เวลาจบจะน้อยกว่าเวลาเริ่ม
            if t_end < t_start:
                t_end += timedelta(days=1) # บวกเพิ่ม 1 วัน
                
            diff = (t_end - t_start).total_seconds() / 3600 # แปลงเป็นชั่วโมง
            self.ent_hours.delete(0, tk.END)
            self.ent_hours.insert(0, f"{diff:.2f}")
            
        except ValueError:
            pass

    def _add_item(self):
        h_str = self.ent_hours.get()
        try:
            hours = float(h_str)
            if hours <= 0: raise ValueError
        except:
            messagebox.showwarning("แจ้งเตือน", "กรุณาระบุจำนวนชั่วโมงให้ถูกต้อง")
            return

        item = {
            "start": self.cb_start.get(),
            "end": self.cb_end.get(),
            "hours": hours,
            "desc": self.ent_desc.get()
        }
        self.details_data.append(item)
        self._refresh_table()
        
        # Clear Inputs
        self.ent_hours.delete(0, tk.END)
        self.ent_desc.delete(0, tk.END)
        self.cb_start.set("")
        self.cb_end.set("")

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
            self.tree.insert("", "end", values=(
                item['start'], item['end'], f"{item['hours']:.2f}", item['desc']
            ))
            total += item['hours']
        self.lbl_total.config(text=f"รวม OT: {total:.2f} ชม.")

    def _confirm_save(self):
        self.on_save(self.date_obj, self.details_data)
        self.destroy()

class OTManagerWindow(tk.Toplevel):
    """หน้าต่างหลัก: ปฏิทินจัดการ OT รายวัน"""
    def __init__(self, parent, emp_id, month, year):
        super().__init__(parent)
        self.emp_id = emp_id
        self.month = month
        self.year = year
        
        self.emp_info = hr_database.load_single_employee(emp_id)
        if not self.emp_info: self.destroy(); return

        emp_name = f"{self.emp_info['fname']} {self.emp_info['lname']}"
        self.title(f"⏱️ จัดการ OT - {emp_name} ({month}/{year})")
        self.geometry("600x700")
        
        self.row_widgets = {}
        self._build_ui()
        
    def _build_ui(self):
        header = ttk.Frame(self, padding=15)
        header.pack(fill="x")
        ttk.Label(header, text=f"พนักงาน: {self.emp_info['fname']} {self.emp_info['lname']}", font=("", 12, "bold")).pack(side="left")
        
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Headers
        h_frame = ttk.Frame(self.scroll_frame)
        h_frame.pack(fill="x", pady=5)
        ttk.Label(h_frame, text="วันที่", width=15, anchor="center", font=("", 10, "bold")).pack(side="left", padx=5)
        ttk.Label(h_frame, text="OT (ชม.)", width=15, anchor="center", font=("", 10, "bold")).pack(side="left", padx=5)
        
        # Rows
        num_days = calendar.monthrange(self.year, self.month)[1]
        for day in range(1, num_days + 1):
            current_date = date(self.year, self.month, day)
            self._create_row(day, current_date)
            
    def _create_row(self, day, current_date):
        # ดึงข้อมูล OT รวม (จาก Daily Records เดิม หรือคำนวณใหม่ก็ได้)
        # เพื่อความแม่นยำ เราดึงรายละเอียดแล้วรวมใหม่ดีกว่า
        details = hr_database.get_ot_details_list(self.emp_id, current_date)
        total_hrs = sum(d['period_hours'] for d in details)
        
        row = ttk.Frame(self.scroll_frame)
        row.pack(fill="x", pady=2)
        
        date_str = f"{day}/{self.month}/{self.year+543}"
        ttk.Label(row, text=date_str, width=15, anchor="center").pack(side="left", padx=5)
        
        lbl_hrs = ttk.Label(row, text=f"{total_hrs:.2f}", width=15, anchor="center", font=("", 10, "bold"), foreground="red" if total_hrs > 0 else "black")
        lbl_hrs.pack(side="left", padx=5)
        
        ttk.Button(row, text="📝 จัดการเวลา", command=lambda d=current_date: self._open_details(d)).pack(side="left", padx=5)
        
        self.row_widgets[current_date] = {
            "label": lbl_hrs,
            "data": details
        }

    def _open_details(self, date_obj):
        current_data = self.row_widgets[date_obj]["data"]
        
        def on_save(target_date, new_details):
            hr_database.save_ot_details_list(self.emp_id, target_date, new_details)
            new_total = sum(x['hours'] for x in new_details)
            
            self.row_widgets[target_date]["label"].config(text=f"{new_total:.2f}", foreground="red" if new_total > 0 else "black")
            
            # Reload to update cache
            updated_data = hr_database.get_ot_details_list(self.emp_id, target_date)
            # (Mapping keys ให้ตรงกับ Popup)
            mapped_data = []
            for d in updated_data:
                mapped_data.append({
                    'start': d['start_time'], 'end': d['end_time'], 
                    'hours': d['period_hours'], 'desc': d['description']
                })
            self.row_widgets[target_date]["data"] = mapped_data

        # แปลง data ก่อนส่งเข้า Popup
        popup_data = []
        for d in current_data:
             # ถ้า data มาจาก DB keys จะเป็น start_time, period_hours
             # ถ้า data มาจาก Memory keys จะเป็น start, hours
             # ต้อง Normalize
             start = d.get('start_time') or d.get('start')
             end = d.get('end_time') or d.get('end')
             hours = d.get('period_hours') or d.get('hours')
             desc = d.get('description') or d.get('desc')
             popup_data.append({'start': start, 'end': end, 'hours': hours, 'desc': desc})
             
        OTDetailsPopup(self, self.emp_id, date_obj, popup_data, on_save)