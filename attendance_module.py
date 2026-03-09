# (ไฟล์: attendance_module.py)
# (เวอร์ชันอัปเกรด - รองรับ "ลารายชั่วโมง" และ "ลาไม่รับค่าจ้าง")

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog 
from custom_widgets import DateDropdown 
import hr_database
from datetime import datetime, time, timedelta
import os
import shutil

NETWORK_UPLOAD_PATH = r"\\192.168.1.51\HR_System_Documents"

class AttendanceModule(ttk.Frame):
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user

        self.THAI_MONTHS = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
            5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
            9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }
        self.MONTH_TO_INT = {v: k for k, v in self.THAI_MONTHS.items()}

        self.selected_emp_id = tk.StringVar()
        self.selected_emp_name = tk.StringVar()

        self.current_att_warning_doc_id = None
        self.current_att_warning_file_path = None
        
        # (3. สร้าง List เวลา 00:00 - 23:45)
        self.time_options = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]
        
        self._create_main_layout()
        self._build_list_panel()
        self._build_form_panel()
        self._load_employee_list()

    def _open_history_window(self):
        """เปิดหน้าต่างจัดการประวัติการลา พร้อมตัวกรอง เดือน/ปี และฟังก์ชันแก้ไข"""
        win = tk.Toplevel(self)
        win.title("📜 จัดการประวัติการลา")
        win.geometry("1100x700") # ขยายขนาดนิดหน่อย
        win.transient(self)
        win.grab_set()

        # --- ส่วนตัวกรอง (Filter) ---
        filter_frame = ttk.LabelFrame(win, text=" ตัวกรองข้อมูล ", padding=10)
        filter_frame.pack(fill="x", padx=10, pady=5)

        # 1. เลือกพนักงาน
        ttk.Label(filter_frame, text="พนักงาน:").pack(side="left", padx=5)
        emps = hr_database.load_all_employees()
        emp_list = ["--- แสดงทั้งหมด ---"] + [f"{e['id']} - {e['fname']} {e['lname']}" for e in emps]
        self.history_emp_filter = ttk.Combobox(filter_frame, width=30, state="readonly", values=emp_list)
        self.history_emp_filter.pack(side="left", padx=5)
        self.history_emp_filter.set("--- แสดงทั้งหมด ---")

        # 2. เลือกเดือน (เพิ่มใหม่)
        ttk.Label(filter_frame, text="เดือน:").pack(side="left", padx=(15, 5))
        months = list(self.THAI_MONTHS.values())
        self.filter_month = ttk.Combobox(filter_frame, values=["ทุกเดือน"] + months, width=12, state="readonly")
        self.filter_month.set(self.THAI_MONTHS[datetime.now().month]) # Default เดือนปัจจุบัน
        self.filter_month.pack(side="left", padx=5)

        # 3. เลือกปี (เพิ่มใหม่)
        ttk.Label(filter_frame, text="ปี:").pack(side="left", padx=(15, 5))
        current_year = datetime.now().year + 543
        years = [str(y) for y in range(current_year, current_year - 10, -1)]
        self.filter_year = ttk.Combobox(filter_frame, values=years, width=8, state="readonly")
        self.filter_year.set(str(current_year))
        self.filter_year.pack(side="left", padx=5)

        # ปุ่มค้นหา
        ttk.Button(filter_frame, text="🔍 ค้นหา", command=lambda: load_data()).pack(side="left", padx=15)

        # --- ตารางข้อมูล ---
        tree_frame = ttk.Frame(win, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        cols = ("id", "emp_id", "name", "date", "type", "days", "reason")
        self.history_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        
        headers = ["ID", "รหัส", "ชื่อ-นามสกุล", "วันที่ลา", "ประเภท", "จำนวนวัน", "หมายเหตุ"]
        widths = [50, 80, 200, 100, 100, 80, 250]
        
        for col, head, w in zip(cols, headers, widths):
            self.history_tree.heading(col, text=head)
            self.history_tree.column(col, width=w)
        
        # Scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=vsb.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def load_data(event=None):
            for i in self.history_tree.get_children(): self.history_tree.delete(i)
            
            # รับค่าจาก Filter
            selected_emp = self.history_emp_filter.get()
            target_id = selected_emp.split(" - ")[0] if selected_emp != "--- แสดงทั้งหมด ---" else None
            
            # แปลงเดือนไทยเป็นเลข (ถ้าเลือก "ทุกเดือน" ให้เป็น None)
            selected_month_str = self.filter_month.get()
            target_month = self.MONTH_TO_INT.get(selected_month_str) if selected_month_str != "ทุกเดือน" else None
            
            # แปลงปี พ.ศ. เป็น ค.ศ.
            target_year = int(self.filter_year.get()) - 543
            
            # เรียกฟังก์ชันใหม่ใน Database (ต้องไปเพิ่มฟังก์ชันนี้ใน hr_database.py ด้วยนะครับ)
            # ถ้ายังไม่ได้ทำ ให้ใช้ของเดิมไปก่อน แล้วค่อยไปแก้ทีหลัง
            # history = hr_database.get_leave_history_filtered(target_id, target_month, target_year) 
            
            # *ชั่วคราว: ใช้ฟังก์ชันเดิมแล้วกรองด้วย Python (ถ้าข้อมูลไม่เยอะมาก)*
            all_history = hr_database.get_employee_leave_history(target_id) # ฟังก์ชันเดิมดึงมาหมด
            
            for item in all_history:
                # กรองเดือน/ปี
                l_date = item['leave_date'] # เป็น datetime object
                if target_year and l_date.year != target_year: continue
                if target_month and l_date.month != target_month: continue
                
                self.history_tree.insert("", "end", values=(
                    item['leave_id'], item['emp_id'], f"{item['fname']} {item['lname']}",
                    hr_database.date_to_thai_str(item['leave_date']),
                    item['leave_type'], item['num_days'], item['reason']
                ))

        # Bind Events (กดเลือกแล้วโหลดเลย)
        self.history_emp_filter.bind("<<ComboboxSelected>>", load_data)
        self.filter_month.bind("<<ComboboxSelected>>", load_data)
        self.filter_year.bind("<<ComboboxSelected>>", load_data)
        
        load_data() # โหลดครั้งแรก

        # --- ปุ่มควบคุมด้านล่าง ---
        btn_frame = ttk.Frame(win, padding=10)
        btn_frame.pack(fill="x")
        
        def delete_entry():
            selected = self.history_tree.selection()
            if not selected:
                messagebox.showwarning("เตือน", "กรุณาเลือกรายการที่ต้องการลบ")
                return
            
            vals = self.history_tree.item(selected[0])['values']
            if messagebox.askyesno("ยืนยัน", f"ต้องการลบประวัติของ {vals[2]} วันที่ {vals[3]} ใช่หรือไม่?"):
                if hr_database.delete_leave_record_by_id(vals[0]):
                    messagebox.showinfo("สำเร็จ", "ลบข้อมูลเรียบร้อยแล้ว")
                    load_data()

        def edit_entry():
            """ฟังก์ชันแก้ไขรายการที่เลือก"""
            selected = self.history_tree.selection()
            if not selected:
                messagebox.showwarning("เตือน", "กรุณาเลือกรายการที่ต้องการแก้ไข")
                return
            
            # ดึงข้อมูลจาก Treeview
            vals = self.history_tree.item(selected[0])['values']
            item_id = vals[0] # คอลัมน์แรกคือ ID (ตรวจสอบว่าใช่ ID ของตาราง leave_records จริงไหม)

            print(f"DEBUG: เลือกรายการ ID จาก Treeview = {item_id}") # <--- เช็คว่าได้ ID ถูกไหม
            
            self._open_edit_leave_popup(item_id, win, load_data)
        ttk.Button(btn_frame, text="❌ ลบรายการ", command=delete_entry).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="✏️ แก้ไข", command=edit_entry).pack(side="right", padx=5) # ปุ่มใหม่
        ttk.Button(btn_frame, text="🔄 รีเฟรช", command=load_data).pack(side="right", padx=5)

    def _open_edit_leave_popup(self, leave_id, parent_win, refresh_callback):
        """(ฉบับแก้ไข) หน้าต่าง Popup สำหรับแก้ไขข้อมูลการลา"""
        
        # 1. ดึงข้อมูลเก่าจาก DB
        print(f"DEBUG: กำลังดึงข้อมูล leave_id = {leave_id}") # Debug
        record = hr_database.get_leave_record_by_id(leave_id)
        
        # ถ้าไม่เจอข้อมูล (record เป็น None) ให้แจ้งเตือนและจบการทำงาน
        if not record:
            messagebox.showerror("Error", f"ไม่พบข้อมูลรหัส: {leave_id}\n(อาจถูกลบไปแล้ว หรือชื่อตาราง DB ไม่ถูกต้อง)")
            return

        # 2. สร้างหน้าต่าง Popup
        edit_win = tk.Toplevel(parent_win)
        edit_win.title(f"✏️ แก้ไขรายการ ID: {leave_id}")
        edit_win.geometry("450x400")
        edit_win.transient(parent_win)
        edit_win.grab_set() # ล็อคหน้าต่างแม่
        
        # จัดกึ่งกลางจอ
        try:
            edit_win.update_idletasks()
            x = parent_win.winfo_rootx() + (parent_win.winfo_width() // 2) - (edit_win.winfo_width() // 2)
            y = parent_win.winfo_rooty() + (parent_win.winfo_height() // 2) - (edit_win.winfo_height() // 2)
            edit_win.geometry(f"+{x}+{y}")
        except: pass
        
        form_frame = ttk.Frame(edit_win, padding=20)
        form_frame.pack(fill="both", expand=True)
        
        # --- สร้าง Form ---
        
        # 3. ประเภทการลา
        ttk.Label(form_frame, text="ประเภทการลา:", font=("", 10, "bold")).pack(anchor="w", pady=(0, 5))
        cb_type = ttk.Combobox(form_frame, values=["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาไม่รับค่าจ้าง", "อื่นๆ"], state="readonly", font=("", 10))
        cb_type.set(record.get('leave_type', ''))
        cb_type.pack(fill="x", pady=5)
        
        # 4. จำนวนวัน
        ttk.Label(form_frame, text="จำนวนวัน:", font=("", 10, "bold")).pack(anchor="w", pady=(10, 5))
        ent_days = ttk.Entry(form_frame, font=("", 10))
        ent_days.insert(0, str(record.get('num_days', 0)))
        ent_days.pack(fill="x", pady=5)
        
        # 5. หมายเหตุ
        ttk.Label(form_frame, text="สาเหตุ/หมายเหตุ:", font=("", 10, "bold")).pack(anchor="w", pady=(10, 5))
        ent_reason = ttk.Entry(form_frame, font=("", 10))
        ent_reason.insert(0, record.get('reason', ''))
        ent_reason.pack(fill="x", pady=5)
        
        # --- ฟังก์ชันบันทึก ---
        def save_edit():
            new_type = cb_type.get()
            new_reason = ent_reason.get()
            
            # เช็คค่าตัวเลข
            try:
                new_days = float(ent_days.get())
                if new_days <= 0: raise ValueError
            except:
                messagebox.showerror("ข้อผิดพลาด", "กรุณากรอก 'จำนวนวัน' เป็นตัวเลขที่มากกว่า 0")
                return
            
            # สั่ง update ลง Database
            # (ต้องมีฟังก์ชัน update_leave_record ใน hr_database.py ด้วย)
            success = hr_database.update_leave_record(leave_id, new_type, new_days, new_reason)
            
            if success:
                messagebox.showinfo("สำเร็จ", "บันทึกการแก้ไขเรียบร้อยแล้ว ✅")
                edit_win.destroy()     # ปิด Popup
                refresh_callback()     # รีโหลดตารางหลักใหม่
            else:
                messagebox.showerror("ล้มเหลว", "ไม่สามารถบันทึกข้อมูลได้\n(กรุณาตรวจสอบการเชื่อมต่อ Database)")

        # ปุ่มบันทึก
        ttk.Button(form_frame, text="💾 บันทึกการแก้ไข", command=save_edit, style="Success.TButton").pack(pady=30, fill="x", ipady=5)

    def _create_main_layout(self):
        """สร้างโครงสร้างหลักแบบ 2 คอลัมน์ (Master-Detail)"""
        
        list_container = ttk.Frame(self, width=400)
        list_container.pack(side="left", fill="y", padx=15, pady=10)
        
        form_container = ttk.Frame(self)
        form_container.pack(side="right", fill="both", expand=True, padx=(0, 15), pady=10)

        self.form_panel = form_container 
        self.list_panel = list_container

    def _build_list_panel(self):
        """สร้าง UI ด้านซ้าย (รายชื่อพนักงาน)"""
        
        search_frame = ttk.Frame(self.list_panel)
        search_frame.pack(fill="x", pady=(5, 10))
        
        ttk.Label(search_frame, text="🔍 ค้นหา:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, width=30, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.search_entry.bind("<KeyRelease>", self._search_employees)

        tree_frame = ttk.LabelFrame(self.list_panel, text="  เลือกพนักงาน  ", padding=10)
        tree_frame.pack(fill="both", expand=True)
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        self.employee_tree = ttk.Treeview(
            tree_container,
            columns=("id", "name", "position"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            height=25
        )
        self.employee_tree.heading("id", text="รหัส")
        self.employee_tree.heading("name", text="ชื่อ-นามสกุล")
        self.employee_tree.heading("position", text="ตำแหน่ง")

        self.employee_tree.column("id", width=60, anchor="center")
        self.employee_tree.column("name", width=180, anchor="w")
        self.employee_tree.column("position", width=120, anchor="w")
        
        self.employee_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.employee_tree.yview)
        
        self.employee_tree.bind("<<TreeviewSelect>>", self._on_employee_selected)

    def _build_form_panel(self):
        """สร้าง UI ด้านขวา (ฟอร์มกรอกข้อมูล) - แก้ไขปุ่มซ้ำและเพิ่มระบบจัดการประวัติ"""
        
        header_frame = ttk.Frame(self.form_panel)
        header_frame.pack(fill="x", pady=(5, 10))
        
        ttk.Label(header_frame, text="บันทึกข้อมูลสำหรับ:", font=("Segoe UI", 12)).pack(side="left")
        self.selected_emp_label = ttk.Label(header_frame, textvariable=self.selected_emp_name, 
                                              font=("Segoe UI", 12, "bold"), foreground="#007bff")
        self.selected_emp_label.pack(side="left", padx=5)
        self.selected_emp_name.set("--- (ยังไม่ได้เลือกพนักงาน) ---") 

        canvas = tk.Canvas(self.form_panel)
        scrollbar = ttk.Scrollbar(self.form_panel, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas, padding=(20, 10))
        
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # --- 1. กรอบสรุปประวัติประจำปี ---
        report_frame = ttk.LabelFrame(self.scroll_frame, text="  📊 สรุปประวัติ (Report)  ", padding=20)
        report_frame.pack(fill="x", pady=(0, 15))
        
        self.report_year_label = ttk.Label(report_frame, text="สรุปประจำปี ....", font=("Segoe UI", 11, "bold"))
        self.report_year_label.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        
        ttk.Label(report_frame, text="ลาป่วย:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="e", padx=(10,5))
        self.report_leave_sick_val = ttk.Label(report_frame, text="- วัน", font=("Segoe UI", 10, "bold"), width=10, anchor="w")
        self.report_leave_sick_val.grid(row=1, column=1, sticky="w")
        
        ttk.Label(report_frame, text="ลากิจ:", font=("Segoe UI", 10)).grid(row=2, column=0, sticky="e", padx=(10,5))
        self.report_leave_biz_val = ttk.Label(report_frame, text="- วัน", font=("Segoe UI", 10, "bold"), width=10, anchor="w")
        self.report_leave_biz_val.grid(row=2, column=1, sticky="w")
        
        ttk.Label(report_frame, text="ลาพักร้อน:", font=("Segoe UI", 10)).grid(row=3, column=0, sticky="e", padx=(10,5))
        self.report_leave_vac_val = ttk.Label(report_frame, text="- วัน", font=("Segoe UI", 10, "bold"), width=10, anchor="w")
        self.report_leave_vac_val.grid(row=3, column=1, sticky="w")
        
        ttk.Label(report_frame, text="รวมวันลา:", font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky="e", padx=(10,5), pady=5)
        self.report_leave_total_val = ttk.Label(report_frame, text="- วัน", font=("Segoe UI", 10, "bold"), width=10, anchor="w")
        self.report_leave_total_val.grid(row=4, column=1, sticky="w", pady=5)

        ttk.Label(report_frame, text="มาสาย (รวม):", font=("Segoe UI", 10)).grid(row=1, column=2, sticky="e", padx=(20,5))
        self.report_late_times_val = ttk.Label(report_frame, text="- ครั้ง", font=("Segoe UI", 10, "bold"), width=12, anchor="w")
        self.report_late_times_val.grid(row=1, column=3, sticky="w")
        
        ttk.Label(report_frame, text="รวมนาทีสาย:", font=("Segoe UI", 10)).grid(row=2, column=2, sticky="e", padx=(20,5))
        self.report_late_mins_val = ttk.Label(report_frame, text="- นาที", font=("Segoe UI", 10, "bold"), width=12, anchor="w")
        self.report_late_mins_val.grid(row=2, column=3, sticky="w")

        ttk.Label(report_frame, text="เฉลี่ยสาย:", font=("Segoe UI", 10)).grid(row=3, column=2, sticky="e", padx=(20,5))
        self.report_late_avg_val = ttk.Label(report_frame, text="- นาที/ครั้ง", font=("Segoe UI", 10, "bold"), width=12, anchor="w")
        self.report_late_avg_val.grid(row=3, column=3, sticky="w")
        
        ttk.Label(report_frame, text="ใบเตือน:", font=("Segoe UI", 10, "bold")).grid(row=4, column=2, sticky="e", padx=(20,5), pady=5)
        self.report_warn_val = ttk.Label(report_frame, text="- ครั้ง", font=("Segoe UI", 10, "bold"), width=12, anchor="w")
        self.report_warn_val.grid(row=4, column=3, sticky="w", pady=5)
        
        # --- 2. กรอบบันทึกการลา ---
        leave_frame = ttk.LabelFrame(self.scroll_frame, text="  📝 บันทึกการลา  ", padding=20)
        leave_frame.pack(fill="x", pady=(10, 15))
        
        row = 0
        ttk.Label(leave_frame, text="ลาตั้งแต่วันที่:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.att_leave_date = DateDropdown(leave_frame, font=("Segoe UI", 10))
        self.att_leave_date.grid(row=row, column=1, sticky="w", pady=10)

        self.att_leave_date.day_var.trace_add("write", self._on_end_date_changed)
        self.att_leave_date.month_var.trace_add("write", self._on_end_date_changed)
        self.att_leave_date.year_var.trace_add("write", self._on_end_date_changed)

        ttk.Label(leave_frame, text="ถึงวันที่:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.att_leave_date_end = DateDropdown(leave_frame, font=("Segoe UI", 10))
        self.att_leave_date_end.grid(row=row, column=3, sticky="w", pady=10)
        
        self.att_leave_date_end.day_var.trace_add("write", self._on_end_date_changed)
        self.att_leave_date_end.month_var.trace_add("write", self._on_end_date_changed)
        self.att_leave_date_end.year_var.trace_add("write", self._on_end_date_changed)
        
        row += 1
        ttk.Label(leave_frame, text="ประเภท:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.att_leave_type = ttk.Combobox(leave_frame, width=20, font=("Segoe UI", 10),
                                           values=["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาไม่รับค่าจ้าง", "ลาอื่นๆ"], state="readonly")
        self.att_leave_type.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(leave_frame, text="ระยะเวลา:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.att_leave_duration_type = ttk.Combobox(leave_frame, width=18, font=("Segoe UI", 10),
                                           values=["เต็มวัน (1.0)", "ครึ่งวัน (0.5)", "ระบุเวลา (ชม.)"], 
                                           state="readonly")
        self.att_leave_duration_type.grid(row=row, column=3, sticky="w", pady=10)
        self.att_leave_duration_type.set("เต็มวัน (1.0)")
        self.att_leave_duration_type.bind("<<ComboboxSelected>>", self._toggle_leave_time_entries)

        self.att_leave_time_frame = ttk.Frame(leave_frame)
        self.att_leave_time_frame.grid(row=row+1, column=3, sticky="w", pady=0, padx=0)
        
        ttk.Label(self.att_leave_time_frame, text="ตั้งแต่:").pack(side="left")
        self.att_leave_start_time = ttk.Combobox(self.att_leave_time_frame, values=self.time_options, 
                                                 width=6, font=("Segoe UI", 10))
        self.att_leave_start_time.pack(side="left", padx=5)
        
        ttk.Label(self.att_leave_time_frame, text="ถึง:").pack(side="left")
        self.att_leave_end_time = ttk.Combobox(self.att_leave_time_frame, values=self.time_options, 
                                               width=6, font=("Segoe UI", 10))
        self.att_leave_end_time.pack(side="left", padx=5)
        self.att_leave_time_frame.grid_remove() 
        
        row += 2 
        ttk.Label(leave_frame, text="เหตุผล:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.att_leave_reason = tk.Text(leave_frame, width=50, height=3, font=("Segoe UI", 10))
        self.att_leave_reason.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        
        row += 1
        # ✅ รวมกลุ่มปุ่มจัดการการลาไว้ที่เดียว
        leave_action_btn_frame = ttk.Frame(leave_frame)
        leave_action_btn_frame.grid(row=row, column=1, columnspan=3, sticky="e", pady=10)

        self.save_leave_btn = ttk.Button(leave_action_btn_frame, text="💾 บันทึกการลา", command=self._save_leave_record, 
                                       width=15, style="Primary.TButton", state="disabled") 
        self.save_leave_btn.pack(side="right", padx=5)

        ttk.Button(leave_action_btn_frame, text="📜 ดูประวัติ/แก้ไข/ลบ", 
                   command=self._open_history_window, width=18).pack(side="right", padx=5)

        # --- 3. กรอบบันทึกการมาสาย ---
        late_frame = ttk.LabelFrame(self.scroll_frame, text="  🏃 บันทึกการมาสาย  ", padding=20)
        late_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(late_frame, text="วันที่มาสาย:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.att_late_date = DateDropdown(late_frame, font=("Segoe UI", 10))
        self.att_late_date.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(late_frame, text="จำนวนนาที:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.att_late_minutes = ttk.Entry(late_frame, width=15, font=("Segoe UI", 10))
        self.att_late_minutes.grid(row=row, column=3, sticky="w", pady=10)
        row += 1
        ttk.Label(late_frame, text="เหตุผล:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.att_late_reason = tk.Text(late_frame, width=50, height=3, font=("Segoe UI", 10))
        self.att_late_reason.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        row += 1
        self.save_late_btn = ttk.Button(late_frame, text="💾 บันทึกการมาสาย", command=self._save_late_record, 
                                   width=20, style="Primary.TButton", state="disabled")
        self.save_late_btn.grid(row=row, column=1, columnspan=3, sticky="e", pady=10)

        # --- 4. กรอบบันทึกใบเตือน ---
        warn_frame = ttk.LabelFrame(self.scroll_frame, text="  📜 หนังสือเตือน  ", padding=20)
        warn_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(warn_frame, text="วันที่เตือน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.att_warn_date = DateDropdown(warn_frame, font=("Segoe UI", 10))
        self.att_warn_date.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(warn_frame, text="ระดับ:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        
        self.warning_options_list = [
            "ตักเตือนทางวาจา ครั้ง 1",
            "ตักเตือนทางวาจา ครั้ง 2",
            "ตักเตือนทางวาจา มากกว่า 2 ครั้ง",
            "หนังสือตักเตือนครั้งที่ 1",
            "หนังสือตักเตือนครั้งที่ 2",
            "หนังสือเลิกจ้าง"
        ]
        self.att_warn_level = ttk.Combobox(warn_frame, width=30, font=("Segoe UI", 10),
                                           values=self.warning_options_list, state="readonly")
        self.att_warn_level.grid(row=row, column=3, sticky="w", pady=10)
        row += 1
        ttk.Label(warn_frame, text="เหตุผล/คำอธิบาย:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.att_warn_reason = tk.Text(warn_frame, width=50, height=3, font=("Segoe UI", 10))
        self.att_warn_reason.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        row += 1
        self.att_warn_doc_label = ttk.Label(warn_frame, text="ไฟล์แนบ (ใบเตือน):", font=("Segoe UI", 10))
        self.att_warn_doc_label.grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        
        self.att_warn_file_display = ttk.Label(warn_frame, text="ยังไม่มีไฟล์อัปโหลด", font=("Segoe UI", 9, "italic"), foreground="gray", width=40, anchor="w")
        self.att_warn_file_display.grid(row=row, column=1, columnspan=2, sticky="w", pady=(10,0))
        
        self.att_warn_btn_frame = ttk.Frame(warn_frame) 
        self.att_warn_btn_frame.grid(row=row, column=3, sticky="w", pady=5)
        
        self.att_warn_upload_btn = ttk.Button(self.att_warn_btn_frame, text="อัปโหลด", command=self._att_upload_warning_doc, width=10, state="disabled")
        self.att_warn_upload_btn.pack(side="left", padx=(0, 5))
        
        self.att_warn_view_btn = ttk.Button(self.att_warn_btn_frame, text="ดูไฟล์", command=self._att_view_warning_doc, width=10, state="disabled")
        self.att_warn_view_btn.pack(side="left", padx=5)
        
        self.att_warn_delete_btn = ttk.Button(self.att_warn_btn_frame, text="ลบ", command=self._att_delete_warning_doc, width=5, state="disabled")
        self.att_warn_delete_btn.pack(side="left", padx=5)
        row += 1
        self.save_warn_btn = ttk.Button(warn_frame, text="💾 บันทึกใบเตือน", command=self._save_warning_record, 
                                   width=20, style="Primary.TButton", state="disabled")
        self.save_warn_btn.grid(row=row, column=1, columnspan=3, sticky="e", pady=10)

        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))

    # --- !! (5. เพิ่มฟังก์ชันใหม่นี้) !! ---

    def _toggle_leave_time_entries(self, event=None):
        """(เหมือนเดิม) ซ่อน/แสดง ช่องกรอกเวลาลา"""
        if self.att_leave_duration_type.get() == "ระบุเวลา (ชม.)":
            self.att_leave_time_frame.grid() 
        else:
            self.att_leave_time_frame.grid_remove() 
            self.att_leave_start_time.set("")
            self.att_leave_end_time.set("")

    def _on_end_date_changed(self, *args):
        """(แก้ไข V2) ปลดล็อกให้เลือก 'ครึ่งวัน' หรือ 'ระบุเวลา' ได้เสมอ แม้จะลาหลายวัน"""
        try:
            # ไม่ว่าจะเลือกวันเดียวหรือหลายวัน ให้เปิดให้เลือก Duration ได้เสมอ
            self.att_leave_duration_type.config(state="readonly")
            
            # เช็คเพิ่มเติม: ถ้าเลือกแบบ "ระบุเวลา" ให้โชว์ช่องเวลาด้วย
            self._toggle_leave_time_entries()

        except Exception:
            self.att_leave_duration_type.config(state="readonly")

    def _on_mousewheel(self, event, widget):
        try:
            widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass 

    def _load_employee_list(self):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
            
        try:
            all_employees = hr_database.load_all_employees()
            for emp in all_employees:
                display_name = f"{emp.get('fname','')} {emp.get('lname','')}"
                self.employee_tree.insert("", "end", iid=emp.get("id"), values=(
                    emp.get("id", ""),
                    display_name,
                    emp.get("position", "") 
                ))
        except Exception as e:
            messagebox.showerror("Load Error", f"ไม่สามารถโหลดรายชื่อพนักงานได้:\n{e}")

        self._load_att_warning_doc_status()

    def _on_employee_selected(self, event=None):
        selection = self.employee_tree.selection()
        if not selection:
            self.selected_emp_id.set("")
            self.selected_emp_name.set("--- (ยังไม่ได้เลือกพนักงาน) ---")
            self.save_leave_btn.config(state="disabled")
            self.save_late_btn.config(state="disabled")
            self.save_warn_btn.config(state="disabled")
            self.att_warn_upload_btn.config(state="disabled")
            self._clear_forms()
            self._load_attendance_report("") 
            self._load_att_warning_doc_status() 
            return

        item = self.employee_tree.item(selection[0])
        emp_id = item["values"][0]
        emp_name = item["values"][1]
        
        if emp_id:
            self.selected_emp_id.set(emp_id)
            self.selected_emp_name.set(f"{emp_name} (ID: {emp_id})")
            
            self.save_leave_btn.config(state="normal")
            self.save_late_btn.config(state="normal")
            self.save_warn_btn.config(state="normal")
            self.att_warn_upload_btn.config(state="normal")
            
            self._clear_forms()
            self._load_attendance_report(emp_id) 
            self._load_att_warning_doc_status() 
            
        else:
             self.selected_emp_id.set("")
             self.selected_emp_name.set("--- (ยังไม่ได้เลือกพนักงาน) ---")
             self.save_leave_btn.config(state="disabled")
             self.save_late_btn.config(state="disabled")
             self.save_warn_btn.config(state="disabled")
             self.att_warn_upload_btn.config(state="disabled")
             self._clear_forms()
             self._load_attendance_report("")
             self._load_att_warning_doc_status()

    def _search_employees(self, event=None):
        search_term = self.search_entry.get().lower()
        
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
            
        all_employees = hr_database.load_all_employees()
        for emp in all_employees:
            display_name = f"{emp.get('fname','')} {emp.get('lname','')}"
            
            if (search_term in emp.get("id", "").lower() or
                search_term in display_name.lower() or
                search_term in emp.get("position", "").lower()):
                
                self.employee_tree.insert("", "end", iid=emp.get("id"), values=(
                    emp.get("id", ""),
                    display_name,
                    emp.get("position", "") 
                ))
    
    def _clear_forms(self):
        """(Helper) ล้างค่าในฟอร์มกรอกข้อมูล 3 อัน (ด้านขวา)"""
        try:
            self.att_leave_date.clear()
            self.att_leave_type.set("")
            # --- !! (6. แก้ไข: ล้าง UI การลา) !! ---
            self.att_leave_duration_type.set("เต็มวัน (1.0)")
            self._toggle_leave_time_entries() # (สั่งซ่อน)
            # --- (จบส่วนแก้ไข) ---
            self.att_leave_reason.delete("1.0", tk.END)
            
            self.att_late_date.clear()
            self.att_late_minutes.delete(0, tk.END)
            self.att_late_reason.delete("1.0", tk.END)

            self.att_warn_date.clear()
            self.att_warn_level.set("") 
            self.att_warn_reason.delete("1.0", tk.END)
            
            self._load_att_warning_doc_status()
        except AttributeError:
            pass 
    
    def _load_attendance_report(self, emp_id):
        if not emp_id:
            self.report_year_label.config(text="สรุปประจำปี ....")
            self.report_leave_sick_val.config(text="- วัน")
            self.report_leave_biz_val.config(text="- วัน")
            self.report_leave_vac_val.config(text="- วัน")
            self.report_leave_total_val.config(text="- วัน")
            self.report_late_times_val.config(text="- ครั้ง")
            self.report_late_mins_val.config(text="- นาที")
            self.report_late_avg_val.config(text="- นาที/ครั้ง")
            self.report_warn_val.config(text="- ครั้ง")
            return
            
        current_year = datetime.now().year
        
        summary = hr_database.get_attendance_summary(emp_id, current_year)
        
        try:
            self.report_year_label.config(text=f"สรุปประจำปี {current_year + 543}")
            
            self.report_leave_sick_val.config(text=f"{summary['leave_by_type'].get('ลาป่วย', 0.0):.1f} วัน")
            self.report_leave_biz_val.config(text=f"{summary['leave_by_type'].get('ลากิจ', 0.0):.1f} วัน")
            self.report_leave_vac_val.config(text=f"{summary['leave_by_type'].get('ลาพักร้อน', 0.0):.1f} วัน")
            self.report_leave_total_val.config(text=f"{summary['total_leave_days']:.1f} วัน")
            
            self.report_late_times_val.config(text=f"{summary['total_late_times']} ครั้ง")
            self.report_late_mins_val.config(text=f"{summary['total_late_minutes']} นาที")
            self.report_late_avg_val.config(text=f"{summary['avg_late_minutes']:.1f} นาที/ครั้ง")
            
            self.report_warn_val.config(text=f"{summary['total_warnings']} ครั้ง")

        except AttributeError:
            print("DEBUG: Labels for report are not created yet.")
        except Exception as e:
            messagebox.showerror("Report UI Error", f"ไม่สามารถอัปเดตหน้า Report ได้:\n{e}")
            
    # --- !! (7. "ผ่าตัด" ฟังก์ชันนี้) !! ---
    def _save_leave_record(self):
        emp_id = self.selected_emp_id.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาเลือกพนักงานก่อน")
            return
            
        try:
            start_date = self.att_leave_date.get_date() 
            end_date = self.att_leave_date_end.get_date()
        except Exception:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่ลาให้ครบถ้วน")
            return
            
        if not start_date:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือก 'ลาตั้งแต่วันที่'")
            return

        leave_type = self.att_leave_type.get()
        duration_type = self.att_leave_duration_type.get()
        reason = self.att_leave_reason.get("1.0", "end-1c")

        if not leave_type:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือก 'ประเภท' การลา")
            return

        # --- 1. คำนวณ "จำนวนวันต่อครั้ง" (Per Day Amount) ---
        days_per_record = 1.0
        l_start = None
        l_end = None

        try:
            if duration_type == "เต็มวัน (1.0)":
                days_per_record = 1.0
            
            elif duration_type == "ครึ่งวัน (0.5)":
                days_per_record = 0.5
            
            elif duration_type == "ระบุเวลา (ชม.)":
                start_str = self.att_leave_start_time.get()
                end_str = self.att_leave_end_time.get()
                
                if not start_str or not end_str:
                    messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาระบุเวลา เริ่มต้น และ สิ้นสุด")
                    return
                
                l_start = datetime.strptime(start_str, '%H:%M').time()
                l_end = datetime.strptime(end_str, '%H:%M').time()
                
                if l_start >= l_end:
                    messagebox.showwarning("เวลาผิดพลาด", "เวลาสิ้นสุด ต้องมากกว่าเวลาเริ่มต้น")
                    return
                
                # คำนวณสัดส่วนตาม 8 ชม. งาน
                dummy_date = datetime.today().date()
                duration = datetime.combine(dummy_date, l_end) - datetime.combine(dummy_date, l_start)
                hours_total = duration.total_seconds() / 3600
                days_per_record = round(hours_total / 8.0, 4)

        except Exception as e:
            messagebox.showerror("ข้อมูลผิดพลาด", f"รูปแบบเวลาไม่ถูกต้อง: {e}")
            return

        # --- 2. คำนวณยอดรวมทั้งหมด (Total Days) เพื่อเช็คโควตา ---
        total_req_days = days_per_record # กรณีวันเดียว
        
        if end_date and end_date >= start_date:
            day_count = (end_date - start_date).days + 1
            total_req_days = days_per_record * day_count
        
        # --- 3. ตรวจสอบโควตา ---
        current_year = start_date.year
        is_pass, msg, remaining = hr_database.check_leave_quota_status(emp_id, current_year, leave_type, total_req_days)
        
        if not is_pass:
            if not messagebox.askyesno("โควตาหมด!", 
                                       f"{msg}\n\nคุณต้องการบันทึกเป็น 'ลาโดยไม่รับค่าจ้าง (Leave without Pay)' แทนหรือไม่?"):
                return
            leave_type = f"{leave_type} (เกินสิทธิ์)" # หรือเปลี่ยนเป็น "ลาไม่รับค่าจ้าง" ตาม Logic บริษัท

        # --- 4. เริ่มบันทึก (วนลูปตามช่วงวันที่เลือก) ---
        try:
            if end_date and end_date >= start_date: # แบบช่วง (Range)
                current_date = start_date
                success_count = 0
                while current_date <= end_date:
                    # ข้ามวันอาทิตย์ (ตัวอย่าง: ถ้าบริษัทหยุดวันอาทิตย์)
                    if current_date.weekday() != 6: 
                        # บันทึกทีละวัน ด้วยยอด days_per_record (เช่น 0.5)
                        hr_database.add_employee_leave(
                            emp_id, current_date, leave_type, days_per_record, reason, l_start, l_end
                        )
                        success_count += 1
                    current_date += timedelta(days=1)
                
                messagebox.showinfo("สำเร็จ", f"บันทึกการลาแบบช่วงเรียบร้อย ({success_count} วันทำงาน)")

            else: # แบบวันเดียว (Single Day)
                success = hr_database.add_employee_leave(
                    emp_id, start_date, leave_type, days_per_record, reason, l_start, l_end
                )
                if success: messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลการลาเรียบร้อยแล้ว")
            
            self._clear_forms()
            self._load_attendance_report(emp_id)

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}")

    def _save_late_record(self):
        emp_id = self.selected_emp_id.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาเลือกพนักงานก่อน")
            return
            
        late_date = self.att_late_date.get_date()
        if not late_date:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่มาสายให้ครบถ้วน")
            return
            
        minutes_str = self.att_late_minutes.get()
        reason = self.att_late_reason.get("1.0", "end-1c")
        if not minutes_str:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกจำนวนนาทีที่สาย")
            return
        try:
            minutes_late = int(minutes_str)
        except ValueError:
            messagebox.showwarning("ข้อมูลผิดพลาด", "จำนวนนาทีต้องเป็นตัวเลข")
            return
            
        success = hr_database.add_employee_late(emp_id, late_date, minutes_late, reason)
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลการมาสายเรียบร้อยแล้ว")
            
            self.att_late_date.clear()
            self.att_late_minutes.delete(0, tk.END)
            self.att_late_reason.delete("1.0", tk.END)
            
            self._load_attendance_report(emp_id) 

    def _save_warning_record(self):
        emp_id = self.selected_emp_id.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาเลือกพนักงานก่อน")
            return
        
        warn_date = self.att_warn_date.get_date()
        if not warn_date:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่เตือนให้ครบถ้วน")
            return
            
        reason = self.att_warn_reason.get("1.0", "end-1c")
        level = self.att_warn_level.get() 
        
        if not reason or not level: 
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกเหตุผล และเลือกระดับการเตือน")
            return
            
        success = hr_database.add_employee_warning(emp_id, warn_date, reason, level)
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลใบเตือนเรียบร้อยแล้ว")
            
            self.att_warn_date.clear() 
            self.att_warn_reason.delete("1.0", tk.END)
            self.att_warn_level.set("") 
            
            self._load_attendance_report(emp_id)

    def _load_att_warning_doc_status(self):
        emp_id = self.selected_emp_id.get() 
        
        if not emp_id:
            self.att_warn_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.att_warn_view_btn.config(state="disabled")
            self.att_warn_delete_btn.config(state="disabled")
            self.att_warn_upload_btn.config(text="อัปโหลด", state="disabled") 
            self.current_att_warning_doc_id = None
            self.current_att_warning_file_path = None
            return

        doc_info = hr_database.get_document_by_description(emp_id, "เอกสารใบเตือน")
        
        if doc_info and doc_info.get('file_path'):
            self.current_att_warning_doc_id = doc_info['doc_id']
            self.current_att_warning_file_path = doc_info['file_path']
            filename = os.path.basename(doc_info['file_path'])
            self.att_warn_file_display.config(text=filename, foreground="#007bff")
            self.att_warn_view_btn.config(state="normal")
            self.att_warn_delete_btn.config(state="normal")
            self.att_warn_upload_btn.config(text="แทนที่", state="normal")
        else:
            self.att_warn_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.att_warn_view_btn.config(state="disabled")
            self.att_warn_delete_btn.config(state="disabled")
            self.att_warn_upload_btn.config(text="อัปโหลด", state="normal") 
            self.current_att_warning_doc_id = None
            self.current_att_warning_file_path = None

    def _att_upload_warning_doc(self):
        emp_id = self.selected_emp_id.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาเลือกพนักงานก่อนอัปโหลดไฟล์")
            return
        if not os.path.exists(NETWORK_UPLOAD_PATH):
            messagebox.showerror("Network Error", f"ไม่พบโฟลเดอร์สำหรับอัปโหลดที่ Server:\n{NETWORK_UPLOAD_PATH}")
            return

        source_path = filedialog.askopenfilename(
            title="เลือกไฟล์ (PDF, JPG, PNG)",
            filetypes=[("เอกสาร", "*.pdf *.jpg *.jpeg *.png"), ("All Files", "*.*")]
        )
        if not source_path: return

        original_filename, file_extension = os.path.splitext(os.path.basename(source_path))
        timestamp = int(datetime.now().timestamp())
        unique_filename = f"{emp_id}_warning_{timestamp}{file_extension}"
        destination_path = os.path.join(NETWORK_UPLOAD_PATH, unique_filename)

        try:
            shutil.copy(source_path, destination_path)
            
            if self.current_att_warning_doc_id:
                old_path = self.current_att_warning_file_path
                hr_database.delete_document(self.current_att_warning_doc_id)
                if old_path and os.path.exists(old_path):
                    try: os.remove(old_path)
                    except Exception as e: print(f"Warning: ไม่สามารถลบไฟล์เก่า {old_path} ได้: {e}")

            success = hr_database.add_employee_document(emp_id, "เอกสารใบเตือน", destination_path)
            
            if success:
                messagebox.showinfo("สำเร็จ", f"อัปโหลดไฟล์ {unique_filename} เรียบร้อย")
                self._load_att_warning_doc_status()
            
        except Exception as e:
            messagebox.showerror("Upload Failed", f"ไม่สามารถอัปโหลดไฟล์ได้:\n{e}")

    def _att_view_warning_doc(self):
        if not self.current_att_warning_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่ระบุ")
            return
        try:
            if not os.path.exists(self.current_att_warning_file_path):
                 messagebox.showerror("เปิดไฟล์ล้มเหลว", f"ไม่พบไฟล์ที่: {self.current_att_warning_file_path}")
                 return
            os.startfile(self.current_att_warning_file_path)
        except Exception as e:
            messagebox.showerror("เปิดไฟล์ล้มเหลว", f"ไม่สามารถเปิดไฟล์ได้:\n{e}")

    def _att_delete_warning_doc(self):
        if not self.current_att_warning_doc_id or not self.current_att_warning_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่จะลบ")
            return
        
        filename = os.path.basename(self.current_att_warning_file_path)
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบไฟล์ '{filename}' ใช่หรือไม่?"):
            return

        try:
            success_db = hr_database.delete_document(self.current_att_warning_doc_id)
            if success_db:
                if os.path.exists(self.current_att_warning_file_path):
                    try: os.remove(self.current_att_warning_file_path)
                    except Exception as e: messagebox.showwarning("ลบไฟล์สำเร็จ (แต่)", f"ลบข้อมูลไฟล์ออกจากระบบแล้ว\nแต่ไม่สามารถลบไฟล์จริงออกจาก Server ได้: {e}")
                else: messagebox.showinfo("สำเร็จ", "ลบข้อมูลไฟล์เรียบร้อย")
                
                self._load_att_warning_doc_status()
            else:
                messagebox.showerror("ล้มเหลว", "ไม่สามารถลบข้อมูลออกจากฐานข้อมูลได้")
        except Exception as e:
            messagebox.showerror("ลบไฟล์ล้มเหลว", f"เกิดข้อผิดพลาด: {e}")

    def _load_employee_dropdown(self):
        self._load_employee_list()


# เพิ่มใน attendance_module.py

class AttendanceHistoryWindow(tk.Toplevel):
    def __init__(self, parent, db_callback):
        super().__init__(parent)
        self.title("📜 จัดการประวัติการลา/เวลาทำงาน")
        self.geometry("900x500")
        self.db_callback = db_callback # สำหรับส่งค่ากลับไปแก้ไขที่หน้าหลัก
        
        self._create_widgets()
        self._load_data()

    def _create_widgets(self):
        # ส่วน Filter
        filter_frame = ttk.Frame(self, padding=10)
        filter_frame.pack(fill="x")
        
        ttk.Label(filter_frame, text="ค้นหาประวัติการลาทั้งหมด").pack(side="left")
        ttk.Button(filter_frame, text="🔄 รีเฟรช", command=self._load_data).pack(side="right")

        # ตารางแสดงข้อมูล
        columns = ("id", "emp_id", "name", "date", "type", "days", "reason")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        
        self.tree.heading("id", text="ID")
        self.tree.heading("emp_id", text="รหัสพนักงาน")
        self.tree.heading("name", text="ชื่อ-สกุล")
        self.tree.heading("date", text="วันที่ลา")
        self.tree.heading("type", text="ประเภท")
        self.tree.heading("days", text="จำนวนวัน")
        
        self.tree.column("id", width=50)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # ปุ่มจัดการ
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="❌ ลบรายการที่เลือก", command=self._delete_item).pack(side="right", padx=5)
        ttk.Label(self, text="💡 ดับเบิลคลิกเพื่อแก้ไขรายการ (ฟีเจอร์นี้จะส่งข้อมูลกลับไปหน้าหลัก)", foreground="gray").pack(pady=5)

    def _load_data(self):
        # ล้างตาราง
        for i in self.tree.get_children(): self.tree.delete(i)
        # ดึงข้อมูลจาก DB
        history = hr_database.get_employee_leave_history()
        for item in history:
            self.tree.insert("", "end", values=(
                item['leave_id'], item['emp_id'], f"{item['fname']} {item['lname']}",
                item['leave_date'], item['leave_type'], item['num_days'], item['reason']
            ))

    def _delete_item(self):
        selected = self.tree.selection()
        if not selected: return
        
        leave_id = self.tree.item(selected[0])['values'][0]
        if messagebox.askyesno("ยืนยัน", "คุณต้องการลบรายการลานี้ใช่หรือไม่?"):
            if hr_database.delete_leave_record(leave_id):
                messagebox.showinfo("สำเร็จ", "ลบข้อมูลเรียบร้อยแล้ว")
                self._load_data()