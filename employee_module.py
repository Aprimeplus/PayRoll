# (ไฟล์: employee_module.py)
# (เวอร์ชันอัปเกรด - เปลี่ยน DateEntry เป็น DateDropdown ทั้งหมด)
# (เวอร์ชันอัปเกรด - รองรับ "ลารายชั่วโมง" ในแท็บ 5)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog # <--- (เพิ่ม simpledialog)
from datetime import datetime
import pandas as pd             
from fpdf import FPDF          
import os
import hr_database
import json
from custom_widgets import DateDropdown # <--- (Import เข้ามา)
import shutil
from datetime import datetime, timedelta, time # <--- (เพิ่ม time)

NETWORK_UPLOAD_PATH = r"\\192.168.1.51\HR_System_Documents"

class EmployeeModule(ttk.Frame):

    def _toggle_sales_options(self):
        """ซ่อน/แสดง ตัวเลือก Sale ตาม Checkbox"""
        if self.is_sales_var.get():
            self.sales_options_frame.grid() # แสดง
        else:
            self.sales_options_frame.grid_remove() # ซ่อน
            self.sale_type_var.set("")
            self.commission_plan_var.set("")
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user 

        loaded_settings = hr_database.load_allowance_settings()
        self.welfare_options = [item['name'] for item in loaded_settings]
        
        # (สร้าง List เวลา 00:00 - 23:45)
        self.time_options = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]

        # (เราลบ _configure_styles() ออกไปแล้ว)
        
        # (ย้าย score_validator มาไว้ที่นี่)
        self.score_validator = (self.register(self._validate_score), '%P')

        # === พื้นที่เนื้อหา ===
        self._create_top_bar() 

        self.content_area = ttk.Frame(self)
        self.content_area.pack(fill="both", expand=True)

        # --- (นี่คือส่วนที่ขาดหายไปในโค้ดของคุณ) ---
        self.list_page = ttk.Frame(self.content_area)
        self.list_page.grid(row=0, column=0, sticky="nsew")

        self.form_page = ttk.Frame(self.content_area)
        self.form_page.grid(row=0, column=0, sticky="nsew")

        self.approval_page = ttk.Frame(self.content_area)
        self.approval_page.grid(row=0, column=0, sticky="nsew")

        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self._build_list_page()
        self._build_form_page_with_tabs()
        
        if self.current_user.get("role") == 'approver':
            self._build_approval_page() 

        self.current_warning_doc_id = None
        self.current_warning_file_path = None

        self._show_list_page()

    def _create_top_bar(self):
        """สร้างแถบบนที่สวยงาม"""
        top_bar = ttk.Frame(self)
        top_bar.pack(fill="x", pady=(0, 10))
        
        title_frame = ttk.Frame(top_bar)
        title_frame.pack(side="left", padx=15, pady=12)
        ttk.Label(title_frame, text="👥 จัดการข้อมูลพนักงาน", style="Header.TLabel").pack(side="left")
        
        btn_frame = ttk.Frame(top_bar)
        btn_frame.pack(side="right", padx=15, pady=8)
        ttk.Button(btn_frame, text="➕ สร้างพนักงานใหม่", command=self._show_form_page_new,
                   style="Success.TButton").pack(side="right")

    def _show_list_page(self):
        """ยกหน้ารายชื่อ (List Page) ขึ้นมาแสดง"""
        self.list_page.tkraise()
        all_employees = hr_database.load_all_employees() 
        self.update_employee_list(all_employees) 

    def _show_form_page_new(self):
        self.clear_form()
        self.emp_id_entry.focus()
        self.form_page.tkraise()

    def _load_and_show_form(self):
        selection = self.employee_tree.selection()
        if not selection:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกพนักงานที่ต้องการดูข้อมูล")
            return

        item = self.employee_tree.item(selection[0])
        emp_id = item["values"][0]
        employee_data = hr_database.load_single_employee(emp_id)
        
        if not employee_data:
            messagebox.showerror("Load Error", f"ไม่พบข้อมูลพนักงาน ID: {emp_id}")
            return

        self.load_employee_data(employee_data) # โหลดข้อมูลเข้าฟอร์ม
        self.form_page.tkraise()

    def _build_list_page(self):
        """สร้างหน้ารายชื่อพนักงาน (List Page) แบบจัดลำดับใหม่ (Fix ปุ่มตกจอ)"""
        
        # 1. สร้าง Container หลัก
        main_container = ttk.Frame(self.list_page)
        main_container.pack(fill="both", expand=True, padx=15, pady=10)
        
        # 2. ส่วนค้นหา (Search Bar) ด้านบน
        search_frame = ttk.Frame(main_container)
        search_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(search_frame, text="🔍 ค้นหา:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, width=40, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", padx=(0, 10))
        self.search_entry.bind("<Return>", lambda event: self._search_employees()) 
        
        ttk.Button(search_frame, text="ค้นหา", width=10, command=self._search_employees).pack(side="left")
        ttk.Button(search_frame, text="ล้าง", width=10, command=self._clear_search).pack(side="left", padx=5)
        ttk.Button(search_frame, text="💾 Export Excel", command=self._export_to_excel).pack(side="right", padx=5)
        ttk.Button(search_frame, text="📂 Mass Import", command=self._open_mass_import_window).pack(side="right", padx=5)
        ttk.Button(search_frame, text="💾 Export Excel", command=self._export_to_excel).pack(side="right", padx=10)
        
        # 3. กรอบตารางและปุ่มสั่งงาน
        tree_frame = ttk.LabelFrame(main_container, text=" รายชื่อพนักงาน ", padding=15)
        tree_frame.pack(fill="both", expand=True)

        # --- [จุดแก้ไขสำคัญ 1] วางปุ่มด้านล่าง (Action Panel) ก่อน! ---
        # เพื่อจองพื้นที่ด้านล่างไว้เสมอ ไม่ให้ตารางดันจนตกจอ
        action_panel = ttk.LabelFrame(tree_frame, text=" คำสั่งจัดการ ", padding=10)
        action_panel.pack(side="bottom", fill="x", pady=(10, 0))

        btn_edit = ttk.Button(action_panel, text="📝 ดู/แก้ไขข้อมูล", command=self._load_and_show_form, width=20)
        btn_edit.pack(side="left", padx=5)
        
        btn_delete = ttk.Button(action_panel, text="🗑️ ลบพนักงาน", command=self.delete_employee, width=15)
        btn_delete.pack(side="left", padx=5)

        # --- [จุดแก้ไขสำคัญ 2] วางสรุปยอด (Summary) ไว้เหนือปุ่มกด ---
        summary_frame = ttk.Frame(tree_frame)
        summary_frame.pack(side="bottom", fill="x", pady=(5, 0))
        self.summary_label = ttk.Label(summary_frame, text="📊 จำนวนพนักงานทั้งหมด: 0 คน", 
                                     font=("Segoe UI", 9), foreground="#7f8c8d")
        self.summary_label.pack(side="left", padx=5)

        # --- [จุดแก้ไขสำคัญ 3] วางตาราง (Treeview) เป็นลำดับสุดท้าย ---
        # และลด height เหลือ 10 (เพื่อให้มันไม่สูงเกินจอ Notebook)
        # expand=True จะทำให้มันยืดเต็มพื้นที่ที่เหลือเองครับ
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(side="top", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        self.employee_tree = ttk.Treeview(
            tree_container,
            columns=("id", "name", "phone", "position", "department", "status", "id_card", "salary"), 
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            height=10 # <--- ลดลงจาก 20 เหลือ 10 เพื่อเซฟที่
        )
        
        # ตั้งค่าหัวตาราง (เหมือนเดิม)
        self.employee_tree.heading("id", text="รหัส")
        self.employee_tree.heading("name", text="ชื่อ-นามสกุล")
        self.employee_tree.heading("phone", text="เบอร์โทร")
        self.employee_tree.heading("position", text="ตำแหน่ง")
        self.employee_tree.heading("department", text="ฝ่าย")
        self.employee_tree.heading("status", text="สถานะ")
        self.employee_tree.heading("id_card", text="บัตรประชาชน")
        self.employee_tree.heading("salary", text="เงินเดือน")
        
        self.employee_tree.column("id", width=80, anchor="center")
        self.employee_tree.column("name", width=220, anchor="w")
        self.employee_tree.column("phone", width=120, anchor="center")
        self.employee_tree.column("position", width=180, anchor="w")
        self.employee_tree.column("department", width=150, anchor="w")
        self.employee_tree.column("status", width=120, anchor="center")
        self.employee_tree.column("id_card", width=150, anchor="center")
        self.employee_tree.column("salary", width=100, anchor="e")
        
        self.employee_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.employee_tree.yview)
        scrollbar_x.config(command=self.employee_tree.xview)
        
        self.employee_tree.bind("<Double-1>", lambda event: self._load_and_show_form())

    def _open_mass_import_window(self):
        """เปิดหน้าต่างเลือกไฟล์ Excel เพื่ออัปเดตข้อมูล (Mass Upsert)"""
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel ข้อมูลพนักงาน (Mass Update)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        
        if not file_path: return
        
        if not messagebox.askyesno("ยืนยันการนำเข้า", 
                                   "ระบบจะทำการ 'ทับข้อมูลเดิม' ด้วยข้อมูลใน Excel\n"
                                   "กรุณาตรวจสอบว่ารหัสพนักงานถูกต้อง\n\nต้องการดำเนินการต่อหรือไม่?"):
            return
            
        # เรียกฟังก์ชันที่สร้างไว้ใน hr_database
        # (ต้องมั่นใจว่าแก้ hr_database.py ไปแล้วนะครับ)
        try:
            count, msg = hr_database.mass_upsert_employees_from_excel(file_path)
            
            if count > 0:
                messagebox.showinfo("สำเร็จ", msg)
                self._show_list_page() # รีเฟรชหน้ารายชื่อใหม่
            else:
                messagebox.showwarning("แจ้งเตือน", msg)
        except AttributeError:
            messagebox.showerror("Error", "ไม่พบฟังก์ชัน mass_upsert ใน hr_database\nกรุณาอัปเดตไฟล์ hr_database.py ก่อน")

    def _search_employees(self):
        search_term = self.search_entry.get().strip()
        if not search_term:
            results = hr_database.load_all_employees()
        else:
            results = hr_database.search_employees(search_term)
        self.update_employee_list(results)

    def _clear_search(self):
        self.search_entry.delete(0, tk.END)
        all_employees = hr_database.load_all_employees()
        self.update_employee_list(all_employees)

    def _build_form_page_with_tabs(self):
        header_frame = ttk.Frame(self.form_page)
        header_frame.pack(fill="x", padx=15, pady=(10, 0))
        
        ttk.Label(header_frame, text="📋 แบบฟอร์มข้อมูลพนักงาน", style="Header.TLabel").pack(side="left")
        ttk.Button(header_frame, text="← กลับ", command=self._show_list_page, width=12).pack(side="right")

        self.notebook = ttk.Notebook(self.form_page)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=10)

        tab1 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab1, text="  👤 ข้อมูลส่วนตัว  ")
        self._build_personal_tab(tab1)

        tab2 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab2, text="  💼 ข้อมูลการจ้าง  ")
        self._build_employment_tab(tab2)

        tab_assets = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab_assets, text="  💻 ทรัพย์สินที่รับผิดชอบ  ")
        self._build_assets_tab(tab_assets)

        tab3 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab3, text="  💰 เงินเดือน/สวัสดิการ  ")
        self._build_salary_tab(tab3)

        tab4 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab4, text="  🏦 ธนาคารและอื่นๆ  ")
        self._build_bank_tab(tab4)

        tab5 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab5, text="  🕒 ประวัติ ลา/สาย/เตือน  ")
        self._build_attendance_tab(tab5)

        self._create_form_buttons()

    def _build_personal_tab(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # --- ข้อมูลพื้นฐาน ---
        basic_frame = ttk.LabelFrame(scroll_frame, text="  ข้อมูลพื้นฐาน  ", padding=20)
        basic_frame.pack(fill="x", pady=(0, 15))

        row = 0
        ttk.Label(basic_frame, text="รหัสพนักงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        self.emp_id_entry = ttk.Entry(basic_frame, width=25, font=("Segoe UI", 10))
        self.emp_id_entry.grid(row=row, column=1, sticky="w", pady=8)
        ttk.Label(basic_frame, text="*", foreground="red").grid(row=row, column=2, sticky="w")

        row += 1
        ttk.Label(basic_frame, text="คำนำหน้า:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        
        prefix_frame = ttk.Frame(basic_frame)
        prefix_frame.grid(row=row, column=1, sticky="w", pady=8)
        
        self.combo_prefix = ttk.Combobox(prefix_frame, values=["นาย", "นาง", "นางสาว", "อื่นๆ"], state="readonly", width=10, font=("Segoe UI", 10))
        self.combo_prefix.pack(side="left")
        self.combo_prefix.bind("<<ComboboxSelected>>", self._on_prefix_change)
        
        self.entry_prefix_other = ttk.Entry(prefix_frame, width=15, font=("Segoe UI", 10))

        row += 1
        ttk.Label(basic_frame, text="ชื่อ:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        self.fname_entry = ttk.Entry(basic_frame, width=30, font=("Segoe UI", 10))
        self.fname_entry.grid(row=row, column=1, sticky="w", pady=8)
        ttk.Label(basic_frame, text="*", foreground="red").grid(row=row, column=2, sticky="w")

        ttk.Label(basic_frame, text="ชื่อเล่น:", font=("Segoe UI", 10)).grid(row=row, column=3, sticky="e", padx=(20, 10), pady=8)
        self.nickname_entry = ttk.Entry(basic_frame, width=20, font=("Segoe UI", 10))
        self.nickname_entry.grid(row=row, column=4, sticky="w", pady=8)

        row += 1
        ttk.Label(basic_frame, text="นามสกุล:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        self.lname_entry = ttk.Entry(basic_frame, width=30, font=("Segoe UI", 10))
        self.lname_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=8)

        row += 1
        ttk.Label(basic_frame, text="วันเกิด:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        birth_frame = ttk.Frame(basic_frame)
        birth_frame.grid(row=row, column=1, columnspan=4, sticky="w", pady=8)
        
        self.birth_entry = DateDropdown(birth_frame, font=("Segoe UI", 10))
        self.birth_entry.pack(side="left")
        ttk.Button(birth_frame, text="คำนวณอายุ", command=self.calculate_age, width=12).pack(side="left", padx=(10, 5))
        self.age_label = ttk.Label(birth_frame, text="-", font=("Segoe UI", 10, "bold"), foreground="#27ae60")
        self.age_label.pack(side="left", padx=10)
        
        self.birth_entry.day_var.trace_add("write", self.calculate_age)
        self.birth_entry.month_var.trace_add("write", self.calculate_age)
        self.birth_entry.year_var.trace_add("write", self.calculate_age)

        row += 1
        ttk.Label(basic_frame, text="เลขบัตรประชาชน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        self.id_card_entry = ttk.Entry(basic_frame, width=30, font=("Segoe UI", 10))
        self.id_card_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=8)

        row += 1
        ttk.Label(basic_frame, text="เบอร์โทรศัพท์:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=8)
        self.phone_entry = ttk.Entry(basic_frame, width=25, font=("Segoe UI", 10))
        self.phone_entry.grid(row=row, column=1, sticky="w", pady=8)

        # --- (!!! เพิ่มใหม่: บุคคลติดต่อฉุกเฉิน & อ้างอิง !!!) ---
        contact_frame = ttk.LabelFrame(scroll_frame, text="  ☎️ บุคคลติดต่อฉุกเฉิน & อ้างอิง  ", padding=20)
        contact_frame.pack(fill="x", pady=(0, 15))
        
        # (ซ้าย: ฉุกเฉิน)
        ttk.Label(contact_frame, text="[ กรณีฉุกเฉิน ]", font=("Segoe UI", 10, "bold"), foreground="red").grid(row=0, column=0, sticky="w")
        
        ttk.Label(contact_frame, text="ชื่อ-สกุล:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.emergency_name = ttk.Entry(contact_frame, width=25)
        self.emergency_name.grid(row=1, column=1, sticky="w", pady=5)
        
        ttk.Label(contact_frame, text="เบอร์โทร:", font=("Segoe UI", 10)).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.emergency_phone = ttk.Entry(contact_frame, width=20)
        self.emergency_phone.grid(row=2, column=1, sticky="w", pady=5)
        
        ttk.Label(contact_frame, text="ความสัมพันธ์:", font=("Segoe UI", 10)).grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.emergency_relation = ttk.Entry(contact_frame, width=20)
        self.emergency_relation.grid(row=3, column=1, sticky="w", pady=5)

        # (ขวา: อ้างอิง)
        ttk.Label(contact_frame, text="[ บุคคลอ้างอิง ]", font=("Segoe UI", 10, "bold"), foreground="blue").grid(row=0, column=2, sticky="w", padx=(40,0))
        
        ttk.Label(contact_frame, text="ชื่อ-สกุล:", font=("Segoe UI", 10)).grid(row=1, column=2, sticky="e", padx=5, pady=5)
        self.ref_name = ttk.Entry(contact_frame, width=25)
        self.ref_name.grid(row=1, column=3, sticky="w", pady=5)
        
        ttk.Label(contact_frame, text="เบอร์โทร:", font=("Segoe UI", 10)).grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.ref_phone = ttk.Entry(contact_frame, width=20)
        self.ref_phone.grid(row=2, column=3, sticky="w", pady=5)
        
        ttk.Label(contact_frame, text="ความสัมพันธ์:", font=("Segoe UI", 10)).grid(row=3, column=2, sticky="e", padx=5, pady=5)
        self.ref_relation = ttk.Entry(contact_frame, width=20)
        self.ref_relation.grid(row=3, column=3, sticky="w", pady=5)
        # -------------------------------------------------------

        # --- ที่อยู่ ---
        address_frame = ttk.LabelFrame(scroll_frame, text="  ที่อยู่  ", padding=20)
        address_frame.pack(fill="both", expand=True, pady=(0, 15))

        ttk.Label(address_frame, text="ที่อยู่ตามบัตรประชาชน:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="ne", padx=(0, 10), pady=8)
        self.address_text = tk.Text(address_frame, width=50, height=4, font=("Segoe UI", 10))
        self.address_text.grid(row=0, column=1, sticky="w", pady=8)

        ttk.Label(address_frame, text="ที่อยู่ปัจจุบัน:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="ne", padx=(0, 10), pady=8)
        self.current_address_text = tk.Text(address_frame, width=50, height=4, font=("Segoe UI", 10))
        self.current_address_text.grid(row=1, column=1, sticky="w", pady=8)
        
        copy_btn = ttk.Button(address_frame, text="📋 คัดลอกจากที่อยู่บัตร", width=20,
                             command=lambda: self.current_address_text.delete("1.0", tk.END) or 
                                           self.current_address_text.insert("1.0", self.address_text.get("1.0", "end-1c")))
        copy_btn.grid(row=2, column=1, sticky="w", pady=5)

        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _build_employment_tab(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        employ_frame = ttk.LabelFrame(scroll_frame, text="  รายละเอียดการจ้างงาน  ", padding=20)
        employ_frame.pack(fill="x", pady=(0, 15))

        row = 0
        ttk.Label(employ_frame, text="ประเภทการจ้าง:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.emp_type = ttk.Combobox(employ_frame, width=35, font=("Segoe UI", 10), values=[
            "พนักงานประจำ", "พนักงานสัญญาจ้างรายปี",
            "สัญญาจ้างเหมารายเดือน", "สัญญาจ้างเหมารายวัน", "สัญญาจ้างเหมารายปี",
            "ที่ปรึกษา"  
        ], state="readonly")
        self.emp_type.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(employ_frame, text="วันที่เริ่มงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        start_frame = ttk.Frame(employ_frame)
        start_frame.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        
        # --- !! (แก้ไข: เปลี่ยนเป็น DateDropdown) !! ---
        self.start_entry = DateDropdown(start_frame, font=("Segoe UI", 10))
        self.start_entry.pack(side="left")
        ttk.Button(start_frame, text="คำนวณอายุงาน", command=self.calc_exp, width=15).pack(side="left", padx=(10, 5))
        self.exp_label = ttk.Label(start_frame, text="-", font=("Segoe UI", 10, "bold"), foreground="#3498db")
        self.exp_label.pack(side="left", padx=10)
        # --- (จบส่วนแก้ไข) ---
        
        row += 1
        ttk.Label(employ_frame, text="ตำแหน่ง:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.position_entry = ttk.Entry(employ_frame, width=35, font=("Segoe UI", 10))
        self.position_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)
        row += 1
        ttk.Label(employ_frame, text="ฝ่ายงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.department_entry = ttk.Entry(employ_frame, width=35, font=("Segoe UI", 10))
        self.department_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)

        row += 1
        ttk.Label(employ_frame, text="สถานที่ทำงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.work_location_combo = ttk.Combobox(employ_frame, width=35, font=("Segoe UI", 10),
                                                values=["สำนักงานใหญ่", "คลังสินค้า"], 
                                                state="readonly")
        self.work_location_combo.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)

        row += 1
        ttk.Label(employ_frame, text="สถานะพนักงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.status_combo = ttk.Combobox(employ_frame, width=35, font=("Segoe UI", 10), values=[
            "เป็นพนักงานอยู่", "ผ่านช่วงทดลองงาน", "ระหว่างทดลองงาน", "พ้นสภาพพนักงาน" , "นักศึกษาฝึกงาน"
        ], state="readonly")
        self.status_combo.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        self.status_combo.bind("<<ComboboxSelected>>", self._toggle_termination_fields) 
        
        self.termination_fields_row = row + 1 
        self.term_date_label = ttk.Label(employ_frame, text="วันที่พ้นสภาพ:", font=("Segoe UI", 10))
        
        # --- !! (แก้ไข: เปลี่ยนเป็น DateDropdown) !! ---
        self.term_date_entry = DateDropdown(employ_frame, font=("Segoe UI", 10))
        # --- (จบส่วนแก้ไข) ---
        
        self.term_reason_label = ttk.Label(employ_frame, text="เหตุผล:", font=("Segoe UI", 10))
        self.term_reason_entry = ttk.Entry(employ_frame, width=35, font=("Segoe UI", 10))
        self._toggle_termination_fields()
        
        self.guarantee_frame = ttk.LabelFrame(scroll_frame, text="  🔒 ข้อมูลการค้ำประกัน  ", padding=20)
        self.guarantee_frame.pack(fill="x", pady=(0, 15))
        
        self.guarantee_var = tk.BooleanVar()
        self.guarantee_check = ttk.Checkbutton(self.guarantee_frame, text="มีการค้ำประกัน", 
                                               variable=self.guarantee_var, 
                                               command=self._toggle_guarantee_fields)
        self.guarantee_check.grid(row=0, column=0, columnspan=4, sticky="w", pady=5)
        self.guarantor_label = ttk.Label(self.guarantee_frame, text="ผู้ค้ำประกัน:", font=("Segoe UI", 10))
        self.guarantor_entry = ttk.Entry(self.guarantee_frame, width=35, font=("Segoe UI", 10))
        self.guarantee_amount_label = ttk.Label(self.guarantee_frame, text="วงเงินค้ำประกัน:", font=("Segoe UI", 10))
        self.guarantee_amount_entry = ttk.Entry(self.guarantee_frame, width=25, font=("Segoe UI", 10))
        self.guarantee_doc_label = ttk.Label(self.guarantee_frame, text="เอกสารค้ำประกัน:", font=("Segoe UI", 10))
        self.guarantee_file_display = ttk.Label(self.guarantee_frame, text="ยังไม่มีไฟล์อัปโหลด", font=("Segoe UI", 9, "italic"), foreground="gray", width=40, anchor="w")
        self.current_guarantee_doc_id = None
        self.current_guarantee_file_path = None
        self.btn_file_frame = ttk.Frame(self.guarantee_frame) 
        self.guarantee_upload_btn = ttk.Button(self.btn_file_frame, text="อัปโหลด", command=self._upload_guarantee_doc, width=10)
        self.guarantee_view_btn = ttk.Button(self.btn_file_frame, text="ดูไฟล์", command=self._view_guarantee_doc, width=10)
        self.guarantee_delete_btn = ttk.Button(self.btn_file_frame, text="ลบ", command=self._delete_guarantee_doc, width=5)
        self.guarantee_upload_btn.pack(side="left", padx=(0, 5))
        self.guarantee_view_btn.pack(side="left", padx=5)
        self.guarantee_delete_btn.pack(side="left", padx=5)
        self._toggle_guarantee_fields()

        self.probation_frame = ttk.LabelFrame(scroll_frame, text="  ⏱️ ข้อมูลการทดลองงาน  ", padding=20)
        self.probation_frame.pack(fill="x", pady=(0, 15))

        row = 0
        ttk.Label(self.probation_frame, text="ระยะเวลาทดลองงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.probation_days_combo = ttk.Combobox(self.probation_frame, width=15, font=("Segoe UI", 10),
                                                 values=["90", "120", "30", "60", "180"], state="readonly")
        self.probation_days_combo.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(self.probation_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="w", padx=5)
        row += 1
        ttk.Label(self.probation_frame, text="วันที่สิ้นสุดทดลองงาน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.probation_end_date_label = ttk.Label(self.probation_frame, text="-", font=("Segoe UI", 10, "bold"), foreground="#e67e22")
        self.probation_end_date_label.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)
        row += 1
        ttk.Label(self.probation_frame, text="ประเมินผล (รอบ 1):", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        score_frame = ttk.Frame(self.probation_frame)
        score_frame.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)
        self.probation_score_entry = ttk.Entry(score_frame, width=15, font=("Segoe UI", 10))
        self.probation_score_entry.pack(side="left")
        self.probation_status_label = ttk.Label(score_frame, text="", font=("Segoe UI", 10, "bold"))
        self.probation_status_label.pack(side="left", padx=10)
        self.probation_score_entry.bind("<KeyRelease>", self._update_probation_status)
        row += 1
        self.probation_score_2_row = row 
        self.probation_score_2_label = ttk.Label(self.probation_frame, text="ประเมินผล (รอบ 2):", font=("Segoe UI", 10))
        self.score_frame_2 = ttk.Frame(self.probation_frame)
        self.probation_score_2_entry = ttk.Entry(self.score_frame_2, width=15, font=("Segoe UI", 10)) 
        self.probation_score_2_entry.pack(side="left")
        self.probation_status_2_label = ttk.Label(self.score_frame_2, text="", font=("Segoe UI", 10, "bold")) 
        self.probation_status_2_label.pack(side="left", padx=10)
        self.probation_score_2_entry.bind("<KeyRelease>", self._update_probation_status)
        
        # --- !! (แก้ไข: ผูก Event กับ DateDropdown) !! ---
        self.start_entry.day_var.trace_add("write", self._calculate_probation_end_date)
        self.start_entry.month_var.trace_add("write", self._calculate_probation_end_date)
        self.start_entry.year_var.trace_add("write", self._calculate_probation_end_date)
        # --- (จบส่วนแก้ไข) ---
        
        self.probation_days_combo.bind("<<ComboboxSelected>>", 
            lambda e: self._chain_events(e, self._calculate_probation_end_date, self._toggle_probation_scores))
        
        training_frame = ttk.LabelFrame(scroll_frame, text="  📚 ประวัติการฝึกอบรม  ", padding=20)
        training_frame.pack(fill="x", pady=(0, 15))

        # ส่วนกรอกข้อมูล
        tf_input = ttk.Frame(training_frame)
        tf_input.pack(fill="x", pady=(0, 10))
        
        ttk.Label(tf_input, text="วันที่อบรม:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
        self.train_date_input = DateDropdown(tf_input, font=("Segoe UI", 10))
        self.train_date_input.pack(side="left", padx=(0, 10))
        
        ttk.Label(tf_input, text="ชื่อหลักสูตร:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
        self.train_name_input = ttk.Entry(tf_input, width=30, font=("Segoe UI", 10))
        self.train_name_input.pack(side="left", padx=(0, 10))
        
        ttk.Label(tf_input, text="มูลค่า (บาท):", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
        self.train_cost_input = ttk.Entry(tf_input, width=10, font=("Segoe UI", 10))
        self.train_cost_input.pack(side="left", padx=(0, 10))
        
        ttk.Button(tf_input, text="➕ เพิ่มรายการ", command=self._add_training_record, style="Success.TButton").pack(side="left")

        # ส่วนตารางแสดงผล
        tf_tree_frame = ttk.Frame(training_frame)
        tf_tree_frame.pack(fill="x")
        
        self.training_tree = ttk.Treeview(tf_tree_frame, columns=("date", "name", "cost"), show="headings", height=5)
        self.training_tree.heading("date", text="วันที่")
        self.training_tree.heading("name", text="ชื่อหลักสูตร")
        self.training_tree.heading("cost", text="มูลค่า (บาท)")
        
        self.training_tree.column("date", width=100, anchor="center")
        self.training_tree.column("name", width=300, anchor="w")
        self.training_tree.column("cost", width=100, anchor="e")
        
        self.training_tree.pack(side="left", fill="x", expand=True)
        
        tf_scroll = ttk.Scrollbar(tf_tree_frame, orient="vertical", command=self.training_tree.yview)
        tf_scroll.pack(side="right", fill="y")
        self.training_tree.configure(yscrollcommand=tf_scroll.set)
        
        ttk.Button(training_frame, text="🗑️ ลบรายการที่เลือก", command=self._delete_training_record).pack(anchor="w", pady=5)
        
        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _build_assets_tab(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # --- กรอบอุปกรณ์ IT ---
        it_frame = ttk.LabelFrame(scroll_frame, text="  💻 อุปกรณ์ IT และการสื่อสาร  ", padding=20)
        it_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(it_frame, text="คอมพิวเตอร์ (รุ่น/Serial):", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.asset_computer = ttk.Entry(it_frame, width=35, font=("Segoe UI", 10))
        self.asset_computer.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(it_frame, text="โทรศัพท์ (รุ่น/IMEI):", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.asset_phone = ttk.Entry(it_frame, width=35, font=("Segoe UI", 10))
        self.asset_phone.grid(row=row, column=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(it_frame, text="เบอร์โทรศัพท์(บริษัท):", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.asset_number = ttk.Entry(it_frame, width=25, font=("Segoe UI", 10))
        self.asset_number.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(it_frame, text="ค่ายมือถือ:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.asset_carrier = ttk.Combobox(it_frame, width=20, values=["AIS", "True", "DTAC"], font=("Segoe UI", 10))
        self.asset_carrier.grid(row=row, column=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(it_frame, text="ประเภทซิม:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.asset_sim = ttk.Combobox(it_frame, width=20, values=["รายเดือน", "เติมเงิน","รายเดือนกลุ่ม"], font=("Segoe UI", 10))
        self.asset_sim.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(it_frame, text="Email บริษัท:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.asset_email = ttk.Entry(it_frame, width=35, font=("Segoe UI", 10))
        self.asset_email.grid(row=row, column=3, sticky="w", pady=10)
        
        # --- กรอบบัญชีโซเชียลและอื่นๆ ---
        social_frame = ttk.LabelFrame(scroll_frame, text="  🌐 บัญชีโซเชียลและอื่นๆ  ", padding=20)
        social_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(social_frame, text="Line ID:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.asset_line = ttk.Entry(social_frame, width=25, font=("Segoe UI", 10))
        self.asset_line.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(social_frame, text="เบอร์ที่ผูกกับ Line:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.asset_line_phone = ttk.Entry(social_frame, width=25, font=("Segoe UI", 10))
        self.asset_line_phone.grid(row=row, column=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(social_frame, text="Facebook Link/Name:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.asset_facebook = ttk.Entry(social_frame, width=35, font=("Segoe UI", 10))
        self.asset_facebook.grid(row=row, column=1, sticky="w", pady=10)
        
        ttk.Label(social_frame, text="รหัสบัตรพนักงาน (Access Card):", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.asset_card_id = ttk.Entry(social_frame, width=25, font=("Segoe UI", 10))
        self.asset_card_id.grid(row=row, column=3, sticky="w", pady=10)

        row += 1
        ttk.Label(social_frame, text="อื่นๆ / หมายเหตุ:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.asset_others = tk.Text(social_frame, width=50, height=3, font=("Segoe UI", 10))
        self.asset_others.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _build_salary_tab(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        salary_frame = ttk.LabelFrame(scroll_frame, text="  💰 เงินเดือนปัจจุบัน  ", padding=20)
        salary_frame.pack(fill="x", pady=(0, 15))
        ttk.Label(salary_frame, text="เงินเดือน:", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="e", padx=(0, 10), pady=10)
        self.salary_entry = ttk.Entry(salary_frame, width=25, font=("Segoe UI", 11))
        self.salary_entry.grid(row=0, column=1, sticky="w", pady=10)
        ttk.Label(salary_frame, text="บาท", font=("Segoe UI", 11)).grid(row=0, column=2, sticky="w", padx=5, pady=10)
        sales_frame = ttk.LabelFrame(scroll_frame, text="  📈 ข้อมูลฝ่ายขาย (Sales Info)  ", padding=20)
        sales_frame.pack(fill="x", pady=(0, 15))

        self.is_sales_var = tk.BooleanVar()
        self.sale_type_var = tk.StringVar()
        self.commission_plan_var = tk.StringVar()

        # Checkbox: เป็นพนักงานขายหรือไม่
        cb_is_sales = ttk.Checkbutton(sales_frame, text="เป็นพนักงานฝ่ายขาย (Sales)", 
                                      variable=self.is_sales_var, command=self._toggle_sales_options)
        cb_is_sales.grid(row=0, column=0, columnspan=3, sticky="w", pady=5)

        # Frame ย่อยสำหรับตัวเลือก (ซ่อน/แสดง)
        self.sales_options_frame = ttk.Frame(sales_frame)
        self.sales_options_frame.grid(row=1, column=0, columnspan=3, sticky="w", padx=20)

        # Sale Type (Radio)
        ttk.Label(self.sales_options_frame, text="ประเภท Sale:").grid(row=0, column=0, sticky="w", pady=5)
        rb_in = ttk.Radiobutton(self.sales_options_frame, text="Inbound", variable=self.sale_type_var, value="Inbound")
        rb_in.grid(row=0, column=1, padx=10)
        rb_out = ttk.Radiobutton(self.sales_options_frame, text="Outbound", variable=self.sale_type_var, value="Outbound")
        rb_out.grid(row=0, column=2, padx=10)

        # Commission Plan (Dropdown)
        ttk.Label(self.sales_options_frame, text="Commission Plan:").grid(row=1, column=0, sticky="w", pady=5)
        plan_options = ["Plan A", "Plan B", "Plan C", "Plan D", "Plan F"]
        self.plan_combo = ttk.Combobox(self.sales_options_frame, values=plan_options, textvariable=self.commission_plan_var, state="readonly", width=15)
        self.plan_combo.grid(row=1, column=1, columnspan=2, sticky="w", padx=10)

        self._toggle_sales_options()
        welfare_frame = ttk.LabelFrame(scroll_frame, text="  🎁 สวัสดิการและค่าใช้จ่าย  ", padding=20)
        welfare_frame.pack(fill="x", pady=(0, 15))
        self.welfare_vars = []
        self.welfare_amount_entries = []
        for i, option in enumerate(self.welfare_options):
            var = tk.BooleanVar()
            self.welfare_vars.append(var)
            cb = ttk.Checkbutton(welfare_frame, text=option, variable=var)
            cb.grid(row=i, column=0, sticky="w", padx=5, pady=5)
            amount_entry = ttk.Entry(welfare_frame, width=15, font=("Segoe UI", 10))
            amount_entry.grid(row=i, column=1, padx=10, pady=5)
            ttk.Label(welfare_frame, text="บาท").grid(row=i, column=2, sticky="w", padx=5, pady=5)
            self.welfare_amount_entries.append(amount_entry)
        history_frame = ttk.LabelFrame(scroll_frame, text="  📊 ประวัติการปรับเงินเดือนรายปี  ", padding=20)
        history_frame.pack(fill="x", pady=(0, 15))
        self.salary_history = []
        years = [2569, 2570, 2571, 2572, 2573]
        ttk.Label(history_frame, text="ครั้งที่", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=10, pady=5)
        ttk.Label(history_frame, text="ปี", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(history_frame, text="เงินเดือนใหม่", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, padx=10, pady=5)
        ttk.Label(history_frame, text="ค่าตำแหน่ง", font=("Segoe UI", 10, "bold")).grid(row=0, column=3, padx=10, pady=5)
        ttk.Label(history_frame, text="ตำแหน่งใหม่", font=("Segoe UI", 10, "bold")).grid(row=0, column=4, padx=10, pady=5) 
        ttk.Label(history_frame, text="ผลประเมิน (Manual)", font=("Segoe UI", 10, "bold")).grid(row=0, column=5, padx=10, pady=5) 
        for i in range(5):
            ttk.Label(history_frame, text=f"{i+1}", font=("Segoe UI", 10)).grid(row=i+1, column=0, padx=10, pady=5)
            year_combo = ttk.Combobox(history_frame, values=years, width=12, state="readonly", font=("Segoe UI", 10))
            year_combo.grid(row=i+1, column=1, padx=10, pady=5)
            salary_entry = ttk.Entry(history_frame, width=20, font=("Segoe UI", 10))
            salary_entry.grid(row=i+1, column=2, padx=10, pady=5)
            allowance_entry = ttk.Entry(history_frame, width=20, font=("Segoe UI", 10))
            allowance_entry.grid(row=i+1, column=3, padx=10, pady=5)
            new_pos_entry = ttk.Entry(history_frame, width=20, font=("Segoe UI", 10))
            new_pos_entry.grid(row=i+1, column=4, padx=10, pady=5)
            score_frame = ttk.Frame(history_frame)
            score_frame.grid(row=i+1, column=5, padx=10, pady=5, sticky="w") 
            score_entry = ttk.Entry(score_frame, width=15, font=("Segoe UI", 10),
                                      validate="key", 
                                      validatecommand=self.score_validator)
            score_entry.pack(side="left")
            ttk.Label(score_frame, text="/ 1000", font=("Segoe UI", 10, "italic"), foreground="gray").pack(side="left", padx=(5,0))
            self.salary_history.append({
                "year": year_combo, 
                "salary": salary_entry, 
                "position_allowance": allowance_entry,
                "new_position": new_pos_entry,      
                "assessment_score": score_entry   
            })
            canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
            canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _validate_score(self, new_value):
        if new_value == "":
            return True
        try:
            value = float(new_value)
            if 0 <= value <= 1000:
                return True
            else:
                return False
        except ValueError:
            return False

    def _build_bank_tab(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        health_frame = ttk.LabelFrame(scroll_frame, text="  ⚕️ ข้อมูลสุขภาพ  ", padding=20)
        health_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(health_frame, text="ปัญหาด้านสุขภาพ:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="e", padx=(0, 10), pady=10)
        self.health_combo = ttk.Combobox(health_frame, width=20, font=("Segoe UI", 10), 
                                        values=["ไม่มี", "มี"], state="readonly")
        self.health_combo.grid(row=0, column=1, sticky="w", pady=10)
        self.health_combo.bind("<<ComboboxSelected>>", self.toggle_health_detail)
        self.health_detail_label = ttk.Label(health_frame, text="รายละเอียด:", font=("Segoe UI", 10))
        self.health_detail_entry = ttk.Entry(health_frame, width=50, font=("Segoe UI", 10))
        
        bank_frame = ttk.LabelFrame(scroll_frame, text="  🏦 ข้อมูลบัญชีเงินเดือน  ", padding=20)
        bank_frame.pack(fill="x", pady=(0, 15))
        row = 0
        ttk.Label(bank_frame, text="หมายเลขบัญชี:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.account_entry = ttk.Entry(bank_frame, width=30, font=("Segoe UI", 10))
        self.account_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)
        row += 1
        ttk.Label(bank_frame, text="ธนาคาร:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.bank_entry = ttk.Entry(bank_frame, width=30, font=("Segoe UI", 10))
        self.bank_entry.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(bank_frame, text="สาขา:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.branch_entry = ttk.Entry(bank_frame, width=25, font=("Segoe UI", 10))
        self.branch_entry.grid(row=row, column=3, sticky="w", pady=10)
        row += 1
        ttk.Label(bank_frame, text="ชื่อบัญชี:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.account_name_entry = ttk.Entry(bank_frame, width=35, font=("Segoe UI", 10))
        self.account_name_entry.grid(row=row, column=1, columnspan=2, sticky="w", pady=10)
        row += 1
        ttk.Label(bank_frame, text="ประเภทบัญชี:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.account_type_var = tk.StringVar()
        type_frame = ttk.Frame(bank_frame)
        type_frame.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        ttk.Radiobutton(type_frame, text="ออมทรัพย์", variable=self.account_type_var, 
                       value="ออมทรัพย์").pack(side="left", padx=(0, 20))
        ttk.Radiobutton(type_frame, text="กระแสรายวัน", variable=self.account_type_var, 
                       value="กระแสรายวัน").pack(side="left")
        
        leave_frame = ttk.LabelFrame(scroll_frame, text="  🌴 จำนวนสิทธิ์การลา (ต่อปี)  ", padding=20)
        leave_frame.pack(fill="x", pady=(0, 15))
        row = 0
        ttk.Label(leave_frame, text="ลาพักร้อนประจำปี:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.leave_annual_entry = ttk.Entry(leave_frame, width=10, font=("Segoe UI", 10))
        self.leave_annual_entry.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(leave_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="w", padx=5, pady=10)
        row += 1
        ttk.Label(leave_frame, text="ลาป่วย:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.leave_sick_entry = ttk.Entry(leave_frame, width=10, font=("Segoe UI", 10))
        self.leave_sick_entry.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(leave_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="w", padx=5, pady=10)
        row += 1
        ttk.Label(leave_frame, text="ลากิจ:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.leave_personal_entry = ttk.Entry(leave_frame, width=10, font=("Segoe UI", 10))
        self.leave_personal_entry.grid(row=row, column=1, sticky="w", pady=10)
        ttk.Label(leave_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="w", padx=5, pady=10)
        row = 0
        ttk.Label(leave_frame, text="ลาบวช:", font=("Segoe UI", 10)).grid(row=row, column=3, sticky="e", padx=(20, 10), pady=10)
        self.leave_ordination_entry = ttk.Entry(leave_frame, width=10, font=("Segoe UI", 10))
        self.leave_ordination_entry.grid(row=row, column=4, sticky="w", pady=10)
        ttk.Label(leave_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=5, sticky="w", padx=5, pady=10)
        row += 1
        ttk.Label(leave_frame, text="ลาคลอด:", font=("Segoe UI", 10)).grid(row=row, column=3, sticky="e", padx=(20, 10), pady=10)
        self.leave_maternity_entry = ttk.Entry(leave_frame, width=10, font=("Segoe UI", 10))
        self.leave_maternity_entry.grid(row=row, column=4, sticky="w", pady=10)
        ttk.Label(leave_frame, text="วัน", font=("Segoe UI", 10)).grid(row=row, column=5, sticky="w", padx=5, pady=10)

        sso_frame = ttk.LabelFrame(scroll_frame, text="  📋 ข้อมูลประกันสังคม  ", padding=20)
        sso_frame.pack(fill="x", pady=(0, 15))

        row = 0
        ttk.Label(sso_frame, text="วันแจ้งเข้าผู้ประกันตน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        
        # --- !! (แก้ไข: เปลี่ยนเป็น DateDropdown) !! ---
        self.sso_start_entry = DateDropdown(sso_frame, font=("Segoe UI", 10))
        self.sso_start_entry.grid(row=row, column=1, sticky="w", pady=10)     

        ttk.Label(sso_frame, text="วันที่ทำการ:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        
        self.sso_start_action_entry = DateDropdown(sso_frame, font=("Segoe UI", 10))
        self.sso_start_action_entry.grid(row=row, column=3, sticky="w", pady=10)      

        row += 1
        ttk.Label(sso_frame, text="วันสิ้นสุดผู้ประกันตน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        
        self.sso_end_entry = DateDropdown(sso_frame, font=("Segoe UI", 10))
        self.sso_end_entry.grid(row=row, column=1, sticky="w", pady=10)      

        ttk.Label(sso_frame, text="วันที่ทำการ:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        
        self.sso_end_action_entry = DateDropdown(sso_frame, font=("Segoe UI", 10))
        self.sso_end_action_entry.grid(row=row, column=3, sticky="w", pady=10)
        # --- (จบส่วนแก้ไข) ---
        
        row += 1
        ttk.Label(sso_frame, text="โรงพยาบาล (ประกันสังคม):", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.sso_hospital_entry = ttk.Entry(sso_frame, width=35, font=("Segoe UI", 10))
        self.sso_hospital_entry.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        
        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    # --- !! (เริ่ม "ผ่าตัด" แท็บ 5) !! ---
    def _build_attendance_tab(self, parent):
        """(ส่วนที่ 3) Tab ใหม่สำหรับบันทึก ลา/สาย/ใบเตือน"""
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # --- (!! "ผ่าตัด" กรอบบันทึกการลา) !! ---
        leave_frame = ttk.LabelFrame(scroll_frame, text="  📝 บันทึกการลา (ป่วย/กิจ/พักร้อน)  ", padding=20)
        leave_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(leave_frame, text="วันที่ลา:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        self.att_leave_date = DateDropdown(leave_frame, font=("Segoe UI", 10))
        self.att_leave_date.grid(row=row, column=1, sticky="w", pady=10)

        ttk.Label(leave_frame, text="ประเภท:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        self.att_leave_type = ttk.Combobox(leave_frame, width=20, font=("Segoe UI", 10),
                                           values=["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาอื่นๆ"], state="readonly")
        self.att_leave_type.grid(row=row, column=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(leave_frame, text="ระยะเวลา:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        
        self.att_leave_duration_type = ttk.Combobox(leave_frame, width=18, font=("Segoe UI", 10),
                                           values=["เต็มวัน (1.0)", "ครึ่งวัน (0.5)", "ระบุเวลา (ชม.)"], 
                                           state="readonly")
        self.att_leave_duration_type.grid(row=row, column=1, sticky="w", pady=10)
        self.att_leave_duration_type.set("เต็มวัน (1.0)")
        self.att_leave_duration_type.bind("<<ComboboxSelected>>", self._toggle_leave_time_entries)

        # (ซ่อนไว้ก่อน)
        self.att_leave_time_frame = ttk.Frame(leave_frame)
        self.att_leave_time_frame.grid(row=row, column=2, columnspan=2, sticky="w", pady=0, padx=10)
        
        ttk.Label(self.att_leave_time_frame, text="ตั้งแต่:").pack(side="left")
        self.att_leave_start_time = ttk.Combobox(self.att_leave_time_frame, values=self.time_options, 
                                                 width=6, font=("Segoe UI", 10))
        self.att_leave_start_time.pack(side="left", padx=5)
        
        ttk.Label(self.att_leave_time_frame, text="ถึง:").pack(side="left")
        self.att_leave_end_time = ttk.Combobox(self.att_leave_time_frame, values=self.time_options, 
                                               width=6, font=("Segoe UI", 10))
        self.att_leave_end_time.pack(side="left", padx=5)
        
        self.att_leave_time_frame.grid_remove() # (ซ่อนไว้ก่อน)
        
        row += 1
        ttk.Label(leave_frame, text="เหตุผล:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.att_leave_reason = tk.Text(leave_frame, width=50, height=3, font=("Segoe UI", 10))
        self.att_leave_reason.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)

        row += 1
        ttk.Button(leave_frame, text="💾 บันทึกการลา", command=self._save_leave_record, 
                   width=20, style="Primary.TButton").grid(row=row, column=1, columnspan=3, sticky="e", pady=10)
        # --- (จบส่วน "ผ่าตัด" การลา) ---

        # --- กรอบสำหรับ "บันทึกการมาสาย" ---
        late_frame = ttk.LabelFrame(scroll_frame, text="  🏃 บันทึกการมาสาย  ", padding=20)
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
        ttk.Button(late_frame, text="💾 บันทึกการมาสาย", command=self._save_late_record, 
                   width=20, style="Primary.TButton").grid(row=row, column=1, columnspan=3, sticky="e", pady=10)

        # --- กรอบสำหรับ "บันทึกใบเตือน" ---
        warn_frame = ttk.LabelFrame(scroll_frame, text="  📜 บันทึกใบเตือน (ใบมีด)  ", padding=20)
        warn_frame.pack(fill="x", pady=(0, 15))
        
        row = 0
        ttk.Label(warn_frame, text="วันที่เตือน:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=10)
        
        self.att_warn_date = DateDropdown(warn_frame, font=("Segoe UI", 10))
        self.att_warn_date.grid(row=row, column=1, sticky="w", pady=10)

        ttk.Label(warn_frame, text="ระดับ:", font=("Segoe UI", 10)).grid(row=row, column=2, sticky="e", padx=(20, 10), pady=10)
        warning_options_list = [
            "ตักเตือนทางวาจา ครั้ง 1",
            "ตักเตือนทางวาจา ครั้ง 2",
            "ตักเตือนทางวาจา มากกว่า 2 ครั้ง",
            "หนังสือตักเตือนครั้งที่ 1",
            "หนังสือตักเตือนครั้งที่ 2",
            "หนังสือเลิกจ้าง"
        ]
        self.att_warn_level = ttk.Combobox(warn_frame, width=30, font=("Segoe UI", 10),
                                           values=warning_options_list, state="readonly")
        self.att_warn_level.grid(row=row, column=3, sticky="w", pady=10)
        
        row += 1
        ttk.Label(warn_frame, text="เหตุผล/คำอธิบาย:", font=("Segoe UI", 10)).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        self.att_warn_reason = tk.Text(warn_frame, width=50, height=3, font=("Segoe UI", 10))
        self.att_warn_reason.grid(row=row, column=1, columnspan=3, sticky="w", pady=10)
        row += 1
        self.warn_doc_label = ttk.Label(warn_frame, text="ไฟล์แนบ (ใบเตือน):", font=("Segoe UI", 10))
        self.warn_doc_label.grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=10)
        
        self.warn_file_display = ttk.Label(warn_frame, text="ยังไม่มีไฟล์อัปโหลด", font=("Segoe UI", 9, "italic"), foreground="gray", width=40, anchor="w")
        self.warn_file_display.grid(row=row, column=1, columnspan=2, sticky="w", pady=(10,0))
        
        self.warn_btn_frame = ttk.Frame(warn_frame) 
        self.warn_btn_frame.grid(row=row, column=3, sticky="w", pady=5)
        
        self.warn_upload_btn = ttk.Button(self.warn_btn_frame, text="อัปโหลด", command=self._upload_warning_doc, width=10)
        self.warn_upload_btn.pack(side="left", padx=(0, 5))
        
        self.warn_view_btn = ttk.Button(self.warn_btn_frame, text="ดูไฟล์", command=self._view_warning_doc, width=10)
        self.warn_view_btn.pack(side="left", padx=5)
        
        self.warn_delete_btn = ttk.Button(self.warn_btn_frame, text="ลบ", command=self._delete_warning_doc, width=5)
        self.warn_delete_btn.pack(side="left", padx=5)

        row += 1
        ttk.Button(warn_frame, text="💾 บันทึกใบเตือน", command=self._save_warning_record, 
                   width=20, style="Primary.TButton").grid(row=row, column=1, columnspan=3, sticky="e", pady=10)


        report_frame = ttk.LabelFrame(scroll_frame, text="  📊 สรุปประวัติ (Report)  ", padding=20)
        report_frame.pack(fill="x", pady=(0, 15))
        ttk.Label(report_frame, text="(ส่วนสรุป Report จะพัฒนาในขั้นตอนถัดไป)").pack(padx=10, pady=10)


        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

                                
    def _create_form_buttons(self):
        """สร้างปุ่ม Action Bar ด้านล่างฟอร์ม (จัด Layout ใหม่กันตกขอบ)"""
        
        # ใช้ LabelFrame เพื่อแบ่งส่วนชัดเจน
        action_bar = ttk.LabelFrame(self.form_page, text=" บันทึกและจัดการข้อมูล ", padding=10)
        action_bar.pack(fill="x", side="bottom", padx=15, pady=15) # side="bottom" เพื่อให้เกาะล่างเสมอ

        # โซนซ้าย: ปุ่มหลัก (บันทึก / ล้าง)
        left_zone = ttk.Frame(action_bar)
        left_zone.pack(side="left")

        self.btn_save = ttk.Button(left_zone, text="💾 บันทึกข้อมูล", command=self.save_employee, 
                                  width=18, style="Success.TButton")
        self.btn_save.pack(side="left", padx=5)

        self.btn_clear = ttk.Button(left_zone, text="🗑️ ล้างฟอร์ม", command=self.clear_form, 
                                   width=15)
        self.btn_clear.pack(side="left", padx=5)

        # โซนขวา: ปุ่มเสริม (Export PDF)
        right_zone = ttk.Frame(action_bar)
        right_zone.pack(side="right")

        self.btn_pdf = ttk.Button(right_zone, text="📄 Export PDF (รายนี้)", 
                                 command=self._export_to_pdf, width=20)
        self.btn_pdf.pack(side="right", padx=5)

    # === ฟังก์ชัน Helper === 
    def _add_training_record(self):
        date_val = self.train_date_input.get()
        name_val = self.train_name_input.get().strip()
        cost_val = self.train_cost_input.get().strip()
        
        if not name_val:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกชื่อหลักสูตร")
            return
            
        if not cost_val: cost_val = "0"
        try:
            float(cost_val)
        except ValueError:
            messagebox.showwarning("ข้อมูลผิดพลาด", "มูลค่าต้องเป็นตัวเลข")
            return

        self.training_tree.insert("", "0", values=(date_val, name_val, cost_val))
        
        # ล้างค่าหลังเพิ่ม
        self.train_date_input.clear()
        self.train_name_input.delete(0, tk.END)
        self.train_cost_input.delete(0, tk.END)

    def _on_prefix_change(self, event=None):
        if self.combo_prefix.get() == "อื่นๆ":
            self.entry_prefix_other.pack(side="left", padx=5)
        else:
            self.entry_prefix_other.pack_forget()
            self.entry_prefix_other.delete(0, tk.END)

    def _delete_training_record(self):
        selected = self.training_tree.selection()
        if not selected:
            messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือกรายการที่ต้องการลบ")
            return
        for item in selected:
            self.training_tree.delete(item)

    def _on_date_selected(self, event):
        pass

    def _chain_events(self, event, *callbacks):
        for callback in callbacks:
            callback(event)

    def _on_mousewheel(self, event, widget):
        try:
            widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass 

    def _export_to_excel(self):
        employees_to_export = hr_database.load_all_employees() 
        if not employees_to_export:
            messagebox.showwarning("⚠️ ไม่มีข้อมูล", "ไม่มีข้อมูลพนักงานสำหรับ Export")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="บันทึกเป็น Excel"
        )
        
        if not file_path:
            return 

        try:
            processed_data = []
            welfare_options = [
                "ค่าที่พัก", "สิทธิ์ที่พักฟรี", "ค่าอาหาร", "ค่าเดินทาง",
                "ค่าน้ำมัน","ค่าน้ำมัน (Fleet Card)","ค่าที่พักต่างจังหวัด", "ค่าเสื่อมรถยนต์",
                "ค่าโทรศัพท์", "เบี้ยเลี้ยง", "ค่าตำแหน่ง" 
            ]
            
            for emp in employees_to_export:
                flat_emp = emp.copy()
                
                welfare_flags = flat_emp.pop("welfare", [])
                welfare_amounts = flat_emp.pop("welfare_amounts", [])
                for i, option in enumerate(welfare_options):
                    flat_emp[f"สวัสดิการ_{option}_สถานะ"] = "มี" if i < len(welfare_flags) and welfare_flags[i] else "ไม่มี"
                    flat_emp[f"สวัสดิการ_{option}_จำนวนเงิน"] = welfare_amounts[i] if i < len(welfare_amounts) else ""
                
                salary_history = flat_emp.pop("salary_history", [])
                for i, history in enumerate(salary_history):
                    flat_emp[f"ประวัติเงินเดือน_{i+1}_ปี"] = history.get("year", "")
                    flat_emp[f"ประวัติเงินเดือน_{i+1}_เงินเดือน"] = history.get("salary", "")
                    flat_emp[f"ประวัติเงินเดือน_{i+1}_ค่าตำแหน่ง"] = history.get("position_allowance", "")
                    flat_emp["guarantee_enabled"] = "มี" if flat_emp.get("guarantee_enabled") else "ไม่มี"
                
                processed_data.append(flat_emp)
                
            df = pd.DataFrame(processed_data)

            column_mapping = {
                "id": "รหัสพนักงาน", "fname": "ชื่อ", "nickname": "ชื่อเล่น", "lname": "นามสกุล",
                "birth": "วันเกิด", "age": "อายุ", "id_card": "เลขบัตรประชาชน",
                "phone": "เบอร์โทรศัพท์", "address": "ที่อยู่ตามบัตร", "current_address": "ที่อยู่ปัจจุบัน",
                
                # (ข้อมูลใหม่)
                "emergency_name": "ชื่อผู้ติดต่อฉุกเฉิน", "emergency_phone": "เบอร์ฉุกเฉิน", "emergency_relation": "ความสัมพันธ์(ฉุกเฉิน)",
                "ref_name": "ชื่อบุคคลอ้างอิง", "ref_phone": "เบอร์อ้างอิง", "ref_relation": "ความสัมพันธ์(อ้างอิง)",
                
                "emp_type": "ประเภทการจ้าง", "start_date": "วันเริ่มงาน", "exp": "อายุงาน",
                "position": "ตำแหน่ง", "department": "ฝ่าย", "status": "สถานะ",
                "termination_date": "วันพ้นสภาพ", "termination_reason": "เหตุผล",
                "guarantee_enabled": "ค้ำประกัน", "guarantor_name": "ผู้ค้ำ",
                "guarantee_amount": "วงเงินค้ำ", "salary": "เงินเดือน",
                "health": "ปัญหาสุขภาพ", "health_detail": "รายละเอียดสุขภาพ",
                "account": "เลขบัญชี", "bank": "ธนาคาร", "branch": "สาขา",
                "account_name": "ชื่อบัญชี", "account_type": "ประเภทบัญชี",
                "sso_start": "วันแจ้งเข้า", "sso_end": "วันแจ้งออก",
                "sso_start_action": "วันทำรายการเข้า", "sso_end_action": "วันทำรายการออก",
                "leave_annual_days": "พักร้อน(วัน)", "leave_sick_days": "ป่วย(วัน)",
                "leave_ordination_days": "บวช(วัน)", "leave_maternity_days": "คลอด(วัน)" ,
                "leave_personal_days": "กิจ(วัน)"
            }
            
            df = df.rename(columns=column_mapping)
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("✅ สำเร็จ", f"บันทึกไฟล์ Excel เรียบร้อยแล้ว\nที่: {file_path}")

        except Exception as e:
            messagebox.showerror("❌ เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้:\n{e}")

    def _export_to_pdf(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            messagebox.showwarning("⚠️ ไม่มีข้อมูล", "กรุณาเปิดข้อมูลพนักงานที่ต้องการ Export ก่อน")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            title="บันทึกเป็น PDF",
            initialfile=f"ประวัติพนักงาน_{emp_id}.pdf"
        )
        
        if not file_path: return 

        try:
            pdf = FPDF()
            pdf.add_page()
            
            # (ใช้ Path ฟอนต์ที่ถูกต้องตามเครื่องคุณ)
            font_path = os.path.join(os.path.dirname(__file__), "THSarabunNew.ttf")
            if not os.path.exists(font_path):
                # Fallback ไปที่โฟลเดอร์ resources ถ้ามี
                font_path = os.path.join(os.path.dirname(__file__), "resources", "THSarabunNew.ttf")
                
            if not os.path.exists(font_path):
                messagebox.showerror("❌ ไม่พบฟอนต์", "ไม่พบไฟล์ THSarabunNew.ttf")
                return
                
            pdf.add_font("THSarabun", "", font_path, uni=True)
            pdf.set_font("THSarabun", size=16)
            
            pdf.set_font_size(24)
            pdf.cell(0, 10, f"ประวัติพนักงาน: {self.fname_entry.get()} {self.lname_entry.get()}", 
                     ln=True, align='C')
            pdf.ln(10) 
            
            def write_field(label, value):
                pdf.set_font("THSarabun", size=16)
                pdf.cell(50, 10, f"{label}:", border=0)
                pdf.set_font("THSarabun", size=16)
                pdf.cell(0, 10, value, ln=True, border=0)

            # --- 1. ข้อมูลส่วนตัว ---
            pdf.set_font_size(18)
            pdf.cell(0, 10, "ข้อมูลส่วนตัว", ln=True)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
            pdf.ln(2)

            write_field("รหัสพนักงาน", self.emp_id_entry.get())
            write_field("ชื่อ-นามสกุล", f"{self.fname_entry.get()} {self.lname_entry.get()} ( {self.nickname_entry.get()} )")
            write_field("เบอร์โทร", self.phone_entry.get())
            write_field("เลขบัตรประชาชน", self.id_card_entry.get())
            write_field("ที่อยู่ปัจจุบัน", self.current_address_text.get("1.0", "end-1c").replace("\n", " "))

            # --- (!!! เพิ่มใหม่: บุคคลติดต่อฉุกเฉิน !!!) ---
            pdf.ln(5)
            pdf.set_font_size(18)
            pdf.cell(0, 10, "บุคคลติดต่อฉุกเฉิน & อ้างอิง", ln=True)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
            pdf.ln(2)
            
            write_field("ฉุกเฉิน: ชื่อ", self.emergency_name.get())
            write_field("ฉุกเฉิน: เบอร์", self.emergency_phone.get())
            write_field("ฉุกเฉิน: สัมพันธ์", self.emergency_relation.get())
            pdf.ln(2)
            write_field("อ้างอิง: ชื่อ", self.ref_name.get())
            write_field("อ้างอิง: เบอร์", self.ref_phone.get())
            write_field("อ้างอิง: สัมพันธ์", self.ref_relation.get())
            # -------------------------------------------

            # --- 2. ข้อมูลการจ้าง ---
            pdf.ln(5)
            pdf.set_font_size(18)
            pdf.cell(0, 10, "ข้อมูลการจ้าง", ln=True)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
            pdf.ln(2)
            
            write_field("ตำแหน่ง", self.position_entry.get())
            write_field("ฝ่าย", self.department_entry.get())
            write_field("สถานะ", self.status_combo.get())
            if self.status_combo.get() == "พ้นสภาพพนักงาน":
                write_field(" - วันที่พ้นสภาพ", self.term_date_entry.get())
                write_field(" - เหตุผล", self.term_reason_entry.get())
            write_field("วันที่เริ่มงาน", self.start_entry.get())
            write_field("เงินเดือน", f"{self.salary_entry.get()} บาท")
            
            # --- 3. ข้อมูลค้ำประกัน ---
            pdf.ln(5)
            pdf.set_font_size(18)
            pdf.cell(0, 10, "ข้อมูลการค้ำประกัน", ln=True)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
            pdf.ln(2)

            if self.guarantee_var.get():
                write_field("ผู้ค้ำประกัน", self.guarantor_entry.get())
                write_field("วงเงินค้ำประกัน", f"{self.guarantee_amount_entry.get()} บาท")
            else:
                write_field("การค้ำประกัน", "ไม่มี")
            
            pdf.output(file_path)
            messagebox.showinfo("✅ สำเร็จ", f"บันทึกไฟล์ PDF เรียบร้อยแล้ว\nที่: {file_path}")

        except Exception as e:
            messagebox.showerror("❌ เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ PDF ได้:\n{e}")

    def calculate_age(self, *args): # (เพิ่ม *args)
        try:
            birth_str = self.birth_entry.get()
            if not birth_str:
                self.age_label.config(text="ยังไม่เลือกวันเกิด")
                return
                
            day, month, year_be = map(int, birth_str.split('/'))
            year_ce = int(year_be) - 543
            birth_date = datetime(year_ce, int(month), int(day))
            today = datetime.today()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            self.age_label.config(text=f"{age} ปี")
        except Exception:
            self.age_label.config(text="รูปแบบไม่ถูกต้อง")

    def _check_score(self, score_entry, status_label):
        score_text = score_entry.get()
        score_text = score_text.replace("%", "").strip()
        
        try:
            score_value = float(score_text)
            if score_value >= 75:
                status_label.config(text="✔ ผ่าน", foreground="#27ae60") 
            else:
                status_label.config(text="❌ ไม่ผ่าน", foreground="#c0392b") 
        except ValueError:
            status_label.config(text="", foreground="black")
    
    def _update_probation_status(self, event):
        widget = event.widget
        if widget == self.probation_score_entry:
            self._check_score(self.probation_score_entry, self.probation_status_label)
        elif widget == self.probation_score_2_entry:
            self._check_score(self.probation_score_2_entry, self.probation_status_2_label)
    
    def _toggle_probation_scores(self, event=None):
        try:
            days_str = self.probation_days_combo.get()
            
            if days_str in ("120", "180"): 
                self.probation_score_2_label.grid(row=self.probation_score_2_row, column=0, 
                                                  sticky="e", padx=(0, 10), pady=10)
                self.score_frame_2.grid(row=self.probation_score_2_row, column=1, 
                                                  columnspan=2, sticky="w", pady=10)
            else:
                self.probation_score_2_label.grid_forget()
                self.score_frame_2.grid_forget()
                
                self.probation_score_2_entry.delete(0, tk.END)
                self.probation_status_2_label.config(text="") 

        except AttributeError:
            pass

    def _calculate_probation_end_date(self, *args): # (เพิ่ม *args)
        try:
            start_str = self.start_entry.get()
            days_str = self.probation_days_combo.get()

            if not start_str or not days_str:
                self.probation_end_date_label.config(text="-")
                return

            probation_days = int(days_str)
            
            day, month, year_be = map(int, start_str.split('/'))
            year_ce = int(year_be) - 543
            start_date = datetime(year_ce, int(month), int(day))

            end_date = start_date + timedelta(days=probation_days)
            
            end_date_be_str = f"{end_date.day:02d}/{end_date.month:02d}/{end_date.year + 543}"
            
            self.probation_end_date_label.config(text=end_date_be_str)

        except Exception as e:
            print(f"DEBUG: _calculate_probation_end_date Error: {e}")
            self.probation_end_date_label.config(text="ข้อมูลไม่ถูกต้อง")

    def calc_exp(self, *args): # (เพิ่ม *args)
        try:
            start_str = self.start_entry.get()
            if not start_str:
                self.exp_label.config(text="ยังไม่เลือกวันเริ่มงาน")
                return
                
            day, month, year_be = map(int, start_str.split('/'))
            year_ce = int(year_be) - 543
            start_date = datetime(year_ce, int(month), int(day))
            today = datetime.today()

            years = today.year - start_date.year
            months = today.month - start_date.month
            if months < 0:
                years -= 1
                months += 12

            days = today.day - start_date.day
            if days < 0:
                months -= 1
                if months < 0:
                    years -= 1
                    months += 12
                prev_month_date = today.replace(day=1) - timedelta(days=1)
                days_in_prev_month = prev_month_date.day
                days += days_in_prev_month

            self.exp_label.config(text=f"{years} ปี {months} เดือน {days} วัน")
        except Exception:
            self.exp_label.config(text="รูปแบบไม่ถูกต้อง")

    def toggle_health_detail(self, event=None):
        value = self.health_combo.get()
        if value == "มี":
            self.health_detail_label.grid(row=1, column=0, sticky="e", padx=(0, 10), pady=10)
            self.health_detail_entry.grid(row=1, column=1, columnspan=3, sticky="w", pady=10)
        else:
            self.health_detail_label.grid_forget()
            self.health_detail_entry.grid_forget()
            try:
                self.health_detail_entry.delete(0, tk.END)
            except Exception:
                pass
     
    def _toggle_termination_fields(self, event=None):
        try:
            base_row = self.termination_fields_row
        except AttributeError:
            return 

        status = self.status_combo.get()
        if status == "พ้นสภาพพนักงาน":
            self.term_date_label.grid(row=base_row, column=0, sticky="e", padx=(0, 10), pady=10)
            self.term_date_entry.grid(row=base_row, column=1, sticky="w", pady=10)
            
            self.term_reason_label.grid(row=base_row + 1, column=0, sticky="e", padx=(0, 10), pady=10)
            self.term_reason_entry.grid(row=base_row + 1, column=1, columnspan=3, sticky="w", pady=10)
        else:
            self.term_date_label.grid_forget()
            self.term_date_entry.grid_forget()
            self.term_reason_label.grid_forget()
            self.term_reason_entry.grid_forget()
            try:
                self.term_date_entry.delete(0, tk.END)
                self.term_reason_entry.delete(0, tk.END)
            except Exception:
                pass

    def _toggle_guarantee_fields(self):
        if self.guarantee_var.get():
            self.guarantor_label.grid(row=1, column=0, sticky="e", padx=(20, 10), pady=10)
            self.guarantor_entry.grid(row=1, column=1, columnspan=3, sticky="w", pady=10)
            self.guarantee_amount_label.grid(row=2, column=0, sticky="e", padx=(20, 10), pady=10)
            self.guarantee_amount_entry.grid(row=2, column=1, sticky="w", pady=10)
            self.guarantee_doc_label.grid(row=3, column=0, sticky="ne", padx=(20, 10), pady=10)
            self.guarantee_file_display.grid(row=3, column=1, sticky="w", pady=(10,0))
            self.btn_file_frame.grid(row=3, column=2, columnspan=2, sticky="w", pady=5)
            self._load_guarantee_doc_status()
        else:
            self.guarantor_label.grid_forget()
            self.guarantor_entry.grid_forget()
            self.guarantee_amount_label.grid_forget()
            self.guarantee_amount_entry.grid_forget()
            self.guarantee_doc_label.grid_forget()
            self.guarantee_file_display.grid_forget()
            self.btn_file_frame.grid_forget() 
            self.current_guarantee_doc_id = None 
            self.current_guarantee_file_path = None
            try:
                self.guarantor_entry.delete(0, tk.END)
                self.guarantee_amount_entry.delete(0, tk.END)
            except Exception:
                pass

    def _load_guarantee_doc_status(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            self.guarantee_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.guarantee_view_btn.config(state="disabled")
            self.guarantee_delete_btn.config(state="disabled")
            self.guarantee_upload_btn.config(text="อัปโหลด")
            self.current_guarantee_doc_id = None
            self.current_guarantee_file_path = None
            return
        doc_info = hr_database.get_document_by_description(emp_id, "เอกสารค้ำประกัน")
        if doc_info and doc_info.get('file_path'):
            self.current_guarantee_doc_id = doc_info['doc_id']
            self.current_guarantee_file_path = doc_info['file_path']
            filename = os.path.basename(doc_info['file_path']) 
            self.guarantee_file_display.config(text=filename, foreground="#007bff") 
            self.guarantee_view_btn.config(state="normal")
            self.guarantee_delete_btn.config(state="normal")
            self.guarantee_upload_btn.config(text="แทนที่") 
        else:
            self.guarantee_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.guarantee_view_btn.config(state="disabled")
            self.guarantee_delete_btn.config(state="disabled")
            self.guarantee_upload_btn.config(text="อัปโหลด")
            self.current_guarantee_doc_id = None
            self.current_guarantee_file_path = None

    def _upload_guarantee_doc(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณากรอกรหัสพนักงานก่อนอัปโหลดไฟล์")
            return
        if not os.path.exists(NETWORK_UPLOAD_PATH):
            messagebox.showerror("Network Error",
                                 f"ไม่พบโฟลเดอร์สำหรับอัปโหลดที่ Server:\n{NETWORK_UPLOAD_PATH}\n"
                                 f"กรุณาตรวจสอบการเชื่อมต่อเครือข่าย หรือติดต่อผู้ดูแลระบบ")
            return
        source_path = filedialog.askopenfilename(
            title="เลือกไฟล์ (PDF, JPG, PNG)",
            filetypes=[("เอกสาร", "*.pdf *.jpg *.jpeg *.png"), ("All Files", "*.*")]
        )
        if not source_path:
            return 
        original_filename, file_extension = os.path.splitext(os.path.basename(source_path))
        timestamp = int(datetime.now().timestamp())
        unique_filename = f"{emp_id}_guarantee_{timestamp}{file_extension}"
        destination_path = os.path.join(NETWORK_UPLOAD_PATH, unique_filename)
        try:
            shutil.copy(source_path, destination_path)
            if self.current_guarantee_doc_id:
                old_path = self.current_guarantee_file_path
                hr_database.delete_document(self.current_guarantee_doc_id)
                if old_path and os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                    except Exception as e:
                        print(f"Warning: ไม่สามารถลบไฟล์เก่า {old_path} ได้: {e}")
            success = hr_database.add_employee_document(emp_id, "เอกสารค้ำประกัน", destination_path)
            if success:
                messagebox.showinfo("สำเร็จ", f"อัปโหลดไฟล์ {unique_filename} เรียบร้อย")
                self._load_guarantee_doc_status() 
        except Exception as e:
            messagebox.showerror("Upload Failed", f"ไม่สามารถอัปโหลดไฟล์ได้:\n{e}")

    def _view_guarantee_doc(self):
        if not self.current_guarantee_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่ระบุ")
            return
        try:
            if not os.path.exists(self.current_guarantee_file_path):
                 messagebox.showerror("เปิดไฟล์ล้มเหลว", 
                                    f"ไม่พบไฟล์ที่: {self.current_guarantee_file_path}\n"
                                    f"ไฟล์อาจถูกย้ายหรือลบไปแล้ว")
                 return
            os.startfile(self.current_guarantee_file_path)
        except Exception as e:
            messagebox.showerror("เปิดไฟล์ล้มเหลว", f"ไม่สามารถเปิดไฟล์ได้:\n{e}")

    def _delete_guarantee_doc(self):
        if not self.current_guarantee_doc_id or not self.current_guarantee_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่จะลบ")
            return
        filename = os.path.basename(self.current_guarantee_file_path)
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบไฟล์ '{filename}' ใช่หรือไม่?"):
            return
        try:
            success_db = hr_database.delete_document(self.current_guarantee_doc_id)
            if success_db:
                if os.path.exists(self.current_guarantee_file_path):
                    try:
                        os.remove(self.current_guarantee_file_path)
                    except Exception as e:
                        messagebox.showwarning("ลบไฟล์สำเร็จ (แต่)",
                                               f"ลบข้อมูลไฟล์ออกจากระบบแล้ว\n"
                                               f"แต่ไม่สามารถลบไฟล์จริงออกจาก Server ได้: {e}")
                else:
                     messagebox.showinfo("สำเร็จ", "ลบข้อมูลไฟล์เรียบร้อย")
                self._load_guarantee_doc_status() 
            else:
                messagebox.showerror("ล้มเหลว", "ไม่สามารถลบข้อมูลออกจากฐานข้อมูลได้")
        except Exception as e:
            messagebox.showerror("ลบไฟล์ล้มเหลว", f"เกิดข้อผิดพลาด: {e}")
    
    def _load_warning_doc_status(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            self.warn_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.warn_view_btn.config(state="disabled")
            self.warn_delete_btn.config(state="disabled")
            self.warn_upload_btn.config(text="อัปโหลด")
            self.current_warning_doc_id = None
            self.current_warning_file_path = None
            return
        doc_info = hr_database.get_document_by_description(emp_id, "เอกสารใบเตือน")
        if doc_info and doc_info.get('file_path'):
            self.current_warning_doc_id = doc_info['doc_id']
            self.current_warning_file_path = doc_info['file_path']
            filename = os.path.basename(doc_info['file_path'])
            self.warn_file_display.config(text=filename, foreground="#007bff")
            self.warn_view_btn.config(state="normal")
            self.warn_delete_btn.config(state="normal")
            self.warn_upload_btn.config(text="แทนที่")
        else:
            self.warn_file_display.config(text="ยังไม่มีไฟล์อัปโหลด", foreground="gray")
            self.warn_view_btn.config(state="disabled")
            self.warn_delete_btn.config(state="disabled")
            self.warn_upload_btn.config(text="อัปโหลด")
            self.current_warning_doc_id = None
            self.current_warning_file_path = None

    def _upload_warning_doc(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณากรอกรหัสพนักงานก่อนอัปโหลดไฟล์")
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
            if self.current_warning_doc_id:
                old_path = self.current_warning_file_path
                hr_database.delete_document(self.current_warning_doc_id)
                if old_path and os.path.exists(old_path):
                    try: os.remove(old_path)
                    except Exception as e: print(f"Warning: ไม่สามารถลบไฟล์เก่า {old_path} ได้: {e}")
            success = hr_database.add_employee_document(emp_id, "เอกสารใบเตือน", destination_path)
            if success:
                messagebox.showinfo("สำเร็จ", f"อัปโหลดไฟล์ {unique_filename} เรียบร้อย")
                self._load_warning_doc_status()
        except Exception as e:
            messagebox.showerror("Upload Failed", f"ไม่สามารถอัปโหลดไฟล์ได้:\n{e}")

    def _view_warning_doc(self):
        if not self.current_warning_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่ระบุ")
            return
        try:
            if not os.path.exists(self.current_warning_file_path):
                 messagebox.showerror("เปิดไฟล์ล้มเหลว", f"ไม่พบไฟล์ที่: {self.current_warning_file_path}")
                 return
            os.startfile(self.current_warning_file_path)
        except Exception as e:
            messagebox.showerror("เปิดไฟล์ล้มเหลว", f"ไม่สามารถเปิดไฟล์ได้:\n{e}")

    def _delete_warning_doc(self):
        if not self.current_warning_doc_id or not self.current_warning_file_path:
            messagebox.showwarning("ไม่พบไฟล์", "ไม่พบข้อมูลไฟล์ที่จะลบ")
            return
        filename = os.path.basename(self.current_warning_file_path)
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบไฟล์ '{filename}' ใช่หรือไม่?"):
            return
        try:
            success_db = hr_database.delete_document(self.current_warning_doc_id)
            if success_db:
                if os.path.exists(self.current_warning_file_path):
                    try: os.remove(self.current_warning_file_path)
                    except Exception as e: messagebox.showwarning("ลบไฟล์สำเร็จ (แต่)", f"ลบข้อมูลไฟล์ออกจากระบบแล้ว\nแต่ไม่สามารถลบไฟล์จริงออกจาก Server ได้: {e}")
                else: messagebox.showinfo("สำเร็จ", "ลบข้อมูลไฟล์เรียบร้อย")
                self._load_warning_doc_status()
            else:
                messagebox.showerror("ล้มเหลว", "ไม่สามารถลบข้อมูลออกจากฐานข้อมูลได้")
        except Exception as e:
            messagebox.showerror("ลบไฟล์ล้มเหลว", f"เกิดข้อผิดพลาด: {e}")

    def save_employee(self):
        if not self.emp_id_entry.get() or not self.fname_entry.get():
            messagebox.showwarning("⚠️ คำเตือน", "กรุณากรอกรหัสพนักงานและชื่อ")
            return
        
        # --- (!!! ส่วนที่ขาดไป 1: ดึงข้อมูลจากตารางฝึกอบรม !!!) ---
        training_data = []
        for item in self.training_tree.get_children():
            vals = self.training_tree.item(item)["values"]
            training_data.append({
                "date": vals[0],
                "course_name": vals[1],
                "cost": vals[2]
            })

        final_prefix = self.combo_prefix.get()
        if final_prefix == "อื่นๆ":
            final_prefix = self.entry_prefix_other.get().strip()
        # --- (จบส่วนที่เพิ่ม) ---

        employee_data = {
            "id": self.emp_id_entry.get(),
            "prefix": final_prefix,
            "fname": self.fname_entry.get(),
            "nickname": self.nickname_entry.get(),
            "lname": self.lname_entry.get(),
            "emergency_name": self.emergency_name.get(),
            "emergency_phone": self.emergency_phone.get(),
            "emergency_relation": self.emergency_relation.get(),
            "ref_name": self.ref_name.get(),
            "ref_phone": self.ref_phone.get(),
            "ref_relation": self.ref_relation.get(),
            # --- (!!! ส่วนที่แก้ไข 2: ลบวงเล็บ () ออก !!!) ---
            "training_history": training_data, 
            # --- (จบส่วนแก้ไข) ---
            "assets": {
                "computer": self.asset_computer.get(),
                "phone": self.asset_phone.get(),
                "number": self.asset_number.get(),
                "carrier": self.asset_carrier.get(),
                "sim": self.asset_sim.get(),
                "email": self.asset_email.get(),
                "line": self.asset_line.get(),
                "line_phone": self.asset_line_phone.get(),
                "facebook": self.asset_facebook.get(),
                "card_id": self.asset_card_id.get(),
                "others": self.asset_others.get("1.0", "end-1c")
            },
            "birth": self.birth_entry.get(),
            "age": self.age_label.cget("text"),
            "id_card": self.id_card_entry.get(),
            "phone": self.phone_entry.get(),
            "address": self.address_text.get("1.0", "end-1c"),
            "current_address": self.current_address_text.get("1.0", "end-1c"),
            "emp_type": self.emp_type.get(),
            "start_date": self.start_entry.get(),
            "exp": self.exp_label.cget("text"),
            "position": self.position_entry.get(),
            "department": self.department_entry.get(),
            "work_location": self.work_location_combo.get(),
            "status": self.status_combo.get(),
            "termination_date": self.term_date_entry.get() if self.status_combo.get() == "พ้นสภาพพนักงาน" else "",
            "termination_reason": self.term_reason_entry.get() if self.status_combo.get() == "พ้นสภาพพนักงาน" else "",
            "guarantee_enabled": self.guarantee_var.get(),
            "guarantor_name": self.guarantor_entry.get(),
            "guarantee_amount": self.guarantee_amount_entry.get(),
            "probation_days": self.probation_days_combo.get(),
            "probation_end_date": self.probation_end_date_label.cget("text"), 
            "probation_assessment_score": self.probation_score_entry.get(),
            "probation_assessment_score_2": self.probation_score_2_entry.get(),
            "salary": self.salary_entry.get(),
            "health": self.health_combo.get(),
            "health_detail": self.health_detail_entry.get() if self.health_combo.get() == "มี" else "",
            "account": self.account_entry.get(),
            "bank": self.bank_entry.get(),
            "branch": self.branch_entry.get(),
            "account_name": self.account_name_entry.get(),
            "account_type": self.account_type_var.get(),
            "sso_start": self.sso_start_entry.get(),
            "sso_end": self.sso_end_entry.get(),
            "sso_start_action": self.sso_start_action_entry.get(),
            "sso_end_action": self.sso_end_action_entry.get(),
            "sso_hospital": self.sso_hospital_entry.get(),
            "leave_annual": self.leave_annual_entry.get(),
            "leave_sick": self.leave_sick_entry.get(),
            "leave_ordination": self.leave_ordination_entry.get(),
            "leave_maternity": self.leave_maternity_entry.get(),
            "leave_personal": self.leave_personal_entry.get(),
            "welfare_options": self.welfare_options,
            "welfare": [var.get() for var in self.welfare_vars],
            "welfare_amounts": [entry.get() for entry in self.welfare_amount_entries],
            "salary_history": [
                {
                    "year": h["year"].get(), 
                    "salary": h["salary"].get(),
                    "position_allowance": h["position_allowance"].get(),
                    "new_position": h["new_position"].get(),
                    "assessment_score": h["assessment_score"].get()
                } 
                for h in self.salary_history
            ]
        }
        try:
            success, message = hr_database.save_employee(employee_data, self.current_user)
            if success:
                messagebox.showinfo("✅ สำเร็จ", message) 
                self.clear_form()
                self._show_list_page() 
        except Exception as e:
            messagebox.showerror("Save Error", f"เกิดข้อผิดพลาดไม่คาดคิด: {e}")
            
    def _toggle_leave_time_entries(self, event=None):
        """(ใหม่) ซ่อน/แสดง ช่องกรอกเวลาลา"""
        if self.att_leave_duration_type.get() == "ระบุเวลา (ชม.)":
            self.att_leave_time_frame.grid() # (แสดง)
        else:
            self.att_leave_time_frame.grid_remove() # (ซ่อน)
            self.att_leave_start_time.set("")
            self.att_leave_end_time.set("")

    def _save_leave_record(self):
        emp_id = self.emp_id_entry.get() # (ใช้ emp_id_entry จากฟอร์มหลัก)
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาโหลดข้อมูลพนักงานก่อนบันทึกการลา")
            return
            
        leave_date = self.att_leave_date.get_date() 
        if not leave_date:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่ลาให้ครบถ้วน")
            return
        
        leave_type = self.att_leave_type.get()
        duration_type = self.att_leave_duration_type.get()
        reason = self.att_leave_reason.get("1.0", "end-1c")

        if not leave_type or not duration_type:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก วันที่, ประเภท, และระยะเวลา")
            return
        
        num_days = 0.0
        leave_start_time = None
        leave_end_time = None
        
        try:
            if duration_type == "เต็มวัน (1.0)":
                num_days = 1.0
            
            elif duration_type == "ครึ่งวัน (0.5)":
                num_days = 0.5
            
            elif duration_type == "ระบุเวลา (ชม.)":
                start_str = self.att_leave_start_time.get()
                end_str = self.att_leave_end_time.get()
                
                if not start_str or not end_str:
                    messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาระบุเวลา เริ่มต้น และ สิ้นสุด")
                    return
                
                leave_start_time = datetime.strptime(start_str, '%H:%M').time()
                leave_end_time = datetime.strptime(end_str, '%H:%M').time()
                
                if leave_start_time >= leave_end_time:
                    messagebox.showwarning("เวลาผิดพลาด", "เวลาสิ้นสุด ต้องมากกว่าเวลาเริ่มต้น")
                    return
                
                dummy_date = datetime.today().date()
                duration = datetime.combine(dummy_date, leave_end_time) - datetime.combine(dummy_date, leave_start_time)
                hours_total = duration.total_seconds() / 3600
                
                WORK_HOURS_PER_DAY = 8.0 
                num_days = round(hours_total / WORK_HOURS_PER_DAY, 4)
                
        except Exception as e:
            messagebox.showerror("ข้อมูลผิดพลาด", f"ไม่สามารถแปลงข้อมูล: {e}")
            return
            
        success = hr_database.add_employee_leave(
            emp_id, leave_date, leave_type, num_days, reason, 
            leave_start_time, leave_end_time
        )
        
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลการลาเรียบร้อยแล้ว")
            
            # (ล้างค่าในฟอร์มย่อย)
            self.att_leave_date.clear()
            self.att_leave_type.set("")
            self.att_leave_duration_type.set("เต็มวัน (1.0)")
            self._toggle_leave_time_entries() # (สั่งซ่อน)
            self.att_leave_reason.delete("1.0", tk.END)
    # --- (จบฟังก์ชัน "ผ่าตัด" การลา) ---
        
    def _save_late_record(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาโหลดข้อมูลพนักงานก่อนบันทึกการมาสาย")
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

    def _save_warning_record(self):
        emp_id = self.emp_id_entry.get()
        if not emp_id:
            messagebox.showwarning("ข้อผิดพลาด", "กรุณาโหลดข้อมูลพนักงานก่อนบันทึกใบเตือน")
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

    def _build_approval_page(self):
        main_container = ttk.Frame(self.approval_page)
        main_container.pack(fill="both", expand=True, padx=15, pady=10)
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(header_frame, text="✅ รายการรออนุมัติการแก้ไข", style="Header.TLabel").pack(side="left")
        ttk.Button(header_frame, text="🔄 โหลดใหม่", command=self._load_pending_changes).pack(side="right") 
        tree_frame = ttk.LabelFrame(main_container, text="  คำขอแก้ไขข้อมูลพนักงาน  ", padding=15)
        tree_frame.pack(fill="both", expand=True)
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        self.approval_tree = ttk.Treeview(
            tree_container,
            columns=("change_id", "emp_id", "emp_name", "requested_by", "timestamp"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            height=20
        )
        self.approval_tree.heading("change_id", text="ID คำขอ")
        self.approval_tree.heading("emp_id", text="รหัส พนง.")
        self.approval_tree.heading("emp_name", text="ชื่อ-สกุล พนง.")
        self.approval_tree.heading("requested_by", text="ผู้ขอแก้ไข (HR)")
        self.approval_tree.heading("timestamp", text="เวลาที่ขอ")
        self.approval_tree.column("change_id", width=60, anchor="center")
        self.approval_tree.column("emp_id", width=80, anchor="center")
        self.approval_tree.column("emp_name", width=200, anchor="w")
        self.approval_tree.column("requested_by", width=120, anchor="center")
        self.approval_tree.column("timestamp", width=150, anchor="center")
        self.approval_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.approval_tree.yview)
        scrollbar_x.config(command=self.approval_tree.xview)
        btn_frame = ttk.Frame(tree_frame)
        btn_frame.pack(fill="x", pady=(15, 5))
        ttk.Button(btn_frame, text="🔍 ดูรายละเอียด", command=self._view_change_details, width=18).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="✔️ อนุมัติ", command=self._approve_selected_change, width=15, style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ ปฏิเสธ", command=self._reject_selected_change, width=15).pack(side="left", padx=5)
        self._load_pending_changes()

    def _load_pending_changes(self):
        for item in self.approval_tree.get_children():
            self.approval_tree.delete(item)
        pending_list = hr_database.get_pending_changes()
        for item in pending_list:
            req_time = item.get('request_timestamp', None)
            time_str = req_time.strftime('%Y-%m-%d %H:%M') if req_time else '-'
            emp_name = f"{item.get('fname', '')} {item.get('lname', '')}"
            self.approval_tree.insert("", "end", values=(
                item.get('change_id'),
                item.get('emp_id'),
                emp_name,
                item.get('requested_by'),
                time_str
            ))

    def _get_selected_change_id(self):
        selection = self.approval_tree.selection()
        if not selection:
            messagebox.showwarning("⚠️ ไม่ได้เลือก", "กรุณาเลือกรายการที่ต้องการ")
            return None
        item = self.approval_tree.item(selection[0])
        return item["values"][0] 

    def _view_change_details(self):
        change_id = self._get_selected_change_id()
        if not change_id: return
        new_data, current_data = hr_database.get_change_details(change_id)
        if not new_data: return 
        detail_window = tk.Toplevel(self)
        detail_window.title(f"รายละเอียดคำขอแก้ไข #{change_id}")
        detail_window.geometry("800x600")
        detail_window.transient(self)
        detail_window.grab_set()
        text_widget = tk.Text(detail_window, wrap="word", font=("Consolas", 10))
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        details = f"คำขอแก้ไข ID: {change_id}\n"
        details += f"รหัสพนักงาน: {new_data.get('id')}\n\n"
        details += "{:<25} | {:<30} | {:<30}\n".format("ฟิลด์", "ข้อมูลปัจจุบัน", "ข้อมูลที่ขอแก้ไข")
        details += "-"*80 + "\n"
        all_keys = set(current_data.keys() if current_data else []) | set(new_data.keys())
        ignore_keys = {'welfare', 'welfare_amounts', 'salary_history', 'welfare_options'} 
        for key in sorted(list(all_keys)):
             if key not in ignore_keys:
                current_val = str(current_data.get(key, '')) if current_data else 'N/A'
                new_val = str(new_data.get(key, ''))
                if current_val != new_val: 
                    details += "{:<25} | {:<30} | {:<30} <--- CHANGED\n".format(key, current_val, new_val)
                else:
                    details += "{:<25} | {:<30} | {:<30}\n".format(key, current_val, new_val)
        text_widget.insert("1.0", details)
        text_widget.config(state="disabled") 
        ttk.Button(detail_window, text="ปิด", command=detail_window.destroy).pack(pady=10)

    def show_approval_page(self):
        if self.current_user.get("role") == 'approver':
            self._load_pending_changes() 
            self.approval_page.tkraise()
        else:
            messagebox.showwarning("Permission Denied", "คุณไม่มีสิทธิ์เข้าถึงหน้านี้")

    def _approve_selected_change(self):
        change_id = self._get_selected_change_id()
        if not change_id: return

        print(f"DEBUG: กำลังพยายามอนุมัติ Change ID: {change_id}") # <--- Debug

        if messagebox.askyesno("ยืนยันการอนุมัติ", f"คุณต้องการอนุมัติคำขอแก้ไข ID: {change_id} ใช่หรือไม่?"):
            try:
                # ส่ง username ของคนกดอนุมัติไปด้วย
                approver = self.current_user['username']
                print(f"DEBUG: ผู้กดอนุมัติคือ {approver}") # <--- Debug
                
                # เรียกฟังก์ชันฐานข้อมูล
                success = hr_database.approve_change(change_id, approver)
                
                print(f"DEBUG: ผลลัพธ์จาก Database: {success}") # <--- Debug

                if success:
                    messagebox.showinfo("✅ สำเร็จ", "อนุมัติการแก้ไขเรียบร้อยแล้ว\nข้อมูลพนักงานถูกอัปเดตแล้ว")
                    self._load_pending_changes() # รีเฟรชตารางรออนุมัติ
                    
                    # (Optional) ถ้าต้องการให้รีเฟรชรายการพนักงานหลักด้วย (ถ้าทำได้)
                    # self.update_employee_list(hr_database.load_all_employees()) 
                else:
                    messagebox.showerror("ผิดพลาด", "ระบบแจ้งว่าการอนุมัติไม่สำเร็จ\n(กรุณาดู Error ในจอดำ)")
            
            except Exception as e:
                print(f"DEBUG: เกิด Error ร้ายแรงที่หน้าจอ: {e}")
                import traceback
                traceback.print_exc()
                messagebox.showerror("Exception", f"เกิดข้อผิดพลาดร้ายแรง:\n{e}")

    def _reject_selected_change(self):
        change_id = self._get_selected_change_id()
        if not change_id: return
        if messagebox.askyesno("ยืนยันการปฏิเสธ", f"คุณต้องการปฏิเสธคำขอแก้ไข ID: {change_id} ใช่หรือไม่?"):
            success = hr_database.reject_change(change_id, self.current_user['username'])
            if success:
                messagebox.showinfo("🗑️ สำเร็จ", "ปฏิเสธคำขอแก้ไขเรียบร้อยแล้ว")
                self._load_pending_changes() 


    def load_employee_data(self, employee):
        self.clear_form()
        
        # --- 1. ข้อมูลพื้นฐาน ---
        self.emp_id_entry.insert(0, str(employee.get("id", "") or ""))
        prefix = str(employee.get("prefix", "") or "")
        if prefix in ["นาย", "นาง", "นางสาว"]:
            self.combo_prefix.set(prefix)
            self.entry_prefix_other.pack_forget()
        elif not prefix:
            self.combo_prefix.set("")
            self.entry_prefix_other.pack_forget()
        else:
            # ถ้าเป็นคำอื่นๆ (เช่น ดร., ทพ.)
            self.combo_prefix.set("อื่นๆ")
            self.entry_prefix_other.pack(side="left", padx=5)
            self.entry_prefix_other.delete(0, tk.END)
            self.entry_prefix_other.insert(0, prefix)
        self.fname_entry.insert(0, str(employee.get("fname", "") or ""))
        self.nickname_entry.insert(0, str(employee.get("nickname", "") or ""))
        self.lname_entry.insert(0, str(employee.get("lname", "") or ""))
        self.birth_entry.set_date_from_str(str(employee.get("birth", "") or ""))
        self.age_label.config(text=str(employee.get("age", "-")))
        self.id_card_entry.insert(0, str(employee.get("id_card", "") or ""))
        self.phone_entry.insert(0, str(employee.get("phone", "") or ""))
        
        # --- 2. บุคคลติดต่อฉุกเฉิน & อ้างอิง ---
        try:
            self.emergency_name.delete(0, tk.END)
            self.emergency_name.insert(0, str(employee.get("emergency_name", "") or ""))
            
            self.emergency_phone.delete(0, tk.END)
            self.emergency_phone.insert(0, str(employee.get("emergency_phone", "") or ""))
            
            self.emergency_relation.delete(0, tk.END)
            self.emergency_relation.insert(0, str(employee.get("emergency_relation", "") or ""))
            
            self.ref_name.delete(0, tk.END)
            self.ref_name.insert(0, str(employee.get("ref_name", "") or ""))
            
            self.ref_phone.delete(0, tk.END)
            self.ref_phone.insert(0, str(employee.get("ref_phone", "") or ""))
            
            self.ref_relation.delete(0, tk.END)
            self.ref_relation.insert(0, str(employee.get("ref_relation", "") or ""))
        except AttributeError:
            pass 

        # --- 3. ที่อยู่ ---
        self.address_text.delete("1.0", tk.END)
        self.address_text.insert("1.0", str(employee.get("address", "") or ""))
        self.current_address_text.delete("1.0", tk.END)
        self.current_address_text.insert("1.0", str(employee.get("current_address", "") or ""))
        
        # --- 4. ข้อมูลการจ้าง ---
        self.emp_type.set(str(employee.get("emp_type", "") or ""))
        self.start_entry.set_date_from_str(str(employee.get("start_date", "") or ""))
        self.exp_label.config(text=str(employee.get("exp", "-")))
        self.position_entry.insert(0, str(employee.get("position", "") or ""))
        self.department_entry.insert(0, str(employee.get("department", "") or ""))
        self.work_location_combo.set(str(employee.get("work_location", "") or ""))
        self.status_combo.set(str(employee.get("status", "") or ""))
        
        self.term_date_entry.set_date_from_str(str(employee.get("termination_date", "") or ""))
        self.term_reason_entry.delete(0, tk.END)
        self.term_reason_entry.insert(0, str(employee.get("termination_reason", "") or ""))
        self._toggle_termination_fields() 

        # --- 5. เงินเดือน/สวัสดิการ ---
        self.salary_entry.insert(0, str(employee.get("salary", "") or ""))
        
        welfare_flags = employee.get("welfare", [False] * len(self.welfare_vars))
        welfare_amounts = employee.get("welfare_amounts", [""] * len(self.welfare_amount_entries))
        for i, var in enumerate(self.welfare_vars):
            var.set(welfare_flags[i] if i < len(welfare_flags) else False)
        for i, entry in enumerate(self.welfare_amount_entries):
            entry.delete(0, tk.END)
            entry.insert(0, str(welfare_amounts[i] if i < len(welfare_amounts) else ""))

        for i, hist_widget in enumerate(self.salary_history):
            hist_data = employee.get("salary_history", [])
            hist_widget["salary"].delete(0, tk.END)
            hist_widget["position_allowance"].delete(0, tk.END)
            hist_widget["new_position"].delete(0, tk.END)
            hist_widget["assessment_score"].delete(0, tk.END)
            
            if i < len(hist_data):
                hist_widget["year"].set(str(hist_data[i].get("year", "") or ""))
                hist_widget["salary"].insert(0, str(hist_data[i].get("salary") or ""))
                hist_widget["position_allowance"].insert(0, str(hist_data[i].get("position_allowance") or ""))
                hist_widget["new_position"].insert(0, str(hist_data[i].get("new_position") or ""))
                hist_widget["assessment_score"].insert(0, str(hist_data[i].get("assessment_score") or ""))
            else:
                hist_widget["year"].set("")

        # --- 6. ธนาคารและอื่นๆ ---
        health_status = str(employee.get("health", "ไม่มี") or "ไม่มี")
        self.health_combo.set(health_status)
        self.toggle_health_detail()
        if health_status == "มี":
            self.health_detail_entry.delete(0, tk.END)
            self.health_detail_entry.insert(0, str(employee.get("health_detail", "") or ""))
            
        self.account_entry.insert(0, str(employee.get("account", "") or ""))
        self.bank_entry.insert(0, str(employee.get("bank", "") or ""))
        self.branch_entry.insert(0, str(employee.get("branch", "") or ""))
        self.account_name_entry.insert(0, str(employee.get("account_name", "") or ""))
        self.account_type_var.set(str(employee.get("account_type", "") or ""))
        
        self.sso_start_entry.set_date_from_str(str(employee.get("sso_start", "") or ""))
        self.sso_end_entry.set_date_from_str(str(employee.get("sso_end", "") or ""))
        self.sso_start_action_entry.set_date_from_str(str(employee.get("sso_start_action", "") or ""))
        self.sso_end_action_entry.set_date_from_str(str(employee.get("sso_end_action", "") or ""))
        self.sso_hospital_entry.insert(0, str(employee.get("sso_hospital", "") or ""))
        
        self.leave_annual_entry.insert(0, str(employee.get("leave_annual", "") or ""))
        self.leave_sick_entry.insert(0, str(employee.get("leave_sick", "") or ""))
        self.leave_ordination_entry.insert(0, str(employee.get("leave_ordination", "") or ""))
        self.leave_maternity_entry.insert(0, str(employee.get("leave_maternity", "") or ""))
        self.leave_personal_entry.insert(0, str(employee.get("leave_personal", "") or ""))

        # --- 7. ค้ำประกัน & ทดลองงาน ---
        self.guarantee_var.set(bool(employee.get("guarantee_enabled", False)))
        self._toggle_guarantee_fields()
        if self.guarantee_var.get():
            self.guarantor_entry.delete(0, tk.END)
            self.guarantor_entry.insert(0, str(employee.get("guarantor_name", "") or ""))
            self.guarantee_amount_entry.delete(0, tk.END)
            self.guarantee_amount_entry.insert(0, str(employee.get("guarantee_amount", "") or ""))
            self._load_guarantee_doc_status()

        self.probation_days_combo.set(str(employee.get("probation_days", "90") or "90"))
        self.probation_score_entry.delete(0, tk.END)
        self.probation_score_entry.insert(0, str(employee.get("probation_assessment_score") or ""))
        self.probation_score_2_entry.delete(0, tk.END)
        self.probation_score_2_entry.insert(0, str(employee.get("probation_assessment_score_2") or ""))
        
        self._toggle_probation_scores()
        self._calculate_probation_end_date()
        self._check_score(self.probation_score_entry, self.probation_status_label)
        self._check_score(self.probation_score_2_entry, self.probation_status_2_label)
        self._load_warning_doc_status()

        # --- 8. ประวัติการฝึกอบรม ---
        for item in self.training_tree.get_children():
            self.training_tree.delete(item)
        train_history = employee.get("training_history", [])
        for record in train_history:
            self.training_tree.insert("", "end", values=(
                str(record.get("date", "") or ""),
                str(record.get("course_name", "") or ""),
                str(record.get("cost", "0.00") or "0.00")
            ))

        # --- 9. (!!! เพิ่มส่วนนี้ !!!) ทรัพย์สิน (Assets) ---
        assets = employee.get("assets", {})
        try:
            self.asset_computer.delete(0, tk.END); self.asset_computer.insert(0, str(assets.get("computer", "") or ""))
            self.asset_phone.delete(0, tk.END); self.asset_phone.insert(0, str(assets.get("phone", "") or ""))
            self.asset_number.delete(0, tk.END); self.asset_number.insert(0, str(assets.get("number", "") or ""))
            self.asset_carrier.set(str(assets.get("carrier", "") or ""))
            self.asset_sim.set(str(assets.get("sim", "") or ""))
            self.asset_email.delete(0, tk.END); self.asset_email.insert(0, str(assets.get("email", "") or ""))
            self.asset_line.delete(0, tk.END); self.asset_line.insert(0, str(assets.get("line", "") or ""))
            self.asset_line_phone.delete(0, tk.END); self.asset_line_phone.insert(0, str(assets.get("line_phone", "") or ""))
            self.asset_facebook.delete(0, tk.END); self.asset_facebook.insert(0, str(assets.get("facebook", "") or ""))
            self.asset_card_id.delete(0, tk.END); self.asset_card_id.insert(0, str(assets.get("card_id", "") or ""))
            self.asset_others.delete("1.0", tk.END); self.asset_others.insert("1.0", str(assets.get("others", "") or ""))
        except AttributeError:
            pass

    def delete_employee(self):
        selection = self.employee_tree.selection()
        if not selection:
            messagebox.showwarning("⚠️ คำเตือน", "กรุณาเลือกพนักงานที่ต้องการลบ")
            return
        item = self.employee_tree.item(selection[0])
        emp_id = item["values"][0]
        emp_name = item["values"][1]
        if messagebox.askyesno("ยืนยัน", f"ต้องการลบข้อมูลพนักงาน {emp_name} (ID: {emp_id}) ใช่หรือไม่?"):
            success = hr_database.delete_employee(emp_id)
            if success:
                messagebox.showinfo("สำเร็จ", "ลบข้อมูลแล้ว")
                all_employees = hr_database.load_all_employees()
                self.update_employee_list(all_employees)
            if self.emp_id_entry.get() == emp_id:
                self.clear_form()

    def clear_form(self):
        self.combo_prefix.set("")
        self.entry_prefix_other.delete(0, tk.END)
        self.entry_prefix_other.pack_forget()
        self.emergency_name.delete(0, tk.END)
        self.emergency_phone.delete(0, tk.END)
        self.emergency_relation.delete(0, tk.END)
        self.ref_name.delete(0, tk.END)
        self.ref_phone.delete(0, tk.END)
        self.ref_relation.delete(0, tk.END)
        self.emp_id_entry.delete(0, tk.END)
        self.fname_entry.delete(0, tk.END)
        self.nickname_entry.delete(0, tk.END)
        self.lname_entry.delete(0, tk.END)
        self.birth_entry.delete(0, tk.END)
        self.age_label.config(text="-")
        self.id_card_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.address_text.delete("1.0", tk.END)
        self.current_address_text.delete("1.0", tk.END)
        self.emp_type.set("")
        self.start_entry.delete(0, tk.END)
        self.exp_label.config(text="-")
        self.position_entry.delete(0, tk.END)
        self.department_entry.delete(0, tk.END)
        self.work_location_combo.set("")
        self.status_combo.set("")
        self._toggle_termination_fields()
        self.salary_entry.delete(0, tk.END)
        for item in self.training_tree.get_children():
            self.training_tree.delete(item)
        self.train_date_input.clear()
        self.train_name_input.delete(0, tk.END)
        self.train_cost_input.delete(0, tk.END)
        for var in self.welfare_vars:
            var.set(False)
        for entry in self.welfare_amount_entries:
            entry.delete(0, tk.END)
        for hist in self.salary_history:
            hist["year"].set("")
            hist["salary"].delete(0, tk.END)
            hist["position_allowance"].delete(0, tk.END)
            hist["new_position"].delete(0, tk.END)      
            hist["assessment_score"].delete(0, tk.END)
        self.health_combo.set("ไม่มี")
        self.toggle_health_detail()
        self.account_entry.delete(0, tk.END)
        self.bank_entry.delete(0, tk.END)
        self.branch_entry.delete(0, tk.END)
        self.account_name_entry.delete(0, tk.END)
        self.account_type_var.set("")
        self.sso_start_entry.delete(0, tk.END)
        self.sso_end_entry.delete(0, tk.END)
        self.sso_start_action_entry.delete(0, tk.END)
        self.sso_end_action_entry.delete(0, tk.END)
        self.sso_hospital_entry.delete(0, tk.END)
        self.leave_annual_entry.delete(0, tk.END)
        self.leave_sick_entry.delete(0, tk.END)
        self.leave_ordination_entry.delete(0, tk.END)
        self.leave_maternity_entry.delete(0, tk.END)
        self.leave_personal_entry.delete(0, tk.END)
        self.guarantee_var.set(False)
        self._toggle_guarantee_fields() 
        self._load_warning_doc_status()
        self._load_guarantee_doc_status()
        self.probation_days_combo.set("90") 
        self.probation_end_date_label.config(text="-")
        self.probation_score_entry.delete(0, tk.END)
        self.probation_status_label.config(text="")
        self.probation_score_2_entry.delete(0, tk.END)
        self.probation_status_2_label.config(text="")
        self._toggle_probation_scores()
        
        # (ล้างฟอร์มย่อยในแท็บ 5)
        try:
            self.att_leave_date.clear()
            self.att_leave_type.set("")
            self.att_leave_duration_type.set("เต็มวัน (1.0)")
            self._toggle_leave_time_entries()
            self.att_leave_reason.delete("1.0", tk.END)
            
            self.att_late_date.clear()
            self.att_late_minutes.delete(0, tk.END)
            self.att_late_reason.delete("1.0", tk.END)

            self.att_warn_date.clear()
            self.att_warn_level.set("") 
            self.att_warn_reason.delete("1.0", tk.END)
        
        except AttributeError:
            pass # (ถ้า UI ยังไม่ถูกสร้าง)
            
        try:
            self.asset_computer.delete(0, tk.END)
            self.asset_phone.delete(0, tk.END)
            self.asset_number.delete(0, tk.END)
            self.asset_carrier.set("")
            self.asset_sim.set("")
            self.asset_email.delete(0, tk.END)
            self.asset_line.delete(0, tk.END)
            self.asset_line_phone.delete(0, tk.END)
            self.asset_facebook.delete(0, tk.END)
            self.asset_card_id.delete(0, tk.END)
            self.asset_others.delete("1.0", tk.END)
        except AttributeError:
            pass    
        
    def update_employee_list(self, employees):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        for emp in employees:
            display_name = f"{emp.get('fname','')} {emp.get('lname','')}"
            nick = (emp.get("nickname") or "").strip()
            if nick:
                display_name = f"{emp.get('fname','')} ({nick}) {emp.get('lname','')}"
            status = emp.get("status", "")
            self.employee_tree.insert("", "end", values=(
                emp.get("id", ""),
                display_name,
                emp.get("phone", ""),
                emp.get("position", ""),
                emp.get("department", ""),
                status,
                emp.get("id_card", ""),
                emp.get("salary", "") 
            ))
        total_emp = len(employees)
        self.summary_label.config(text=f"📊 จำนวนพนักงานทั้งหมด: {total_emp} คน")
