# (ไฟล์ใหม่: company_profile_module.py)

import tkinter as tk
from tkinter import ttk, messagebox
import hr_database 
from custom_widgets import DateDropdown # (ใช้ Widget ที่เราสร้างไว้)
from datetime import datetime
import calendar 

class CompanyProfileModule(ttk.Frame):
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user

        # (ตัวแปรสำหรับแปลงค่า)
        self.THAI_MONTHS = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
            5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
            9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }
        
        # --- สร้าง UI ---
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="🏢 ข้อมูลบริษัท (Company Profile)", 
                  font=("Segoe UI", 16, "bold"), foreground="#2c3e50").pack(anchor="w", pady=(0, 15))

        # --- สร้าง Notebook (แท็บ) ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 0))

        # --- แท็บ 1: ข้อมูลบริษัท (ที่อยู่, Tax, Report) ---
        self.tab_info = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_info, text="  📄 ข้อมูลบริษัทและสาขา  ")
        self._build_company_info_tab(self.tab_info) 

        # --- แท็บ 2: ตั้งค่าวันหยุด (ย้ายมาจาก Time Processor) ---
        self.tab_holiday = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_holiday, text="  📅 ตั้งค่าวันหยุดประจำปี  ")
        self._build_holiday_tab(self.tab_holiday)

    
    def _build_company_info_tab(self, parent_tab):
        """สร้าง UI สำหรับแท็บ 'ข้อมูลบริษัท'"""
        
        # (ใช้ Canvas + Scrollbar)
        canvas = tk.Canvas(parent_tab)
        scrollbar = ttk.Scrollbar(parent_tab, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # --- (กรอบ 6.1: ข้อมูลบริษัท / ที่อยู่สาขา) ---
        loc_frame = ttk.LabelFrame(scroll_frame, text="  6.1 ข้อมูลบริษัท / ที่อยู่สาขา  ", padding=20)
        loc_frame.pack(fill="x", pady=(0, 15))

        # (ส่วนฟอร์มสำหรับ "เพิ่ม")
        add_loc_frame = ttk.Frame(loc_frame)
        add_loc_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(add_loc_frame, text="ชื่อสถานที่:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.loc_name_entry = ttk.Entry(add_loc_frame, width=25, font=("Segoe UI", 10))
        self.loc_name_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(add_loc_frame, text="ประเภท:", font=("Segoe UI", 10)).grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.loc_type_combo = ttk.Combobox(add_loc_frame, values=["สำนักงานใหญ่", "สาขา", "คลังสินค้า"], 
                                           width=15, state="readonly", font=("Segoe UI", 10))
        self.loc_type_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(add_loc_frame, text="Google Link:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.loc_link_entry = ttk.Entry(add_loc_frame, width=50, font=("Segoe UI", 10))
        self.loc_link_entry.grid(row=1, column=1, columnspan=3, sticky="w", padx=5, pady=5)

        ttk.Button(add_loc_frame, text="💾 บันทึก/อัปเดต สาขา", command=self._save_location).grid(row=1, column=4, sticky="w", padx=10, pady=5)

        # (ส่วนตาราง "แสดงผล")
        loc_tree_frame = ttk.Frame(loc_frame)
        loc_tree_frame.pack(fill="x", expand=True, pady=(10,0))
        
        loc_scrollbar = ttk.Scrollbar(loc_tree_frame, orient="vertical")
        self.loc_tree = ttk.Treeview(
            loc_tree_frame,
            columns=("id", "name", "type", "link"),
            show="headings",
            yscrollcommand=loc_scrollbar.set,
            height=5
        )
        loc_scrollbar.config(command=self.loc_tree.yview)
        loc_scrollbar.pack(side="right", fill="y")
        
        self.loc_tree.heading("id", text="ID")
        self.loc_tree.heading("name", text="ชื่อสถานที่")
        self.loc_tree.heading("type", text="ประเภท")
        self.loc_tree.heading("link", text="Google Link")
        self.loc_tree.column("id", width=50, anchor="center")
        self.loc_tree.column("name", width=200, anchor="w")
        self.loc_tree.column("type", width=150, anchor="w")
        self.loc_tree.column("link", width=300, anchor="w")
        self.loc_tree.pack(fill="x", expand=True)

        ttk.Button(loc_frame, text="🗑️ ลบสาขาที่เลือก", command=self._delete_location).pack(anchor="w", pady=5)
        self.loc_tree.bind("<Double-1>", self._load_location_to_form) # (ดับเบิลคลิกเพื่อแก้ไข)
        
        
        # --- (กรอบ 6.2: เลขประจำตัวผู้เสียภาษี) ---
        tax_frame = ttk.LabelFrame(scroll_frame, text="  6.2 เลขประจำตัวผู้เสียภาษี  ", padding=20)
        tax_frame.pack(fill="x", pady=15)
        
        ttk.Label(tax_frame, text="Tax ID:", font=("Segoe UI", 10)).pack(side="left", padx=5)
        self.tax_id_entry = ttk.Entry(tax_frame, width=30, font=("Segoe UI", 10))
        self.tax_id_entry.pack(side="left", padx=5)
        ttk.Button(tax_frame, text="💾 บันทึก Tax ID", command=self._save_tax_id).pack(side="left", padx=10)


        # --- (กรอบ 6.3: จำนวนพนักงาน) ---
        report_frame = ttk.LabelFrame(scroll_frame, text="  6.3 จำนวนพนักงานแต่ละแผนก  ", padding=20)
        report_frame.pack(fill="both", expand=True, pady=15)
        
        report_filter_frame = ttk.Frame(report_frame)
        report_filter_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(report_filter_frame, text="เลือกปี (พ.ศ.):", font=("Segoe UI", 10)).pack(side="left", padx=5)
        current_year_be = datetime.now().year + 543
        year_values = [str(y) for y in range(current_year_be + 1, current_year_be - 5, -1)]
        self.report_year_combo = ttk.Combobox(report_filter_frame, values=year_values, width=10, state="readonly", font=("Segoe UI", 10))
        self.report_year_combo.set(str(current_year_be))
        self.report_year_combo.pack(side="left", padx=5)
        ttk.Button(report_filter_frame, text="🔄 โหลดรายงาน", command=self._load_dept_report).pack(side="left", padx=5)
        
        # (ตารางแสดงผล Report)
        report_tree_frame = ttk.Frame(report_frame)
        report_tree_frame.pack(fill="both", expand=True, pady=(10,0))
        
        report_scrollbar = ttk.Scrollbar(report_tree_frame, orient="vertical")
        self.report_tree = ttk.Treeview(
            report_tree_frame,
            columns=("dept", "count"),
            show="headings",
            yscrollcommand=report_scrollbar.set,
            height=10
        )
        report_scrollbar.config(command=self.report_tree.yview)
        report_scrollbar.pack(side="right", fill="y")
        
        self.report_tree.heading("dept", text="แผนก/ฝ่าย")
        self.report_tree.heading("count", text="จำนวน (คน)")
        self.report_tree.column("dept", width=300, anchor="w")
        self.report_tree.column("count", width=100, anchor="center")
        self.report_tree.pack(fill="both", expand=True)
        
        # (Label สรุปยอดรวม)
        self.report_total_label = ttk.Label(report_frame, text="รวมทั้งหมด: 0 คน", font=("Segoe UI", 10, "bold"))
        self.report_total_label.pack(anchor="e", pady=5)
        
        
        # --- (ผูก Canvas) ---
        canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: self._on_mousewheel(ev, c)))
        canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- (โหลดข้อมูลครั้งแรก) ---
        self._load_company_info_data()

    
    def _build_holiday_tab(self, parent_tab):
        """
        (ย้ายมา) สร้าง UI สำหรับแท็บ 'ตั้งค่าวันหยุด'
        (โค้ดนี้คัดลอกมาจาก time_processor_module.py)
        """
        
        main_holiday_frame = ttk.Frame(parent_tab)
        main_holiday_frame.pack(fill="both", expand=True)
        main_holiday_frame.columnconfigure(1, weight=1) 
        main_holiday_frame.rowconfigure(1, weight=1) 

        add_frame = ttk.LabelFrame(main_holiday_frame, text="  เพิ่ม/แก้ไข วันหยุด  ", padding=15)
        add_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 15), pady=5)
        
        ttk.Label(add_frame, text="เลือกวันที่:", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 5))
        
        self.holiday_date_entry = DateDropdown(add_frame, font=("Segoe UI", 10))
        self.holiday_date_entry.pack(anchor="w", fill="x", pady=5)
        
        ttk.Label(add_frame, text="คำอธิบาย (เช่น วันพ่อ):", font=("Segoe UI", 10)).pack(anchor="w", pady=(10, 5))
        self.holiday_desc_entry = ttk.Entry(add_frame, width=30, font=("Segoe UI", 10))
        self.holiday_desc_entry.pack(anchor="w", fill="x", pady=5)
        
        ttk.Button(add_frame, text="💾 บันทึกวันหยุด", command=self._save_new_holiday, 
                   style="Success.TButton").pack(anchor="w", fill="x", pady=(15, 5))

        filter_frame = ttk.Frame(main_holiday_frame)
        filter_frame.grid(row=0, column=1, sticky="new", pady=5)
        
        ttk.Label(filter_frame, text="เลือกปี (พ.ศ.) เพื่อแสดงผล:", font=("Segoe UI", 10)).pack(side="left", padx=5)
        current_year_be = datetime.now().year + 543
        year_values = [str(y) for y in range(current_year_be + 1, current_year_be - 5, -1)]
        self.holiday_year_combo = ttk.Combobox(filter_frame, values=year_values, width=10, state="readonly", font=("Segoe UI", 10))
        self.holiday_year_combo.set(str(current_year_be))
        self.holiday_year_combo.pack(side="left", padx=5)
        
        ttk.Button(filter_frame, text="🔄 โหลดข้อมูล", command=self._load_holidays_to_tree).pack(side="left", padx=5)

        list_frame = ttk.LabelFrame(main_holiday_frame, text="  รายการวันหยุดประจำปี (ที่เลือก)  ", padding=15)
        list_frame.grid(row=1, column=1, sticky="nsew", pady=(10, 5))
        list_frame.rowconfigure(0, weight=1) 
        list_frame.columnconfigure(0, weight=1)

        tree_container = ttk.Frame(list_frame)
        tree_container.grid(row=0, column=0, sticky="nsew")
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        
        self.holiday_tree = ttk.Treeview(
            tree_container,
            columns=("id", "date", "desc"),
            show="headings",
            yscrollcommand=scrollbar.set,
            height=15
        )
        self.holiday_tree.heading("id", text="ID")
        self.holiday_tree.heading("date", text="วันที่")
        self.holiday_tree.heading("desc", text="คำอธิบาย")
        self.holiday_tree.column("id", width=50, anchor="center")
        self.holiday_tree.column("date", width=120, anchor="w")
        self.holiday_tree.column("desc", width=300, anchor="w")
        
        self.holiday_tree.pack(fill="both", expand=True)
        scrollbar.config(command=self.holiday_tree.yview)
        
        ttk.Button(list_frame, text="🗑️ ลบวันที่เลือก", command=self._delete_selected_holiday).grid(row=1, column=0, sticky="w", pady=(10,0))
        
        self._load_holidays_to_tree() # (โหลดข้อมูลครั้งแรก)

    
    # --- (ฟังก์ชัน Helper สำหรับ Scrollbar) ---
    def _on_mousewheel(self, event, widget):
        try:
            widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass 

    # --- (ฟังก์ชัน Helper สำหรับ "แท็บ 1: ข้อมูลบริษัท") ---
    
    def _load_company_info_data(self):
        """(ใหม่) โหลดข้อมูล Tax ID และ ที่อยู่สาขา"""
        # 1. โหลด Tax ID
        tax_id = hr_database.get_company_setting("tax_id")
        self.tax_id_entry.delete(0, tk.END)
        self.tax_id_entry.insert(0, tax_id)
        
        # 2. โหลดที่อยู่
        self._load_locations_to_tree()
        
        # 3. โหลด Report
        self._load_dept_report()

    def _save_tax_id(self):
        """(ใหม่) บันทึก Tax ID"""
        tax_id = self.tax_id_entry.get().strip()
        if not tax_id:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก Tax ID")
            return
        
        success = hr_database.save_company_setting("tax_id", tax_id)
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึก Tax ID เรียบร้อย")

    def _load_locations_to_tree(self):
        """(ใหม่) โหลดข้อมูลสาขาลงใน Treeview"""
        for item in self.loc_tree.get_children():
            self.loc_tree.delete(item)
            
        locations = hr_database.get_company_locations()
        for loc in locations:
            self.loc_tree.insert("", "end", iid=loc['loc_id'], values=(
                loc['loc_id'],
                loc['loc_name'],
                loc.get('loc_type', ''),
                loc.get('google_link', '')
            ))
            
    def _load_location_to_form(self, event=None):
        """(ใหม่) เมื่อดับเบิลคลิก ให้นำข้อมูลไปใส่ในฟอร์มแก้ไข"""
        selection = self.loc_tree.selection()
        if not selection: return
        
        item = self.loc_tree.item(selection[0])
        loc_id, name, loc_type, link = item["values"]
        
        self.loc_name_entry.delete(0, tk.END)
        self.loc_name_entry.insert(0, name)
        self.loc_type_combo.set(loc_type)
        self.loc_link_entry.delete(0, tk.END)
        self.loc_link_entry.insert(0, link)

    def _save_location(self):
        """(ใหม่) บันทึกข้อมูลสาขา (เพิ่มใหม่ หรือ อัปเดต)"""
        name = self.loc_name_entry.get().strip()
        loc_type = self.loc_type_combo.get()
        link = self.loc_link_entry.get().strip()
        
        if not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก 'ชื่อสถานที่'")
            return
            
        success = hr_database.add_company_location(name, loc_type, link)
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลสาขาเรียบร้อย")
            self.loc_name_entry.delete(0, tk.END)
            self.loc_type_combo.set("")
            self.loc_link_entry.delete(0, tk.END)
            self._load_locations_to_tree()
            
    def _delete_location(self):
        """(ใหม่) ลบข้อมูลสาขาที่เลือก"""
        selection = self.loc_tree.selection()
        if not selection:
            messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือกสาขาที่ต้องการลบในตาราง")
            return
            
        item = self.loc_tree.item(selection[0])
        loc_id = item["values"][0]
        loc_name = item["values"][1]
        
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบสาขานี้ใช่หรือไม่?\n\n- {loc_name}"):
            return
            
        success = hr_database.delete_company_location(loc_id)
        if success:
            messagebox.showinfo("สำเร็จ", "ลบข้อมูลสาขาเรียบร้อย")
            self._load_locations_to_tree()

    def _load_dept_report(self):
        """(ใหม่) โหลด Report (6.3) นับจำนวนพนักงาน"""
        try:
            year_be = int(self.report_year_combo.get())
        except Exception:
            messagebox.showwarning("ผิดพลาด", "กรุณาเลือกปี (พ.ศ.) ให้ถูกต้อง")
            return
            
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
            
        report_data = hr_database.get_employee_count_by_dept(year_be)
        
        total_count = 0
        for row in report_data:
            dept_name = row['dept']
            count = row['count']
            self.report_tree.insert("", "end", values=(
                dept_name,
                f"{count} คน"
            ))
            total_count += count
            
        self.report_total_label.config(text=f"รวมทั้งหมด: {total_count} คน")


    # --- (ฟังก์ชัน Helper สำหรับ "แท็บ 2: วันหยุด") ---
    # (โค้ดนี้คัดลอกมาจาก time_processor_module.py)

    def _load_holidays_to_tree(self):
        try:
            year_be_str = self.holiday_year_combo.get()
            year_ce = int(year_be_str) - 543
        except Exception:
            messagebox.showwarning("ผิดพลาด", "กรุณาเลือกปี (พ.ศ.) ให้ถูกต้อง")
            return

        for item in self.holiday_tree.get_children():
            self.holiday_tree.delete(item)
            
        holidays = hr_database.get_company_holidays(year_ce)
        for holiday in holidays:
            date_obj = holiday['holiday_date']
            date_be_str = f"{date_obj.day:02d}/{date_obj.month:02d}/{date_obj.year + 543}"
            
            self.holiday_tree.insert("", "end", iid=holiday['holiday_id'], values=(
                holiday['holiday_id'],
                date_be_str,
                holiday['description']
            ))

    def _save_new_holiday(self):
        try:
            holiday_date_obj = self.holiday_date_entry.get_date()
            if not holiday_date_obj: 
                raise ValueError("วันที่ไม่ได้เลือก")
        except Exception:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกวันที่จากปฏิทินให้ครบถ้วน")
            return
            
        description = self.holiday_desc_entry.get().strip()
        if not description:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกคำอธิบายวันหยุด")
            return
            
        success = hr_database.add_company_holiday(holiday_date_obj, description)
        if success:
            messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลวันหยุดเรียบร้อย")
            self.holiday_date_entry.clear()
            self.holiday_desc_entry.delete(0, tk.END)
            self._load_holidays_to_tree() 
            
    def _delete_selected_holiday(self):
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือกวันหยุดที่ต้องการลบในตาราง")
            return
            
        item = self.holiday_tree.item(selection[0])
        holiday_id = item["values"][0]
        holiday_desc = f"{item['values'][1]} ({item['values'][2]})"
        
        if not messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบวันหยุดนี้ใช่หรือไม่?\n\n- {holiday_desc}"):
            return
            
        success = hr_database.delete_company_holiday(holiday_id)
        if success:
            messagebox.showinfo("สำเร็จ", "ลบวันหยุดเรียบร้อย")
            self._load_holidays_to_tree()