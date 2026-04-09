# (ไฟล์: time_processor_module.py)
# (เวอร์ชันอัปเกรด - แก้ไข AttributeError ของ tksheet และ Grid Layout)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from custom_widgets import DateDropdown 
import hr_database 
from datetime import datetime
import os 
import calendar 
from tksheet import Sheet 

class TimeProcessorModule(ttk.Frame):
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user
        
        self.raw_log_data = [] 
        self.last_summary_report = [] 

        self.THAI_MONTHS = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
            5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
            9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }
        self.MONTH_TO_INT = {v: k for k, v in self.THAI_MONTHS.items()}

        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        self._build_processing_tab(main_frame)


    def _build_processing_tab(self, parent_tab):
        """สร้าง UI สำหรับแท็บ 'ประมวลผลเวลา'"""
        
        upload_frame = ttk.LabelFrame(parent_tab, text="  ขั้นตอนที่ 1: อัปโหลด Log จากเครื่องสแกน  ", padding=15)
        upload_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Button(upload_frame, text="เลือกไฟล์ Excel (.xlsx, .csv)", 
                   command=self._import_fingerprint_file, width=30).pack(side="left", padx=10)
        self.save_to_db_btn = ttk.Button(upload_frame, text="💾 บันทึก Log ลงฐานข้อมูล", 
                                 command=self._save_logs_to_db, state="disabled")
        self.save_to_db_btn.pack(side="left", padx=10)
        self.upload_status_label = ttk.Label(upload_frame, text="ยังไม่ได้เลือกไฟล์", font=("Segoe UI", 10, "italic"), foreground="gray")
        self.upload_status_label.pack(side="left")

        process_frame = ttk.LabelFrame(parent_tab, text="  ขั้นตอนที่ 2: ประมวลผล ขาด/สาย  ", padding=15)
        process_frame.pack(fill="x", pady=15)
        
        process_frame.columnconfigure(1, weight=1)
        process_frame.columnconfigure(3, weight=1)
        process_frame.columnconfigure(4, weight=1)

        filter_frame = ttk.LabelFrame(process_frame, text="  ตัวกรองด่วน  ", padding=10)
        filter_frame.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(0, 15), padx=5)

        ttk.Label(filter_frame, text="ปี (พ.ศ.):").pack(side="left", padx=(5,5))
        current_year_be = datetime.now().year + 543
        year_values = [str(y) for y in range(current_year_be + 1, current_year_be - 5, -1)]
        self.year_combo = ttk.Combobox(filter_frame, values=year_values, width=8, state="readonly", font=("Segoe UI", 10))
        self.year_combo.set(str(current_year_be))
        self.year_combo.pack(side="left", padx=5)
        ttk.Label(filter_frame, text="เดือน:").pack(side="left", padx=5)
        month_values = list(self.THAI_MONTHS.values())
        self.month_combo = ttk.Combobox(filter_frame, values=month_values, width=15, state="readonly", font=("Segoe UI", 10))
        self.month_combo.set(self.THAI_MONTHS[datetime.now().month])
        self.month_combo.pack(side="left", padx=5)
        btn_frame = ttk.Frame(filter_frame)
        btn_frame.pack(side="left", padx=10)
        ttk.Button(btn_frame, text="1-15", command=self._set_date_1_15, width=8).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="16-สิ้นเดือน", command=self._set_date_16_end, width=10).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="ทั้งเดือน", command=self._set_date_month, width=8).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="ทั้งปี", command=self._set_date_year, width=8).pack(side="left", padx=2)

        ttk.Label(process_frame, text="ประมวลผลตั้งแต่วันที่:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="e", padx=5, pady=10)
        
        self.start_date_entry = DateDropdown(process_frame, font=("Segoe UI", 10))
        self.start_date_entry.grid(row=1, column=1, sticky="w", pady=10)

        ttk.Label(process_frame, text="ถึงวันที่:", font=("Segoe UI", 10)).grid(row=1, column=2, sticky="e", padx=5, pady=10)
        
        self.end_date_entry = DateDropdown(process_frame, font=("Segoe UI", 10))
        self.end_date_entry.grid(row=1, column=3, sticky="w", pady=10)
        
        self.process_btn = ttk.Button(process_frame, text="🚀 เริ่มประมวลผล", 
                                      command=self._run_processing, style="Success.TButton", 
                                      state="normal") 
        self.process_btn.grid(row=1, column=4, sticky="ns", padx=(20, 10), pady=10)

        self.export_btn = ttk.Button(process_frame, text="📄 Export Excel", 
                                     command=self._export_summary_to_excel, state="disabled") 
        self.export_btn.grid(row=1, column=5, sticky="ns", padx=(0, 10), pady=10)
        
        self.result_frame = ttk.LabelFrame(parent_tab, text="  ผลลัพธ์การประมวลผล (สรุปการสาย)  ", padding=15)
        self.result_frame.pack(fill="both", expand=True, pady=(0, 10))

        tree_container = ttk.Frame(self.result_frame)
        tree_container.pack(fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        self.result_tree = ttk.Treeview(
            tree_container,
            columns=("id", "name", "emp_type", "department", "position", "late_min", "late_hr", "absent"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            height=15 
        )
        self.result_tree.heading("id", text="รหัสพนักงาน")
        self.result_tree.heading("name", text="ชื่อ-นามสกุล")
        self.result_tree.heading("emp_type", text="ประเภทการจ้าง")
        self.result_tree.heading("department", text="แผนก")
        self.result_tree.heading("position", text="ตำแหน่ง")
        self.result_tree.heading("late_min", text="สาย (นาที)")
        self.result_tree.heading("late_hr", text="สาย (ชม.)")
        self.result_tree.heading("absent", text="ขาด (วัน)")
        
        self.result_tree.column("id", width=100, anchor="center")
        self.result_tree.column("name", width=250, anchor="w")
        self.result_tree.column("emp_type", width=120, anchor="w")
        self.result_tree.column("department", width=120, anchor="w")
        self.result_tree.column("position", width=150, anchor="w")
        self.result_tree.column("late_min", width=100, anchor="e")
        self.result_tree.column("late_hr", width=100, anchor="e")
        self.result_tree.column("absent", width=100, anchor="center")
        
        self.result_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.result_tree.yview)
        scrollbar_x.config(command=self.result_tree.xview)
        self.result_tree.tag_configure('striped', background='#f0f0f0')

        self.result_tree.bind("<Double-1>", self._show_attendance_details)
    
    def _import_fingerprint_file(self):
        """ฟังก์ชันหลักสำหรับเลือกไฟล์และจัดการโหมดการนำเข้า (Mass Edit)"""
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์สแกนนิ้ว (Excel/CSV)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
        )
        if not file_path: return

        # 1. โหลดข้อมูลเข้า Memory
        self._load_file(file_path)
        
        if not self.raw_log_data: 
            return # ถ้าโหลดไม่ผ่าน จบการทำงาน

        try:
            # หาช่วงวันที่จากข้อมูลในไฟล์
            all_timestamps = [item[1] for item in self.raw_log_data]
            min_date = min(all_timestamps).date()
            max_date = max(all_timestamps).date()
            
            # 2. ถาม User: ต้องการโหมดไหน?
            msg = (f"ข้อมูลในไฟล์: {min_date.strftime('%d/%m/%Y')} ถึง {max_date.strftime('%d/%m/%Y')}\n"
                   f"จำนวน: {len(self.raw_log_data)} รายการ\n\n"
                   "คุณต้องการทำรายการแบบไหน?\n"
                   "✅ [YES] = ลบข้อมูลเก่าในระบบทิ้ง แล้วลงข้อมูลใหม่ (Mass Edit/Replace)\n"
                   "❌ [NO]  = เพิ่มข้อมูลใหม่เข้าไปต่อท้าย (Append)")
            
            is_replace = messagebox.askyesno("ยืนยันการนำเข้า", msg)
            
            if is_replace:
                hr_database.delete_scan_logs_range(min_date, max_date)
                
            # 3. บันทึกข้อมูลลง DB
            count = hr_database.insert_scan_logs(self.raw_log_data)
            
            messagebox.showinfo("สำเร็จ", f"นำเข้าข้อมูลเรียบร้อย {count} รายการ")
            
            # 4. สั่งประมวลผลทันที
            self._save_logs_to_db()

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")
    
    def _load_file(self, file_path):
        """
        โหลดไฟล์ Log (Excel/CSV) และแปลงข้อมูลให้พร้อมบันทึก
        - รองรับ Format วันที่หลากหลาย
        - จัดการ Error และแจ้งเตือนผู้ใช้
        """
        # (ถ้า file_path ไม่ถูกส่งมา ให้เปิด Dialog ถาม)
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="เลือกไฟล์ Log (Excel หรือ CSV)",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
        
        if not file_path: return

        try:
            self.upload_status_label.config(text="⏳ กำลังโหลดไฟล์...", foreground="orange")
            self.update_idletasks()
            
            # 1. อ่านไฟล์ตามนามสกุล
            if file_path.lower().endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(file_path, encoding='tis-620')
            else:
                df = pd.read_excel(file_path)

            total_rows = len(df)
            self.upload_status_label.config(
                text=f"⏳ อ่านไฟล์สำเร็จ ({total_rows} แถว) กำลังตรวจสอบคอลัมน์...", 
                foreground="orange"
            )
            self.update_idletasks()

            # 2. Map ชื่อคอลัมน์ (รองรับชื่อภาษาไทย/อังกฤษ)
            column_mapping = {
                'รหัสพนักงาน': ['รหัสพนักงาน', 'ID', 'รหัส', 'รับ', 'EmpID', 'User ID', 'AC-No.'], 
                'วันที่': ['วันที่', 'Date', 'วัน', 'CheckDate'], 
                'เวลาเข้างาน': ['เวลาเข้างาน', 'เวลาเข้า', 'เวลาช้างาน', 'เวลาเช้างาน', 'CheckIn', 'InTime'], 
                'เวลาออกงาน': ['เวลาออกงาน', 'เวลาออก', 'เวลาออกงาม', 'CheckOut', 'OutTime'] 
            }

            actual_columns = {}
            for required_col, possible_names in column_mapping.items():
                found = False
                for col in df.columns:
                    col_clean = str(col).strip()
                    if col_clean in possible_names:
                        actual_columns[required_col] = col_clean
                        found = True
                        break
                if not found:
                    self.upload_status_label.config(text="❌ เกิดข้อผิดพลาด", foreground="red")
                    messagebox.showerror("Format Error", 
                                       f"ไม่พบคอลัมน์: {required_col}\n\n"
                                       f"ไฟล์ของคุณมีคอลัมน์: {list(df.columns)}\n\n"
                                       f"กรุณาตรวจสอบชื่อคอลัมน์ในไฟล์")
                    return

            df = df.rename(columns={v: k for k, v in actual_columns.items()})

            self.upload_status_label.config(
                text=f"⏳ กำลังประมวลผล {total_rows} แถว...", 
                foreground="orange"
            )
            self.update_idletasks()

            self.raw_log_data = []
            skipped_rows = 0
            processed_count = 0
            
            BAD_TIME_VALUES = {"", "0", "0:00", "nan", "nat", "none", "-"} 

            # กำหนด Format วันที่ที่เป็นไปได้ (เรียงตามลำดับความน่าจะเป็น)
            # หมายเหตุ: %d=วัน, %m=เดือน, %Y=ปี ค.ศ. 4 หลัก
            DATE_FORMATS = [
                '%d/%m/%Y %H:%M:%S', # 15/01/2026 08:30:00 (แบบละเอียด)
                '%d/%m/%Y %H:%M',    # 15/01/2026 08:30    (แบบปกติ)
                '%Y-%m-%d %H:%M:%S', # 2026-01-15 08:30:00 (แบบ Database)
                '%Y-%m-%d %H:%M',    # 2026-01-15 08:30
                '%m/%d/%Y %H:%M:%S', # 01/15/2026 08:30:00 (แบบ US)
                '%m/%d/%Y %H:%M',    # 01/15/2026 08:30
            ]

            # 3. วนลูปแปลงข้อมูล
            for idx, row in df.iterrows():
                processed_count += 1
                if processed_count % 100 == 0:
                    self.upload_status_label.config(
                        text=f"⏳ ประมวลผล {processed_count}/{total_rows} แถว...", 
                        foreground="orange"
                    )
                    self.update_idletasks()
                
                try:
                    emp_id = str(row['รหัสพนักงาน']).strip()
                    date_str = str(row['วันที่']).strip()
                    # ถ้า date_str มีเวลาติดมาด้วย ให้ตัดออกเอาแค่วันที่ (กรณี Excel อ่านมาเป็น datetime)
                    if " " in date_str:
                        date_str = date_str.split(" ")[0]
                except Exception:
                    continue 

                if not emp_id or not date_str or emp_id.lower() == 'nan' or date_str.lower() == 'nan':
                    continue 

                # ฟังก์ชันย่อยสำหรับแปลงเวลา
                def parse_time(time_val):
                    time_s = str(time_val).strip()
                    if time_s.lower() in BAD_TIME_VALUES: return None
                    
                    # รวมร่าง วันที่ + เวลา
                    dt_str = f"{date_str} {time_s}".replace(".", ":") # แก้ 08.30 -> 08:30
                    
                    parsed_dt = None
                    # ลองแปลงตาม Format ที่กำหนด
                    for fmt in DATE_FORMATS:
                        try:
                            parsed_dt = pd.to_datetime(dt_str, format=fmt)
                            break
                        except: continue
                    
                    # ถ้ายังไม่ได้ ให้ pandas เดา (Fallback)
                    if parsed_dt is None:
                        try:
                            # ระบุ dayfirst=True เพื่อบอกว่า 15/01 คือ 15 ม.ค.
                            parsed_dt = pd.to_datetime(dt_str, dayfirst=True, errors='coerce')
                        except: pass
                    
                    if parsed_dt is not None and not pd.isna(parsed_dt):
                        return parsed_dt
                    return None

                # แปลงเวลาเข้า
                ts_in = parse_time(row['เวลาเข้างาน'])
                if ts_in:
                    self.raw_log_data.append((emp_id, ts_in))
                
                # แปลงเวลาออก
                ts_out = parse_time(row['เวลาออกงาน'])
                if ts_out:
                    self.raw_log_data.append((emp_id, ts_out))

                # ถ้าทั้งเข้าและออกไม่ได้เลย ให้นับว่าข้าม
                if not ts_in and not ts_out:
                    skipped_rows += 1

            # 4. สรุปผล
            self.upload_status_label.config(
                text=f"✅ ไฟล์: {os.path.basename(file_path)} (พบ {len(self.raw_log_data)} Log สแกน)", 
                foreground="green"
            )
            
            if skipped_rows > 0:
                print(f"Info: ข้าม {skipped_rows} รายการที่เวลาไม่สมบูรณ์")

            self.save_to_db_btn.config(state="normal")
            self.process_btn.config(state="normal") 

        except Exception as e:
            messagebox.showerror("File Error", f"ไม่สามารถอ่านไฟล์ได้:\n{e}")
            self.upload_status_label.config(text="❌ เกิดข้อผิดพลาดในการอ่านไฟล์", foreground="red")
            import traceback; traceback.print_exc()
    
    def _save_logs_to_db(self):
        if not self.raw_log_data:
            messagebox.showwarning("ไม่มีข้อมูล", "กรุณาโหลดไฟล์ Excel ก่อน")
            return
        if not messagebox.askyesno("ยืนยัน", f"ต้องการบันทึก Log ดิบ {len(self.raw_log_data)} รายการลงฐานข้อมูลหรือไม่?"):
            return
        try:
            success_count = hr_database.insert_scan_logs(self.raw_log_data)
            messagebox.showinfo("สำเร็จ", f"บันทึก Log ดิบ {success_count} รายการเรียบร้อยแล้ว\n(ข้อมูลที่ซ้ำจะถูกข้าม)")
            self.raw_log_data = []
            self.save_to_db_btn.config(state="disabled")
            self.upload_status_label.config(text="บันทึกลง DB แล้ว (กรุณาเลือกไฟล์ใหม่)", foreground="blue")
        except Exception as e:
            messagebox.showerror("Database Error", f"ไม่สามารถบันทึก Log ได้:\n{e}")

    def _run_processing(self):
        """เริ่มประมวลผลและแสดงลงตารางทันที (แก้ไขแล้ว)"""
        try:
            # 1. ดึงและเช็ควันที่
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            
            if not start_date or not end_date: 
                messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกช่วงวันที่ให้ถูกต้อง")
                return
            
            if start_date > end_date:
                messagebox.showwarning("วันที่ผิดพลาด", "วันที่เริ่มต้น ต้องมาก่อนวันที่สิ้นสุด")
                return
            
            # 2. เริ่มคำนวณ (เปลี่ยนเคอร์เซอร์เป็นนาฬิกาทราย)
            self.config(cursor="wait")
            self.update()
            
            # เรียกฟังก์ชันคำนวณหลักจาก Database
            summary_report = hr_database.process_attendance_summary(start_date, end_date)

            # คืนค่าเคอร์เซอร์
            self.config(cursor="") 

            # 3. อัปเดตข้อมูลลงตัวแปร
            self.last_summary_report = summary_report 
            self.export_btn.config(state="normal")  
            
            # 4. เคลียร์ตารางเก่า
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)

            self.result_frame.config(text=f"  ผลลัพธ์การประมวลผล ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})  ")
            
            # 5. วนลูปเติมข้อมูลใหม่ลงตาราง (แทนที่ _display_summary)
            if summary_report:
                for i, report in enumerate(summary_report):
                    tag_to_use = 'striped' if i % 2 == 0 else '' 
                    
                    self.result_tree.insert("", "end", iid=report['emp_id'], values=(
                        report['emp_id'],
                        report['name'],
                        report.get('emp_type', ''),     
                        report.get('department', ''), 
                        report.get('position', ''),   
                        f"{report['total_late_minutes']:.0f}", 
                        f"{report['total_late_hours']:.2f}",  
                        report['absent_days']
                    ), tags=(tag_to_use,))
                
                messagebox.showinfo("สำเร็จ", f"ประมวลผลเรียบร้อย {len(summary_report)} คน")
            else:
                self.result_frame.config(text=f"  ผลลัพธ์: ไม่พบข้อมูล ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})  ")
                messagebox.showinfo("แจ้งเตือน", "ไม่พบข้อมูลพนักงาน หรือไม่มีการสแกนนิ้วในช่วงเวลานี้")

        except Exception as e:
            self.config(cursor="")
            self.export_btn.config(state="disabled") 
            import traceback; traceback.print_exc()
            messagebox.showerror("Processing Error", f"เกิดข้อผิดพลาดขณะประมวลผล:\n{e}")
            
    def _export_summary_to_excel(self):
        if not self.last_summary_report:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลสรุปผลลัพธ์ที่จะ Export")
            return
            
        start_str = self.start_date_entry.get().replace("/", "-")
        end_str = self.end_date_entry.get().replace("/", "-")
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="บันทึกสรุป ขาด/สาย เป็น Excel",
            initialfile=f"สรุปขาดสาย_{start_str}_ถึง_{end_str}.xlsx"
        )
        if not file_path:
            return 
        try:
            df = pd.DataFrame(self.last_summary_report)
            
            column_mapping = {
                "emp_id": "รหัสพนักงาน",
                "name": "ชื่อ-นามสกุล",
                "emp_type": "ประเภทการจ้าง",  
                "department": "แผนก",      
                "position": "ตำแหน่ง",      
                "total_late_minutes": "สาย (นาที)",
                "total_late_hours": "สาย (ชั่วโมง)",
                "absent_days": "ขาดงาน (วัน)"
            }
            df = df.rename(columns=column_mapping)
            
            final_columns = [
                "รหัสพนักงาน", "ชื่อ-นามสกุล", "ประเภทการจ้าง", "แผนก", "ตำแหน่ง",
                "สาย (นาที)", "สาย (ชั่วโมง)", "ขาดงาน (วัน)"
            ]
            
            df = df[[col for col in final_columns if col in df.columns]]
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("✅ สำเร็จ", f"บันทึกไฟล์ Excel เรียบร้อยแล้ว\nที่: {file_path}")
        except Exception as e:
            messagebox.showerror("❌ เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้:\n{e}")

    def _get_selected_dates(self):
        try:
            year_be = int(self.year_combo.get())
            year_ce = year_be - 543
            month_name = self.month_combo.get()
            month_int = self.MONTH_TO_INT[month_name]
            return year_ce, month_int
        except Exception as e:
            messagebox.showerror("ข้อมูลไม่ครบ", f"กรุณาเลือกปีและเดือนให้ถูกต้อง: {e}")
            return None, None

    def _update_date_entries(self, start_date, end_date):
        self.start_date_entry.set_date(start_date)
        self.end_date_entry.set_date(end_date)
        
    def _set_date_1_15(self):
        year_ce, month_int = self._get_selected_dates()
        if not year_ce: return
        start_date = datetime(year_ce, month_int, 1)
        end_date = datetime(year_ce, month_int, 15)
        self._update_date_entries(start_date, end_date)

    def _set_date_16_end(self):
        year_ce, month_int = self._get_selected_dates()
        if not year_ce: return
        start_date = datetime(year_ce, month_int, 16)
        last_day = calendar.monthrange(year_ce, month_int)[1]
        end_date = datetime(year_ce, month_int, last_day)
        self._update_date_entries(start_date, end_date)

    def _set_date_month(self):
        year_ce, month_int = self._get_selected_dates()
        if not year_ce: return
        start_date = datetime(year_ce, month_int, 1)
        last_day = calendar.monthrange(year_ce, month_int)[1]
        end_date = datetime(year_ce, month_int, last_day)
        self._update_date_entries(start_date, end_date)
        
    def _set_date_year(self):
        try:
            year_be = int(self.year_combo.get())
            year_ce = year_be - 543
        except Exception:
            messagebox.showerror("ข้อมูลไม่ครบ", "กรุณาเลือกปีให้ถูกต้อง")
            return
        start_date = datetime(year_ce, 1, 1)
        end_date = datetime(year_ce, 12, 31)
        self._update_date_entries(start_date, end_date)

    def _show_attendance_details(self, event):
        """
        (ฉบับแก้ไข V76.1 - Fix Penalty Display Bug & Update Half-day to 4.5 hrs)
        - แก้ไขปัญหา: ช่อง 'สาย(นาที)' และ 'หัก(ชม.)' ไม่แสดงค่า
        - สาเหตุ: Date Key ไม่ตรงกัน (Object vs Thai String) -> แก้ให้แปลงเป็น String พ.ศ. ก่อน Map
        - แก้ไขเวลาลาครึ่งวัน (0.5) ให้แสดงผลเป็น 4.5 ชม.
        """
        from datetime import date, datetime # Import เพิ่มกันเหนียว

        selection = self.result_tree.selection()
        if not selection: return
        
        emp_id = selection[0] 
        
        emp_data = None
        for report in self.last_summary_report:
            if report['emp_id'] == emp_id:
                emp_data = report
                break
        
        if not emp_data:
            messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบข้อมูลรายละเอียดสำหรับพนักงานนี้")
            return

        emp_name = emp_data.get('name', emp_id)
        emp_type = emp_data.get('emp_type', '')
        
        # ข้อมูลที่คำนวณเสร็จแล้ว (มี Late/Penalty ครบ)
        details_list_calculated = emp_data.get('details', [])
        
        # --- [FIX START] สร้าง Map โดยแปลง Date Object เป็น String (dd/mm/yyyy พ.ศ.) ---
        calculated_map = {}
        for row in details_list_calculated:
            d = row.get('date')
            if d:
                key_str = ""
                # ถ้าเป็น Date Object (ปกติจะเป็นตัวนี้)
                if hasattr(d, 'day'):
                    key_str = f"{d.day:02d}/{d.month:02d}/{d.year + 543}"
                else:
                    # เผื่อเป็น String
                    key_str = str(d)
                
                calculated_map[key_str] = row
        # --- [FIX END] ---

        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
        except: return

        # ดึงข้อมูลสดจาก DB
        daily_records = hr_database.get_daily_records_range(emp_id, start_date, end_date)

        is_daily_emp = "รายวัน" in str(emp_type) or "Daily" in str(emp_type)
        
        win = tk.Toplevel(self)
        win.title(f"รายละเอียดปฏิทินทำงาน - {emp_name} ({emp_type})")
        win.geometry("1550x750")
        win.transient(self) 
        win.grab_set()      
        
        sheet_frame = ttk.Frame(win, padding=(15, 15, 15, 0))
        sheet_frame.pack(fill="both", expand=True)

        sheet = Sheet(sheet_frame, 
                      show_toolbar=False, show_top_left=False, show_row_index=False,
                      show_header=True, expand="both")
        sheet.pack(fill="both", expand=True) 

        # --- Headers ---
        if is_daily_emp:
            headers = ["วันที่", "สถานะการทำงาน", "เวลาเข้า", "เวลาออก", 
                       "ชม.ทำงาน", "ชม.ลา", "สาย(นาที)", "หัก(ชม.)",
                       "เริ่ม OT", "ออก OT", "ชม.OT", "อนุมัติ OT", "หมายเหตุ"] # 🛠️ เพิ่ม "หมายเหตุ"
        else:
            headers = ["วันที่", "สถานะการทำงาน", "เวลาเข้า", "เวลาออก", 
                       "ชม.ทำงาน", "ชม.ลา", "สาย(นาที)", "หัก(ชม.)", "หมายเหตุ"] # 🛠️ เพิ่ม "หมายเหตุ"
            
        sheet.headers(headers)
        
        # Config Widths
        sheet.column_width(column=0, width=90)
        sheet.column_width(column=1, width=280)
        sheet.column_width(column=2, width=70); sheet.column_width(column=3, width=70)
        sheet.column_width(column=4, width=80); sheet.column_width(column=5, width=60)
        sheet.column_width(column=6, width=70); sheet.column_width(column=7, width=60)
        
        if is_daily_emp:
            sheet.column_width(column=8, width=70); sheet.column_width(column=9, width=70)
            sheet.column_width(column=10, width=70); sheet.column_width(column=11, width=100)
        
        # --- Prepare Data ---
        sheet_data = []
        is_diligence_failed = False
        fail_reasons = [] 

        if not daily_records:
            empty_row = [""] * len(headers)
            empty_row[1] = "(ไม่พบข้อมูล)"
            sheet_data.append(empty_row)
        else:
            for item in daily_records:
                # สร้าง Key String ให้ตรงกับ Map ที่ทำไว้ด้านบน
                d_str = item['work_date'].strftime("%d/%m/%Y")
                d_str_thai = f"{item['work_date'].day:02d}/{item['work_date'].month:02d}/{item['work_date'].year + 543}"
                
                status_text = item.get('status', 'ปกติ')
                scan_in = item.get('work_in_time') or "-"
                scan_out = item.get('work_out_time') or "-"
                
                # [LOOKUP CORRECTLY NOW]
                calc_row = calculated_map.get(d_str_thai, {})
                
                # --- Diligence Logic ---
                if "สาย" in status_text or "ขาด" in status_text or "ลา" in status_text or "ไม่ครบ" in status_text:
                    is_diligence_failed = True
                    reason = status_text.split('(')[0].strip()
                    if "สาย" in status_text: reason = "มาสาย"
                    elif "ไม่ครบ" in status_text: reason = "กลับก่อน"
                    fail_msg = f"{item['work_date'].day}: {reason}"
                    if fail_msg not in fail_reasons: fail_reasons.append(fail_msg)

                if scan_in and scan_in != "-":
                    try:
                        time_in_obj = datetime.strptime(scan_in[:5], "%H:%M").time() 
                        limit_time = datetime.strptime("08:00", "%H:%M").time()
                        if time_in_obj > limit_time:
                            is_diligence_failed = True
                            reason = f"{item['work_date'].day}: สาย({scan_in[:5]})"
                            is_duplicate = False
                            for r in fail_reasons:
                                if f"{item['work_date'].day}:" in r and "สาย" in r: is_duplicate = True
                            if not is_duplicate: fail_reasons.append(reason)
                    except: pass

                # --- Work Hours ---
                work_hrs_str = "-"
                if scan_in != "-" and scan_out != "-":
                    try:
                        t_in = datetime.strptime(scan_in[:5], "%H:%M")
                        t_out = datetime.strptime(scan_out[:5], "%H:%M")
                        diff = t_out - t_in
                        total_seconds = diff.total_seconds()
                        noon_start = t_in.replace(hour=12, minute=0)
                        noon_end = t_in.replace(hour=13, minute=0)
                        if t_in < noon_end and t_out > noon_start:
                             overlap_start = max(t_in, noon_start)
                             overlap_end = min(t_out, noon_end)
                             if (overlap_end - overlap_start).total_seconds() > 0:
                                 total_seconds -= (overlap_end - overlap_start).total_seconds()
                        hours = int(total_seconds // 3600)
                        minutes = int((total_seconds % 3600) // 60)
                        work_hrs_str = f"{hours}ชม. {minutes}น."
                    except: pass

                # --- [แก้ไขจุดนี้ให้เป็น 4.5 ชม.] ---
                leave_hrs_str = ""
                if "ลา" in status_text and "(" in status_text:
                     if "0.5" in status_text: leave_hrs_str = "4 ชม."  # <--- จุดที่ต้องแก้
                     elif "1.0" in status_text: leave_hrs_str = "8 ชม."

                # --- Penalty Columns ---
                
                # 1. สาย (นาที)
                actual_late_val = calc_row.get('actual_late_mins', 0)
                # ถ้าลาแล้ว (ที่ไม่ใช่ลาแล้วมาสาย) ไม่โชว์นาทีสาย
                if status_text.startswith("ลา") and "มาสาย" not in status_text:
                    actual_late_val = 0
                actual_late_str = f"{actual_late_val}" if actual_late_val > 0 else ""
                
                # 2. หัก (ชม.)
                penalty_val = calc_row.get('penalty_hrs', 0)
                # ถ้าลาป่วย/ลากิจ ไม่โชว์ยอดหักเงิน (เพราะหักวันลา) แต่ถ้า ลาไม่รับค่าจ้าง/สาย/ออกก่อน ต้องโชว์
                if status_text.startswith("ลา") and "ลาไม่รับค่าจ้าง" not in status_text:
                     penalty_val = 0
                
                # กรณีพิเศษ: ดึงยอดหักจาก Text (เช่น "ลาไม่รับค่าจ้าง (หัก 2 ชม.)") ถ้าใน calc_row ไม่มี
                if penalty_val == 0 and "หัก" in status_text and "ชม." in status_text:
                    import re
                    match = re.search(r"หัก\s*([\d\.]+)\s*ชม", status_text)
                    if match:
                        try: penalty_val = float(match.group(1))
                        except: pass

                penalty_str = f"{penalty_val:.2f}" if penalty_val > 0 else ""

                row_vals = [
                    d_str, 
                    status_text, 
                    scan_in, scan_out,
                    work_hrs_str, leave_hrs_str,
                    actual_late_str, penalty_str 
                ]
                
                if is_daily_emp:
                    ot_hrs = float(item.get('ot_hours', 0))
                    ot_hrs_str = f"{ot_hrs:.2f}" if ot_hrs > 0 else ""
                    
                    display_ot_in = item.get('ot_in_time') or ""
                    display_ot_out = item.get('ot_out_time') or ""
                    if ot_hrs > 0 and not display_ot_in: display_ot_in = scan_in
                    if ot_hrs > 0 and not display_ot_out: display_ot_out = scan_out
                    
                    is_approved = item.get('is_ot_approved', False)
                    approval_text = ""
                    if ot_hrs > 0:
                        approval_text = "✅ อนุมัติ" if is_approved else "❌ ไม่อนุมัติ"
                    
                    row_vals.extend([display_ot_in, display_ot_out, ot_hrs_str, approval_text])
                
                remark_text = item.get('remark') or ""
                row_vals.append(remark_text)
                sheet_data.append(row_vals)
        
        sheet.set_sheet_data(sheet_data)
        
        # --- Highlight & Readonly ---
        for i, row_vals in enumerate(sheet_data):
            row_status = str(row_vals[1])
            t_in_chk = str(row_vals[2]) 
            
            bg, fg = "#ffffff", "#000000"
            if 'ขาดงาน' in row_status: bg, fg = '#fddfe2', '#9f1f2e'
            elif 'สาย' in row_status: bg, fg = '#fff4de', '#a05f00' 
            elif 'ออกก่อน' in row_status: bg, fg = '#fff4de', '#a05f00'
            elif 'ไม่ครบ' in row_status: bg, fg = '#fff4de', '#a05f00'
            elif 'ลา' in row_status: bg, fg = '#e0f0ff', '#00529e'
            elif 'วันหยุด' in row_status: bg, fg = '#ffffff', 'gray'
            elif is_daily_emp and len(row_vals) > 10 and row_vals[10] != "": bg = '#f0fff0' 
            elif i % 2 == 1: bg = '#f0f0f0'
            
            # Highlight เวลาสาย (08:00)
            if t_in_chk != "-" and t_in_chk:
                try:
                    if datetime.strptime(t_in_chk[:5], "%H:%M").time() > datetime.strptime("08:00", "%H:%M").time():
                        fg = "#FF0000" 
                except: pass

            sheet.highlight_rows(rows=[i], bg=bg, fg=fg)
            
        if is_daily_emp:
            sheet.readonly_columns(columns=[0, 4, 5, 6, 7, 10]) 
        else:
            sheet.readonly_columns(columns=[0, 4, 5, 6, 7]) 
        
        # --- Dropdowns ---
        leave_types = ["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาคลอด", "ลาบวช", "ลาอื่นๆ", "ลาไม่รับค่าจ้าง"]
        status_options_base = ["ปกติ", "ขาดงาน", "มาสาย", "ออกก่อนเวลา/ชม.ไม่ครบ"]
        for lt in leave_types:
            status_options_base.append(f"{lt} (เต็มวัน)")
            status_options_base.append(f"{lt} (ครึ่งวัน)")
            status_options_base.append(f"{lt} (ระบุเวลา)") 

        approval_options = ["✅ อนุมัติ", "❌ ไม่อนุมัติ"]

        total_rows = sheet.get_total_rows()
        for i in range(total_rows):
            if i >= len(sheet_data): continue
            
            curr_stat = str(sheet_data[i][1])
            sheet.create_dropdown(r=i, c=1, values=status_options_base, set_value=curr_stat, state="readonly")
            
            if is_daily_emp:
                ot_val_str = str(sheet_data[i][10])
                if ot_val_str:
                    curr_appr = sheet_data[i][11]
                    if not curr_appr: curr_appr = "❌ ไม่อนุมัติ"
                    sheet.create_dropdown(r=i, c=11, values=approval_options, set_value=curr_appr, state="readonly")

        sheet.enable_bindings("single", "drag_select", "row_select", "column_width_resize", "arrowkeys", "edit_cell")

        # --- Diligence Summary (Copy ส่วนเดิมมาวางต่อได้เลย) ---
        diligence_frame = ttk.LabelFrame(win, text="  🏆 สรุปเบี้ยขยัน (Diligence Allowance)  ", padding=10)
        diligence_frame.pack(fill="x", padx=15, pady=5)
        
        if not is_daily_emp:
            ttk.Label(diligence_frame, text="* พนักงานรายเดือน ไม่ได้รับสิทธิ์เบี้ยขยัน", foreground="gray").pack(anchor="w")
        else:
            current_streak = hr_database.get_employee_diligence_streak(emp_id)
            diligence_amount = 0
            step_msg = ""
            status_text = ""
            status_color = ""
            
            if not is_diligence_failed:
                if current_streak == 0: diligence_amount = 300; step_msg = "เริ่มต้น (เดือนที่ 1)"
                elif current_streak == 1: diligence_amount = 400; step_msg = "ต่อเนื่อง (เดือนที่ 2)"
                else: diligence_amount = 500; step_msg = f"สูงสุด (ต่อเนื่องเดือนที่ {current_streak + 1})"
                status_text = "✅ ผ่านเกณฑ์ (ไม่ขาด/ลา/สาย/เข้าทัน 8 โมง)"; status_color = "green"
            else:
                reason_str = ", ".join(fail_reasons[:3]) 
                if len(fail_reasons) > 3: reason_str += "..."
                status_text = f"❌ ไม่ผ่านเกณฑ์ ({reason_str})"; status_color = "red"
                step_msg = "เดือนหน้าเริ่มนับใหม่ที่ 300"

            row1 = ttk.Frame(diligence_frame); row1.pack(fill="x")
            ttk.Label(row1, text="สถานะเดือนนี้:", width=15, font=("", 10, "bold")).pack(side="left")
            ttk.Label(row1, text=status_text, foreground=status_color, font=("", 10, "bold")).pack(side="left")
            
            row2 = ttk.Frame(diligence_frame); row2.pack(fill="x")
            ttk.Label(row2, text="ยอดเงินที่ได้:", width=15).pack(side="left")
            ttk.Label(row2, text=f"{diligence_amount:,.2f} บาท", font=("", 11, "bold"), foreground="blue").pack(side="left")
            ttk.Label(row2, text=f"  ({step_msg})").pack(side="left")
            
            row3 = ttk.Frame(diligence_frame); row3.pack(fill="x")
            ttk.Label(row3, text="สถิติเดิม:", width=15).pack(side="left")
            ttk.Label(row3, text=f"ทำต่อเนื่องมาแล้ว {current_streak} เดือน").pack(side="left")

        # Footer Buttons
        btn_frame = ttk.Frame(win, padding=10)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="📄 Export Excel", command=lambda: self._export_details_to_excel(daily_records, emp_name)).pack(side="left")
        
        ttk.Button(btn_frame, text="💾 บันทึกการแก้ไข", 
                   command=lambda: self._save_details_from_popup(sheet, daily_records, emp_id, win, is_daily_emp), 
                   style="Success.TButton").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="ปิด", command=win.destroy).pack(side="right")
    
    
    def _parse_date_be(self, date_str_be):
        """(Helper) แปลง 'dd/mm/yyyy' (พ.ศ.) เป็น date object (ค.ศ.)"""
        try:
            day, month, year_be = map(int, date_str_be.split('/'))
            year_ce = year_be - 543
            return datetime(year_ce, month, day).date()
        except Exception:
            return None

    def _parse_leave_type(self, status_str):
        """(Helper) แปลง 'ลา (ลากิจ) (0.5' หรือ 'ลา (ลากิจ)' เป็น 'ลากิจ'"""
        
        # (เวอร์ชันอัปเกรด: ฉลาดขึ้น)
        
        if not status_str.startswith("ลา ("):
            return None
        
        # ตัด "ลา (" (4 ตัวอักษร) ที่จุดเริ่มต้น
        temp_str = status_str[4:] 
        
        # หาตำแหน่งของ ")" ตัวแรก
        end_paren_index = temp_str.find(")")
        
        if end_paren_index == -1:
            # ไม่เจอ ')' เลย (เช่น "ลา (ลากิจ (0.5" - รูปแบบผิด)
            return None 
        
        # ดึง "ลากิจ" (หรือ "ลาป่วย" ฯลฯ) ออกมา
        leave_type = temp_str[:end_paren_index] 
        return leave_type

    def _save_details_from_popup(self, sheet, original_details_list, emp_id, popup_window, is_daily_emp=False):
        """
        (ฉบับแก้ไข V68.0 - Keep Scan Time Logic & Full Version)
        - รองรับการบันทึกสถานะลา โดยไม่ลบเวลาเข้า-ออกทิ้ง
        - แก้ปัญหา KeyError date/work_date
        - รองรับสถานะ 'ระบุเวลา'
        """
        try:
            if not messagebox.askyesno("ยืนยันการบันทึก", "ระบบจะบันทึกสถานะการลาและคำนวณ OT ใหม่\nต้องการบันทึกใช่หรือไม่?", parent=popup_window):
                return 
            
            # 1. ดึงข้อมูลจากตาราง (Sheet) ที่แก้ไขแล้ว
            new_data_list_of_lists = sheet.get_sheet_data()
            
            # 2. กำหนด Headers
            if is_daily_emp:
                headers = [
                    "date_str", "status", "scan_in", "scan_out", 
                    "work_hrs", "leave_hours", "actual_late_mins", "penalty_hrs",
                    "ot_in", "ot_out", "ot_hrs", "ot_approved", "remark" # 🛠️ เพิ่ม "remark"
                ]
            else:
                headers = [
                    "date_str", "status", "scan_in", "scan_out", 
                    "work_hrs", "leave_hours", "actual_late_mins", "penalty_hrs", "remark" # 🛠️ เพิ่ม "remark"
                ]
            
            # สร้าง Map วันที่ -> ข้อมูลใหม่
            new_data_map = {}
            for row_list in new_data_list_of_lists:
                while len(row_list) < len(headers): row_list.append("") 
                row_dict = {headers[i]: str(row_vals).strip() for i, row_vals in enumerate(row_list)}
                if row_dict.get('date_str'): 
                    new_data_map[row_dict['date_str']] = row_dict

            changes_detected = 0
            
            # 3. วนลูปเทียบกับข้อมูลเดิม
            for original_row in original_details_list:
                
                # --- [KEY FIX] จับคู่วันที่ ---
                date_obj = None
                new_row = None
                
                if 'work_date' in original_row:
                    wd = original_row['work_date']
                    date_obj = wd
                    key_be = f"{wd.day:02d}/{wd.month:02d}/{wd.year + 543}"
                    key_ad = wd.strftime("%d/%m/%Y")
                    
                    if key_be in new_data_map: new_row = new_data_map[key_be]
                    elif key_ad in new_data_map: new_row = new_data_map[key_ad]
                else:
                    d_str = str(original_row.get('date', '')).strip()
                    if d_str in new_data_map:
                        new_row = new_data_map[d_str]
                        date_obj = self._parse_date_be(d_str)

                if not new_row or not date_obj: continue 

                # Helper
                def get_val(row, key_list):
                    for k in key_list:
                        if k in row and row[k] is not None: return str(row[k]).strip()
                    return ""

                val_status_old = str(original_row.get('status') or "").strip()
                val_in_old = get_val(original_row, ['work_in_time', 'scan_in'])
                val_out_old = get_val(original_row, ['work_out_time', 'scan_out'])
                
                # ค่าใหม่
                val_status_new = new_row['status']
                if val_status_new == "เลือกสถานะ": val_status_new = val_status_old
                val_in_new = new_row['scan_in']
                if val_in_new in ["None", "-", ""]: val_in_new = ""
                val_out_new = new_row['scan_out']
                if val_out_new in ["None", "-", ""]: val_out_new = ""
                
                # Normalize ค่าเดิม
                if val_in_old == "-": val_in_old = ""
                if val_out_old == "-": val_out_old = ""

                # เช็คการเปลี่ยนแปลง
                status_changed = (val_status_old != val_status_new)
                scan_in_changed = (val_in_old != val_in_new)
                scan_out_changed = (val_out_old != val_out_new)

                val_remark_old = str(original_row.get('remark') or "").strip()
                val_remark_new = new_row.get('remark', "").strip()
                remark_changed = (val_remark_old != val_remark_new)

                ot_changed = False
                new_calculated_ot_hours = 0.0
                val_ot_in_new = ""
                val_ot_out_new = ""
                val_approved_new = False

                if is_daily_emp:
                    val_ot_in_new = new_row.get('ot_in', '')
                    if val_ot_in_new in ["None", "-", ""]: val_ot_in_new = ""
                    val_ot_out_new = new_row.get('ot_out', '')
                    if val_ot_out_new in ["None", "-", ""]: val_ot_out_new = ""
                    
                    val_ot_in_old = get_val(original_row, ['ot_in_time', 'ot_in'])
                    if val_ot_in_old == "-": val_ot_in_old = ""
                    val_ot_out_old = get_val(original_row, ['ot_out_time', 'ot_out'])
                    if val_ot_out_old == "-": val_ot_out_old = ""
                    
                    val_appr_str = new_row.get('ot_approved', '')
                    val_approved_new = (val_appr_str == "✅ อนุมัติ")
                    val_approved_old = bool(original_row.get('is_ot_approved', False))
                    
                    if val_ot_in_new and val_ot_out_new:
                        new_calculated_ot_hours = self._calculate_time_diff(val_ot_in_new, val_ot_out_new)
                    
                    old_ot_hours = float(original_row.get('ot_hours', 0) or 0)
                    
                    if (val_ot_in_new != val_ot_in_old) or \
                       (val_ot_out_new != val_ot_out_old) or \
                       (abs(new_calculated_ot_hours - old_ot_hours) > 0.01) or \
                       (val_approved_new != val_approved_old):
                        ot_changed = True

                if not (status_changed or scan_in_changed or scan_out_changed or ot_changed or remark_changed):
                    continue
                
                changes_detected += 1
                
                # --- เริ่มบันทึก ---
                def is_leave_status(s):
                    return "ลา" in s and ("(" in s or "เต็มวัน" in s or "ครึ่งวัน" in s or "ระบุเวลา" in s)

                if status_changed:
                    # 1. ลบใบลาเก่าทิ้ง (ถ้ามี)
                    if is_leave_status(val_status_old) and not is_leave_status(val_status_new):
                        hr_database.delete_leave_record_on_date(emp_id, date_obj)
                    
                    # 2. เพิ่มใบลาใหม่ (ถ้าเลือก)
                    elif is_leave_status(val_status_new):
                        leave_type = ""
                        num_days = 1.0 
                        
                        if "(ครึ่งวัน)" in val_status_new:
                            num_days = 0.5
                            leave_type = val_status_new.replace(" (ครึ่งวัน)", "").strip()
                        elif "(เต็มวัน)" in val_status_new:
                            num_days = 1.0
                            leave_type = val_status_new.replace(" (เต็มวัน)", "").strip()
                        elif "(ระบุเวลา)" in val_status_new:
                            num_days = 1.0 
                            leave_type = val_status_new.replace(" (ระบุเวลา)", "").strip()
                        elif "(" in val_status_new:
                            leave_type = self._parse_leave_type(val_status_new)
                        
                        if leave_type:
                            # [FIX V68] ไม่ต้องลบ Scan Logs ทิ้งแล้ว! (เพื่อให้เวลาเข้า-ออกยังอยู่)
                            # hr_database.delete_scan_logs_on_date(emp_id, date_obj) 
                            
                            hr_database.add_employee_leave(emp_id, date_obj, leave_type, num_days, "แก้ไขผ่าน Pop-up")

                # 3. บันทึกเวลา Scan (ถ้ามีการแก้ด้วยมือ)
                if scan_in_changed or scan_out_changed:
                    hr_database.delete_scan_logs_on_date(emp_id, date_obj)
                    if val_in_new:
                        try:
                            t = datetime.strptime(val_in_new, '%H:%M').time()
                            dt = datetime.combine(date_obj, t)
                            hr_database.add_manual_scan_log(emp_id, dt)
                        except ValueError: pass
                    if val_out_new:
                        try:
                            t = datetime.strptime(val_out_new, '%H:%M').time()
                            dt = datetime.combine(date_obj, t)
                            hr_database.add_manual_scan_log(emp_id, dt)
                        except ValueError: pass

                # 4. บันทึก OT
                if is_daily_emp and ot_changed:
                    if hasattr(hr_database, 'update_employee_ot_times'):
                         hr_database.update_employee_ot_times(emp_id, date_obj, val_ot_in_new, val_ot_out_new, new_calculated_ot_hours)
                    if hasattr(hr_database, 'update_ot_approval_status'):
                         hr_database.update_ot_approval_status(emp_id, date_obj, val_approved_new)
                
                if remark_changed:
                    hr_database.update_daily_remark(emp_id, date_obj, val_remark_new)

            if changes_detected > 0:
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลเรียบร้อย ({changes_detected} รายการ)", parent=popup_window)
                popup_window.destroy()
                self._run_processing() 
            else:
                messagebox.showinfo("ไม่เปลี่ยนแปลง", "ไม่พบการเปลี่ยนแปลงข้อมูล\n(อาจเพราะค่าใหม่เหมือนค่าเดิม)", parent=popup_window)

        except Exception as e:
            print(f"Save Error: {e}")
            import traceback; traceback.print_exc()
            messagebox.showerror("เกิดข้อผิดพลาด", f"บันทึกข้อมูลไม่ได้:\n{e}", parent=popup_window)

    def _export_details_to_excel(self, details_list, emp_name):
        """(ฉบับแก้ไข V65.0 - Fix KeyError date/work_date ใน Export)"""
        
        if not details_list:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลรายละเอียดที่จะ Export")
            return
        
        # Helper หาวันที่ (รองรับทั้ง object และ string)
        def get_date_str(row):
            if 'work_date' in row: return row['work_date'].strftime("%d-%m-%Y")
            return str(row.get('date', '')).replace("/", "-")

        start_date_str = get_date_str(details_list[0])
        end_date_str = get_date_str(details_list[-1])
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="บันทึกรายละเอียดปฏิทินทำงาน",
            initialfile=f"ปฏิทิน_{emp_name}_{start_date_str}_ถึง_{end_date_str}.xlsx"
        )
        if not file_path: return 
            
        try:
            # --- 1. เตรียมข้อมูล ---
            export_data = []
            for row in details_list:
                item = row.copy()
                
                # Normalize Keys (แปลงจาก DB Format -> Excel Format)
                if 'work_date' in item:
                    item['date'] = item['work_date'].strftime("%d/%m/%Y")
                    item['scan_in'] = item.get('work_in_time')
                    item['scan_out'] = item.get('work_out_time')
                    item['actual_late_mins'] = 0 
                    item['penalty_hrs'] = 0 
                    
                # คำนวณ Work Hours
                scan_in = item.get('scan_in') or item.get('work_in_time') or ""
                scan_out = item.get('scan_out') or item.get('work_out_time') or ""
                
                work_hrs_str = ""
                if scan_in and scan_out and scan_in != "-" and scan_out != "-":
                    try:
                        t_in = datetime.strptime(scan_in[:5], "%H:%M")
                        t_out = datetime.strptime(scan_out[:5], "%H:%M")
                        diff = t_out - t_in
                        total_seconds = diff.total_seconds()
                        hours = int(total_seconds // 3600)
                        minutes = int((total_seconds % 3600) // 60)
                        work_hrs_str = f"{hours}ชม. {minutes}น."
                    except: pass
                item['work_hours'] = work_hrs_str

                # Leave Hours
                status_text = item.get('status', '')
                leave_hrs_str = ""
                if "ลา" in status_text and "(" in status_text:
                     if "0.5" in status_text: leave_hrs_str = "4 ชม." # <--- จุดที่ต้องแก้
                     elif "1.0" in status_text: leave_hrs_str = "8 ชม."
                item['leave_hours'] = leave_hrs_str
                item['remark'] = item.get('remark') or ""
                
                export_data.append(item)

            # --- 2. สร้าง DataFrame ---
            df = pd.DataFrame(export_data)
            
            # --- 3. เปลี่ยนชื่อคอลัมน์ ---
            column_mapping = {
                "date": "วันที่",
                "status": "สถานะการทำงาน",
                "scan_in": "เวลาเข้างาน",
                "scan_out": "เวลาออกงาน",
                "work_hours": "ชม.ทำงาน",
                "leave_hours": "ชม.ลา",
                "actual_late_mins": "สาย (นาทีจริง)",
                "penalty_hrs": "ชม. ที่หัก",
                "ot_hours": "OT (ชม.)",
                "ot_in_time": "OT เข้า",
                "ot_out_time": "OT ออก",
                "remark": "หมายเหตุ" # 🛠️ [เพิ่มบรรทัดนี้]
            }
            # Rename เฉพาะที่มี
            df = df.rename(columns={k:v for k,v in column_mapping.items() if k in df.columns})
            
            # เลือกคอลัมน์ที่จะโชว์ (เพิ่ม "หมายเหตุ" ต่อท้าย)
            desired_cols = ["วันที่", "สถานะการทำงาน", "เวลาเข้างาน", "เวลาออกงาน", "ชม.ทำงาน", "ชม.ลา", "สาย (นาทีจริง)", "ชม. ที่หัก", "OT (ชม.)", "หมายเหตุ"]
            cols_to_use = [c for c in desired_cols if c in df.columns]
            df = df[cols_to_use]

            # --- 4. บันทึก ---
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("✅ สำเร็จ", f"บันทึกไฟล์ Excel เรียบร้อยแล้ว\nที่: {file_path}")
            
        except Exception as e:
            messagebox.showerror("❌ เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้:\n{e}")
            
            # --- 5. Clean ข้อมูลตัวเลข ---
            def clean_values(p):
                try:
                    val = float(p)
                    return val if val > 0 else ''
                except:
                    return ''
            
            if "สาย (นาทีจริง)" in df.columns:
                df["สาย (นาทีจริง)"] = df["สาย (นาทีจริง)"].apply(clean_values)
            if "ชม. ที่หัก" in df.columns:
                df["ชม. ที่หัก"] = df["ชม. ที่หัก"].apply(clean_values)

            # --- 6. บันทึก ---
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("✅ สำเร็จ", f"บันทึกไฟล์ Excel เรียบร้อยแล้ว\nที่: {file_path}")
            
        except Exception as e:
            messagebox.showerror("❌ เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้:\n{e}")
    
    def _process_data(self):
        """เริ่มประมวลผลข้อมูล (ฉบับปรับปรุง V3 - เช็ควันที่และแจ้งเตือนละเอียด)"""
        
        # 1. เช็ควันที่ก่อน (สำคัญที่สุด)
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            
            if not start_date or not end_date:
                messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือก วันที่เริ่มต้น และ วันที่สิ้นสุด")
                return

            # (!!! เพิ่มการเช็ค !!!) ห้ามเลือกวันที่ย้อนกลับ
            if start_date > end_date:
                messagebox.showerror("วันที่ผิดพลาด", "วันที่เริ่มต้น ต้องมาก่อน วันที่สิ้นสุด")
                return

        except Exception as e:
            messagebox.showerror("Error", f"รูปแบบวันที่ไม่ถูกต้อง: {e}")
            return

        # 2. ตรวจสอบข้อมูลนำเข้า (Excel/Scanner)
        if self.raw_log_data:
            # กรณีมีข้อมูลใหม่ -> บันทึกลง Database ก่อน
            try:
                inserted = hr_database.insert_scan_logs(self.raw_log_data)
                print(f"DEBUG: Inserted {inserted} logs.")
            except Exception as e:
                messagebox.showerror("DB Error", f"บันทึก Log ลงฐานข้อมูลไม่สำเร็จ:\n{e}")
                return
        else:
            # กรณีไม่มีข้อมูลใหม่ -> ถาม User ว่าจะเอาข้อมูลเก่าใน DB มาคิดไหม?
            # (ถ้ากด No -> จบการทำงานทันที)
            if not messagebox.askyesno("ยืนยัน", "คุณยังไม่ได้นำเข้าไฟล์ Log ใหม่ในรอบนี้\nต้องการประมวลผลจาก 'ข้อมูลเดิม' ในฐานข้อมูลหรือไม่?"):
                return

        # 3. เริ่มกระบวนการคำนวณ (Process)
        try:
            # เปลี่ยน Cursor เป็นรูปนาฬิกาทราย (Loading)
            self.config(cursor="wait")
            self.update()
            
            print(f"DEBUG: กำลังเรียก hr_database.process_attendance_summary({start_date}, {end_date})")
            
            # เรียกสมองคำนวณจาก hr_database
            summary_data = hr_database.process_attendance_summary(start_date, end_date)
            
            # คืนค่า Cursor ปกติ
            self.config(cursor="") 
            
            if not summary_data:
                messagebox.showinfo("ไม่พบข้อมูล", f"ประมวลผลเสร็จสิ้น\nแต่ไม่พบพนักงาน หรือไม่มีการสแกนนิ้วในช่วง {start_date} ถึง {end_date}")
                return

            # 4. แสดงผลลัพธ์ลงตาราง
            self.last_summary_report = summary_data
            
            
            messagebox.showinfo("สำเร็จ", f"ประมวลผลเรียบร้อยแล้ว!\nพนักงานทั้งหมด: {len(summary_data)} คน")

        except Exception as e:
            # ถ้าพัง ให้คืนค่า Cursor และฟ้อง Error
            self.config(cursor="")
            import traceback
            traceback.print_exc() # ปริ้นท์ลงจอดำด้วยเผื่อดูรายละเอียด
            messagebox.showerror("เกิดข้อผิดพลาด", f"ระบบไม่สามารถประมวลผลได้:\n{e}")

    def _calculate_time_diff(self, start_str, end_str):
        """คำนวณระยะห่างระหว่างเวลา 2 ค่า (คืนค่าเป็นชั่วโมง float)"""
        try:
            if not start_str or not end_str: return 0.0
            
            t_start = datetime.strptime(start_str, "%H:%M")
            t_end = datetime.strptime(end_str, "%H:%M")
            
            # กรณีข้ามวัน (เช่น เข้า 23:00 ออก 01:00)
            if t_end < t_start:
                t_end += pd.Timedelta(days=1)
                
            diff = t_end - t_start
            hours = diff.total_seconds() / 3600.0
            return round(hours, 2)
        except:
            return 0.0
