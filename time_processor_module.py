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
                   command=self._load_file, width=30).pack(side="left", padx=10)
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
    
    
    def _load_file(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Log (Excel หรือ CSV)",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            self.upload_status_label.config(text="⏳ กำลังโหลดไฟล์...", foreground="orange")
            self.update_idletasks()
            
            if file_path.endswith('.csv'):
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

            column_mapping = {
                'รหัสพนักงาน': ['รหัสพนักงาน', 'ID', 'รหัส', 'รับ'],
                'วันที่': ['วันที่', 'Date', 'วัน'],
                'เวลาเข้างาน': ['เวลาเข้างาน', 'เวลาเข้า', 'เวลาช้างาน', 'เวลาเช้างาน'],
                'เวลาออกงาน': ['เวลาออกงาน', 'เวลาออก', 'เวลาออกงาม']
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
            
            BAD_TIME_VALUES = {"", "0", "0:00", "nan", "nat", "none"} 

            print("=== ตัวอย่างข้อมูล 3 แถวแรก ===")
            for idx in range(min(3, len(df))):
                row = df.iloc[idx]
                print(f"แถว {idx+1}:")
                print(f"  รหัสพนักงาน: '{row['รหัสพนักงาน']}' (type: {type(row['รหัสพนักงาน'])})")
                print(f"  วันที่: '{row['วันที่']}' (type: {type(row['วันที่'])})")
                print(f"  เวลาเข้างาน: '{row['เวลาเข้างาน']}' (type: {type(row['เวลาเข้างาน'])})")
                print(f"  เวลาออกงาน: '{row['เวลาออกงาน']}' (type: {type(row['เวลาออกงาน'])})")
            print("=" * 50) 

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
                except Exception:
                    continue 

                if not emp_id or not date_str or emp_id.lower() == 'nan' or date_str.lower() == 'nan':
                    continue 

                time_in_str = str(row['เวลาเข้างาน']).strip()
                if time_in_str.lower() not in BAD_TIME_VALUES:
                    try: 
                        datetime_in_str = f"{date_str} {time_in_str}"
                        datetime_in_str_formatted = datetime_in_str.replace(".", ":") 
                        
                        formats_to_try = [
                            '%d/%m/%Y %H:%M',   
                            '%-d/%m/%Y %H:%M',  
                            '%d/%-m/%Y %H:%M',  
                            '%-d/%-m/%Y %H:%M', 
                        ]
                        
                        ts_in = None
                        for fmt in formats_to_try:
                            try:
                                ts_in = pd.to_datetime(datetime_in_str_formatted, format=fmt)
                                break
                            except:
                                continue
                        
                        if ts_in is None or pd.isna(ts_in):
                            ts_in = pd.to_datetime(datetime_in_str_formatted, dayfirst=True, errors='coerce')
                        
                        if ts_in is not None and not pd.isna(ts_in):
                            self.raw_log_data.append((emp_id, ts_in))
                        else:
                            skipped_rows += 1
                            
                    except Exception as e_in:
                        print(f"Skipping IN-log format error: {e_in} | Data: {datetime_in_str}")
                        skipped_rows += 1

                time_out_str = str(row['เวลาออกงาน']).strip()
                if time_out_str.lower() not in BAD_TIME_VALUES:
                    try:
                        datetime_out_str = f"{date_str} {time_out_str}"
                        datetime_out_str_formatted = datetime_out_str.replace(".", ":")
                        
                        formats_to_try = [
                            '%d/%m/%Y %H:%M',
                            '%-d/%m/%Y %H:%M',
                            '%d/%-m/%Y %H:%M',
                            '%-d/%-m/%Y %H:%M',
                        ]
                        
                        ts_out = None
                        for fmt in formats_to_try:
                            try:
                                ts_out = pd.to_datetime(datetime_out_str_formatted, format=fmt)
                                break
                            except:
                                continue
                        
                        if ts_out is None or pd.isna(ts_out):
                            ts_out = pd.to_datetime(datetime_out_str_formatted, dayfirst=True, errors='coerce')
                        
                        if ts_out is not None and not pd.isna(ts_out):
                            self.raw_log_data.append((emp_id, ts_out))
                        else:
                            skipped_rows += 1
                            
                    except Exception as e_out:
                        print(f"Skipping OUT-log format error: {e_out} | Data: {datetime_out_str}")
                        skipped_rows += 1
            
            self.upload_status_label.config(
                text=f"✅ ไฟล์: {os.path.basename(file_path)} (พบ {len(self.raw_log_data)} Log สแกน)", 
                foreground="green"
            )
            
            if skipped_rows > 0:
                skip_percentage = (skipped_rows / total_rows * 100) if total_rows > 0 else 0
                if skip_percentage > 10:
                    messagebox.showwarning(
                        "ข้ามบางรายการ", 
                        f"ข้าม {skipped_rows} รายการ ({skip_percentage:.1f}%)\n"
                        f"เนื่องจากรูปแบบเวลาไม่ถูกต้องหรือข้อมูลว่างเปล่า"
                    )
                else:
                    print(f"Info: ข้าม {skipped_rows} รายการที่มีเวลาว่างหรือผิดรูปแบบ")

            self.save_to_db_btn.config(state="normal")
            self.process_btn.config(state="normal") 

        except Exception as e:
            messagebox.showerror("File Error", f"ไม่สามารถอ่านไฟล์ได้:\n{e}")
            self.upload_status_label.config(text="❌ เกิดข้อผิดพลาดในการอ่านไฟล์", foreground="red")
    
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
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            if not start_date or not end_date: 
                raise ValueError("วันที่ไม่ได้เลือก หรือเลือกไม่ครบ")
        except Exception:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกช่วงวันที่ (วัน/เดือน/ปี) ให้ถูกต้อง")
            return
            
        if start_date > end_date:
            messagebox.showwarning("วันที่ผิดพลาด", "วันที่เริ่มต้น ต้องมาก่อนวันที่สิ้นสุด")
            return
            
        try:
            summary_report = hr_database.process_attendance_summary(start_date, end_date)

            print("--- ผลลัพธ์จาก 'การคำนวน' ---")
            print(summary_report[:5]) 

            self.last_summary_report = summary_report 
            self.export_btn.config(state="normal")  
            
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)

            self.result_frame.config(text=f"  ผลลัพธ์การประมวลผล ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})  ")
            
            if summary_report:
                for i, report in enumerate(summary_report):
                    tag_to_use = 'striped' if i % 2 == 0 else '' 
                    
                    self.result_tree.insert("", "end", iid=report['emp_id'], values=(
                        report['emp_id'],
                        report['name'],
                        report['emp_type'],     
                        report['department'], 
                        report['position'],   
                        f"{report['total_late_minutes']:.0f}", 
                        f"{report['total_late_hours']:.2f}",  
                        report['absent_days']
                    ), tags=(tag_to_use,))
            else:
                self.result_frame.config(text=f"  ผลลัพธ์: ไม่พบข้อมูลพนักงาน ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})  ")

        except Exception as e:
            self.export_btn.config(state="disabled") 
            messagebox.showerror("Processing Error", f"เกิดข้อผิดพลาดขณะประมวลผล:\n{e}")
        finally:
            pass
            
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
        """(ใหม่) เมื่อดับเบิลคลิกที่ตาราง ให้เปิดหน้าต่าง Pop-up (แบบตาราง tksheet)"""
        
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
        details_list_original = emp_data.get('details', [])
        
        win = tk.Toplevel(self)
        win.title(f"รายละเอียดปฏิทินทำงาน - {emp_name}")
        win.geometry("1350x550") # <--- ขยายหน้าต่างให้กว้างขึ้นอีกนิด
        win.transient(self) 
        win.grab_set()      
        
        sheet_frame = ttk.Frame(win, padding=(15, 15, 15, 0))
        sheet_frame.pack(fill="both", expand=True)

        sheet = Sheet(sheet_frame, 
                      show_toolbar=False,
                      show_top_left=False,
                      show_row_index=False,
                      show_header=True,
                      expand="both" 
                     )
        sheet.pack(fill="both", expand=True) 

        # --- 4. เตรียมข้อมูลสำหรับ tksheet (เพิ่มคอลัมน์ใหม่!) ---
        
        headers = ["วันที่", "สถานะการทำงาน", "เวลาเข้า", "เวลาออก", 
                   "ชม.ทำงาน", "ชม.ลา", "สาย(นาที)", "หัก(ชม.)"]
        sheet.headers(headers)
        
        # กำหนดความกว้าง (ปรับให้พอดีกับหน้าจอ)
        sheet.column_width(column=0, width=100) # วันที่
        sheet.column_width(column=1, width=450) # สถานะ (กว้างสุด)
        sheet.column_width(column=2, width=80)  # เข้า
        sheet.column_width(column=3, width=80)  # ออก
        sheet.column_width(column=4, width=100) # ชม.ทำงาน (ใหม่)
        sheet.column_width(column=5, width=80)  # ชม.ลา (ใหม่)
        sheet.column_width(column=6, width=80)  # สาย
        sheet.column_width(column=7, width=80)  # หัก
        
        sheet_data = []
        if not details_list_original:
            sheet_data.append(("", "(ไม่พบข้อมูล)", "", "", "", "", "", ""))
        else:
            for row_data in details_list_original:
                scan_in = row_data.get('scan_in', '')
                scan_out = row_data.get('scan_out', '')
                
                # --- (1) คำนวณชั่วโมงทำงาน (Work Hours) ---
                work_hrs_str = ""
                if scan_in and scan_out:
                    try:
                        t_in = datetime.strptime(scan_in, "%H:%M")
                        t_out = datetime.strptime(scan_out, "%H:%M")
                        # คำนวณผลต่าง (timedelta)
                        diff = t_out - t_in
                        total_seconds = diff.total_seconds()
                        
                        # หักเวลาพัก 1 ชม. (ถ้าทำงานข้ามเที่ยง)
                        # (Logic ง่ายๆ: ถ้าเข้าก่อน 12:00 และออกหลัง 13:00)
                        noon_start = t_in.replace(hour=12, minute=0)
                        noon_end = t_in.replace(hour=13, minute=0)
                        
                        if t_in < noon_end and t_out > noon_start:
                             # หักเวลาพักออก (สูงสุด 60 นาที)
                             overlap_start = max(t_in, noon_start)
                             overlap_end = min(t_out, noon_end)
                             break_seconds = (overlap_end - overlap_start).total_seconds()
                             if break_seconds > 0:
                                 total_seconds -= break_seconds
                        
                        hours = int(total_seconds // 3600)
                        minutes = int((total_seconds % 3600) // 60)
                        work_hrs_str = f"{hours}ชม. {minutes}น."
                    except:
                        work_hrs_str = "-"

                # --- (2) ดึงชั่วโมงลา (Leave Hours) ---
                # (สมมติว่าเราดึงจากสถานะ เช่น "ลา (ลากิจ) (08:00-11:00)")
                leave_hrs_str = ""
                status_text = row_data.get('status', '')
                if "ลา" in status_text and "(" in status_text:
                     # พยายามดึงตัวเลขชั่วโมงจากข้อความ (ถ้ามี)
                     # (นี่เป็นวิธีคร่าวๆ ในอนาคตอาจต้องส่งค่ามาจาก hr_database โดยตรงจะแม่นกว่า)
                     if "0.5" in status_text: leave_hrs_str = "4 ชม."
                     elif "1.0" in status_text: leave_hrs_str = "8 ชม."
                     # (กรณีลาชั่วโมง: คำนวณจากเวลาในวงเล็บยากหน่อยใน client-side)
                
                actual_late_val = row_data.get('actual_late_mins', 0)
                actual_late_str = f"{actual_late_val:.0f}" if actual_late_val > 0 else ""
                
                hours_val = row_data.get('penalty_hrs', 0.0)
                hours_str = f"{hours_val:.2f}" if hours_val > 0 else ""
                
                sheet_data.append([
                    row_data.get('date', ''), 
                    status_text, 
                    scan_in,
                    scan_out,
                    work_hrs_str,   # (ใหม่)
                    leave_hrs_str,  # (ใหม่)
                    actual_late_str, 
                    hours_str        
                ])
        
        sheet.set_sheet_data(sheet_data)
        
        # (โค้ด Highlight สี - เหมือนเดิม)
        for i, row_data in enumerate(details_list_original):
            row_status = row_data.get('status', '')
            bg_color = "#ffffff"
            fg_color = "#000000"
            if 'ขาดงาน' in row_status:
                bg_color = '#fddfe2'; fg_color = '#9f1f2e'
            elif 'มาสาย' in row_status:
                bg_color = '#fff4de'; fg_color = '#a05f00'
            elif 'ลา' in row_status: 
                bg_color = '#e0f0ff'; fg_color = '#00529e'
            elif 'วันหยุด' in row_status:
                bg_color = '#ffffff'; fg_color = 'gray'
            elif i % 2 == 1: bg_color = '#f0f0f0'
            sheet.highlight_rows(rows=[i], bg=bg_color, fg=fg_color)
            
        # ล็อคคอลัมน์ (ยกเว้นสถานะและเวลา)
        sheet.readonly_columns(columns=[0, 4, 5, 6, 7]) 
        
        leave_types = ["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาอื่นๆ"]
        status_options_base = ["วันทำงาน", "ขาดงาน", "มาสาย"] + [f"ลา ({t})" for t in leave_types]
        
        total_data_rows = sheet.get_total_rows() 
        
        for i in range(total_data_rows):
            original_status = str(sheet_data[i][1]).strip()
            # Logic Skip Dropdown (V17.0)
            is_complex_status = "(" in original_status and len(original_status) > 15
            if is_complex_status: continue 
            
            sheet.create_dropdown(
                r=i, c=1, values=status_options_base, set_value=original_status, state="readonly"
            )

        sheet.enable_bindings("single", "drag_select", "row_select", "column_width_resize", "arrowkeys", "right_click_popup_menu", "edit_cell")

        # (ส่วนปุ่มล่างสุด - เหมือนเดิม)
        btn_frame = ttk.Frame(win, padding=(15, 10, 15, 15))
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="📄 Export Excel", command=lambda: self._export_details_to_excel(details_list_original, emp_name)).pack(side="left")
        ttk.Button(btn_frame, text="💾 บันทึกการแก้ไข", command=lambda d=details_list_original, s=sheet, e=emp_id, w=win: self._save_details_from_popup(s, d, e, w), style="Success.TButton").pack(side="left", padx=10)
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

    def _save_details_from_popup(self, sheet, original_details_list, emp_id, popup_window):
        """(ฉบับ Debug V4) เพิ่มการ Print เพื่อตรวจสอบว่าทำไมไม่บันทึก"""
        
        try:
            if not messagebox.askyesno("ยืนยันการบันทึก",
                                     "คุณต้องการบันทึกการเปลี่ยนแปลงทั้งหมดนี้ใช่หรือไม่?",
                                      parent=popup_window):
                return 
            
            # 1. ดึงข้อมูลจากตาราง
            new_data_list_of_lists = sheet.get_sheet_data()
            print(f"DEBUG: จำนวนแถวในตาราง = {len(new_data_list_of_lists)}") # <--- เช็กจำนวนแถว
            
            # 2. สร้าง Dictionary
            new_data_map = {}
            headers = [
                "date", "status", "scan_in", "scan_out", 
                "work_hours", "leave_hours", 
                "actual_late_mins", "penalty_hrs"
            ]
            
            for row_list in new_data_list_of_lists:
                while len(row_list) < 8: row_list.append("") # เติมให้ครบ
                row_dict = {headers[i]: str(val).strip() for i, val in enumerate(row_list)} # แปลงเป็น string และตัดช่องว่าง
                
                date_key = row_dict['date']
                if date_key: new_data_map[date_key] = row_dict

            changes_detected = 0
            
            # 3. วนลูปเปรียบเทียบ
            for original_row in original_details_list:
                date_str = str(original_row['date']).strip()
                
                if date_str not in new_data_map:
                    print(f"DEBUG: หาวันที่ {date_str} ไม่เจอในตารางใหม่ (ข้าม)")
                    continue
                
                new_row = new_data_map[date_str]
                date_obj = self._parse_date_be(date_str)
                if not date_obj: continue 

                # --- เปรียบเทียบ (แบบละเอียด) ---
                # 1. สถานะ
                val_status_old = str(original_row['status']).strip()
                val_status_new = new_row['status']
                if val_status_new == "เลือกสถานะ": val_status_new = val_status_old # คืนค่าเดิมถ้าไม่ได้เลือก
                
                status_changed = (val_status_old != val_status_new)

                # 2. เวลาเข้า
                val_in_old = str(original_row.get('scan_in') or "").strip()
                val_in_new = new_row['scan_in']
                if val_in_new == "None": val_in_new = "" # แก้บั๊กบางทีตารางส่ง "None" มาเป็น text
                scan_in_changed = (val_in_old != val_in_new)

                # 3. เวลาออก
                val_out_old = str(original_row.get('scan_out') or "").strip()
                val_out_new = new_row['scan_out']
                if val_out_new == "None": val_out_new = ""
                scan_out_changed = (val_out_old != val_out_new)

                # ถ้าไม่มีอะไรเปลี่ยนเลย -> ข้าม
                if not (status_changed or scan_in_changed or scan_out_changed):
                    continue 
                
                # ถ้ามาถึงตรงนี้ แปลว่าเจอการเปลี่ยนแปลง!
                print(f"DEBUG: พบการแก้ไขที่วันที่ {date_str}")
                print(f"   - Status: '{val_status_old}' -> '{val_status_new}' (Changed: {status_changed})")
                print(f"   - In:     '{val_in_old}' -> '{val_in_new}' (Changed: {scan_in_changed})")
                print(f"   - Out:    '{val_out_old}' -> '{val_out_new}' (Changed: {scan_out_changed})")
                
                changes_detected += 1
                
                # --- บันทึกลงฐานข้อมูล ---
                new_status_is_leave = "ลา" in val_status_new and "(" in val_status_new
                original_status_is_leave = "ลา" in val_status_old and "(" in val_status_old

                # 1. จัดการสถานะลา
                if status_changed:
                    if (not new_status_is_leave) and original_status_is_leave:
                        # ยกเลิกการลา (ลบออกจาก DB)
                        hr_database.delete_leave_record_on_date(emp_id, date_obj)
                    elif new_status_is_leave:
                        # เปลี่ยนเป็นลา หรือ เปลี่ยนประเภทลา (บันทึกทับ)
                        leave_type = self._parse_leave_type(val_status_new)
                        if leave_type:
                            # ลบสแกนเก่าออกก่อน (ถ้าลาเต็มวันมักไม่มีสแกน)
                            if not scan_in_changed and not scan_out_changed: # ลบเฉพาะถ้าไม่ได้ตั้งใจแก้เวลาด้วย
                                hr_database.delete_scan_logs_on_date(emp_id, date_obj)
                            hr_database.add_employee_leave(emp_id, date_obj, leave_type, 1.0, "แก้ไขผ่าน Pop-up")

                # 2. จัดการเวลาสแกน (ถ้ามีการแก้ไขเวลา)
                if scan_in_changed or scan_out_changed:
                    # ลบอันเก่าทิ้งก่อนเสมอ เพื่อกันซ้ำ
                    hr_database.delete_scan_logs_on_date(emp_id, date_obj)
                    
                    # บันทึกเวลาเข้า
                    if val_in_new:
                        try:
                            t = datetime.strptime(val_in_new, '%H:%M').time()
                            dt = datetime.combine(date_obj, t)
                            hr_database.add_manual_scan_log(emp_id, dt)
                        except ValueError: 
                            print(f"   ❌ รูปแบบเวลาเข้าไม่ถูกต้อง: {val_in_new}")

                    # บันทึกเวลาออก
                    if val_out_new:
                        try:
                            t = datetime.strptime(val_out_new, '%H:%M').time()
                            dt = datetime.combine(date_obj, t)
                            hr_database.add_manual_scan_log(emp_id, dt)
                        except ValueError: 
                            print(f"   ❌ รูปแบบเวลาออกไม่ถูกต้อง: {val_out_new}")
                        
            if changes_detected > 0:
                print(f"DEBUG: บันทึกเสร็จสิ้น {changes_detected} รายการ")
                messagebox.showinfo("สำเร็จ", f"บันทึกการแก้ไข {changes_detected} รายการเรียบร้อย", parent=popup_window)
                popup_window.destroy()
                self._run_processing() 
            else:
                print("DEBUG: ไม่พบความเปลี่ยนแปลงใดๆ")
                messagebox.showinfo("ไม่เปลี่ยนแปลง", "ไม่พบการเปลี่ยนแปลงข้อมูล (หรือข้อมูลเหมือนเดิม)", parent=popup_window)

        except Exception as e:
            print(f"Save Error Details: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}", parent=popup_window)

    def _export_details_to_excel(self, details_list, emp_name):
        """(ใหม่) ส่งออกข้อมูล "ปฏิทินสถานะ" (จาก Pop-up) เป็น Excel"""
        
        if not details_list:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบข้อมูลรายละเอียดที่จะ Export")
            return
            
        start_date_str = details_list[0]['date'].replace("/", "-")
        end_date_str = details_list[-1]['date'].replace("/", "-")
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="บันทึกรายละเอียดปฏิทินทำงาน",
            initialfile=f"ปฏิทิน_{emp_name}_{start_date_str}_ถึง_{end_date_str}.xlsx"
        )
        if not file_path:
            return 
            
        try:
            # --- 1. เตรียมข้อมูล (คำนวณคอลัมน์เพิ่มให้เหมือนหน้าจอ) ---
            export_data = []
            for row in details_list:
                # คัดลอกข้อมูลเดิมมา
                item = row.copy()
                
                # (ก) คำนวณชม.ทำงาน (Work Hours)
                scan_in = item.get('scan_in', '')
                scan_out = item.get('scan_out', '')
                work_hrs_str = ""
                if scan_in and scan_out:
                    try:
                        t_in = datetime.strptime(scan_in, "%H:%M")
                        t_out = datetime.strptime(scan_out, "%H:%M")
                        diff = t_out - t_in
                        total_seconds = diff.total_seconds()
                        
                        # หักพักเที่ยง 12:00-13:00
                        noon_start = t_in.replace(hour=12, minute=0)
                        noon_end = t_in.replace(hour=13, minute=0)
                        if t_in < noon_end and t_out > noon_start:
                             overlap_start = max(t_in, noon_start)
                             overlap_end = min(t_out, noon_end)
                             break_seconds = (overlap_end - overlap_start).total_seconds()
                             if break_seconds > 0:
                                 total_seconds -= break_seconds
                        
                        hours = int(total_seconds // 3600)
                        minutes = int((total_seconds % 3600) // 60)
                        work_hrs_str = f"{hours}ชม. {minutes}น."
                    except:
                        work_hrs_str = "-"
                item['work_hours'] = work_hrs_str

                # (ข) คำนวณชม.ลา (Leave Hours)
                status_text = item.get('status', '')
                leave_hrs_str = ""
                if "ลา" in status_text and "(" in status_text:
                     if "0.5" in status_text: leave_hrs_str = "4 ชม."
                     elif "1.0" in status_text: leave_hrs_str = "8 ชม."
                item['leave_hours'] = leave_hrs_str
                
                export_data.append(item)

            # --- 2. สร้าง DataFrame ---
            df = pd.DataFrame(export_data)
            
            # --- 3. เปลี่ยนชื่อคอลัมน์ (Map) ---
            column_mapping = {
                "date": "วันที่",
                "status": "สถานะการทำงาน",
                "scan_in": "เวลาเข้างาน",
                "scan_out": "เวลาออกงาน",
                "work_hours": "ชม.ทำงาน",    # (เพิ่มใหม่)
                "leave_hours": "ชม.ลา",       # (เพิ่มใหม่)
                "actual_late_mins": "สาย (นาทีจริง)",
                "penalty_hrs": "ชม. ที่หัก"
            }
            df = df.rename(columns=column_mapping)
            
            # --- 4. เลือกและเรียงลำดับคอลัมน์ ---
            final_columns = [
                "วันที่", "สถานะการทำงาน", "เวลาเข้างาน", "เวลาออกงาน", 
                "ชม.ทำงาน", "ชม.ลา",         # (เพิ่มใหม่)
                "สาย (นาทีจริง)", "ชม. ที่หัก"
            ]
            
            # กรองเฉพาะคอลัมน์ที่มีอยู่จริง (ป้องกัน Error)
            cols_to_use = [col for col in final_columns if col in df.columns]
            df = df[cols_to_use]
            
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
            self._display_summary(summary_data)
            
            messagebox.showinfo("สำเร็จ", f"ประมวลผลเรียบร้อยแล้ว!\nพนักงานทั้งหมด: {len(summary_data)} คน")

        except Exception as e:
            # ถ้าพัง ให้คืนค่า Cursor และฟ้อง Error
            self.config(cursor="")
            import traceback
            traceback.print_exc() # ปริ้นท์ลงจอดำด้วยเผื่อดูรายละเอียด
            messagebox.showerror("เกิดข้อผิดพลาด", f"ระบบไม่สามารถประมวลผลได้:\n{e}")