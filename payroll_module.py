# (ไฟล์: payroll_module.py)
# (เวอร์ชัน V15.0 - เพิ่มฟังก์ชันพิมพ์สลิป PDF พร้อมโลโก้)
from reportlab.lib.pagesizes import A4, landscape
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog 
from custom_widgets import DateDropdown 
import hr_database
from datetime import datetime
import calendar
import pandas as pd
import os
from fpdf import FPDF  # (ต้องมีไลบรารีนี้)
from daily_timesheet import DailyTimesheetWindow
import smtplib
import ssl
from email.message import EmailMessage
from tksheet import Sheet
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
import io
from bahttext import bahttext

class PayrollModule(ttk.Frame):

    def _calculate_hourly_rate(self, salary):
        """
        คำนวณอัตราค่าจ้างรายชั่วโมง (ตามกฎหมายแรงงาน หาร 30 วัน 8 ชม.)
        """
        try:
            salary_val = float(salary)
            # หาร 30 วัน แล้วหาร 8 ชั่วโมง
            return (salary_val / 30) / 8
        except:
            return 0.0

    def _calculate_deduction(self, salary, missing_hours):
        """คำนวณยอดเงินที่ต้องหัก"""
        hourly_rate = self._calculate_hourly_rate(salary)
        return hourly_rate * float(missing_hours)

    def _calculate_tax_step_ladder(self, net_income):
        """คำนวณภาษีตามขั้นบันได (Step Ladder) ปีปัจจุบัน"""
        
        # ตารางภาษี: (เพดานเงินได้, อัตราภาษี)
        # 0 - 150,000 = ยกเว้น (0%)
        # 150,001 - 300,000 = 5%
        # 300,001 - 500,000 = 10%
        # ... ไปเรื่อยๆ
        
        tax = 0.0
        
        # ขั้นที่ 1: 0 - 150,000 (ยกเว้น)
        if net_income <= 150000:
            return 0.0
        
        # ขั้นที่ 2: 150,001 - 300,000 (5%)
        # คำนวณส่วนที่เกิน 150,000 แต่ไม่เกิน 300,000
        amount = min(net_income, 300000) - 150000
        tax += amount * 0.05
        if net_income <= 300000: return tax
        
        # ขั้นที่ 3: 300,001 - 500,000 (10%)
        amount = min(net_income, 500000) - 300000
        tax += amount * 0.10
        if net_income <= 500000: return tax
        
        # ขั้นที่ 4: 500,001 - 750,000 (15%)
        amount = min(net_income, 750000) - 500000
        tax += amount * 0.15
        if net_income <= 750000: return tax
        
        # ขั้นที่ 5: 750,001 - 1,000,000 (20%)
        amount = min(net_income, 1000000) - 750000
        tax += amount * 0.20
        if net_income <= 1000000: return tax
        
        # ขั้นที่ 6: 1,000,001 - 2,000,000 (25%)
        amount = min(net_income, 2000000) - 1000000
        tax += amount * 0.25
        if net_income <= 2000000: return tax
        
        # ขั้นที่ 7: 2,000,001 - 5,000,000 (30%)
        amount = min(net_income, 5000000) - 2000000
        tax += amount * 0.30
        if net_income <= 5000000: return tax
        
        # ขั้นที่ 8: 5,000,001 ขึ้นไป (35%)
        amount = net_income - 5000000
        tax += amount * 0.35
        
        return tax


    def _open_email_approval_window(self):
        """(Approver) หน้าจอตรวจสอบและอนุมัติการส่งอีเมล"""
        if self.current_user['role'] != 'approver':
            messagebox.showerror("สิทธิ์ไม่ถึง", "เฉพาะผู้อนุมัติ (Approver) เท่านั้น")
            return

        win = tk.Toplevel(self)
        win.title("📧 อนุมัติการส่งสลิปเงินเดือน")
        win.geometry("900x500")
        
        columns = ("queue_id", "id", "name", "email", "path")
        tree = ttk.Treeview(win, columns=columns, show="headings")
        
        tree.heading("queue_id", text="Queue ID")
        tree.heading("id", text="รหัสพนักงาน")
        tree.heading("name", text="ชื่อ-สกุล")
        tree.heading("email", text="อีเมลปลายทาง")
        tree.heading("path", text="ไฟล์ PDF (คลิกเพื่อดู)")
        
        tree.column("queue_id", width=0, stretch=False)
        tree.column("id", width=80)
        tree.column("name", width=200)
        tree.column("email", width=200)
        tree.column("path", width=300)
        
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # โหลดข้อมูล
        pending_list = hr_database.get_pending_emails()
        for item in pending_list:
            tree.insert("", "end", values=(
                item['queue_id'],
                item['emp_id'], 
                f"{item['fname']} {item['lname']}", 
                item['receiver_email'],
                item['pdf_path']
            ))
        
            
        # --- ฟังก์ชันกดดู PDF ---
        def preview_pdf(event):
            item_id = tree.selection()
            if not item_id: return
            vals = tree.item(item_id[0], "values")
            pdf_path = vals[4]
            
            if os.path.exists(pdf_path):
                try:
                    os.startfile(pdf_path)
                except Exception as e:
                    messagebox.showerror("Error", f"เปิดไฟล์ไม่ได้: {e}")
            else:
                messagebox.showerror("Error", "หาไฟล์ไม่เจอ (อาจถูกลบหรือ Path ผิด)")
        
        tree.bind("<Double-1>", preview_pdf)

        # --- ปุ่มสั่งการ ---
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", pady=10, padx=10)
        
        # (ฟังก์ชันใหม่: ปฏิเสธ/ลบรายการ)
        def reject_selection():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือกรายการที่ต้องการปฏิเสธ")
                return

            if not messagebox.askyesno("ยืนยัน", f"ต้องการ 'ปฏิเสธ' รายการที่เลือก ({len(selection)} รายการ) ใช่หรือไม่?"):
                return

            for item_id in selection:
                vals = tree.item(item_id)['values']
                queue_id = vals[0] # ID ที่ซ่อนไว้
                
                # อัปเดตสถานะใน DB เป็น 'rejected'
                hr_database.update_email_status(queue_id, 'rejected')
                
                # ลบออกจากหน้าจอ
                tree.delete(item_id)
            
            messagebox.showinfo("สำเร็จ", "ปฏิเสธรายการเรียบร้อยแล้ว")

        # (ฟังก์ชันเดิม: อนุมัติและส่ง)
        def approve_and_send():
            items = tree.get_children()
            if not items:
                messagebox.showinfo("ว่างเปล่า", "ไม่มีรายการรออนุมัติ")
                return

            if not messagebox.askyesno("ยืนยัน", f"ต้องการอนุมัติและส่งอีเมลทั้ง {len(items)} รายการ ใช่หรือไม่?"):
                return

            # [HARDCODE MODE] สำหรับทดสอบ
            sender_email = "pakawantomdee@gmail.com"       # <---  แก้เป็น Gmail 
            sender_password = "error"   # <---  แก้เป็นรหัส App 16 หลัก
            
            sent_count = 0
            error_count = 0
            
            context = ssl.create_default_context()
            try:
                win.config(cursor="wait")
                win.update()

                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                    smtp.login(sender_email, sender_password)
                    
                    for item_id in items:
                        vals = tree.item(item_id)['values']
                        q_id = vals[0]
                        emp_id = vals[1]
                        emp_name = vals[2]
                        email_to = vals[3]
                        pdf_file = vals[4]
                        
                        if not email_to:
                            error_count += 1
                            continue

                        try:
                            msg = EmailMessage()
                            msg['Subject'] = f"สลิปเงินเดือน - {emp_name}"
                            msg['From'] = sender_email
                            msg['To'] = email_to
                            msg.set_content(f"เรียนคุณ {emp_name},\n\nทางบริษัทขอนำส่งสลิปเงินเดือนประจำงวดครับ\n(เอกสารแนบ)\n\nฝ่ายบุคคล")

                            if os.path.exists(pdf_file):
                                with open(pdf_file, 'rb') as f:
                                    file_data = f.read()
                                    file_name = os.path.basename(pdf_file)
                                msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)
                                
                                smtp.send_message(msg)
                                hr_database.update_email_status(q_id, 'sent')
                                sent_count += 1
                                tree.delete(item_id)
                                win.update()
                            else:
                                error_count += 1
                                print(f"File not found: {pdf_file}")

                        except Exception as e:
                            error_count += 1
                            print(f"Failed to send to {emp_name}: {e}")

                win.config(cursor="")
                messagebox.showinfo("เสร็จสิ้น", f"ส่งสำเร็จ: {sent_count} รายการ\nผิดพลาด: {error_count} รายการ", parent=win)
                if not tree.get_children(): win.destroy()

            except Exception as e:
                win.config(cursor="")
                messagebox.showerror("SMTP Error", f"เกิดข้อผิดพลาดในการเชื่อมต่ออีเมล:\n{e}\n(ตรวจสอบ Email/App Password)", parent=win)

        # ปุ่มสั่งการ 
        ttk.Button(btn_frame, text="✅ อนุมัติและส่งทั้งหมด", command=approve_and_send, style="Success.TButton").pack(side="right", padx=5)
        
        ttk.Button(btn_frame, text="❌ ปฏิเสธ/ลบ (เลือกรายการ)", command=reject_selection).pack(side="right", padx=5)
        
        ttk.Label(btn_frame, text="💡 คลิกเลือกรายการเพื่อลบ หรือ กดส่งทั้งหมด", foreground="gray").pack(side="left")

    def _request_email_approval(self):
        """(HR) สร้างไฟล์ PDF แล้วส่งคำขอไปให้ Approver"""
        
        # 1. เช็คเดือน/ปี
        y_ce, m_int = self._get_selected_dates()
        if not y_ce: return

        if not messagebox.askyesno("ยืนยัน", "ต้องการสร้างใบคำขอส่งสลิปเงินเดือนทางอีเมล ใช่หรือไม่?"):
            return
            
        # ตรวจสอบโฟลเดอร์กลาง (ต้องใช้ Path กลางเพื่อให้ Approver เปิดดูได้)
        # เช่น \\192.168.1.51\HR_Documents\Temp_Payslips
        shared_folder = r"\\192.168.1.51\HR_System_Documents\Temp_Payslips"
        if not os.path.exists(shared_folder):
            try:
                os.makedirs(shared_folder)
            except:
                messagebox.showerror("Error", f"ไม่สามารถเข้าถึงโฟลเดอร์กลาง: {shared_folder}")
                return

        success_count = 0
        
        # 2. วนลูปสร้าง PDF ทีละคน
        # (สมมติวนลูปจาก self.last_payroll_results)
        if not self.last_payroll_results:
             messagebox.showwarning("เตือน", "กรุณากดคำนวณเงินเดือนก่อน")
             return

        for data in self.last_payroll_results:
            emp_id = data['emp_id']
            
            # ดึงอีเมลพนักงาน (สมมติว่าอยู่ใน data หรือต้องไป query เพิ่ม)
            # สมมติว่าใน data ไม่มี ต้องไปดึงจาก asset หรือ employee info
            emp_assets = hr_database.load_single_employee(emp_id).get('assets', {})
            receiver_email = emp_assets.get('email', '')
            
            if not receiver_email:
                print(f"Skip {emp_id}: No email")
                continue

            # ตั้งชื่อไฟล์
            filename = f"Payslip_{emp_id}_{y_ce}_{m_int}.pdf"
            filepath = os.path.join(shared_folder, filename)
            
            try:
                # สร้าง PDF (ใช้ฟังก์ชันเดิมที่มีอยู่ แต่ส่ง list คนเดียว)
                self._generate_pdf([data], filepath)
                
                # บันทึกลง DB
                hr_database.add_email_request(
                    emp_id, m_int, y_ce, filepath, receiver_email, 
                    self.current_user['username']
                )
                success_count += 1
                
            except Exception as e:
                print(f"Error generating {emp_id}: {e}")

        messagebox.showinfo("สำเร็จ", f"ส่งคำขออนุมัติเรียบร้อย {success_count} รายการ\nรอ Approver ตรวจสอบ")
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user

        # เก็บข้อมูล Input ละเอียด: { 'EMP001': { 'ot': 500, 'tax': 100, ... }, ... }
        self.payroll_inputs = {} 
        self.last_payroll_results = []

        self.THAI_MONTHS = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
            5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
            9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }
        self.MONTH_TO_INT = {v: k for k, v in self.THAI_MONTHS.items()}

        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        self._build_control_panel(main_frame)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(15, 0))

        self.tab1 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab1, text="  ขั้นตอนที่ 1: 📝 ป้อนข้อมูลรายรับ/รายจ่าย  ")
        self._build_input_tab(self.tab1)

        self.tab2 = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab2, text="  ขั้นตอนที่ 2: 📊 ผลการคำนวณและ Export  ")
        self._build_results_tab(self.tab2)

    def _build_control_panel(self, parent):
        process_frame = ttk.LabelFrame(parent, text="  รอบการคำนวณ  ", padding=15)
        process_frame.pack(fill="x")
        
        filter_frame = ttk.LabelFrame(process_frame, text="  ตัวกรองด่วน  ", padding=10)
        filter_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_frame, text="ปี (พ.ศ.):").pack(side="left", padx=(5,5))
        current_year_be = datetime.now().year + 543
        year_values = [str(y) for y in range(current_year_be + 1, current_year_be - 5, -1)]
        self.year_combo = ttk.Combobox(filter_frame, values=year_values, width=8, state="readonly", font=("Segoe UI", 10))
        self.year_combo.set(str(current_year_be))
        self.year_combo.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="เดือน:").pack(side="left", padx=5)
        self.month_combo = ttk.Combobox(filter_frame, values=list(self.THAI_MONTHS.values()), width=15, state="readonly", font=("Segoe UI", 10))
        self.month_combo.set(self.THAI_MONTHS[datetime.now().month])
        self.month_combo.pack(side="left", padx=5)

        btn_frame = ttk.Frame(filter_frame)
        btn_frame.pack(side="left", padx=10)
        ttk.Button(btn_frame, text="1-15", command=self._set_date_1_15, width=8).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="16-สิ้นเดือน", command=self._set_date_16_end, width=10).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="ทั้งเดือน", command=self._set_date_month, width=8).pack(side="left", padx=2)

        date_frame = ttk.Frame(process_frame)
        date_frame.pack(fill="x")
        ttk.Label(date_frame, text="คำนวณตั้งแต่วันที่:", font=("Segoe UI", 10)).pack(side="left", padx=5)
        self.start_date_entry = DateDropdown(date_frame, font=("Segoe UI", 10))
        self.start_date_entry.pack(side="left", padx=5)
        ttk.Label(date_frame, text="ถึงวันที่:", font=("Segoe UI", 10)).pack(side="left", padx=5)
        self.end_date_entry = DateDropdown(date_frame, font=("Segoe UI", 10))
        self.end_date_entry.pack(side="left", padx=5)

    def _build_input_tab(self, parent_tab):
        """สร้าง Tab สำหรับเลือกพนักงานและกรอกข้อมูล (แยกแท็บย่อย: ประจำ vs จ้างเหมา/ที่ปรึกษา)"""
        
        # 1. ส่วนหัว
        top_frame = ttk.Frame(parent_tab)
        top_frame.pack(fill="x", pady=(0, 10))

        # ปุ่มคำสั่ง
        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(side="left")

        ttk.Button(btn_frame, text="🔄 โหลดรายชื่อทั้งหมด", command=self._load_employees_to_input_tree).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="💰 ดึงยอด A+ Smart", command=self._sync_commission_from_asmart, style="Primary.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📅 บันทึก Timesheet", command=self._open_daily_timesheet).pack(side="left", padx=5)
        
        # ปุ่มเลือก (ทำงานกับแท็บที่เปิดอยู่)
        ttk.Button(btn_frame, text="☑ เลือกทั้งหมด", command=lambda: self._toggle_all_selection(True)).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="☐ ไม่เลือก", command=lambda: self._toggle_all_selection(False)).pack(side="left", padx=2)

        # 2. สร้าง Notebook แยกแท็บย่อย
        self.input_notebook = ttk.Notebook(parent_tab)
        self.input_notebook.pack(fill="both", expand=True)

        # --- Tab 1: พนักงานทั่วไป ---
        self.tab_regular = ttk.Frame(self.input_notebook, padding=5)
        self.input_notebook.add(self.tab_regular, text=" 👥 พนักงานประจำ / รายวัน ")
        # สร้างตาราง 1 (เรียก Helper)
        self.tree_reg = self._create_treeview(self.tab_regular) 

        # --- Tab 2: จ้างเหมา / ที่ปรึกษา ---
        self.tab_contract = ttk.Frame(self.input_notebook, padding=5)
        self.input_notebook.add(self.tab_contract, text=" 🤝 สัญญาจ้างเหมา / ที่ปรึกษา ")
        # สร้างตาราง 2 (เรียก Helper)
        self.tree_cont = self._create_treeview(self.tab_contract) 

        # Footer
        ttk.Label(parent_tab, text="💡 แยกแท็บตามประเภทสัญญา | คลิกช่อง 'เลือก' เพื่อคำนวณ | ดับเบิลคลิกชื่อ เพื่อกรอกเงินเพิ่ม/หัก", foreground="gray").pack(pady=5)
    
    def _on_tree_click(self, event):
        """จัดการคลิกที่ตาราง: ถ้าคลิกช่องแรก ให้สลับ ☑ / ☐"""
        region = self.input_tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.input_tree.identify_column(event.x)
            
            # ถ้าคลิกคอลัมน์ที่ 1 (ช่อง select)
            if col == "#1":
                item_id = self.input_tree.identify_row(event.y)
                if not item_id: return
                
                # เช็คว่าเป็น separator หรือไม่
                if 'separator' in self.input_tree.item(item_id, "tags"): return

                # สลับค่า
                current_val = self.input_tree.item(item_id, "values")[0]
                new_val = "☐" if current_val == "☑" else "☑"
                
                # อัปเดตค่าในตาราง (รักษาค่าอื่นๆ ไว้)
                old_values = self.input_tree.item(item_id, "values")
                self.input_tree.item(item_id, values=(new_val, *old_values[1:]))

    def _toggle_all_selection(self, select_all=True):
        symbol = "☑" if select_all else "☐"
        for item in self.input_tree.get_children():
            if 'separator' in self.input_tree.item(item, "tags"): continue
            old_vals = self.input_tree.item(item, "values")
            self.input_tree.item(item, values=(symbol, *old_vals[1:]))

    def _build_results_tab(self, parent_tab):
        """สร้างหน้าจอแสดงผลลัพธ์และปุ่มควบคุม (Results Tab)"""
        
        # สร้าง Container หลักสำหรับส่วนปุ่มกด
        # ใช้ LabelFrame เพื่อตีกรอบให้ดูเป็นสัดส่วน และไม่กินพื้นที่ Grid มากเกินไป
        control_panel = ttk.LabelFrame(parent_tab, text=" แผงควบคุม (Control Panel) ", padding=10)
        control_panel.pack(fill="x", pady=(0, 10))
        
        # --- แถวที่ 1: การจัดการหลัก (คำนวณ / ยืนยันการจ่าย / ประวัติ) ---
        row1 = ttk.Frame(control_panel)
        row1.pack(fill="x", pady=2)
        
        ttk.Label(row1, text="1. จัดการเงินเดือน:", width=15, font=("", 9, "bold")).pack(side="left")
        
        # ปุ่มคำณวน
        ttk.Button(row1, text="🚀 คำนวณเงินเดือน", command=self._run_payroll_calculation, style="Success.TButton").pack(side="left", padx=2)
        
        # [แก้ไข] เปลี่ยนชื่อปุ่มเป็น "ยืนยันการจ่าย"
        self.save_db_btn = ttk.Button(row1, text="✅ ยืนยันการจ่าย (บันทึก)", 
                                      command=self._save_payroll_to_database, 
                                      state="disabled")
        self.save_db_btn.pack(side="left", padx=2)
        
        # ปุ่มดูประวัติ
        ttk.Button(row1, text="📜 ดูประวัติย้อนหลัง", command=self._open_history_window).pack(side="left", padx=2)
        
        # คั่นด้วย Separator แนวตั้ง
        ttk.Separator(row1, orient="vertical").pack(side="left", fill="y", padx=5)

        self.export_btn = ttk.Button(row1, text="📄 Export Excel", command=self._export_payroll_to_excel, state="disabled")
        self.export_btn.pack(side="left", padx=2)
        
        self.print_btn = ttk.Button(row1, text="🖨️ พิมพ์สลิป (PDF)", command=self._print_selected_payslip, state="disabled")
        self.print_btn.pack(side="left", padx=2)

        # --- แถวที่ 2: นำส่งสรรพากร/ประกันสังคม (รายเดือน) ---
        row2 = ttk.Frame(control_panel)
        row2.pack(fill="x", pady=2)
        
        ttk.Label(row2, text="2. นำส่งรายเดือน:", width=15, font=("", 9, "bold")).pack(side="left")

        self.pnd1_btn = ttk.Button(row2, text="🏛️ ใบปะหน้า ภ.ง.ด.1", command=self._print_pnd1_summary, state="disabled")
        self.pnd1_btn.pack(side="left", padx=2)

        self.pnd3_btn = ttk.Button(row2, text="🏛️ ใบปะหน้า ภ.ง.ด.3", command=self._print_pnd3_summary, state="disabled")
        self.pnd3_btn.pack(side="left", padx=2)

        self.sso_btn = ttk.Button(row2, text="🏥 สปส. 1-10 (Excel)", command=self._export_sso_excel, state="disabled")
        self.sso_btn.pack(side="left", padx=2)

        # --- แถวที่ 3: รายปี / ฟีเจอร์พิเศษ (Email) ---
        row3 = ttk.Frame(control_panel)
        row3.pack(fill="x", pady=2)
        
        ttk.Label(row3, text="3. รายปี/อื่นๆ:", width=15, font=("", 9, "bold")).pack(side="left")

        self.pnd1k_btn = ttk.Button(row3, text="📄 ภ.ง.ด.1ก (Excel)", command=self._export_pnd1k_excel)
        self.pnd1k_btn.pack(side="left", padx=2)

        self.pnd1k_pdf_btn = ttk.Button(row3, text="📄 ภ.ง.ด.1ก (PDF)", command=self._print_pnd1k_pdf)
        self.pnd1k_pdf_btn.pack(side="left", padx=2)

        self.btn_50tawi = ttk.Button(row3, text="📄 ใบ 50 ทวิ (รายคน)", command=self._print_50tawi_pdf)
        self.btn_50tawi.pack(side="left", padx=2)
        
        # คั่นด้วย Separator แนวตั้ง
        ttk.Separator(row3, orient="vertical").pack(side="left", fill="y", padx=5)

        self.email_req_btn = ttk.Button(row3, text="📧 ขอส่งสลิป (Email)", command=self._request_email_approval, state="disabled")
        self.email_req_btn.pack(side="left", padx=2)

        # แสดงปุ่มอนุมัติเฉพาะ Role Approver
        if self.current_user['role'] == 'approver':
            self.email_approve_btn = ttk.Button(row3, text="✅ อนุมัติการส่งเมล", command=self._open_email_approval_window)
            self.email_approve_btn.pack(side="left", padx=2)

        # --- ส่วนแสดงผลตาราง (Sheet) ---
        # ใช้ Frame หุ้มเพื่อให้ขยายเต็มพื้นที่ที่เหลือ
        sheet_frame = ttk.Frame(parent_tab)
        sheet_frame.pack(fill="both", expand=True, pady=(5,0))
        
        self.results_sheet = Sheet(sheet_frame,
                                   show_x_scrollbar=True,
                                   show_y_scrollbar=True,
                                   headers=None,
                                   theme="light blue"
                                  )
        self.results_sheet.pack(fill="both", expand=True)
        
        # ตั้งค่า Binding (การคลิก/เลือก)
        self.results_sheet.enable_bindings(
            "single_select",
            "row_select",
            "column_width_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "copy"
        )
        
        # Bind Event ดับเบิ้ลคลิกเพื่อดูรายละเอียด
        self.results_sheet.bind("<Double-1>", self._on_result_double_click)
        self.results_sheet.extra_bindings("cell_double_click", func=self._on_result_double_click)
    def _on_result_double_click(self, event=None):
        """ทำงานเมื่อดับเบิลคลิก (เวอร์ชั่นรองรับ ค่าเที่ยว, OT, เบี้ยขยัน)"""
        # print(f"\n--- 🖱️ DEBUG: Checking Click ---") 
        row = None
        col = None

        if not hasattr(event, 'x') or not hasattr(event, 'y'):
            return

        # --- 1. พยายามหา Row/Col ---
        try:
            # ลองแบบส่ง Event (มาตรฐานใหม่)
            row = self.results_sheet.identify_row(event)
        except:
            try:
                # ลองแบบส่งพิกัด y (มาตรฐานเก่า/ภายใน)
                if hasattr(self.results_sheet, 'MT'):
                    row = self.results_sheet.MT.identify_row(y=event.y)
                else:
                    row = self.results_sheet.identify_row(event.y)
            except: pass

        try:
            # ลองหา Column
            if hasattr(self.results_sheet, 'identify_column'):
                col = self.results_sheet.identify_column(event)
            elif hasattr(self.results_sheet, 'identify_col'):
                col = self.results_sheet.identify_col(event)
            elif hasattr(self.results_sheet, 'MT'):
                if hasattr(self.results_sheet.MT, 'identify_col'):
                    col = self.results_sheet.MT.identify_col(x=event.x)
                elif hasattr(self.results_sheet.MT, 'identify_column'):
                    col = self.results_sheet.MT.identify_column(x=event.x)
        except: pass

        if row is None or col is None: return
        if not self.last_payroll_results or row >= len(self.last_payroll_results): return

        # --- 2. ตรวจสอบชื่อหัวตารางและเปิด Popup ---
        try:
            headers = self.results_sheet.headers()
            if col < len(headers):
                clicked_header_name = headers[col]
                # print(f"   > Clicked Header: {clicked_header_name}")
                
                # กรณี 1: ค่าเที่ยวรถ
                if "ค่าเที่ยว" in clicked_header_name:
                    self._show_driving_details_popup(row)
                
                # กรณี 2: OT หรือ เบี้ยขยัน (เปิด Popup เวลาทำงาน)
                elif "OT" in clicked_header_name or "เบี้ยขยัน" in clicked_header_name:
                    self._show_attendance_details_popup(row, clicked_header_name)
                    
        except Exception as e:
            print(f"Error handling double click: {e}")

    def _show_driving_details_popup(self, row_index):
        """(คงเดิม) แสดง Popup รายละเอียดเที่ยวรถ"""
        payroll_data = self.last_payroll_results[row_index]
        emp_id = payroll_data['emp_id']
        emp_name = payroll_data.get('name', '-')
        
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
        except: return

        details = hr_database.get_driving_details_range(emp_id, start_date, end_date)
        if not details:
            messagebox.showinfo("ไม่มีข้อมูล", f"ไม่พบรายการเที่ยวรถของ\n{emp_name}")
            return

        win = tk.Toplevel(self)
        win.title(f"🚚 รายละเอียดเที่ยวรถ - {emp_name}")
        win.geometry("700x400")
        
        ttk.Label(win, text=f"ประวัติเที่ยวรถ: {emp_name}", font=("", 12, "bold")).pack(pady=10)
        ttk.Label(win, text=f"ช่วงวันที่: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", foreground="gray").pack(pady=(0,10))

        cols = ("date", "license", "type", "driver", "cost", "service", "total")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
        
        tree.heading("date", text="วันที่")
        tree.heading("license", text="ทะเบียน")
        tree.heading("type", text="ประเภท")
        tree.heading("driver", text="คนขับ")
        tree.heading("cost", text="ค่าเที่ยว")
        tree.heading("service", text="ค่าบริการ")
        tree.heading("total", text="รวม")
        
        tree.column("date", width=80, anchor="center")
        tree.column("license", width=100)
        tree.column("type", width=80)
        tree.column("driver", width=120)
        tree.column("cost", width=70, anchor="e")
        tree.column("service", width=70, anchor="e")
        tree.column("total", width=70, anchor="e")
        
        tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        total_sum = 0.0
        for item in details:
            d_str = item['work_date'].strftime("%d/%m/%Y")
            cost = float(item.get('trip_cost', 0))
            serv = float(item.get('service_fee', 0))
            total = cost + serv
            total_sum += total
            
            tree.insert("", "end", values=(
                d_str, item.get('license_plate', '-'), item.get('car_type', '-'),
                item.get('driver_name', '-'), f"{cost:,.2f}", f"{serv:,.2f}", f"{total:,.2f}"
            ))
            
        ttk.Label(win, text=f"รวมยอดสุทธิ: {total_sum:,.2f} บาท", font=("", 11, "bold"), foreground="green").pack(pady=10, anchor="e", padx=20)
    
    def _show_attendance_details_popup(self, row_index, title_prefix="Attendance"):
        """(ใหม่) แสดง Popup รายละเอียดการเข้างาน/OT/เบี้ยขยัน"""
        # 1. เตรียมข้อมูล
        payroll_data = self.last_payroll_results[row_index]
        emp_id = payroll_data['emp_id']
        emp_name = payroll_data.get('name', '-')
        
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
        except: return

        # 2. ดึงข้อมูลจาก DB (ต้องสร้างฟังก์ชันนี้ใน hr_database.py ด้วย)
        # เราจะดึงรายการรายวันในช่วงวันที่เลือกมาแสดง
        daily_records = hr_database.get_daily_records_range(emp_id, start_date, end_date)
        
        if not daily_records:
            messagebox.showinfo("ไม่มีข้อมูล", f"ไม่พบประวัติเวลาทำงานของ\n{emp_name}")
            return

        # 3. สร้างหน้าต่าง
        win = tk.Toplevel(self)
        win.title(f"📅 รายละเอียดเวลาทำงาน - {emp_name}")
        win.geometry("800x500")
        
        header_text = f"ประวัติการเข้างานและ OT: {emp_name}"
        if "เบี้ยขยัน" in title_prefix:
            header_text += " (ตรวจสอบเงื่อนไขเบี้ยขยัน)"
            
        ttk.Label(win, text=header_text, font=("", 12, "bold")).pack(pady=10)
        ttk.Label(win, text=f"ช่วงวันที่: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", foreground="gray").pack(pady=(0,10))

        # ตาราง
        cols = ("date", "status", "in", "out", "ot_hours", "late_mins", "is_approved")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
        
        tree.heading("date", text="วันที่")
        tree.heading("status", text="สถานะ")
        tree.heading("in", text="เข้างาน")
        tree.heading("out", text="ออกงาน")
        tree.heading("ot_hours", text="ชม. OT")
        tree.heading("late_mins", text="สาย (นาที)")
        tree.heading("is_approved", text="อนุมัติ OT")
        
        tree.column("date", width=90, anchor="center")
        tree.column("status", width=120)
        tree.column("in", width=70, anchor="center")
        tree.column("out", width=70, anchor="center")
        tree.column("ot_hours", width=70, anchor="center")
        tree.column("late_mins", width=70, anchor="center")
        tree.column("is_approved", width=80, anchor="center")
        
        tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        # ใส่ข้อมูล
        total_ot = 0.0
        total_late = 0
        diligence_fail_count = 0
        
        for item in daily_records:
            d_str = item['work_date'].strftime("%d/%m/%Y")
            status = item.get('status', 'ปกติ')
            t_in = item.get('work_in_time') or "-"
            t_out = item.get('work_out_time') or "-"
            
            # ถ้าไม่มี เวลาเข้าออกปกติ ลองดูจาก OT time หรือ Log (ถ้า DB เก็บแยก)
            # ในที่นี้สมมติว่า function get_daily_records_range เตรียมมาให้แล้ว
            
            ot = float(item.get('ot_hours', 0))
            # late = int(item.get('late_minutes', 0)) # ถ้ามีฟิลด์นี้
            late = 0 # สมมติ 0 ไปก่อนถ้ายังไม่มีฟิลด์จริง
            
            is_appr = "✅" if item.get('is_ot_approved') else "-"
            if ot > 0 and not item.get('is_ot_approved'): is_appr = "รออนุมัติ"
            
            # Highlight แถวที่มีปัญหา (สำหรับเบี้ยขยัน)
            tags = ()
            if "สาย" in status or "ขาด" in status or "ลา" in status:
                tags = ('bad',)
                diligence_fail_count += 1
            
            tree.insert("", "end", values=(
                d_str, status, t_in, t_out, f"{ot:.2f}", late, is_appr
            ), tags=tags)
            
            total_ot += ot
            total_late += late

        tree.tag_configure('bad', foreground='red')

        # สรุปท้ายตาราง
        summary_frame = ttk.Frame(win)
        summary_frame.pack(fill="x", padx=20, pady=15) # เพิ่ม pady ให้สวยงาม
        
        # ซ้าย: รวม OT
        ttk.Label(summary_frame, text=f"รวม OT: {total_ot:.2f} ชม.", font=("Segoe UI", 11, "bold")).pack(side="left")
        
        # ขวา: สถานะเบี้ยขยัน + Streak
        if "เบี้ยขยัน" in title_prefix:
            # 1. แสดงสถานะผ่าน/ไม่ผ่าน
            result_text = "✅ ผ่านเกณฑ์เดือนนี้" if diligence_fail_count == 0 else f"❌ ไม่ผ่าน ({diligence_fail_count} วัน)"
            color = "green" if diligence_fail_count == 0 else "red"
            
            ttk.Label(summary_frame, text=result_text, font=("Segoe UI", 11, "bold"), foreground=color).pack(side="right", padx=(10, 0))
            
            # 2. (เพิ่มใหม่) แสดงค่า Streak และจำนวนเงิน
            # ถ้าผ่านเดือนนี้ ให้ไปคำนวณ Streak มาโชว์
            if diligence_fail_count == 0:
                try:
                    # ดึงเดือน/ปี จาก start_date ที่เลือกอยู่
                    m = start_date.month
                    y = start_date.year
                    
                    # เรียกฟังก์ชันใหม่จาก Database
                    streak, money = hr_database.get_diligence_streak_info(emp_id, m, y)
                    
                    streak_text = f"🔥 ทำดีต่อเนื่อง: {streak} เดือน (รับ {money:,.0f} บ.)"
                    ttk.Label(summary_frame, text=streak_text, font=("Segoe UI", 11, "bold"), foreground="#FF8C00").pack(side="right", padx=10)
                    
                except Exception as e:
                    print(f"Cannot get streak: {e}")

    def _export_pnd1k_excel(self):
        """ออกรายงาน ภ.ง.ด. 1ก (รายปี) เป็น Excel"""
        
        # 1. ถามปี พ.ศ.
        current_year_be = datetime.now().year + 543
        year_str = simpledialog.askstring("เลือกปีภาษี", f"กรุณากรอกปี พ.ศ. ที่ต้องการออกรายงาน (เช่น {current_year_be}):", initialvalue=str(current_year_be))
        
        if not year_str or not year_str.isdigit(): return
        year_be = int(year_str)
        year_ce = year_be - 543 # แปลงกลับเป็น ค.ศ. เพื่อ query DB

        # 2. ดึงข้อมูล
        data_list = hr_database.get_annual_pnd1k_data(year_ce)
        
        if not data_list:
            messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติการจ่ายเงินเดือนในปี {year_be}")
            return

        # 3. เลือกที่เซฟ
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"PND1K_Year_{year_be}.xlsx",
            title=f"บันทึก ภ.ง.ด. 1ก ปี {year_be}"
        )
        if not save_path: return

        try:
            # 4. เตรียมข้อมูลลง Excel
            rows = []
            seq = 1
            total_inc = 0
            total_tax = 0
            total_sso = 0
            total_fund = 0
            
            for item in data_list:
                inc = float(item['annual_income'] or 0)
                tax = float(item['annual_tax'] or 0)
                sso = float(item['annual_sso'] or 0)
                fund = float(item['annual_fund'] or 0)
                
                total_inc += inc
                total_tax += tax
                total_sso += sso
                total_fund += fund
                
                rows.append({
                    "ลำดับ": seq,
                    "เลขประจำตัวประชาชน": item.get('id_card', ''),
                    "ชื่อ": item.get('fname', ''),
                    "นามสกุล": item.get('lname', ''),
                    "ที่อยู่": item.get('address', ''),
                    "วันเดือนปีที่จ่าย": "ตลอดปีภาษี",
                    "จำนวนเงินได้ที่จ่าย (ทั้งปี)": inc,
                    "ภาษีที่หักนำส่ง (ทั้งปี)": tax,
                    "ประกันสังคม (ทั้งปี)": sso,
                    "กองทุนสำรองฯ (ทั้งปี)": fund,
                    "เงื่อนไข": "1"
                })
                seq += 1
                
            df = pd.DataFrame(rows)
            
            # เพิ่มแถว Total
            total_row = {
                "ลำดับ": "", "เลขประจำตัวประชาชน": "", "ชื่อ": ">>> รวมทั้งสิ้น <<<", "นามสกุล": "", "ที่อยู่": "",
                "วันเดือนปีที่จ่าย": "",
                "จำนวนเงินได้ที่จ่าย (ทั้งปี)": total_inc,
                "ภาษีที่หักนำส่ง (ทั้งปี)": total_tax,
                "ประกันสังคม (ทั้งปี)": total_sso,
                "กองทุนสำรองฯ (ทั้งปี)": total_fund,
                "เงื่อนไข": ""
            }
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            # บันทึก
            df.to_excel(save_path, index=False)
            
            if messagebox.askyesno("สำเร็จ", f"บันทึกเรียบร้อยแล้วที่:\n{save_path}\n\nต้องการเปิดไฟล์เลยหรือไม่?"):
                os.startfile(save_path)

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")
     
    # --- ส่วน Logic ---
    def _open_history_window(self):
        """เปิดหน้าต่างดูประวัติเงินเดือนย้อนหลัง (V91.0 เพิ่มปุ่มพิมพ์สลิปย้อนหลัง)"""
        win = tk.Toplevel(self)
        win.title("📜 ประวัติการจ่ายเงินเดือน (Payroll History)")
        win.geometry("1400x700") 
        
        # --- Filter Frame ---
        top_frame = ttk.Frame(win, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="เลือกงวดที่จ่าย:").pack(side="left")
        
        # ปี
        current_year = datetime.now().year + 543
        years = [str(y) for y in range(current_year, current_year-5, -1)]
        cb_year = ttk.Combobox(top_frame, values=years, width=6, state="readonly")
        cb_year.set(current_year)
        cb_year.pack(side="left", padx=5)
        
        # เดือน
        months = list(self.THAI_MONTHS.values())
        cb_month = ttk.Combobox(top_frame, values=months, width=12, state="readonly")
        cb_month.set(self.THAI_MONTHS[datetime.now().month])
        cb_month.pack(side="left", padx=5)

        # เลือกรอบวิก
        ttk.Label(top_frame, text="รอบ:").pack(side="left", padx=(10, 2))
        cb_period = ttk.Combobox(top_frame, values=["ทั้งเดือน", "วิก 1 (1-15)", "วิก 2 (16-สิ้นเดือน)"], width=15, state="readonly")
        cb_period.set("ทั้งเดือน")
        cb_period.pack(side="left", padx=5)
        
        # 🛠️ ตัวแปรเก็บข้อมูลประวัติดิบ เพื่อนำไปส่งให้ PDF
        win.current_records = []

        # --- Sheet แสดงข้อมูล ---
        sheet_frame = ttk.Frame(win)
        sheet_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        history_sheet = Sheet(sheet_frame, headers=None, theme="light blue")
        history_sheet.pack(fill="both", expand=True)
        history_sheet.enable_bindings("single_select", "row_select", "column_width_resize", "arrowkeys", "copy")

        # 🛠️ [NEW] ฟังก์ชันสำหรับพิมพ์สลิปจากหน้าประวัติโดยตรง
        def print_history_payslip():
            if not win.current_records:
                messagebox.showwarning("เตือน", "ไม่มีข้อมูลประวัติ กรุณาค้นหาก่อนครับ", parent=win)
                return

            selected_indexes = history_sheet.get_selected_rows(return_tuple=True)
            
            # ถ้าเลือกแถวสุดท้ายที่เป็น TOTAL ให้ตัดออก
            if selected_indexes and selected_indexes[-1] == len(win.current_records):
                selected_indexes = selected_indexes[:-1]

            if not selected_indexes:
                if not messagebox.askyesno("พิมพ์ทั้งหมด?", "คุณไม่ได้เลือกพนักงาน\nต้องการพิมพ์สลิปของ 'ทุกคน' ในงวดนี้หรือไม่?", parent=win):
                    return
                selected_indexes = range(len(win.current_records))
            
            data_to_print = []
            for i in selected_indexes:
                r = win.current_records[i]
                
                # แปลงข้อมูลจาก DB ให้ตรงกับรูปแบบที่ _generate_pdf ต้องการ
                mapped_data = {
                    "emp_id": r["emp_id"],
                    "name": f"{r.get('fname','')} {r.get('lname','')}",
                    "position": r.get('position', '-'),
                    "department": r.get('department', '-'),
                    "base_salary": float(r.get("base_salary", 0) or 0),
                    "position_allowance": float(r.get("position_allowance", 0) or 0),
                    "ot": float(r.get("ot_pay", 0) or 0),
                    "commission": float(r.get("commission", 0) or 0),
                    "bonus": float(r.get("bonus", 0) or 0),
                    "incentive": float(r.get("incentive", 0) or 0),
                    "diligence": float(r.get("diligence", 0) or 0),
                    "driving_allowance": float(r.get("driving_allowance", 0) or 0),
                    "other_income": float(r.get("other_income", 0) or 0),
                    "total_income": float(r.get("total_income", 0) or 0),
                    "sso": float(r.get("sso_deduct", 0) or 0),
                    "pnd1": float(r.get("tax_deduct", 0) or 0), 
                    "pnd3": 0.0, 
                    "provident_fund": float(r.get("provident_fund", 0) or 0),
                    "loan": float(r.get("loan_deduct", 0) or 0),
                    "late_deduct": float(r.get("late_deduct", 0) or 0),
                    "other_deduct": float(r.get("other_deduct", 0) or 0),
                    "total_deduct": float(r.get("total_deduct", 0) or 0),
                    "net_salary": float(r.get("net_salary", 0) or 0),
                    "remark": str(r.get("remark", ""))
                }
                data_to_print.append(mapped_data)

            # หาวันที่จ่ายเงินจากประวัติ
            pd_date = win.current_records[0].get('payment_date')
            pay_date_str = f"{pd_date.day}/{pd_date.month}/{pd_date.year+543}" if pd_date else "-"

            filename_prefix = f"Payslip_History_{data_to_print[0]['emp_id']}" if len(data_to_print)==1 else "Payslip_History_All"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile=f"{filename_prefix}.pdf",
                title="บันทึกสลิปเงินเดือนย้อนหลัง"
            )
            if not save_path: return
            
            try:
                self._generate_pdf(data_to_print, save_path, pay_date_str)
                messagebox.showinfo("สำเร็จ", f"สร้างสลิปเงินเดือนย้อนหลังเรียบร้อยแล้ว:\n{save_path}", parent=win)
                os.startfile(save_path)
            except Exception as e:
                messagebox.showerror("PDF Error", f"เกิดข้อผิดพลาดในการสร้างสลิปย้อนหลัง:\n{e}", parent=win)

        # ฟังก์ชันโหลดข้อมูล
        def load_history():
            try:
                y_be = int(cb_year.get())
                y_ce = y_be - 543
                m_name = cb_month.get()
                m_int = self.MONTH_TO_INT[m_name]
                
                p_text = cb_period.get()
                calc_mode = 0
                if "วิก 1" in p_text: calc_mode = 1
                elif "วิก 2" in p_text: calc_mode = 2
            except: return

            records = hr_database.get_monthly_payroll_records(m_int, y_ce, calc_mode)
            
            win.current_records = records # 🛠️ เก็บประวัติไว้ใช้ตอนพิมพ์สลิป
            
            if not records:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติการจ่ายเงินของ {m_name} {y_be} ({p_text})")
                history_sheet.set_sheet_data([])
                return

            t_base = 0.0; t_pos = 0.0; t_ot = 0.0; t_comm = 0.0; t_incen = 0.0
            t_dil = 0.0; t_bonus = 0.0; t_other_inc = 0.0; t_drive = 0.0; t_total_inc = 0.0
            t_sso = 0.0; t_tax = 0.0; t_prov = 0.0; t_loan = 0.0; t_late = 0.0
            t_other_ded = 0.0; t_total_ded = 0.0; t_net = 0.0

            sheet_data = []
            pay_date_display = "-"
            
            for r in records:
                if pay_date_display == "-" and r.get('payment_date'):
                    pd_date = r['payment_date'] 
                    pay_date_display = f"{pd_date.day}/{pd_date.month}/{pd_date.year+543}"

                t_base += float(r.get('base_salary', 0))
                t_pos += float(r.get('position_allowance', 0))
                t_ot += float(r.get('ot_pay', 0))
                t_comm += float(r.get('commission', 0))
                t_incen += float(r.get('incentive', 0))
                t_dil += float(r.get('diligence', 0))
                t_bonus += float(r.get('bonus', 0))
                t_other_inc += float(r.get('other_income', 0))
                t_drive += float(r.get('driving_allowance', 0))
                t_total_inc += float(r.get('total_income', 0))
                
                t_sso += float(r.get('sso_deduct', 0))
                t_tax += float(r.get('tax_deduct', 0))
                t_prov += float(r.get('provident_fund', 0))
                t_loan += float(r.get('loan_deduct', 0))
                t_late += float(r.get('late_deduct', 0))
                t_other_ded += float(r.get('other_deduct', 0))
                t_total_ded += float(r.get('total_deduct', 0))
                t_net += float(r.get('net_salary', 0))

                fullname = f"{r.get('fname','')} {r.get('lname','')}"
                row = [
                    r['emp_id'], fullname,
                    f"{r.get('base_salary', 0):,.2f}", f"{r.get('position_allowance', 0):,.2f}",
                    f"{r.get('ot_pay', 0):,.2f}", f"{r.get('commission', 0):,.2f}", 
                    f"{r.get('incentive', 0):,.2f}", f"{r.get('diligence', 0):,.2f}",
                    f"{r.get('bonus', 0):,.2f}", f"{r.get('other_income', 0):,.2f}", 
                    f"{r.get('driving_allowance', 0):,.2f}", 
                    f"{r.get('total_income', 0):,.2f}", 
                    f"{r.get('sso_deduct', 0):,.2f}", f"{r.get('tax_deduct', 0):,.2f}", 
                    f"{r.get('provident_fund', 0):,.2f}", f"{r.get('loan_deduct', 0):,.2f}", 
                    f"{r.get('late_deduct', 0):,.2f}", f"{r.get('other_deduct', 0):,.2f}",
                    f"{r.get('total_deduct', 0):,.2f}", 
                    f"{r.get('net_salary', 0):,.2f}",
                    r.get('remark', '') 
                ]
                sheet_data.append(row)
            
            if sheet_data:
                summary_row = [
                    "TOTAL", "รวมทั้งสิ้น",
                    f"{t_base:,.2f}", f"{t_pos:,.2f}",
                    f"{t_ot:,.2f}", f"{t_comm:,.2f}", 
                    f"{t_incen:,.2f}", f"{t_dil:,.2f}",
                    f"{t_bonus:,.2f}", f"{t_other_inc:,.2f}", 
                    f"{t_drive:,.2f}", 
                    f"{t_total_inc:,.2f}", 
                    f"{t_sso:,.2f}", f"{t_tax:,.2f}", 
                    f"{t_prov:,.2f}", f"{t_loan:,.2f}", 
                    f"{t_late:,.2f}", f"{t_other_ded:,.2f}",
                    f"{t_total_ded:,.2f}", 
                    f"{t_net:,.2f}",
                    ""
                ]
                sheet_data.append(summary_row)

            headers = [
                "รหัส", "ชื่อ-สกุล", "เงินเดือน", "ค่าตำแหน่ง", "OT", "คอมมิชชั่น", 
                "Incentive", "เบี้ยขยัน", "โบนัส", "อื่นๆ(รับ)", "ค่าเที่ยว", "รวมรับ", 
                "ประกันสังคม", "ภาษี", "กองทุนฯ", "เงินกู้", "มาสาย/ลา", "อื่นๆ(หัก)", "รวมหัก", 
                "สุทธิ", "หมายเหตุ"
            ]
            history_sheet.headers(headers)
            history_sheet.set_sheet_data(sheet_data)
            
            history_sheet.highlight_columns(columns=list(range(2, 12)), bg="#e6f7ff", fg="black") 
            history_sheet.highlight_columns(columns=list(range(12, 19)), bg="#fff7e6", fg="black") 
            history_sheet.highlight_columns(columns=[19], bg="#ffffcc", fg="black") 
            if sheet_data: 
                history_sheet.highlight_rows(rows=[len(sheet_data)-1], bg="#ccffcc", fg="black") 
            
            win.title(f"📜 ประวัติการจ่ายเงินเดือน - งวด {m_name} {y_be} | {p_text} (จ่ายเมื่อ: {pay_date_display})")

        ttk.Button(top_frame, text="🔍 ค้นหาประวัติ", command=load_history).pack(side="left", padx=10)
        
        # 🛠️ [NEW] ปุ่มพิมพ์สลิปย้อนหลัง นำไปวางข้างๆ ปุ่มค้นหาประวัติ
        ttk.Button(top_frame, text="🖨️ พิมพ์สลิปย้อนหลัง (PDF)", command=print_history_payslip, style="Primary.TButton").pack(side="left", padx=10)
        
        load_history()

    def _open_input_popup(self, event):
        """
        (ฉบับแก้ไข V90.0 - แยกคอลัมน์ วิก 1 และ วิก 2)
        """
        selection = self.input_tree.selection()
        if not selection: return
        item_id = selection[0]
        tree_values = self.input_tree.item(item_id, "values")
        
        emp_id_real = str(tree_values[1]) 
        emp_name = str(tree_values[2])
        
        # ดึงข้อมูลเดิมจาก Memory (ถ้าไม่มีก็สร้างใหม่)
        curr_data = self.payroll_inputs.get(emp_id_real, {})

        popup = tk.Toplevel(self)
        popup.title(f"บันทึกรายรับ/รายจ่าย - {emp_name}")
        popup.geometry("600x750") # 🛠️ ขยายหน้าต่างให้กว้างขึ้น
        popup.transient(self)
        popup.grab_set()
        
        main_frame = ttk.Frame(popup, padding=20)
        main_frame.pack(fill="both", expand=True)
        
        entries_1 = {} # เก็บช่องกรอกวิก 1
        entries_2 = {} # เก็บช่องกรอกวิก 2

        # --- หัวตาราง ---
        header_f = ttk.Frame(main_frame)
        header_f.pack(fill="x", pady=(0, 10))
        ttk.Label(header_f, text="รายการ (ระบุตัวเลข)", width=22, font=("", 10, "bold")).pack(side="left")
        ttk.Label(header_f, text="วิก 1 (วันที่ 1-15)", width=18, font=("", 10, "bold")).pack(side="left", padx=5)
        ttk.Label(header_f, text="วิก 2 (16-สิ้นเดือน)", width=18, font=("", 10, "bold")).pack(side="left", padx=5)

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=5)

        # --- ส่วนรายได้ ---
        ttk.Label(main_frame, text="รายได้ (Addition)", font=("", 10, "bold"), foreground="green").pack(anchor="w", pady=(5,5))
        fields_inc = [
            ("ค่าล่วงเวลา (OT)", "ot"), ("คอมมิชชั่น", "commission"), 
            ("Incentive", "incentive"), ("เบี้ยขยัน (Manual)", "diligence"), 
            ("โบนัส", "bonus"), ("ค่าที่พักนอก", "housing_income"),
            ("เงินได้อื่นๆ (รับ)", "other_income")
        ]

        for label, key in fields_inc:
            f = ttk.Frame(main_frame); f.pack(fill="x", pady=2)
            ttk.Label(f, text=label, width=22).pack(side="left")
            
            # วิก 1
            ent1 = ttk.Entry(f, width=15); ent1.pack(side="left", padx=5)
            val1 = curr_data.get(f"{key}_1", 0.0)
            ent1.insert(0, f"{val1:.2f}" if val1 != 0 else "0")
            entries_1[key] = ent1
            
            # วิก 2
            ent2 = ttk.Entry(f, width=15); ent2.pack(side="left", padx=5)
            val2 = curr_data.get(f"{key}_2", 0.0)
            ent2.insert(0, f"{val2:.2f}" if val2 != 0 else "0")
            entries_2[key] = ent2

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=10)
        
        # --- ส่วนเงินหัก ---
        ttk.Label(main_frame, text="รายการหัก (Deduction)", font=("", 10, "bold"), foreground="red").pack(anchor="w", pady=(0,5))
        fields_ded = [
            ("ภาษี (Manual)", "tax"), ("กองทุนสำรองฯ", "provident_fund"), 
            ("หักเงินกู้/ยืม", "loan"), ("ค่าที่พักใน", "housing_deduct"), ("หักอื่นๆ", "other_deduct")
        ]

        for label, key in fields_ded:
            f = ttk.Frame(main_frame); f.pack(fill="x", pady=2)
            ttk.Label(f, text=label, width=22).pack(side="left")
            
            # วิก 1
            ent1 = ttk.Entry(f, width=15); ent1.pack(side="left", padx=5)
            val1 = curr_data.get(f"{key}_1", 0.0)
            ent1.insert(0, f"{val1:.2f}" if val1 != 0 else "0")
            entries_1[key] = ent1
            
            # วิก 2
            ent2 = ttk.Entry(f, width=15); ent2.pack(side="left", padx=5)
            val2 = curr_data.get(f"{key}_2", 0.0)
            ent2.insert(0, f"{val2:.2f}" if val2 != 0 else "0")
            entries_2[key] = ent2

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=10)
        
        # --- อื่นๆ ---
        ttk.Label(main_frame, text="อื่นๆ (Others)", font=("", 10, "bold"), foreground="blue").pack(anchor="w", pady=(0,5))
        
        f_remark = ttk.Frame(main_frame); f_remark.pack(fill="x", pady=2)
        ttk.Label(f_remark, text="หมายเหตุ (ในตาราง)", width=22).pack(side="left", anchor="n")
        remark_entry = tk.Text(f_remark, width=40, height=2, font=("Segoe UI", 10)) 
        remark_entry.pack(side="left")
        remark_entry.insert("1.0", curr_data.get("remark", "")) 

        f_note = ttk.Frame(main_frame); f_note.pack(fill="x", pady=2)
        ttk.Label(f_note, text="โน้ตภายใน (ดูเอง)", width=22).pack(side="left", anchor="n")
        note_entry = tk.Text(f_note, width=40, height=2, font=("Segoe UI", 10)) 
        note_entry.pack(side="left")
        note_entry.insert("1.0", curr_data.get("note", ""))

        def save_popup():
            try:
                new_data = {}
                # บันทึกของวิก 1 และวิก 2 แยกรหัสกัน (เช่น ot_1, ot_2)
                all_keys = [k for _, k in fields_inc] + [k for _, k in fields_ded]
                for k in all_keys:
                    new_data[f"{k}_1"] = float(entries_1[k].get().replace(',', '') or 0.0)
                    new_data[f"{k}_2"] = float(entries_2[k].get().replace(',', '') or 0.0)
                
                # บันทึกหมายเหตุ
                new_data['remark'] = remark_entry.get("1.0", tk.END).strip() 
                new_data['note'] = note_entry.get("1.0", tk.END).strip()
                
                self.payroll_inputs[emp_id_real] = new_data
                self.input_tree.set(item_id, column="status", value="✅ บันทึกแล้ว")
                popup.destroy()
                messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลแล้ว กรุณากดปุ่ม 'ประมวลผล' เพื่ออัปเดตยอด")
            except: messagebox.showerror("Error", "กรุณากรอกตัวเลขให้ถูกต้อง (ห้ามมีตัวหนังสือในช่องยอดเงิน)")

        ttk.Button(main_frame, text="💾 บันทึกยอด", command=save_popup).pack(pady=20)

    def _run_payroll_calculation(self):
        """
        ฟังก์ชันหลักคำนวณเงินเดือน - ฉบับสมบูรณ์ V90.0 (รองรับแบ่งวิก 1 และ วิก 2 อัตโนมัติ)
        """
        # --- 1. ตรวจสอบวันที่ และ เช็คโหมดวิก (Period Mode) ---
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            if not start_date or not end_date: 
                messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกวันที่ให้ครบ")
                return
            current_month = start_date.month
            current_year = start_date.year 
            
            # 🛠️ [NEW] เช็คว่ากำลังคิดเงินเดือนของช่วงไหน
            calc_mode = 0 # 0 = เต็มเดือน (เอาวิก 1 + วิก 2), 1 = วิก 1 (เอาแค่ _1), 2 = วิก 2 (เอาแค่ _2)
            if start_date.day <= 15 and end_date.day <= 15:
                calc_mode = 1 # เลือกคำนวณ 1-15
            elif start_date.day >= 16:
                calc_mode = 2 # เลือกคำนวณ 16-สิ้นเดือน
            else:
                calc_mode = 0 # เลือกคำนวณคร่อม 2 รอบ (เช่น 1-31)
                
        except Exception as e:
            print(f"Date Error: {e}")
            return

        # 2. กวาดรายชื่อพนักงานที่ถูกเลือก (ที่ติ๊ก ☑)
        selected_iids = []
        def collect_ids(tree_widget):
            ids = []
            if tree_widget is not None:
                for iid in tree_widget.get_children():
                    vals = tree_widget.item(iid, "values")
                    if vals and vals[0] == "☑":  
                        ids.append(iid)
            return ids

        if hasattr(self, 'tree_reg') and hasattr(self, 'tree_cont'):
            selected_iids.extend(collect_ids(self.tree_reg))
            selected_iids.extend(collect_ids(self.tree_cont))
        elif hasattr(self, 'input_tree'):
            selected_iids.extend(collect_ids(self.input_tree))
        
        if not selected_iids:
            messagebox.showwarning("เตือน", "กรุณาเลือกพนักงานที่ต้องการคำนวณอย่างน้อย 1 คน")
            return

        # --- 3. เตรียมโครงสร้างเก็บผลรวมและ Config ---
        try:
            allowance_settings = hr_database.load_allowance_settings()
            taxable_map = { item['name']: item['is_taxable'] for item in allowance_settings }
        except: taxable_map = {} 
        
        self.last_payroll_results = []
        sheet_data = []
        
        total_sum = {
            "base_salary": 0.0, "position": 0.0, "ot": 0.0, 
            "bonus": 0.0, "commission": 0.0, "other_income": 0.0,
            "welfare_taxable": 0.0, "welfare_nontaxable": 0.0,
            "total_income": 0.0, "net_salary": 0.0,
            "sso": 0.0, "tax": 0.0, "pnd1": 0.0, "pnd3": 0.0, 
            "provident_fund": 0.0, "loan": 0.0, "late": 0.0, 
            "other_deduct": 0.0, "total_deduct": 0.0,
            "driving_allowance": 0.0
        }

        # --- 4. เริ่ม Loop คำนวณรายคน ---
        current_tree = self.input_tree if hasattr(self, 'input_tree') else self.tree_reg

        for iid in selected_iids:
            node_tree = self.tree_reg if iid in self.tree_reg.get_children() else self.tree_cont
            real_emp_id = node_tree.item(iid, "values")[1]
            
            # 🛠️ [NEW] ดึงค่าจาก Popup และกรองตาม calc_mode ที่เลือกไว้
            raw_input = self.payroll_inputs.get(real_emp_id, {})
            user_in = {}
            user_in['remark'] = raw_input.get('remark', '')
            user_in['note'] = raw_input.get('note', '')
            
            # รายการฟิลด์ตัวเลขทั้งหมด
            numeric_fields = ['ot', 'commission', 'incentive', 'diligence', 'bonus', 'housing_income', 'other_income', 'tax', 'provident_fund', 'loan', 'housing_deduct', 'other_deduct']
            
            for field in numeric_fields:
                val1 = float(raw_input.get(f"{field}_1", 0.0))
                val2 = float(raw_input.get(f"{field}_2", 0.0))
                
                if calc_mode == 1:   # คิดเงินวิก 1
                    user_in[field] = val1
                elif calc_mode == 2: # คิดเงินวิก 2
                    user_in[field] = val2
                else:                # คิดเงินเต็มเดือน
                    user_in[field] = val1 + val2
            
            # เรียกคำนวณ Backend 
            res = hr_database.calculate_payroll_for_employee(real_emp_id, start_date, end_date, user_in)
            
            if res:
                emp_info = hr_database.load_single_employee(real_emp_id)
                emp_name = f"{emp_info.get('fname','')} {emp_info.get('lname','')}"
                res['name'] = emp_name
                
                # ดึงข้อมูลประกอบ
                if emp_info:
                    res['position'] = emp_info.get('position', '-')
                    res['department'] = emp_info.get('department', '-')
                    res['bank_account'] = emp_info.get('bank_account_no') or emp_info.get('account') or '-'
                    res['bank_name'] = emp_info.get('bank_name') or emp_info.get('bank') or '-'
                    
                    # คำนวณสวัสดิการ (Welfare)
                    w_tax_sum = 0.0; w_nontax_sum = 0.0
                    w_opts = emp_info.get('welfare_options', [])
                    w_flags = emp_info.get('welfare', [])
                    w_amts = emp_info.get('welfare_amounts', [])
                    for idx, amt_str in enumerate(w_amts):
                        if idx < len(w_flags) and w_flags[idx]:
                            amt = float(amt_str or 0)
                            w_name = w_opts[idx] if idx < len(w_opts) else ""
                            if taxable_map.get(w_name, True): w_tax_sum += amt
                            else: w_nontax_sum += amt
                
                # สะสมยอดลง Summary
                total_sum['base_salary'] += res['base_salary']
                total_sum['position'] += res.get('position_allowance', 0)
                total_sum['ot'] += res['ot']
                total_sum['commission'] += res['commission']
                total_sum['bonus'] += res['bonus']
                total_sum['other_income'] += res['other_income']
                total_sum['driving_allowance'] += res.get('driving_allowance', 0)
                total_sum['total_income'] += res['total_income']
                total_sum['sso'] += res['sso']
                total_sum['tax'] += res['tax']
                total_sum['pnd1'] += res.get('pnd1', 0)
                total_sum['pnd3'] += res.get('pnd3', 0)
                total_sum['provident_fund'] += res['provident_fund']
                total_sum['loan'] += res['loan']
                total_sum['late'] += res['late_deduct']
                total_sum['total_deduct'] += res['total_deduct']
                total_sum['net_salary'] += res['net_salary']

                # จัดข้อมูลลงตารางผลลัพธ์
                # 🛠️ อื่นๆ(รับ) ในตาราง = Other Income จาก Popup + สวัสดิการ
                display_other_sheet = res['other_income']
                
                res['remark'] = user_in.get('remark', '')
                res['note'] = user_in.get('note', '')

                row_data = [
                    real_emp_id, emp_name,
                    f"{res['base_salary']:,.2f}", f"{res.get('position_allowance',0):,.2f}",
                    f"{res['ot']:,.2f}", f"{res['commission']:,.2f}", 
                    f"{res.get('incentive',0):,.2f}", f"{res.get('diligence',0):,.2f}",
                    f"{res['bonus']:,.2f}", f"{display_other_sheet:,.2f}", 
                    f"{res.get('driving_allowance',0):,.2f}", 
                    f"{res['total_income']:,.2f}", 
                    f"{res['sso']:,.2f}", 
                    f"{res.get('pnd1',0):,.2f}", 
                    f"{res.get('pnd3',0):,.2f}", 
                    f"{res['provident_fund']:,.2f}", f"{res['loan']:,.2f}", 
                    f"{res['late_deduct']:,.2f}", f"{res['other_deduct']:,.2f}",
                    f"{res['total_deduct']:,.2f}", f"{res['net_salary']:,.2f}",
                    res['remark']  # 🛠️ เอาใส่ไว้ขวาสุด
                ]
                sheet_data.append(row_data)
                self.last_payroll_results.append(res)

        # --- 5. สรุปยอดบรรทัดสุดท้าย (TOTAL) ---
        summary_row = [
            "TOTAL", "รวมทั้งสิ้น",
            f"{total_sum['base_salary']:,.2f}", f"{total_sum['position']:,.2f}",
            f"{total_sum['ot']:,.2f}", f"{total_sum['commission']:,.2f}", 
            "-", "-", f"{total_sum['bonus']:,.2f}",
            f"{(total_sum['other_income']):,.2f}", f"{total_sum['driving_allowance']:,.2f}",
            f"{total_sum['total_income']:,.2f}",
            f"{total_sum['sso']:,.2f}", f"{total_sum['pnd1']:,.2f}", f"{total_sum['pnd3']:,.2f}", 
            f"{total_sum['provident_fund']:,.2f}", f"{total_sum['loan']:,.2f}",
            f"{total_sum['late']:,.2f}", f"{total_sum['other_deduct']:,.2f}",
            f"{total_sum['total_deduct']:,.2f}", f"{total_sum['net_salary']:,.2f}",
            ""  # 🛠️ เพิ่มช่องว่างตรงคอลัมน์หมายเหตุ
        ]
        sheet_data.append(summary_row)

        # --- 6. อัปเดตตาราง Sheet ---
        headers = ["รหัส", "ชื่อ-สกุล", "เงินเดือน", "ค่าตำแหน่ง", "OT", "คอมมิชชั่น", "Incentive", "เบี้ยขยัน", "โบนัส", "อื่นๆ(รับ)", "ค่าเที่ยว", "รวมรับ", "ประกันสังคม", "ภ.ง.ด.1", "ภ.ง.ด.3", "กองทุนฯ", "เงินกู้", "มาสาย/ลา", "อื่นๆ(หัก)", "รวมหัก", "สุทธิ", "หมายเหตุ"] # 🛠️ เพิ่ม "หมายเหตุ" ตรงนี้
        self.results_sheet.headers(headers) 
        self.results_sheet.set_sheet_data(sheet_data)
        
        # ใส่สี Highlight ตาราง
        self.results_sheet.highlight_columns(columns=list(range(2, 11)), bg="#e6f7ff", fg="black") # รายรับสีฟ้า
        self.results_sheet.highlight_columns(columns=list(range(12, 19)), bg="#fff7e6", fg="black") # รายหักสีส้ม
        self.results_sheet.highlight_columns(columns=[20], bg="#ffffcc", fg="black") # สุทธิสีเหลือง
        if sheet_data: self.results_sheet.highlight_rows(rows=[len(sheet_data)-1], bg="#ccffcc", fg="black") # แถวรวมสีเขียว

        # --- 7. เปิดใช้งานปุ่มอื่นๆ ---
        self.export_btn.config(state="normal")
        self.print_btn.config(state="normal")
        self.pnd1_btn.config(state="normal")
        self.pnd3_btn.config(state="normal")
        self.save_db_btn.config(state="normal")
        self.sso_btn.config(state="normal")
        self.email_req_btn.config(state="normal")
        
        self.notebook.select(self.tab2) # กระโดดไปแท็บผลลัพธ์
        messagebox.showinfo("สำเร็จ", f"คำนวณเงินเดือนเรียบร้อยแล้ว {len(self.last_payroll_results)} รายการ")
        
    def _export_payroll_to_excel(self):
        if not self.last_payroll_results: 
            messagebox.showwarning("ไม่มีข้อมูล", "กรุณากดคำนวณเงินเดือนก่อนส่งออกไฟล์")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel", "*.xlsx")],
            title="บันทึกตารางเงินเดือนเป็น Excel"
        )
        if not file_path: return
        
        try:
            # 1. เตรียมข้อมูล
            df = pd.DataFrame(self.last_payroll_results)
            
            # 2. เปลี่ยนชื่อคอลัมน์ให้เป็นภาษาไทย
            cols = {
                "emp_id": "รหัส", "name": "ชื่อ-สกุล",
                "base_salary": "เงินเดือน", "position_allowance": "ค่าตำแหน่ง",
                "ot": "OT", "commission": "คอมมิชชั่น","incentive": "Incentive", "diligence": "เบี้ยขยัน", "bonus": "โบนัส", "other_income": "อื่นๆ(รับ)",
                "driving_allowance": "ค่าเที่ยวรถ",
                "total_income": "รวมรับ",
                "sso": "ประกันสังคม", "provident_fund": "กองทุนสำรองฯ",
                "loan": "เงินกู้", "late_deduct": "ขาด/สาย", "other_deduct": "อื่นๆ(หัก)",
                "total_deduct": "รวมหัก", "net_salary": "สุทธิ",
                "remark": "หมายเหตุ" # 🛠️ [เพิ่มบรรทัดนี้]
            }
            
            # กรองเอาเฉพาะคอลัมน์ที่มีอยู่จริง (ป้องกัน Error)
            valid_cols = [c for c in cols.keys() if c in df.columns]
            df = df[valid_cols]
            df = df.rename(columns=cols)
            
            # 3. บันทึกไฟล์ (จุดที่เคย Error)
            df.to_excel(file_path, index=False)
            
            # 4. แจ้งผลสำเร็จ และถามว่าจะเปิดไฟล์เลยไหม
            if messagebox.askyesno("สำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{file_path}\n\nต้องการเปิดไฟล์เลยหรือไม่?"):
                os.startfile(file_path)
                
        except PermissionError:
            messagebox.showerror("บันทึกไม่ได้", 
                                 f"ไม่สามารถบันทึกไฟล์ได้!\n\nสาเหตุ: ไฟล์ '{os.path.basename(file_path)}' กำลังถูกเปิดใช้งานอยู่\n\nวิธีแก้: กรุณาปิดโปรแกรม Excel แล้วลองใหม่อีกครั้ง")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")
        
    def _save_payroll_to_database(self):
        """บันทึกผลคำนวณ (Confirm Payment) ลงฐานข้อมูล"""
        if not self.last_payroll_results:
            messagebox.showwarning("เตือน", "ไม่มีข้อมูลให้บันทึก กรุณาคำนวณก่อน")
            return
            
        # 1. ดึงเดือน/ปี ที่คำนวณ
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date() # 🛠️ [FIX] เพิ่มบรรทัดนี้เข้ามาครับ!
            
            m_int = start_date.month
            y_ce = start_date.year
            month_name = list(self.THAI_MONTHS.values())[m_int - 1]
        except:
            return

        # 2. ถามยืนยัน + ถามวันที่จ่ายเงิน (สำคัญ!)
        if not messagebox.askyesno("ยืนยันการจ่ายเงิน", 
                                   f"คุณต้องการ 'ปิดงวดและบันทึกข้อมูล' ประจำเดือน {month_name} {y_ce+543}\n"
                                   f"จำนวน {len(self.last_payroll_results)} รายการ ใช่หรือไม่?"):
            return

        # เรียก Popup ถามวันที่ (ที่ทำไว้ในรอบก่อน)
        pay_date_str = self._ask_pay_date() 
        if not pay_date_str: return # ถ้ากดยกเลิก ก็จบ
        
        # แปลง string วันที่ (dd/mm/yyyy พ.ศ.) กลับเป็น date object (ค.ศ.)
        try:
            d, m, y_be = map(int, pay_date_str.split('/'))
            pay_date = datetime(y_be - 543, m, d).date()
        except:
            pay_date = datetime.now().date()

        # 3. บันทึกลง Database
        success_count = 0
        
        # 🛠️ [NEW] หาว่าตอนนี้เรากำลังบันทึกวิกไหน (คำนวณจากวันที่เลือกไว้)
        calc_mode = 0
        if start_date.day <= 15 and end_date.day <= 15: calc_mode = 1
        elif start_date.day >= 16: calc_mode = 2

        for item in self.last_payroll_results:
            # 🛠️ ส่ง calc_mode ไปให้ Database
            ok = hr_database.save_monthly_payroll(item['emp_id'], m_int, y_ce, pay_date, item, calc_mode)
            if ok: success_count += 1
            
        messagebox.showinfo("สำเร็จ", f"บันทึกประวัติการจ่ายเงินเรียบร้อยแล้ว\nจำนวน: {success_count} รายการ\n(วันที่จ่าย: {pay_date_str})")

    # --- (!!! ส่วนใหม่: พิมพ์สลิปเงินเดือน PDF !!!) ---

    def _ask_pay_date(self):
        """สร้าง Popup ถามวันที่จ่ายเงิน (คืนค่าเป็น string 'dd/mm/yyyy')"""
        win = tk.Toplevel(self)
        win.title("ระบุวันที่จ่ายเงิน")
        win.geometry("300x150")
        win.transient(self)
        win.grab_set()
        
        # จัดกึ่งกลางจอ
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

        ttk.Label(win, text="กรุณาระบุวันที่จ่ายเงิน (Payment Date):", font=("Segoe UI", 10)).pack(pady=15)
        
        # ใช้ DateDropdown ที่เราทำไว้
        date_picker = DateDropdown(win, font=("Segoe UI", 10))
        date_picker.pack(pady=5)
        # ตั้งค่าเริ่มต้นเป็นวันนี้
        date_picker.set_date(datetime.now())

        self._temp_pay_date = None

        def on_ok():
            self._temp_pay_date = date_picker.get() # ได้ string dd/mm/yyyy (พ.ศ.)
            win.destroy()

        def on_cancel():
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="ตกลง", command=on_ok, style="Success.TButton", width=10).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="ยกเลิก", command=on_cancel, width=10).pack(side="left", padx=5)

        self.wait_window(win)
        return self._temp_pay_date
    
    def _print_selected_payslip(self):
        # 1. เช็คการเลือก
        selected_indexes = self.results_sheet.get_selected_rows(return_tuple=True)
        
        data_to_print = []
        if not selected_indexes:
            if not messagebox.askyesno("พิมพ์ทั้งหมด?", "คุณไม่ได้เลือกพนักงาน\nต้องการพิมพ์สลิปของ 'ทุกคน' ในรายการหรือไม่?"):
                return
            data_to_print = self.last_payroll_results
            filename_prefix = "Payslip_All"
        else:
            data_to_print = [self.last_payroll_results[i] for i in selected_indexes]
            filename_prefix = f"Payslip_{data_to_print[0]['emp_id']}" if len(data_to_print)==1 else "Payslip_Selected"

        if not data_to_print: return
        
        # --- [NEW] 2. ถามวันที่จ่ายเงิน ---
        pay_date_str = self._ask_pay_date()
        if not pay_date_str: return # ถ้ากดยกเลิก ก็จบการทำงาน
        
        # 3. เลือกที่เซฟไฟล์
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"{filename_prefix}_{datetime.now().strftime('%Y%m%d')}.pdf",
            title="บันทึกสลิปเงินเดือน"
        )
        if not save_path: return
        
        try:
            # --- [NEW] ส่งวันที่จ่ายเงินไปด้วย ---
            self._generate_pdf(data_to_print, save_path, pay_date_str)
            
            messagebox.showinfo("สำเร็จ", f"สร้างไฟล์ PDF เรียบร้อยแล้วที่:\n{save_path}")
            os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("PDF Error", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{e}")

    def _generate_pdf(self, data_list, filepath, pay_date_str):
        """
        สร้างไฟล์ PDF สลิปเงินเดือน (ฉบับแก้ไข V19.8 - Deductions Update)
        - เพิ่มช่อง: 'ค่าที่พักใน' ในส่วนเงินหัก
        - แก้ไขชื่อ: 'หักเงินกู้ยืม' -> 'หักเงินกู้ยืม (กยศ.)'
        - ปรับ Logic การดึงค่าที่พักในจาก Input Popup
        """
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=False)
        
        # --- 1. โหลดฟอนต์ ---
        base_path = os.path.dirname(__file__)
        resource_path = os.path.join(base_path, "resources")
        font_path_reg = os.path.join(resource_path, "THSarabunNew.ttf")
        
        if not os.path.exists(font_path_reg): 
            font_path_reg = os.path.join(base_path, "THSarabunNew.ttf")
            
        if not os.path.exists(font_path_reg):
            raise Exception("ไม่พบไฟล์ฟอนต์ THSarabunNew.ttf")

        pdf.add_font("THSarabun", "", font_path_reg, uni=True)
        pdf.add_font("THSarabun", "B", font_path_reg, uni=True) 

        # --- 2. โหลดโลโก้ ---
        logo_path = os.path.join(base_path, "company_logo.png")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(base_path, "company_logo.jpg")

        # --- Helper: จัดรูปแบบเงิน ---
        def fmt_money(val):
            return f"{val:,.2f}" if isinstance(val, (int, float)) and val > 0 else "-"

        # --- ฟังก์ชันวาดแบบฟอร์ม (Nested Function) ---
        def draw_slip_form(current_data, start_y, copy_label, emp_welfare_map):
            
            # เตรียมข้อมูล "งวดเดือน"
            try:
                s_date = self.start_date_entry.get_date()
                month_th = list(self.THAI_MONTHS.values())[s_date.month - 1]
                period_str = f"{month_th} {s_date.year + 543}"
            except:
                period_str = "-"

            b_name = current_data.get('bank_name', '-')
            b_acc = current_data.get('bank_account', '-')

            # ------------------------------------------------
            # ส่วนหัว (Header)
            # ------------------------------------------------
            if os.path.exists(logo_path):
                pdf.image(logo_path, x=15, y=start_y + 5, w=20)
            
            pdf.set_xy(0, start_y + 5)
            pdf.set_font("THSarabun", "B", 20)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(0, 8, "บริษัท เอไพร์ม พลัส จํากัด", ln=True, align='C')
            pdf.set_font("THSarabun", "B", 16)
            pdf.cell(0, 8, f"ใบจ่ายเงินเดือน (Pay Slip) {copy_label}", ln=True, align='C')

            # (ขวาบน)
            pdf.set_font("THSarabun", "B", 14)
            pdf.set_xy(140, start_y + 6)
            pdf.cell(60, 6, f"เดือน : {period_str}", border=0, align='R')
            pdf.set_xy(140, start_y + 12)
            pdf.cell(60, 6, f"วันที่จ่าย : {pay_date_str}", border=0, align='R')

            # ------------------------------------------------
            # กรอบข้อมูลพนักงาน
            # ------------------------------------------------
            box_top = start_y + 22
            box_h = 24  
            
            pdf.set_draw_color(0)
            pdf.set_line_width(0.1)
            pdf.rect(10, box_top, 190, box_h) 
            
            pdf.line(10, box_top + 8, 200, box_top + 8)
            pdf.line(10, box_top + 16, 200, box_top + 16)
            
            pdf.line(40, box_top, 40, box_top + box_h)
            pdf.line(105, box_top, 105, box_top + box_h) 
            pdf.line(135, box_top, 135, box_top + box_h) 

            pdf.set_font("THSarabun", "", 14)
            
            # บรรทัด 1
            pdf.set_xy(10, box_top + 1); pdf.cell(30, 6, "  รหัสพนักงาน :", border=0)
            pdf.set_xy(40, box_top + 1); pdf.cell(65, 6, f"  {current_data.get('emp_id', '-')}", border=0)
            
            pdf.set_xy(105, box_top + 1); pdf.cell(30, 6, "  ตำแหน่ง :", border=0)
            pdf.set_xy(135, box_top + 1); pdf.cell(65, 6, f"  {current_data.get('position','-')}", border=0)
            
            # บรรทัด 2
            pdf.set_xy(10, box_top + 9); pdf.cell(30, 6, "  ชื่อ - นามสกุล :", border=0)
            display_name = current_data.get('name', '')
            if not display_name:
                display_name = f"{current_data.get('fname', '')} {current_data.get('lname', '')}".strip()
            pdf.set_xy(40, box_top + 9); pdf.cell(65, 6, f"  {display_name}", border=0)
            
            pdf.set_xy(105, box_top + 9); pdf.cell(30, 6, "  แผนก :", border=0)
            pdf.set_xy(135, box_top + 9); pdf.cell(65, 6, f"  {current_data.get('department','-')}", border=0)

            # บรรทัด 3
            pdf.set_xy(10, box_top + 17); pdf.cell(30, 6, "  ธนาคาร :", border=0)
            pdf.set_xy(40, box_top + 17); pdf.cell(65, 6, f"  {b_name}", border=0)

            pdf.set_xy(105, box_top + 17); pdf.cell(30, 6, "  เลขที่บัญชี :", border=0)
            pdf.set_xy(135, box_top + 17); pdf.cell(65, 6, f"  {b_acc}", border=0)

            # ------------------------------------------------
            # ตารางรายการ (Table)
            # ------------------------------------------------
            tbl_top = box_top + box_h + 2 
            row_h = 5.5 
            
            # หัวตาราง
            pdf.rect(10, tbl_top, 95, 8)
            pdf.rect(105, tbl_top, 95, 8)
            pdf.set_font("THSarabun", "B", 16)
            pdf.set_xy(10, tbl_top); pdf.cell(95, 8, "เงินได้ (Earnings)", border=0, align='C')
            pdf.set_xy(105, tbl_top); pdf.cell(95, 8, "เงินหัก (Deductions)", border=0, align='C')

            # ไส้ในตาราง
            body_top = tbl_top + 8 
            max_rows = 10 
            
            # --- 1. คำนวณฝั่งรายได้ (Incomes) ---
            val_wear = emp_welfare_map.get('ค่าเสื่อมรถยนต์', 0.0) 
            val_allowance = emp_welfare_map.get('เบี้ยเลี้ยง', 0.0) 
            val_housing_in = emp_welfare_map.get('ค่าที่พัก', 0.0) 
            
            if 'pdf_other_income' in current_data:
                total_other_pdf = current_data['pdf_other_income']
                val_other_final = total_other_pdf - val_wear - val_allowance - val_housing_in
            else:
                val_other_final = 0.0
            if val_other_final < 0: val_other_final = 0 

            incomes = [ 
                ("เงินเดือน", current_data.get('base_salary', 0)), 
                ("ค่าตำแหน่ง", current_data.get('position_allowance', 0)), 
                ("ค่าล่วงเวลา", current_data.get('ot', 0)), 
                ("คอมมิชชั่น", current_data.get('commission', 0)), 
                ("โบนัส", current_data.get('bonus', 0)), 
                ("ค่าสึกหรอ", val_wear),      
                ("เบี้ยเลี้ยง", val_allowance),  
                ("ค่าที่พักนอก", val_housing_in),  
                ("เงินได้อื่นๆ", val_other_final)
            ]
            
            # --- 2. คำนวณฝั่งรายหัก (Deductions) ---
            # ดึง 'ค่าที่พักใน' จาก Input ที่บันทึกไว้
            emp_id = current_data.get('emp_id')
            val_housing_deduct = 0.0
            if hasattr(self, 'payroll_inputs') and emp_id in self.payroll_inputs:
                val_housing_deduct = float(self.payroll_inputs[emp_id].get('housing_deduct', 0.0))
            
            # แยกยอดออกจาก 'หักอื่นๆ'
            total_other_deduct = current_data.get('other_deduct', 0.0)
            val_other_deduct_final = total_other_deduct - val_housing_deduct
            if val_other_deduct_final < 0: val_other_deduct_final = 0

            deductions = [
                ("ประกันสังคม", current_data.get('sso', 0)),
                ("ภาษีเงินได้", 0),
                ("  - ภ.ง.ด. 1", current_data.get('pnd1', 0)),
                ("  - ภ.ง.ด. 3", current_data.get('pnd3', 0)),
                ("สำรองเลี้ยงชีพ", current_data.get('provident_fund', 0)), 
                ("หักเงินกู้ยืม (กยศ.)", current_data.get('loan', 0)),  # [RENAMED]
                ("ค่าที่พักใน", val_housing_deduct),                   # [NEW]
                ("ขาด/ลา/สาย", current_data.get('late_deduct', 0)), 
                ("หักอื่นๆ", val_other_deduct_final)                   # [UPDATED]
            ]

            pdf.set_font("THSarabun", "", 14)
            
            for i in range(max_rows):
                curr_y = body_top + (i * row_h)
                
                pdf.rect(10, curr_y, 190, row_h)
                pdf.line(105, curr_y, 105, curr_y + row_h) 
                pdf.line(90, curr_y, 90, curr_y + row_h)   
                pdf.line(185, curr_y, 185, curr_y + row_h)  
                
                # ฝั่งรับ
                if i < len(incomes):
                    label, val = incomes[i]
                    text_offset = (row_h - 5) / 2 
                    pdf.set_xy(10, curr_y + text_offset)
                    pdf.cell(55, 5, f"  {label}", border=0, align='L')
                    pdf.set_xy(65, curr_y + text_offset)
                    pdf.cell(25, 5, fmt_money(val), border=0, align='R')
                    pdf.set_xy(90, curr_y + text_offset)
                    pdf.cell(15, 5, "บาท", border=0, align='C')

                # ฝั่งหัก
                if i < len(deductions):
                    l2, v2 = deductions[i]
                    text_offset = (row_h - 5) / 2
                    pdf.set_xy(105, curr_y + text_offset)
                    pdf.cell(55, 5, f"  {l2}", border=0, align='L')
                    
                    show_val = fmt_money(v2)
                    if show_val != "-":
                        pdf.set_xy(160, curr_y + text_offset)
                        pdf.cell(25, 5, show_val, border=0, align='R')
                        pdf.set_xy(185, curr_y + text_offset)
                        pdf.cell(15, 5, "บาท", border=0, align='C')

            # ------------------------------------------------
            # ส่วนสรุปยอด (Totals)
            # ------------------------------------------------
            totals_y = body_top + (max_rows * row_h)
            
            pdf.set_fill_color(240, 240, 240) 
            pdf.rect(10, totals_y, 190, 7, 'F')
            pdf.rect(10, totals_y, 190, 7)
            pdf.line(105, totals_y, 105, totals_y + 7)

            pdf.set_font("THSarabun", "B", 14)
            
            # รวมรับ
            pdf.set_xy(10, totals_y)
            pdf.cell(55, 7, "  รวมเงินได้", 0, 0, 'L')
            pdf.set_xy(65, totals_y)
            pdf.cell(25, 7, fmt_money(current_data.get('total_income', 0)), 0, 0, 'R')
            pdf.set_xy(90, totals_y)
            pdf.cell(15, 7, "บาท", 0, 0, 'C')
            
            # รวมหัก
            pdf.set_xy(105, totals_y)
            pdf.cell(55, 7, "  รวมเงินหัก", 0, 0, 'L')
            pdf.set_xy(160, totals_y)
            pdf.cell(25, 7, fmt_money(current_data.get('total_deduct', 0)), 0, 0, 'R')
            pdf.set_xy(185, totals_y)
            pdf.cell(15, 7, "บาท", 0, 0, 'C')

            # ------------------------------------------------
            # ยอดสุทธิ (Net Salary)
            # ------------------------------------------------
            net_y = totals_y + 7
            pdf.set_fill_color(220, 220, 220) 
            pdf.rect(105, net_y, 95, 8, 'F')
            pdf.rect(105, net_y, 95, 8)
            
            pdf.set_xy(105, net_y)
            pdf.cell(55, 8, "  รวมเงินได้สุทธิ", 0, 0, 'L')
            pdf.set_xy(160, net_y)
            pdf.cell(25, 8, fmt_money(current_data.get('net_salary', 0)), 0, 0, 'R')
            pdf.set_xy(185, net_y)
            pdf.cell(15, 8, "บาท", 0, 0, 'C')

            # ------------------------------------------------
            # ลายเซ็น (Signature)
            # ------------------------------------------------
            sig_y = net_y + 10 
            pdf.set_font("THSarabun", "", 12)
            pdf.set_xy(10, sig_y)
            pdf.cell(60, 5, "ลงชื่อผู้จ่ายเงิน ...........................................", 0, 0, 'L')
            pdf.set_xy(105, sig_y)
            pdf.cell(60, 5, "ลงชื่อผู้รับเงิน ...........................................", 0, 0, 'L')

        # --- Main Loop: Generate Page ---
        for data in data_list:
            emp_welfare_map = {} 
            
            # ตรวจสอบและดึงข้อมูล
            if 'position' not in data or 'department' not in data or not data.get('bank_account'):
                emp_id = data.get('emp_id')
                if emp_id:
                    info = hr_database.load_single_employee(emp_id)
                    if info:
                        data['position'] = info.get('position', '-')
                        data['department'] = info.get('department', '-')
                        data['bank_account'] = info.get('bank_account_no') or info.get('account') or '-'
                        data['bank_name'] = info.get('bank_name') or info.get('bank') or '-'
                        
                        if 'name' not in data or not data['name']:
                             data['name'] = f"{info.get('fname','')} {info.get('lname','')}".strip()

                        w_opts = info.get('welfare_options', [])
                        w_flags = info.get('welfare', [])
                        w_amts = info.get('welfare_amounts', [])
                        
                        for i, name in enumerate(w_opts):
                            if i < len(w_flags) and w_flags[i]:
                                try:
                                    val = float(w_amts[i])
                                    emp_welfare_map[name] = val
                                except: pass

            pdf.add_page()
            
            # วาดต้นฉบับ
            draw_slip_form(data, 5, "(ต้นฉบับ)", emp_welfare_map)
            
            # เส้นประตัดกระดาษ
            pdf.set_draw_color(100)
            pdf.dashed_line(5, 148, 205, 148, dash_length=2, space_length=2)
            pdf.set_font("THSarabun", "", 10)
            pdf.text(185, 147, "ตัดตามรอยประ")
            
            # วาดสำเนา
            draw_slip_form(data, 155, "(สำเนา)", emp_welfare_map)
            
            # Timestamp ท้ายกระดาษ
            pdf.set_xy(10, 290)
            pdf.set_font("THSarabun", "", 10)
            pdf.cell(0, 5, f"พิมพ์เมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align='R')

        # Save File
        pdf.output(filepath)
    # (ใน payroll_module.py) -> แทนที่ฟังก์ชันนี้ได้เลยครับ

    def _load_employees_to_input_tree(self):
        """โหลดรายชื่อ แยกตามแท็บ (ฉบับสมบูรณ์ V18.7 - Clean)"""
        
        # 1. ตรวจสอบก่อนว่าสร้างตารางหรือยัง (กัน Error)
        if not hasattr(self, 'tree_reg') or not hasattr(self, 'tree_cont'):
            return

        # 2. ล้างข้อมูลเก่าทั้ง 2 ตาราง
        for i in self.tree_reg.get_children(): self.tree_reg.delete(i)
        for i in self.tree_cont.get_children(): self.tree_cont.delete(i)
        
        self.payroll_inputs = {}
        
        # 3. ดึงข้อมูลพนักงานทั้งหมด
        emps = hr_database.load_all_employees()
        
        # กลุ่มสัญญาพิเศษ
        special_types = ["สัญญาจ้างเหมา", "สัญญาจ้างเหมารายเดือน", "ที่ปรึกษา"]

        for emp in emps:
            # ข้ามคนที่ลาออก
            if emp.get('status') in ['พ้นสภาพพนักงาน', 'ลาออก']: continue
            
            emp_type = emp.get('emp_type', 'รายเดือน')
            emp_id = emp['id']
            name = f"{emp['fname']} {emp['lname']}"
            
            # เช็คสถานะการกรอก
            status_text = "-"
            if emp_id in self.payroll_inputs: status_text = "✅"

            # ข้อมูลที่จะใส่ในตาราง
            values = ("☑", emp_id, name, emp_type, status_text)
            
            # 4. แยกถังลงตาราง
            if emp_type in special_types:
                # ลง Tab 2 (Contract/Consultant)
                self.tree_cont.insert("", "end", iid=emp_id, values=values, tags=('row_special',))
            else:
                # ลง Tab 1 (Regular)
                self.tree_reg.insert("", "end", iid=emp_id, values=values, tags=('row_normal',))

    def _get_selected_dates(self):
        try:
            y = int(self.year_combo.get()) - 543
            m = self.MONTH_TO_INT[self.month_combo.get()]
            return y, m
        except: return None, None

    def _set_date_1_15(self):
        y, m = self._get_selected_dates()
        if y: 
            self.start_date_entry.set_date(datetime(y, m, 1))
            self.end_date_entry.set_date(datetime(y, m, 15))

    def _set_date_16_end(self):
        y, m = self._get_selected_dates()
        if y:
            last = calendar.monthrange(y, m)[1]
            self.start_date_entry.set_date(datetime(y, m, 16))
            self.end_date_entry.set_date(datetime(y, m, last))

    def _set_date_month(self):
        y, m = self._get_selected_dates()
        if y:
            last = calendar.monthrange(y, m)[1]
            self.start_date_entry.set_date(datetime(y, m, 1))
            self.end_date_entry.set_date(datetime(y, m, last))
    
    def _calculate_smart_tax(self, current_income, current_sso, current_pfund, 
                             ytd_income, ytd_tax_paid, ytd_sso, ytd_pfund, 
                             month_idx, is_resigned, other_allowances=0):
        
        # 1. ประมาณการรายได้ทั้งปี (Annualized Income)
        total_income_ytd = ytd_income + current_income
        total_sso_ytd = ytd_sso + current_sso
        total_pfund_ytd = ytd_pfund + current_pfund
        
        if is_resigned or month_idx == 12:
            annual_income = total_income_ytd
            annual_sso = total_sso_ytd
            annual_pfund = total_pfund_ytd
        else:
            m = max(1, month_idx)
            annual_income = (total_income_ytd / m) * 12
            annual_sso = (total_sso_ytd / m) * 12
            annual_pfund = (total_pfund_ytd / m) * 12

        # Cap ประกันสังคม (ไม่เกิน 9,000 ต่อปี)
        if annual_sso > 9000: annual_sso = 9000

        # 2. หักค่าใช้จ่าย (50% ไม่เกิน 100,000)
        expenses = min(annual_income * 0.5, 100000)

        # 3. หักค่าลดหย่อน (ส่วนตัว 60,000 + SSO + P.Fund + อื่นๆ)
        total_deductions = 60000 + annual_sso + annual_pfund + other_allowances

        # 4. เงินได้สุทธิ
        net_taxable = max(0, annual_income - expenses - total_deductions)

        # 5. คำนวณภาษีทั้งปี (Step Ladder)
        annual_tax = self._calculate_tax_step_ladder(net_taxable)

        # 6. หาภาษีงวดนี้ (วิธีสะสมยอด)
        # --- [แก้ไข] ประกาศตัวแปรไว้ก่อนเพื่อกัน Error ---
        expected_tax_ytd = 0.0 
        # -----------------------------------------------

        if is_resigned or month_idx == 12:
            # กรณีเดือนสุดท้าย หรือ ลาออก -> ภาษีที่ควรจ่ายทั้งปี คือ annual_tax เลย
            expected_tax_ytd = annual_tax 
            tax_this_month = annual_tax - ytd_tax_paid
        else:
            # กรณีระหว่างปี -> เฉลี่ยตามเดือน
            expected_tax_ytd = (annual_tax / 12) * month_idx
            tax_this_month = expected_tax_ytd - ytd_tax_paid

        final_tax = max(0, tax_this_month)

        # --- [DEBUG ZONE] แสดงที่มาของตัวเลข ---
        print("-" * 50)
        print(f"DEBUG TAX CALCULATION (Month {month_idx})")
        print(f"1. รายได้เดือนนี้: {current_income:,.2f}")
        print(f"2. ประมาณการรายได้ทั้งปี: {annual_income:,.2f}")
        print(f"3. หักค่าใช้จ่าย: {expenses:,.2f}")
        print(f"4. หักค่าลดหย่อน (รวม): {total_deductions:,.2f} (ประกันสังคมปีละ: {annual_sso})")
        print(f"5. เงินได้สุทธิ (Net Taxable): {net_taxable:,.2f}")
        print(f"6. ภาษีทั้งปี (Calculated Year Tax): {annual_tax:,.2f}")
        
        if annual_tax > 0:
            print(f"   -> เฉลี่ยต่อเดือน ({annual_tax} / 12): {annual_tax/12:,.2f}") 
        
        # ตอนนี้เรียกใช้ได้แล้ว ไม่ Error
        print(f"7. ภาษีสะสมที่ควรจ่ายถึงเดือนนี้: {expected_tax_ytd:,.2f}")
        print(f"8. ภาษีที่จ่ายไปแล้ว (YTD Paid): {ytd_tax_paid:,.2f}")
        print(f"9. สรุปภาษีงวดนี้: {final_tax:,.2f}")
        print("-" * 50)
        # ----------------------------------------

        return final_tax
    
    def _print_pnd1_summary(self):
        """(ฉบับแก้ไข 100%) ออกรายงาน ภ.ง.ด. 1 (PDF) พร้อมรายชื่อพนักงาน + ยอดรวม"""
        print("--- DEBUG: START PND1 PDF GENERATION ---")
        
        if not self.last_payroll_results:
            messagebox.showwarning("เตือน", "กรุณากดคำนวณเงินเดือนก่อน")
            return
        
        # --- 1. เตรียมข้อมูล ---
        pnd1_list = self.last_payroll_results
        total_emp = len(pnd1_list)
        grand_total_income = 0.0
        grand_total_tax = 0.0

        processed_list = []
        for emp in pnd1_list:
            # คำนวณรายได้สุทธิสำหรับภาษี (ถ้ามี Commission แยกก็ลบออก หรือตามกฎบริษัท)
            # ในที่นี้สมมติว่า total_income คือยอดรวมทั้งหมดที่จ่าย
            income_for_tax = float(emp.get('total_income', 0)) 
            tax_amount = float(emp.get('pnd1', 0))
            
            grand_total_income += income_for_tax
            grand_total_tax += tax_amount
            
            # ดึงชื่อ-นามสกุล (แก้ปัญหา KeyError: 'name')
            emp_id = emp.get('emp_id', '')
            # พยายามหาชื่อจากผลลัพธ์ก่อน ถ้าไม่มีไปดึงจาก DB
            emp_name = emp.get('name', '') 
            if not emp_name:
                # ดึงสดจาก DB
                emp_info = hr_database.load_single_employee(emp_id)
                if emp_info:
                    emp_name = f"{emp_info.get('fname', '')} {emp_info.get('lname', '')}"
                else:
                    emp_name = "ไม่ระบุชื่อ"

            processed_list.append({
                "id": emp_id,
                "name": emp_name,
                "income": income_for_tax,
                "tax": tax_amount
            })

        # --- 2. เลือกไฟล์ ---
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"PND1_Report_{datetime.now().strftime('%Y%m')}.pdf",
            title="บันทึกรายงาน ภ.ง.ด. 1"
        )
        if not save_path: return

        try:
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=False) # เราจะคุมเอง
            
            # --- โหลดฟอนต์ ---
            base_path = os.path.dirname(__file__)
            resource_path = os.path.join(base_path, "resources")
            font_path = os.path.join(resource_path, "THSarabunNew.ttf")
            
            if not os.path.exists(font_path): 
                font_path = os.path.join(base_path, "THSarabunNew.ttf")
            
            if not os.path.exists(font_path):
                messagebox.showerror("Font Error", f"ไม่พบไฟล์ฟอนต์ที่: {font_path}")
                return

            pdf.add_font("THSarabun", "", font_path, uni=True)
            
            # Config Columns
            COLS = [15, 30, 75, 35, 35]
            
            def fmt_money(val): return f"{val:,.2f}"

            # --- ฟังก์ชันวาดหัวกระดาษ ---
            def draw_page_header(page_num):
                pdf.set_xy(0, 15)
                pdf.set_text_color(0, 0, 0)
                
                pdf.set_font("THSarabun", "", 20)
                pdf.text(80, 20, "บริษัท เอไพร์ม พลัส จํากัด")
                
                pdf.set_font("THSarabun", "", 16)
                pdf.text(65, 28, "ใบแนบ ภ.ง.ด. 1 (รายละเอียดการหักภาษี ณ ที่จ่าย)")
                
                try:
                    s_date = self.start_date_entry.get_date()
                    month_th = list(self.THAI_MONTHS.values())[s_date.month - 1]
                    period_str = f"ประจำเดือน {month_th} พ.ศ. {s_date.year + 543}"
                except: period_str = "-"
                
                pdf.set_font("THSarabun", "", 14)
                pdf.text(85, 35, f"งวด: {period_str} (หน้าที่ {page_num})")

            # --- ฟังก์ชันวาดหัวตาราง ---
            def draw_table_header_fixed(y_pos):
                pdf.set_draw_color(0, 0, 0)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("THSarabun", "", 14) 

                cur_x = 10
                for width in COLS:
                    pdf.rect(cur_x, y_pos, width, 8)
                    cur_x += width
                
                text_y = y_pos + 6
                pdf.text(13, text_y, "ลำดับ")
                pdf.text(32, text_y, "รหัสพนักงาน")
                pdf.text(80, text_y, "ชื่อ - นามสกุล ผู้มีเงินได้")
                pdf.text(138, text_y, "จำนวนเงินได้")
                pdf.text(175, text_y, "ภาษีที่หัก")

            # --- เริ่มวาด ---
            pdf.add_page()
            draw_page_header(pdf.page_no())
            
            current_y = 45 
            draw_table_header_fixed(current_y)
            current_y += 8
            
            row_height = 7
            bottom_margin = 250 # เหลือที่ไว้สรุปยอดหน้าสุดท้าย

            # --- Loop Data ---
            for i, item in enumerate(processed_list):
                # เช็คหน้ากระดาษ
                if current_y + row_height > bottom_margin:
                    pdf.add_page()
                    draw_page_header(pdf.page_no())
                    current_y = 45
                    draw_table_header_fixed(current_y)
                    current_y += 8

                pdf.set_font("THSarabun", "", 14)
                pdf.set_xy(10, current_y)
                
                pdf.cell(COLS[0], row_height, str(i+1), 1, 0, 'C')
                
                if len(item['id']) > 10: pdf.set_font("THSarabun", "", 12)
                pdf.cell(COLS[1], row_height, item['id'], 1, 0, 'C')
                pdf.set_font("THSarabun", "", 14)
                
                pdf.cell(COLS[2], row_height, f"  {item['name']}", 1, 0, 'L')
                pdf.cell(COLS[3], row_height, fmt_money(item['income']), 1, 0, 'R')
                pdf.cell(COLS[4], row_height, fmt_money(item['tax']), 1, 0, 'R')
                
                current_y += row_height

            # --- ส่วนสรุปยอด (Summary) ---
            if current_y + 35 > 270: # ถ้าที่ไม่พอสรุป ให้ขึ้นหน้าใหม่
                pdf.add_page()
                draw_page_header(pdf.page_no())
                current_y = 45
            else:
                current_y += 5 

            pdf.set_font("THSarabun", "", 14)
            pdf.text(10, current_y + 6, "สรุปรายการนำส่งรวม:")
            current_y += 8
            
            w_box1 = 120 
            w_box2 = 35 
            w_box3 = 35 
            h_box = 18
            
            pdf.set_draw_color(0)
            pdf.rect(10, current_y, w_box1, h_box)               
            pdf.rect(10 + w_box1, current_y, w_box2, h_box)      
            pdf.rect(10 + w_box1 + w_box2, current_y, w_box3, h_box) 
            
            label_y = current_y + 6
            pdf.text(50, label_y, "รวมจำนวนราย (ราย)")      
            pdf.text(10 + w_box1 + 8, label_y, "รวมเงินได้") 
            pdf.text(10 + w_box1 + w_box2 + 10, label_y, "รวมภาษี") 
            
            val_y = current_y + 14
            pdf.set_font("THSarabun", "", 16)
            
            pdf.text(65, val_y, str(total_emp))
            
            income_txt = fmt_money(grand_total_income)
            income_w = pdf.get_string_width(income_txt)
            pdf.text(163 - income_w, val_y, income_txt)
            
            tax_txt = fmt_money(grand_total_tax)
            tax_w = pdf.get_string_width(tax_txt)
            pdf.text(198 - tax_w, val_y, tax_txt)

            # --- ลายเซ็น ---
            sig_y = current_y + 30
            pdf.set_font("THSarabun", "", 14)
            pdf.text(120, sig_y, "ลงชื่อ ....................................................... ผู้มีอำนาจลงนาม")
            pdf.text(125, sig_y + 7, f"( วันที่พิมพ์เอกสาร: {datetime.now().strftime('%d/%m/%Y')} )")

            pdf.output(save_path)
            os.startfile(save_path)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"สร้างใบปะหน้าไม่สำเร็จ: {e}")
        
    def _print_pnd3_summary(self):
        if not self.last_payroll_results: return
        
        # กรองคนที่มี PND3 > 0
        pnd3_list = [x for x in self.last_payroll_results if x.get('pnd3', 0) > 0]
        
        if not pnd3_list:
            messagebox.showwarning("ไม่พบข้อมูล", "ในรอบเดือนนี้ ไม่มีรายการหัก ภ.ง.ด. 3 (จากค่าคอมมิชชั่น)")
            return

        total_emp = len(pnd3_list)
        total_income = sum(x['commission'] for x in pnd3_list) # ยอดจ่าย (เฉพาะคอมมิชชั่น)
        total_tax = sum(x['pnd3'] for x in pnd3_list)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"PND3_Cover_{datetime.now().strftime('%Y%m')}.pdf",
            title="บันทึกใบปะหน้า ภ.ง.ด. 3"
        )
        if not save_path: return

        try:
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            
            # โหลดฟอนต์ (Logic เดิม)
            base_path = os.path.dirname(__file__)
            resource_path = os.path.join(base_path, "resources")
            font_path_reg = os.path.join(resource_path, "THSarabunNew.ttf")
            if not os.path.exists(font_path_reg): font_path_reg = os.path.join(base_path, "THSarabunNew.ttf")
            font_path_bold = os.path.join(resource_path, "THSarabunNew Bold.ttf")
            if not os.path.exists(font_path_bold): font_path_bold = os.path.join(base_path, "THSarabunNew Bold.ttf")
            if not os.path.exists(font_path_bold): font_path_bold = font_path_reg

            pdf.add_font("THSarabun", "", font_path_reg, uni=True)
            pdf.add_font("THSarabun", "B", font_path_bold, uni=True)

            # หัวกระดาษ
            pdf.set_font("THSarabun", "B", 22)
            pdf.set_xy(0, 15)
            pdf.cell(0, 10, "สรุปรายการภาษีเงินได้หัก ณ ที่จ่าย (ภ.ง.ด. 3)", ln=True, align='C')
            
            pdf.set_font("THSarabun", "", 16)
            try:
                s_date = self.start_date_entry.get_date()
                month_th = list(self.THAI_MONTHS.values())[s_date.month - 1]
                period_str = f"เดือน {month_th} พ.ศ. {s_date.year + 543}"
            except: period_str = "-"
            pdf.cell(0, 10, f"ประจำงวด: {period_str}", ln=True, align='C')
            
            # ตารางสรุป (Manual XY)
            start_y = 45 
            box_w = 160
            center_x = (210 - box_w) / 2
            pdf.set_draw_color(0); pdf.set_line_width(0.2)
            
            def draw_row(y, label, value, is_bold=False):
                pdf.set_xy(center_x, y)
                pdf.rect(center_x, y, box_w, 12)
                pdf.set_font("THSarabun", "B" if is_bold else "", 16)
                pdf.set_xy(center_x + 5, y + 1)
                pdf.cell(box_w * 0.6, 10, label, border=0)
                pdf.set_font("THSarabun", "B", 16)
                pdf.set_xy(center_x + (box_w * 0.6), y + 1)
                pdf.cell(box_w * 0.4 - 5, 10, value, border=0, align='R')

            draw_row(start_y, "จำนวนผู้มีเงินได้ (ราย)", f"{total_emp}  ราย")
            draw_row(start_y + 12, "รวมเงินได้ทั้งสิ้น (ยอดคอมมิชชั่น)", f"{total_income:,.2f}  บาท")
            draw_row(start_y + 24, "รวมภาษีที่หักนำส่ง (ภ.ง.ด. 3)", f"{total_tax:,.2f}  บาท", is_bold=True)

            # หมายเหตุ
            note_y = start_y + 40
            pdf.set_font("THSarabun", "", 14)
            pdf.set_xy(center_x, note_y)
            pdf.cell(0, 8, "* เอกสารนี้เป็นใบสรุปสำหรับตรวจสอบภายใน เพื่อนำตัวเลขไปกรอกลงในแบบฟอร์ม", border=0)
            pdf.set_xy(center_x + 3, note_y + 6)
            pdf.cell(0, 8, "ภ.ง.ด. 3 ของกรมสรรพากร (หัก ณ ที่จ่าย 3% จากค่าคอมมิชชั่น)", border=0)

            sig_y = 200
            pdf.set_xy(0, sig_y)
            pdf.cell(0, 8, "ลงชื่อ ....................................................... ผู้จัดทำ", ln=0, align='C')
            pdf.set_xy(0, sig_y + 8)
            pdf.cell(0, 8, f"( วันที่พิมพ์: {datetime.now().strftime('%d/%m/%Y')} )", ln=0, align='C')

            pdf.output(save_path)
            os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("Error", f"สร้างใบปะหน้าไม่สำเร็จ: {e}")

    def _sync_commission_from_asmart(self):
        """ปุ่มสำหรับดึงข้อมูลจาก A+ Smart (อัปเดต: ค้นหาด้วย 'ชื่อ' แทน 'รหัส')"""
        
        # 1. เช็คว่าเลือกเดือน/ปี หรือยัง
        y_ce, m_int = self._get_selected_dates()
        if not y_ce: return

        if not messagebox.askyesno("ยืนยัน", f"ต้องการดึงยอดคอมมิชชั่นประจำเดือน {m_int}/{y_ce}\nจากระบบ A+ Smart (192.168.1.51) หรือไม่?"):
            return

        updated_count = 0
        
        # รวมตารางที่มีทั้งหมดเพื่อวนลูปทีเดียว
        target_trees = []
        if hasattr(self, 'tree_reg'): target_trees.append(self.tree_reg)
        if hasattr(self, 'tree_cont'): target_trees.append(self.tree_cont)

        # 2. วนลูปทุกตาราง (ประจำ + จ้างเหมา)
        for tree in target_trees:
            for item_id in tree.get_children():
                vals = tree.item(item_id, "values")
                
                emp_id = vals[1] 
                emp_name = vals[2] # ดึงชื่อ-นามสกุลจากคอลัมน์ที่ 3
                
                # 3. ไปดึงยอดเงินจาก DB โดยส่ง 'ชื่อ' ไปแทน
                comm_amt = hr_database.get_commission_from_asmart_by_name(emp_name, m_int, y_ce)
                
                if comm_amt > 0:
                    if emp_id not in self.payroll_inputs:
                        self.payroll_inputs[emp_id] = {}
                    
                    # 🛠️ [FIX] ยัดเงินเข้า 'วิก 2' (commission_2) เพื่อให้ตารางคำนวณและหน้า Popup มองเห็น!
                    self.payroll_inputs[emp_id]['commission_2'] = comm_amt
                    
                    tree.set(item_id, column="status", value=f"✅ A+ Smart: {comm_amt:,.2f}")
                    updated_count += 1

        if updated_count > 0:
            messagebox.showinfo("สำเร็จ", f"ดึงข้อมูลคอมมิชชั่นเรียบร้อย {updated_count} รายการ")
        else:
            messagebox.showinfo("ไม่พบข้อมูล", "ไม่พบยอดคอมมิชชั่นในช่วงเวลานี้ (หรือชื่อพนักงานไม่ตรงกันเลย)")

    def _open_daily_timesheet(self):
        """เปิดหน้าต่างบันทึกงานรายวัน (Daily Timesheet) - ฉบับรองรับ 2 แท็บ"""
        
        # 1. ตรวจสอบว่าผู้ใช้กำลังเปิดแท็บไหนอยู่ (0=ประจำ, 1=จ้างเหมา)
        current_tab_index = self.input_notebook.index(self.input_notebook.select())
        
        # 2. เลือกตาราง (Treeview) ให้ตรงกับแท็บ
        if current_tab_index == 0:
            target_tree = self.tree_reg  # แท็บพนักงานประจำ
        else:
            target_tree = self.tree_cont # แท็บจ้างเหมา
            
        # 3. ดึงรายการที่เลือกจากตารางนั้น
        selection = target_tree.selection()
        
        if not selection:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกพนักงานที่ต้องการบันทึกงานรายวัน")
            return
            
        # ดึงรหัสพนักงาน (Values: 0=checkbox, 1=id, 2=name, ...)
        emp_id = target_tree.item(selection[0], "values")[1] 
        
        # 4. เช็คว่าเลือกเดือน/ปีหรือยัง
        y_ce, m_int = self._get_selected_dates()
        if not y_ce: 
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือก ปีและเดือน ที่ต้องการบันทึกก่อน")
            return

        # เปิดหน้าต่างใหม่
        DailyTimesheetWindow(self, emp_id, m_int, y_ce)
    
    def _print_50tawi_pdf(self):
        """ออกหนังสือรับรอง 50 ทวิ (Overlay) - V10 (เพิ่ม Bahttext)"""
        
        # 1. เช็คการเลือกพนักงาน
        selected_indexes = self.results_sheet.get_selected_rows(return_tuple=True)
        if not selected_indexes:
            messagebox.showwarning("เตือน", "กรุณาเลือกพนักงานในตารางผลลัพธ์ 1 คน")
            return
        
        selected_data = self.last_payroll_results[selected_indexes[0]]
        emp_id = selected_data['emp_id']
        
        # 2. ถามปีภาษี
        current_year = datetime.now().year
        year_str = simpledialog.askstring("ปีภาษี", f"กรุณากรอกปี พ.ศ. (เช่น {current_year+543}):", initialvalue=str(current_year+543))
        if not year_str or not year_str.isdigit(): return
        year_be = int(year_str)
        year_ce = year_be - 543

        # 3. ดึงข้อมูล
        emp_data = hr_database.get_employee_annual_summary(emp_id, year_ce)
        
        if not emp_data:
            messagebox.showerror("Error", "ไม่พบข้อมูลการจ่ายเงินในปีที่ระบุ")
            return

        # 4. หาไฟล์ Template
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "approve_wh3_081156.pdf") 
        if not os.path.exists(template_path):
             template_path = os.path.join(base_dir, "resources", "approve_wh3_081156.pdf")
             if not os.path.exists(template_path):
                messagebox.showerror("Error", "ไม่พบไฟล์ Template: approve_wh3_081156.pdf")
                return

        try:
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=A4)
            
            # โหลดฟอนต์
            font_path = os.path.join(base_dir, "resources", "THSarabunNew.ttf")
            if not os.path.exists(font_path): font_path = os.path.join(base_dir, "THSarabunNew.ttf")
            pdfmetrics.registerFont(TTFont('THSarabun', font_path))
            
            # ==================================================================================
            # 🎯 1. ดึงข้อมูลบริษัทจาก Database (แทนการ Hardcode)
            # ==================================================================================
            tax_id_db = hr_database.get_company_setting("tax_id")
            PAYER_TAX_ID = tax_id_db if tax_id_db else "0205558005856" 
            
            locations = hr_database.get_company_locations()
            hq_info = next((loc for loc in locations if loc['loc_type'] == 'สำนักงานใหญ่'), None)
            
            if hq_info:
                PAYER_NAME = hq_info['loc_name']
                # (ถ้ามีที่อยู่จริงใน DB ก็ดึงมาใส่ตรงนี้ครับ)
                PAYER_ADDR = "9/106 ซอยเอกชัย 119 แยก 1 แขวงบางบอนใต้ เขตบางบอน กรุงเทพมหานคร 10150" 
            else:
                PAYER_NAME = "บริษัท เอไพร์ม พลัส จำกัด"
                PAYER_ADDR = "9/106 ซอยเอกชัย 119 แยก 1 แขวงบางบอนใต้ เขตบางบอน กรุงเทพมหานคร 10150"

            # ==================================================================================
            # 🎯 2. คำนวณตัวเลขต่างๆ (Sequence Logic)
            # ==================================================================================
            seq_no = emp_data.get('sequence_no', 0)
            yy = str(year_be)[-2:] 
            seq_text_3digit = f"{seq_no:03d}"
            book_no_text = f"{yy}/01"
            running_no_text = f"AP{yy}-{seq_no:02d}"

            # ==================================================================================
            # 🎯 3. โซนตั้งค่าพิกัด (Coordinates Config)
            # ==================================================================================
            PAYER_ID_X = 376; PAYER_ID_Y = 747          
            PAYER_NAME_X = 60; PAYER_NAME_Y = 730        
            PAYER_ADDR_X = 60; PAYER_ADDR_Y = 708        

            X_BOOK_NO = 490; Y_BOOK_NO = 783
            X_RUNNING_NO = 493; Y_RUNNING_NO = 768
            
            X_SEQ_NO = 80; Y_SEQ_NO = 605 
            
            ID_X = 377; ID_Y = 678                
            NAME_X = 60; NAME_Y = 660              
            ADDR_X = 60; ADDR_Y = 631              

            ID_SPACING = 10.9; ID_GROUP_GAP = 10.3

            Y_INCOME_ROW_1 = 538
            X_DATE = 330
            X_AMOUNT = 487.5; X_TAX = 557.5

            Y_TOTAL = 181
            Y_SSO = 144.5
            X_SSO = 400; X_FUND = 430  
            
            # --- [NEW] พิกัดสำหรับ "ตัวอักษรภาษาไทย" ---
            X_BAHT_TEXT = 210  # ปรับซ้ายขวาตรงนี้ (ค่าเดิมประมาณนี้)
            Y_BAHT_TEXT = 160  # ปรับขึ้นลงตรงนี้ (อยู่เหนือช่องรวมเงิน)                      
            
            # Helper Function วาดเลขห่างๆ
            def draw_id_card_spaced(c, x, y, text, spacing=13, group_gap=8):
                c.setFont('THSarabun', 16) 
                text = str(text).replace("-", "").strip()
                curr_x = x
                jump_indices = [0, 4, 9, 11]
                for i, char in enumerate(text):
                    c.drawString(curr_x, y, char)
                    step = spacing
                    if i in jump_indices: step += group_gap
                    curr_x += step

            # ==================================================================================
            # 🖌️ เริ่มวาดข้อมูล
            # ==================================================================================
            
            # 1. วาด "เล่มที่" / "เลขที่" / "ลำดับที่"
            c.setFont('THSarabun', 16)
            c.drawRightString(X_BOOK_NO + 60, Y_BOOK_NO, book_no_text)     
            c.setFont('THSarabun', 14)  
            c.drawRightString(X_RUNNING_NO + 60, Y_RUNNING_NO, running_no_text)
            c.setFont('THSarabun', 16) 
            c.drawString(X_SEQ_NO, Y_SEQ_NO, seq_text_3digit)
            
            # 2. ข้อมูลบริษัท
            draw_id_card_spaced(c, PAYER_ID_X, PAYER_ID_Y, PAYER_TAX_ID, spacing=ID_SPACING, group_gap=ID_GROUP_GAP)
            c.setFont('THSarabun', 14)
            c.drawString(PAYER_NAME_X, PAYER_NAME_Y, PAYER_NAME)
            c.drawString(PAYER_ADDR_X, PAYER_ADDR_Y, PAYER_ADDR)

            # 3. ข้อมูลพนักงาน
            emp_card_id = emp_data.get('id_card', "") or ""
            draw_id_card_spaced(c, ID_X, ID_Y, emp_card_id, spacing=ID_SPACING, group_gap=ID_GROUP_GAP)
            c.setFont('THSarabun', 14)
            c.drawString(NAME_X, NAME_Y, f"{emp_data.get('fname','')} {emp_data.get('lname','')}")
            c.drawString(ADDR_X, ADDR_Y, emp_data.get('address','') or "-")
            
            # 4. ส่วนเงินได้ (ตารางกลาง)
            THAI_MONTHS_SHORT = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
            start_month = int(emp_data.get('start_month', 1)) 
            end_month = int(emp_data.get('end_month', 12))
            
            period_text = ""
            if start_month == 1 and end_month == 12:
                period_text = f"ตลอดปี {year_be}"
            else:
                try:
                    s_name = THAI_MONTHS_SHORT[start_month]
                    e_name = THAI_MONTHS_SHORT[end_month]
                    period_text = f"{s_name} - {e_name} {year_be}"
                except:
                    period_text = f"ตลอดปี {year_be}"

            c.setFont('THSarabun', 12)
            c.drawString(X_DATE, Y_INCOME_ROW_1, period_text)

            c.setFont('THSarabun', 14)
            income_val = float(emp_data.get('total_income', 0))
            tax_val = float(emp_data.get('total_tax', 0))
            
            c.drawRightString(X_AMOUNT, Y_INCOME_ROW_1, f"{income_val:,.2f}")
            c.drawRightString(X_TAX, Y_INCOME_ROW_1, f"{tax_val:,.2f}")

            # 5. ยอดรวม (บรรทัดล่าง)
            c.drawRightString(X_AMOUNT, Y_TOTAL, f"{income_val:,.2f}")
            c.drawRightString(X_TAX, Y_TOTAL, f"{tax_val:,.2f}")

            # --- [NEW] พิมพ์ตัวอักษรภาษาไทย (Bahttext) ---
            try:
                # แปลงยอดภาษีรวม เป็นภาษาไทย (หรือจะใช้ income_val ก็ได้ถ้าต้องการยอดเงินได้)
                # ปกติช่อง "รวมเงินภาษีที่หักนำส่ง (ตัวอักษร)" จะอยู่ข้างล่างยอดรวมภาษี
                text_baht = bahttext(tax_val) 
                
                c.setFont('THSarabun', 14)
                c.drawString(X_BAHT_TEXT, Y_BAHT_TEXT, f"({text_baht})")
            except Exception as e:
                print(f"Bahttext Error: {e}")
            # ---------------------------------------------

            # 6. ประกันสังคม / กองทุน
            sso_val = float(emp_data.get('total_sso', 0))
            fund_val = float(emp_data.get('total_fund', 0))
            
            if sso_val > 0:
                c.drawRightString(X_SSO, Y_SSO, f"{sso_val:,.2f}")
            if fund_val > 0:
                c.drawRightString(X_FUND, Y_SSO, f"{fund_val:,.2f}")

            # 7. ลงชื่อ
            # วันที่จ่าย/ลงชื่อ (ถ้าต้องการ Fix วันที่พิมพ์)
            # c.drawString(X_DATE_SIGN, Y_DATE_SIGN, datetime.now().strftime("%d/%m/%Y"))

            c.save()
            packet.seek(0)

            # รวมร่าง PDF
            new_pdf = PdfReader(packet)
            existing_pdf = PdfReader(open(template_path, "rb"))
            output = PdfWriter()
            page = existing_pdf.pages[0]
            page.merge_page(new_pdf.pages[0])
            output.add_page(page)
            
            # บันทึก
            clean_run_no = running_no_text.replace('/', '-')
            save_filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile=f"50Tawi_{clean_run_no}_{emp_id}.pdf",
                title="บันทึกหนังสือรับรอง 50 ทวิ"
            )
            
            if save_filename:
                with open(save_filename, "wb") as f:
                    output.write(f)
                if messagebox.askyesno("สำเร็จ", "สร้างใบ 50 ทวิ เรียบร้อยแล้ว\nต้องการเปิดดูเลยหรือไม่?"):
                    os.startfile(save_filename)

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    def _export_sso_excel(self):
        """ออกรายงานนำส่งเงินสมทบประกันสังคม (สปส. 1-10) เป็น Excel"""
        if not self.last_payroll_results:
            messagebox.showwarning("เตือน", "กรุณากดคำนวณเงินเดือนก่อนครับ")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"SSO_Report_{datetime.now().strftime('%Y%m')}.xlsx",
            title="บันทึกรายงาน สปส. 1-10"
        )
        if not save_path: return

        try:
            data_rows = []
            seq = 1
            
            for item in self.last_payroll_results:
                emp_id = item['emp_id']
                
                # ดึงข้อมูลส่วนตัว
                emp_info = hr_database.load_single_employee(emp_id)
                id_card = emp_info.get('id_card', '') if emp_info else ''
                fname = emp_info.get('fname', '') if emp_info else ''
                lname = emp_info.get('lname', '') if emp_info else ''
                
                # ดึงยอดที่หักจริง
                sso_amount = float(item.get('sso', 0))
                if sso_amount <= 0: continue # ถ้าไม่หักประกันสังคม ไม่ต้องส่งชื่อ
                
                # คำนวณฐานค่าจ้าง (Base Wage) ย้อนกลับจากยอดหัก (5%)
                # หรือจะใช้ item['base_salary'] ก็ได้ แต่ต้อง Cap ที่ 15,000
                wage_base = sso_amount / 0.05 
                if wage_base > 15000: wage_base = 15000
                
                row = {
                    "ลำดับ": seq,
                    "เลขบัตรประชาชน": id_card,
                    "คำนำหน้า": "", 
                    "ชื่อผู้ประกันตน": fname,
                    "นามสกุลผู้ประกันตน": lname,
                    "ค่าจ้างที่นำส่งเงินสมทบ": wage_base,
                    "เงินสมทบผู้ประกันตน": sso_amount,
                    "เงินสมทบนายจ้าง": sso_amount, # นายจ้างจ่ายเท่าลูกจ้าง
                    "รวมเงินสมทบ": sso_amount * 2,
                    "สถานะ": "ปกติ"
                }
                data_rows.append(row)
                seq += 1
            
            if not data_rows:
                messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบรายการหักประกันสังคมในงวดนี้")
                return

            df = pd.DataFrame(data_rows)
            
            # เพิ่มแถวรวม (Total) ท้ายตาราง
            total_row = {
                "ลำดับ": "", "เลขประจำตัวประชาชน": "", "คำนำหน้า": "", "ชื่อผู้ประกันตน": ">>> รวมทั้งสิ้น <<<", "นามสกุลผู้ประกันตน": "",
                "ค่าจ้างที่นำส่งเงินสมทบ": df["ค่าจ้างที่นำส่งเงินสมทบ"].sum(),
                "เงินสมทบผู้ประกันตน": df["เงินสมทบผู้ประกันตน"].sum(),
                "เงินสมทบนายจ้าง": df["เงินสมทบนายจ้าง"].sum(),
                "รวมเงินสมทบ": df["รวมเงินสมทบ"].sum(),
            }
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            # บันทึก
            df.to_excel(save_path, index=False)
            
            if messagebox.askyesno("สำเร็จ", f"บันทึกรายงานเรียบร้อยที่:\n{save_path}\n\nเปิดดูเลยหรือไม่?"):
                os.startfile(save_path)

        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {e}")

    def _print_pnd1k_pdf(self):
        """ออกรายงาน ภ.ง.ด. 1ก (Overlay) - ฉบับสมบูรณ์ (รวมยอดแยกตามหน้า)"""
        
        # ต้องแน่ใจว่า import landscape มาแล้วด้านบนไฟล์:
        # from reportlab.lib.pagesizes import A4, landscape

        # ==========================================
        # 🟢 Helper Function: วาดข้อความแบบกำหนดระยะห่าง
        # ==========================================
        def draw_spaced(c, x, y, text, char_space):
            c.saveState()
            t = c.beginText(x, y)
            t.setFont("THSarabun", 14) 
            t.setCharSpace(char_space) 
            t.textOut(text)
            c.drawText(t)
            c.restoreState()
        # ==========================================

        # 1. ถามปีภาษี
        current_year_be = datetime.now().year + 543
        year_str = simpledialog.askstring("เลือกปีภาษี", f"กรุณากรอกปี พ.ศ. (เช่น {current_year_be}):", initialvalue=str(current_year_be))
        if not year_str or not year_str.isdigit(): return
        year_be = int(year_str)
        year_ce = year_be - 543 

        # 2. ดึงข้อมูล
        data_list = hr_database.get_annual_pnd1k_data(year_ce)
        if not data_list:
            messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติการจ่ายเงินเดือนในปี {year_be}")
            return

        # 3. ตรวจสอบไฟล์ Template
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "pnd1k.pdf")
        if not os.path.exists(template_path):
             template_path = os.path.join(base_dir, "resources", "pnd1k.pdf")
        
        if not os.path.exists(template_path):
            messagebox.showerror("Error", "ไม่พบไฟล์แบบฟอร์มต้นฉบับ (pnd1k.pdf)")
            return

        # 4. เลือกตำแหน่งบันทึก
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"PND1K_{year_be}_Completed.pdf",
            title="บันทึก ภ.ง.ด. 1ก"
        )
        if not save_path: return

        try:
            # ลงทะเบียน Font
            font_path = os.path.join(base_dir, "resources", "THSarabunNew.ttf")
            if not os.path.exists(font_path): font_path = os.path.join(base_dir, "THSarabunNew.ttf")
            pdfmetrics.registerFont(TTFont('THSarabun', font_path))

            output_writer = PdfWriter()
            
            # ==========================================
            # 🎯 CONFIGURATION
            # ==========================================
            Y_START = 413           
            ROW_HEIGHT = 39         
            MAX_ROW_PER_PAGE = 7    
            
            # พิกัดเลขหน้า
            Y_PAGE = 475          
            X_PAGE_CURR = 701       
            X_PAGE_TOTAL = 765      

            # พิกัดข้อมูลอื่นๆ
            X_SEQ = 80              
            X_ID_1 = 111; X_ID_2 = 127; ID_SPACE_2 = 6.9 
            X_ID_3 = 180; ID_SPACE_3 = 6.9; X_ID_4 = 247; ID_SPACE_4 = 4.7 
            X_ID_5 = 278            
            X_FNAME = 302; X_LNAME = 452
            X_ADDRESS = 315; Y_ADDR_OFFSET = 15
            X_INCOME = 680; X_TAX = 780; X_COND = 790
            Y_TOTAL_ROW = 135       
            # ==========================================
            
            # LOGIC Pagination
            chunks = [data_list[i:i + MAX_ROW_PER_PAGE] for i in range(0, len(data_list), MAX_ROW_PER_PAGE)]
            total_pages = len(chunks) 
            seq_global = 1 

            # เริ่มวนลูปสร้างทีละหน้า
            for page_idx, batch in enumerate(chunks):
                current_page_num = page_idx + 1 
                
                packet = io.BytesIO()
                
                # ✅ ใช้ landscape(A4)
                c = canvas.Canvas(packet, pagesize=landscape(A4)) 
                
                # --- พิมพ์เลขหน้า ---
                c.saveState()
                c.setFont("THSarabun", 18) 
                c.drawString(X_PAGE_CURR, Y_PAGE, str(current_page_num)) 
                c.drawString(X_PAGE_TOTAL, Y_PAGE, str(total_pages))     
                c.restoreState()

                c.setFont("THSarabun", 14) 
                current_y = Y_START
                
                # --- วนลูปรายชื่อ ---
                for item in batch:
                    inc = float(item.get('annual_income', 0) or 0)
                    tax = float(item.get('annual_tax', 0) or 0)
                    
                    c.drawCentredString(X_SEQ, current_y, str(seq_global))
                    
                    tid = str(item.get('id_card', '')).replace('-', '').replace(' ', '').strip()
                    if len(tid) == 13:
                        c.drawString(X_ID_1, current_y, tid[0])
                        draw_spaced(c, X_ID_2, current_y, tid[1:5], ID_SPACE_2)
                        draw_spaced(c, X_ID_3, current_y, tid[5:10], ID_SPACE_3)
                        draw_spaced(c, X_ID_4, current_y, tid[10:12], ID_SPACE_4)
                        c.drawString(X_ID_5, current_y, tid[12])
                    else:
                        c.drawString(X_ID_1, current_y, tid) 

                    c.drawString(X_FNAME, current_y, item.get('fname', ''))
                    c.drawString(X_LNAME, current_y, item.get('lname', ''))
                    
                    addr_y = current_y - Y_ADDR_OFFSET
                    c.setFont("THSarabun", 10) 
                    c.drawString(X_ADDRESS, addr_y, item.get('address', '-'))
                    c.setFont("THSarabun", 14) 

                    c.drawRightString(X_INCOME, current_y, f"{inc:,.2f}")
                    c.drawRightString(X_TAX, current_y, f"{tax:,.2f}")
                    c.drawCentredString(X_COND, current_y, "1")

                    current_y -= ROW_HEIGHT
                    seq_global += 1

                # =========================================================
                # 🎯 แก้ไข: สรุปยอดรวม "ตามหน้า" (Page Total) ไม่ใช่สะสม
                # =========================================================
                
                # 1. คำนวณผลรวมเฉพาะข้อมูลในหน้านี้ (batch)
                page_income = sum(float(item.get('annual_income', 0) or 0) for item in batch)
                page_tax = sum(float(item.get('annual_tax', 0) or 0) for item in batch)

                # 2. สั่งพิมพ์ลงทุกหน้าเลย (ไม่ต้องเช็คว่าเป็นหน้าสุดท้ายไหม)
                c.drawRightString(X_INCOME, Y_TOTAL_ROW, f"{page_income:,.2f}")
                c.drawRightString(X_TAX, Y_TOTAL_ROW, f"{page_tax:,.2f}")

                # =========================================================

                c.save()
                packet.seek(0)

                template_reader = PdfReader(open(template_path, "rb"))
                bg_page = template_reader.pages[0] 
                overlay_reader = PdfReader(packet)
                bg_page.merge_page(overlay_reader.pages[0])
                output_writer.add_page(bg_page)

            with open(save_path, "wb") as f_out:
                output_writer.write(f_out)

            if messagebox.askyesno("สำเร็จ", f"สร้างไฟล์เรียบร้อยแล้วที่:\n{save_path}\n\nต้องการเปิดดูเลยหรือไม่?"):
                os.startfile(save_path)

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{e}")

    def _create_treeview(self, parent_frame):
        """Helper สร้าง Treeview แบบมี Checkbox"""
        container = ttk.Frame(parent_frame)
        container.pack(fill="both", expand=True)
        
        scroll_y = ttk.Scrollbar(container, orient="vertical")
        scroll_y.pack(side="right", fill="y")
        
        tree = ttk.Treeview(container, columns=("select", "id", "name", "type", "status"), show="headings", yscrollcommand=scroll_y.set, height=12)
        scroll_y.config(command=tree.yview)

        tree.heading("select", text="เลือก"); tree.column("select", width=50, anchor="center")
        tree.heading("id", text="รหัส"); tree.column("id", width=80, anchor="center")
        tree.heading("name", text="ชื่อ-นามสกุล"); tree.column("name", width=200, anchor="w")
        tree.heading("type", text="ประเภท"); tree.column("type", width=150, anchor="center")
        tree.heading("status", text="สถานะ"); tree.column("status", width=120, anchor="center")
        
        tree.pack(side="left", fill="both", expand=True)
        
        # Bindings (ต้องใช้ Wrapper เพราะเรามี 2 ตาราง)
        tree.bind("<ButtonRelease-1>", lambda event, t=tree: self._on_tree_click_wrapper(event, t))
        tree.bind("<Double-1>", lambda event, t=tree: self._open_input_popup_wrapper(event, t))
        
        tree.tag_configure('row_normal', foreground='black')
        tree.tag_configure('row_special', foreground='#8B4513')
        return tree

    def _on_tree_click_wrapper(self, event, tree_widget):
        region = tree_widget.identify("region", event.x, event.y)
        if region == "cell":
            col = tree_widget.identify_column(event.x)
            if col == "#1":
                item_id = tree_widget.identify_row(event.y)
                if not item_id: return
                curr = tree_widget.item(item_id, "values")[0]
                new_val = "☐" if curr == "☑" else "☑"
                old_vals = tree_widget.item(item_id, "values")
                tree_widget.item(item_id, values=(new_val, *old_vals[1:]))

    def _open_input_popup_wrapper(self, event, tree_widget):
        self.input_tree = tree_widget # Trick: หลอกฟังก์ชันเดิมว่ากำลังใช้ตารางนี้
        self._open_input_popup(event)
    
    def _toggle_all_selection(self, select_all=True):
        # ดูว่า Tab ไหนเปิดอยู่ ก็เลือกเฉพาะตารางนั้น
        current_tab_index = self.input_notebook.index(self.input_notebook.select())
        target_tree = self.tree_reg if current_tab_index == 0 else self.tree_cont
        symbol = "☑" if select_all else "☐"
        for item in target_tree.get_children():
            old_vals = target_tree.item(item, "values")
            target_tree.item(item, values=(symbol, *old_vals[1:]))