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
            sender_password = "utak iewz bejb qvnm"   # <---  แก้เป็นรหัส App 16 หลัก
            
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
        # สร้าง Frame ส่วนหัวสำหรับปุ่มและตัวกรอง
        top_frame = ttk.Frame(parent_tab)
        top_frame.pack(fill="x", pady=(0, 10))

        # --- [เพิ่มใหม่] ส่วนตัวกรองประเภทพนักงาน ---
        filter_frame = ttk.LabelFrame(top_frame, text=" กรองรายชื่อ ", padding=5)
        filter_frame.pack(side="left", padx=(0, 10))

        ttk.Label(filter_frame, text="ประเภท:").pack(side="left", padx=5)
        
        self.emp_type_var = tk.StringVar(value="ทั้งหมด")
        self.emp_type_combo = ttk.Combobox(filter_frame, textvariable=self.emp_type_var, 
                                           values=["ทั้งหมด", "รายเดือน", "รายวัน"], 
                                           state="readonly", width=10)
        self.emp_type_combo.pack(side="left", padx=5)
        self.emp_type_combo.bind("<<ComboboxSelected>>", lambda e: self._load_employees_to_input_tree()) # เลือกปุ๊บโหลดใหม่ปั๊บ

        # --- ปุ่มคำสั่งเดิม (ย้ายมาใส่ top_frame ต่อจาก filter) ---
        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(side="left")

        ttk.Button(btn_frame, text="🔄 โหลดรายชื่อ", command=self._load_employees_to_input_tree).pack(side="left", padx=5)

        ttk.Button(btn_frame, text="💰 ดึงยอด A+ Smart", 
                   command=self._sync_commission_from_asmart,
                   style="Primary.TButton").pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="📅 บันทึก Timesheet", 
                   command=self._open_daily_timesheet).pack(side="left", padx=5)
        
        # --- ตาราง Treeview (เหมือนเดิม แต่เพิ่มคอลัมน์ "ประเภท") ---
        tree_container = ttk.Frame(parent_tab)
        tree_container.pack(fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        # เพิ่มคอลัมน์ "type"
        self.input_tree = ttk.Treeview(
            tree_container,
            columns=("id", "name", "type", "status"), # <--- เพิ่ม "type"
            show="headings",
            yscrollcommand=scrollbar_y.set,
            height=15
        )
        self.input_tree.heading("id", text="รหัส")
        self.input_tree.heading("name", text="ชื่อ-นามสกุล")
        self.input_tree.heading("type", text="ประเภท") # <--- เพิ่มหัวข้อ
        self.input_tree.heading("status", text="สถานะการป้อนข้อมูล")
        
        self.input_tree.column("id", width=80, anchor="center")
        self.input_tree.column("name", width=200, anchor="w")
        self.input_tree.column("type", width=100, anchor="center") # <--- กำหนดความกว้าง
        self.input_tree.column("status", width=150, anchor="center")
        
        self.input_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.input_tree.yview)
        self.input_tree.tag_configure('separator', background='#dddddd', foreground='#555555', font=("Segoe UI", 9, "bold"))
        self.input_tree.tag_configure('row_monthly', foreground='black') 
        self.input_tree.tag_configure('row_daily', foreground='#8B4513')
        
        self.input_tree.bind("<Double-1>", self._open_input_popup)
        ttk.Label(parent_tab, text="💡 ดับเบิลคลิกที่รายชื่อ เพื่อกรอก OT, โบนัส, ภาษี(ก้าวหน้า), และรายการหักอื่นๆ", foreground="gray").pack(pady=5)

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
        """เปิดหน้าต่างดูประวัติเงินเดือนย้อนหลัง"""
        win = tk.Toplevel(self)
        win.title("📜 ประวัติการจ่ายเงินเดือน (Payroll History)")
        win.geometry("1200x700")
        
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
        
        # --- Sheet แสดงข้อมูล ---
        sheet_frame = ttk.Frame(win)
        sheet_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        history_sheet = Sheet(sheet_frame, headers=None, theme="light blue")
        history_sheet.pack(fill="both", expand=True)
        history_sheet.enable_bindings("single_select", "row_select", "column_width_resize", "arrowkeys", "copy")

        # ฟังก์ชันโหลดข้อมูล
        def load_history():
            try:
                y_be = int(cb_year.get())
                y_ce = y_be - 543
                m_name = cb_month.get()
                m_int = self.MONTH_TO_INT[m_name]
            except: return

            # ดึงข้อมูลจาก DB
            records = hr_database.get_monthly_payroll_records(m_int, y_ce)
            
            if not records:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติการจ่ายเงินของงวด {m_name} {y_be}")
                history_sheet.set_sheet_data([])
                return

            # แปลงข้อมูลลงตาราง
            sheet_data = []
            pay_date_display = "-"
            
            for r in records:
                # ดึงวันที่จ่ายมาโชว์สักหน่อย (เอาจากคนแรกก็ได้)
                if pay_date_display == "-" and r.get('payment_date'):
                    pd = r['payment_date'] # date object
                    pay_date_display = f"{pd.day}/{pd.month}/{pd.year+543}"

                fullname = f"{r.get('fname','')} {r.get('lname','')}"
                row = [
                    r['emp_id'], fullname,
                    f"{r['base_salary']:,.2f}", 
                    f"{r['ot_pay']:,.2f}", 
                    f"{r['commission']:,.2f}", 
                    f"{r['total_income']:,.2f}", # รวมรับ
                    f"{r['sso_deduct']:,.2f}", 
                    f"{r['tax_deduct']:,.2f}", 
                    f"{r['total_deduct']:,.2f}", # รวมหัก
                    f"{r['net_salary']:,.2f}"     # สุทธิ
                ]
                sheet_data.append(row)
            
            # ตั้งชื่อหัวตาราง
            headers = [
                "รหัส", "ชื่อ-สกุล", 
                "เงินเดือน", "OT", "คอมมิชชั่น", "รวมรับ",
                "ประกันสังคม", "ภาษี", "รวมหัก", "สุทธิ (รับจริง)"
            ]
            history_sheet.headers(headers)
            history_sheet.set_sheet_data(sheet_data)
            
            # บอกวันที่จ่ายที่หัวหน้าต่าง
            win.title(f"📜 ประวัติการจ่ายเงินเดือน - งวด {m_name} {y_be} (จ่ายเมื่อ: {pay_date_display})")

        ttk.Button(top_frame, text="🔍 ค้นหาประวัติ", command=load_history).pack(side="left", padx=10)
        
        # โหลดครั้งแรกเลย
        load_history()

    def _open_input_popup(self, event):
        selection = self.input_tree.selection()
        if not selection: return
        
        emp_id = selection[0]
        emp_name = self.input_tree.item(emp_id, "values")[1]
        curr_data = self.payroll_inputs.get(emp_id, {})

        popup = tk.Toplevel(self)
        popup.title(f"บันทึกรายรับ/รายจ่าย - {emp_name}")
        popup.geometry("500x450")
        popup.transient(self)
        popup.grab_set()
        
        frame = ttk.Frame(popup, padding=20)
        frame.pack(fill="both", expand=True)
        
        entries = {}
        ttk.Label(frame, text="รายได้ (Addition)", font=("", 10, "bold"), foreground="green").grid(row=0, column=0, sticky="w", pady=(0,10))
        fields_inc = [
            ("ค่าล่วงเวลา (OT)", "ot"), 
            ("คอมมิชชั่น", "commission"), 
            ("Incentive", "incentive"), 
            ("เบี้ยขยัน", "diligence"), 
            ("โบนัส", "bonus"), 
            ("เงินได้อื่นๆ", "other_income")
        ]
        row = 1
        for label, key in fields_inc:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            ent = ttk.Entry(frame, width=20)
            ent.grid(row=row, column=1, sticky="w", padx=5)
            ent.insert(0, str(curr_data.get(key, 0.0)))
            entries[key] = ent
            row += 1
            
        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=15)
        row += 1
        
        ttk.Label(frame, text="รายการหัก (Deduction)", font=("", 10, "bold"), foreground="red").grid(row=row, column=0, sticky="w", pady=(0,10))
        row += 1
        fields_ded = [("ภาษี ณ ที่จ่าย", "tax"), ("กองทุนสำรองเลี้ยงชีพ", "provident_fund"),
                      ("หักเงินกู้ (กยศ.)ยืม", "loan"), ("หักอื่นๆ", "other_deduct")]
        for label, key in fields_ded:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            ent = ttk.Entry(frame, width=20)
            ent.grid(row=row, column=1, sticky="w", padx=5)
            ent.insert(0, str(curr_data.get(key, 0.0)))
            entries[key] = ent
            row += 1
            
        def save_popup():
            try:
                new_data = {}
                has_data = False
                for key, ent in entries.items():
                    val = float(ent.get() or 0)
                    new_data[key] = val
                    if val > 0: has_data = True
                self.payroll_inputs[emp_id] = new_data
                status = "✅ บันทึกแล้ว" if has_data else "-"
                self.input_tree.set(emp_id, column="status", value=status)
                popup.destroy()
            except ValueError:
                messagebox.showerror("Error", "กรุณากรอกตัวเลขเท่านั้น (ถ้าไม่มีให้ใส่ 0)")

        ttk.Button(frame, text="💾 บันทึก", command=save_popup, style="Success.TButton").grid(row=row+1, column=0, columnspan=2, pady=20)

    def _run_payroll_calculation(self):
        """
        ฟังก์ชันหลักคำนวณเงินเดือน (Main Payroll Engine) - ฉบับแก้ไข (Fix Separator Bug)
        - กรองแถว Separator ทิ้ง ไม่ให้นำมาคำนวณ
        """
        # --- 1. ตรวจสอบวันที่และการตั้งค่าเบื้องต้น ---
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            if not start_date or not end_date: 
                messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกวันที่ให้ครบ")
                return
            
            current_month = start_date.month
            current_year = start_date.year 
        except: 
            return

        # --- [แก้ไขจุดที่ 1.5] กรองรายชื่อ (ตัดเส้นแบ่งออก) ---
        all_items = self.input_tree.get_children()
        employee_ids = []
        
        for item_iid in all_items:
            # ตรวจสอบ Tag ของแถวนั้น
            tags = self.input_tree.item(item_iid, "tags")
            
            # ถ้ามี tag ชื่อ 'separator' ให้ข้ามไปเลย (ไม่เอามาคำนวณ)
            if 'separator' in tags:
                continue
                
            # ถ้าเป็นคนจริงๆ ให้เก็บ ID ไว้ (เพราะเราตั้ง iid=emp_id ตอน insert)
            employee_ids.append(item_iid)

        if not employee_ids:
            messagebox.showwarning("แจ้งเตือน", "ไม่พบรายชื่อพนักงานที่จะคำนวณ (กรุณาโหลดรายชื่อก่อน)")
            return

        # --- 2. โหลดค่า Config ---
        try:
            allowance_settings = hr_database.load_allowance_settings()
            taxable_map = { item['name']: item['is_taxable'] for item in allowance_settings }
        except:
            taxable_map = {} 
            
        try:
            sso_cfg = hr_database.load_sso_config(current_year) 
            sso_rate = sso_cfg.get('rate', 5.0) / 100.0
            sso_max_base = sso_cfg.get('max_salary', 15000)
            sso_min_base = sso_cfg.get('min_salary', 1650)
        except:
            sso_rate = 0.05
            sso_max_base = 15000
            sso_min_base = 1650

        # --- 3. เริ่ม Loop คำนวณทีละคน ---
        self.last_payroll_results = []
        sheet_data = []
        
        total_sum = {
            "base_salary": 0.0, "position": 0.0, "ot": 0.0, 
            "bonus": 0.0, "commission": 0.0, "other_income": 0.0,
            "welfare_taxable": 0.0, "welfare_nontaxable": 0.0,
            "total_income": 0.0, "net_salary": 0.0,
            "sso": 0.0, "tax": 0.0, "provident_fund": 0.0, 
            "loan": 0.0, "late": 0.0, "other_deduct": 0.0, "total_deduct": 0.0
        }

        for emp_id in employee_ids:
            user_in = self.payroll_inputs.get(emp_id, {})
            res = hr_database.calculate_payroll_for_employee(emp_id, start_date, end_date, user_in)
            
            if res:
                emp_info = hr_database.load_single_employee(emp_id)
                emp_name = self.input_tree.item(emp_id, "values")[1]
                res['name'] = emp_name
                
                is_resigned = False
                if emp_info and emp_info.get('status') in ['พ้นสภาพพนักงาน', 'ลาออก']:
                     is_resigned = True

                # คำนวณสวัสดิการ
                welfare_taxable_sum = 0.0
                welfare_nontaxable_sum = 0.0
                if emp_info:
                    w_flags = emp_info.get('welfare', [])
                    w_amounts = emp_info.get('welfare_amounts', [])
                    w_options = emp_info.get('welfare_options', [])
                    for idx, amt_str in enumerate(w_amounts):
                        if idx < len(w_flags) and w_flags[idx] and idx < len(w_options):
                            try:
                                amt = float(amt_str or 0)
                                if amt > 0:
                                    w_name = w_options[idx]
                                    if taxable_map.get(w_name, True):
                                        welfare_taxable_sum += amt
                                    else:
                                        welfare_nontaxable_sum += amt
                            except: pass

                # คำนวณประกันสังคม
                sso_wage_base = res['base_salary']
                if sso_wage_base > sso_max_base: sso_calc_base = sso_max_base
                elif sso_wage_base < sso_min_base: sso_calc_base = sso_min_base
                else: sso_calc_base = sso_wage_base
                
                current_sso = int((sso_calc_base * sso_rate) + 0.5)
                res['sso'] = current_sso 

                # คำนวณภาษี
                income_for_tax = (
                    res['base_salary'] + res['position_allowance'] + res['ot'] + 
                    res['bonus'] + res['incentive'] + res['diligence'] + 
                    res['other_income'] + welfare_taxable_sum
                )
                
                ytd_income, ytd_tax, ytd_sso = hr_database.get_ytd_summary(emp_id, current_year, current_month)
                
                pnd1_calc = self._calculate_smart_tax(
                    current_income=income_for_tax,
                    current_sso=res['sso'],
                    current_pfund=res['provident_fund'],
                    ytd_income=ytd_income,
                    ytd_tax_paid=ytd_tax,
                    ytd_sso=ytd_sso,
                    ytd_pfund=0, 
                    month_idx=current_month,
                    is_resigned=is_resigned,
                    other_allowances=0 
                )
                pnd3_calc = res['commission'] * 0.03
                res['pnd1'] = pnd1_calc
                res['pnd3'] = pnd3_calc
                res['tax'] = pnd1_calc + pnd3_calc

                # สรุปยอด
                res['total_income'] = (
                    income_for_tax + res['commission'] + 
                    welfare_nontaxable_sum + res.get('driving_allowance', 0)
                )
                res['total_deduct'] = (
                    res['sso'] + res['tax'] + res['provident_fund'] + 
                    res['loan'] + res['late_deduct'] + res['other_deduct']
                )
                res['net_salary'] = res['total_income'] - res['total_deduct']
                
                self.last_payroll_results.append(res)
                
                # ยอดรวม
                total_sum['base_salary'] += res['base_salary']
                total_sum['position'] += res['position_allowance']
                total_sum['ot'] += res['ot']
                total_sum['bonus'] += res['bonus']
                total_sum['commission'] += res['commission']
                total_sum['welfare_taxable'] += welfare_taxable_sum
                total_sum['welfare_nontaxable'] += welfare_nontaxable_sum
                total_sum['total_income'] += res['total_income']
                total_sum['sso'] += res['sso']
                total_sum['tax'] += res['tax']
                total_sum['provident_fund'] += res['provident_fund']
                total_sum['loan'] += res['loan']
                total_sum['late'] += res['late_deduct']
                total_sum['total_deduct'] += res['total_deduct']
                total_sum['net_salary'] += res['net_salary']

                # ข้อมูลลงตาราง
                display_other = (
                    res['other_income'] + welfare_taxable_sum + 
                    welfare_nontaxable_sum + res.get('driving_allowance', 0)
                )
                
                row_data = [
                    emp_id, emp_name,
                    f"{res['base_salary']:,.2f}", f"{res['position_allowance']:,.2f}",
                    f"{res['ot']:,.2f}", f"{res['commission']:,.2f}", 
                    f"{res.get('incentive',0):,.2f}", f"{res.get('diligence',0):,.2f}",
                    f"{res['bonus']:,.2f}", f"{display_other:,.2f}", 
                    f"{res.get('driving_allowance',0):,.2f}",
                    f"{res['total_income']:,.2f}", 
                    f"{res['sso']:,.2f}", f"{res['pnd1']:,.2f}", f"{res['pnd3']:,.2f}",
                    f"{res['provident_fund']:,.2f}", f"{res['loan']:,.2f}", 
                    f"{res['late_deduct']:,.2f}", f"{res['other_deduct']:,.2f}",
                    f"{res['total_deduct']:,.2f}", f"{res['net_salary']:,.2f}"    
                ]
                sheet_data.append(row_data)

        # --- 4. เพิ่มแถวสรุป (Total Row) ---
        display_total_other = (total_sum['other_income'] + total_sum['welfare_taxable'] + total_sum['welfare_nontaxable'])
        summary_row = [
            "TOTAL", "รวมทั้งสิ้น",
            f"{total_sum['base_salary']:,.2f}", f"{total_sum['position']:,.2f}",
            f"{total_sum['ot']:,.2f}", f"{total_sum['commission']:,.2f}", 
            "-", "-", f"{total_sum['bonus']:,.2f}",
            f"{display_total_other:,.2f}", "-",
            f"{total_sum['total_income']:,.2f}",
            f"{total_sum['sso']:,.2f}", f"{total_sum['tax']:,.2f}", "-",
            f"{total_sum['provident_fund']:,.2f}", f"{total_sum['loan']:,.2f}",
            f"{total_sum['late']:,.2f}", f"{total_sum['other_deduct']:,.2f}",
            f"{total_sum['total_deduct']:,.2f}", f"{total_sum['net_salary']:,.2f}"
        ]
        sheet_data.append(summary_row)

        # --- 5. อัปเดตหน้าจอ ---
        # 5.1 ตั้งค่าหัวตาราง (Headers)
        headers = [
            "รหัส", "ชื่อ-สกุล",
            "เงินเดือน", "ค่าตำแหน่ง", "OT", "คอมมิชชั่น", "Incentive", "เบี้ยขยัน", "โบนัส", "อื่นๆ(รับ)", "ค่าเที่ยว",
            "รวมรับ",
            "ประกันสังคม", "ภ.ง.ด.1", "ภ.ง.ด.3", "กองทุนฯ", "เงินกู้", "มาสาย/ลา", "อื่นๆ(หัก)",
            "รวมหัก",
            "สุทธิ"
        ]
        self.results_sheet.headers(headers) 

        # 5.2 ใส่ข้อมูล
        self.results_sheet.set_sheet_data(sheet_data)
        
        # 5.3 ใส่สี (Highlight)
        self.results_sheet.highlight_columns(columns=list(range(2, 11)), bg="#e6f7ff", fg="black") # รายรับ
        self.results_sheet.highlight_columns(columns=list(range(12, 19)), bg="#fff7e6", fg="black") # รายหัก
        self.results_sheet.highlight_columns(columns=[20], bg="#ffffcc", fg="black") # สุทธิ
        
        # Highlight แถว Total
        if sheet_data:
            last_row_idx = len(sheet_data) - 1
            self.results_sheet.highlight_rows(rows=[last_row_idx], bg="#ccffcc", fg="black")

        # เปิดปุ่มใช้งาน
        self.export_btn.config(state="normal")
        self.print_btn.config(state="normal")
        self.pnd1_btn.config(state="normal")
        self.pnd3_btn.config(state="normal")
        self.save_db_btn.config(state="normal")
        self.sso_btn.config(state="normal")
        
        self.notebook.select(self.tab2)
        messagebox.showinfo("สำเร็จ", f"คำนวณเงินเดือนเรียบร้อยแล้ว\n(ใช้เรทประกันสังคมปี {current_year+543})")

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
                "total_deduct": "รวมหัก", "net_salary": "สุทธิ"
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
        for item in self.last_payroll_results:
            # ส่ง pay_date ที่เลือกเข้าไปบันทึก
            ok = hr_database.save_monthly_payroll(item['emp_id'], m_int, y_ce, pay_date, item)
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

        pay_date = pay_date_str
        try:
            s_date = self.start_date_entry.get_date()
            month_th = list(self.THAI_MONTHS.values())[s_date.month - 1]
            period_str = f"{month_th} {s_date.year + 543}"
        except:
            period_str = "-"

        def fmt_money(val):
            return f"{val:,.2f}" if isinstance(val, (int, float)) and val > 0 else "-"

        # --- ฟังก์ชันวาดสลิป (Nested Function) ---
        def draw_slip_form(current_data, start_y, copy_label):
            # (รับ current_data เข้ามาเป็น argument เพื่อความชัวร์)
            
            # 1. Header Info
            if os.path.exists(logo_path):
                pdf.image(logo_path, x=15, y=start_y + 5, w=20)
            
            pdf.set_xy(0, start_y + 5)
            pdf.set_font("THSarabun", "B", 20)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(0, 8, "บริษัท เอไพร์ม พลัส จํากัด", ln=True, align='C')
            pdf.set_font("THSarabun", "B", 16)
            pdf.cell(0, 8, f"ใบจ่ายเงินเดือน (Pay Slip) {copy_label}", ln=True, align='C')

            # 2. Employee Box
            box_top = start_y + 22
            box_h = 16
            pdf.set_draw_color(0)
            pdf.set_line_width(0.1)
            pdf.rect(10, box_top, 190, box_h) 
            pdf.line(10, box_top + 8, 200, box_top + 8)
            pdf.line(40, box_top, 40, box_top + 16)
            pdf.line(105, box_top, 105, box_top + 16)
            pdf.line(135, box_top, 135, box_top + 16)

            pdf.set_font("THSarabun", "", 14)
            pdf.set_xy(10, box_top + 1); pdf.cell(30, 6, "  รหัสพนักงาน :", border=0)
            pdf.set_xy(40, box_top + 1); pdf.cell(65, 6, f"  {current_data.get('emp_id', '-')}", border=0)
            pdf.set_xy(105, box_top + 1); pdf.cell(30, 6, "  ตำแหน่ง :", border=0)
            pdf.set_xy(135, box_top + 1); pdf.cell(65, 6, f"  {current_data.get('position','-')}", border=0)
            
            pdf.set_xy(10, box_top + 9); pdf.cell(30, 6, "  ชื่อ - นามสกุล :", border=0)
            
            # --- (จุดที่เคย Error: เรียกใช้ชื่ออย่างปลอดภัย) ---
            display_name = current_data.get('name', '')
            if not display_name:
                display_name = f"{current_data.get('fname', '')} {current_data.get('lname', '')}".strip()
            if not display_name: 
                display_name = "ไม่ระบุชื่อ"
            # --------------------------------------------------
            
            pdf.set_xy(40, box_top + 9); pdf.cell(65, 6, f"  {display_name}", border=0)
            
            pdf.set_xy(105, box_top + 9); pdf.cell(30, 6, "  แผนก :", border=0)
            pdf.set_xy(135, box_top + 9); pdf.cell(65, 6, f"  {current_data.get('department','-')}", border=0)

            pdf.set_xy(10, box_top + 18)
            pdf.cell(95, 6, f"วันที่จ่าย : {pay_date}")
            pdf.set_xy(105, box_top + 18)
            pdf.cell(95, 6, f"ค่าจ้างเดือน : {period_str}")

            # --- 3. Table Header ---
            tbl_top = box_top + 28
            row_h = 7
            
            pdf.rect(10, tbl_top, 95, 8)   # กรอบซ้าย
            pdf.rect(105, tbl_top, 95, 8)  # กรอบขวา
            
            pdf.set_font("THSarabun", "B", 16)
            pdf.set_xy(10, tbl_top)
            pdf.cell(95, 8, "เงินได้ (Earnings)", border=0, align='C')
            
            pdf.set_xy(105, tbl_top)
            pdf.cell(95, 8, "เงินหัก (Deductions)", border=0, align='C')

            # --- 4. Data Rows ---
            body_top = tbl_top + 8 
            max_rows = 8
            
            incomes = [ 
                ("เงินเดือน", current_data.get('base_salary', 0)), 
                ("ค่าตำแหน่ง", current_data.get('position_allowance', 0)), 
                ("ค่าล่วงเวลา", current_data.get('ot', 0)), 
                ("คอมมิชชั่น", current_data.get('commission', 0)), 
                ("โบนัส", current_data.get('bonus', 0)), 
                ("เงินได้อื่นๆ", current_data.get('other_income', 0))
            ]
            deductions = [
                ("ประกันสังคม", current_data.get('sso', 0)),
                ("ภาษีเงินได้", 0),
                ("  - ภ.ง.ด. 1", current_data.get('pnd1', 0)),
                ("  - ภ.ง.ด. 3", current_data.get('pnd3', 0)),
                ("สำรองเลี้ยงชีพ", current_data.get('provident_fund', 0)), 
                ("หักเงินกู้ยืม", current_data.get('loan', 0)), 
                ("ขาด/ลา/สาย", current_data.get('late_deduct', 0)), 
                ("หักอื่นๆ", current_data.get('other_deduct', 0))
            ]

            pdf.set_font("THSarabun", "", 14)
            
            for i in range(max_rows):
                curr_y = body_top + (i * row_h)
                
                pdf.rect(10, curr_y, 190, row_h)
                pdf.line(105, curr_y, 105, curr_y + row_h)
                pdf.line(90, curr_y, 90, curr_y + row_h)
                pdf.line(185, curr_y, 185, curr_y + row_h)
                
                if i < len(incomes):
                    label, val = incomes[i]
                    pdf.set_xy(10, curr_y)
                    pdf.cell(55, row_h, f"  {label}", border=0, align='L')
                    pdf.set_xy(65, curr_y)
                    pdf.cell(25, row_h, fmt_money(val), border=0, align='R')
                    pdf.set_xy(90, curr_y)
                    pdf.cell(15, row_h, "บาท", border=0, align='C')

                if i < len(deductions):
                    l2, v2 = deductions[i]
                    pdf.set_xy(105, curr_y)
                    pdf.cell(55, row_h, f"  {l2}", border=0, align='L')
                    
                    show_val = fmt_money(v2)
                    if show_val != "-":
                        pdf.set_xy(160, curr_y)
                        pdf.cell(25, row_h, show_val, border=0, align='R')
                        pdf.set_xy(185, curr_y)
                        pdf.cell(15, row_h, "บาท", border=0, align='C')

            # 5. Totals
            totals_y = body_top + (max_rows * row_h)
            
            pdf.set_fill_color(240, 240, 240) 
            pdf.rect(10, totals_y, 190, 7, 'F')
            pdf.rect(10, totals_y, 190, 7)
            pdf.line(105, totals_y, 105, totals_y + 7)

            pdf.set_font("THSarabun", "B", 14)
            
            pdf.set_xy(10, totals_y)
            pdf.cell(55, 7, "  รวมเงินได้", 0, 0, 'L')
            pdf.set_xy(65, totals_y)
            pdf.cell(25, 7, fmt_money(current_data.get('total_income', 0)), 0, 0, 'R')
            pdf.set_xy(90, totals_y)
            pdf.cell(15, 7, "บาท", 0, 0, 'C')
            
            pdf.set_xy(105, totals_y)
            pdf.cell(55, 7, "  รวมเงินหัก", 0, 0, 'L')
            pdf.set_xy(160, totals_y)
            pdf.cell(25, 7, fmt_money(current_data.get('total_deduct', 0)), 0, 0, 'R')
            pdf.set_xy(185, totals_y)
            pdf.cell(15, 7, "บาท", 0, 0, 'C')

            # 6. Net Salary
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

            # 7. Signature
            sig_y = net_y + 15
            pdf.set_font("THSarabun", "", 12)
            pdf.set_xy(10, sig_y)
            pdf.cell(60, 5, "ลงชื่อผู้จ่ายเงิน ...........................................", 0, 0, 'L')
            pdf.set_xy(105, sig_y)
            pdf.cell(60, 5, "ลงชื่อผู้รับเงิน ...........................................", 0, 0, 'L')

        # --- Loop Generate Page ---
        for data in data_list:
            # (!!! เพิ่ม Logic กู้คืนชื่อ !!!)
            if 'name' not in data or not data['name']:
                # พยายามสร้างชื่อจาก fname/lname ถ้ามี
                if 'fname' in data and 'lname' in data:
                    data['name'] = f"{data['fname']} {data['lname']}".strip()
                else:
                    # ถ้าไม่มีจริงๆ ให้วิ่งไปดึงจาก DB
                    emp_id = data.get('emp_id')
                    if emp_id:
                        info = hr_database.load_single_employee(emp_id)
                        if info:
                            data['name'] = f"{info.get('fname','')} {info.get('lname','')}".strip()
                            data['position'] = info.get('position', '-')
                            data['department'] = info.get('department', '-')

            pdf.add_page()
            
            # ส่ง data เข้าไปวาด (แทนที่จะใช้ตัวแปร global loop)
            draw_slip_form(data, 5, "(ต้นฉบับ)")
            
            pdf.set_draw_color(100)
            pdf.dashed_line(5, 148, 205, 148, dash_length=2, space_length=2)
            pdf.set_font("THSarabun", "", 10)
            pdf.text(185, 147, "ตัดตามรอยประ")
            
            draw_slip_form(data, 153, "(สำเนา)")
            
            pdf.set_xy(10, 290)
            pdf.set_font("THSarabun", "", 10)
            pdf.cell(0, 5, f"พิมพ์เมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align='R')

        pdf.output(filepath)
            
    # (ฟังก์ชัน Helper เดิม)
    # (ใน payroll_module.py) -> แทนที่ฟังก์ชันนี้ได้เลยครับ

    def _load_employees_to_input_tree(self):
        """โหลดรายชื่อแบบจัดกลุ่ม (Grouped): รายเดือนไว้บน / รายวันไว้ล่าง"""
        
        # 1. ล้างตารางเก่า
        for item in self.input_tree.get_children(): 
            self.input_tree.delete(item)
            
        self.payroll_inputs = {} 
        
        # 2. ดึงข้อมูลทั้งหมด
        emps = hr_database.load_all_employees()
        filter_type = self.emp_type_var.get() # "ทั้งหมด", "รายเดือน", "รายวัน"
        
        # 3. เตรียมตะกร้าแยกกลุ่ม
        list_monthly = []
        list_daily = []

        for emp in emps:
            # ข้ามคนที่ลาออก
            if emp.get('status') in ['พ้นสภาพพนักงาน', 'ลาออก']: continue

            # เช็คประเภท
            emp_type = emp.get('emp_type', 'รายเดือน') # ถ้าไม่มีให้เหมาเป็นรายเดือน
            
            # จัดลงตะกร้า
            if "รายวัน" in emp_type:
                list_daily.append(emp)
            else:
                list_monthly.append(emp)

        # 4. ฟังก์ชันย่อยสำหรับใส่ข้อมูลลงตาราง
        def insert_emp(emp_obj, row_tag):
            # เช็คสถานะการกรอก (Logic เดิม)
            status_text = "-"
            if emp_obj['id'] in self.payroll_inputs:
                status_text = "✅ บันทึกร่างแล้ว"
            
            self.input_tree.insert("", "end", iid=emp_obj['id'], values=(
                emp_obj['id'], 
                f"{emp_obj['fname']} {emp_obj['lname']}", 
                emp_obj.get('emp_type', '-'), 
                status_text
            ), tags=(row_tag,))

        # 5. เริ่มแสดงผลตาม Filter
        
        # --- กรณีเลือกดู "รายเดือน" หรือ "ทั้งหมด" ---
        if filter_type in ["ทั้งหมด", "รายเดือน"]:
            # เรียงตามรหัสพนักงาน (หรือชื่อ) ก่อนแสดง
            list_monthly.sort(key=lambda x: x['id'])
            
            for emp in list_monthly:
                insert_emp(emp, 'row_monthly')

        # --- กรณีเลือก "ทั้งหมด" และมีข้อมูลทั้ง 2 กลุ่ม -> สร้างเส้นคั่น ---
        if filter_type == "ทั้งหมด" and list_monthly and list_daily:
            # ใส่แถว Dummy เป็นเส้นคั่น
            self.input_tree.insert("", "end", values=(
                "-----", 
                "⬇⬇⬇  พนักงานรายวัน (Daily)  ⬇⬇⬇", 
                "-----", 
                ""
            ), tags=('separator',))

        # --- กรณีเลือกดู "รายวัน" หรือ "ทั้งหมด" ---
        if filter_type in ["ทั้งหมด", "รายวัน"]:
            # เรียงตามรหัส
            list_daily.sort(key=lambda x: x['id'])
            
            for emp in list_daily:
                insert_emp(emp, 'row_daily')

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
        if is_resigned or month_idx == 12:
            tax_this_month = annual_tax - ytd_tax_paid
        else:
            # (ภาษีทั้งปี / 12 * จำนวนเดือนที่ผ่านไป) - ภาษีที่จ่ายแล้ว
            expected_tax_ytd = (annual_tax / 12) * month_idx
            tax_this_month = expected_tax_ytd - ytd_tax_paid

        return max(0, tax_this_month)
    
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
        """ปุ่มสำหรับดึงข้อมูลจาก A+ Smart"""
        
        # 1. เช็คว่าเลือกเดือน/ปี หรือยัง
        y_ce, m_int = self._get_selected_dates()
        if not y_ce: return

        if not messagebox.askyesno("ยืนยัน", f"ต้องการดึงยอดคอมมิชชั่นประจำเดือน {m_int}/{y_ce}\nจากระบบ A+ Smart (192.168.1.51) หรือไม่?"):
            return

        updated_count = 0
        
        # 2. วนลูปพนักงานทุกคนในตาราง
        for item_id in self.input_tree.get_children():
            emp_id = self.input_tree.item(item_id, "values")[0] # รหัสพนักงาน
            
            # 3. ไปดึงยอดเงิน
            comm_amt = hr_database.get_commission_from_asmart(emp_id, m_int, y_ce)
            
            if comm_amt > 0:
                # 4. อัปเดตเข้าตัวแปร payroll_inputs
                if emp_id not in self.payroll_inputs:
                    self.payroll_inputs[emp_id] = {}
                
                # บันทึกยอดลงช่อง commission
                self.payroll_inputs[emp_id]['commission'] = comm_amt
                
                # อัปเดตสถานะในตารางให้รู้ว่าดึงมาแล้ว
                self.input_tree.set(emp_id, column="status", value=f"✅ A+ Smart: {comm_amt:,.2f}")
                updated_count += 1

        if updated_count > 0:
            messagebox.showinfo("สำเร็จ", f"ดึงข้อมูลคอมมิชชั่นเรียบร้อย {updated_count} รายการ")
        else:
            messagebox.showinfo("ไม่พบข้อมูล", "ไม่พบยอดคอมมิชชั่นในช่วงเวลานี้ (หรือรหัสพนักงานไม่ตรงกัน)")

    def _open_daily_timesheet(self):
        """เปิดหน้าต่างบันทึกงานรายวัน (Daily Timesheet)"""
        selection = self.input_tree.selection()
        if not selection:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกพนักงานที่ต้องการบันทึกงานรายวัน")
            return
            
        emp_id = self.input_tree.item(selection[0], "values")[0] # รหัสพนักงาน
        
        # เช็คว่าเลือกเดือน/ปีหรือยัง
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