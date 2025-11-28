# (ไฟล์: payroll_module.py)
# (เวอร์ชัน V15.0 - เพิ่มฟังก์ชันพิมพ์สลิป PDF พร้อมโลโก้)

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

class PayrollModule(ttk.Frame):

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
        btn_frame = ttk.Frame(parent_tab)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="🔄 โหลดรายชื่อพนักงาน", command=self._load_employees_to_input_tree).pack(side="left")

        ttk.Button(btn_frame, text="💰 ดึงยอด A+ Smart", 
                   command=self._sync_commission_from_asmart,
                   style="Primary.TButton").pack(side="left", padx=10)
        
        ttk.Button(btn_frame, text="📅 บันทึกงานรายวัน (Timesheet)", 
                   command=self._open_daily_timesheet).pack(side="left", padx=10)
        
        tree_container = ttk.Frame(parent_tab)
        tree_container.pack(fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(tree_container, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        self.input_tree = ttk.Treeview(
            tree_container,
            columns=("id", "name", "status"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            height=15
        )
        self.input_tree.heading("id", text="รหัส")
        self.input_tree.heading("name", text="ชื่อ-นามสกุล")
        self.input_tree.heading("status", text="สถานะการป้อนข้อมูล")
        
        self.input_tree.column("id", width=80, anchor="center")
        self.input_tree.column("name", width=250, anchor="w")
        self.input_tree.column("status", width=200, anchor="center")
        
        self.input_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=self.input_tree.yview)
        
        self.input_tree.bind("<Double-1>", self._open_input_popup)
        ttk.Label(parent_tab, text="💡 ดับเบิลคลิกที่รายชื่อ เพื่อกรอก OT, โบนัส, ภาษี(ก้าวหน้า), และรายการหักอื่นๆ", foreground="gray").pack(pady=5)

    def _build_results_tab(self, parent_tab):
        btn_frame = ttk.Frame(parent_tab)
        btn_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(btn_frame, text="🚀 คำนวณเงินเดือน", command=self._run_payroll_calculation, style="Success.TButton").pack(side="left")
        self.export_btn = ttk.Button(btn_frame, text="📄 Export Excel", command=self._export_payroll_to_excel, state="disabled")
        self.export_btn.pack(side="left", padx=10)
        
        self.print_btn = ttk.Button(btn_frame, text="🖨️ พิมพ์สลิป (PDF)", command=self._print_selected_payslip, state="disabled")
        self.print_btn.pack(side="left", padx=10)

        self.pnd1_btn = ttk.Button(btn_frame, text="🏛️ ใบปะหน้า ภ.ง.ด.1", command=self._print_pnd1_summary, state="disabled")
        self.pnd1_btn.pack(side="left", padx=10)

        self.pnd3_btn = ttk.Button(btn_frame, text="🏛️ ใบปะหน้า ภ.ง.ด.3", command=self._print_pnd3_summary, state="disabled")
        self.pnd3_btn.pack(side="left", padx=10)

        self.save_db_btn = ttk.Button(btn_frame, text="💾 บันทึกงวดบัญชี (DB)", command=self._save_payroll_to_database, state="disabled")
        self.save_db_btn.pack(side="left", padx=10)

        ttk.Button(btn_frame, text="📜 ดูประวัติย้อนหลัง", command=self._open_history_window).pack(side="left", padx=10)

        self.pnd1k_btn = ttk.Button(btn_frame, text="📄 ภ.ง.ด.1ก (รายปี)", command=self._export_pnd1k_excel)
        self.pnd1k_btn.pack(side="left", padx=10)

        self.pnd1k_pdf_btn = ttk.Button(btn_frame, text="📄 ภ.ง.ด.1ก (PDF)", command=self._print_pnd1k_pdf)
        self.pnd1k_pdf_btn.pack(side="left", padx=5)

        self.email_req_btn = ttk.Button(btn_frame, text="📧 ขอส่งสลิป (Email)", command=self._request_email_approval, state="disabled")
        self.email_req_btn.pack(side="left", padx=10)

        if self.current_user['role'] == 'approver':
            self.email_approve_btn = ttk.Button(btn_frame, text="✅ อนุมัติการส่งเมล", command=self._open_email_approval_window)
            self.email_approve_btn.pack(side="left", padx=10)

        # --- ส่วนแสดงผลตาราง (ใช้ Sheet อันเดียว) ---
        sheet_frame = ttk.Frame(parent_tab)
        sheet_frame.pack(fill="both", expand=True)
        
        self.results_sheet = Sheet(sheet_frame,
                                   show_x_scrollbar=True,
                                   show_y_scrollbar=True,
                                   headers=None,
                                   theme="light blue"
                                  )
        self.results_sheet.pack(fill="both", expand=True)
        self.results_sheet.enable_bindings("single", "row_select", "column_width_resize", "arrowkeys", "copy")
    
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
        win.title("📜 ประวัติเงินเดือนย้อนหลัง (Payroll History)")
        win.geometry("1200x700")
        
        # --- Filter Frame ---
        top_frame = ttk.Frame(win, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="เลือกงวด:").pack(side="left")
        
        # ปี
        current_year = datetime.now().year + 543
        years = [str(y) for y in range(current_year, current_year-5, -1)]
        cb_year = ttk.Combobox(top_frame, values=years, width=6, state="readonly")
        cb_year.set(current_year)
        cb_year.pack(side="left", padx=5)
        
        # เดือน
        months = list(self.THAI_MONTHS.values())
        cb_month = ttk.Combobox(top_frame, values=months, width=10, state="readonly")
        cb_month.set(self.THAI_MONTHS[datetime.now().month])
        cb_month.pack(side="left", padx=5)
        
        # --- Sheet แสดงข้อมูล ---
        sheet_frame = ttk.Frame(win)
        sheet_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        history_sheet = Sheet(sheet_frame,
                              headers=None,
                              theme="light blue")
        history_sheet.pack(fill="both", expand=True)
        history_sheet.enable_bindings("single", "row_select", "column_width_resize", "arrowkeys", "copy")

        # ฟังก์ชันโหลดข้อมูล
        def load_history():
            try:
                y_be = int(cb_year.get())
                y_ce = y_be - 543
                m_name = cb_month.get()
                m_int = self.MONTH_TO_INT[m_name]
            except: return

            records = hr_database.get_monthly_payroll_records(m_int, y_ce)
            
            if not records:
                messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติเงินเดือนของงวด {m_name} {y_be}")
                history_sheet.set_sheet_data([])
                return

            # แปลงข้อมูลลงตาราง
            sheet_data = []
            for r in records:
                fullname = f"{r.get('fname','')} {r.get('lname','')}"
                row = [
                    r['emp_id'], fullname,
                    f"{r['base_salary']:,.2f}", f"{r['position_allowance']:,.2f}",
                    f"{r['ot_pay']:,.2f}", f"{r['commission']:,.2f}", f"{r['bonus']:,.2f}",
                    f"{r['other_income']:,.2f}", f"{r['driving_allowance']:,.2f}",
                    f"{r['total_income']:,.2f}",
                    f"{r['sso_deduct']:,.2f}", f"{r['tax_deduct']:,.2f}", 
                    f"{r['provident_fund']:,.2f}", f"{r['loan_deduct']:,.2f}",
                    f"{r['late_deduct']:,.2f}", f"{r['other_deduct']:,.2f}",
                    f"{r['total_deduct']:,.2f}",
                    f"{r['net_salary']:,.2f}"
                ]
                sheet_data.append(row)
            
            headers = [
                "รหัส", "ชื่อ-สกุล", 
                "เงินเดือน", "ค่าตำแหน่ง", "OT", "คอมฯ", "โบนัส", "อื่นๆ(รับ)", "ค่าเที่ยว", "รวมรับ",
                "ประกันสังคม", "ภาษี", "กองทุนฯ", "เงินกู้", "ขาด/สาย", "อื่นๆ(หัก)", "รวมหัก", "สุทธิ"
            ]
            history_sheet.headers(headers)
            history_sheet.set_sheet_data(sheet_data)
            
            # ใส่สีสวยๆ เหมือนหน้าหลัก
            history_sheet.highlight_columns(columns=list(range(2, 10)), bg="#e6f7ff", fg="black") # ฟ้า (รายได้)
            history_sheet.highlight_columns(columns=list(range(10, 17)), bg="#fff7e6", fg="black") # ส้ม (รายหัก)
            history_sheet.highlight_columns(columns=[17], bg="#ffffcc", fg="black") # เหลือง (สุทธิ)

        ttk.Button(top_frame, text="🔍 ค้นหา", command=load_history).pack(side="left", padx=10)

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
        fields_inc = [("ค่าล่วงเวลา (OT)", "ot"), ("คอมมิชชั่น", "commission"), 
                      ("โบนัส", "bonus"), ("เงินได้อื่นๆ", "other_income")]
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
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            if not start_date or not end_date: 
                messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกวันที่ให้ครบ")
                return
        except: return

        employee_ids = self.input_tree.get_children()
        if not employee_ids:
            messagebox.showwarning("แจ้งเตือน", "กรุณาโหลดรายชื่อพนักงานก่อน")
            return

        # ล้างข้อมูลเก่า
        self.last_payroll_results = []
        sheet_data = []
        
        # --- ตัวแปรสำหรับเก็บยอดรวม (Grand Total) ---
        total_sum = {
            "base_salary": 0.0, "position_allowance": 0.0,
            "ot": 0.0, "commission": 0.0, "bonus": 0.0, 
            "other_income": 0.0, "driving_allowance": 0.0, "total_income": 0.0,
            "sso": 0.0, "pnd1": 0.0, "pnd3": 0.0, "provident_fund": 0.0,
            "loan": 0.0, "late_deduct": 0.0, "other_deduct": 0.0, "total_deduct": 0.0,
            "net_salary": 0.0
        }

        for emp_id in employee_ids:
            user_inputs = self.payroll_inputs.get(emp_id, {})
            res = hr_database.calculate_payroll_for_employee(emp_id, start_date, end_date, user_inputs)
            
            if res:
                emp_info = hr_database.load_single_employee(emp_id)
                name = f"{emp_info.get('fname', '')} {emp_info.get('lname', '')}"
                
                income_for_pnd1 = res['total_income']
                pnd3_calc = 0.0
                
                # คำนวณภาษีแบบก้าวหน้า
                pnd1_calc = self._calculate_smart_tax(income_for_pnd1, res['sso'])
                
                res['pnd1'] = pnd1_calc
                res['pnd3'] = pnd3_calc
                res['tax'] = pnd1_calc + pnd3_calc
                
                res['total_deduct'] = (
                    res['sso'] + res['tax'] + res['provident_fund'] + 
                    res['loan'] + res['late_deduct'] + res['other_deduct']
                )
                res['net_salary'] = res['total_income'] - res['total_deduct']
                self.last_payroll_results.append(res)

                # --- บวกยอดรวมเข้ากองกลาง ---
                total_sum["base_salary"] += res.get('base_salary', 0)
                total_sum["position_allowance"] += res.get('position_allowance', 0)
                total_sum["ot"] += res.get('ot', 0)
                total_sum["commission"] += res.get('commission', 0)
                total_sum["bonus"] += res.get('bonus', 0)
                total_sum["other_income"] += res.get('other_income', 0)
                total_sum["driving_allowance"] += res.get('driving_allowance', 0)
                total_sum["total_income"] += res.get('total_income', 0)
                
                total_sum["sso"] += res.get('sso', 0)
                total_sum["pnd1"] += res.get('pnd1', 0)
                total_sum["pnd3"] += res.get('pnd3', 0)
                total_sum["provident_fund"] += res.get('provident_fund', 0)
                total_sum["loan"] += res.get('loan', 0)
                total_sum["late_deduct"] += res.get('late_deduct', 0)
                total_sum["other_deduct"] += res.get('other_deduct', 0)
                total_sum["total_deduct"] += res.get('total_deduct', 0)
                
                total_sum["net_salary"] += res.get('net_salary', 0)

                # สร้างแถวข้อมูล (รายคน)
                row = [
                    emp_id, name,
                    f"{res['base_salary']:,.2f}", f"{res['position_allowance']:,.2f}",
                    f"{res['ot']:,.2f}", f"{res['commission']:,.2f}", f"{res['bonus']:,.2f}", 
                    f"{res['other_income']:,.2f}", f"{res.get('driving_allowance', 0):,.2f}",
                    f"{res['total_income']:,.2f}", 
                    f"{res['sso']:,.2f}", f"{res['pnd1']:,.2f}", f"{res['pnd3']:,.2f}",
                    f"{res['provident_fund']:,.2f}", f"{res['loan']:,.2f}", 
                    f"{res['late_deduct']:,.2f}", f"{res['other_deduct']:,.2f}",
                    f"{res['total_deduct']:,.2f}", 
                    f"{res['net_salary']:,.2f}"    
                ]
                sheet_data.append(row)

        # --- (!!! เพิ่มแถวสรุปยอดรวม - บรรทัดสุดท้าย !!!) ---
        summary_row = [
            "TOTAL", "รวมทั้งสิ้น",
            f"{total_sum['base_salary']:,.2f}", f"{total_sum['position_allowance']:,.2f}",
            f"{total_sum['ot']:,.2f}", f"{total_sum['commission']:,.2f}", f"{total_sum['bonus']:,.2f}",
            f"{total_sum['other_income']:,.2f}", f"{total_sum['driving_allowance']:,.2f}",
            f"{total_sum['total_income']:,.2f}",
            f"{total_sum['sso']:,.2f}", f"{total_sum['pnd1']:,.2f}", f"{total_sum['pnd3']:,.2f}",
            f"{total_sum['provident_fund']:,.2f}", f"{total_sum['loan']:,.2f}",
            f"{total_sum['late_deduct']:,.2f}", f"{total_sum['other_deduct']:,.2f}",
            f"{total_sum['total_deduct']:,.2f}",
            f"{total_sum['net_salary']:,.2f}"
        ]
        sheet_data.append(summary_row)

        # ตั้งค่า Sheet
        headers = [
            "รหัส", "ชื่อ-สกุล", 
            "เงินเดือน", "ค่าตำแหน่ง", "OT", "คอมฯ", "โบนัส", "อื่นๆ(รับ)", "ค่าเที่ยว", "รวมรับ",
            "ประกันสังคม", "ภ.ง.ด.1", "ภ.ง.ด.3", "กองทุนฯ", "เงินกู้", "ขาด/สาย", "อื่นๆ(หัก)", "รวมหัก",
            "สุทธิ"
        ]
        self.results_sheet.headers(headers)
        self.results_sheet.set_sheet_data(sheet_data)
        
        # --- ใส่สีคอลัมน์ ---
        self.results_sheet.highlight_columns(columns=list(range(2, 10)), bg="#e6f7ff", fg="black") # ฟ้าอ่อน
        self.results_sheet.highlight_columns(columns=list(range(10, 18)), bg="#fff7e6", fg="black") # ส้มอ่อน
        self.results_sheet.highlight_columns(columns=[18], bg="#ffffcc", fg="black") # เหลืองอ่อน
        
        # --- (!!! ใส่สีเขียวให้แถวสุดท้าย !!!) ---
        last_row_idx = len(sheet_data) - 1
        self.results_sheet.highlight_rows(rows=[last_row_idx], bg="#ccffcc", fg="black") # สีเขียวอ่อน
        # -------------------------------------

        # เปิดปุ่ม
        self.export_btn.config(state="normal")
        self.print_btn.config(state="normal")
        if hasattr(self, 'save_db_btn'): 
            self.save_db_btn.config(state="normal")
        self.pnd1_btn.config(state="normal")
        self.pnd3_btn.config(state="normal")
        self.email_req_btn.config(state="normal")
        
        self.notebook.select(self.tab2)
        messagebox.showinfo("สำเร็จ", "คำนวณเงินเดือนเรียบร้อยแล้ว")

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
                "ot": "OT", "commission": "คอมมิชชั่น", "bonus": "โบนัส", "other_income": "อื่นๆ(รับ)",
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
        """บันทึกผลคำนวณปัจจุบันลงฐานข้อมูล (payroll_records)"""
        if not self.last_payroll_results:
            messagebox.showwarning("เตือน", "ไม่มีข้อมูลให้บันทึก กรุณาคำนวณก่อน")
            return
            
        # ดึงเดือน/ปี ที่คำนวณ
        y_ce, m_int = self._get_selected_dates()
        if not y_ce: return
        
        month_name = list(self.THAI_MONTHS.values())[m_int - 1]
        
        if not messagebox.askyesno("ยืนยันการบันทึก", 
                                   f"คุณต้องการบันทึกงวดเงินเดือน {month_name} {y_ce+543}\n"
                                   f"จำนวน {len(self.last_payroll_results)} รายการ ลงฐานข้อมูลใช่หรือไม่?\n\n"
                                   f"(หากมีข้อมูลเก่าของเดือนนี้ ระบบจะบันทึกทับ)"):
            return
            
        success_count = 0
        pay_date = datetime.now().date() # ใช้วันที่ปัจจุบันเป็นวันที่จ่าย
        
        for item in self.last_payroll_results:
            # บันทึกทีละคน
            ok = hr_database.save_monthly_payroll(item['emp_id'], m_int, y_ce, pay_date, item)
            if ok: success_count += 1
            
        messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลเรียบร้อย {success_count} รายการ")

    # --- (!!! ส่วนใหม่: พิมพ์สลิปเงินเดือน PDF !!!) ---
    
    def _print_selected_payslip(self):
        # (ดึง Index แถวที่เลือกจาก Sheet)
        selected_indexes = self.results_sheet.get_selected_rows(return_tuple=True)
        
        if not selected_indexes:
            if not messagebox.askyesno("พิมพ์ทั้งหมด?", "คุณไม่ได้เลือกพนักงาน\nต้องการพิมพ์สลิปของ 'ทุกคน' ในรายการหรือไม่?"):
                return
            data_to_print = self.last_payroll_results
            filename_prefix = "Payslip_All"
        else:
            # ดึงข้อมูลจาก list ตาม index ที่เลือก
            data_to_print = [self.last_payroll_results[i] for i in selected_indexes]
            filename_prefix = f"Payslip_{data_to_print[0]['emp_id']}" if len(data_to_print)==1 else "Payslip_Selected"

        if not data_to_print: return
        
        # (ส่วนบันทึกไฟล์และเรียก _generate_pdf เหมือนเดิม ไม่ต้องแก้)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"{filename_prefix}_{datetime.now().strftime('%Y%m%d')}.pdf",
            title="บันทึกสลิปเงินเดือน"
        )
        if not save_path: return
        
        try:
            self._generate_pdf(data_to_print, save_path)
            messagebox.showinfo("สำเร็จ", f"สร้างไฟล์ PDF เรียบร้อยแล้วที่:\n{save_path}")
            os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("PDF Error", f"เกิดข้อผิดพลาดในการสร้าง PDF:\n{e}")

    def _generate_pdf(self, data_list, filepath):
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

        pay_date = datetime.now().strftime("%d/%m/%Y")
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
    def _load_employees_to_input_tree(self):
        for item in self.input_tree.get_children(): self.input_tree.delete(item)
        self.payroll_inputs = {}
        emps = hr_database.load_all_employees()
        for emp in emps:
            if emp.get('status') not in ['พ้นสภาพพนักงาน', 'ลาออก']:
                self.input_tree.insert("", "end", iid=emp['id'], values=(emp['id'], f"{emp['fname']} {emp['lname']}", "-"))

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
    
    def _calculate_smart_tax(self, monthly_income, monthly_sso):
        """
        คำนวณภาษี ภ.ง.ด. 1 แบบอัตราก้าวหน้า (Progressive Rate)
        อ้างอิงฐานภาษีปี 2567
        """
        # 1. ประมาณการรายได้ทั้งปี
        annual_income = monthly_income * 12
        
        # 2. หักค่าใช้จ่าย (50% ไม่เกิน 100,000)
        expenses = min(annual_income * 0.5, 100000)
        
        # 3. หักค่าลดหย่อน (พื้นฐาน)
        # - ส่วนตัว 60,000
        # - ประกันสังคม (ประมาณการทั้งปี)
        annual_sso = monthly_sso * 12
        allowances = 60000 + annual_sso
        
        # 4. เงินได้สุทธิ (Net Taxable Income)
        net_taxable = annual_income - expenses - allowances
        
        if net_taxable <= 0:
            return 0.0
            
        # 5. คำนวณภาษีตามขั้นบันได
        tax_accumulated = 0.0
        
        # Step 1: 0 - 150,000 (ยกเว้น 0%)
        if net_taxable <= 150000:
            return 0.0
        net_taxable -= 150000
        
        # Step 2: 150,001 - 300,000 (5%) -> Max 7,500
        step_amount = min(net_taxable, 150000)
        tax_accumulated += step_amount * 0.05
        net_taxable -= step_amount
        if net_taxable <= 0: return tax_accumulated / 12
        
        # Step 3: 300,001 - 500,000 (10%) -> Max 20,000
        step_amount = min(net_taxable, 200000)
        tax_accumulated += step_amount * 0.10
        net_taxable -= step_amount
        if net_taxable <= 0: return tax_accumulated / 12
        
        # Step 4: 500,001 - 750,000 (15%) -> Max 37,500
        step_amount = min(net_taxable, 250000)
        tax_accumulated += step_amount * 0.15
        net_taxable -= step_amount
        if net_taxable <= 0: return tax_accumulated / 12
        
        # Step 5: 750,001 - 1,000,000 (20%) -> Max 50,000
        step_amount = min(net_taxable, 250000)
        tax_accumulated += step_amount * 0.20
        net_taxable -= step_amount
        if net_taxable <= 0: return tax_accumulated / 12
        
        # Step 6: 1,000,001 - 2,000,000 (25%) -> Max 250,000
        step_amount = min(net_taxable, 1000000)
        tax_accumulated += step_amount * 0.25
        net_taxable -= step_amount
        if net_taxable <= 0: return tax_accumulated / 12
        
        # Step 7: 2,000,001 ขึ้นไป (คิดสูงสุดที่ 30-35% ตัดจบที่ 30% สำหรับเคสทั่วไป)
        tax_accumulated += net_taxable * 0.30
        
        # หาร 12 เพื่อหัก ณ ที่จ่ายเดือนนี้
        return tax_accumulated / 12
    
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
    
    def _print_pnd1k_pdf(self):
        """ออกรายงาน ภ.ง.ด. 1ก (รายปี) เป็น PDF (รวมใบปะหน้า + ใบแนบ)"""
        
        # 1. ถามปี พ.ศ.
        current_year_be = datetime.now().year + 543
        year_str = simpledialog.askstring("เลือกปีภาษี", f"กรุณากรอกปี พ.ศ. ที่ต้องการออกรายงาน (เช่น {current_year_be}):", initialvalue=str(current_year_be))
        
        if not year_str or not year_str.isdigit(): return
        year_be = int(year_str)
        year_ce = year_be - 543 

        # 2. ดึงข้อมูลทั้งปี
        data_list = hr_database.get_annual_pnd1k_data(year_ce)
        
        if not data_list:
            messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบประวัติการจ่ายเงินเดือนในปี {year_be}")
            return

        # 3. คำนวณยอดรวมก่อน
        total_emp = len(data_list)
        grand_total_income = sum(float(item['annual_income'] or 0) for item in data_list)
        grand_total_tax = sum(float(item['annual_tax'] or 0) for item in data_list)

        # 4. เลือกที่เซฟ
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"PND1K_Year_{year_be}.pdf",
            title=f"บันทึก ภ.ง.ด. 1ก ปี {year_be} (PDF)"
        )
        if not save_path: return

        try:
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=False) # คุมหน้าเอง

            # --- โหลดฟอนต์ ---
            base_path = os.path.dirname(__file__)
            resource_path = os.path.join(base_path, "resources")
            font_path_reg = os.path.join(resource_path, "THSarabunNew.ttf")
            if not os.path.exists(font_path_reg): font_path_reg = os.path.join(base_path, "THSarabunNew.ttf")
            font_path_bold = os.path.join(resource_path, "THSarabunNew Bold.ttf")
            if not os.path.exists(font_path_bold): font_path_bold = os.path.join(base_path, "THSarabunNew Bold.ttf")
            if not os.path.exists(font_path_bold): font_path_bold = font_path_reg

            pdf.add_font("THSarabun", "", font_path_reg, uni=True)
            pdf.add_font("THSarabun", "B", font_path_bold, uni=True)

            def fmt_money(val): return f"{val:,.2f}"

            # ==========================================
            #  ส่วนที่ 1: ใบปะหน้า (Cover Sheet)
            # ==========================================
            pdf.add_page()
            
            # หัวกระดาษ
            pdf.set_font("THSarabun", "B", 22)
            pdf.set_xy(0, 20)
            pdf.cell(0, 10, "ใบสรุป ภ.ง.ด. 1ก (รายปี)", ln=True, align='C')
            
            pdf.set_font("THSarabun", "", 16)
            pdf.cell(0, 10, f"ประจำปีภาษี: {year_be}", ln=True, align='C')
            pdf.ln(10)

            # กล่องยอดรวม
            start_y = pdf.get_y()
            box_w = 160
            center_x = (210 - box_w) / 2

            def draw_cover_row(label, value, is_bold=False):
                x = center_x
                y = pdf.get_y()
                pdf.rect(x, y, box_w, 12)
                
                pdf.set_xy(x + 5, y + 2)
                pdf.set_font("THSarabun", "B" if is_bold else "", 16)
                pdf.cell(100, 8, label, border=0)
                
                pdf.set_xy(x + 105, y + 2)
                pdf.set_font("THSarabun", "B", 16)
                pdf.cell(50, 8, value, border=0, align='R')
                pdf.ln(12)

            draw_cover_row("1. จำนวนรายผู้มีเงินได้ทั้งหมด", f"{total_emp}  ราย")
            draw_cover_row("2. รวมเงินได้ทั้งสิ้นที่จ่ายตลอดปี", f"{grand_total_income:,.2f}  บาท")
            draw_cover_row("3. รวมภาษีที่นำส่งทั้งสิ้น", f"{grand_total_tax:,.2f}  บาท", is_bold=True)

            # ลายเซ็น
            pdf.ln(20)
            pdf.set_font("THSarabun", "", 14)
            pdf.set_x(center_x)
            pdf.cell(0, 8, "ลงชื่อ ....................................................... ผู้มีหน้าที่หักภาษี ณ ที่จ่าย", ln=True, align='C')
            pdf.cell(0, 8, f"( วันที่พิมพ์: {datetime.now().strftime('%d/%m/%Y')} )", ln=True, align='C')

            # ==========================================
            #  ส่วนที่ 2: ใบแนบ (Attachment List)
            # ==========================================
            pdf.add_page() # ขึ้นหน้าใหม่สำหรับรายการ

            # กำหนดคอลัมน์: ลำดับ, บัตร ปชช, ชื่อ-สกุล, วันที่จ่าย, เงินได้ทั้งปี, ภาษีทั้งปี
            col_w = [10, 35, 60, 25, 30, 30]
            headers = ["ลำดับ", "เลขบัตรประชาชน", "ชื่อ-นามสกุล", "วันเดือนปี", "เงินได้ทั้งปี", "ภาษีทั้งปี"]

            def draw_attach_header():
                pdf.set_font("THSarabun", "B", 18)
                pdf.cell(0, 10, f"ใบแนบ ภ.ง.ด. 1ก ประจำปี {year_be}", ln=True, align='C')
                pdf.ln(2)
                
                # หัวตาราง
                pdf.set_fill_color(230, 230, 230)
                pdf.set_font("THSarabun", "B", 14)
                for i, h in enumerate(headers):
                    pdf.cell(col_w[i], 8, h, border=1, align='C', fill=True)
                pdf.ln()

            draw_attach_header()

            # วนลูปข้อมูล
            pdf.set_font("THSarabun", "", 14)
            seq = 1
            current_y = pdf.get_y()
            row_h = 7
            bottom_margin = 270

            for item in data_list:
                if current_y + row_h > bottom_margin:
                    pdf.add_page()
                    draw_attach_header()
                    current_y = pdf.get_y()

                inc = float(item['annual_income'] or 0)
                tax = float(item['annual_tax'] or 0)
                fullname = f"{item.get('fname','')} {item.get('lname','')}"
                id_card = item.get('id_card', '-')

                pdf.cell(col_w[0], row_h, str(seq), 1, 0, 'C')
                
                if len(id_card) > 13: pdf.set_font("THSarabun", "", 12)
                pdf.cell(col_w[1], row_h, id_card, 1, 0, 'C')
                pdf.set_font("THSarabun", "", 14)
                
                pdf.cell(col_w[2], row_h, f"  {fullname}", 1, 0, 'L')
                pdf.cell(col_w[3], row_h, "ตลอดปี", 1, 0, 'C')
                pdf.cell(col_w[4], row_h, fmt_money(inc), 1, 0, 'R')
                pdf.cell(col_w[5], row_h, fmt_money(tax), 1, 0, 'R')
                pdf.ln()
                
                current_y += row_h
                seq += 1

            # บรรทัดยอดรวมท้ายตาราง
            pdf.set_font("THSarabun", "B", 14)
            pdf.set_fill_color(204, 255, 204)
            
            pdf.cell(sum(col_w[:4]), 8, "รวมยอดทั้งสิ้น", 1, 0, 'R', fill=True)
            pdf.cell(col_w[4], 8, fmt_money(grand_total_income), 1, 0, 'R', fill=True)
            pdf.cell(col_w[5], 8, fmt_money(grand_total_tax), 1, 0, 'R', fill=True)

            pdf.output(save_path)
            os.startfile(save_path)

        except Exception as e:
            messagebox.showerror("Error", f"สร้าง PDF ภ.ง.ด.1ก ไม่สำเร็จ:\n{e}")