# (ไฟล์: transport_module.py)
# (เวอร์ชัน Debug - เช็คว่าทำไมปุ่มไม่ขึ้น)

import tkinter as tk
from tkinter import ttk, messagebox
from custom_widgets import DateDropdown 
import hr_database
from datetime import datetime
from daily_timesheet import DailyTimesheetWindow 

# (!!! สำคัญ !!!) ต้องมีไฟล์ ot_manager.py อยู่ข้างๆ ไฟล์นี้
try:
    from ot_manager import OTManagerWindow 
    print("DEBUG: Import OTManagerWindow สำเร็จ!") # <--- เช็ค 1
except ImportError as e:
    print(f"DEBUG: Error Importing OTManagerWindow: {e}")
    # (ถ้า Import ไม่ได้ โปรแกรมอาจจะทำงานต่อแต่ปุ่มจะพัง)

class TransportModule(ttk.Frame):
    
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user
        print("DEBUG: กำลังเริ่ม TransportModule...") # <--- เช็ค 2

        self.THAI_MONTHS = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
            5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
            9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }
        self.MONTH_TO_INT = {v: k for k, v in self.THAI_MONTHS.items()}

        # --- UI Layout ---
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill="x", pady=(0, 20))
        ttk.Label(header_frame, text="🚛 ระบบจัดการเที่ยวรถ (Transport Management)", 
                  style="Header.TLabel", font=("", 16, "bold")).pack(side="left")
        
        user_info = f"User: {current_user['username']} ({current_user['role']})"
        ttk.Label(header_frame, text=user_info, foreground="gray").pack(side="right")

        # Control Panel
        self._build_control_panel(main_frame)
        
        # List Panel
        self._build_list_panel(main_frame)

    def _build_control_panel(self, parent):
        filter_frame = ttk.LabelFrame(parent, text="  เลือกช่วงเวลา  ", padding=15)
        filter_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_frame, text="ปี (พ.ศ.):").pack(side="left", padx=(5,5))
        current_year_be = datetime.now().year + 543
        year_values = [str(y) for y in range(current_year_be + 1, current_year_be - 5, -1)]
        self.year_combo = ttk.Combobox(filter_frame, values=year_values, width=8, state="readonly")
        self.year_combo.set(str(current_year_be))
        self.year_combo.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="เดือน:").pack(side="left", padx=5)
        self.month_combo = ttk.Combobox(filter_frame, values=list(self.THAI_MONTHS.values()), width=15, state="readonly")
        self.month_combo.set(self.THAI_MONTHS[datetime.now().month])
        self.month_combo.pack(side="left", padx=5)

    def _build_list_panel(self, parent):
        # ปุ่มเครื่องมือ
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(btn_frame, text="🔄 โหลดรายชื่อพนักงาน", command=self._load_employees).pack(side="left")
        
        # ปุ่มเดิม
        ttk.Button(btn_frame, text="📝 บันทึกเที่ยวรถ", 
                   command=self._open_daily_timesheet, 
                   style="Primary.TButton").pack(side="left", padx=10)


        # ตารางรายชื่อ
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("id", "name", "pos"), show="headings", height=20)
        self.tree.heading("id", text="รหัส")
        self.tree.heading("name", text="ชื่อ-นามสกุล")
        self.tree.heading("pos", text="ตำแหน่ง")
        
        self.tree.column("id", width=100, anchor="center")
        self.tree.column("name", width=300)
        self.tree.column("pos", width=200)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.tree.bind("<Double-1>", lambda e: self._open_daily_timesheet())

    def _load_employees(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        
        emps = hr_database.load_all_employees()
        for emp in emps:
            if emp.get('status') not in ['พ้นสภาพพนักงาน', 'ลาออก']:
                self.tree.insert("", "end", values=(
                    emp['id'], 
                    f"{emp['fname']} {emp['lname']}",
                    emp.get('position', '-')
                ))

    def _open_daily_timesheet(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("เตือน", "กรุณาเลือกพนักงานก่อนครับ")
            return
        emp_id = self.tree.item(selection[0], "values")[0]
        try:
            y_be = int(self.year_combo.get())
            y_ce = y_be - 543
            m_name = self.month_combo.get()
            m_int = self.MONTH_TO_INT[m_name]
        except:
            messagebox.showerror("Error", "กรุณาเลือก ปี/เดือน ให้ถูกต้อง")
            return
        DailyTimesheetWindow(self, emp_id, m_int, y_ce)

    def _open_ot_manager(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("เตือน", "กรุณาเลือกพนักงานก่อนครับ")
            return
        emp_id = self.tree.item(selection[0], "values")[0]
        try:
            y_be = int(self.year_combo.get())
            y_ce = y_be - 543
            m_name = self.month_combo.get()
            m_int = self.MONTH_TO_INT[m_name]
        except:
            messagebox.showerror("Error", "กรุณาเลือก ปี/เดือน ให้ถูกต้อง")
            return
        
        print(f"DEBUG: กำลังเปิด OT Manager สำหรับ {emp_id}") # <--- เช็ค 5
        OTManagerWindow(self, emp_id, m_int, y_ce)