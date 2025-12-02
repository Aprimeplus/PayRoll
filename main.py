# (ไฟล์: main.py)
# (เวอร์ชันอัปเดต - ***บังคับ Theme "clam"***)

import tkinter as tk
from tkinter import ttk, messagebox 
import sys
from transport_module import TransportModule
from employee_module import EmployeeModule
import hr_database
from login_screen import LoginScreen 
from attendance_module import AttendanceModule
from time_processor_module import TimeProcessorModule
from payroll_module import PayrollModule 
from company_profile_module import CompanyProfileModule 
from user_management_module import UserManagementModule
from dashboard_module import DashboardModule

class MainApp(tk.Tk):
    def __init__(self, current_user):
        super().__init__()
        self.title("Company Management System")
        self.geometry("1400x850")
        self.current_user = current_user

        # (!!!) เรียกใช้การตั้งค่าสไตล์
        self._configure_global_styles() 

        self._create_main_layout()
        self.frames = {}
        self._create_module_frames()

        # --- (แก้ไขตรงนี้) เลือกหน้าแรกตาม Role ---
        if self.current_user['role'] == 'dispatcher':
            self.show_frame("TransportModule")
        else:
            self.show_frame("DashboardModule")
        # ---------------------------------------

        # ส่วนแสดงชื่อ User (เหมือนเดิม)
        user_label = ttk.Label(self.sidebar_frame,
                               text=f"User: {self.current_user['username']} ({self.current_user['role']})",
                               background="#f0f0f0", anchor="w", padding=(10, 0))
        user_label.pack(side="bottom", fill="x", pady=(0, 10))

    def _configure_global_styles(self):
        """(ใหม่) ตั้งค่าสไตล์กลางสำหรับ Theme 'clam'"""
        style = ttk.Style(self)
        
        # 1. ตั้งค่าสี
        primary_color = "#007bff"
        success_color = "#28a745"
        light_bg = "#ffffff"
        striped_bg = "#f0f0f0"
        dark_text = "#000000" # <--- (สีดำ ตามที่คุณแนะนำ)
        
        # 2. ตั้งค่า TButton (ปุ่มธรรมดา)
        style.configure("TButton", font=("Segoe UI", 10), padding=5)
        
        # 3. สร้าง Style "Success.TButton" (สำหรับปุ่มสีเขียว)
        darker_green = '#218838'
        
        # (!!! แก้ไข: เปลี่ยน 'foreground' เป็นสีดำ และ 'background' เป็นสีขาว !!!)
        style.configure("Success.TButton", 
                        font=("Segoe UI", 10, "bold"), 
                        background=light_bg,       # <--- ตั้ง background เป็นสีขาว
                        foreground=dark_text,      # <--- ตั้ง foreground เป็นสีดำ
                        bordercolor=success_color, # <--- "ขอบ" เป็นสีเขียว
                        darkcolor=light_bg,
                        lightcolor=light_bg)
                        
        style.map("Success.TButton",
                  background=[
                      ('!active', light_bg),  # <--- พื้นหลังขาว
                      ('active', striped_bg)  # <--- พื้นหลังเทา (ตอนกด)
                  ],
                  foreground=[
                      ('!active', dark_text), # <--- ตัวอักษรดำ
                      ('active', dark_text)
                  ],
                  bordercolor=[
                      ('!active', success_color), # <--- ขอบเขียว
                      ('active', darker_green)    # <--- ขอบเขียวเข้ม
                  ])

        # 4. สร้าง Style "Primary.TButton" (สำหรับปุ่มสีน้ำเงิน)
        darker_blue = '#0069d9'
        
        # (!!! แก้ไข: เปลี่ยน 'foreground' เป็นสีดำ และ 'background' เป็นสีขาว !!!)
        style.configure("Primary.TButton",
                        font=("Segoe UI", 10),
                        background=light_bg,       # <--- ตั้ง background เป็นสีขาว
                        foreground=dark_text,      # <--- ตั้ง foreground เป็นสีดำ
                        bordercolor=primary_color, # <--- "ขอบ" เป็นสีน้ำเงิน
                        darkcolor=light_bg,
                        lightcolor=light_bg)
                        
        style.map("Primary.TButton",
                  background=[
                      ('!active', light_bg),  # <--- พื้นหลังขาว
                      ('active', striped_bg)  # <--- พื้นหลังเทา (ตอนกด)
                  ],
                  foreground=[
                      ('!active', dark_text), # <--- ตัวอักษรดำ
                      ('active', dark_text)
                  ],
                  bordercolor=[
                      ('!active', primary_color), # <--- ขอบน้ำเงิน
                      ('active', darker_blue)     # <--- ขอบน้ำเงินเข้ม
                  ])

        # 5. (!!!) แก้ปัญหา Treeview ลายทาง (Striped) ใน Theme 'clam' (!!!)
        # (โค้ดส่วนนี้เหมือนเดิม)
        style.configure("Treeview", 
                        background=light_bg, 
                        fieldbackground=light_bg,
                        font=("Segoe UI", 10),
                        rowheight=25) 
        
        style.configure("Treeview.Heading", 
                        font=("Segoe UI", 10, "bold"), 
                        padding=5)

        # 6. ตั้งค่า Label
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground="#2c3e50")
        style.configure("Section.TLabel", font=("Segoe UI", 11, "bold"), foreground="#34495e")

        # 7. ตั้งค่า Sidebar
        style.configure("Sidebar.TFrame", background=striped_bg)

    def _create_main_layout(self):
        self.sidebar_frame = ttk.Frame(self, width=200, style="Sidebar.TFrame")
        self.sidebar_frame.pack(side="left", fill="y")
        self.sidebar_frame.pack_propagate(False)

        ttk.Separator(self, orient="vertical").pack(side="left", fill="y")

        self.content_area = ttk.Frame(self, padding=10)
        self.content_area.pack(side="right", fill="both", expand=True)

        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)
        
        # ใส่แค่หัวข้อ "เมนูหลัก" พอครับ
        ttk.Label(self.sidebar_frame, text="เมนูหลัก", font=("-size 14 -weight bold"), 
                  background="#f0f0f0", padding=10).pack(pady=10)

    def _create_sidebar_btn(self, text, frame_name):
        """ฟังก์ชันช่วยสร้างปุ่มเมนู Sidebar"""
        btn = ttk.Button(self.sidebar_frame, text=text, command=lambda: self.show_frame(frame_name))
        btn.pack(fill="x", padx=10, pady=5)    
    
    def call_employee_module_method(self, method_name):
        if "EmployeeModule" in self.frames:
            employee_frame = self.frames["EmployeeModule"]
            if hasattr(employee_frame, method_name):
                method = getattr(employee_frame, method_name)
                return method
        return lambda: messagebox.showerror("Error", "Employee module not loaded correctly.")

    def _create_module_frames(self):
        """สร้างหน้าจอต่างๆ ตามสิทธิ์การใช้งาน (Role)"""
        container = self.content_area
        user_role = self.current_user['role']

        # --- CASE 1: Admin/HR/Approver ---
        if user_role in ['hr', 'approver']:
            # ... (ส่วนโหลด Frame เดิม ไม่ต้องแก้) ...
            self.frames["DashboardModule"] = DashboardModule(container, self, self.current_user)
            self.frames["EmployeeModule"] = EmployeeModule(container, self, self.current_user)
            self.frames["AttendanceModule"] = AttendanceModule(container, self, self.current_user)
            self.frames["TimeProcessorModule"] = TimeProcessorModule(container, self, self.current_user)
            self.frames["PayrollModule"] = PayrollModule(container, self, self.current_user)
            self.frames["CompanyProfileModule"] = CompanyProfileModule(container, self, self.current_user)

            for frame in self.frames.values():
                frame.grid(row=0, column=0, sticky="nsew")

            # โหลด UserManagement เฉพาะ Approver
            if user_role == 'approver':
                self.frames["UserManagementModule"] = UserManagementModule(container, self, self.current_user)
                self.frames["UserManagementModule"].grid(row=0, column=0, sticky="nsew")

            # --- สร้างปุ่มเมนูทั่วไป (HR/Approver ใช้เหมือนกัน) ---
            self._create_sidebar_btn("📊 แดชบอร์ด", "DashboardModule")
            self._create_sidebar_btn("👥 ข้อมูลพนักงาน", "EmployeeModule")
            self._create_sidebar_btn("🕒 บันทึกเวลา/การลา", "AttendanceModule")
            self._create_sidebar_btn("⚙️ ประมวลผลเวลา", "TimeProcessorModule")
            self._create_sidebar_btn("💰 คำนวณเงินเดือน", "PayrollModule")
            self._create_sidebar_btn("🏢 ข้อมูลบริษัท", "CompanyProfileModule")
            
            # --- (!!! ส่วนที่คุณต้องการ !!!) ---
            # ถ้าเป็น Approver ให้เพิ่มเส้นคั่นและปุ่มพิเศษ
            if user_role == 'approver':
                # 1. เส้นคั่น
                ttk.Label(self.sidebar_frame, text="--------", background="#f0f0f0").pack(pady=5)
                
                # 2. ปุ่มจัดการผู้ใช้
                self._create_sidebar_btn("🔐 จัดการผู้ใช้", "UserManagementModule")
                
                # 3. ปุ่มอนุมัติการแก้ไข (เรียกฟังก์ชันข้าม Module)
                ttk.Button(self.sidebar_frame, text="✅ อนุมัติการแก้ไข",
                           command=lambda: self.frames["EmployeeModule"].show_approval_page()
                          ).pack(fill="x", padx=10, pady=5)

        # --- CASE 2: Dispatcher ---
        elif user_role == 'dispatcher':
            self.frames["TransportModule"] = TransportModule(container, self, self.current_user)
            self.frames["TransportModule"].grid(row=0, column=0, sticky="nsew")
            self._create_sidebar_btn("🚛 จัดการเที่ยวรถ", "TransportModule")
            self.show_frame("TransportModule") # บังคับเปิดหน้าแรก

        # ปุ่ม Logout
        logout_btn = ttk.Button(self.sidebar_frame, text="🚪 ออกจากระบบ", command=self.logout)
        logout_btn.pack(side="bottom", fill="x", padx=10, pady=20)

    def show_frame(self, frame_name):
        frame = self.frames[frame_name]
        frame.tkraise()  

        if frame_name == "DashboardModule" and hasattr(frame, 'refresh_data'):
             frame.refresh_data()
        
        if frame_name == "EmployeeModule" and hasattr(frame, '_show_list_page'):
            frame._show_list_page() 
            
        if frame_name == "AttendanceModule" and hasattr(frame, '_load_employee_dropdown'):
             frame._load_employee_dropdown()
             
        if frame_name == "CompanyProfileModule" and hasattr(frame, '_load_company_info_data'):
             frame._load_company_info_data()

    def logout(self):
        self.destroy()
        start_application()
    
    def call_employee_module_method(self, method_name):
        if "EmployeeModule" in self.frames:
            employee_frame = self.frames["EmployeeModule"]
            if hasattr(employee_frame, method_name):
                method = getattr(employee_frame, method_name)
                return method
        return lambda: messagebox.showerror("Error", "Employee module not loaded correctly.")

def start_application():
    hr_database.init_db()
    print("DEBUG: 1. start_application() กำลังทำงาน") 
    root = tk.Tk()

    # --- !! (นี่คือ 2 บรรทัดที่สำคัญที่สุด) !! ---
    # (สั่งให้แอปใช้ Theme "clam" ซึ่งรองรับ gridlines)
    style = ttk.Style(root)
    style.theme_use("clam") 
    # --- (จบส่วนที่เพิ่ม) ---
    
    root.withdraw()
    
    print("DEBUG: 2. กำลังจะเปิดหน้า Login...") 
    login_dialog = LoginScreen(root)
    user_info = login_dialog.user_info
    print(f"DEBUG: 3. ปิดหน้า Login แล้ว, user_info คือ: {user_info}") 

    if user_info:
        print("DEBUG: 4. Login สำเร็จ, กำลังเปิด MainApp...") 
        root.destroy()
        app = MainApp(user_info)
        app.mainloop()
        print("DEBUG: 5. ปิด MainApp แล้ว") 
    else:
        print("DEBUG: 4. Login ไม่ผ่าน หรือกดยกเลิก, กำลังปิดโปรแกรม") 
        root.destroy()
        sys.exit()

if __name__ == "__main__":
    start_application()