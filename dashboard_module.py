# (ไฟล์ใหม่: dashboard_module.py)

import tkinter as tk
from tkinter import ttk
import hr_database
from datetime import datetime

# Import Matplotlib สำหรับกราฟ
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

# ตั้งค่าฟอนต์ภาษาไทยให้ Matplotlib (สำคัญมาก ไม่งั้นจะเป็นสี่เหลี่ยม)
# ลองใช้ Tahoma หรือ Segoe UI ที่มีใน Windows
plt.rcParams['font.family'] = 'Tahoma' 

class DashboardModule(ttk.Frame):
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user
        
        self._build_ui()
        
    def _build_ui(self):
        # 1. Header
        header_frame = ttk.Frame(self, padding=20)
        header_frame.pack(fill="x")
        
        welcome_text = f"👋 ยินดีต้อนรับ, {self.current_user['username']} ({self.current_user['role']})"
        ttk.Label(header_frame, text="📊 แดชบอร์ดภาพรวม (Dashboard)", 
                  style="Header.TLabel").pack(side="left")
        ttk.Label(header_frame, text=welcome_text, 
                  font=("Segoe UI", 10, "bold"), foreground="gray").pack(side="right")

        # 2. Main Content (Scrollable)
        # (เผื่อจอเล็ก เราจะใส่ Scrollbar)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- ส่วนที่ 1: การ์ดสรุป (KPI Cards) ---
        cards_frame = ttk.Frame(self.scroll_frame, padding=20)
        cards_frame.pack(fill="x")
        
        # สร้างการ์ด 3 ใบ
        self.card_total = self._create_card(cards_frame, "พนักงานทั้งหมด", "0 คน", "#3498db", 0)
        self.card_leave = self._create_card(cards_frame, "ลาวันนี้", "0 คน", "#e67e22", 1)
        self.card_late = self._create_card(cards_frame, "มาสายวันนี้", "0 คน", "#e74c3c", 2)
        
        # --- ส่วนที่ 2: กราฟและแจ้งเตือน (แบ่งซ้าย-ขวา) ---
        content_frame = ttk.Frame(self.scroll_frame, padding=20)
        content_frame.pack(fill="both", expand=True)
        
        # (ซ้าย: กราฟวงกลม)
        chart_frame = ttk.LabelFrame(content_frame, text=" สัดส่วนพนักงานตามแผนก ", padding=10)
        chart_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas_chart = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas_chart.get_tk_widget().pack(fill="both", expand=True)
        
        # (ขวา: Panel รวมแจ้งเตือน) - [แก้ไขส่วนนี้]
        right_panel = ttk.Frame(content_frame)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # 1. ตารางแจ้งเตือนผ่านโปร (ด้านบน)
        alert_frame = ttk.LabelFrame(right_panel, text=" 🔔 ใกล้ผ่านโปร (30 วัน) ", padding=10)
        alert_frame.pack(side="top", fill="both", expand=True, pady=(0, 10))
        
        columns = ("name", "dept", "date")
        # ลดความสูงลงเหลือ 6 บรรทัด เพื่อแบ่งพื้นที่
        self.alert_tree = ttk.Treeview(alert_frame, columns=columns, show="headings", height=6) 
        self.alert_tree.heading("name", text="ชื่อ-สกุล")
        self.alert_tree.heading("dept", text="แผนก")
        self.alert_tree.heading("date", text="วันครบกำหนด")
        self.alert_tree.column("name", width=120)
        self.alert_tree.column("dept", width=80)
        self.alert_tree.column("date", width=80)
        self.alert_tree.pack(fill="both", expand=True)
        
        # 2. [เพิ่มใหม่] ตารางแจ้งเตือนไม่สแกนนิ้ว (ด้านล่าง)
        missing_frame = ttk.LabelFrame(right_panel, text=" ⚠️ ไม่สแกนนิ้ว/ขาดงาน (เดือนนี้) ", padding=10)
        missing_frame.pack(side="bottom", fill="both", expand=True)
        
        m_columns = ("date", "name", "dept")
        self.missing_tree = ttk.Treeview(missing_frame, columns=m_columns, show="headings", height=8)
        self.missing_tree.heading("date", text="วันที่")
        self.missing_tree.heading("name", text="ชื่อ-สกุล")
        self.missing_tree.heading("dept", text="แผนก")
        
        self.missing_tree.column("date", width=80)
        self.missing_tree.column("name", width=120)
        self.missing_tree.column("dept", width=80)
        
        # ใส่ Scrollbar ให้ตารางนี้หน่อยเผื่อข้อมูลเยอะ
        m_scroll = ttk.Scrollbar(missing_frame, orient="vertical", command=self.missing_tree.yview)
        self.missing_tree.configure(yscrollcommand=m_scroll.set)
        self.missing_tree.pack(side="left", fill="both", expand=True)
        m_scroll.pack(side="right", fill="y")
        
        # ปุ่มรีเฟรช
        ttk.Button(self.scroll_frame, text="🔄 รีเฟรชข้อมูล", command=self.refresh_data).pack(pady=10)

        # โหลดข้อมูลครั้งแรก
        self.refresh_data()

    def _create_card(self, parent, title, value, color, col_index):
        """สร้าง Widget การ์ดสรุปตัวเลข"""
        frame = tk.Frame(parent, bg="white", relief="raised", borderwidth=1)
        frame.grid(row=0, column=col_index, sticky="ew", padx=10)
        parent.columnconfigure(col_index, weight=1)
        
        # แถบสีด้านซ้าย
        tk.Frame(frame, bg=color, width=5).pack(side="left", fill="y")
        
        content = tk.Frame(frame, bg="white", padx=20, pady=10)
        content.pack(side="left", fill="both", expand=True)
        
        tk.Label(content, text=title, font=("Segoe UI", 10), bg="white", fg="gray").pack(anchor="w")
        val_label = tk.Label(content, text=value, font=("Segoe UI", 20, "bold"), bg="white", fg="#2c3e50")
        val_label.pack(anchor="w")
        
        return val_label # คืนค่า Label เพื่อเอาไปอัปเดตทีหลัง

    def refresh_data(self):
        """ดึงข้อมูลจาก DB มาอัปเดตหน้าจอ"""
        # ดึงข้อมูลล่าสุดจาก Database
        stats = hr_database.get_dashboard_stats()
        
        # =========================================
        # 1. อัปเดตการ์ดตัวเลข (KPI Cards)
        # =========================================
        self.card_total.config(text=f"{stats.get('total_employees', 0)} คน")
        self.card_leave.config(text=f"{stats.get('on_leave_today', 0)} คน")
        self.card_late.config(text=f"{stats.get('late_today', 0)} คน")
        
        # =========================================
        # 2. อัปเดตกราฟวงกลม (Pie Chart)
        # =========================================
        self.ax.clear() # ล้างกราฟเก่าก่อนวาดใหม่
        
        dept_data = stats.get('dept_counts', [])
        
        if dept_data:
            # เตรียมข้อมูลสำหรับกราฟ
            labels = [d['dept'] for d in dept_data]
            sizes = [d['count'] for d in dept_data]
            total_employees = stats.get('total_employees', 1) # ป้องกันหารด้วย 0

            # ฟังก์ชันจัดรูปแบบป้ายกำกับ (% และ จำนวนคน)
            def format_label(pct):
                absolute_count = int(round(pct / 100. * total_employees))
                return f'{pct:.1f}% ({absolute_count} คน)'
            
            # สีของกราฟ
            colors = ['#ff9999','#66b3ff','#99ff99','#ffcc99', '#c2c2f0', '#ffb3e6']
            
            # วาดกราฟ
            try:
                self.ax.pie(sizes, 
                            labels=labels, 
                            autopct=format_label, 
                            startangle=90, 
                            colors=colors,
                            labeldistance=1.1,       # ระยะห่างชื่อแผนก
                            pctdistance=0.75,        # ระยะห่างตัวเลข %
                            wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
                            ) 
                self.ax.axis('equal') # บังคับให้เป็นวงกลม
            except Exception as e:
                print(f"Graph Error: {e}")
                self.ax.text(0.5, 0.5, "Error วาดกราฟ", ha='center')
        else:
            # กรณีไม่มีข้อมูลพนักงานเลย
            self.ax.text(0.5, 0.5, "ไม่มีข้อมูลพนักงาน", ha='center', fontdict={'size': 12})
            
        self.canvas_chart.draw() # สั่งให้ Canvas วาดภาพใหม่
        
        # =========================================
        # 3. อัปเดตตาราง: ใกล้ผ่านโปร (30 วัน)
        # =========================================
        # ล้างข้อมูลเก่าในตารางทิ้งก่อน
        for item in self.alert_tree.get_children():
            self.alert_tree.delete(item)
            
        # ใส่ข้อมูลใหม่ (แก้ไข: ลบส่วนที่ซ้ำออกแล้ว)
        for emp in stats.get('probation_upcoming', []):
            d = emp.get('probation_end_date')
            if d:
                date_str = f"{d.day}/{d.month}/{d.year + 543}"
            else:
                date_str = "-"
            
            self.alert_tree.insert("", "end", values=(
                f"{emp.get('fname', '')} {emp.get('lname', '')}",
                emp.get('department', '-'),
                date_str
            ))
            
        # =========================================
        # 4. อัปเดตตาราง: ไม่สแกนนิ้ว/ขาดงาน (เดือนนี้)
        # =========================================
        # ล้างข้อมูลเก่าในตารางทิ้งก่อน
        for item in self.missing_tree.get_children():
            self.missing_tree.delete(item)
            
        # ใส่ข้อมูลใหม่ (ใช้ .get() เพื่อความปลอดภัย)
        for miss in stats.get('missing_scans', []):
            d = miss.get('work_date')
            if d:
                date_str = f"{d.day}/{d.month}/{d.year + 543}"
            else:
                date_str = "-"
                
            self.missing_tree.insert("", "end", values=(
                date_str,
                f"{miss.get('fname', '')} {miss.get('lname', '')}",
                miss.get('department', '-')
            ))