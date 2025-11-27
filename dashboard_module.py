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
        
        # (ขวา: ตารางแจ้งเตือน)
        alert_frame = ttk.LabelFrame(content_frame, text=" 🔔 แจ้งเตือน: ใกล้ผ่านโปร (30 วัน) ", padding=10)
        alert_frame.pack(side="right", fill="both", expand=True)
        
        columns = ("name", "dept", "date")
        self.alert_tree = ttk.Treeview(alert_frame, columns=columns, show="headings", height=15)
        self.alert_tree.heading("name", text="ชื่อ-สกุล")
        self.alert_tree.heading("dept", text="แผนก")
        self.alert_tree.heading("date", text="วันครบกำหนด")
        
        self.alert_tree.column("name", width=150)
        self.alert_tree.column("dept", width=100)
        self.alert_tree.column("date", width=100)
        
        self.alert_tree.pack(fill="both", expand=True)
        
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
        stats = hr_database.get_dashboard_stats()
        
        # 1. อัปเดตการ์ด (เหมือนเดิม)
        self.card_total.config(text=f"{stats['total_employees']} คน")
        self.card_leave.config(text=f"{stats['on_leave_today']} คน")
        self.card_late.config(text=f"{stats['late_today']} คน")
        
        # 2. อัปเดตกราฟ
        self.ax.clear()
        dept_data = stats['dept_counts']
        
        if dept_data:
            labels = [d['dept'] for d in dept_data]
            sizes = [d['count'] for d in dept_data]
            total_employees = stats['total_employees'] 

            def format_label(pct):
                """ฟังก์ชัน Custom Autopct"""
                absolute_count = int(round(pct / 100. * total_employees))
                return f'{pct:.1f}% ({absolute_count} คน)'
            
            colors = ['#ff9999','#66b3ff','#99ff99','#ffcc99', '#c2c2f0', '#ffb3e6']
            
            # (!!! เพิ่มพารามิเตอร์จัดการระยะห่างตรงนี้ !!!)
            self.ax.pie(sizes, 
                        labels=labels, 
                        autopct=format_label, 
                        startangle=90, 
                        colors=colors,
                        labeldistance=1.05,       # ขยับชื่อแผนก (labels) ออกไปเล็กน้อย (แก้ไขปัญหาซ้อนทับ)
                        pctdistance=0.75,         # ขยับ %/จำนวนคน (autopct) เข้ามาใกล้ศูนย์กลาง
                        wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'} # เพิ่มเส้นแบ่ง (ทำให้ดูเป็นระเบียบ)
                        ) 
            
            self.ax.axis('equal') 
        else:
            self.ax.text(0.5, 0.5, "ไม่มีข้อมูล", ha='center')
            
        self.canvas_chart.draw()
        
        # 3. อัปเดตแจ้งเตือน (เหมือนเดิม)
        for item in self.alert_tree.get_children():
            self.alert_tree.delete(item)
            
        for emp in stats['probation_upcoming']:
            # แปลงวันที่เป็น พ.ศ.
            d = emp['probation_end_date']
            date_str = f"{d.day}/{d.month}/{d.year + 543}"
            
            self.alert_tree.insert("", "end", values=(
                f"{emp['fname']} {emp['lname']}",
                emp['department'],
                date_str
            ))