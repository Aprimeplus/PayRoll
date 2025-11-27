import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

class DailyJobDemoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DEMO: บันทึกการทำงานรายวัน (Daily Job Entry)")
        self.geometry("1300x700")
        
        # --- ตั้งค่าข้อมูลสมมติ (Config) ---
        self.daily_wage = 500.0          # ค่าแรงรายวัน
        self.hourly_rate = self.daily_wage / 8
        self.ot_rate = self.hourly_rate * 1.5  # OT 1.5 เท่า
        self.car_rates = {"-": 0, "กระบะ": 50, "เฮี๊ยบ": 100} # เรทค่าเที่ยว
        
        # เก็บตัวแปรของ Widget ทุกช่องเพื่อนำมาคำนวณ
        self.row_vars = [] 
        
        self._build_ui()
        
    def _build_ui(self):
        # 1. Header (ข้อมูลพนักงานสมมติ)
        header_frame = ttk.Frame(self, padding=15)
        header_frame.pack(fill="x")
        
        lbl_style = ("Segoe UI", 12)
        ttk.Label(header_frame, text="ชื่อพนักงาน: นายทดสอบ ระบบดี", font=("Segoe UI", 14, "bold")).pack(side="left", padx=10)
        ttk.Label(header_frame, text=f"ค่าแรงรายวัน: {self.daily_wage:,.0f} บาท", font=lbl_style, foreground="green").pack(side="left", padx=10)
        ttk.Label(header_frame, text=f"ค่า OT/ชม.: {self.ot_rate:,.2f} บาท", font=lbl_style, foreground="orange").pack(side="left", padx=10)
        
        # 2. Table Header
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=10)
        
        # สร้าง Canvas + Scrollbar (เผื่อวันเยอะ)
        canvas = tk.Canvas(table_frame)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        self.scroll_inner = ttk.Frame(canvas)
        
        self.scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # หัวตาราง
        headers = ["วันที่", "สถานะ", "ค่าแรง(บาท)", "|", "OT เริ่ม", "OT ถึง", "รวม OT(ชม.)", "ค่า OT(บาท)", "|", "รถที่ขับ", "จำนวนเที่ยว", "ค่าเที่ยว(บาท)", "|", "รวมรับสุทธิ"]
        widths =  [5, 12, 10, 2, 8, 8, 10, 10, 2, 10, 8, 10, 2, 12]
        
        header_row = ttk.Frame(self.scroll_inner)
        header_row.pack(fill="x", pady=5)
        
        for i, txt in enumerate(headers):
            lbl = ttk.Label(header_row, text=txt, font=("Segoe UI", 10, "bold"), width=widths[i], anchor="center")
            lbl.pack(side="left", padx=2)

        # 3. Generate Rows (สร้างแถววันที่ 1-31 สมมติ)
        for day in range(1, 32): # สมมติเดือนที่มี 31 วัน
            self._create_daily_row(day)

        # 4. Footer Summary
        footer_frame = ttk.Frame(self, padding=15)
        footer_frame.pack(fill="x")
        
        self.lbl_total_wage = ttk.Label(footer_frame, text="รวมค่าแรง: 0.00", font=("Segoe UI", 11))
        self.lbl_total_wage.pack(side="left", padx=15)
        
        self.lbl_total_ot = ttk.Label(footer_frame, text="รวม OT: 0.00", font=("Segoe UI", 11))
        self.lbl_total_ot.pack(side="left", padx=15)
        
        self.lbl_total_drive = ttk.Label(footer_frame, text="รวมค่าเที่ยว: 0.00", font=("Segoe UI", 11))
        self.lbl_total_drive.pack(side="left", padx=15)
        
        self.lbl_grand_total = ttk.Label(footer_frame, text="ยอดรวมทั้งสิ้น: 0.00 บาท", font=("Segoe UI", 14, "bold"), foreground="blue")
        self.lbl_grand_total.pack(side="right", padx=15)
        
        # คำนวณครั้งแรก
        self._recalc_all()

    def _create_daily_row(self, day):
        row_frame = ttk.Frame(self.scroll_inner)
        row_frame.pack(fill="x", pady=1)
        
        # ตัวแปรเก็บค่าในแถวนี้
        vars = {
            "status": tk.StringVar(value="ทำงาน"),
            "wage": tk.DoubleVar(value=self.daily_wage),
            "ot_start": tk.StringVar(),
            "ot_end": tk.StringVar(),
            "ot_hrs": tk.DoubleVar(value=0.0),
            "ot_amt": tk.DoubleVar(value=0.0),
            "car_type": tk.StringVar(value="-"),
            "trips": tk.IntVar(value=0),
            "drive_amt": tk.DoubleVar(value=0.0),
            "net_total": tk.DoubleVar(value=0.0)
        }
        
        # 1. วันที่
        ttk.Label(row_frame, text=f"{day}", width=5, anchor="center").pack(side="left", padx=2)
        
        # 2. สถานะ (ทำงาน/หยุด/ลา)
        cb_status = ttk.Combobox(row_frame, textvariable=vars["status"], values=["ทำงาน", "หยุด", "ลา", "ขาด"], width=10, state="readonly")
        cb_status.pack(side="left", padx=2)
        cb_status.bind("<<ComboboxSelected>>", lambda e: self._on_row_change(vars))
        
        # 3. ค่าแรง (Auto)
        entry_wage = ttk.Entry(row_frame, textvariable=vars["wage"], width=10, state="readonly", justify="right")
        entry_wage.pack(side="left", padx=2)
        
        ttk.Label(row_frame, text="|", width=2).pack(side="left")
        
        # 4. OT Start
        ent_ot_start = ttk.Entry(row_frame, textvariable=vars["ot_start"], width=8, justify="center")
        ent_ot_start.pack(side="left", padx=2)
        ent_ot_start.bind("<FocusOut>", lambda e: self._on_row_change(vars)) # คำนวณเมื่อพิมพ์เสร็จ
        
        # 5. OT End
        ent_ot_end = ttk.Entry(row_frame, textvariable=vars["ot_end"], width=8, justify="center")
        ent_ot_end.pack(side="left", padx=2)
        ent_ot_end.bind("<FocusOut>", lambda e: self._on_row_change(vars))
        
        # 6. OT Hrs (Auto)
        ttk.Label(row_frame, textvariable=vars["ot_hrs"], width=10, anchor="e").pack(side="left", padx=2)
        
        # 7. OT Amt (Auto)
        ttk.Label(row_frame, textvariable=vars["ot_amt"], width=10, anchor="e", foreground="orange").pack(side="left", padx=2)
        
        ttk.Label(row_frame, text="|", width=2).pack(side="left")
        
        # 8. Car Type
        cb_car = ttk.Combobox(row_frame, textvariable=vars["car_type"], values=list(self.car_rates.keys()), width=8, state="readonly")
        cb_car.pack(side="left", padx=2)
        cb_car.bind("<<ComboboxSelected>>", lambda e: self._on_row_change(vars))
        
        # 9. Trips
        ent_trips = ttk.Entry(row_frame, textvariable=vars["trips"], width=8, justify="center")
        ent_trips.pack(side="left", padx=2)
        ent_trips.bind("<KeyRelease>", lambda e: self._on_row_change(vars))
        
        # 10. Drive Amt (Auto)
        ttk.Label(row_frame, textvariable=vars["drive_amt"], width=10, anchor="e", foreground="green").pack(side="left", padx=2)
        
        ttk.Label(row_frame, text="|", width=2).pack(side="left")
        
        # 11. NET TOTAL
        ttk.Label(row_frame, textvariable=vars["net_total"], width=12, anchor="e", font=("", 9, "bold")).pack(side="left", padx=2)

        self.row_vars.append(vars)

    def _on_row_change(self, vars):
        """Logic คำนวณรายบรรทัด"""
        
        # 1. คำนวณค่าแรง
        status = vars["status"].get()
        if status == "ทำงาน":
            vars["wage"].set(self.daily_wage)
        else:
            vars["wage"].set(0.0)
            
        # 2. คำนวณ OT (Time Diff)
        start_str = vars["ot_start"].get()
        end_str = vars["ot_end"].get()
        ot_hours = 0.0
        ot_money = 0.0
        
        if start_str and end_str:
            try:
                t_start = datetime.strptime(start_str, "%H:%M")
                t_end = datetime.strptime(end_str, "%H:%M")
                diff = (t_end - t_start).total_seconds() / 3600
                if diff > 0:
                    ot_hours = round(diff, 2)
            except ValueError:
                pass # ใส่เวลาผิด ไม่คิดเงิน
        
        ot_money = ot_hours * self.ot_rate
        vars["ot_hrs"].set(f"{ot_hours:.2f}")
        vars["ot_amt"].set(round(ot_money, 2))
        
        # 3. คำนวณค่าเที่ยวรถ
        car = vars["car_type"].get()
        try:
            trips = int(vars["trips"].get())
        except:
            trips = 0
        
        rate = self.car_rates.get(car, 0)
        drive_money = rate * trips
        vars["drive_amt"].set(drive_money)
        
        # 4. รวมสุทธิรายวัน
        total = vars["wage"].get() + ot_money + drive_money
        vars["net_total"].set(f"{total:,.2f}")
        
        # 5. อัปเดตยอดรวมด้านล่าง
        self._recalc_all()

    def _recalc_all(self):
        sum_wage = 0.0
        sum_ot = 0.0
        sum_drive = 0.0
        
        for v in self.row_vars:
            sum_wage += v["wage"].get()
            sum_ot += v["ot_amt"].get()
            sum_drive += v["drive_amt"].get()
            
        grand_total = sum_wage + sum_ot + sum_drive
        
        self.lbl_total_wage.config(text=f"รวมค่าแรง: {sum_wage:,.2f}")
        self.lbl_total_ot.config(text=f"รวม OT: {sum_ot:,.2f}")
        self.lbl_total_drive.config(text=f"รวมค่าเที่ยว: {sum_drive:,.2f}")
        self.lbl_grand_total.config(text=f"ยอดรวมทั้งสิ้น: {grand_total:,.2f} บาท")

if __name__ == "__main__":
    app = DailyJobDemoApp()
    app.mainloop()