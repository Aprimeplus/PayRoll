# (ไฟล์ใหม่: custom_widgets.py)

import tkinter as tk
from tkinter import ttk
from datetime import datetime
import calendar

class DateDropdown(ttk.Frame):
    """
    Widget (แบบ Frame) ที่รวม Combobox 3 อัน (Day, Month, Year)
    เพื่อเลียนแบบการทำงานของ tkcalendar.DateEntry
    - .get()       -> คืนค่า string "dd/mm/yyyy" (พ.ศ.)
    - .get_date()  -> คืนค่า datetime.date object (ค.ศ.)
    - .insert(0, text) -> ตั้งค่าจาก string "dd/mm/yyyy" (พ.ศ.)
    - .delete(0, 'end') -> ล้างค่า
    """
    
    THAI_MONTHS = [
        'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน',
        'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม',
        'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ]
    MONTH_TO_INT = {name: i+1 for i, name in enumerate(THAI_MONTHS)}
    
    def __init__(self, parent, width=12, font=("Segoe UI", 10), **kwargs):
        super().__init__(parent, **kwargs)
        
        # --- ตัวแปรเก็บค่า ---
        self.day_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.year_var = tk.StringVar()

        # --- สร้าง Lists ---
        self.day_list = list(range(1, 32))
        self.month_list = self.THAI_MONTHS
        
        # (สร้าง list ปี พ.ศ. 40 ปีย้อนหลัง และ 5 ปีล่วงหน้า)
        current_year_be = datetime.now().year + 543
        self.year_list = list(range(current_year_be + 5, current_year_be - 40, -1))

        # --- สร้าง Widgets ---
        self.day_combo = ttk.Combobox(self, values=self.day_list, 
                                      textvariable=self.day_var, 
                                      width=4, font=font, state="readonly")
        
        self.month_combo = ttk.Combobox(self, values=self.month_list, 
                                        textvariable=self.month_var, 
                                        width=12, font=font, state="readonly")
        
        self.year_combo = ttk.Combobox(self, values=self.year_list, 
                                       textvariable=self.year_var, 
                                       width=6, font=font, state="readonly")

        # --- จัดวาง ---
        self.day_combo.pack(side="left", padx=(0, 2))
        self.month_combo.pack(side="left", padx=2)
        self.year_combo.pack(side="left", padx=(2, 0))
        
        # --- ผูก Events ---
        self.month_var.trace_add("write", self._update_days)
        self.year_var.trace_add("write", self._update_days)

    def _update_days(self, *args):
        """(Internal) อัปเดตจำนวนวันในเดือน ตามเดือนและปีที่เลือก"""
        try:
            month_name = self.month_var.get()
            year_be_str = self.year_var.get()
            
            if not month_name or not year_be_str:
                self.day_combo.config(values=self.day_list)
                return

            month_int = self.MONTH_TO_INT[month_name]
            year_ce = int(year_be_str) - 543
            
            # (คำนวณวันสุดท้ายของเดือน)
            days_in_month = calendar.monthrange(year_ce, month_int)[1]
            new_day_list = list(range(1, days_in_month + 1))
            
            # (อัปเดต Combobox วัน)
            self.day_combo.config(values=new_day_list)

            # (ถ้าวันที่เลือกไว้เดิม เกินจำนวนวันสูงสุด ให้ล้างค่า)
            if self.day_var.get() and int(self.day_var.get()) > days_in_month:
                self.day_var.set("")
                
        except Exception as e:
            print(f"Error updating days: {e}")
            self.day_combo.config(values=self.day_list)

    def get_date(self):
        """คืนค่าเป็น datetime.date object (ค.ศ.)"""
        try:
            day = int(self.day_var.get())
            month_name = self.month_var.get()
            year_be = int(self.year_var.get())
            
            month_int = self.MONTH_TO_INT[month_name]
            year_ce = year_be - 543
            
            return datetime(year_ce, month_int, day).date()
        except Exception:
            # (ถ้าค่าไม่ครบหรือไม่ถูกต้อง)
            return None

    def get(self):
        """คืนค่าเป็น string "dd/mm/yyyy" (พ.ศ.)"""
        try:
            day = int(self.day_var.get())
            month_name = self.month_var.get()
            year_be = int(self.year_var.get())
            
            month_int = self.MONTH_TO_INT[month_name]
            
            return f"{day:02d}/{month_int:02d}/{year_be}"
        except Exception:
            return ""

    def set_date(self, date_obj):
        """ตั้งค่าจาก datetime.date object (ค.ศ.)"""
        if not date_obj:
            self.clear()
            return
            
        try:
            self.day_var.set(str(date_obj.day))
            self.month_var.set(self.THAI_MONTHS[date_obj.month - 1])
            self.year_var.set(str(date_obj.year + 543))
            self._update_days() # (สำคัญ)
        except Exception as e:
            print(f"Error setting date: {e}")
            self.clear()

    def set_date_from_str(self, date_str_be):
        """ตั้งค่าจาก string "dd/mm/yyyy" (พ.ศ.)"""
        if not date_str_be:
            self.clear()
            return

        try:
            day, month, year_be = map(int, date_str_be.split('/'))
            self.day_var.set(str(day))
            self.month_var.set(self.THAI_MONTHS[month - 1])
            self.year_var.set(str(year_be))
            self._update_days() # (สำคัญ)
        except Exception as e:
            print(f"Error setting date from string: {e}")
            self.clear()

    def insert(self, index, text):
        """เลียนแบบ .insert() เพื่อรับค่า string"""
        if index == 0:
            self.set_date_from_str(text)

    def delete(self, start, end):
        """เลียนแบบ .delete(0, 'end')"""
        self.clear()

    def clear(self):
        """ล้างค่าทั้งหมด"""
        self.day_var.set("")
        self.month_var.set("")
        self.year_var.set("")
        self.day_combo.config(values=self.day_list)