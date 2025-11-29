import os
import io
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter

# !!! สำคัญ: ต้องมีไฟล์ hr_database.py อยู่ในโฟลเดอร์เดียวกัน
import hr_database 

# ==============================================================================
# 📍 โซนตั้งค่าพิกัด (รวมทั้ง บริษัท และ พนักงาน - แยกอิสระ)
# ==============================================================================

# ---------------------------------------------------------
# 🏢 1. ข้อมูลบริษัท (ผู้จ่ายเงิน)
# ---------------------------------------------------------
PAYER_TAX_ID = "0123456789000"       
PAYER_NAME = "บริษัท เอไพร์ม พลัส จำกัด"
PAYER_ADDR = "123/45 ถ.สุขุมวิท แขวงคลองเตย เขตคลองเตย กทม. 10110"

# 1.1 เลขประจำตัวผู้เสียภาษีบริษัท (13 หลัก)
PAYER_ID_X = 376          # แนวนอน
PAYER_ID_Y = 747          # แนวตั้ง (ปรับจูนแล้วจาก 747 -> 731 เพื่อไม่ให้ชนขอบบน)

# 1.2 ชื่อบริษัท
PAYER_NAME_X = 60         # แนวนอน
PAYER_NAME_Y = 730        # แนวตั้ง (ต่ำกว่าเลขบัตรนิดหน่อย)

# 1.3 ที่อยู่บริษัท
PAYER_ADDR_X = 60         # แนวนอน
PAYER_ADDR_Y = 708        # แนวตั้ง (ต่ำลงมาจากชื่อ)

# ---------------------------------------------------------
# 👤 2. ข้อมูลพนักงาน (ผู้รับเงิน)
# ---------------------------------------------------------
# 2.1 พิกัดเลขบัตรประชาชน (13 หลัก)
ID_X = 377          # แนวนอAPน 
ID_Y = 678          # แนวตั้ง 

# 2.2 พิกัดชื่อ - นามสกุล
NAME_X = 60         # แนวนอน
NAME_Y = 660        # แนวตั้ง 

# 2.3 พิกัดที่อยู่
ADDR_X = 60         # แนวนอน
ADDR_Y = 631        # แนวตั้ง 

# ---------------------------------------------------------
# 📏 3. การเว้นระยะตัวเลข (ใช้ร่วมกันทั้ง บริษัท และ พนักงาน)
# ---------------------------------------------------------
ID_SPACING = 10.9   # ระยะห่างระหว่างตัวเลขปกติ
ID_GROUP_GAP = 10.3 # ระยะกระโดดข้ามขีด

# ---------------------------------------------------------
# ⚙️ 4. ตั้งค่าทั่วไป
# ---------------------------------------------------------
SHOW_GRID = True    # เปิดตารางไม้บรรทัด (True/False)

# ==============================================================================

def draw_id_card_spaced(c, x, y, text, spacing=13, group_gap=8):
    """ฟังก์ชันวาดเลขบัตรแบบกระโดดข้ามขีด"""
    c.setFont('THSarabun', 16)
    text = str(text).replace("-", "").strip()
    curr_x = x
    
    # index ที่ต้องกระโดด: หลังตัวที่ 1, 5, 10, 12 (Index: 0, 4, 9, 11)
    jump_indices = [0, 4, 9, 11]

    for i, char in enumerate(text):
        c.drawString(curr_x, y, char)
        step = spacing
        if i in jump_indices:
            step += group_gap
        curr_x += step

def create_test_pdf_real_data():
    # --- 1. สร้างหน้าต่างถามข้อมูล ---
    root = tk.Tk()
    root.withdraw()

    emp_id_input = simpledialog.askstring("Input", "กรุณากรอกรหัสพนักงาน (เช่น AP082):")
    if not emp_id_input: return

    current_year = datetime.now().year + 543
    year_str = simpledialog.askstring("Input", f"กรุณากรอกปี พ.ศ. (เช่น {current_year}):", initialvalue=str(current_year))
    if not year_str or not year_str.isdigit(): return
    
    year_be = int(year_str)
    year_ce = year_be - 543

    # --- 2. ดึงข้อมูลพนักงานจาก Database ---
    try:
        emp_data = hr_database.get_employee_annual_summary(emp_id_input, year_ce)
    except Exception as e:
        messagebox.showerror("Database Error", f"เชื่อมต่อฐานข้อมูลไม่ได้: {e}")
        return

    if not emp_data:
        messagebox.showerror("Not Found", f"ไม่พบข้อมูลการจ่ายเงินของ {emp_id_input} ในปี {year_be}")
        return

    # --- 3. เริ่มกระบวนการสร้าง PDF ---
    base_dir = os.getcwd()
    template_path = os.path.join(base_dir, "approve_wh3_081156.pdf")
    if not os.path.exists(template_path):
        template_path = os.path.join(base_dir, "resources", "approve_wh3_081156.pdf")
    
    font_path = os.path.join(base_dir, "THSarabunNew.ttf")
    if not os.path.exists(font_path):
        font_path = os.path.join(base_dir, "resources", "THSarabunNew.ttf")

    if not os.path.exists(template_path) or not os.path.exists(font_path):
        messagebox.showerror("Error", "หาไฟล์ Template หรือ Font ไม่เจอ")
        return

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    pdfmetrics.registerFont(TTFont('THSarabun', font_path))

    # --- Option: วาด Grid ไม้บรรทัด ---
    if SHOW_GRID:
        c.setFont('Helvetica', 8)
        c.setStrokeColorRGB(1, 0, 0)
        c.setLineWidth(0.5)
        for x in range(0, 600, 10):
            if x % 50 == 0:
                c.line(x, 0, x, 900)
                c.drawString(x, 5, str(x))
            elif x % 10 == 0:
                c.setStrokeColorRGB(1, 0.8, 0.8)
                c.line(x, 0, x, 900)
                c.setStrokeColorRGB(1, 0, 0)
        for y in range(0, 900, 10):
            if y % 20 == 0:
                c.line(0, y, 600, y)
                c.drawString(5, y, str(y))
            elif y % 10 == 0:
                c.setStrokeColorRGB(1, 0.8, 0.8)
                c.line(0, y, 600, y)
                c.setStrokeColorRGB(1, 0, 0)

    # =========================================================
    # 🔴 วาดข้อมูลบริษัท (ผู้จ่าย) - ใช้พิกัดแยกอิสระ
    # =========================================================
    
    # 1. เลขภาษีบริษัท
    draw_id_card_spaced(c, PAYER_ID_X, PAYER_ID_Y, PAYER_TAX_ID, spacing=ID_SPACING, group_gap=ID_GROUP_GAP)
    
    c.setFont('THSarabun', 14)
    # 2. ชื่อบริษัท (ตำแหน่งอิสระ)
    c.drawString(PAYER_NAME_X, PAYER_NAME_Y, PAYER_NAME) 
    
    # 3. ที่อยู่บริษัท (ตำแหน่งอิสระ)
    c.drawString(PAYER_ADDR_X, PAYER_ADDR_Y, PAYER_ADDR)

    # =========================================================
    # 🔵 วาดข้อมูลพนักงาน (ผู้รับ) - ใช้พิกัดแยกอิสระ
    # =========================================================
    
    # 1. เลขบัตรประชาชน
    card_id = emp_data.get('id_card', '')
    draw_id_card_spaced(c, ID_X, ID_Y, card_id, spacing=ID_SPACING, group_gap=ID_GROUP_GAP)

    # 2. ชื่อ-นามสกุล
    full_name = f"{emp_data.get('fname', '')} {emp_data.get('lname', '')}"
    c.setFont('THSarabun', 14)
    c.drawString(NAME_X, NAME_Y, full_name)

    # 3. ที่อยู่
    address = emp_data.get('address', '-')
    c.drawString(ADDR_X, ADDR_Y, address)

    # 4. ข้อมูลการเงิน
    Y_INC = 538
    c.drawString(330, Y_INC, f"ตลอดปี {year_be}")
    
    total_income = emp_data.get('total_income', 0.0)
    total_tax = emp_data.get('total_tax', 0.0)
    
    c.drawRightString(480, Y_INC, f"{total_income:,.2f}")
    c.drawRightString(550, Y_INC, f"{total_tax:,.2f}")

    # ยอดรวมด้านล่าง
    Y_TOTAL = 248
    c.drawRightString(480, Y_TOTAL, f"{total_income:,.2f}")
    c.drawRightString(550, Y_TOTAL, f"{total_tax:,.2f}")

    c.save()
    packet.seek(0)

    # 3. รวมร่างกับ Template
    new_pdf = PdfReader(packet)
    existing_pdf = PdfReader(open(template_path, "rb"))
    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    # 4. บันทึกและเปิดไฟล์
    output_filename = f"TEST_50Tawi_{emp_id_input}_{year_be}.pdf"
    with open(output_filename, "wb") as f:
        output.write(f)
    
    print(f"✅ สร้างไฟล์สำเร็จ: {output_filename}")
    os.startfile(output_filename)

if __name__ == "__main__":
    create_test_pdf_real_data()