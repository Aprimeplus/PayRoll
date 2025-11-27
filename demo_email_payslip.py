import smtplib
import ssl
from email.message import EmailMessage
import os
from fpdf import FPDF
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

def create_dummy_pdf(filename):
    """สร้างไฟล์ PDF จำลอง (รองรับภาษาไทย)"""
    pdf = FPDF()
    pdf.add_page()
    
    # --- 1. พยายามโหลดฟอนต์ไทย ---
    # หาไฟล์ฟอนต์ในโฟลเดอร์ปัจจุบัน หรือ โฟลเดอร์ resources
    font_path = "THSarabunNew.ttf"
    if not os.path.exists(font_path):
        font_path = os.path.join("resources", "THSarabunNew.ttf")
    
    if os.path.exists(font_path):
        # ถ้าเจอฟอนต์ -> ใช้ฟอนต์ไทย (พิมพ์ไทยได้)
        pdf.add_font("THSarabun", "", font_path, uni=True)
        pdf.set_font("THSarabun", size=20)
        header_text = "TEST PAYSLIP / ใบแจ้งเงินเดือนทดสอบ"
    else:
        # ถ้าไม่เจอ -> ใช้ Arial (พิมพ์ไทยไม่ได้ ต้องใช้ภาษาอังกฤษ)
        pdf.set_font("Arial", size=12)
        header_text = "TEST PAYSLIP (Font not found)"
        print("⚠️ ไม่พบไฟล์ฟอนต์ THSarabunNew.ttf เลยใช้ Arial แทน")

    # --- 2. สร้างเนื้อหา ---
    pdf.cell(0, 10, txt=header_text, ln=1, align="C")
    
    # (วันที่และเนื้อหาภาษาอังกฤษ ใช้ฟอนต์เดิมต่อได้)
    pdf.cell(0, 10, txt=f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=1, align="C")
    pdf.cell(0, 10, txt="This is a test document sent from Python.", ln=1, align="C")
    
    pdf.output(filename)
    print(f"✅ สร้างไฟล์ {filename} สำเร็จ")

def send_email_demo():
    # --- 1. รับข้อมูลจากผู้ใช้ (ผ่าน Popup) ---
    root = tk.Tk()
    root.withdraw() # ซ่อนหน้าต่างหลัก

    sender_email = simpledialog.askstring("ข้อมูลผู้ส่ง", "กรุณากรอก Gmail ผู้ส่ง:\n(เช่น hr@gmail.com)")
    if not sender_email: return

    # !!! ตรงนี้ต้องใช้ App Password (ไม่ใช่รหัสผ่านเข้าเมลปกติ) !!!
    password = simpledialog.askstring("รหัสผ่าน", "กรุณากรอก App Password (16 หลัก):\n(ไม่ใช่รหัสผ่านปกติ)", show='*')
    if not password: return

    receiver_email = simpledialog.askstring("ผู้รับ", "ส่งไปที่อีเมลไหน?:\n(เช่น your_email@gmail.com)")
    if not receiver_email: return

    # --- 2. สร้างไฟล์ PDF จำลอง ---
    pdf_filename = "test_payslip.pdf"
    try:
        create_dummy_pdf(pdf_filename)
    except Exception as e:
        messagebox.showerror("PDF Error", f"สร้าง PDF ไม่ได้: {e}")
        return

    # --- 3. เตรียมเนื้อหาอีเมล ---
    subject = "ทดสอบส่งสลิปเงินเดือน (Test Payslip)"
    body = """
    เรียน พนักงาน,

    นี่คืออีเมลทดสอบจากระบบ HR System
    กรุณาตรวจสอบไฟล์สลิปเงินเดือนที่แนบมา

    ขอบคุณครับ
    (ระบบอัตโนมัติ)
    """

    em = EmailMessage()
    em['From'] = sender_email
    em['To'] = receiver_email
    em['Subject'] = subject
    em.set_content(body)

    # --- 4. แนบไฟล์ PDF ---
    try:
        with open(pdf_filename, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(pdf_filename)
        
        em.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)
    except Exception as e:
        messagebox.showerror("File Error", f"อ่านไฟล์ PDF ไม่ได้: {e}")
        return

    # --- 5. ส่งออกผ่าน Gmail SMTP ---
    context = ssl.create_default_context()
    
    try:
        print("⏳ กำลังเชื่อมต่อ Server และส่งอีเมล...")
        # ใช้ Gmail SMTP (smtp.gmail.com, port 465)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender_email, password)
            smtp.send_message(em)
        
        print("🎉 ส่งอีเมลสำเร็จ!")
        messagebox.showinfo("สำเร็จ", f"ส่งอีเมลไปที่ {receiver_email} เรียบร้อยแล้ว!\nกรุณาเช็ก Inbox (หรือ Junk Mail)")
        
    except Exception as e:
        print(f"❌ ส่งไม่สำเร็จ: {e}")
        messagebox.showerror("ล้มเหลว", f"เกิดข้อผิดพลาด:\n{e}\n\n(เช็ก App Password หรือการตั้งค่า Gmail)")

    # ลบไฟล์ทดสอบทิ้ง
    if os.path.exists(pdf_filename):
        os.remove(pdf_filename)

if __name__ == "__main__":
    send_email_demo()