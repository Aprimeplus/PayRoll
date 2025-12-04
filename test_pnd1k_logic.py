import os
import io
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4, landscape
from pypdf import PdfReader, PdfWriter

# ==========================================
# ฟังก์ชันช่วย: วาดตารางพิกัด (Grid)
# ==========================================
def draw_debug_grid(c, p_width, p_height):
    c.saveState()
    width = int(p_width)
    height = int(p_height)
    
    # เส้นหลัก
    c.setStrokeColorRGB(1, 0, 0, 0.6)
    c.setFillColorRGB(1, 0, 0, 1)
    c.setFont("Helvetica", 8)
    c.setLineWidth(0.5)
    
    for x in range(0, width + 50, 50):
        c.line(x, 0, x, height)
        c.drawString(x + 2, height - 20, str(x))
        c.drawString(x + 2, 20, str(x))

    for y in range(0, height + 50, 50):
        c.line(0, y, width, y)
        c.drawString(5, y + 2, str(y))
        c.drawString(width - 25, y + 2, str(y))

    # เส้นย่อย
    c.setStrokeColorRGB(1, 0, 0, 0.2)
    c.setLineWidth(0.2)
    
    for x in range(0, width + 50, 10):
        if x % 50 != 0: c.line(x, 0, x, height)
    for y in range(0, height + 50, 10):
        if y % 50 != 0: c.line(0, y, width, y)

    c.restoreState()

# ==========================================
# ฟังก์ชันช่วยใหม่: วาดข้อความแบบกำหนดระยะห่างตัวอักษร (Char Spacing)
# ==========================================
def draw_spaced(c, x, y, text, char_space):
    """วาดข้อความโดยถ่างระยะห่างระหว่างตัวอักษร"""
    c.saveState()
    t = c.beginText(x, y)
    t.setFont("THSarabun", 14) # ใช้ฟอนต์เดียวกับที่ลงทะเบียน
    t.setCharSpace(char_space) # กำหนดความห่าง (Spacing)
    t.textOut(text)
    c.drawText(t)
    c.restoreState()

# ==========================================
# 1. สร้าง Fake Data
# ==========================================
def get_fake_employees():
    raw_data = [
        {"id": "EMP01", "fname": "สมชาย", "lname": "พี่ใหญ่", "id_card": "1100100001111", "income": 800000, "tax": 25000, "start_date": "2010-01-01", "address": "123 ถ.สุขุมวิท เขตคลองเตย กทม. 10110"},
        {"id": "EMP02", "fname": "สมศักดิ์", "lname": "รุ่นสอง", "id_card": "1100200002222", "income": 600000, "tax": 15000, "start_date": "2012-06-15", "address": "45/2 หมู่ 5 ต.บางเขน อ.เมือง นนทบุรี 11000"},
        {"id": "EMP03", "fname": "สมศรี", "lname": "รุ่นสาม", "id_card": "1100300003333", "income": 500000, "tax": 12000, "start_date": "2015-03-01", "address": "789 ถ.เพชรเกษม เขตหนองแขม กทม. 10160"},
        {"id": "EMP04", "fname": "สมฤดี", "lname": "รุ่นสี่", "id_card": "1100400004444", "income": 450000, "tax": 9000, "start_date": "2016-08-01", "address": "99 ซ.อารีย์ เขตพญาไท กทม. 10400"},
        {"id": "EMP05", "fname": "สมปอง", "lname": "น้องกลาง", "id_card": "1100500005555", "income": 400000, "tax": 8000, "start_date": "2018-05-01", "address": "88 หมู่บ้านสุขสันต์ บางแค กทม. 10160"},
        {"id": "EMP06", "fname": "สมบูรณ์", "lname": "กลางค่อนใหม่", "id_card": "1100600006666", "income": 380000, "tax": 6000, "start_date": "2019-11-01", "address": "101/1 ถ.ลาดพร้าว เขตจตุจักร กทม. 10900"},
        {"id": "EMP07", "fname": "สมคิด", "lname": "กลางๆ", "id_card": "1100700007777", "income": 350000, "tax": 5000, "start_date": "2021-04-01", "address": "555 ถ.วิภาวดี เขตหลักสี่ กทม. 10210"},
        {"id": "EMP08", "fname": "สมรักษ์", "lname": "กลางๆ", "id_card": "1100800008888", "income": 300000, "tax": 3000, "start_date": "2022-02-01", "address": "222 คอนโดหรู ริมแม่น้ำเจ้าพระยา กทม. 10600"},
        {"id": "EMP09", "fname": "สมหญิง", "lname": "เกือบใหม่", "id_card": "1100900009999", "income": 250000, "tax": 1000, "start_date": "2023-01-01", "address": "333 ถ.รามคำแหง เขตบางกะปิ กทม. 10240"},
        {"id": "EMP10", "fname": "น้องใหม่", "lname": "ล่าสุด", "id_card": "1101000000000", "income": 180000, "tax": 0, "start_date": "2024-01-01", "address": "444 หมู่ 8 ต.สำโรงเหนือ จ.สมุทรปราการ 10270"},
    ]
    return raw_data

def generate_test_pdf():
    print("--- เริ่มต้นการทดสอบ (เพิ่ม X_SEQ กลับเข้ามาแล้ว) ---")
    
    data = get_fake_employees()
    data.sort(key=lambda x: x['start_date'])
    
    TEMPLATE_FILE = "pnd1k.pdf"
    FONT_FILE = "THSarabunNew.ttf" 
    OUTPUT_FILE = "TEST_PND1K_SPACED.pdf"
    
    if not os.path.exists(TEMPLATE_FILE) or not os.path.exists(FONT_FILE):
        print("❌ ไม่พบไฟล์ PDF หรือ Font")
        return

    # --- โซนตั้งค่าพิกัด ---
    Y_START = 413
    ROW_HEIGHT = 39  
    MAX_ROW_PER_PAGE = 7 
    
    # >>> ใส่ตัวแปรที่หายไปกลับคืนมา <<<
    X_SEQ = 80
    # >>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    # >>> [โซนปรับระยะห่างตัวเลข] <<<
    # กลุ่ม 1: เลขตัวเดียว (ไม่ต้องถ่าง)
    X_ID_1 = 111         
    
    # กลุ่ม 2: 4 ตัว (1234) -> ถ่างออก
    X_ID_2 = 127         # จุดเริ่ม
    ID_SPACE_2 = 6.9     # *** ลองปรับเลขนี้ดูครับ (ยิ่งมาก ยิ่งห่าง) ***
    
    # กลุ่ม 3: 5 ตัว (12345) -> ถ่างออก
    X_ID_3 = 180         
    ID_SPACE_3 = 6.9    # *** ลองปรับเลขนี้ดูครับ ***
    
    # กลุ่ม 4: 2 ตัว (12) -> ถ่างออก
    X_ID_4 = 247         
    ID_SPACE_4 = 4.7     # *** ลองปรับเลขนี้ดูครับ ***
    
    # กลุ่ม 5: เลขตัวเดียว (ไม่ต้องถ่าง)
    X_ID_5 = 278        
    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    X_FNAME = 302  
    X_LNAME = 452  
    
    X_ADDRESS = 315    
    Y_ADDR_OFFSET = 15 
    
    X_INCOME = 680; X_TAX = 780; X_COND = 790
    Y_TOTAL_ROW = 135
    
    Y_PAGE = 475          
    X_PAGE_CURR = 701 
    X_PAGE_TOTAL = 765  
    # --------------------------------

    try:
        pdfmetrics.registerFont(TTFont('THSarabun', FONT_FILE))
        
        template_reader_init = PdfReader(open(TEMPLATE_FILE, "rb"))
        first_page = template_reader_init.pages[0]
        page_width = float(first_page.mediabox.width)
        page_height = float(first_page.mediabox.height)
        
        output_writer = PdfWriter()
        grand_total_income = sum(d['income'] for d in data)
        grand_total_tax = sum(d['tax'] for d in data)
        
        chunks = [data[i:i + MAX_ROW_PER_PAGE] for i in range(0, len(data), MAX_ROW_PER_PAGE)]
        total_pages = len(chunks)
        seq_global = 1 

        for page_idx, batch in enumerate(chunks):
            current_page_num = page_idx + 1
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=(page_width, page_height))
            
            #draw_debug_grid(c, page_width, page_height)

            c.saveState() 
            c.setFillColorRGB(0, 0 , 0)
            c.setFont("THSarabun", 18)   
            c.drawString(X_PAGE_CURR, Y_PAGE, str(current_page_num))
            c.drawString(X_PAGE_TOTAL, Y_PAGE, str(total_pages))
            c.restoreState() 

            c.setFont("THSarabun", 14)
            current_y = Y_START
            
            for person in batch:
                c.drawCentredString(X_SEQ, current_y, str(seq_global))
                
                # >>> ส่วนที่เปลี่ยน: ใช้ draw_spaced แทน drawString <<<
                tid = person['id_card']
                c.drawString(X_ID_1, current_y, tid[0])           # 1 ตัว
                
                draw_spaced(c, X_ID_2, current_y, tid[1:5], ID_SPACE_2)   # 4 ตัว (ถ่าง)
                draw_spaced(c, X_ID_3, current_y, tid[5:10], ID_SPACE_3)  # 5 ตัว (ถ่าง)
                draw_spaced(c, X_ID_4, current_y, tid[10:12], ID_SPACE_4) # 2 ตัว (ถ่าง)
                
                c.drawString(X_ID_5, current_y, tid[12])          # 1 ตัว
                # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                
                c.drawString(X_FNAME, current_y, person['fname'])
                c.drawString(X_LNAME, current_y, person['lname'])
                
                c.setFont("THSarabun", 10)
                c.drawString(X_ADDRESS, current_y - Y_ADDR_OFFSET, person['address']) 
                c.setFont("THSarabun", 14)

                c.drawRightString(X_INCOME, current_y, f"{person['income']:,.2f}")
                c.drawRightString(X_TAX, current_y, f"{person['tax']:,.2f}")
                c.drawCentredString(X_COND, current_y, "1")
                
                current_y -= ROW_HEIGHT
                seq_global += 1
            
            if current_page_num == total_pages:
                c.drawRightString(X_INCOME, Y_TOTAL_ROW, f"{grand_total_income:,.2f}")
                c.drawRightString(X_TAX, Y_TOTAL_ROW, f"{grand_total_tax:,.2f}")

            c.save()
            packet.seek(0)
            
            template_reader = PdfReader(open(TEMPLATE_FILE, "rb"))
            bg_page = template_reader.pages[0] 
            overlay_reader = PdfReader(packet)
            bg_page.merge_page(overlay_reader.pages[0])
            output_writer.add_page(bg_page)

        with open(OUTPUT_FILE, "wb") as f:
            output_writer.write(f)
            
        print(f"✅ สร้างไฟล์ทดสอบเรียบร้อย: {OUTPUT_FILE}")
        os.startfile(OUTPUT_FILE)

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    generate_test_pdf()