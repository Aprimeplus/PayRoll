import psycopg2
from datetime import datetime, timedelta

# ตั้งค่า Connection
DB_CONFIG = {
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "host": "192.168.1.51",
    "port": "5432"
}

def check_probation():
    conn = psycopg2.connect(**DB_CONFIG)
    cursor = conn.cursor()
    
    today = datetime.now().date()
    next_30 = today + timedelta(days=30)
    
    print(f"--- 🕒 ช่วงเวลาที่ตรวจสอบ: {today} ถึง {next_30} ---")
    
    # 1. เช็กคนที่มีสถานะ 'ระหว่างทดลองงาน' ทั้งหมด
    print("\n--- 1. รายชื่อคนที่มีสถานะ 'ระหว่างทดลองงาน' ---")
    cursor.execute("SELECT fname, status, probation_end_date FROM employees WHERE status = 'ระหว่างทดลองงาน'")
    rows = cursor.fetchall()
    if not rows:
        print("❌ ไม่พบพนักงานที่มีสถานะ 'ระหว่างทดลองงาน' เลย")
    else:
        for r in rows:
            fname, status, end_date = r
            print(f"ชื่อ: {fname} | สถานะ: {status} | วันครบกำหนด: {end_date} (Type: {type(end_date)})")
            
            # วิเคราะห์
            if end_date is None:
                print("   -> ⚠️ วันที่ว่างเปล่า (NULL) - เลยไม่แจ้งเตือน")
            elif end_date < today:
                print("   -> ❌ ผ่านไปแล้ว (Expired) - เลยไม่แจ้งเตือน")
            elif end_date > next_30:
                print("   -> ❌ ยังอีกนาน (เกิน 30 วัน) - เลยไม่แจ้งเตือน")
            else:
                print("   -> ✅ เข้าเงื่อนไข! (ควรแสดงผล)")

    conn.close()

if __name__ == "__main__":
    check_probation()