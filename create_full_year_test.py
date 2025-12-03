import psycopg2
from datetime import datetime, timedelta, time, date
import calendar
import random

# --- ตั้งค่าการเชื่อมต่อ Database ---
DB_CONFIG = {
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "host": "192.168.1.51",
    "port": "5432"
}

EMP_ID = "TEST_YEAR_001"
YEAR = 2024  # ปีที่สร้างข้อมูล (เวลาทดสอบอย่าลืมเลือกปีนี้นะครับ)

def get_db_connection():
    try:
        return psycopg2.connect(**DB_CONFIG)
    except Exception as e:
        print(f"Error: {e}")
        return None

def create_full_year_data():
    conn = get_db_connection()
    if not conn: return
    cursor = conn.cursor()

    print(f"\n--- 🔄 กำลังสร้างข้อมูล 'พนักงานรายวัน' แบบ Full Option ({YEAR}) ---")

    # 1. ล้างข้อมูลเก่าทิ้งให้หมด (Clean Up)
    print("1. ล้างข้อมูลเก่า...")
    tables = ["employees", "time_attendance_logs", "employee_daily_records", 
              "employee_leave_records", "employee_late_records", 
              "employee_warning_records", "employee_ot_details",
              "payroll_records", "employee_welfare", "salary_history", 
              "employee_company_assets"]
    
    # ลบเฉพาะพนักงานคนนี้
    for t in tables:
        cursor.execute(f"DELETE FROM {t} WHERE emp_id = %s", (EMP_ID,))
    
    # 2. สร้างข้อมูลพนักงานแบบ 'จัดเต็ม' (Full Profile)
    print("2. สร้างประวัติพนักงาน (สมชาย รายวัน)...")
    
    sql_employee = """
        INSERT INTO employees (
            emp_id, fname, lname, nickname, emp_type, 
            position, department, start_date, salary, 
            birth_date, age, id_card, phone, 
            address, current_address, 
            bank_name, bank_account_no, bank_account_name, bank_branch, bank_account_type,
            sso_hospital, sso_start_date, 
            leave_annual_days, leave_sick_days, leave_personal_days,
            work_location, diligence_streak, status,
            emergency_contact_name, emergency_contact_phone, emergency_contact_relation
        ) VALUES (
            %s, 'สมชาย', 'ใจสู้', 'ชาย', 'พนักงานรายวัน', 
            'พนักงานคลังสินค้า', 'Warehouse', '2023-01-01', 500, 
            '1990-05-20', '34', '1-1002-00345-67-8', '081-234-5678', 
            '123/45 หมู่ 8 ต.บางพลี อ.บางพลี จ.สมุทรปราการ', 'ที่อยู่เดียวกับทะเบียนบ้าน', 
            'ธนาคารกสิกรไทย', '123-4-56789-0', 'นายสมชาย ใจสู้', 'สาขาเมกาบางนา', 'ออมทรัพย์',
            'รพ.จุฬารัตน์ 3', '2023-04-01', 
            6, 30, 6, 
            'คลังสินค้า', 0, 'ปกติ',
            'นางสมหญิง ใจสู้', '089-987-6543', 'ภรรยา'
        );
    """
    cursor.execute(sql_employee, (EMP_ID,))

    # 3. ใส่ข้อมูลเสริม (สวัสดิการ, ประวัติเงินเดือน, ทรัพย์สิน)
    print("3. เพิ่มข้อมูลเสริม (สวัสดิการ/ประวัติเงินเดือน)...")
    
    # - ประวัติการปรับเงินเดือน (ปีที่แล้ว 450 -> ปีนี้ 500)
    cursor.execute("""
        INSERT INTO salary_history (emp_id, adjustment_year, new_salary, position_allowance, new_position, assessment_score)
        VALUES (%s, '2023', 450, 0, 'พนักงานคลังสินค้า', 'A')
    """, (EMP_ID,))

    # - สวัสดิการ (เบี้ยขยัน, ค่าเดินทาง)
    welfares = [
        ('เบี้ยขยัน', 1, 0),     # มีสิทธิ์ (แต่ยอดคำนวณอัตโนมัติ)
        ('ค่าเดินทาง', 1, 1000), # ได้ fix 1,000 บาท
        ('ค่าโทรศัพท์', 0, 0)    # ไม่มีสิทธิ์
    ]
    for name, has, amt in welfares:
        cursor.execute("INSERT INTO employee_welfare (emp_id, welfare_name, has_welfare, amount) VALUES (%s, %s, %s, %s)", 
                       (EMP_ID, name, has, amt))

    # - ทรัพย์สินบริษัท
    cursor.execute("""
        INSERT INTO employee_company_assets (emp_id, phone_info, other_details)
        VALUES (%s, 'Samsung A12 (เครื่องส่วนกลาง)', 'เสื้อยูนิฟอร์ม 3 ตัว')
    """, (EMP_ID,))

    # 4. สร้างข้อมูลการลงเวลา 12 เดือน (Generate Logs)
    print(f"4. จำลองการทำงานตลอดปี {YEAR} (12 เดือน)...")
    
    # Scenario ในแต่ละเดือน
    scenarios = {
        1: "pass",      # ม.ค. ✅ ผ่าน (Streak 0->1) | เงิน 300
        2: "pass",      # ก.พ. ✅ ผ่าน (Streak 1->2) | เงิน 400
        3: "late",      # มี.ค. ❌ สาย (Streak -> 0)
        4: "pass",      # เม.ย. ✅ ผ่าน (Streak 0->1) | เงิน 300
        5: "absent",    # พ.ค. ❌ ขาด (Streak -> 0)
        6: "pass",      # มิ.ย. ✅ ผ่าน (Streak 0->1) | เงิน 300
        7: "leave",     # ก.ค. ❌ ลาป่วย (Streak -> 0)
        8: "pass",      # ส.ค. ✅ ผ่าน (Streak 0->1) | เงิน 300
        9: "pass",      # ก.ย. ✅ ผ่าน (Streak 1->2) | เงิน 400
        10: "pass",     # ต.ค. ✅ ผ่าน (Streak 2->3) | เงิน 500
        11: "pass",     # พ.ย. ✅ ผ่าน (Streak 3->4) | เงิน 500 (Max)
        12: "ot_heavy"  # ธ.ค. ✅ ผ่าน + ทำ OT เยอะ (Streak 4->5) | เงิน 500
    }

    for month in range(1, 13):
        num_days = calendar.monthrange(YEAR, month)[1]
        start_date = date(YEAR, month, 1)
        end_date = date(YEAR, month, num_days)
        scenario = scenarios.get(month, "pass")
        
        # print(f"   - สร้างเดือน {month}/{YEAR}: Scenario = {scenario}")

        current = start_date
        while current <= end_date:
            # วันอาทิตย์ = วันหยุด
            if current.weekday() == 6: 
                cursor.execute("INSERT INTO employee_daily_records (emp_id, work_date, status) VALUES (%s, %s, 'วันหยุด') ON CONFLICT DO NOTHING", (EMP_ID, current))
                current += timedelta(days=1)
                continue

            # เวลามาตรฐาน (07:45 - 17:15)
            t_in = datetime.combine(current, time(7, 45))
            t_out = datetime.combine(current, time(17, 15))
            
            # --- ปรับเวลาตาม Scenario ---
            if scenario == "late" and current.day == 5:
                t_in = datetime.combine(current, time(8, 10)) # สาย 10 นาที

            elif scenario == "absent" and current.day == 10:
                # ขาดงาน (ไม่ลงเวลา)
                current += timedelta(days=1)
                continue
                
            elif scenario == "leave" and current.day == 15:
                # ลาป่วย
                cursor.execute("""
                    INSERT INTO employee_leave_records (emp_id, leave_date, leave_type, num_days, reason)
                    VALUES (%s, %s, 'ลาป่วย', 1.0, 'ปวดหัว ตัวร้อน')
                """, (EMP_ID, current))
                # (วันลาไม่ต้องสแกนนิ้ว)
                current += timedelta(days=1)
                continue
            
            elif scenario == "ot_heavy":
                # ธันวาคม งานเยอะ เลิก 2 ทุ่มทุกวัน (OT ~2.5 ชม.)
                t_out = datetime.combine(current, time(20, 0))

            # Insert Logs
            cursor.execute("INSERT INTO time_attendance_logs (emp_id, scan_timestamp) VALUES (%s, %s)", (EMP_ID, t_in))
            cursor.execute("INSERT INTO time_attendance_logs (emp_id, scan_timestamp) VALUES (%s, %s)", (EMP_ID, t_out))
            
            current += timedelta(days=1)

    conn.commit()
    conn.close()
    print("\n✅ สำเร็จ! ข้อมูลพร้อมใช้งานแล้วครับ")
    print(f"   User ID: {EMP_ID} (นายสมชาย ใจสู้)")
    print(f"   ปีข้อมูล: {YEAR} (กรุณาเลือกปีนี้ในโปรแกรม)")
    print("---------------------------------------------------")

if __name__ == "__main__":
    create_full_year_data()