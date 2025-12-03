import hr_database
from datetime import datetime, timedelta, date
import calendar

# --- ตั้งค่าการทดสอบ ---
TEST_EMP_ID = "TEST_AUTO_VERIFY"
CURRENT_M = 12
CURRENT_Y = 2024

def setup_test_employee():
    """เคลียร์ข้อมูลและสร้างพนักงานใหม่"""
    conn = hr_database.get_db_connection()
    cursor = conn.cursor()
    # ล้างข้อมูลเก่า
    for table in ["employees", "employee_daily_records", "employee_leave_records", "employee_late_records"]:
        cursor.execute(f"DELETE FROM {table} WHERE emp_id = %s", (TEST_EMP_ID,))
    
    # สร้างใหม่ (รายวัน)
    cursor.execute("""
        INSERT INTO employees (emp_id, fname, lname, emp_type, salary, department)
        VALUES (%s, 'Test', 'Auto', 'รายวัน', 500, 'IT')
    """, (TEST_EMP_ID,))
    conn.commit()
    conn.close()

def clear_month_history(month, year):
    """ล้างข้อมูลเฉพาะเดือนที่ระบุ (เพื่อให้มั่นใจว่าไม่มีข้อมูลค้าง)"""
    conn = hr_database.get_db_connection()
    cursor = conn.cursor()
    last_day = calendar.monthrange(year, month)[1]
    start_date = date(year, month, 1)
    end_date = date(year, month, last_day)
    
    cursor.execute("DELETE FROM employee_daily_records WHERE emp_id = %s AND work_date BETWEEN %s AND %s", (TEST_EMP_ID, start_date, end_date))
    cursor.execute("DELETE FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (TEST_EMP_ID, start_date, end_date))
    cursor.execute("DELETE FROM employee_late_records WHERE emp_id = %s AND late_date BETWEEN %s AND %s", (TEST_EMP_ID, start_date, end_date))
    conn.commit()
    conn.close()

def simulate_month_history(month, year, condition="GOOD"):
    """สร้างประวัติจำลอง: GOOD (ปกติ), BAD (สาย), EMPTY (ไม่มีข้อมูล)"""
    conn = hr_database.get_db_connection()
    cursor = conn.cursor()
    
    # ล้างก่อนสร้างเสมอ
    clear_month_history(month, year)
    
    if condition == "GOOD":
        # สร้างวันทำงานปกติ 1 วัน
        work_date = date(year, month, 1)
        cursor.execute("INSERT INTO employee_daily_records (emp_id, work_date, status, ot_hours) VALUES (%s, %s, 'ปกติ', 0)", (TEST_EMP_ID, work_date))
        
    elif condition == "BAD":
        # สร้างวันสาย 1 วัน (มี record ทำงาน แต่สถานะสาย + มี record ในตารางสาย)
        work_date = date(year, month, 5)
        cursor.execute("INSERT INTO employee_daily_records (emp_id, work_date, status, ot_hours) VALUES (%s, %s, 'สาย', 0)", (TEST_EMP_ID, work_date))
        cursor.execute("INSERT INTO employee_late_records (emp_id, late_date, minutes_late, reason) VALUES (%s, %s, 30, 'Late')", (TEST_EMP_ID, work_date))
        
    # condition == "EMPTY" ไม่ต้องทำอะไร (เพราะเราล้างไปแล้ว)
    
    conn.commit()
    conn.close()

def run_full_verification():
    print("\n" + "="*80)
    print("🚀 เริ่มต้นการตรวจสอบระบบเบี้ยขยัน (7 Scenarios)")
    print("="*80)
    
    setup_test_employee()
    
    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 1: เพิ่งเริ่มงาน (ไม่มีประวัติเลย)")
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 300)")
    assert reward == 300
    
    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 2: ขยัน 1 เดือน (เดือน 11 ดี, เดือน 10 ว่าง)")
    simulate_month_history(11, 2024, "GOOD")
    simulate_month_history(10, 2024, "EMPTY")
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 400)")
    assert reward == 400

    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 3: ขยัน 2 เดือนติด (เดือน 11 ดี, เดือน 10 ดี)")
    simulate_month_history(11, 2024, "GOOD")
    simulate_month_history(10, 2024, "GOOD")
    simulate_month_history(9, 2024, "EMPTY")
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 500)")
    assert reward == 500

    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 4: ขยัน 3 เดือนติด (เดือน 11, 10, 9 ดี -> ตันเพดาน)")
    simulate_month_history(11, 2024, "GOOD")
    simulate_month_history(10, 2024, "GOOD")
    simulate_month_history(9, 2024, "GOOD")
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 500)")
    assert reward == 500

    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 5: พลาด! (เดือน 11 มาสาย)")
    simulate_month_history(11, 2024, "BAD")
    # เดือนเก่าๆ จะดีแค่ไหนก็ไม่ช่วยอะไร ถ้าเดือนล่าสุดพัง
    simulate_month_history(10, 2024, "GOOD") 
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 300)")
    assert reward == 300 # รีเซ็ตนับ 1 ใหม่

    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 6: เริ่มกลับมาดี (เดือน 11 ดี, แต่เดือน 10 เคยพลาด)")
    simulate_month_history(11, 2024, "GOOD")
    simulate_month_history(10, 2024, "BAD")
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 400)")
    assert reward == 400 # นับต่อเนื่องได้แค่ 1 เดือน (เดือน 11) เจอเดือน 10 ตัดจบ

    # --------------------------------------------------------------------------------
    print("\n🔹 Scenario 7: เว้นช่วงทำงาน (เดือน 11 ดี, เดือน 10 ไม่มาทำ/ไม่มีกะ, เดือน 9 ดี)")
    # กรณีพนักงานรายวัน บางเดือนอาจไม่ได้มาทำ (สัญญาขาดช่วง)
    simulate_month_history(11, 2024, "GOOD")
    simulate_month_history(10, 2024, "EMPTY")
    simulate_month_history(9, 2024, "GOOD")
    
    reward = hr_database.get_auto_diligence_reward(TEST_EMP_ID, CURRENT_M, CURRENT_Y)
    print(f"   => ผลลัพธ์: {reward} (คาดหวัง: 400)")
    # Logic ปัจจุบัน: เจอเดือนว่าง (WorkDays=0) -> STOP 
    # ดังนั้นจะนับได้แค่เดือน 11 เดือนเดียว = Streak 1 -> ได้ 400
    assert reward == 400 

    print("\n" + "="*80)
    print("✅✅✅ ผ่านทุกการทดสอบ (System Verified) ✅✅✅")
    print("="*80)

if __name__ == "__main__":
    try:
        run_full_verification()
    except AssertionError as e:
        print(f"\n❌ การทดสอบล้มเหลว! {e}")
    except Exception as e:
        print(f"\n❌ Error: {e}")