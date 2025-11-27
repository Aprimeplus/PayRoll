# (ไฟล์: auto_verify_payroll.py)
# (เวอร์ชันอัปเกรด - ตรวจสอบทั้ง "สาย" และ "ขาดงาน")

import hr_database
from datetime import date
import psycopg2

def get_real_employees():
    """ค้นหาพนักงานจริงๆ ในระบบมาให้อัตโนมัติ"""
    conn = hr_database.get_db_connection()
    if not conn: return None, None
    
    monthly_emp = None
    daily_emp = None
    
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT emp_id, fname, lname, emp_type, salary FROM employees")
            all_emps = cursor.fetchall()
            
            print(f"--- 🕵️‍♂️ พบพนักงานทั้งหมด {len(all_emps)} คนในระบบ ---")
            
            for emp in all_emps:
                emp_id, fname, lname, emp_type, salary = emp
                emp_type_str = str(emp_type)
                
                # หาตัวแทนพนักงานรายวัน
                if ("รายวัน" in emp_type_str or "daily" in emp_type_str.lower()) and not daily_emp:
                    daily_emp = emp_id
                    print(f"   👉 พบตัวแทนรายวัน: {fname} ({emp_id})")
                
                # หาตัวแทนพนักงานรายเดือน (ที่ไม่ใช่รายวัน)
                elif ("รายวัน" not in emp_type_str and "daily" not in emp_type_str.lower()) and not monthly_emp:
                    monthly_emp = emp_id
                    print(f"   👉 พบตัวแทนรายเดือน: {fname} ({emp_id})")
                
                if daily_emp and monthly_emp: break
                
    except Exception as e:
        print(f"Error finding employees: {e}")
    finally:
        conn.close()
        
    return monthly_emp, daily_emp

def test_calculation(emp_id):
    if not emp_id: return

    print(f"\n-------------------------------------------------------")
    print(f"🔍 กำลังตรวจสอบการคำนวณของ: {emp_id}")
    
    # 1. ดึงข้อมูลพนักงานมาดูประเภท
    emp = hr_database.load_single_employee(emp_id)
    if not emp:
        print("❌ ไม่พบข้อมูลรายละเอียดพนักงาน")
        return

    salary = float(emp.get('salary', 0))
    emp_type = emp.get('emp_type', 'ไม่ระบุ')
    print(f"👤 ชื่อ: {emp['fname']} {emp['lname']}")
    print(f"📋 ประเภท: {emp_type}")
    print(f"💰 ฐานเงินเดือน/ค่าแรง: {salary:,.2f} บาท")
    
    # 2. คำนวณเรทที่ควรจะเป็น (Manual Check)
    expected_hourly = 0
    expected_daily = 0
    
    if "รายวัน" in emp_type or "daily" in str(emp_type).lower():
        expected_daily = salary
        expected_hourly = salary / 8
        print(f"   📝 [Logic รายวัน] ควรตกวันละ {expected_daily:.2f}, ชม.ละ {expected_hourly:.2f}")
    else:
        expected_daily = salary / 30
        expected_hourly = expected_daily / 8
        print(f"   📝 [Logic รายเดือน] ควรตกวันละ {expected_daily:.2f}, ชม.ละ {expected_hourly:.2f}")

    # 3. สั่งให้ระบบคำนวณจริง (ช่วงเดือนปัจจุบัน)
    start = date(2025, 10, 1) 
    end = date(2025, 10, 31) 
    
    result = hr_database.calculate_payroll_for_employee(emp_id, start, end, 0, 0)
    
    if result:
        penalty_hours = result['debug_penalty_hours']
        absent_days = result['debug_absent_days'] # ดึงวันขาดงานมาด้วย
        actual_deduction = result['late_absent_deduction']
        
        print(f"\n📊 ผลลัพธ์จากระบบ:")
        print(f"   - มาสาย:      {penalty_hours:.2f} ชม.")
        print(f"   - ขาดงาน:     {absent_days:.2f} วัน")
        print(f"   - หักเงินรวม: {actual_deduction:,.2f} บาท")
        
        # 4. คำนวณย้อนกลับเพื่อพิสูจน์ (Proof)
        calculated_deduction = (absent_days * expected_daily) + (penalty_hours * expected_hourly)
        
        print(f"\n🧮 พิสูจน์คำตอบ:")
        print(f"   ({absent_days} วัน x {expected_daily:.2f}) + ({penalty_hours} ชม. x {expected_hourly:.2f})")
        print(f"   = {calculated_deduction:,.2f} บาท")
        
        if abs(actual_deduction - calculated_deduction) < 0.01:
            print("   ✅ ถูกต้อง! (ยอดเงินตรงกันเป๊ะ)")
        else:
            diff = actual_deduction - calculated_deduction
            print(f"   ❌ ไม่ตรงกัน! (ต่างกัน {diff:.2f} บาท)")
            
    else:
        print("❌ เกิดข้อผิดพลาดในการคำนวณ")

if __name__ == "__main__":
    monthly_id, daily_id = get_real_employees()
    
    if monthly_id: test_calculation(monthly_id)
    else: print("\n⚠️ ไม่พบพนักงานรายเดือนในระบบ")
        
    if daily_id: test_calculation(daily_id)
    else: print("\n⚠️ ไม่พบพนักงานรายวันในระบบ")