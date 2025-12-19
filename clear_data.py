import hr_database
import psycopg2

def clear_test_data():
    conn = hr_database.get_db_connection()
    if not conn:
        print("❌ เชื่อมต่อ Database ไม่ได้")
        return

    try:
        with conn.cursor() as cursor:
            print("⏳ กำลังเริ่มล้างข้อมูล...")
            
            # 1. ลบข้อมูลพนักงานและข้อมูลที่เกี่ยวข้องทั้งหมด (Payroll, Leave, OT, ประวัติ ฯลฯ)
            # CASCADE จะช่วยลบข้อมูลในตารางลูกให้เองโดยอัตโนมัติ
            cursor.execute("TRUNCATE TABLE employees CASCADE;")
            
            # 2. ลบข้อมูล Log ต่างๆ ที่อาจไม่ได้ผูก Foreign Key ไว้
            cursor.execute("TRUNCATE TABLE time_attendance_logs RESTART IDENTITY;")
            cursor.execute("TRUNCATE TABLE audit_logs RESTART IDENTITY;")
            cursor.execute("TRUNCATE TABLE email_queue RESTART IDENTITY;")
            
            # (หมายเหตุ: ตาราง users, company_settings, company_holidays, company_locations จะยังอยู่)
            
        conn.commit()
        print("✅ ลบข้อมูลพนักงานและข้อมูล Transaction ออกหมดแล้ว!")
        print("   (ข้อมูล User, วันหยุด, และการตั้งค่าบริษัท ยังอยู่ครบ)")
        
    except Exception as e:
        conn.rollback()
        print(f"❌ เกิดข้อผิดพลาด: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    confirm = input("⚠️ คำเตือน: ข้อมูลพนักงานและเงินเดือนทั้งหมดจะถูกลบถาวร!\nพิมพ์ 'DELETE' เพื่อยืนยัน: ")
    if confirm == 'DELETE':
        clear_test_data()
    else:
        print("ยกเลิกรายการ")