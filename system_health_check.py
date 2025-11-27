# (ไฟล์: system_health_check.py)
import psycopg2
from psycopg2 import extras

# --- ตั้งค่าการเชื่อมต่อ ---
DB_CONFIG = {
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "host": "192.168.1.51",
    "port": "5432"
}

# รายชื่อตารางที่ "ต้องมี" ในระบบ
REQUIRED_TABLES = [
    "employees", 
    "employee_welfare", 
    "salary_history", 
    "employee_training_records",  # (ตารางใหม่)
    "employee_company_assets",    # (ตารางใหม่)
    "time_attendance_logs", 
    "employee_leave_records", 
    "employee_late_records", 
    "employee_warning_records",
    "users", 
    "pending_employee_changes", 
    "employee_documents",
    "company_holidays", 
    "company_settings", 
    "company_locations"
]

# คอลัมน์สำคัญที่ "ต้องมี" (ที่เราเพิ่งเพิ่มไป)
REQUIRED_COLUMNS = {
    "employees": [
        "emergency_contact_name", "ref_person_name", # (เพิ่งเพิ่ม)
        "probation_end_date", "work_location"
    ],
    "employee_company_assets": ["computer_info", "line_id"],
    "employee_training_records": ["course_name", "cost"]
}

# คอลัมน์ที่ "ต้องเป็นชนิด DATE" (ผลจากการ Migration)
DATE_COLUMNS = {
    "employees": ["birth_date", "start_date", "termination_date", "probation_end_date"],
    "employee_leave_records": ["leave_date"],
    "company_holidays": ["holiday_date"]
}

def check_system():
    print("\n🏥 --- เริ่มต้นการตรวจสอบสุขภาพระบบ (System Health Check) ---")
    print(f"📡 กำลังเชื่อมต่อฐานข้อมูลที่: {DB_CONFIG['host']}...")
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("✅ การเชื่อมต่อ: สำเร็จ\n")
    except Exception as e:
        print(f"❌ การเชื่อมต่อ: ล้มเหลว ({e})")
        return

    # --- 1. ตรวจสอบตาราง (Tables) ---
    print("📋 [1. ตรวจสอบตารางในฐานข้อมูล]")
    cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'")
    existing_tables = [row[0] for row in cursor.fetchall()]
    
    missing_tables = []
    for table in REQUIRED_TABLES:
        if table in existing_tables:
            # นับจำนวนข้อมูล
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {table}")
                count = cursor.fetchone()[0]
                print(f"   ✅ พบตาราง: {table:<30} | มีข้อมูล: {count} รายการ")
            except:
                print(f"   ✅ พบตาราง: {table:<30} | (ตรวจสอบจำนวนไม่ได้)")
        else:
            print(f"   ❌ ไม่พบตาราง: {table}")
            missing_tables.append(table)
            
    if missing_tables:
        print(f"   ⚠️ สรุป: มีตารางหายไป {len(missing_tables)} ตาราง! (ระบบอาจทำงานไม่สมบูรณ์)")
    else:
        print("   ✨ สรุป: ตารางครบถ้วนสมบูรณ์")

    # --- 2. ตรวจสอบคอลัมน์สำคัญ (Schema) ---
    print("\n🔍 [2. ตรวจสอบคอลัมน์ที่เพิ่งเพิ่มใหม่]")
    all_cols_ok = True
    for table, columns in REQUIRED_COLUMNS.items():
        if table not in existing_tables: continue
        
        cursor.execute(f"SELECT column_name FROM information_schema.columns WHERE table_name = '{table}'")
        existing_columns_in_db = [row[0] for row in cursor.fetchall()]
        
        for col in columns:
            if col in existing_columns_in_db:
                print(f"   ✅ {table}.{col:<25} -> มีอยู่จริง")
            else:
                print(f"   ❌ {table}.{col:<25} -> ไม่พบ!")
                all_cols_ok = False
    
    if not all_cols_ok:
        print("   ⚠️ คำเตือน: มีคอลัมน์สำคัญขาดหายไป กรุณาตรวจสอบ init_db หรือรัน SQL เพิ่มเติม")

    # --- 3. ตรวจสอบชนิดข้อมูลวันที่ (Date Type Migration) ---
    print("\n📅 [3. ตรวจสอบชนิดข้อมูลวันที่ (ต้องเป็น DATE เท่านั้น)]")
    all_dates_ok = True
    for table, columns in DATE_COLUMNS.items():
        if table not in existing_tables: continue
        
        for col in columns:
            cursor.execute(f"""
                SELECT data_type FROM information_schema.columns 
                WHERE table_name = '{table}' AND column_name = '{col}'
            """)
            result = cursor.fetchone()
            if result:
                dtype = result[0]
                if dtype == 'date':
                    print(f"   ✅ {table}.{col:<20} : {dtype.upper()}")
                else:
                    print(f"   ❌ {table}.{col:<20} : {dtype.upper()} (ควรเป็น DATE)")
                    all_dates_ok = False
            else:
                print(f"   ⚠️ ไม่พบคอลัมน์ {col} ใน {table}")

    if all_dates_ok:
        print("   ✨ ยอดเยี่ยม: การ Migration วันที่สมบูรณ์ 100%")
    else:
        print("   ⚠️ คำเตือน: ยังมีบางคอลัมน์เป็น TEXT อยู่")

    conn.close()
    print("\n🏁 --- การตรวจสอบเสร็จสิ้น ---")

if __name__ == "__main__":
    check_system()