import psycopg2
from psycopg2 import sql, extras
import tkinter as tk
from tkinter import messagebox
import json 
import hashlib 
from datetime import datetime, time, timedelta 
import pandas as pd

# --- !! สำคัญ !! ---
# กรอกรายละเอียดการเชื่อมต่อ PostgreSQL ของคุณที่นี่
def get_db_connection():
    try:
        conn = psycopg2.connect(
            dbname="aplus_com_test",
            user="app_user",
            password="cailfornia123",
            host="192.168.1.51",
            port="5432"
        )
        return conn
    except Exception as e:
        messagebox.showerror("Connection Error", f"ไม่สามารถเชื่อมต่อฐานข้อมูลได้:\n{e}")
        return None
         
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(stored_hash, provided_password):
    return stored_hash == hash_password(provided_password)

def date_to_thai_str(date_obj):
    """แปลง date object (ค.ศ.) จาก DB -> String (พ.ศ.) เพื่อแสดงผล"""
    if not date_obj: return ""
    try:
        return f"{date_obj.day:02d}/{date_obj.month:02d}/{date_obj.year + 543}"
    except: return ""

def thai_str_to_date(date_str):
    """แปลง String (พ.ศ.) จาก UI -> date object (ค.ศ.) เพื่อบันทึก"""
    if not date_str: return None
    try:
        d, m, y_be = map(int, date_str.split('/'))
        return datetime(y_be - 543, m, d).date()
    except: return None

def insert_scan_logs(log_data_list):
    """
    (ส่วนที่ 4) บันทึกข้อมูล Log ดิบ (จาก Excel) ลงในตาราง time_attendance_logs
    log_data_list คือ list of tuples: [(emp_id, scan_timestamp), ...]
    """
    if not log_data_list:
        return 0
        
    conn = get_db_connection()
    if not conn: return 0

    inserted_count = 0
    try:
        with conn.cursor() as cursor:
            # ใช้ execute_values เพื่อประสิทธิภาพสูงสุดในการ Insert ข้อมูลจำนวนมาก
            insert_query = """
                INSERT INTO time_attendance_logs (emp_id, scan_timestamp)
                VALUES %s
                ON CONFLICT DO NOTHING; 
            """
            extras.execute_values(
                cursor,
                insert_query,
                log_data_list,
                template="(%s, %s)",
                page_size=100 
            )
            inserted_count = cursor.rowcount 
            conn.commit()
            
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Insert Logs)", f"ไม่สามารถบันทึกข้อมูล Log ได้:\n{e}")
        return 0
    finally:
        if conn: conn.close()
        
    return inserted_count

def init_db():
    """สร้างตารางทั้งหมดหากยังไม่มี (รันครั้งเดียวตอนเปิดแอป)"""
    conn = get_db_connection()
    if not conn: return
    try:
        with conn.cursor() as cursor:
            # 1. ตารางพนักงาน
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employees (
                    emp_id TEXT PRIMARY KEY, fname TEXT, nickname TEXT, lname TEXT,
                    birth_date DATE, age TEXT, id_card TEXT, phone TEXT, address TEXT,
                    current_address TEXT, emp_type TEXT, start_date DATE, work_exp TEXT,
                    position TEXT, department TEXT, status TEXT, salary REAL,termination_date DATE, 
                    termination_reason TEXT,
                    health_status TEXT, health_detail TEXT, bank_account_no TEXT,
                    bank_name TEXT, bank_branch TEXT, bank_account_name TEXT,
                    bank_account_type TEXT, sso_start_date DATE, sso_end_date DATE,
                    sso_start_action_date DATE, sso_end_action_date DATE,
                    leave_annual_days INTEGER, leave_sick_days INTEGER,
                    leave_ordination_days INTEGER, leave_maternity_days INTEGER,leave_personal_days INTEGER,
                    guarantee_enabled INTEGER DEFAULT 0,
                    guarantor_name TEXT,
                    guarantee_amount REAL
                );
            """)
            
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS email_queue (
                        queue_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        period_month INTEGER,
                        period_year INTEGER,
                        pdf_path TEXT NOT NULL,
                        receiver_email TEXT,
                        status TEXT DEFAULT 'pending', -- pending, approved, rejected, sent
                        requested_by TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                    );
                """)
            except Exception as e:
                print(f"Warning (Create email_queue): {e}")
                conn.rollback()

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_daily_records (
                    record_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    work_date DATE NOT NULL,
                    status TEXT,             -- 'ทำงาน', 'หยุด', 'ลา', 'ขาด'
                    ot_hours REAL DEFAULT 0,
                    car_type TEXT,           -- '-', 'กระบะ', 'เฮี๊ยบ'
                    trip_count INTEGER DEFAULT 0,
                    UNIQUE(emp_id, work_date),  -- ห้ามซ้ำวันเดียวกัน
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            try:
                cursor.execute("""
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS is_sales BOOLEAN DEFAULT FALSE,
                    ADD COLUMN IF NOT EXISTS sale_type TEXT,       -- 'Inbound', 'Outbound'
                    ADD COLUMN IF NOT EXISTS commission_plan TEXT; -- 'Plan A', 'Plan B', ...
                """)
            except Exception: conn.rollback()

            # 2. ตารางสวัสดิการ
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_welfare (
                    welfare_id SERIAL PRIMARY KEY, emp_id TEXT NOT NULL,
                    welfare_name TEXT NOT NULL, has_welfare INTEGER NOT NULL,
                    amount REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS payroll_records (
                    payroll_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    period_month INTEGER NOT NULL,
                    period_year INTEGER NOT NULL,
                    payment_date DATE,
                    base_salary REAL DEFAULT 0,
                    position_allowance REAL DEFAULT 0,
                    ot_pay REAL DEFAULT 0,
                    commission REAL DEFAULT 0,
                    bonus REAL DEFAULT 0,
                    driving_allowance REAL DEFAULT 0,
                    other_income REAL DEFAULT 0,
                    total_income REAL DEFAULT 0,
                    sso_deduct REAL DEFAULT 0,
                    tax_deduct REAL DEFAULT 0,       
                    provident_fund REAL DEFAULT 0,
                    loan_deduct REAL DEFAULT 0,
                    late_deduct REAL DEFAULT 0,
                    other_deduct REAL DEFAULT 0,
                    total_deduct REAL DEFAULT 0,
                    net_salary REAL DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE,
                    UNIQUE(emp_id, period_month, period_year)
                );
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS payroll_records (
                    payroll_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    period_month INTEGER NOT NULL,
                    period_year INTEGER NOT NULL,  -- ปี ค.ศ.
                    payment_date DATE,
                    
                    -- รายได้
                    base_salary REAL DEFAULT 0,
                    position_allowance REAL DEFAULT 0,
                    ot_pay REAL DEFAULT 0,
                    commission REAL DEFAULT 0,
                    bonus REAL DEFAULT 0,
                    driving_allowance REAL DEFAULT 0,
                    other_income REAL DEFAULT 0,
                    total_income REAL DEFAULT 0,
                    
                    -- รายหัก
                    sso_deduct REAL DEFAULT 0,
                    tax_deduct REAL DEFAULT 0,       
                    provident_fund REAL DEFAULT 0,
                    loan_deduct REAL DEFAULT 0,
                    late_deduct REAL DEFAULT 0,
                    other_deduct REAL DEFAULT 0,
                    total_deduct REAL DEFAULT 0,
                    
                    -- สุทธิ
                    net_salary REAL DEFAULT 0,
                    
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE,
                    UNIQUE(emp_id, period_month, period_year) -- ห้ามจ่ายซ้ำเดือนเดิม
                );
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_driving_details (
                    detail_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    work_date DATE NOT NULL,
                    car_type TEXT,
                    license_plate TEXT,
                    driver_name TEXT,
                    delivery_date DATE,
                    trip_cost REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_driving_records (
                    record_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    drive_date DATE NOT NULL,
                    car_type TEXT NOT NULL,  -- 'กระบะ' หรือ 'เฮี๊ยบ'
                    trips_count INTEGER DEFAULT 1,
                    total_amount REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 3. ตารางประวัติเงินเดือน
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS salary_history (
                    history_id SERIAL PRIMARY KEY, emp_id TEXT NOT NULL,
                    adjustment_year TEXT, new_salary REAL,position_allowance REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 4. ตารางประวัติฝึกอบรม
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_training_records (
                    id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    training_date DATE,
                    course_name TEXT,
                    cost REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 5. ตารางทรัพย์สินบริษัท (ที่ขาดหายไป)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_company_assets (
                    asset_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL UNIQUE, 
                    computer_info TEXT,
                    phone_info TEXT,
                    phone_number TEXT,
                    sim_type TEXT,
                    carrier TEXT,
                    company_email TEXT,
                    line_id TEXT,
                    line_linked_phone TEXT,
                    facebook TEXT,
                    employee_card_id TEXT,
                    other_details TEXT,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 6. ตารางผู้ใช้งานระบบ
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    user_id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK (role IN ('hr', 'approver'))
                );
            """)

            # 7. ตารางรายการรออนุมัติ
            cursor.execute("""
                 CREATE TABLE IF NOT EXISTS pending_employee_changes (
                    change_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    change_data JSONB NOT NULL,
                    requested_by TEXT NOT NULL,
                    request_timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    status TEXT DEFAULT 'pending' NOT NULL CHECK (status IN ('pending', 'approved', 'rejected')),
                    approved_by TEXT,
                    approval_timestamp TIMESTAMP,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 8. ตารางเอกสารแนบ
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_documents (
                    doc_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    description TEXT, 
                    file_path TEXT NOT NULL,
                    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)
            
            # --- Update Columns (ถ้าตารางมีอยู่แล้วแต่คอลัมน์ไม่ครบ) ---
            try:
                cursor.execute("""
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS probation_days INTEGER DEFAULT 90,
                    ADD COLUMN IF NOT EXISTS probation_end_date DATE,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score TEXT,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score_2 TEXT,
                    ADD COLUMN IF NOT EXISTS sso_hospital TEXT,
                    ADD COLUMN IF NOT EXISTS work_location TEXT; 
                    ADD COLUMN IF NOT EXISTS emergency_contact_name TEXT,
                    ADD COLUMN IF NOT EXISTS emergency_contact_phone TEXT,
                    ADD COLUMN IF NOT EXISTS emergency_contact_relation TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_name TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_phone TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_relation TEXT;
                """)
            except Exception: conn.rollback()
            
            try:
                cursor.execute("""
                    ALTER TABLE salary_history
                    ADD COLUMN IF NOT EXISTS new_position TEXT,
                    ADD COLUMN IF NOT EXISTS assessment_score TEXT;
                """)
            except Exception: conn.rollback()
                
            # 9. ตารางวันหยุดบริษัท
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS company_holidays (
                        holiday_id SERIAL PRIMARY KEY,
                        holiday_date DATE NOT NULL UNIQUE,
                        description TEXT
                    );
                """)
            except Exception: conn.rollback()    
            
            # 10. ตารางการตั้งค่าและสาขา
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS company_settings (
                        setting_key TEXT PRIMARY KEY,
                        setting_value TEXT
                    );
                """)
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS company_locations (
                        loc_id SERIAL PRIMARY KEY,
                        loc_name TEXT NOT NULL,
                        loc_type TEXT,
                        google_link TEXT,
                        CONSTRAINT unique_loc_name UNIQUE(loc_name)
                    );
                """)
            except Exception: conn.rollback()

            # 11. ตารางการลา
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS employee_leave_records (
                        leave_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        leave_date DATE NOT NULL,
                        leave_type TEXT NOT NULL, 
                        num_days REAL DEFAULT 1.0,
                        reason TEXT,
                        FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                    );
                """)
            except Exception: conn.rollback()

            try:
                cursor.execute("""
                    ALTER TABLE employee_leave_records
                    ADD CONSTRAINT unique_emp_leave_date UNIQUE (emp_id, leave_date);
                """)
            except Exception: conn.rollback() 
            
            try:
                cursor.execute("""
                    ALTER TABLE employee_leave_records
                    ADD COLUMN IF NOT EXISTS leave_start_time TIME,
                    ADD COLUMN IF NOT EXISTS leave_end_time TIME;
                """)
            except Exception: conn.rollback()
                
            # 12. ตารางการมาสายและใบเตือน
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS employee_late_records (
                        late_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        late_date DATE NOT NULL,
                        minutes_late INTEGER,
                        reason TEXT,
                        FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                    );
                """)
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS employee_warning_records (
                        warning_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        warning_date DATE NOT NULL,
                        reason TEXT NOT NULL,
                        warning_level TEXT, 
                        FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                    );
                """)
            except Exception: conn.rollback()

            # 13. ตาราง Log สแกนนิ้ว
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS time_attendance_logs (
                        log_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        scan_timestamp TIMESTAMP NOT NULL
                    );
                """)
                # ลบ constraint เก่า (ถ้ามี)
                try:
                    cursor.execute("""
                        ALTER TABLE time_attendance_logs 
                        DROP CONSTRAINT IF EXISTS time_attendance_logs_emp_id_fkey;
                    """)
                except Exception: conn.rollback() 

                # เพิ่ม constraint กันซ้ำ
                try:
                    cursor.execute("""
                        ALTER TABLE time_attendance_logs
                        ADD CONSTRAINT unique_log_entry UNIQUE (emp_id, scan_timestamp);
                    """)
                except Exception: conn.rollback()
            except Exception: conn.rollback()

            # 14. เพิ่ม User เริ่มต้น
            cursor.execute(
                """
                INSERT INTO users (username, password_hash, role) 
                VALUES (%s, %s, %s)
                ON CONFLICT (username) DO UPDATE SET
                    password_hash = EXCLUDED.password_hash,
                    role = EXCLUDED.role
                """,
                ('hr_user', hash_password('password123'), 'hr')
            )
            cursor.execute(
                """
                INSERT INTO users (username, password_hash, role) 
                VALUES (%s, %s, %s)
                ON CONFLICT (username) DO UPDATE SET
                    password_hash = EXCLUDED.password_hash,
                    role = EXCLUDED.role
                """,
                ('approver_user', hash_password('approve123'), 'approver')
            )

        conn.commit()
    except Exception as e:
        messagebox.showerror("DB Init Error", f"ไม่สามารถสร้างตารางได้:\n{e}")
    finally:
        if conn: conn.close()

def get_company_setting(key):
    conn = get_db_connection()
    if not conn: return ""
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT setting_value FROM company_settings WHERE setting_key = %s", (key,))
            result = cursor.fetchone()
            return result[0] if result else ""
    except Exception as e:
        print(f"Error loading setting {key}: {e}")
        return ""
    finally:
        if conn: conn.close()

def calculate_payroll_for_employee(emp_id, start_date, end_date, user_inputs=None):
    """
    (สมองหลัก Payroll V20.0 - Full Option)
    - รองรับ OT Auto
    - รองรับ ค่าเที่ยว (กระบะ/เฮี๊ยบ/ฟรี)
    - รองรับ การหักสาย/ขาดงาน (คืนชีพ Logic กลับมา)
    """
    if user_inputs is None: user_inputs = {}

    # 1. กำหนดกฎการเข้างาน
    WORK_RULES = {
        "สำนักงานใหญ่": { "standard_in": time(9,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120, "work_hours_per_day": 8.0, "required_duration_minutes": 480, "break_start": time(12,0), "break_end": time(13,0) },
        "คลังสินค้า": { "standard_in": time(8,0), "tier_1_cutoff": time(8,30), "penalty_1_mins": 60, "penalty_2_mins": 120, "work_hours_per_day": 8.0, "required_duration_minutes": 480, "break_start": time(12,0), "break_end": time(13,0) },
        "default": { "standard_in": time(9,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120, "work_hours_per_day": 8.0, "required_duration_minutes": 480, "break_start": time(12,0), "break_end": time(13,0) }
    }

    result = {
        "emp_id": emp_id, 
        "base_salary": 0.0, "position_allowance": 0.0,
        "ot": 0.0, "commission": 0.0, "bonus": 0.0, "other_income": 0.0, 
        "driving_allowance": 0.0, 
        "total_income": 0.0,
        "sso": 0.0, "tax": 0.0, "provident_fund": 0.0, "loan": 0.0, 
        "late_deduct": 0.0, "other_deduct": 0.0, 
        "total_deduct": 0.0,
        "net_salary": 0.0,
        "debug_penalty_hours": 0.0, "debug_absent_days": 0.0, "debug_worked_days": 0.0 
    }

    conn = get_db_connection()
    if not conn: return result

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. ข้อมูลพนักงาน
            emp_info = load_single_employee(emp_id)
            if not emp_info or not emp_info.get("salary"): return result

            salary_from_db = float(emp_info.get("salary", 0.0))
            emp_type = emp_info.get("emp_type", "")
            
            cursor.execute("SELECT position_allowance FROM salary_history WHERE emp_id = %s ORDER BY history_id DESC LIMIT 1", (emp_id,))
            pos_row = cursor.fetchone()
            result["position_allowance"] = float(pos_row[0]) if pos_row and pos_row[0] else 0.0

            # 2. User Inputs
            manual_ot = float(user_inputs.get('ot', 0)) 
            result["commission"] = float(user_inputs.get('commission', 0))
            result["bonus"] = float(user_inputs.get('bonus', 0))
            result["other_income"] = float(user_inputs.get('other_income', 0))
            result["tax"] = float(user_inputs.get('tax', 0))
            result["provident_fund"] = float(user_inputs.get('provident_fund', 0))
            result["loan"] = float(user_inputs.get('loan', 0))
            result["other_deduct"] = float(user_inputs.get('other_deduct', 0))

            # 3. ดึงข้อมูลประกอบ (วันหยุด, ลา, สแกนนิ้ว, ค่าเที่ยว)
            work_location = emp_info.get('work_location', "default") or "default"
            if work_location not in WORK_RULES: work_location = "default"
            rules = WORK_RULES[work_location]
            
            cursor.execute("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holiday_dates_set = {row['holiday_date'] for row in cursor.fetchall()}
            
            cursor.execute("SELECT leave_date, num_days, leave_type, leave_start_time, leave_end_time FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            leave_records_dict = {}
            for row in cursor.fetchall():
                leave_records_dict[row['leave_date']] = dict(row)

            cursor.execute("SELECT scan_timestamp FROM time_attendance_logs WHERE emp_id = %s AND scan_timestamp::date BETWEEN %s AND %s ORDER BY scan_timestamp", (emp_id, start_date, end_date))
            scans = cursor.fetchall()
            scans_by_date = {}
            for scan in scans:
                d = scan['scan_timestamp'].date()
                if d not in scans_by_date: scans_by_date[d] = []
                scans_by_date[d].append(scan['scan_timestamp'])

            cursor.execute("""
                SELECT work_date, total_amount, ot_hours 
                FROM employee_daily_records 
                WHERE emp_id = %s AND work_date BETWEEN %s AND %s
            """, (emp_id, start_date, end_date))
            daily_records_dict = {}
            for row in cursor.fetchall():
                daily_records_dict[row['work_date']] = dict(row)

            # --- 4. เริ่มคำนวณ (Loop รายวัน) ---
            total_penalty_minutes = 0.0
            total_absent_days = 0.0
            actual_worked_days = 0.0
            auto_driving_allowance = 0.0
            total_ot_money = 0.0
            
            all_dates = pd.date_range(start_date, end_date, freq='D')

            for current_date in all_dates:
                c_date = current_date.date()
                is_holiday = (current_date.weekday() == 6 or c_date in holiday_dates_set)
                
                daily_rec = daily_records_dict.get(c_date)
                leave_info = leave_records_dict.get(c_date)
                scans_today = scans_by_date.get(c_date)
                
                # A. ค่าเที่ยว
                if daily_rec:
                    auto_driving_allowance += float(daily_rec.get('total_amount', 0) or 0)

                # B. OT
                ot_hrs = float(daily_rec.get('ot_hours', 0) or 0) if daily_rec else 0
                if ot_hrs > 0:
                    # สูตร OT (1.5 เท่า)
                    base_calc = salary_from_db / 30.0 if "รายวัน" not in emp_type else salary_from_db
                    hourly_rate = base_calc / 8.0
                    daily_ot_pay = ot_hrs * hourly_rate * 1.5
                    total_ot_money += daily_ot_pay

                # C. เช็คการมาทำงาน / สาย / ขาด
                is_present = False
                if scans_today: is_present = True
                if daily_rec and (daily_rec.get('total_amount',0) > 0 or daily_rec.get('ot_hours',0) > 0): is_present = True
                
                # ถ้าลา ก็ถือว่าไม่ขาด (แต่จะหักเงินไหม อยู่ที่ประเภทลา)
                is_leave = False
                if leave_info: is_leave = True

                if not is_holiday and not is_leave:
                    if not is_present:
                        # ขาดงาน (Absent)
                        total_absent_days += 1.0
                    elif scans_today:
                        # มาทำงาน -> เช็คสาย (Late)
                        t_in = min(scans_today).time()
                        
                        # Logic เช็คสายตามกฎ (Tier 1/2)
                        # ถ้าลาชั่วโมง (เช่น ลาเช้า) ต้องขยับเวลาเข้างานตามจริง
                        effective_start = rules["standard_in"]
                        
                        if t_in > rules["tier_1_cutoff"]:
                            total_penalty_minutes += rules["penalty_2_mins"]
                        elif t_in > rules["standard_in"]:
                            total_penalty_minutes += rules["penalty_1_mins"]
                            
                        actual_worked_days += 1.0
                    else:
                         # มาทำงาน (มี record) แต่ไม่สแกน (เช่น ขับรถ) -> นับว่ามา ไม่สาย
                         actual_worked_days += 1.0
                
                elif is_holiday and is_present:
                    # มาทำวันหยุด
                    actual_worked_days += 1.0 # (รายวันได้ตังเพิ่ม)

            # --- 5. สรุปยอด ---
            result["driving_allowance"] = auto_driving_allowance
            result["ot"] = manual_ot + total_ot_money
            
            # คำนวณหักสาย/ขาด (Late Deduct)
            penalty_hours = total_penalty_minutes / 60.0
            
            # สูตรคำนวณเงินหัก
            deduct_amount = 0.0
            
            if salary_from_db > 0:
                if "รายวัน" in emp_type or "daily" in str(emp_type).lower():
                    # รายวัน:
                    # - ฐานเงินเดือน = ค่าแรง x วันที่มาทำจริง
                    # - สาย = หักเป็นนาที/ชั่วโมง จากวันที่มา
                    # - ขาด = ไม่ได้ตังอยู่แล้ว (ไม่ต้องหักซ้ำ)
                    result["base_salary"] = salary_from_db * actual_worked_days
                    
                    daily_rate = salary_from_db
                    hourly_rate = daily_rate / 8.0
                    deduct_amount = penalty_hours * hourly_rate
                    
                else:
                    # รายเดือน:
                    # - ฐานเงินเดือน = เต็มเดือน
                    # - สาย = หักตามจริง
                    # - ขาด = หักตามวัน (Daily Rate)
                    result["base_salary"] = salary_from_db
                    
                    daily_rate = salary_from_db / 30.0
                    hourly_rate = daily_rate / 8.0
                    
                    deduct_absent = total_absent_days * daily_rate
                    deduct_late = penalty_hours * hourly_rate
                    deduct_amount = deduct_absent + deduct_late

            result["late_deduct"] = deduct_amount

            # --- 6. Final Sum ---
            result["sso"] = round(result["base_salary"] * 0.05)
            if result["sso"] > 750: result["sso"] = 750
            if result["sso"] < 83: result["sso"] = 83

            result["total_income"] = (
                result["base_salary"] + result["position_allowance"] + 
                result["ot"] + result["commission"] + result["bonus"] + result["other_income"] +
                result["driving_allowance"] 
            )
            result["total_deduct"] = (
                result["sso"] + result["tax"] + result["provident_fund"] + 
                result["late_deduct"] + result["loan"] + result["other_deduct"]
            )
            result["net_salary"] = result["total_income"] - result["total_deduct"]

    except Exception as e:
        print(f"Payroll Error: {e}")
        return None
    finally:
        if conn: conn.close()
        
    return result

def save_company_setting(key, value):
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO company_settings (setting_key, setting_value)
                VALUES (%s, %s)
                ON CONFLICT (setting_key) DO UPDATE SET
                    setting_value = EXCLUDED.setting_value;
                """,
                (key, value)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Setting)", f"ไม่สามารถบันทึก Setting ได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def get_company_locations():
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("SELECT * FROM company_locations ORDER BY loc_id ASC")
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถโหลดข้อมูลสาขาได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def add_company_location(name, loc_type, link):
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO company_locations (loc_name, loc_type, google_link)
                VALUES (%s, %s, %s)
                ON CONFLICT (loc_name) DO UPDATE SET
                    loc_type = EXCLUDED.loc_type,
                    google_link = EXCLUDED.google_link;
                """,
                (name, loc_type, link)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Location)", f"ไม่สามารถบันทึกข้อมูลสาขาได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def delete_company_location(loc_id):
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM company_locations WHERE loc_id = %s", (loc_id,))
            conn.commit()
            return cursor.rowcount > 0
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Location)", f"ไม่สามารถลบข้อมูลสาขาได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def get_employee_count_by_dept(year_be):
    """
    (เวอร์ชัน Date Migration) นับพนักงาน ณ สิ้นปี (พ.ศ.) ที่ระบุ
    """
    conn = get_db_connection()
    if not conn: return []
    
    try:
        # แปลงปี พ.ศ. -> วันที่สิ้นปี ค.ศ. (Date Object)
        year_ce = int(year_be) - 543
        end_of_year_date = datetime(year_ce, 12, 31).date()

        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # (ดูสิครับ! SQL สั้นและอ่านง่ายขึ้นเยอะ ไม่ต้องแปลงไปมา)
            query = """
                SELECT 
                    COALESCE(department, 'N/A') as dept, 
                    COUNT(emp_id) as count
                FROM employees
                WHERE
                    start_date <= %s  -- (ใครเริ่มงานก่อนสิ้นปี)
                AND
                    (
                        termination_date IS NULL -- (ยังไม่ออก)
                        OR
                        termination_date > %s -- (หรือออกหลังสิ้นปีนั้น)
                    )
                GROUP BY dept
                ORDER BY dept ASC
            """
            cursor.execute(query, (end_of_year_date, end_of_year_date))
            return [dict(row) for row in cursor.fetchall()]
            
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถนับจำนวนพนักงานได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def authenticate_user(username, password):
    conn = get_db_connection()
    if not conn: return None
    user_info = None
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT user_id, username, password_hash, role FROM users WHERE username = %s", (username,))
            result = cursor.fetchone()
            if result and verify_password(result[2], password):
                user_info = {"user_id": result[0], "username": result[1], "role": result[3]}
    except Exception as e:
        messagebox.showerror("Auth Error", f"เกิดข้อผิดพลาดในการตรวจสอบผู้ใช้:\n{e}")
    finally:
        if conn: conn.close()
    return user_info

# === Employee Data Management ===

def _upsert_employee_data(cursor, data):
    try:
        emp_id = data.get("id")
        emp_cols = [ 
            "emp_id", "fname", "nickname", "lname", "birth_date", "age", "id_card", "phone",
            "address", "current_address", "emp_type", "start_date", "work_exp", "position",
            "department", "status", "salary","termination_date", "termination_reason", "health_status", "health_detail",
            "bank_account_no", "bank_name", "bank_branch", "bank_account_name",
            "is_sales", "sale_type", "commission_plan",
            "bank_account_type", "sso_start_date", "sso_end_date",
            "sso_start_action_date", "sso_end_action_date", "leave_annual_days","sso_hospital",
            "leave_sick_days", "leave_ordination_days", "leave_maternity_days","leave_personal_days",
            "guarantee_enabled", "guarantor_name", "guarantee_amount","probation_days", "probation_end_date",
            "probation_assessment_score", "probation_assessment_score_2",
            "work_location", 
            # คอลัมน์ใหม่
            "emergency_contact_name", "emergency_contact_phone", "emergency_contact_relation",
            "ref_person_name", "ref_person_phone", "ref_person_relation"
        ]
        
        # สร้าง dict เริ่มต้น (ค่าจะเป็น None ถ้าชื่อไม่ตรง)
        sql_data = {col: data.get(col) for col in emp_cols} 
        sql_data['emp_id'] = data.get('id')
        
        # --- (!!! เพิ่มส่วนนี้ครับ: จับคู่ชื่อตัวแปรจาก UI ให้ตรงกับ DB !!!) ---
        sql_data['emergency_contact_name'] = data.get('emergency_name')
        sql_data['emergency_contact_phone'] = data.get('emergency_phone')
        sql_data['emergency_contact_relation'] = data.get('emergency_relation')
        sql_data['ref_person_name'] = data.get('ref_name')
        sql_data['ref_person_phone'] = data.get('ref_phone')
        sql_data['ref_person_relation'] = data.get('ref_relation')
        # --------------------------------------------------------------

        # ... (ส่วนแปลงวันที่ เหมือนเดิม) ...
        sql_data['birth_date'] = thai_str_to_date(data.get('birth'))
        sql_data['start_date'] = thai_str_to_date(data.get('start_date'))
        sql_data['termination_date'] = thai_str_to_date(data.get('termination_date'))
        sql_data['probation_end_date'] = thai_str_to_date(data.get('probation_end_date'))
        
        sql_data['sso_start_date'] = thai_str_to_date(data.get('sso_start'))
        sql_data['sso_end_date'] = thai_str_to_date(data.get('sso_end'))
        sql_data['sso_start_action_date'] = thai_str_to_date(data.get('sso_start_action'))
        sql_data['sso_end_action_date'] = thai_str_to_date(data.get('sso_end_action'))
        
        # ... (ส่วนอื่นๆ เหมือนเดิม ไม่ต้องแก้) ...
        sql_data['work_exp'] = data.get('exp')
        sql_data['health_status'] = data.get('health')
        sql_data['bank_account_no'] = data.get('account')
        sql_data['bank_name'] = data.get('bank')
        sql_data['bank_branch'] = data.get('branch')
        sql_data['bank_account_name'] = data.get('account_name')
        sql_data['bank_account_type'] = data.get('account_type')
        sql_data['sso_hospital'] = data.get('sso_hospital')

        leave_annual_raw = data.get('leave_annual')
        leave_sick_raw = data.get('leave_sick')
        leave_ordination_raw = data.get('leave_ordination')
        leave_maternity_raw = data.get('leave_maternity')
        leave_personal_raw = data.get('leave_personal')

        sql_data['leave_annual_days'] = int(leave_annual_raw or 0) if not isinstance(leave_annual_raw, dict) else 0
        sql_data['leave_sick_days'] = int(leave_sick_raw or 0) if not isinstance(leave_sick_raw, dict) else 0
        sql_data['leave_ordination_days'] = int(leave_ordination_raw or 0) if not isinstance(leave_ordination_raw, dict) else 0
        sql_data['leave_maternity_days'] = int(leave_maternity_raw or 0) if not isinstance(leave_maternity_raw, dict) else 0
        sql_data['leave_personal_days'] = int(leave_personal_raw or 0) if not isinstance(leave_personal_raw, dict) else 0
        
        salary_str = data.get('salary', '')
        sql_data['salary'] = float(salary_str) if salary_str else 0.0
        sql_data['guarantee_enabled'] = 1 if data.get('guarantee_enabled') else 0
        guarantee_amount_str = data.get('guarantee_amount', '')
        sql_data['guarantee_amount'] = float(guarantee_amount_str) if guarantee_amount_str else 0.0
        probation_days_str = data.get('probation_days', '')
        sql_data['probation_days'] = int(probation_days_str) if probation_days_str else 90 
        sql_data['probation_assessment_score'] = data.get('probation_assessment_score')
        sql_data['probation_assessment_score_2'] = data.get('probation_assessment_score_2')
        
        sql_data['work_location'] = data.get('work_location') 
        
        insert_sql_query = sql.SQL("""
            INSERT INTO employees ({columns}) VALUES ({values})
            ON CONFLICT (emp_id) DO UPDATE SET {updates}
        """).format(
            columns=sql.SQL(', ').join(map(sql.Identifier, emp_cols)),
            values=sql.SQL(', ').join(map(sql.Placeholder, emp_cols)),
            updates=sql.SQL(', ').join(
                sql.Composed([sql.Identifier(c), sql.SQL(" = "), sql.Placeholder(c)])
                for c in emp_cols if c != "emp_id"
            )
        )
        cursor.execute(insert_sql_query, sql_data)

        # ... (ส่วน Welfare, Salary History, Training, Assets เหมือนเดิม) ...
        cursor.execute("DELETE FROM employee_welfare WHERE emp_id = %s", (emp_id,))
        welfare_options = data.get("welfare_options", [])
        welfare_flags = data.get("welfare", [])
        welfare_amounts = data.get("welfare_amounts", [])
        for i, name in enumerate(welfare_options):
            has_welfare = 1 if i < len(welfare_flags) and welfare_flags[i] else 0
            amount_str = welfare_amounts[i] if i < len(welfare_amounts) else ""
            amount = float(amount_str) if amount_str else 0.0 
            cursor.execute(
                "INSERT INTO employee_welfare (emp_id, welfare_name, has_welfare, amount) VALUES (%s, %s, %s, %s)",
                (emp_id, name, has_welfare, amount)
            )

        cursor.execute("DELETE FROM salary_history WHERE emp_id = %s", (emp_id,))
        salary_history = data.get("salary_history", [])
        for item in salary_history:
            if item.get("year") or item.get("salary") or item.get("new_position"): 
                salary_str = item.get("salary")
                new_salary = float(salary_str) if salary_str else 0.0 
                allowance_str = item.get("position_allowance")
                position_allowance = float(allowance_str) if allowance_str else 0.0
                new_pos = item.get("new_position", "")
                assess_score = item.get("assessment_score", "")
                cursor.execute(
                    """
                    INSERT INTO salary_history 
                    (emp_id, adjustment_year, new_salary, position_allowance, new_position, assessment_score) 
                    VALUES (%s, %s, %s, %s, %s, %s)
                    """,
                    (emp_id, item.get("year"), new_salary, position_allowance, new_pos, assess_score) 
                )

        cursor.execute("DELETE FROM employee_training_records WHERE emp_id = %s", (emp_id,))
        training_history = data.get("training_history", [])
        for item in training_history:
            t_date = thai_str_to_date(item.get("date"))
            c_name = item.get("course_name", "")
            c_cost = float(item.get("cost", 0) or 0)
            
            cursor.execute(
                """
                INSERT INTO employee_training_records (emp_id, training_date, course_name, cost)
                VALUES (%s, %s, %s, %s)
                """,
                (emp_id, t_date, c_name, c_cost)
            )   

        cursor.execute("DELETE FROM employee_company_assets WHERE emp_id = %s", (emp_id,))
        assets = data.get("assets", {})
        cursor.execute("""
            INSERT INTO employee_company_assets (
                emp_id, computer_info, phone_info, phone_number, sim_type, carrier,
                company_email, line_id, line_linked_phone, facebook, employee_card_id, other_details
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            emp_id,
            assets.get('computer', ''), assets.get('phone', ''), assets.get('number', ''),
            assets.get('sim', ''), assets.get('carrier', ''), assets.get('email', ''),
            assets.get('line', ''), assets.get('line_phone', ''), assets.get('facebook', ''),
            assets.get('card_id', ''), assets.get('others', '')
        ))  

        return True
    except Exception as e:
        print(f"Error in _upsert_employee_data: {e}") 
        messagebox.showerror("Database Upsert Error", f"เกิดข้อผิดพลาดในการบันทึกข้อมูลหลัก:\n{e}")
        return False

def load_single_employee(emp_id):
    conn = get_db_connection()
    if not conn: 
        return None
    
    try:
        emp_id = str(emp_id) 

        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("SELECT * FROM employees WHERE emp_id = %s", (emp_id,))
            emp_row = cursor.fetchone()
            if not emp_row:
                return None
            employee_data = dict(emp_row)
            
            # (กำหนดคีย์ที่คาดว่าจะเป็น String หรือ None)
            string_keys = [
                "fname", "nickname", "lname", "birth_date", "age", "id_card", "phone","sale_type", "commission_plan",
                "address", "current_address", "emp_type", "start_date", "work_exp",
                "position", "department", "status", "termination_reason",
                "health_status", "health_detail", "bank_account_no",
                "bank_name", "bank_branch", "bank_account_name",
                "bank_account_type", "sso_start_date", "sso_end_date",
                "sso_start_action_date", "sso_end_action_date",
                "sso_hospital",
                "guarantor_name", "probation_end_date",
                "probation_assessment_score", "probation_assessment_score_2",
                "work_location"
            ]

            for key in string_keys:
                if employee_data.get(key) is None:
                    employee_data[key] = ""
            
            employee_data['id'] = employee_data.get('emp_id')
            
            # --- (จุดที่แก้ไข: แปลง Date Object (ค.ศ.) เป็น String (พ.ศ.) เพื่อแสดงผล UI) ---
            employee_data['birth'] = date_to_thai_str(employee_data.get('birth_date'))
            employee_data['start_date'] = date_to_thai_str(employee_data.get('start_date'))
            employee_data['termination_date'] = date_to_thai_str(employee_data.get('termination_date'))
            employee_data['probation_end_date'] = date_to_thai_str(employee_data.get('probation_end_date'))
            employee_data['is_sales'] = bool(employee_data.get('is_sales'))
            
            employee_data['sso_start'] = date_to_thai_str(employee_data.get('sso_start_date'))
            employee_data['sso_end'] = date_to_thai_str(employee_data.get('sso_end_date'))
            employee_data['sso_start_action'] = date_to_thai_str(employee_data.get('sso_start_action_date'))
            employee_data['sso_end_action'] = date_to_thai_str(employee_data.get('sso_end_action_date'))
            
            employee_data['exp'] = employee_data.get('work_exp')
            employee_data['health'] = employee_data.get('health_status')
            employee_data['account'] = employee_data.get('bank_account_no')
            employee_data['bank'] = employee_data.get('bank_name')
            employee_data['branch'] = employee_data.get('bank_branch')
            employee_data['sso_hospital'] = employee_data.get('sso_hospital')
            # --- (จบส่วนแปลงวันที่) ---
            
            annual_days = employee_data.get('leave_annual_days')
            employee_data['leave_annual'] = str(annual_days) if annual_days is not None else ""
            sick_days = employee_data.get('leave_sick_days')
            employee_data['leave_sick'] = str(sick_days) if sick_days is not None else ""
            ordination_days = employee_data.get('leave_ordination_days')
            employee_data['leave_ordination'] = str(ordination_days) if ordination_days is not None else ""
            maternity_days = employee_data.get('leave_maternity_days')
            employee_data['leave_maternity'] = str(maternity_days) if maternity_days is not None else ""
            personal_days = employee_data.get('leave_personal_days')
            employee_data['leave_personal'] = str(personal_days) if personal_days is not None else ""
            employee_data['emergency_name'] = employee_data.get('emergency_contact_name', '')
            employee_data['emergency_phone'] = employee_data.get('emergency_contact_phone', '')
            employee_data['emergency_relation'] = employee_data.get('emergency_contact_relation', '')
            employee_data['ref_name'] = employee_data.get('ref_person_name', '')
            employee_data['ref_phone'] = employee_data.get('ref_person_phone', '')
            employee_data['ref_relation'] = employee_data.get('ref_person_relation', '')
            employee_data['guarantee_enabled'] = bool(employee_data.get('guarantee_enabled')) 
            employee_data['guarantor_name'] = employee_data.get('guarantor_name') 
            
            guar_amount = employee_data.get('guarantee_amount')
            employee_data['guarantee_amount'] = str(guar_amount) if guar_amount else ""
            
            prob_days = employee_data.get('probation_days')
            employee_data['probation_days'] = str(prob_days) if prob_days is not None else "90"
            
            employee_data['probation_assessment_score'] = employee_data.get('probation_assessment_score')
            employee_data['probation_assessment_score_2'] = employee_data.get('probation_assessment_score_2')
            
            employee_data['work_location'] = employee_data.get('work_location') 
            
            cursor.execute(
                "SELECT welfare_name, has_welfare, amount FROM employee_welfare WHERE emp_id = %s ORDER BY welfare_id",
                (emp_id,)
            )
            welfare_rows = cursor.fetchall()
            employee_data['welfare_options'] = [row['welfare_name'] for row in welfare_rows]
            employee_data['welfare'] = [bool(row['has_welfare']) for row in welfare_rows]
            employee_data['welfare_amounts'] = [str(row['amount']) if row['amount'] else "" for row in welfare_rows]
            cursor.execute(
                """
                SELECT adjustment_year, new_salary, position_allowance, 
                       new_position, assessment_score 
                FROM salary_history 
                WHERE emp_id = %s ORDER BY history_id
                """,
                (emp_id,)
            )
            history_rows = cursor.fetchall()
            employee_data['salary_history'] = [
                {
                    "year": row['adjustment_year'], 
                    "salary": str(row['new_salary']) if row['new_salary'] else "",
                    "position_allowance": str(row['position_allowance']) if row.get('position_allowance') else "",
                    "new_position": row.get('new_position', ""), 
                    "assessment_score": row.get('assessment_score', "") 
                }
                for row in history_rows
            ]
            cursor.execute(
                "SELECT training_date, course_name, cost FROM employee_training_records WHERE emp_id = %s ORDER BY training_date DESC",
                (emp_id,)
            )
            training_rows = cursor.fetchall()
            employee_data['training_history'] = [
                {
                    "date": date_to_thai_str(row['training_date']),
                    "course_name": row['course_name'],
                    "cost": f"{row['cost']:.2f}" if row['cost'] is not None else "0.00"
                }
                for row in training_rows
            ]
             
            cursor.execute("SELECT * FROM employee_company_assets WHERE emp_id = %s", (emp_id,))
            asset_row = cursor.fetchone()
            
            assets_data = {
                'computer': '', 'phone': '', 'number': '', 'sim': '', 'carrier': '',
                'email': '', 'line': '', 'line_phone': '', 'facebook': '',
                'card_id': '', 'others': ''
            }
            
            if asset_row:
                row_dict = dict(asset_row)
                assets_data = {
                    'computer': row_dict.get('computer_info', '') or '',
                    'phone': row_dict.get('phone_info', '') or '',
                    'number': row_dict.get('phone_number', '') or '',
                    'sim': row_dict.get('sim_type', '') or '',
                    'carrier': row_dict.get('carrier', '') or '',
                    'email': row_dict.get('company_email', '') or '',
                    'line': row_dict.get('line_id', '') or '',
                    'line_phone': row_dict.get('line_linked_phone', '') or '',
                    'facebook': row_dict.get('facebook', '') or '',
                    'card_id': row_dict.get('employee_card_id', '') or '',
                    'others': row_dict.get('other_details', '') or ''
                }
            
            employee_data['assets'] = assets_data

            return employee_data
            
    except Exception as e:
        messagebox.showerror("Load Error", f"ไม่สามารถโหลดข้อมูลพนักงานได้:\n{e}")
        return None
    finally:
        if conn: 
            conn.close()

def save_employee(data, current_user):
    conn = get_db_connection()
    if not conn: return False
    emp_id = data.get("id")
    role = current_user.get("role")
    username = current_user.get("username")
    is_new_employee = not bool(load_single_employee(emp_id)) 
    try:
        with conn.cursor() as cursor:
            if is_new_employee:
                success = _upsert_employee_data(cursor, data)
                if not success: raise Exception("Failed to insert new employee data.")
                message = "บันทึกข้อมูลพนักงานใหม่เรียบร้อยแล้ว"
            elif role == 'hr':
                change_data_json = json.dumps(data)
                cursor.execute(
                    """
                    INSERT INTO pending_employee_changes
                    (emp_id, change_data, requested_by, status)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (emp_id, change_data_json, username, 'pending')
                )
                message = "ส่งคำขอแก้ไขข้อมูลเรียบร้อยแล้ว รอการอนุมัติ"
            elif role == 'approver':
                success = _upsert_employee_data(cursor, data)
                if not success: raise Exception("Failed to update employee data directly.")
                message = "อัปเดตข้อมูลพนักงานเรียบร้อยแล้ว"
            else: 
                 raise Exception("Invalid user role for saving data.")
            conn.commit()
            return True, message 
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Save Error", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}")
        return False, f"เกิดข้อผิดพลาด: {e}"
    finally:
        if conn: conn.close()

def _reconstruct_data(emp_rows, welfare_rows, history_rows):
    employees_dict = {}
    for emp in emp_rows:
        emp_data = dict(emp)
        emp_data['department'] = emp_data.get('department') 
        emp_data['status'] = emp_data.get('status') 
        emp_data['termination_date'] = emp_data.get('termination_date', "") 
        emp_data['termination_reason'] = emp_data.get('termination_reason', "")
        emp_data['id'] = emp_data.get('emp_id')
        emp_data['birth'] = emp_data.get('birth_date')
        emp_data['exp'] = emp_data.get('work_exp')
        emp_data['health'] = emp_data.get('health_status')
        emp_data['account'] = emp_data.get('bank_account_no')
        emp_data['bank'] = emp_data.get('bank_name')
        emp_data['branch'] = emp_data.get('bank_branch')
        emp_data['sso_start'] = emp_data.get('sso_start_date')
        emp_data['sso_end'] = emp_data.get('sso_end_date')
        emp_data['sso_start_action'] = emp_data.get('sso_start_action_date')
        emp_data['sso_end_action'] = emp_data.get('sso_end_action_date')
        emp_data['sso_hospital'] = emp_data.get('sso_hospital', "")
        emp_data['leave_annual'] = emp_data.get('leave_annual_days')
        emp_data['leave_sick'] = emp_data.get('leave_sick_days')
        emp_data['leave_ordination'] = emp_data.get('leave_ordination_days')
        emp_data['leave_maternity'] = emp_data.get('leave_maternity_days')
        emp_data['leave_personal'] = emp_data.get('leave_personal_days')
        emp_data['guarantee_enabled'] = bool(emp_data.get('guarantee_enabled'))
        emp_data['guarantor_name'] = emp_data.get('guarantor_name', "")
        guar_amount_rec = emp_data.get('guarantee_amount')
        emp_data['guarantee_amount'] = str(guar_amount_rec) if guar_amount_rec else ""
        prob_days_rec = emp_data.get('probation_days')
        emp_data['probation_days'] = str(prob_days_rec) if prob_days_rec is not None else "90"
        emp_data['probation_end_date'] = emp_data.get('probation_end_date', "")
        emp_data['probation_assessment_score'] = emp_data.get('probation_assessment_score', "")
        emp_data['probation_assessment_score_2'] = emp_data.get('probation_assessment_score_2', "")
        emp_data['work_location'] = emp_data.get('work_location', "") 
        emp_data['welfare_options'] = []
        emp_data['welfare'] = []
        emp_data['welfare_amounts'] = []
        emp_data['salary_history'] = []
        employees_dict[emp_data['id']] = emp_data
    for wel in welfare_rows:
        emp_id = wel['emp_id']
        if emp_id in employees_dict:
            employees_dict[emp_id]['welfare_options'].append(wel['welfare_name'])
            employees_dict[emp_id]['welfare'].append(bool(wel['has_welfare']))
            employees_dict[emp_id]['welfare_amounts'].append(str(wel['amount']) if wel['amount'] else "")
    for hist in history_rows:
        emp_id = hist['emp_id']
        if emp_id in employees_dict:
            employees_dict[emp_id]['salary_history'].append(
                {
                    "year": hist['adjustment_year'], 
                    "salary": str(hist['new_salary']) if hist['new_salary'] else "",
                    "position_allowance": str(hist['position_allowance']) if hist.get('position_allowance') else "",
                    "new_position": hist.get('new_position', ""), 
                    "assessment_score": hist.get('assessment_score', "") 
                }
            )
    return list(employees_dict.values())

def load_all_employees():
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("SELECT * FROM employees ORDER BY emp_id")
            emp_rows = cursor.fetchall()
            cursor.execute("SELECT * FROM employee_welfare ORDER BY emp_id, welfare_id")
            welfare_rows = cursor.fetchall()
            cursor.execute("SELECT * FROM salary_history ORDER BY emp_id, history_id")
            history_rows = cursor.fetchall()
            return _reconstruct_data(emp_rows, welfare_rows, history_rows)
    except Exception as e:
        messagebox.showerror("Load Error", f"ไม่สามารถโหลดข้อมูลพนักงานทั้งหมดได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def delete_employee(emp_id):
    conn = get_db_connection()
    if not conn: return False
    try:
        emp_id = str(emp_id) 
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM employees WHERE emp_id = %s", (emp_id,))
            conn.commit()
            return cursor.rowcount > 0 
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Delete Error", f"ไม่สามารถลบข้อมูลได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def search_employees(search_term):
    conn = get_db_connection()
    if not conn: return []
    
    like_term = f"%{search_term}%" 
    
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT DISTINCT emp_id FROM employees
                WHERE emp_id ILIKE %s
                   OR fname ILIKE %s
                   OR lname ILIKE %s
                   OR position ILIKE %s
                   OR department ILIKE %s
                   OR id_card ILIKE %s
            """, (like_term, like_term, like_term, like_term, like_term, like_term))
            
            emp_ids = [row['emp_id'] for row in cursor.fetchall()]
            
            if not emp_ids:
                return [] 

            cursor.execute(sql.SQL("SELECT * FROM employees WHERE emp_id IN %s ORDER BY emp_id"), (tuple(emp_ids),))
            emp_rows = cursor.fetchall()
            cursor.execute(sql.SQL("SELECT * FROM employee_welfare WHERE emp_id IN %s ORDER BY emp_id, welfare_id"), (tuple(emp_ids),))
            welfare_rows = cursor.fetchall()
            cursor.execute(sql.SQL("SELECT * FROM salary_history WHERE emp_id IN %s ORDER BY emp_id, history_id"), (tuple(emp_ids),))
            history_rows = cursor.fetchall()
            
            return _reconstruct_data(emp_rows, welfare_rows, history_rows)
            
    except Exception as e:
        messagebox.showerror("Search Error", f"ไม่สามารถค้นหาข้อมูลได้:\n{e}")
        return []
    finally:
        if conn: conn.close()


def get_pending_changes():
    conn = get_db_connection()
    if not conn: return []
    pending_list = []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT pc.change_id, pc.emp_id, e.fname, e.lname, pc.requested_by, pc.request_timestamp
                FROM pending_employee_changes pc
                JOIN employees e ON pc.emp_id = e.emp_id
                WHERE pc.status = 'pending'
                ORDER BY pc.request_timestamp DESC
            """)
            pending_list = cursor.fetchall()
    except Exception as e:
        messagebox.showerror("Load Pending Error", f"ไม่สามารถโหลดรายการรออนุมัติได้:\n{e}")
    finally:
        if conn: conn.close()
    return [dict(row) for row in pending_list]

def get_change_details(change_id):
    conn = get_db_connection()
    if not conn: return None, None
    change_data = None
    current_data = None
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("SELECT emp_id, change_data FROM pending_employee_changes WHERE change_id = %s", (change_id,))
            result = cursor.fetchone()
            if result:
                change_data = result['change_data'] 
                emp_id = result['emp_id']
                current_data = load_single_employee(emp_id) 
    except Exception as e:
        messagebox.showerror("Load Detail Error", f"ไม่สามารถโหลดรายละเอียดการแก้ไขได้:\n{e}")
    finally:
        if conn: conn.close()
    return change_data, current_data 

def approve_change(change_id, approver_username):
    print(f"DEBUG: เข้าสู่ฟังก์ชัน approve_change (ID={change_id}, User={approver_username})") # <--- Debug
    
    conn = get_db_connection()
    if not conn: 
        print("DEBUG: เชื่อมต่อฐานข้อมูลไม่ได้")
        return False
        
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. ดึงข้อมูล JSON ที่รอการแก้ไขออกมา
            cursor.execute("SELECT change_data FROM pending_employee_changes WHERE change_id = %s AND status = 'pending'", (change_id,))
            result = cursor.fetchone()
            
            if not result:
                print("DEBUG: ไม่พบรายการ หรือสถานะไม่ใช่ pending")
                raise Exception("ไม่พบคำขอแก้ไข หรือคำขอไม่ได้อยู่ในสถานะ 'pending'")

            employee_data_to_update = result['change_data']
            print(f"DEBUG: ข้อมูลที่จะอัปเดต: {str(employee_data_to_update)[:100]}...") # ดูข้อมูลย่อๆ

            # 2. สั่งอัปเดตข้อมูลจริง (_upsert_employee_data)
            # (ฟังก์ชันนี้เราแก้ไปแล้ว ให้รองรับชื่อคอลัมน์ emergency_contact_name)
            success = _upsert_employee_data(cursor, employee_data_to_update)
            
            if not success:
                print("DEBUG: _upsert_employee_data ตอบกลับมาว่าล้มเหลว")
                raise Exception("เกิดข้อผิดพลาดขณะอัปเดตข้อมูลพนักงานหลัก (_upsert ล้มเหลว)")

            # 3. เปลี่ยนสถานะคำขอเป็น 'approved'
            cursor.execute(
                """
                UPDATE pending_employee_changes
                SET status = 'approved', approved_by = %s, approval_timestamp = %s
                WHERE change_id = %s
                """,
                (approver_username, datetime.now(), change_id)
            )
            conn.commit()
            print("DEBUG: อนุมัติสำเร็จและ Commit แล้ว")
            return True

    except Exception as e:
        conn.rollback()
        print(f"DEBUG: เกิด Error ใน approve_change: {e}")
        messagebox.showerror("Approve Error", f"ไม่สามารถอนุมัติการแก้ไขได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def reject_change(change_id, rejecter_username):
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                UPDATE pending_employee_changes
                SET status = 'rejected', approved_by = %s, approval_timestamp = %s
                WHERE change_id = %s AND status = 'pending'
                """,
                (rejecter_username, datetime.now(), change_id)
            )
            conn.commit()
            return cursor.rowcount > 0
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Reject Error", f"ไม่สามารถปฏิเสธการแก้ไขได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def add_employee_document(emp_id, description, file_path):
    conn = get_db_connection()
    if not conn: return False
    try:
        emp_id = str(emp_id) 
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO employee_documents (emp_id, description, file_path)
                VALUES (%s, %s, %s)
                """,
                (emp_id, description, file_path)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error", f"ไม่สามารถบันทึกข้อมูลเอกสารได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def get_document_by_description(emp_id, description):
    conn = get_db_connection()
    if not conn: return None
    try:
        emp_id = str(emp_id) 
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute(
                """
                SELECT doc_id, file_path FROM employee_documents
                WHERE emp_id = %s AND description = %s
                ORDER BY uploaded_at DESC
                LIMIT 1
                """,
                (emp_id, description)
            )
            doc = cursor.fetchone()
            return dict(doc) if doc else None
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงข้อมูลเอกสารได้:\n{e}")
        return None
    finally:
        if conn: conn.close()

def delete_document(doc_id):
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM employee_documents WHERE doc_id = %s", (doc_id,))
            conn.commit()
            return cursor.rowcount > 0 
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error", f"ไม่สามารถลบข้อมูลเอกสารได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def add_employee_leave(emp_id, leave_date, leave_type, num_days, reason, 
                       leave_start_time=None, leave_end_time=None):
    """
    (ส่วนที่ 3) บันทึกข้อมูลการลาใหม่ (รองรับการ Update ทับ และ เวลา)
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        emp_id = str(emp_id) 
        with conn.cursor() as cursor:
            
            # (เปลี่ยนกลับไปใช้ "ON CONFLICT (emp_id, leave_date)" (ชื่อคอลัมน์))
            cursor.execute(
                """
                INSERT INTO employee_leave_records
                (emp_id, leave_date, leave_type, num_days, reason, leave_start_time, leave_end_time)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (emp_id, leave_date) DO UPDATE SET 
                    leave_type = EXCLUDED.leave_type,
                    num_days = EXCLUDED.num_days,
                    reason = EXCLUDED.reason,
                    leave_start_time = EXCLUDED.leave_start_time,
                    leave_end_time = EXCLUDED.leave_end_time
                """,
                (emp_id, leave_date, leave_type, num_days, reason, leave_start_time, leave_end_time)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        
        # (Error ที่คุณเห็น [InvalidColumnReference] จะหายไป)
        if "unique_emp_leave_date" in str(e):
             messagebox.showwarning("ผิดพลาด", f"พนักงาน {emp_id} มีการลาในวันที่ {leave_date} อยู่แล้ว (Constraint Error)")
             return False
        
        messagebox.showerror("DB Error (Leave)", f"ไม่สามารถบันทึกข้อมูลการลาได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def add_employee_late(emp_id, late_date, minutes_late, reason):
    """
    (ส่วนที่ 3) บันทึกข้อมูลการมาสายใหม่ ลงในตาราง employee_late_records
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        emp_id = str(emp_id) 
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO employee_late_records
                (emp_id, late_date, minutes_late, reason)
                VALUES (%s, %s, %s, %s)
                """,
                (emp_id, late_date, minutes_late, reason)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Late)", f"ไม่สามารถบันทึกข้อมูลการมาสายได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def add_employee_warning(emp_id, warning_date, reason, warning_level):
    """
    (ส่วนที่ 3) บันทึกข้อมูลใบเตือนใหม่ ลงในตาราง employee_warning_records
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        emp_id = str(emp_id) 
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO employee_warning_records
                (emp_id, warning_date, reason, warning_level)
                VALUES (%s, %s, %s, %s)
                """,
                (emp_id, warning_date, reason, warning_level)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Warning)", f"ไม่สามารถบันทึกข้อมูลใบเตือนได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def get_attendance_summary(emp_id, year):
    """
    (ส่วนที่ 3) สรุปข้อมูล ลา/สาย/ใบเตือน ทั้งหมดของพนักงาน
    โดยกรองตามปี (ค.ศ.) ที่ระบุ
    """
    summary = {
        'total_leave_days': 0.0,
        'leave_by_type': {
            'ลาป่วย': 0.0,
            'ลากิจ': 0.0,
            'ลาพักร้อน': 0.0,
            'ลาอื่นๆ': 0.0
        },
        'total_late_times': 0,
        'total_late_minutes': 0,
        'max_late_minutes': 0,
        'avg_late_minutes': 0.0,
        'total_warnings': 0
    }
    
    conn = get_db_connection()
    if not conn: 
        return summary 

    try:
        emp_id = str(emp_id) 
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            
            cursor.execute(
                """
                SELECT 
                    leave_type, 
                    SUM(num_days) as total_days
                FROM employee_leave_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM leave_date) = %s
                GROUP BY leave_type
                """,
                (emp_id, year)
            )
            leave_results = cursor.fetchall()
            
            total_leave = 0.0
            for row in leave_results:
                leave_type = row['leave_type']
                total_days = row['total_days'] or 0.0
                if leave_type in summary['leave_by_type']:
                    summary['leave_by_type'][leave_type] = float(total_days)
                total_leave += float(total_days)
            summary['total_leave_days'] = total_leave

            cursor.execute(
                """
                SELECT 
                    COUNT(late_id) as total_times,
                    SUM(minutes_late) as total_minutes,
                    MAX(minutes_late) as max_minutes,
                    AVG(minutes_late) as avg_minutes
                FROM employee_late_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM late_date) = %s
                """,
                (emp_id, year)
            )
            late_result = cursor.fetchone()
            if late_result:
                summary['total_late_times'] = int(late_result['total_times'] or 0)
                summary['total_late_minutes'] = int(late_result['total_minutes'] or 0)
                summary['max_late_minutes'] = int(late_result['max_minutes'] or 0)
                summary['avg_late_minutes'] = float(late_result['avg_minutes'] or 0.0)

            cursor.execute(
                """
                SELECT 
                    COUNT(warning_id) as total_count
                FROM employee_warning_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM warning_date) = %s
                """,
                (emp_id, year)
            )
            warn_result = cursor.fetchone()
            if warn_result:
                summary['total_warnings'] = int(warn_result['total_count'] or 0)

    except Exception as e:
        messagebox.showerror("DB Error (Summary)", f"ไม่สามารถดึงข้อมูลสรุปได้:\n{e}")
        return summary 
    finally:
        if conn: conn.close()
        
    return summary


def get_leave_details(emp_id, year):
    """(ส่วนที่ 3) ดึงรายละเอียดการลาทั้งหมดของพนักงานตามปี"""
    conn = get_db_connection()
    if not conn: return []
    try:
        emp_id = str(emp_id) 
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute(
                """
                SELECT leave_date, leave_type, num_days, reason 
                FROM employee_leave_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM leave_date) = %s
                ORDER BY leave_date DESC
                """,
                (emp_id, year)
            )
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงรายละเอียดการลาได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def get_late_details(emp_id, year):
    """(ส่วนที่ 3) ดึงรายละเอียดการมาสายทั้งหมดของพนักงานตามปี"""
    conn = get_db_connection()
    if not conn: return []
    try:
        emp_id = str(emp_id) 
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute(
                """
                SELECT late_date, minutes_late, reason 
                FROM employee_late_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM late_date) = %s
                ORDER BY late_date DESC
                """,
                (emp_id, year)
            )
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงรายละเอียดการมาสายได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def get_warning_details(emp_id, year):
    """(ส่วนที่ 3) ดึงรายละเอียดใบเตือนทั้งหมดของพนักงานตามปี"""
    conn = get_db_connection()
    if not conn: return []
    try:
        emp_id = str(emp_id) 
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute(
                """
                SELECT warning_date, reason, warning_level
                FROM employee_warning_records
                WHERE emp_id = %s AND EXTRACT(YEAR FROM warning_date) = %s
                ORDER BY warning_date DESC
                """,
                (emp_id, year)
            )
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงรายละเอียดใบเตือนได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def get_company_holidays(year):
    """ดึงข้อมูลวันหยุดทั้งหมดในปีที่ระบุ (ปี ค.ศ.)"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute(
                """
                SELECT holiday_id, holiday_date, description 
                FROM company_holidays
                WHERE EXTRACT(YEAR FROM holiday_date) = %s
                ORDER BY holiday_date ASC
                """,
                (year,)
            )
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงข้อมูลวันหยุดได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def add_company_holiday(holiday_date, description):
    """เพิ่มวันหยุดใหม่ (date object, text)"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO company_holidays (holiday_date, description)
                VALUES (%s, %s)
                ON CONFLICT (holiday_date) DO UPDATE SET
                    description = EXCLUDED.description;
                """,
                (holiday_date, description)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Holiday)", f"ไม่สามารถบันทึกวันหยุดได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def delete_company_holiday(holiday_id):
    """ลบวันหยุดจาก ID"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM company_holidays WHERE holiday_id = %s", (holiday_id,))
            conn.commit()
            return cursor.rowcount > 0
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Holiday)", f"ไม่สามารถลบวันหยุดได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def delete_leave_record_on_date(emp_id, leave_date_obj):
    """(ใหม่) ลบข้อมูลการลาในวันที่ระบุ"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "DELETE FROM employee_leave_records WHERE emp_id = %s AND leave_date = %s",
                (str(emp_id), leave_date_obj)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Delete Leave)", f"ไม่สามารถลบข้อมูลการลาได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def delete_scan_logs_on_date(emp_id, scan_date_obj):
    """(ใหม่) ลบข้อมูลสแกนนิ้ว "ทั้งวัน" ของวันที่ระบุ"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "DELETE FROM time_attendance_logs WHERE emp_id = %s AND scan_timestamp::date = %s",
                (str(emp_id), scan_date_obj)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error (Delete Scans)", f"ไม่สามารถลบข้อมูลสแกนได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def add_manual_scan_log(emp_id, scan_timestamp_obj):
    """(ใหม่) เพิ่ม Log สแกนนิ้ว (รายการเดียว)"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO time_attendance_logs (emp_id, scan_timestamp)
                VALUES (%s, %s)
                ON CONFLICT DO NOTHING
                """,
                (str(emp_id), scan_timestamp_obj)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
    if base_salary > 15000:
        base_salary = 15000.0
    deduction = round(base_salary * 0.05) 
    return float(deduction)

def process_attendance_summary(start_date, end_date):
    """
    (ฉบับแก้ไข V14.0 - บันทึก OT ลง Database อัตโนมัติ เฉพาะคลังสินค้า)
    """
    conn = get_db_connection()
    if not conn: return []
    
    summary_report = []

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            
            # 1. ดึงข้อมูลพนักงาน
            cursor.execute("""
                SELECT emp_id, fname, lname, work_location, department, position, emp_type
                FROM employees
                WHERE status IS NULL 
                   OR status NOT IN ('พ้นสภาพพนักงาน', 'ลาออก')
                   OR (termination_date >= %s)
                ORDER BY emp_id ASC
            """, (start_date,))
            employees = [dict(row) for row in cursor.fetchall()]
            
            # 2. ดึงข้อมูลวันหยุด / ลา / สแกนนิ้ว (เหมือนเดิม)
            cursor.execute("SELECT holiday_date, description FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holiday_dict = {row['holiday_date']: row['description'] for row in cursor.fetchall()}
            
            cursor.execute("SELECT emp_id, leave_date, leave_type, num_days FROM employee_leave_records WHERE leave_date BETWEEN %s AND %s", (start_date, end_date))
            leaves = {}
            for row in cursor.fetchall():
                leaves[(row['emp_id'], row['leave_date'])] = row

            cursor.execute("""
                SELECT emp_id, scan_timestamp FROM time_attendance_logs 
                WHERE DATE(scan_timestamp) BETWEEN %s AND %s ORDER BY scan_timestamp ASC
            """, (start_date, end_date))
            logs_map = {}
            for row in cursor.fetchall():
                eid = row['emp_id']
                dt = row['scan_timestamp'].date()
                if (eid, dt) not in logs_map: logs_map[(eid, dt)] = []
                logs_map[(eid, dt)].append(row['scan_timestamp'])

            # 3. กฎการเข้างาน
            WORK_RULES = {
                "สำนักงานใหญ่": { "standard_in": time(9,0), "standard_out": time(18,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
                "คลังสินค้า": { "standard_in": time(8,0), "standard_out": time(17,0), "tier_1_cutoff": time(8,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
                "default": { "standard_in": time(9,0), "standard_out": time(18,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 }
            }

            all_dates = pd.date_range(start_date, end_date).date
            
            for emp in employees:
                emp_id = emp['emp_id']
                emp_name = f"{emp['fname']} {emp['lname']}"
                loc = emp.get('work_location', 'default') or 'default'
                if loc not in WORK_RULES: loc = 'default'
                rule = WORK_RULES[loc]
                
                # เช็คว่าเป็นคนคลังหรือไม่ (เพื่อคิด OT)
                is_warehouse = (loc == "คลังสินค้า")

                total_late_mins = 0
                total_absent_days = 0.0
                daily_details = []
                
                for curr_date in all_dates:
                    status = "ปกติ"
                    late_mins = 0
                    ot_minutes = 0
                    ot_hours_to_save = 0.0 # ตัวแปรสำหรับบันทึก
                    
                    scan_in_str = "-"
                    scan_out_str = "-"
                    
                    day_logs = logs_map.get((emp_id, curr_date), [])
                    leave_info = leaves.get((emp_id, curr_date))
                    is_holiday = curr_date in holiday_dict
                    is_sunday = (curr_date.weekday() == 6)

                    if day_logs:
                        scan_in = min(day_logs).time()
                        scan_out = max(day_logs).time()
                        scan_in_str = scan_in.strftime("%H:%M")
                        if len(day_logs) > 1:
                            scan_out_str = scan_out.strftime("%H:%M")
                            
                            # --- คำนวณ OT (เฉพาะคลัง + มีสแกนออก + ไม่ใช่วันลาเต็มวัน) ---
                            if is_warehouse and not leave_info:
                                if scan_out > rule['standard_out']:
                                    dummy = datetime.today()
                                    t_out = datetime.combine(dummy, scan_out)
                                    t_std = datetime.combine(dummy, rule['standard_out'])
                                    diff = t_out - t_std
                                    ot_minutes = int(diff.total_seconds() / 60)
                                    
                                    # กฎ: ต้องทำเกิน 30 นาทีถึงจะเริ่มนับ (แล้วแต่นโยบาย)
                                    # สมมติ: คิดตามจริงทุกนาที
                                    if ot_minutes > 0:
                                        ot_hours_to_save = round(ot_minutes / 60.0, 2)
                    else:
                        scan_in = None

                    # --- บันทึก OT ลง Database ทันที (Upsert) ---
                    # (จะบันทึกทับลงไปในตารางรายวันเลย เพื่อให้ Payroll ดึงไปใช้)
                    if ot_hours_to_save > 0:
                        cursor.execute("""
                            INSERT INTO employee_daily_records (emp_id, work_date, ot_hours, status)
                            VALUES (%s, %s, %s, 'ทำงาน')
                            ON CONFLICT (emp_id, work_date) DO UPDATE SET
                            ot_hours = EXCLUDED.ot_hours;
                        """, (emp_id, curr_date, ot_hours_to_save))
                    # -------------------------------------------

                    # Logic สถานะ (ย่อ)
                    if leave_info:
                        status = f"ลา {leave_info['leave_type']}"
                    elif is_holiday or is_sunday:
                        status = "วันหยุด"
                        if scan_in: status += " (มาทำ)" # มาทำวันหยุด อาจได้ OT (ยังไม่ได้ทำส่วนคูณ 2 แรง)
                    elif not scan_in:
                        status = "ขาดงาน"
                        total_absent_days += 1.0
                    else:
                        status = "ทำงาน"
                        if scan_in > rule['tier_1_cutoff']: 
                            late_mins = rule['penalty_2_mins']
                            status = "สาย (>30)"
                        elif scan_in > rule['standard_in']: 
                            late_mins = rule['penalty_1_mins']
                            status = "สาย (<30)"
                        total_late_mins += late_mins

                    daily_details.append({
                        "date": curr_date.strftime("%d/%m/%Y"),
                        "status": status,
                        "scan_in": scan_in_str,
                        "scan_out": scan_out_str,
                        "penalty_mins": late_mins,
                        "ot_hrs": ot_hours_to_save # ส่งไปโชว์ในตารางหน้าจอ
                    })

                summary_report.append({
                    "emp_id": emp_id,
                    "name": emp_name,
                    "emp_type": emp.get('emp_type', ''),
                    "department": emp.get('department', '-'),
                    "position": emp.get('position', '-'),
                    "total_late_minutes": total_late_mins,
                    "total_late_hours": total_late_mins / 60.0,
                    "absent_days": total_absent_days,
                    "details": daily_details
                })
            
            conn.commit() # ยืนยันการบันทึก OT

    except Exception as e:
        print(f"Error processing: {e}")
        return []
    finally:
        if conn: conn.close()
        
    return summary_report

def calculate_payroll_for_employee(emp_id, start_date, end_date, user_inputs=None):
    """
    (สมองหลัก Payroll V21.1 - Debug Mode)
    เพิ่ม Print เพื่อตรวจสอบสาเหตุที่ยอดหักเป็น 0
    """
    if user_inputs is None: user_inputs = {}

    # 1. กฎการเข้างาน
    WORK_RULES = {
        "สำนักงานใหญ่": { "standard_in": time(9,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
        "คลังสินค้า": { "standard_in": time(8,0), "tier_1_cutoff": time(8,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
        "default": { "standard_in": time(9,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 }
    }

    result = {
        "emp_id": emp_id, "base_salary": 0.0, "position_allowance": 0.0, "ot": 0.0, "commission": 0.0, "bonus": 0.0, "other_income": 0.0, "driving_allowance": 0.0, "total_income": 0.0, "sso": 0.0, "tax": 0.0, "provident_fund": 0.0, "loan": 0.0, "late_deduct": 0.0, "other_deduct": 0.0, "total_deduct": 0.0, "net_salary": 0.0
    }

    conn = get_db_connection()
    if not conn: return result

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            emp_info = load_single_employee(emp_id)
            if not emp_info or not emp_info.get("salary"): return result

            salary_from_db = float(emp_info.get("salary", 0.0))
            emp_type = emp_info.get("emp_type", "")
            
            # (Debug) เช็คเงินเดือน
            print(f"\n--- เริ่มคำนวณ {emp_id} ---")
            print(f"   > เงินเดือน: {salary_from_db}, ประเภท: {emp_type}")

            # ดึงค่าต่างๆ (ย่อ)
            cursor.execute("SELECT position_allowance FROM salary_history WHERE emp_id = %s ORDER BY history_id DESC LIMIT 1", (emp_id,))
            pos_row = cursor.fetchone()
            result["position_allowance"] = float(pos_row[0]) if pos_row and pos_row[0] else 0.0
            
            manual_ot = float(user_inputs.get('ot', 0))
            result["commission"] = float(user_inputs.get('commission', 0))
            result["bonus"] = float(user_inputs.get('bonus', 0))
            result["other_income"] = float(user_inputs.get('other_income', 0))
            result["tax"] = float(user_inputs.get('tax', 0))
            result["provident_fund"] = float(user_inputs.get('provident_fund', 0))
            result["loan"] = float(user_inputs.get('loan', 0))
            result["other_deduct"] = float(user_inputs.get('other_deduct', 0))

            work_location = emp_info.get('work_location', "default") or "default"
            if work_location not in WORK_RULES: work_location = "default"
            rules = WORK_RULES[work_location]
            
            # ดึงข้อมูลประกอบ
            cursor.execute("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holiday_dates_set = {row['holiday_date'] for row in cursor.fetchall()}
            
            cursor.execute("SELECT leave_date, num_days, leave_type FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            leave_records_dict = {}
            for row in cursor.fetchall(): leave_records_dict[row['leave_date']] = dict(row)

            cursor.execute("SELECT late_date, minutes_late FROM employee_late_records WHERE emp_id = %s AND late_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            manual_late_dict = {}
            for row in cursor.fetchall(): manual_late_dict[row['late_date']] = row['minutes_late']

            cursor.execute("SELECT scan_timestamp FROM time_attendance_logs WHERE emp_id = %s AND scan_timestamp::date BETWEEN %s AND %s ORDER BY scan_timestamp", (emp_id, start_date, end_date))
            scans = cursor.fetchall()
            scans_by_date = {}
            for scan in scans:
                d = scan['scan_timestamp'].date()
                if d not in scans_by_date: scans_by_date[d] = []
                scans_by_date[d].append(scan['scan_timestamp'])

            cursor.execute("SELECT work_date, total_amount, ot_hours FROM employee_daily_records WHERE emp_id = %s AND work_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            daily_records_dict = {}
            for row in cursor.fetchall(): daily_records_dict[row['work_date']] = dict(row)

            # --- เริ่ม Loop ---
            total_penalty_minutes = 0.0
            total_absent_days = 0.0
            actual_worked_days = 0.0
            auto_driving_allowance = 0.0
            total_ot_money = 0.0
            
            all_dates = pd.date_range(start_date, end_date, freq='D')

            for current_date in all_dates:
                c_date = current_date.date()
                is_holiday = (current_date.weekday() == 6 or c_date in holiday_dates_set)
                
                daily_rec = daily_records_dict.get(c_date)
                leave_info = leave_records_dict.get(c_date)
                scans_today = scans_by_date.get(c_date)
                manual_late_mins = manual_late_dict.get(c_date, 0)
                
                # A. ค่าเที่ยว
                if daily_rec: auto_driving_allowance += float(daily_rec.get('total_amount', 0) or 0)

                # B. OT
                ot_hrs = float(daily_rec.get('ot_hours', 0) or 0) if daily_rec else 0
                if ot_hrs > 0:
                    base_calc = salary_from_db / 30.0 if "รายวัน" not in emp_type else salary_from_db
                    hourly_rate = base_calc / 8.0
                    total_ot_money += (ot_hrs * hourly_rate * 1.5)

                # C. เช็คสถานะ
                is_present = False
                if scans_today: is_present = True
                if daily_rec and (daily_rec.get('total_amount',0) > 0 or daily_rec.get('ot_hours',0) > 0): is_present = True
                if manual_late_mins > 0: is_present = True

                is_leave = False
                if leave_info: is_leave = True

                if not is_holiday and not is_leave:
                    if not is_present:
                        # ขาดงาน
                        total_absent_days += 1.0
                        print(f"   - {c_date}: ขาดงาน (Absent)") # Debug
                    else:
                        actual_worked_days += 1.0
                        
                        # เช็คสาย
                        daily_late = 0
                        if scans_today:
                            t_in = min(scans_today).time()
                            if t_in > rules["tier_1_cutoff"]:
                                daily_late = rules["penalty_2_mins"]
                                print(f"   - {c_date}: สาย Tier 2 ({daily_late} นาที) - เข้า {t_in}") # Debug
                            elif t_in > rules["standard_in"]:
                                daily_late = rules["penalty_1_mins"]
                                print(f"   - {c_date}: สาย Tier 1 ({daily_late} นาที) - เข้า {t_in}") # Debug
                        
                        if manual_late_mins > 0:
                            daily_late += manual_late_mins
                            print(f"   - {c_date}: สาย Manual ({manual_late_mins} นาที)") # Debug
                            
                        total_penalty_minutes += daily_late
                elif is_holiday and is_present:
                    actual_worked_days += 1.0

            # --- สรุปยอด ---
            result["driving_allowance"] = auto_driving_allowance
            result["ot"] = manual_ot + total_ot_money
            
            penalty_hours = total_penalty_minutes / 60.0
            deduct_amount = 0.0
            
            if salary_from_db > 0:
                if "รายวัน" in emp_type or "daily" in str(emp_type).lower():
                    result["base_salary"] = salary_from_db * actual_worked_days
                    hourly_rate = salary_from_db / 8.0
                    deduct_amount = penalty_hours * hourly_rate
                    print(f"   > คำนวณรายวัน: หักสาย {deduct_amount} บาท ({penalty_hours} ชม.)") # Debug
                else:
                    result["base_salary"] = salary_from_db
                    daily_rate = salary_from_db / 30.0
                    hourly_rate = daily_rate / 8.0
                    deduct_absent = total_absent_days * daily_rate
                    deduct_late = penalty_hours * hourly_rate
                    deduct_amount = deduct_absent + deduct_late
                    print(f"   > คำนวณรายเดือน: หักขาด {deduct_absent} + หักสาย {deduct_late} = {deduct_amount}") # Debug

            result["late_deduct"] = deduct_amount
            
            # (คำนวณ Net เหมือนเดิม...)
            result["sso"] = round(result["base_salary"] * 0.05)
            if result["sso"] > 750: result["sso"] = 750
            if result["sso"] < 83: result["sso"] = 83

            result["total_income"] = (result["base_salary"] + result["position_allowance"] + result["ot"] + result["commission"] + result["bonus"] + result["other_income"] + result["driving_allowance"])
            result["total_deduct"] = (result["sso"] + result["tax"] + result["provident_fund"] + result["late_deduct"] + result["loan"] + result["other_deduct"])
            result["net_salary"] = result["total_income"] - result["total_deduct"]

    except Exception as e:
        print(f"Payroll Error: {e}")
        return None
    finally:
        if conn: conn.close()
        
    return result

def get_all_users():
    """ดึงรายชื่อผู้ใช้งานทั้งหมด"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("SELECT user_id, username, role FROM users ORDER BY user_id")
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        messagebox.showerror("DB Error", f"ไม่สามารถดึงข้อมูลผู้ใช้ได้:\n{e}")
        return []
    finally:
        if conn: conn.close()

def add_new_user(username, password, role):
    """เพิ่มผู้ใช้งานใหม่ (Password จะถูก Hash)"""
    conn = get_db_connection()
    if not conn: return False
    
    password_hash = hash_password(password) # ใช้ฟังก์ชัน hash เดิมที่มีอยู่แล้ว
    
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO users (username, password_hash, role)
                VALUES (%s, %s, %s)
                """,
                (username, password_hash, role)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        if "unique constraint" in str(e).lower():
             messagebox.showwarning("ซ้ำ", f"ชื่อผู้ใช้ '{username}' มีอยู่ในระบบแล้ว")
        else:
             messagebox.showerror("DB Error", f"ไม่สามารถเพิ่มผู้ใช้ได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def delete_user(user_id):
    """ลบผู้ใช้งาน"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM users WHERE user_id = %s", (user_id,))
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error", f"ไม่สามารถลบผู้ใช้ได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def update_user_password(user_id, new_password):
    """เปลี่ยนรหัสผ่าน (Reset Password)"""
    conn = get_db_connection()
    if not conn: return False
    
    new_hash = hash_password(new_password)
    
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "UPDATE users SET password_hash = %s WHERE user_id = %s",
                (new_hash, user_id)
            )
            conn.commit()
            return True
    except Exception as e:
        conn.rollback()
        messagebox.showerror("DB Error", f"ไม่สามารถเปลี่ยนรหัสผ่านได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def get_dashboard_stats():
    """ดึงข้อมูลสรุปสำหรับ Dashboard"""
    conn = get_db_connection()
    if not conn: return {}

    
    stats = {
        "total_employees": 0,
        "on_leave_today": 0,
        "late_today": 0,
        "probation_upcoming": [],
        "dept_counts": []
    }
    
    try:
        today = datetime.now().date()
        
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. นับจำนวนพนักงานทั้งหมด (ที่ยังไม่ออก)
            cursor.execute("SELECT COUNT(*) FROM employees WHERE status NOT IN ('พ้นสภาพพนักงาน', 'ลาออก') OR status IS NULL")
            stats["total_employees"] = cursor.fetchone()[0]
            
            # 2. นับคนที่ "ลา" วันนี้
            cursor.execute("SELECT COUNT(*) FROM employee_leave_records WHERE leave_date = %s", (today,))
            stats["on_leave_today"] = cursor.fetchone()[0]
            
            # 3. นับคนที่ "สาย" วันนี้
            cursor.execute("SELECT COUNT(*) FROM employee_late_records WHERE late_date = %s", (today,))
            stats["late_today"] = cursor.fetchone()[0]
            
            # 4. กราฟวงกลม: นับจำนวนตามแผนก
            cursor.execute("""
                SELECT COALESCE(department, 'ไม่ระบุ') as dept, COUNT(*) as count 
                FROM employees 
                WHERE status NOT IN ('พ้นสภาพพนักงาน', 'ลาออก') OR status IS NULL
                GROUP BY dept
            """)
            stats["dept_counts"] = [dict(row) for row in cursor.fetchall()]
            
            # 5. แจ้งเตือน: ใกล้ผ่านโปร (ภายใน 30 วันข้างหน้า)
            # (ใช้ Logic วันที่แบบ DATE ที่เราเพิ่งอัปเกรดมา!)
            next_30_days = today + timedelta(days=30)
            cursor.execute("""
                SELECT emp_id, fname, lname, probation_end_date, department
                FROM employees 
                WHERE probation_end_date BETWEEN %s AND %s
                AND status IN ('ระหว่างทดลองงาน')
                ORDER BY probation_end_date ASC
            """, (today, next_30_days))
            stats["probation_upcoming"] = [dict(row) for row in cursor.fetchall()]
            
    except Exception as e:
        print(f"Dashboard DB Error: {e}")
    finally:
        if conn: conn.close()
        
    return stats

def get_asmart_connection():
    """สร้าง Connection ไปหาฐานข้อมูล A+ Smart (เครื่อง 192.168.1.60)"""
    try:
        return psycopg2.connect(
            host="192.168.1.51",      # IP ของ A+ Smart
            dbname="aplus_com_test",
            user="app_user",
            password="cailfornia123",
            port="5432"
        )
    except Exception as e:
        print(f"A+ Smart Connection Error: {e}")
        return None

def get_commission_from_asmart(emp_id, target_month, target_year):
    """
    ดึงยอดคอมมิชชั่นรวมของพนักงาน จาก A+ Smart (แก้ไขชื่อคอลัมน์แล้ว)
    """
    conn = get_asmart_connection()
    if not conn: return 0.0
    
    total_comm = 0.0
    try:
        with conn.cursor() as cursor:
            # (ใช้ชื่อคอลัมน์จริงจากฐานข้อมูล A+ Smart)
            sql = """
                SELECT SUM(final_commission) 
                FROM commissions 
                WHERE sale_key = %s 
                  AND commission_month = %s 
                  AND commission_year = %s
            """
            
            # sale_key คือรหัสพนักงาน (เช่น 'AP082')
            # target_month, target_year คือเดือนปีที่ต้องการดึง
            cursor.execute(sql, (str(emp_id), target_month, target_year))
            result = cursor.fetchone()
            
            if result and result[0]:
                total_comm = float(result[0])
                
    except Exception as e:
        print(f"Error getting commission for {emp_id}: {e}")
    finally:
        conn.close()
        
    return total_comm

def add_driving_record(emp_id, drive_date, car_type, trips):
    """บันทึกค่าเที่ยวรถ (กะบะ 50 / เฮี๊ยบ 100)"""
    RATES = {"กระบะ": 50.0, "เฮี๊ยบ": 100.0}
    rate = RATES.get(car_type, 0.0)
    total_amount = rate * trips
    
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO employee_driving_records (emp_id, drive_date, car_type, trips_count, total_amount)
                VALUES (%s, %s, %s, %s, %s)
            """, (emp_id, drive_date, car_type, trips, total_amount))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error adding driving record: {e}")
        return False
    finally:
        if conn: conn.close()

def get_total_driving_allowance(emp_id, start_date, end_date):
    """คำนวณเงินค่าเที่ยวรวมในช่วงเวลา (ดึงจากตารางรายละเอียด)"""
    conn = get_db_connection()
    if not conn: return 0.0
    try:
        with conn.cursor() as cursor:
            # ดึงผลรวมของ trip_cost จากตารางรายละเอียด
            cursor.execute("""
                SELECT SUM(trip_cost) 
                FROM employee_driving_details
                WHERE emp_id = %s AND work_date BETWEEN %s AND %s
            """, (emp_id, start_date, end_date))
            
            result = cursor.fetchone()
            return float(result[0]) if result and result[0] else 0.0
    except Exception as e:
        print(f"Error calculating total allowance: {e}")
        return 0.0
    finally:
        if conn: conn.close()

def get_daily_records(emp_id, month, year):
    """ดึงข้อมูลรายวัน (เพิ่ม trip_pickup, trip_crane)"""
    conn = get_db_connection()
    if not conn: return {}
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            import calendar
            last_day = calendar.monthrange(year, month)[1]
            start_date = f"{year}-{month:02d}-01"
            end_date = f"{year}-{month:02d}-{last_day}"
            
            # (!!! อัปเดต SQL !!!)
            cursor.execute("""
                SELECT work_date, status, ot_hours, trip_pickup, trip_crane
                FROM employee_daily_records
                WHERE emp_id = %s AND work_date BETWEEN %s AND %s
            """, (emp_id, start_date, end_date))
            
            records = {}
            for row in cursor.fetchall():
                records[row['work_date']] = dict(row)
            return records
    except Exception as e:
        print(f"Error getting daily records: {e}")
        return {}
    finally:
        conn.close()

def save_daily_record(emp_id, work_date, status, ot_hours, trip_pickup, trip_crane):
    """บันทึกข้อมูลรายวัน (แยกกระบะ/เฮี๊ยบ)"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # (!!! อัปเดต SQL !!!)
            cursor.execute("""
                INSERT INTO employee_daily_records 
                (emp_id, work_date, status, ot_hours, trip_pickup, trip_crane)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                    status = EXCLUDED.status,
                    ot_hours = EXCLUDED.ot_hours,
                    trip_pickup = EXCLUDED.trip_pickup,
                    trip_crane = EXCLUDED.trip_crane;
            """, (emp_id, work_date, status, ot_hours, trip_pickup, trip_crane))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving daily record: {e}")
        return False
    finally:
        conn.close()

def add_email_request(emp_id, m, y, pdf_path, email, username):
    """HR สร้างคำขอส่งอีเมล"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # ลบคำขอเก่าของเดือนนี้ทิ้งก่อน (ถ้ามี) จะได้ไม่ซ้ำ
            cursor.execute("""
                DELETE FROM email_queue 
                WHERE emp_id=%s AND period_month=%s AND period_year=%s AND status='pending'
            """, (emp_id, m, y))
            
            cursor.execute("""
                INSERT INTO email_queue (emp_id, period_month, period_year, pdf_path, receiver_email, requested_by)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (emp_id, m, y, pdf_path, email, username))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error adding email request: {e}")
        return False
    finally:
        conn.close()

def get_pending_emails():
    """ดึงรายการรออนุมัติทั้งหมด"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT q.*, e.fname, e.lname 
                FROM email_queue q
                JOIN employees e ON q.emp_id = e.emp_id
                WHERE q.status = 'pending'
                ORDER BY q.created_at DESC
            """)
            return [dict(row) for row in cursor.fetchall()]
    except Exception:
        return []
    finally:
        conn.close()

def update_email_status(queue_id, new_status):
    """เปลี่ยนสถานะ (sent/rejected)"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("UPDATE email_queue SET status=%s WHERE queue_id=%s", (new_status, queue_id))
            conn.commit()
            return True
    except Exception:
        return False
    finally:
        conn.close()

def get_driving_details(emp_id, work_date):
    """ดึงรายการเที่ยวรถทั้งหมดของวันนั้น"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT * FROM employee_driving_details
                WHERE emp_id = %s AND work_date = %s
                ORDER BY detail_id ASC
            """, (emp_id, work_date))
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error getting driving details: {e}")
        return []
    finally:
        conn.close()

def save_driving_details_list(emp_id, work_date, details_list):
    """
    บันทึกรายการเที่ยวรถ (อัปเดต V3: รองรับ is_free)
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # 1. ลบข้อมูลเก่าของวันนั้นทิ้งก่อน
            cursor.execute("DELETE FROM employee_driving_details WHERE emp_id = %s AND work_date = %s", (emp_id, work_date))
            
            # 2. ใส่ข้อมูลใหม่เข้าไปทีละแถว
            for item in details_list:
                # แปลง is_free เป็น Boolean (True/False)
                is_free_val = bool(item.get('is_free', False))
                
                cursor.execute("""
                    INSERT INTO employee_driving_details 
                    (emp_id, work_date, car_type, license_plate, driver_name, delivery_date, trip_cost, ref_doc_type, ref_doc_id, is_free)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    emp_id, work_date, 
                    item.get('car_type'), item.get('license'), item.get('driver'), 
                    item.get('send_date'), item.get('cost'),
                    item.get('doc_type', ''), item.get('doc_id', ''),
                    is_free_val # <--- บันทึกสถานะฟรี
                ))
            
            # 3. อัปเดตยอดรวมลงตารางหลัก
            total_money = sum(x['cost'] for x in details_list)
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, total_amount, status)
                VALUES (%s, %s, %s, 'ทำงาน')
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                total_amount = EXCLUDED.total_amount;
            """, (emp_id, work_date, total_money))
            
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving driving details: {e}")
        return False
    finally:
        conn.close()

def save_monthly_payroll(emp_id, month, year, pay_date, data):
    """
    บันทึกงวดเงินเดือนลง Database (ถ้ามีแล้วให้ Update ทับ)
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO payroll_records (
                    emp_id, period_month, period_year, payment_date,
                    base_salary, position_allowance, ot_pay, commission, bonus,
                    driving_allowance, other_income, total_income,
                    sso_deduct, tax_deduct, provident_fund, loan_deduct,
                    late_deduct, other_deduct, total_deduct, net_salary
                ) VALUES (
                    %s, %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s, %s, %s
                )
                ON CONFLICT (emp_id, period_month, period_year) DO UPDATE SET
                    payment_date = EXCLUDED.payment_date,
                    base_salary = EXCLUDED.base_salary,
                    position_allowance = EXCLUDED.position_allowance,
                    ot_pay = EXCLUDED.ot_pay,
                    commission = EXCLUDED.commission,
                    bonus = EXCLUDED.bonus,
                    driving_allowance = EXCLUDED.driving_allowance,
                    other_income = EXCLUDED.other_income,
                    total_income = EXCLUDED.total_income,
                    sso_deduct = EXCLUDED.sso_deduct,
                    tax_deduct = EXCLUDED.tax_deduct,
                    provident_fund = EXCLUDED.provident_fund,
                    loan_deduct = EXCLUDED.loan_deduct,
                    late_deduct = EXCLUDED.late_deduct,
                    other_deduct = EXCLUDED.other_deduct,
                    total_deduct = EXCLUDED.total_deduct,
                    net_salary = EXCLUDED.net_salary;
            """, (
                emp_id, month, year, pay_date,
                data.get('base_salary', 0), data.get('position_allowance', 0),
                data.get('ot', 0), data.get('commission', 0), data.get('bonus', 0),
                data.get('driving_allowance', 0), data.get('other_income', 0), data.get('total_income', 0),
                data.get('sso', 0), data.get('tax', 0), data.get('provident_fund', 0),
                data.get('loan', 0), data.get('late_deduct', 0), data.get('other_deduct', 0),
                data.get('total_deduct', 0), data.get('net_salary', 0)
            ))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving payroll record: {e}")
        return False
    finally:
        conn.close()

def get_monthly_payroll_records(month, year):
    """ดึงข้อมูลเงินเดือนที่บันทึกไว้ (History) ตามเดือน/ปี"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # ดึงข้อมูลจาก payroll_records + ชื่อพนักงาน
            cursor.execute("""
                SELECT pr.*, e.fname, e.lname, e.id_card, e.position, e.department,
                       e.is_sales, e.sale_type, e.commission_plan
                FROM payroll_records pr
                JOIN employees e ON pr.emp_id = e.emp_id
                WHERE pr.period_month = %s AND pr.period_year = %s
                ORDER BY pr.emp_id ASC
            """, (month, year))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error fetching payroll history: {e}")
        return []
    finally:
        conn.close()

def get_annual_pnd1k_data(year_ce):
    """
    ดึงข้อมูลทำ ภ.ง.ด. 1ก (สรุปรายได้/ภาษี ทั้งปี ของพนักงานแต่ละคน)
    year_ce: ปี ค.ศ. (เช่น 2025)
    """
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT 
                    e.emp_id, e.fname, e.lname, e.id_card, e.address, e.position,
                    SUM(pr.total_income) as annual_income,
                    SUM(pr.tax_deduct) as annual_tax,
                    SUM(pr.sso_deduct) as annual_sso,
                    SUM(pr.provident_fund) as annual_fund
                FROM payroll_records pr
                JOIN employees e ON pr.emp_id = e.emp_id
                WHERE pr.period_year = %s
                GROUP BY e.emp_id, e.fname, e.lname, e.id_card, e.address, e.position
                ORDER BY e.emp_id
            """, (year_ce,))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error fetching PND1K data: {e}")
        return []
    finally:
        conn.close()