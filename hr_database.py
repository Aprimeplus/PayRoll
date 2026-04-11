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
            host="Server-APrime",
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
    """สร้างตารางทั้งหมดหากยังไม่มี (ฉบับแก้ไขสมบูรณ์ V2 - เพิ่มคอลัมน์ OT และ เวลาเข้าออก)"""
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
            
            # Update columns for employees
            try:
                cursor.execute("""
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS is_sales BOOLEAN DEFAULT FALSE,
                    ADD COLUMN IF NOT EXISTS sale_type TEXT,
                    ADD COLUMN IF NOT EXISTS commission_plan TEXT,
                    ADD COLUMN IF NOT EXISTS probation_days INTEGER DEFAULT 90,
                    ADD COLUMN IF NOT EXISTS probation_end_date DATE,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score TEXT,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score_2 TEXT,
                    ADD COLUMN IF NOT EXISTS sso_hospital TEXT,
                    ADD COLUMN IF NOT EXISTS work_location TEXT,
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
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS prefix TEXT; -- <--- เพิ่มบรรทัดนี้
                """)
            except Exception: conn.rollback()

            try:
                cursor.execute("""
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS is_sales BOOLEAN DEFAULT FALSE,
                    ADD COLUMN IF NOT EXISTS sale_type TEXT,
                    ADD COLUMN IF NOT EXISTS commission_plan TEXT,
                    ADD COLUMN IF NOT EXISTS probation_days INTEGER DEFAULT 90,
                    ADD COLUMN IF NOT EXISTS probation_end_date DATE,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score TEXT,
                    ADD COLUMN IF NOT EXISTS probation_assessment_score_2 TEXT,
                    ADD COLUMN IF NOT EXISTS sso_hospital TEXT,
                    ADD COLUMN IF NOT EXISTS work_location TEXT,
                    ADD COLUMN IF NOT EXISTS emergency_contact_name TEXT,
                    ADD COLUMN IF NOT EXISTS emergency_contact_phone TEXT,
                    ADD COLUMN IF NOT EXISTS emergency_contact_relation TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_name TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_phone TEXT,
                    ADD COLUMN IF NOT EXISTS ref_person_relation TEXT,
                    ADD COLUMN IF NOT EXISTS diligence_streak INTEGER DEFAULT 0; 
                """)
            except Exception: conn.rollback()

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    user_id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK (role IN ('hr', 'approver', 'dispatcher')) -- <--- เติมตรงนี้
                );
            """)
            

            # 2. ตาราง Email Queue
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS email_queue (
                        queue_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        period_month INTEGER,
                        period_year INTEGER,
                        pdf_path TEXT NOT NULL,
                        receiver_email TEXT,
                        status TEXT DEFAULT 'pending',
                        requested_by TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                    );
                """)
            except Exception as e:
                print(f"Warning (Create email_queue): {e}")
                conn.rollback()

            # 3. ตารางบันทึกรายวัน (แก้ไข: เพิ่มคอลัมน์ OT และ Work Time ในตัว)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_daily_records (
                    record_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    work_date DATE NOT NULL,
                    status TEXT,             
                    ot_hours REAL DEFAULT 0,
                    car_type TEXT,           
                    trip_count INTEGER DEFAULT 0,
                    trip_pickup INTEGER DEFAULT 0, -- (เพิ่ม)
                    trip_crane INTEGER DEFAULT 0,  -- (เพิ่ม)
                    total_amount REAL DEFAULT 0,   -- (เพิ่ม)
                    ot_in_time TEXT,               -- (เพิ่ม)
                    ot_out_time TEXT,              -- (เพิ่ม)
                    work_in_time TEXT,             -- (เพิ่ม)
                    work_out_time TEXT,  
                    is_ot_approved BOOLEAN DEFAULT FALSE,          -- (เพิ่ม)
                    UNIQUE(emp_id, work_date),  
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)
            try:
                cursor.execute("""
                    ALTER TABLE employee_daily_records
                    ADD COLUMN IF NOT EXISTS is_ot_approved BOOLEAN DEFAULT FALSE;
                """)
            except Exception: conn.rollback()
            
            try:
                cursor.execute("""
                    ALTER TABLE employees
                    ADD COLUMN IF NOT EXISTS company_pays_tax BOOLEAN DEFAULT FALSE; -- <--- เพิ่มบรรทัดนี้
                """)
            except Exception: conn.rollback()

            # (กันเหนียว) สั่ง ALTER อีกรอบ เผื่อตารางสร้างไปแล้วแต่ไม่มีคอลัมน์
            try:
                cursor.execute("""
                    ALTER TABLE employee_daily_records
                    ADD COLUMN IF NOT EXISTS ot_in_time TEXT,
                    ADD COLUMN IF NOT EXISTS ot_out_time TEXT,
                    ADD COLUMN IF NOT EXISTS work_in_time TEXT,
                    ADD COLUMN IF NOT EXISTS work_out_time TEXT,
                    ADD COLUMN IF NOT EXISTS trip_pickup INTEGER DEFAULT 0,
                    ADD COLUMN IF NOT EXISTS trip_crane INTEGER DEFAULT 0,
                    ADD COLUMN IF NOT EXISTS total_amount REAL DEFAULT 0,
                    ADD COLUMN IF NOT EXISTS remark TEXT;
                """)
            except Exception: conn.rollback()

            # 4. ตารางสวัสดิการ
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_welfare (
                    welfare_id SERIAL PRIMARY KEY, emp_id TEXT NOT NULL,
                    welfare_name TEXT NOT NULL, has_welfare INTEGER NOT NULL,
                    amount REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 5. ตาราง Audit Logs
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS audit_logs (
                    log_id SERIAL PRIMARY KEY,
                    action_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    actor_name TEXT,         
                    action_type TEXT,        
                    target_emp_id TEXT,      
                    target_emp_name TEXT,    
                    changed_field TEXT,      
                    old_value TEXT,          
                    new_value TEXT           
                );
            """)

            # 6. ตาราง Payroll Records (รวมยอดสุทธิ)
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
                    incentive REAL DEFAULT 0,    -- (เพิ่ม)
                    diligence REAL DEFAULT 0,    -- (เพิ่ม)
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
            
            # (กันเหนียว) Update Payroll columns
            try:
                cursor.execute("""
                    ALTER TABLE payroll_records
                    ADD COLUMN IF NOT EXISTS incentive REAL DEFAULT 0,
                    ADD COLUMN IF NOT EXISTS diligence REAL DEFAULT 0,
                    ADD COLUMN IF NOT EXISTS remark TEXT;
                """)
            except Exception: conn.rollback()

            # 7. ตารางเที่ยวรถ (Details & Records)
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
                    ref_doc_type TEXT, -- (เพิ่ม)
                    ref_doc_id TEXT,   -- (เพิ่ม)
                    is_free BOOLEAN DEFAULT FALSE,
                    is_service BOOLEAN DEFAULT FALSE,
                    service_fee REAL DEFAULT 0,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)
            
            # Update columns driving details
            try:
                cursor.execute("""
                    ALTER TABLE employee_driving_details
                    ADD COLUMN IF NOT EXISTS ref_doc_type TEXT,
                    ADD COLUMN IF NOT EXISTS ref_doc_id TEXT,
                    ADD COLUMN IF NOT EXISTS is_free BOOLEAN DEFAULT FALSE,
                    ADD COLUMN IF NOT EXISTS is_service BOOLEAN DEFAULT FALSE,
                    ADD COLUMN IF NOT EXISTS service_fee REAL DEFAULT 0;
                """)
            except Exception: conn.rollback()

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS employee_driving_records (
                    record_id SERIAL PRIMARY KEY,
                    emp_id TEXT NOT NULL,
                    drive_date DATE NOT NULL,
                    car_type TEXT NOT NULL, 
                    trips_count INTEGER DEFAULT 1,
                    total_amount REAL,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)

            # 8. ตารางประวัติเงินเดือน
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS salary_history (
                    history_id SERIAL PRIMARY KEY, emp_id TEXT NOT NULL,
                    adjustment_year TEXT, new_salary REAL,position_allowance REAL,
                    new_position TEXT, assessment_score TEXT,
                    FOREIGN KEY (emp_id) REFERENCES employees (emp_id) ON DELETE CASCADE
                );
            """)
            try:
                cursor.execute("""
                    ALTER TABLE salary_history
                    ADD COLUMN IF NOT EXISTS new_position TEXT,
                    ADD COLUMN IF NOT EXISTS assessment_score TEXT;
                """)
            except Exception: conn.rollback()

            # 9. ตารางประวัติฝึกอบรม
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

            # 10. ตารางทรัพย์สิน
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

            # 11. ตาราง Users & Pending Changes
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    user_id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK (role IN ('hr', 'approver'))
                );
            """)

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

            # 12. ตารางเอกสารแนบ
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
                
            # 13. ตารางวันหยุด
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS company_holidays (
                        holiday_id SERIAL PRIMARY KEY,
                        holiday_date DATE NOT NULL UNIQUE,
                        description TEXT
                    );
                """)
            except Exception: conn.rollback()    
            
            # 14. ตาราง Settings & Locations
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

            # 15. ตารางการลา
            try:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS employee_leave_records (
                        leave_id SERIAL PRIMARY KEY,
                        emp_id TEXT NOT NULL,
                        leave_date DATE NOT NULL,
                        leave_type TEXT NOT NULL, 
                        num_days REAL DEFAULT 1.0,
                        reason TEXT,
                        leave_start_time TIME,
                        leave_end_time TIME,
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
                
            # 16. ตารางการมาสายและใบเตือน
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

            # 17. ตาราง Log สแกนนิ้ว
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

            # 18. เพิ่ม User เริ่มต้น (Default Users)
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

def get_employee_diligence_streak(emp_id):
    """ดึงจำนวนเดือนที่ได้เบี้ยขยันต่อเนื่อง"""
    conn = get_db_connection()
    if not conn: return 0
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT diligence_streak FROM employees WHERE emp_id = %s", (str(emp_id),))
            res = cursor.fetchone()
            return int(res[0]) if res and res[0] else 0
    except Exception as e:
        print(f"Error fetching diligence streak: {e}")
        return 0
    finally:
        conn.close()

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
            "emp_id", "prefix", "fname", "nickname", "lname", "birth_date", "age", "id_card", "phone","company_pays_tax",
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
            "ref_person_name", "ref_person_phone", "ref_person_relation",
            "id_card_expiry"  # <--- ✅ [1] เพิ่มตรงนี้
        ]
        
        # สร้าง dict เริ่มต้น (ค่าจะเป็น None ถ้าชื่อไม่ตรง)
        sql_data = {col: data.get(col) for col in emp_cols} 
        sql_data['emp_id'] = data.get('id')
        
        # --- จับคู่ชื่อตัวแปรจาก UI ให้ตรงกับ DB ---
        sql_data['emergency_contact_name'] = data.get('emergency_name')
        sql_data['emergency_contact_phone'] = data.get('emergency_phone')
        sql_data['emergency_contact_relation'] = data.get('emergency_relation')
        sql_data['prefix'] = data.get('prefix', '')
        sql_data['ref_person_name'] = data.get('ref_name')
        sql_data['ref_person_phone'] = data.get('ref_phone')
        sql_data['ref_person_relation'] = data.get('ref_relation')
        # --------------------------------------------------------------

        # --- แปลงวันที่ ---
        sql_data['birth_date'] = thai_str_to_date(data.get('birth'))
        sql_data['start_date'] = thai_str_to_date(data.get('start_date'))
        sql_data['termination_date'] = thai_str_to_date(data.get('termination_date'))
        sql_data['probation_end_date'] = thai_str_to_date(data.get('probation_end_date'))
        
        sql_data['sso_start_date'] = thai_str_to_date(data.get('sso_start'))
        sql_data['sso_end_date'] = thai_str_to_date(data.get('sso_end'))
        sql_data['sso_start_action_date'] = thai_str_to_date(data.get('sso_start_action'))
        sql_data['sso_end_action_date'] = thai_str_to_date(data.get('sso_end_action'))

        sql_data['id_card_expiry'] = thai_str_to_date(data.get('id_card_expiry')) # <--- ✅ [2] เพิ่มตรงนี้

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
                "prefix","fname", "nickname", "lname", "birth_date", "age", "id_card", "phone","sale_type", "commission_plan",
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
            employee_data['id_card_expiry'] = date_to_thai_str(employee_data.get('id_card_expiry')) # <--- เพิ่มตรงนี้
            # --- (จุดที่แก้ไข: แปลง Date Object (ค.ศ.) เป็น String (พ.ศ.) เพื่อแสดงผล UI) ---
            employee_data['birth'] = date_to_thai_str(employee_data.get('birth_date'))
            employee_data['start_date'] = date_to_thai_str(employee_data.get('start_date'))
            employee_data['termination_date'] = date_to_thai_str(employee_data.get('termination_date'))
            employee_data['probation_end_date'] = date_to_thai_str(employee_data.get('probation_end_date'))
            employee_data['is_sales'] = bool(employee_data.get('is_sales'))
            employee_data['company_pays_tax'] = bool(employee_data.get('company_pays_tax', False)) # <--- เพิ่มบรรทัดนี้
            employee_data['sso_start'] = date_to_thai_str(employee_data.get('sso_start_date'))
            employee_data['sso_end'] = date_to_thai_str(employee_data.get('sso_end_date'))
            employee_data['sso_start_action'] = date_to_thai_str(employee_data.get('sso_start_action_date'))
            employee_data['sso_end_action'] = date_to_thai_str(employee_data.get('sso_end_action_date'))
            
            employee_data['exp'] = employee_data.get('work_exp')
            employee_data['health'] = employee_data.get('health_status')
            employee_data['account'] = employee_data.get('bank_account_no')
            employee_data['bank'] = employee_data.get('bank_name')
            employee_data['branch'] = employee_data.get('bank_branch')
            employee_data['account_name'] = employee_data.get('bank_account_name')
            employee_data['account_type'] = employee_data.get('bank_account_type')
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
    """
    บันทึกข้อมูลพนักงาน + พร้อมระบบ Audit Trail (จับผิดการแก้ไข)
    """
    conn = get_db_connection()
    if not conn: return False, "DB Connection Failed"
    
    emp_id = data.get("id")
    role = current_user.get("role")
    username = current_user.get("username")
    
    # 1. ดึงข้อมูลเก่ามาเก็บไว้ก่อน (Old Data)
    old_data = load_single_employee(emp_id)
    is_new_employee = not bool(old_data)
    
    try:
        with conn.cursor() as cursor:
            if is_new_employee:
                # --- กรณีพนักงานใหม่ (CREATE) ---
                success = _upsert_employee_data(cursor, data)
                if success:
                    # บันทึก Log ว่าสร้างพนักงานใหม่
                    emp_name = f"{data.get('fname')} {data.get('lname')}"
                    add_audit_log(username, 'CREATE', emp_id, emp_name, 'All Data', '-', 'New Employee Created')
                    message = "บันทึกข้อมูลพนักงานใหม่เรียบร้อยแล้ว"
                else:
                    raise Exception("Insert failed")

            elif role == 'hr':
                # HR แก้ไข -> ส่งเข้า Pending (เหมือนเดิม)
                change_data_json = json.dumps(data)
                cursor.execute("""
                    INSERT INTO pending_employee_changes (emp_id, change_data, requested_by, status)
                    VALUES (%s, %s, %s, %s)
                """, (emp_id, change_data_json, username, 'pending'))
                message = "ส่งคำขอแก้ไขข้อมูลเรียบร้อยแล้ว รอการอนุมัติ"

            elif role == 'approver' or role == 'admin': # (สมมติ admin ทำได้เลย)
                # --- กรณีแก้ไขข้อมูล (UPDATE) พร้อม Audit Trail ---
                
                # รายชื่อฟิลด์ที่อยากจับตามอง (Sensitive Fields)
                fields_to_track = {
                    'salary': "เงินเดือน",
                    'position': "ตำแหน่ง",
                    'department': "แผนก",
                    'status': "สถานะ",
                    'bank_account_no': "เลขบัญชี",
                    'is_sales': "ฝ่ายขาย",
                    'commission_plan': "แผนคอมมิชชั่น"
                }
                
                # เปรียบเทียบค่า (Compare)
                changes_detected = []
                if old_data:
                    for key, label in fields_to_track.items():
                        # ดึงค่าเก่า/ใหม่ (แปลงเป็น string เพื่อเทียบง่ายๆ)
                        val_old = str(old_data.get(key, '') or '').strip()
                        
                        # (Mapping ชื่อ key ให้ตรงกับ data ที่ส่งมา)
                        data_key = key
                        if key == 'bank_account_no': data_key = 'account' # ชื่อใน dict ไม่ตรง DB
                        
                        val_new = str(data.get(data_key, '') or '').strip()
                        
                        # ถ้าไม่เหมือนกัน -> จดไว้
                        if val_old != val_new:
                            changes_detected.append((label, val_old, val_new))

                # บันทึกข้อมูลจริง
                success = _upsert_employee_data(cursor, data)
                if not success: raise Exception("Update failed")
                
                # ถ้าบันทึกสำเร็จ -> เขียน Audit Log ทีละรายการ
                emp_name = f"{data.get('fname')} {data.get('lname')}"
                for label, v_old, v_new in changes_detected:
                    add_audit_log(username, 'UPDATE', emp_id, emp_name, label, v_old, v_new)
                
                message = "อัปเดตข้อมูลพนักงานเรียบร้อยแล้ว"
                
            else: 
                 raise Exception("Invalid user role")
            
            conn.commit()
            return True, message
            
    except Exception as e:
        conn.rollback()
        print(f"Save Error: {e}")
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

def add_employee_leave(emp_id, leave_date, leave_type, num_days, reason="", leave_start_time=None, leave_end_time=None):
    """
    (ฉบับแก้ไข V71.0 - Support Time Arguments)
    - เพิ่ม parameter: leave_start_time, leave_end_time เพื่อรองรับการส่งค่า 7 ตัวแปร
    - บันทึกเวลาลาลงฐานข้อมูลด้วย (ถ้ามี)
    """
    conn = get_db_connection()
    if not conn: return False
    
    try:
        with conn.cursor() as cursor:
            # 1. Insert/Update Leave Record (เพิ่ม column เวลา)
            cursor.execute("""
                INSERT INTO employee_leave_records 
                (emp_id, leave_date, leave_type, num_days, reason, leave_start_time, leave_end_time)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (emp_id, leave_date) 
                DO UPDATE SET 
                    leave_type = EXCLUDED.leave_type,
                    num_days = EXCLUDED.num_days,
                    reason = EXCLUDED.reason,
                    leave_start_time = EXCLUDED.leave_start_time,
                    leave_end_time = EXCLUDED.leave_end_time;
            """, (emp_id, leave_date, leave_type, num_days, reason, leave_start_time, leave_end_time))
            
            # 2. Update Daily Record Status
            # (ตรวจสอบว่า num_days เป็นตัวเลขหรือไม่ ก่อนจัดรูปแบบ)
            try:
                nd = float(num_days)
            except:
                nd = 1.0

            status_text = f"ลา {leave_type}"
            if nd < 1.0:
                status_text += f" ({nd} วัน)"
            
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, status)
                VALUES (%s, %s, %s)
                ON CONFLICT (emp_id, work_date) 
                DO UPDATE SET status = EXCLUDED.status;
            """, (emp_id, leave_date, status_text))

            # สำคัญ: ไม่ลบ scan logs ทิ้ง (ตามที่คุยกันไว้)
            
            conn.commit()
            print(f"Added leave for {emp_id} on {leave_date}: {leave_type} (Time: {leave_start_time}-{leave_end_time})")
            return True
            
    except Exception as e:
        print(f"Error adding leave: {e}")
        import traceback; traceback.print_exc()
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

def get_work_rules(work_location):
    """
    กำหนดกฎการเข้างานแบบ Centralized
    - คลังสินค้า: เข้า 08:00
    - สำนักงาน/อื่นๆ: เข้า 09:00
    """
    loc = str(work_location).strip() if work_location else ""
    
    # กฎสำหรับคลังสินค้า
    if "คลังสินค้า" in loc:
        return {
            "standard_in": time(8, 0),
            "standard_out": time(17, 0),
            "tier_1_cutoff": time(8, 30), # สายเกินนี้โดนหนัก
            "penalty_1_mins": 60,  # หัก 1 ชม.
            "penalty_2_mins": 120  # หัก 2 ชม.
        }
    
    # กฎ Default (สำนักงานใหญ่)
    return {
        "standard_in": time(9, 0),
        "standard_out": time(18, 0), # สมมติว่าเลิก 18.00 หรือ 17.00 ตามจริง
        "tier_1_cutoff": time(9, 30),
        "penalty_1_mins": 60,
        "penalty_2_mins": 120
    }



def get_auto_diligence_reward(emp_id, current_month, current_year):
    conn = get_db_connection()
    if not conn: return 0

    try:
        consecutive_good_months = 0
        print(f"\n--- DEBUG: Checking Diligence for {emp_id} (Month {current_month}/{current_year}) ---")
        
        for i in range(1, 13): 
            target_month = current_month - i
            target_year = current_year
            while target_month <= 0:
                target_month += 12
                target_year -= 1
            
            with conn.cursor() as cursor:
                import calendar
                last_day = calendar.monthrange(target_year, target_month)[1]
                start_date = datetime(target_year, target_month, 1).date()
                end_date = datetime(target_year, target_month, last_day).date()
                
                # 1. เช็ควันทำงาน
                cursor.execute("""
                    SELECT COUNT(*) FROM employee_daily_records 
                    WHERE emp_id = %s AND work_date BETWEEN %s AND %s
                """, (emp_id, start_date, end_date))
                work_days_count = cursor.fetchone()[0]

                # 2. เช็คประวัติเสีย
                cursor.execute("SELECT COUNT(*) FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
                leave_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM employee_late_records WHERE emp_id = %s AND late_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
                late_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM employee_daily_records WHERE emp_id = %s AND work_date BETWEEN %s AND %s AND (status LIKE '%%สาย%%' OR status LIKE '%%ขาด%%')", (emp_id, start_date, end_date))
                daily_issue = cursor.fetchone()[0]
                
                # 3. [เพิ่มใหม่] เช็คเวลา 08:00
                cursor.execute("""
                    SELECT COUNT(*) FROM (
                        SELECT DATE(scan_timestamp) as d, MIN(scan_timestamp)::time as first_scan
                        FROM time_attendance_logs
                        WHERE emp_id = %s 
                          AND DATE(scan_timestamp) BETWEEN %s AND %s
                        GROUP BY d
                    ) as daily_checks
                    WHERE first_scan > '08:00:00'
                """, (emp_id, start_date, end_date))
                late_8am_count = cursor.fetchone()[0]

                # Debug Print
                print(f"   > Month {target_month}/{target_year}: Work={work_days_count}, Leave={leave_count}, Late={late_count}, DailyStatus={daily_issue}, Late>8:00={late_8am_count}")

                # เงื่อนไข: ต้องมีวันทำงาน และ ไม่มีประวัติเสีย และ ไม่เข้าเกิน 8 โมง
                if work_days_count > 0 and (leave_count == 0 and late_count == 0 and daily_issue == 0 and late_8am_count == 0):
                    consecutive_good_months += 1
                    print(f"     ✅ PASS (Streak: {consecutive_good_months})")
                else:
                    print("     ❌ STOP (Failed criteria)")
                    break
        
        print(f"--- Final Streak: {consecutive_good_months} ---")
        
        if consecutive_good_months == 0: return 300
        elif consecutive_good_months == 1: return 400
        else: return 500
            
    except Exception as e:
        print(f"Auto Diligence Error: {e}")
        return 0
    finally:
        if conn: conn.close()

def get_diligence_streak_info(emp_id, current_month, current_year):
    """
    คำนวณ Streak เบี้ยขยัน (เพิ่มกฎ: ต้องสแกนเข้างานไม่เกิน 08:00:00)
    """
    conn = get_db_connection()
    if not conn: return 0, 0.0

    try:
        consecutive_good_months = 0
        
        # วนลูปย้อนหลัง 12 เดือนเพื่อนับ Streak
        for i in range(1, 13): 
            target_month = current_month - i
            target_year = current_year
            while target_month <= 0:
                target_month += 12
                target_year -= 1
            
            with conn.cursor() as cursor:
                import calendar
                last_day = calendar.monthrange(target_year, target_month)[1]
                start_date = datetime(target_year, target_month, 1).date()
                end_date = datetime(target_year, target_month, last_day).date()
                
                # 1. เช็ควันทำงาน (ต้องมีวันทำงานอย่างน้อย 1 วัน)
                cursor.execute("""
                    SELECT COUNT(*) FROM employee_daily_records 
                    WHERE emp_id = %s AND work_date BETWEEN %s AND %s
                """, (emp_id, start_date, end_date))
                work_days = cursor.fetchone()[0]

                # 2. เช็คประวัติเสีย (ลา / สาย / ขาด)
                cursor.execute("SELECT COUNT(*) FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
                bad_1 = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM employee_late_records WHERE emp_id = %s AND late_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
                bad_2 = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM employee_daily_records WHERE emp_id = %s AND work_date BETWEEN %s AND %s AND (status LIKE '%%สาย%%' OR status LIKE '%%ขาด%%')", (emp_id, start_date, end_date))
                bad_3 = cursor.fetchone()[0]

                # 3. [เพิ่มใหม่] เช็คเวลาเข้างาน (ต้องไม่เกิน 08:00)
                # ดึงจำนวนวันที่สแกนเข้างานครั้งแรกของวัน "เกิน" 08:00
                cursor.execute("""
                    SELECT COUNT(*) FROM (
                        SELECT DATE(scan_timestamp) as d, MIN(scan_timestamp)::time as first_scan
                        FROM time_attendance_logs
                        WHERE emp_id = %s 
                          AND DATE(scan_timestamp) BETWEEN %s AND %s
                        GROUP BY d
                    ) as daily_checks
                    WHERE first_scan > '08:00:00'
                """, (emp_id, start_date, end_date))
                late_arrival_count = cursor.fetchone()[0]
                
                # รวมเงื่อนไข: มีวันทำงาน + ไม่มีประวัติเสีย + ไม่เข้าเกิน 8 โมง
                if work_days > 0 and (bad_1 + bad_2 + bad_3 == 0) and late_arrival_count == 0:
                    consecutive_good_months += 1
                else:
                    break # หยุดนับทันทีถ้าเจอเดือนที่ไม่ผ่านเกณฑ์
        
        # คำนวณเงินรางวัล (ปรับตัวเลขได้ตามต้องการ)
        reward = 0.0
        if consecutive_good_months == 0: reward = 300.0
        elif consecutive_good_months == 1: reward = 400.0
        else: reward = 500.0
            
        return consecutive_good_months, reward
            
    except Exception as e:
        print(f"Streak Info Error: {e}")
        return 0, 0.0
    finally:
        if conn: conn.close()

# (ค้นหาฟังก์ชัน calculate_payroll_for_employee แล้วนำโค้ดนี้ไปวางทับ)

def calculate_payroll_for_employee(emp_id, start_date, end_date, user_inputs=None):
    """
    (ฉบับสมบูรณ์ที่สุด V89.39 - อัปเดตกฎเบี้ยขยันเฉพาะรายวันคลังสินค้า)
    - 🛠️ [FIX] วันหยุดรายวันไม่ถูกนำมาคิดเงิน
    - 🛠️ [FIX] ยอดอื่นๆรับ ไม่เบิ้ล 4000
    - 🛠️ [FIX] เพิ่มการคำนวณค่าล่วงเวลา (OT) จากฐานชั่วโมงในระบบ x 1.5 เท่า
    - 🛠️ [NEW] จำกัดสิทธิ์เบี้ยขยันออโต้ เฉพาะพนักงาน "รายวัน" และ "คลังสินค้า" เท่านั้น
    - 🛠️ [NEW] ลบเงื่อนไขบังคับ 08:00 น. ทิ้ง ให้ยึดตามสถานะในตารางเวลาแทน
    """
    import calendar
    from datetime import date, datetime, time, timedelta
    import math 
    import re 

    if user_inputs is None: user_inputs = {}

    is_sso_exempt = False
    is_daily_calculation = False
    is_pnd3_applicable = False

    result = {
        "emp_id": emp_id, "base_salary": 0.0, "position_allowance": 0.0,
        "ot": 0.0, "commission": 0.0, "bonus": 0.0, "incentive": 0.0, "diligence": 0.0,
        "other_income": 0.0, "driving_allowance": 0.0, "total_income": 0.0,
        "sso": 0.0, "tax": 0.0, "provident_fund": 0.0, "loan": 0.0, 
        "late_deduct": 0.0, "other_deduct": 0.0, "total_deduct": 0.0,
        "net_salary": 0.0, "pnd1": 0.0, "pnd3": 0.0
    }

    conn = get_db_connection()
    if not conn: return result

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            emp_info = load_single_employee(emp_id)
            if not emp_info: return result
            salary_db = float(emp_info.get("salary", 0.0))
            emp_type = str(emp_info.get("emp_type", ""))
            work_loc = str(emp_info.get("work_location", "")) # 🛠️ ดึงสถานที่ทำงานมาเช็ค
            
            is_daily_style = ("รายวัน" in emp_type) or ("Daily" in emp_type)
            is_warehouse = "คลัง" in work_loc # 🛠️ เช็คว่าเป็นคลังสินค้าหรือไม่
            is_contractor = any(k in emp_type for k in ["จ้างเหมา", "ที่ปรึกษา", "Contract", "สัญญาจ้าง"])
            is_fixed = any(k in emp_type for k in ["รายเดือน", "รายปี", "Monthly", "Yearly", "กรรมการ"]) # 🛠️ [เพิ่มคำว่า "กรรมการ" ลงไปในวงเล็บนี้]
            
            is_daily_calculation = is_daily_style or (is_contractor and not is_fixed)
            is_pnd3_applicable = is_contractor and not is_daily_style and not is_fixed
            is_sso_exempt = is_contractor and not is_daily_style and not is_fixed

            if isinstance(start_date, datetime): start_date = start_date.date()
            if isinstance(end_date, datetime): end_date = end_date.date()

            # --- [1] จัดการเบี้ยขยัน (Manual Priority) ---
            val_diligence = user_inputs.get('diligence', 0)
            try:
                manual_diligence = float(val_diligence)
            except:
                manual_diligence = 0.0

            # ตรวจสอบว่ากำลังคิดเงินของ วิก 1 (วันที่ 1-15) หรือไม่
            is_period_1 = (start_date.day <= 15 and end_date.day <= 15)

            if manual_diligence > 0:
                result["diligence"] = manual_diligence
            elif not (is_daily_style and is_warehouse):
                # 🛠️ [RESTORED] ถ้าไม่ใช่รายวัน และ ไม่ใช่คลังสินค้า ตัดเบี้ยขยันเป็น 0 ทันที
                result["diligence"] = 0.0
            elif is_period_1:
                # ถ้าเป็นวิก 1 ตัดเบี้ยขยันออโต้เป็น 0 ทันที (รอไปจ่ายตอนวิก 2)
                result["diligence"] = 0.0
            else:
                # ถ้าเป็นวิก 2 หรือ เต็มเดือน บังคับเช็คประวัติ "ทั้งเดือน" เสมอ
                check_start = start_date.replace(day=1)
                last_day = calendar.monthrange(end_date.year, end_date.month)[1]
                check_end = end_date.replace(day=last_day)

                # 1. เช็คจำนวนวันลา
                cursor.execute("SELECT COUNT(*) FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, check_start, check_end))
                l_count = cursor.fetchone()[0]
                
                # 2. เช็คสถานะรายวัน (สาย, ขาด, ออกก่อน, ไม่ครบ)
                cursor.execute("""
                    SELECT COUNT(*) FROM employee_daily_records 
                    WHERE emp_id = %s AND work_date BETWEEN %s AND %s 
                    AND (status LIKE '%%สาย%%' OR status LIKE '%%ขาด%%' OR status LIKE '%%ออกก่อน%%' OR status LIKE '%%ไม่ครบ%%')
                """, (emp_id, check_start, check_end))
                r_count = cursor.fetchone()[0]
                
                # 3. 🛠️ [FIX BUG] เช็คสแกนนิ้วก่อน 08:00 (โดย "ละเว้น" วันหยุด ไม่เอามาคิดว่าสาย)
                cursor.execute("""
                    SELECT COUNT(*) FROM (
                        SELECT DATE(t.scan_timestamp) as d, MIN(t.scan_timestamp)::time as first_scan 
                        FROM time_attendance_logs t
                        LEFT JOIN employee_daily_records ed ON t.emp_id = ed.emp_id AND DATE(t.scan_timestamp) = ed.work_date
                        WHERE t.emp_id=%s AND DATE(t.scan_timestamp) BETWEEN %s AND %s 
                          AND (ed.status IS NULL OR ed.status NOT LIKE '%%วันหยุด%%')
                        GROUP BY d
                    ) s WHERE first_scan > '08:00:59'""", (emp_id, check_start, check_end))
                late_8_count = cursor.fetchone()[0]
                
                # 🛠️ [DEBUG] พิมพ์บอกในหน้าจอดำ (Console) เพื่อให้เรารู้สาเหตุชัดเจน
                print(f"[{emp_id}] 🔍 เบี้ยขยัน (เต็มเดือน): ลา={l_count} วัน | ประวัติเสีย={r_count} วัน | เข้าหลัง 8 โมง={late_8_count} วัน")

                # ถ้าประวัติขาวสะอาดทั้งเดือน ถึงจะได้เงิน
                if (l_count + r_count + late_8_count) == 0:
                    _, reward = get_diligence_streak_info(emp_id, end_date.month, end_date.year)
                    result["diligence"] = float(reward)
                    print(f"  => ✅ ผ่านเกณฑ์! ได้รับเบี้ยขยัน: {reward} บาท")
                else:
                    print(f"  => ❌ ไม่ผ่านเกณฑ์")
                
                # ถ้าประวัติขาวสะอาดทั้งเดือน ถึงจะได้เงิน
                if (l_count + r_count + late_8_count) == 0:
                    _, reward = get_diligence_streak_info(emp_id, end_date.month, end_date.year)
                    result["diligence"] = float(reward)

            # --- [2] จัดการรายรับอื่นๆ (SUM ONCE) ---
            result["other_income"] = (
                float(user_inputs.get('other_income', 0)) + 
                float(user_inputs.get('housing_income', 0)) + 
                float(user_inputs.get('other_income_manual', 0))
            )
            w_flags = emp_info.get('welfare', [])
            w_amts = emp_info.get('welfare_amounts', [])
            for idx, f in enumerate(w_flags):
                if f: result["other_income"] += float(w_amts[idx] or 0)

            # ดึงค่า OT ที่กรอกมือจาก Popup มาเก็บไว้ก่อน
            manual_ot_money = float(user_inputs.get('ot', 0))

            result["commission"] = float(user_inputs.get('commission', 0))
            result["incentive"] = float(user_inputs.get('incentive', 0))
            result["bonus"] = float(user_inputs.get('bonus', 0))
            result["provident_fund"] = float(user_inputs.get('provident_fund', 0))
            result["loan"] = float(user_inputs.get('loan', 0))
            result["other_deduct"] = float(user_inputs.get('other_deduct', 0)) + float(user_inputs.get('housing_deduct', 0))
            manual_tax = float(user_inputs.get('tax', 0))

            # --- [3] วันทำงาน และ ขาด/ลา/สาย ---
            cursor.execute("SELECT work_date, status, ot_hours, is_ot_approved, total_amount FROM employee_daily_records WHERE emp_id = %s AND work_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            daily_map = {row['work_date']: dict(row) for row in cursor.fetchall()}
            cursor.execute("SELECT leave_date, num_days, leave_type FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
            leave_map = {row['leave_date']: dict(row) for row in cursor.fetchall()}
            cursor.execute("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holiday_dates = {row['holiday_date'] for row in cursor.fetchall()}

            actual_days = 0.0; penalty_hrs = 0.0; absent_days = 0.0; no_pay_days = 0.0; auto_trip = 0.0
            total_approved_ot_hours = 0.0 

            for d in pd.date_range(start_date, end_date).date:
                rec = daily_map.get(d)
                lv = leave_map.get(d)
                is_h = (d.weekday() == 6 or d in holiday_dates)
                
                if lv:
                    num_d = float(lv['num_days'] or 1.0)
                    is_unpaid = 'ลาไม่รับค่าจ้าง' in str(lv['leave_type'])
                    if num_d < 1.0: 
                        actual_days += 1.0
                        if is_unpaid or is_daily_calculation: 
                            penalty_hrs += (num_d * 8.0) # 🛠️ หักเวลาครึ่งวันไปก่อน 4 ชม.
                            # 🛠️ [NEW] ถ้าทำงานไม่ครบอีก (ขาดเพิ่ม) ให้บวกชั่วโมงที่ขาดเข้าไปด้วย!
                            if rec and 'หัก' in str(rec['status']):
                                m = re.search(r"หัก\s*([\d\.]+)\s*ชม", str(rec['status']))
                                if m: penalty_hrs += float(m.group(1))
                    else:
                        if is_unpaid: no_pay_days += 1.0
                        
                    if rec: 
                        auto_trip += float(rec.get('total_amount', 0))
                        if rec.get('is_ot_approved'): total_approved_ot_hours += float(rec.get('ot_hours', 0.0))
                
                elif rec:
                    status_text = str(rec['status'])
                    if 'ขาดงาน' in status_text: 
                        absent_days += 1.0
                    elif 'วันหยุด' in status_text:
                        pass
                    else:
                        actual_days += 1.0
                        m = re.search(r"หัก\s*([\d\.]+)\s*ชม", status_text)
                        if m: penalty_hrs += float(m.group(1))
                    
                    auto_trip += float(rec.get('total_amount', 0))
                    if rec.get('is_ot_approved'):
                        total_approved_ot_hours += float(rec.get('ot_hours', 0.0))

                elif not is_h and not is_fixed: 
                    absent_days += 1.0

            # --- [4] สรุปยอดเงิน และ OT ---
            cursor.execute("SELECT position_allowance FROM salary_history WHERE emp_id = %s ORDER BY history_id DESC LIMIT 1", (emp_id,))
            pa = cursor.fetchone()
            result["position_allowance"] = float(pa[0]) if pa and pa[0] else 0.0

            # 🛠️ [NEW] กฎพิเศษสำหรับ "กรรมการ" (รับเงินเต็มจำนวนเสมอ ไม่หักขาด/ลา/สาย)
            is_director = "กรรมการ" in emp_type
            if is_director:
                penalty_hrs = 0.0  # ล้างเวลาสาย
                absent_days = 0.0  # ล้างวันขาดงาน
                no_pay_days = 0.0  # ล้างวันลาไม่รับค่าจ้าง

            # 🛠️ แยกการคำนวณเรทรายชั่วโมงระหว่าง "รายวัน" และ "รายเดือน"
            if is_daily_calculation:
                day_r = salary_db
                hourly_r = salary_db / 8.0
            else:
                day_r = salary_db / 30.0
                hourly_r = (salary_db / 30.0) / 8.0

            ot_rate = hourly_r * 1.5
            
            # รวมเงิน OT
            auto_ot_money = total_approved_ot_hours * ot_rate
            result["ot"] = auto_ot_money + manual_ot_money

            if is_daily_calculation:
                result["base_salary"] = actual_days * salary_db
                result["late_deduct"] = penalty_hrs * hourly_r # รายวัน หักแค่รายชั่วโมงที่หายไป วันที่ขาดจะไม่ได้เงินอยู่แล้ว
            else:
                result["base_salary"] = salary_db
                result["late_deduct"] = (penalty_hrs * hourly_r) + (absent_days * day_r) + (no_pay_days * day_r)

            result["driving_allowance"] = auto_trip
            
            # ยอดรวมรับทั้งหมด
            result["total_income"] = (
                result["base_salary"] + result["position_allowance"] + result["ot"] + 
                result["commission"] + result["incentive"] + result["diligence"] + 
                result["bonus"] + result["other_income"] + result["driving_allowance"]
            )

            # --- [5] ภาษีและ SSO ---
            if is_pnd3_applicable: result["pnd3"] = result["total_income"] * 0.03
            else: result["pnd1"] = (result["commission"] * 0.03) + manual_tax
            result["tax"] = result["pnd1"] + result["pnd3"]

            if not is_sso_exempt and (end_date.day == calendar.monthrange(end_date.year, end_date.month)[1]):
                sso_config = load_sso_config(end_date.year)
                sso_base = min(max(result["base_salary"] + result["position_allowance"], 1650), float(sso_config.get("max_salary", 15000)))
                result["sso"] = int(sso_base * (float(sso_config.get("rate", 5.0))/100.0) + 0.5)

            result["total_deduct"] = (result["sso"] + result["tax"] + result["provident_fund"] + result["late_deduct"] + result["loan"] + result["other_deduct"])
            result["net_salary"] = result["total_income"] - result["total_deduct"]

    except Exception as e:
        print(f"Payroll Error: {e}")
    return result

def process_attendance_summary(start_date, end_date):
    """
    (ฉบับ V93.0 - Warehouse Strict Late Penalty & No Compensation)
    - 🛠️ ลอจิกพนักงานคลัง: เข้า > 08:00 หัก 0.5 ชม. | เข้า > 08:30 หัก 1.0 ชม. (ไม่ต้องชดเชยเวลาเลิก เลิก 17:00 ปกติ)
    - 🛠️ ลอจิกพนักงานออฟฟิศ: ใช้ระบบ "ชดเชยเวลา" ต้องอยู่ให้ครบ 9 ชม. ไม่งั้นโดนหักเวลาที่ไม่ครบ
    - 🛠️ ตัดวินาทีทิ้ง ป้องกันขาดงาน 30 วินาทีแล้วโดนหัก
    """
    import calendar
    from datetime import date, datetime, time, timedelta
    import pandas as pd
    import math 

    conn = get_db_connection()
    if not conn: return []
    
    summary_report = []

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""SELECT emp_id, fname, lname, work_location, department, position, emp_type 
                              FROM employees 
                              WHERE status IS NULL OR status NOT IN ('พ้นสภาพพนักงาน', 'ลาออก') 
                              OR (termination_date >= %s) ORDER BY emp_id ASC""", (start_date,))
            employees = [dict(row) for row in cursor.fetchall()]
            
            cursor.execute("SELECT holiday_date, description FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holiday_dict = {row['holiday_date']: row['description'] for row in cursor.fetchall()}
            
            cursor.execute("SELECT emp_id, leave_date, leave_type, num_days FROM employee_leave_records WHERE leave_date BETWEEN %s AND %s", (start_date, end_date))
            leaves = {}
            for row in cursor.fetchall(): 
                eid = str(row['emp_id'])
                leaves.setdefault(eid, {})[row['leave_date']] = row

            cursor.execute("""SELECT emp_id, scan_timestamp FROM time_attendance_logs 
                              WHERE DATE(scan_timestamp) BETWEEN %s AND %s ORDER BY scan_timestamp ASC""", (start_date, end_date))
            logs_map = {}
            for row in cursor.fetchall():
                eid = str(row['emp_id'])
                dt = row['scan_timestamp'].date()
                logs_map.setdefault(eid, {}).setdefault(dt, []).append(row['scan_timestamp'])

            cursor.execute("""SELECT emp_id, work_date, ot_hours, ot_in_time, ot_out_time, status, is_ot_approved, total_amount, remark 
                              FROM employee_daily_records WHERE work_date BETWEEN %s AND %s""", (start_date, end_date))
            daily_records_map = {}
            for row in cursor.fetchall():
                daily_records_map[(str(row['emp_id']), row['work_date'])] = dict(row)

            all_dates = pd.date_range(start_date, end_date).date
            
            for emp in employees:
                emp_id = str(emp['emp_id'])
                emp_name = f"{emp['fname']} {emp['lname']}"
                work_loc = emp.get('work_location', '')
                emp_type_str = str(emp.get('emp_type', ''))
                
                # --- 🛠️ 1. จัดกลุ่มพนักงาน ---
                is_warehouse = work_loc and "คลังสินค้า" in work_loc
                is_daily_emp = "รายวัน" in emp_type_str or "Daily" in emp_type_str
                is_contractor = ("จ้างเหมา" in emp_type_str or "สัญญาจ้าง" in emp_type_str) and not ("สำนักงาน" in work_loc or "ออฟฟิศ" in work_loc)
                
                is_flexible_group = is_warehouse or is_daily_emp or is_contractor

                if emp_id == '1007':
                    is_flexible_group = True  # บังคับใช้กฎพนักงานคลัง
                elif emp_id == '1006':
                    is_flexible_group = False
                
                allow_ot_calc = is_flexible_group

                # --- 🛠️ 2. ตั้งค่ากฎเวลาตามกลุ่ม ---
                if is_flexible_group:
                    min_out_time = time(17, 0)    
                    # กฎคลังใหม่ ไม่ใช้ late_cutoff แบบออฟฟิศแล้ว จะไปเช็คตรงๆ ด้านล่างแทน
                else: 
                    min_out_time = time(17, 30)   
                    late_cutoff = time(9, 0)      
                    tier_2_late = time(9, 30)

                total_late_mins_penalty = 0
                total_absent_days = 0.0
                daily_details = []
                
                for curr_date in all_dates:
                    status = "ปกติ"
                    final_penalty_mins = 0
                    ot_hours_to_save = 0.0 
                    scan_in_str = "-"
                    scan_out_str = "-"
                    
                    existing_rec = daily_records_map.get((emp_id, curr_date))
                    if existing_rec:
                        is_ot_approved = bool(existing_rec.get('is_ot_approved', False))
                        if is_ot_approved: ot_hours_to_save = float(existing_rec.get('ot_hours', 0))

                    day_logs = logs_map.get(emp_id, {}).get(curr_date, [])
                    leave_info = leaves.get(emp_id, {}).get(curr_date)
                    is_holiday = curr_date in holiday_dict
                    is_sunday = (curr_date.weekday() == 6)

                    if day_logs:
                        scan_in = min(day_logs).time()
                        scan_out = max(day_logs).time()
                        scan_in_str = scan_in.strftime("%H:%M")
                        if len(day_logs) > 1: scan_out_str = scan_out.strftime("%H:%M")

                        leave_type_str = ""
                        num_days = 1.0
                        if leave_info:
                            leave_type_str = str(leave_info['leave_type']).strip()
                            num_days = float(leave_info.get('num_days', 1.0))
                            status = f"ลา {leave_type_str}"
                            if num_days < 1.0: status += f" ({num_days} วัน)"

                        # 🛠️ [ตัดวินาทีทิ้ง] ป้องกันเศษวินาทีทำให้ชั่วโมงไม่ครบ
                        dummy = datetime.today()
                        t_in_clean = datetime.strptime(scan_in_str, "%H:%M").time()
                        t_out_clean = datetime.strptime(scan_out_str, "%H:%M").time() if len(day_logs) > 1 else t_in_clean
                        
                        t_in_dt = datetime.combine(dummy, t_in_clean)
                        t_out_dt = datetime.combine(dummy, t_out_clean)
                        
                        # --- คำนวณชั่วโมงทำงานพื้นฐาน ---
                        duration_mins = (t_out_dt - t_in_dt).total_seconds() / 60.0
                        deduct_break = 0
                        if t_in_clean < time(12,0) and t_out_clean > time(13,0): deduct_break = 60
                        net_work_mins = max(0, duration_mins - deduct_break)
                        
                        target_mins = 480 
                        is_nopay = (leave_info and "ลาไม่รับค่าจ้าง" in leave_type_str)
                        
                        if not is_nopay and leave_info:
                            if num_days == 0.5: target_mins = 240 
                            elif num_days < 1.0: target_mins = 480 - int(num_days * 8 * 60)

                        penalty_A_late = 0
                        penalty_B_missing = 0
                        deduct_hours_B = 0
                        missing_mins = 0

                        if not (leave_info and num_days < 1.0):
                            # =========================================================
                            # 🛠️ [ลอจิกแยกกลุ่ม] คลังสินค้า VS ออฟฟิศ
                            # =========================================================
                            if is_flexible_group:
                                # กฎ 1: พนักงานคลังสินค้า (และ 1007)
                                required_out_dt = datetime.combine(dummy, min_out_time) # 17:00 เสมอ ไม่ต้องชดเชยเวลา

                                # เช็คสาย หักแบบขั้นบันได
                                if t_in_clean > time(8, 30):
                                    penalty_A_late = 60 # สายเกิน 08:30 หัก 1 ชม.
                                elif t_in_clean > time(8, 0):
                                    penalty_A_late = 30 # สาย 08:01 - 08:30 หัก ครึ่ง ชม.
                                
                                # เช็คออกก่อน 17:00
                                if t_out_dt < required_out_dt:
                                    missing_mins = (required_out_dt - t_out_dt).total_seconds() / 60.0

                            else:
                                # กฎ 2: พนักงานออฟฟิศปกติ (ต้องชดเชยเวลา)
                                shift_end_dt = t_in_dt + timedelta(hours=9) # เข้าตอนไหน + ไป 9 ชม. (รวมพัก)
                                min_out_dt = datetime.combine(dummy, min_out_time) # 17:30
                                required_out_dt = max(min_out_dt, shift_end_dt) # ต้องออกช้ากว่าเวลาเลิกงานขั้นต่ำ และ ครบ 9 ชม.

                                # เช็คมาสาย
                                if t_in_clean > late_cutoff:
                                    if t_in_clean > tier_2_late: penalty_A_late = 120
                                    else: penalty_A_late = 60
                                        
                                # เช็คเวลาออก
                                if t_out_dt < required_out_dt:
                                    missing_mins = (required_out_dt - t_out_dt).total_seconds() / 60.0
                                
                                # เช็คชั่วโมงรวมเผื่อไว้ (ออฟฟิศเท่านั้น)
                                if net_work_mins < (target_mins - 5):
                                    missing_work_mins = target_mins - net_work_mins
                                    missing_mins = max(missing_mins, missing_work_mins)
                            # =========================================================

                            if missing_mins > 0:
                                if is_nopay: deduct_hours_B = math.ceil(missing_mins / 30.0) * 0.5
                                else: deduct_hours_B = math.ceil(missing_mins / 60.0) * 1.0
                                penalty_B_missing = deduct_hours_B * 60

                        # --- [STEP 3] ตัดสินโทษ ---
                        if is_nopay:
                            final_penalty_mins = penalty_B_missing
                            status = f"ลาไม่รับค่าจ้าง (หัก {deduct_hours_B} ชม.)" if deduct_hours_B > 0 else f"ลาไม่รับค่าจ้าง"
                        elif leave_info:
                            final_penalty_mins = 0 
                        else:
                            # 🛠️ เอาโทษที่หนักที่สุดระหว่าง มาสาย กับ ออกก่อนเวลา
                            final_penalty_mins = max(penalty_A_late, penalty_B_missing)
                            
                            if final_penalty_mins == penalty_A_late and penalty_A_late > 0:
                                hrs = penalty_A_late / 60
                                status = f"สาย (หัก {hrs:.1f} ชม.)".replace(".0", "")
                            elif final_penalty_mins == penalty_B_missing and penalty_B_missing > 0:
                                # ถ้าออกก่อนเวลาขั้นต่ำ (17:00/17:30) ให้ใช้คำว่า "ออกก่อนเวลา"
                                if t_out_clean < min_out_time:
                                    status = f"ออกก่อนเวลา (หัก {deduct_hours_B:.1f} ชม.)".replace(".0", "")
                                else:
                                    status = f"ชม.ไม่ครบ (หัก {deduct_hours_B:.1f} ชม.)".replace(".0", "")

                        # คำนวณ OT ยึดจากเวลาที่ควรออก (required_out_dt)
                        if allow_ot_calc and not is_ot_approved and not leave_info and len(day_logs) > 1:
                            if t_out_dt > required_out_dt:
                                raw_ot_mins = int((t_out_dt - required_out_dt).total_seconds() / 60)
                                if raw_ot_mins >= 60: ot_hours_to_save = float(int(raw_ot_mins / 60))

                    else: 
                        if leave_info: 
                             l_type = str(leave_info['leave_type']).strip()
                             num_days = float(leave_info.get('num_days', 1.0))
                             status = f"ลา {l_type}"
                             if num_days < 1.0: status += f" ({num_days} วัน)"
                             if "ลาไม่รับค่าจ้าง" in l_type: final_penalty_mins = num_days * 8 * 60
                        elif is_holiday: status = f"วันหยุด ({holiday_dict[curr_date]})"
                        elif is_sunday: status = "วันหยุด"
                        else:
                            status = "ขาดงาน"
                            total_absent_days += 1.0

                    if "กรรมการ" in emp_type_str:
                        final_penalty_mins = 0
                        if status == "ขาดงาน":
                            total_absent_days -= 1.0  # คืนค่าจำนวนวันขาดงาน
                        if status != "ปกติ":
                            status = "ไม่ต้องลงเวลา"  # เปลี่ยนสถานะให้ดูสุภาพขึ้น

                    if status not in ["ขาดงาน", "วันหยุด", "ไม่ต้องลงเวลา"] and "ใช้สิทธิ์" not in status:
                        total_late_mins_penalty += final_penalty_mins

                    should_save = (ot_hours_to_save > 0) or (day_logs) or (status not in ["ปกติ", "ไม่ต้องลงเวลา"]) or existing_rec
                    if should_save:
                        w_in = scan_in_str if scan_in_str != "-" else None
                        w_out = scan_out_str if scan_out_str != "-" else None
                        total_amt = float(existing_rec.get('total_amount', 0) or 0) if existing_rec else 0.0
                        remark_text = existing_rec.get('remark', '') if existing_rec else '' 
                        
                        cursor.execute("""
                            INSERT INTO employee_daily_records (emp_id, work_date, ot_hours, status, work_in_time, work_out_time, total_amount, remark)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                            ON CONFLICT (emp_id, work_date) DO UPDATE SET
                            ot_hours = EXCLUDED.ot_hours, status = EXCLUDED.status, 
                            work_in_time = EXCLUDED.work_in_time, work_out_time = EXCLUDED.work_out_time,
                            remark = EXCLUDED.remark;
                        """, (emp_id, curr_date, ot_hours_to_save, status, w_in, w_out, total_amt, remark_text))

                    remark_text = existing_rec.get('remark', '') if existing_rec else ''
                    daily_details.append({
                        "date": curr_date, "status": status, "scan_in": scan_in_str, "scan_out": scan_out_str,
                        "penalty_hrs": final_penalty_mins / 60.0, "remark": remark_text
                    })

                summary_report.append({
                    "emp_id": emp_id, "name": emp_name, "emp_type": emp.get('emp_type', ''), 
                    "department": emp.get('department', '-'), "position": emp.get('position', '-'),
                    "total_late_minutes": total_late_mins_penalty, "total_late_hours": total_late_mins_penalty / 60.0,
                    "absent_days": total_absent_days, "details": daily_details
                })
            
            conn.commit()
            return summary_report

    except Exception as e:
        print(f"Error in process_attendance_summary: {e}")
        import traceback; traceback.print_exc()
        return []
    finally:
        conn.close()
        
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



def get_asmart_connection():
    """สร้าง Connection ไปหาฐานข้อมูล A+ Smart (เครื่อง 192.168.1.60)"""
    try:
        return psycopg2.connect(
            host="Server-APrime",      # IP ของ A+ Smart
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

def get_commission_from_asmart_by_name(emp_name, month, year):
    """
    ดึงยอดคอมมิชชั่นจาก A+ Smart โดยใช้ "ชื่อ-นามสกุล"
    (แก้ไข: คืนค่ายอดก่อนหัก 3% โดยเอา net_commission / 0.97)
    """
    conn = get_asmart_connection() # 💡 ต้องเป็น get_asmart_connection()
    if not conn: return 0.0
    
    try:
        with conn.cursor() as cursor:
            # 1. สร้างคีย์เวิร์ดค้นหาแบบฉลาด (Smart Search)
            clean_name = emp_name.replace("นาย", "").replace("นางสาว", "").replace("นาง", "").strip()
            name_parts = clean_name.split()
            
            if len(name_parts) >= 2:
                search_pattern = f"%{name_parts[0]}%{name_parts[-1]}%"
            else:
                search_pattern = f"%{clean_name}%"

            # 🛠️ 2. Query แบบ JOIN (หาร 0.97 เพื่อคืนค่ายอดตั้งต้นก่อนหัก 3%)
            query = """
                SELECT SUM(cpl.net_commission) / 0.97
                FROM commission_payout_logs cpl
                JOIN sales_users u ON cpl.sale_key = u.sale_key
                WHERE u.sale_name LIKE %s 
                  AND EXTRACT(MONTH FROM cpl.timestamp) = %s 
                  AND EXTRACT(YEAR FROM cpl.timestamp) = %s
            """
            
            cursor.execute(query, (search_pattern, month, year))
            result = cursor.fetchone()
            
            if result and result[0]:
                return float(result[0])
            return 0.0

    except Exception as e:
        print(f"Error fetching commission by name: {e}")
        return 0.0
    finally:
        if conn: conn.close()

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

# ไฟล์: hr_database.py

def get_daily_records(emp_id, month, year):
    """ดึงข้อมูลรายวัน (เพิ่ม total_amount เพื่อให้เห็นยอดเงินค่าเที่ยว)"""
    conn = get_db_connection()
    if not conn: return {}
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            import calendar
            last_day = calendar.monthrange(year, month)[1]
            start_date = f"{year}-{month:02d}-01"
            end_date = f"{year}-{month:02d}-{last_day}"
            
            # 🔴 [FIXED] เพิ่ม total_amount ในบรรทัด SELECT
            cursor.execute("""
                SELECT work_date, status, ot_hours, trip_pickup, trip_crane, total_amount
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
    บันทึกรายการเที่ยวรถ (อัปเดต V4: เพิ่มค่าบริการรวมลง)
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # 1. ลบข้อมูลเก่า
            cursor.execute("DELETE FROM employee_driving_details WHERE emp_id = %s AND work_date = %s", (emp_id, work_date))
            
            # 2. ใส่ข้อมูลใหม่
            for item in details_list:
                is_free_val = bool(item.get('is_free', False))
                is_service_val = bool(item.get('is_service', False)) # <--- ใหม่
                service_fee_val = float(item.get('service_fee', 0))  # <--- ใหม่
                
                cursor.execute("""
                    INSERT INTO employee_driving_details 
                    (emp_id, work_date, car_type, license_plate, driver_name, delivery_date, trip_cost, ref_doc_type, ref_doc_id, is_free, is_service, service_fee)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    emp_id, work_date, 
                    item.get('car_type'), item.get('license'), item.get('driver'), 
                    item.get('send_date'), item.get('cost'),
                    item.get('doc_type', ''), item.get('doc_id', ''),
                    is_free_val, is_service_val, service_fee_val
                ))
            
            # 3. อัปเดตยอดรวมลงตารางหลัก (ค่าเที่ยว + ค่าบริการ)
            total_money = sum(x['cost'] + x.get('service_fee', 0) for x in details_list)
            
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

def save_monthly_payroll(emp_id, month, year, pay_date, data, calc_mode=0):
    """บันทึกงวดเงินเดือนลง Database (อัปเดต V90.2: แก้บั๊ก Transaction Aborted)"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # 1. แอบสร้างคอลัมน์และแก้ข้อจำกัด (พร้อมระบบเคลียร์ Error ฐานข้อมูล)
            try:
                cursor.execute("ALTER TABLE payroll_records ADD COLUMN IF NOT EXISTS period_mode INTEGER DEFAULT 0;")
                cursor.execute("ALTER TABLE payroll_records DROP CONSTRAINT IF EXISTS payroll_records_emp_id_period_month_period_year_key;")
                cursor.execute("ALTER TABLE payroll_records ADD CONSTRAINT unique_payroll_period UNIQUE(emp_id, period_month, period_year, period_mode);")
                conn.commit() # ถ้าสำเร็จให้ Commit ก่อนเลย
            except Exception:
                conn.rollback() # 🛠️ [FIX BUG] ถ้าเคยมี Constraint นี้แล้วมัน Error ให้เคลียร์สถานะทิ้งซะ ระบบจะได้ไม่งอแงในคำสั่งถัดไป

            # 2. บันทึกข้อมูล
            cursor.execute("""
                INSERT INTO payroll_records (
                    emp_id, period_month, period_year, period_mode, payment_date,
                    base_salary, position_allowance, ot_pay, commission, 
                    incentive, diligence, bonus,
                    driving_allowance, other_income, total_income,
                    sso_deduct, tax_deduct, provident_fund, loan_deduct,
                    late_deduct, other_deduct, total_deduct, net_salary,
                    remark, note
                ) VALUES (
                    %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, 
                    %s, %s, %s, 
                    %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s, %s, %s, 
                    %s, %s
                )
                ON CONFLICT (emp_id, period_month, period_year, period_mode) DO UPDATE SET
                    payment_date = EXCLUDED.payment_date,
                    base_salary = EXCLUDED.base_salary,
                    position_allowance = EXCLUDED.position_allowance,
                    ot_pay = EXCLUDED.ot_pay,
                    commission = EXCLUDED.commission,
                    incentive = EXCLUDED.incentive,
                    diligence = EXCLUDED.diligence,
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
                    net_salary = EXCLUDED.net_salary,
                    remark = EXCLUDED.remark,
                    note = EXCLUDED.note;
            """, (
                emp_id, month, year, calc_mode, pay_date,
                data.get('base_salary', 0), data.get('position_allowance', 0),
                data.get('ot', 0), data.get('commission', 0), 
                data.get('incentive', 0), data.get('diligence', 0), data.get('bonus', 0), 
                data.get('driving_allowance', 0), data.get('other_income', 0), data.get('total_income', 0),
                data.get('sso', 0), data.get('tax', 0), data.get('provident_fund', 0),
                data.get('loan', 0), data.get('late_deduct', 0), data.get('other_deduct', 0),
                data.get('total_deduct', 0), data.get('net_salary', 0),
                data.get('remark', ''), data.get('note', '')
            ))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving payroll record: {e}")
        return False
    finally:
        if conn: conn.close()

def get_monthly_payroll_records(month, year, calc_mode=0):
    """ดึงข้อมูลตามเดือน/ปี และ รอบวิก (calc_mode)"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # --- 🛠️ [FIX] ตรวจสอบและสร้างคอลัมน์ period_mode หากยังไม่มี ---
            try:
                cursor.execute("ALTER TABLE payroll_records ADD COLUMN IF NOT EXISTS period_mode INTEGER DEFAULT 0;")
                conn.commit()
            except Exception as e:
                conn.rollback() # เคลียร์สถานะ Error ของ Transaction
            # ---------------------------------------------------------

            cursor.execute("""
                SELECT pr.*, e.fname, e.lname, e.id_card, e.position, e.department
                FROM payroll_records pr
                JOIN employees e ON pr.emp_id = e.emp_id
                WHERE pr.period_month = %s AND pr.period_year = %s AND pr.period_mode = %s
                ORDER BY pr.emp_id ASC
            """, (month, year, calc_mode))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error fetching payroll history: {e}")
        return []
    finally:
        if conn: conn.close() # 🛠️ ปิด Connection ให้สมบูร

def get_annual_pnd1k_data(year_ce):
    """
    ดึงข้อมูลทำ ภ.ง.ด. 1ก (สรุปรายได้/ภาษี ทั้งปี)
    (แก้ไข: เพิ่ม start_date เพื่อใช้เรียงลำดับความอาวุโส)
    """
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT 
                    e.emp_id, e.fname, e.lname, e.id_card, e.address, e.position, e.start_date,
                    SUM(pr.total_income) as annual_income,
                    SUM(pr.tax_deduct) as annual_tax,
                    SUM(pr.sso_deduct) as annual_sso,
                    SUM(pr.provident_fund) as annual_fund
                FROM payroll_records pr
                JOIN employees e ON pr.emp_id = e.emp_id
                WHERE pr.period_year = %s
                GROUP BY e.emp_id, e.fname, e.lname, e.id_card, e.address, e.position, e.start_date
                ORDER BY e.start_date ASC  -- เรียงจากเข้าก่อน (อาวุโสสุด) ไปหลัง
            """, (year_ce,))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error fetching PND1K data: {e}")
        return []
    finally:
        conn.close()
    
def check_leave_quota_status(emp_id, year, leave_type, req_days):
    """
    ตรวจสอบสถานะโควตาลา
    Returns: (is_pass, message, remaining_days)
    """
    conn = get_db_connection()
    if not conn: return False, "Database Error", 0
    
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. ดึงโควตาทั้งปี
            cursor.execute("""
                SELECT leave_annual_days, leave_sick_days, leave_personal_days, 
                       leave_maternity_days, leave_ordination_days 
                FROM employees WHERE emp_id = %s
            """, (emp_id,))
            quota_row = cursor.fetchone()
            if not quota_row: return False, "ไม่พบข้อมูลพนักงาน", 0
            
            # Map ชื่อประเภทลา กับ คอลัมน์ใน DB
            type_map = {
                "ลาพักร้อน": "leave_annual_days",
                "ลาป่วย": "leave_sick_days",
                "ลากิจ": "leave_personal_days",
                "ลาคลอด": "leave_maternity_days",
                "ลาบวช": "leave_ordination_days"
            }
            
            col_name = type_map.get(leave_type)
            if not col_name: 
                return True, "ลาประเภทอื่น (ไม่จำกัดโควตา)", 999 # ลาอื่นๆ ไม่เช็ค
                
            total_quota = float(quota_row[col_name] or 0)
            
            # 2. ดึงยอดที่ใช้ไปแล้วในปีนี้
            cursor.execute("""
                SELECT SUM(num_days) 
                FROM employee_leave_records 
                WHERE emp_id = %s AND leave_type = %s AND EXTRACT(YEAR FROM leave_date) = %s
            """, (emp_id, leave_type, year))
            
            used_days = float(cursor.fetchone()[0] or 0)
            
            # 3. คำนวณ
            remaining = total_quota - used_days
            
            if req_days > remaining:
                return False, f"สิทธิ์วันลาไม่พอ! (เหลือ {remaining} วัน, ขอ {req_days} วัน)", remaining
            else:
                return True, f"อนุมัติ (เหลือ {remaining - req_days} วัน)", remaining

    except Exception as e:
        print(f"Quota Check Error: {e}")
        return False, f"Error: {e}", 0
    finally:
        conn.close()

def add_audit_log(actor, action, emp_id, emp_name, field, old_val, new_val):
    """บันทึกเหตุการณ์ลง Audit Log"""
    conn = get_db_connection()
    if not conn: return
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO audit_logs 
                (actor_name, action_type, target_emp_id, target_emp_name, changed_field, old_value, new_value)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (actor, action, emp_id, emp_name, field, str(old_val), str(new_val)))
            conn.commit()
    except Exception as e:
        print(f"Audit Log Error: {e}")
    finally:
        conn.close()

def get_employee_annual_summary(emp_id, year_ce):
    """
    ดึงข้อมูลสรุปรายได้/ภาษี ทั้งปี ของพนักงาน 1 คน (เพื่อทำ 50 ทวิ)
    (Update: เพิ่มการหาลำดับที่ (Sequence No) ให้ตรงกับ ภ.ง.ด. 1ก)
    """
    conn = get_db_connection()
    if not conn: return None
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. ดึงข้อมูลส่วนตัว
            cursor.execute("""
                SELECT fname, lname, id_card, address, position, start_date 
                FROM employees WHERE emp_id = %s
            """, (emp_id,))
            emp = cursor.fetchone()
            if not emp: return None
            
            # 2. ดึงยอดเงินรวมทั้งปี
            cursor.execute("""
                SELECT 
                    SUM(total_income) as total_income,
                    SUM(tax_deduct) as total_tax,
                    SUM(sso_deduct) as total_sso,
                    SUM(provident_fund) as total_fund,
                    MIN(period_month) as start_month,
                    MAX(period_month) as end_month
                FROM payroll_records 
                WHERE emp_id = %s AND period_year = %s
            """, (emp_id, year_ce))
            payroll = cursor.fetchone()
            
            # 3. [NEW] หา "ลำดับที่" ให้ตรงกับ ภ.ง.ด. 1ก
            # โดยการดึงรายการ ภ.ง.ด. 1ก ทั้งหมดของปีนั้น (ซึ่งเรียงตาม start_date แล้ว) มาหา index
            # (นี่คือวิธีที่แม่นยำที่สุดเพื่อให้ลำดับตรงกันเป๊ะ)
            pnd1k_list = get_annual_pnd1k_data(year_ce)
            sequence_no = 0
            for index, item in enumerate(pnd1k_list):
                if str(item['emp_id']) == str(emp_id):
                    sequence_no = index + 1
                    break
            
            # กรณีหาไม่เจอ (เช่น เพิ่งเพิ่มข้อมูลแต่ยังไม่ได้คำนวณเงินเดือน) ให้เป็น 0 หรือ 999
            if sequence_no == 0: sequence_no = 999

            s_month = int(payroll['start_month']) if payroll['start_month'] else 1
            e_month = int(payroll['end_month']) if payroll['end_month'] else 12

            return {
                "fname": emp['fname'],
                "lname": emp['lname'],
                "id_card": emp['id_card'],
                "address": emp['address'],
                "total_income": float(payroll['total_income'] or 0),
                "total_tax": float(payroll['total_tax'] or 0),
                "total_sso": float(payroll['total_sso'] or 0),
                "total_fund": float(payroll['total_fund'] or 0),
                "start_month": s_month,
                "end_month": e_month,
                "sequence_no": sequence_no  # <--- ส่งค่าลำดับกลับไปด้วย
            }
    except Exception as e:
        print(f"Error getting annual summary: {e}")
        return None
    finally:
        conn.close()

def get_ot_details_list(emp_id, work_date):
    """ดึงรายการช่วงเวลา OT ทั้งหมดของวันนั้น"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT * FROM employee_ot_details
                WHERE emp_id = %s AND work_date = %s
                ORDER BY ot_detail_id ASC
            """, (emp_id, work_date))
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error getting OT details: {e}")
        return []
    finally:
        conn.close()

def save_ot_details_list(emp_id, work_date, ot_list):
    """
    บันทึกรายการ OT แบบช่วง (ลบของเก่าในวันนั้นทิ้ง แล้วใส่ใหม่)
    ot_list = [{'start': '17:00', 'end': '19:00', 'hours': 2.0, 'desc': '...'}, ...]
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # 1. ลบข้อมูลเก่าของวันนั้นทิ้งก่อน
            cursor.execute("DELETE FROM employee_ot_details WHERE emp_id = %s AND work_date = %s", (emp_id, work_date))
            
            total_ot_hours = 0.0
            
            # 2. ใส่ข้อมูลใหม่เข้าไปทีละช่วง
            for item in ot_list:
                h = float(item.get('hours', 0))
                total_ot_hours += h
                
                cursor.execute("""
                    INSERT INTO employee_ot_details 
                    (emp_id, work_date, start_time, end_time, period_hours, description)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    emp_id, work_date, 
                    item.get('start'), item.get('end'), h, item.get('desc', '')
                ))
            
            # 3. อัปเดตยอดรวม OT ลงตารางหลัก (employee_daily_records)
            # เพื่อให้ Payroll ดึงไปคำนวณเงินได้เลย โดยไม่ต้องแก้สูตร Payroll
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, ot_hours, status)
                VALUES (%s, %s, %s, 'ทำงาน')
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                ot_hours = EXCLUDED.ot_hours;
            """, (emp_id, work_date, total_ot_hours))
            
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving OT details: {e}")
        return False
    finally:
        conn.close()
    
def get_ot_details_list(emp_id, work_date):
    """ดึงรายการช่วงเวลา OT ทั้งหมดของวันนั้น"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            cursor.execute("""
                SELECT * FROM employee_ot_details
                WHERE emp_id = %s AND work_date = %s
                ORDER BY ot_detail_id ASC
            """, (emp_id, work_date))
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error getting OT details: {e}")
        return []
    finally:
        conn.close()

def save_ot_details_list(emp_id, work_date, ot_list):
    """
    บันทึกรายการ OT แบบช่วง (ลบของเก่าในวันนั้นทิ้ง แล้วใส่ใหม่)
    ot_list = [{'start': '17:00', 'end': '19:00', 'hours': 2.0, 'desc': '...'}, ...]
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # 1. ลบข้อมูลเก่าของวันนั้นทิ้งก่อน
            cursor.execute("DELETE FROM employee_ot_details WHERE emp_id = %s AND work_date = %s", (emp_id, work_date))
            
            total_ot_hours = 0.0
            
            # 2. ใส่ข้อมูลใหม่เข้าไปทีละช่วง
            for item in ot_list:
                h = float(item.get('hours', 0))
                total_ot_hours += h
                
                cursor.execute("""
                    INSERT INTO employee_ot_details 
                    (emp_id, work_date, start_time, end_time, period_hours, description)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    emp_id, work_date, 
                    item.get('start'), item.get('end'), h, item.get('desc', '')
                ))
            
            # 3. อัปเดตยอดรวม OT ลงตารางหลัก (employee_daily_records)
            # เพื่อให้ Payroll ดึงไปคำนวณเงินได้เลย โดยไม่ต้องแก้สูตร Payroll
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, ot_hours, status)
                VALUES (%s, %s, %s, 'ทำงาน')
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                ot_hours = EXCLUDED.ot_hours;
            """, (emp_id, work_date, total_ot_hours))
            
            conn.commit()
            return True
    except Exception as e:
        print(f"Error saving OT details: {e}")
        return False
    finally:
        conn.close()

def update_employee_ot_times(emp_id, work_date, ot_in, ot_out, new_ot_hours):
    """
    (แก้ไข V2) อัปเดตเวลาเข้า-ออก OT และจำนวนชั่วโมง OT (ที่คำนวณมาใหม่)
    """
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # ใช้ UPSERT: อัปเดตเวลา OT และ จำนวนชั่วโมง OT
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, ot_in_time, ot_out_time, ot_hours, status)
                VALUES (%s, %s, %s, %s, %s, 'ทำงาน')
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                    ot_in_time = EXCLUDED.ot_in_time,
                    ot_out_time = EXCLUDED.ot_out_time,
                    ot_hours = EXCLUDED.ot_hours;
            """, (str(emp_id), work_date, ot_in, ot_out, float(new_ot_hours)))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error updating OT times: {e}")
        return False
    finally:
        if conn: conn.close()

def update_ot_approval_status(emp_id, work_date, is_approved):
    """อัปเดตสถานะอนุมัติ OT"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            # อัปเดตเฉพาะฟิลด์ is_ot_approved
            # (ใช้ UPSERT แบบ Update-only ถ้ามี record อยู่แล้ว)
            cursor.execute("""
                UPDATE employee_daily_records
                SET is_ot_approved = %s
                WHERE emp_id = %s AND work_date = %s
            """, (is_approved, str(emp_id), work_date))
            
            # ถ้ายังไม่มี Record (ซึ่งแปลก เพราะ OT ควรถูกสร้างตอนคำนวณแล้ว)
            # เราอาจเลือกที่จะไม่ Insert ใหม่เพื่อความปลอดภัย หรือ Insert ก็ได้
            # ในที่นี้เอาแค่ Update ก่อนครับ
            
            conn.commit()
            return True
    except Exception as e:
        print(f"Error updating OT approval: {e}")
        return False
    finally:
        conn.close()

def get_driving_details_range(emp_id, start_date, end_date):
    """ดึงรายการเที่ยวรถทั้งหมด ในช่วงวันที่ระบุ (สำหรับ Drill-down ดูรายละเอียด)"""
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # ดึงรายละเอียด: วันที่, ทะเบียน, ประเภท, คนขับ, ราคา
            cursor.execute("""
                SELECT 
                    work_date, 
                    car_type, 
                    license_plate, 
                    driver_name,
                    trip_cost,
                    service_fee
                FROM employee_driving_details
                WHERE emp_id = %s 
                  AND work_date BETWEEN %s AND %s
                ORDER BY work_date ASC
            """, (emp_id, start_date, end_date))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error fetching driving details range: {e}")
        return []
    finally:
        if conn: conn.close()

def get_daily_records_range(emp_id, start_date, end_date):
    """
    ดึงข้อมูลรายวัน (Daily Records) ในช่วงเวลาที่กำหนด 
    เพื่อใช้แสดงรายละเอียด OT และ การเข้างาน (เบี้ยขยัน)
    """
    conn = get_db_connection()
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # ดึงข้อมูลจาก employee_daily_records
            # และอาจจะ Join กับตารางอื่นถ้าต้องการข้อมูลเพิ่ม (เช่น Log สแกนนิ้ว)
            # ในที่นี้เอาจาก daily_records เป็นหลักก่อน
            cursor.execute("""
                SELECT 
                    work_date, 
                    status, 
                    ot_hours, 
                    ot_in_time, 
                    ot_out_time, 
                    work_in_time, 
                    work_out_time,
                    is_ot_approved,
                    total_amount,
                    remark -- (เผื่ออยากดูค่าเที่ยวด้วย)
                FROM employee_daily_records
                WHERE emp_id = %s 
                  AND work_date BETWEEN %s AND %s
                ORDER BY work_date ASC
            """, (emp_id, start_date, end_date))
            
            rows = [dict(row) for row in cursor.fetchall()]
            
            # --- (Optional) เสริมข้อมูล: ถ้าวันไหนไม่มีใน daily_records แต่มีสแกนนิ้ว ---
            # อาจจะไปดึง time_attendance_logs มา merge ด้วยก็ได้ 
            # แต่เบื้องต้นเอาแค่นี้ก่อนเพื่อความรวดเร็ว
            
            return rows
            
    except Exception as e:
        print(f"Error fetching daily records range: {e}")
        return []
    finally:
        if conn: conn.close()

def get_ytd_summary(emp_id, year, current_month):
    """
    ดึงยอดสะสมตั้งแต่ต้นปี (ม.ค.) ถึงเดือนก่อนหน้า (current_month - 1)
    เพื่อใช้คำนวณภาษีแบบสะสม
    Returns: (รายได้สะสม, ภาษีที่จ่ายแล้ว, ประกันสังคมสะสม)
    """
    conn = get_db_connection()
    if not conn: return 0.0, 0.0, 0.0
    try:
        with conn.cursor() as cursor:
            # ดึงข้อมูลจากตาราง payroll_records
            cursor.execute("""
                SELECT 
                    COALESCE(SUM(total_income), 0),
                    COALESCE(SUM(tax_deduct), 0),
                    COALESCE(SUM(sso_deduct), 0)
                FROM payroll_records 
                WHERE emp_id = %s 
                  AND period_year = %s 
                  AND period_month < %s
            """, (str(emp_id), year, current_month))
            
            row = cursor.fetchone()
            if row:
                return float(row[0]), float(row[1]), float(row[2])
            return 0.0, 0.0, 0.0
            
    except Exception as e:
        print(f"YTD Error: {e}")
        return 0.0, 0.0, 0.0
    finally:
        if conn: conn.close()

def get_dashboard_stats():
    """ดึงข้อมูลสรุปสำหรับ Dashboard (ฉบับสมบูรณ์ ครบ 6 Query)"""
    conn = get_db_connection()
    if not conn: return {}

    stats = {
        "total_employees": 0,
        "on_leave_today": 0,
        "late_today": 0,
        "probation_upcoming": [],
        "dept_counts": [],
        "missing_scans": [] 
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
            next_30_days = today + timedelta(days=30)
            cursor.execute("""
                SELECT emp_id, fname, lname, probation_end_date, department
                FROM employees 
                WHERE probation_end_date BETWEEN %s AND %s
                AND status IN ('ระหว่างทดลองงาน')
                ORDER BY probation_end_date ASC
            """, (today, next_30_days))
            stats["probation_upcoming"] = [dict(row) for row in cursor.fetchall()]

            # 6. แจ้งเตือน: ไม่สแกนนิ้ว/ขาดงาน (ประจำเดือนนี้) [ที่เพิ่มไปล่าสุด]
            current_month = today.month
            current_year = today.year
            cursor.execute("""
                SELECT e.fname, e.lname, e.department, ed.work_date
                FROM employee_daily_records ed
                JOIN employees e ON ed.emp_id = e.emp_id
                WHERE EXTRACT(MONTH FROM ed.work_date) = %s 
                  AND EXTRACT(YEAR FROM ed.work_date) = %s
                  AND ed.status = 'ขาดงาน'
                ORDER BY ed.work_date DESC
                LIMIT 50
            """, (current_month, current_year))
            stats["missing_scans"] = [dict(row) for row in cursor.fetchall()]
            
    except Exception as e:
        print(f"Dashboard DB Error: {e}")
    finally:
        if conn: conn.close()
        
    return stats

def load_allowance_settings():
    """โหลดรายการสวัสดิการและการตั้งค่าภาษีจาก DB (ถ้าไม่มีให้ใช้ค่า Default)"""
    import json
    conn = get_db_connection()
    if not conn: return []
    
    default_settings = [
        {"name": "ค่าตำแหน่ง", "is_taxable": True},
        {"name": "ค่าที่พัก", "is_taxable": True},
        {"name": "ค่าอาหาร", "is_taxable": True},
        {"name": "ค่าเดินทาง", "is_taxable": True},
        {"name": "ค่าน้ำมัน", "is_taxable": True},
        {"name": "เบี้ยเลี้ยง", "is_taxable": True},
        {"name": "ค่าโทรศัพท์", "is_taxable": True}
    ]
    
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT setting_value FROM company_settings WHERE setting_key = 'allowance_config'")
            result = cursor.fetchone()
            if result and result[0]:
                return json.loads(result[0])
            else:
                # ถ้ายังไม่มีใน DB ให้บันทึกค่า Default ลงไปก่อน
                save_allowance_settings(default_settings)
                return default_settings
    except Exception as e:
        print(f"Error loading allowance settings: {e}")
        return default_settings
    finally:
        if conn: conn.close()

def save_allowance_settings(settings_list):
    """บันทึกการตั้งค่าสวัสดิการลง DB"""
    import json
    return save_company_setting('allowance_config', json.dumps(settings_list, ensure_ascii=False))

# --- [เพิ่มต่อท้ายไฟล์ hr_database.py] ---

def load_sso_config(year):
    """โหลดตั้งค่าประกันสังคมของปีที่ระบุ (คืนค่า Default หากไม่พบ)"""
    import json
    from datetime import datetime
    
    conn = get_db_connection()
    target_year = str(year)
    default_config = {"rate": 5.0, "min_salary": 1650, "max_salary": 15000}
    
    if not conn: return default_config

    try:
        with conn.cursor() as cursor:
            # ดึงก้อน Setting ใหญ่มา
            cursor.execute("SELECT setting_value FROM company_settings WHERE setting_key = 'sso_config_yearly'")
            result = cursor.fetchone()
            
            if result and result[0]:
                all_configs = json.loads(result[0])
                # คืนค่าของปีที่ขอ (หรือคืน Default ถ้าปีนั้นยังไม่ตั้งค่า)
                return all_configs.get(target_year, default_config)
            else:
                return default_config
    except Exception as e:
        print(f"Error loading sso config: {e}")
        return default_config
    finally:
        if conn: conn.close()

def save_sso_config(year, config_dict):
    """บันทึกตั้งค่าประกันสังคมลงในปีที่ระบุ"""
    import json
    
    conn = get_db_connection()
    if not conn: return False
    target_year = str(year)
    
    try:
        # 1. โหลดข้อมูลเก่ามาก่อน (เพื่อไม่ให้ปีอื่นหาย)
        current_data = {}
        with conn.cursor() as cursor:
            cursor.execute("SELECT setting_value FROM company_settings WHERE setting_key = 'sso_config_yearly'")
            result = cursor.fetchone()
            if result and result[0]:
                current_data = json.loads(result[0])
        
        # 2. อัปเดตปีที่ต้องการ
        current_data[target_year] = config_dict
        
        # 3. บันทึกกลับลงไป
        final_json = json.dumps(current_data, ensure_ascii=False)
        return save_company_setting('sso_config_yearly', final_json)
        
    except Exception as e:
        print(f"Error saving sso config: {e}")
        return False
    finally:
        if conn: conn.close()

# (เพิ่มต่อท้ายไฟล์ hr_database.py)

def get_existing_sso_years():
    """ดึงรายการปี (พ.ศ.) ทั้งหมดที่มีการบันทึกค่าไว้ใน DB"""
    import json
    conn = get_db_connection()
    years = set()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT setting_value FROM company_settings WHERE setting_key = 'sso_config_yearly'")
            result = cursor.fetchone()
            if result and result[0]:
                data = json.loads(result[0])
                # แปลง ค.ศ. (Key ใน DB) เป็น พ.ศ. เพื่อแสดงผล
                for y_str in data.keys():
                    if y_str.isdigit():
                        years.add(int(y_str) + 543)
    except: pass
    finally:
        if conn: conn.close()
    return years

# (เพิ่มใน hr_database.py)

def mass_upsert_employees_from_excel(file_path):
    """
    นำเข้าข้อมูลพนักงานจาก Excel แบบ Mass Upsert
    - ถ้ายังไม่มีรหัสพนักงานนี้ -> Insert (เพิ่มใหม่)
    - ถ้ามีรหัสนี้แล้ว -> Update (แก้ไขข้อมูลเดิม)
    """
    import pandas as pd
    conn = get_db_connection()
    if not conn: return 0, "Database Connection Failed"

    try:
        # 1. อ่านไฟล์ Excel
        df = pd.read_excel(file_path)
        
        # แปลงชื่อคอลัมน์ Excel ให้ตรงกับ Database (ปรับแก้ตรงนี้ตามไฟล์ Excel ของคุณ)
        # ตัวอย่าง: Excel Header -> DB Column
        column_map = {
            'รหัส': 'emp_id',
            'รหัสพนักงาน': 'emp_id',
            'ชื่อ': 'fname',
            'นามสกุล': 'lname',
            'ชื่อเล่น': 'nickname',
            'ตำแหน่ง': 'position',
            'แผนก': 'department',
            'เงินเดือน': 'salary',
            'ค่าตำแหน่ง': 'position_allowance',
            'ประเภท': 'emp_type',  # รายวัน/รายเดือน
            'เลขบัตรประชาชน': 'id_card',
            'วันเริ่มงาน': 'start_date'
        }
        df.rename(columns=column_map, inplace=True)
        
        # 2. คัดกรองเอาเฉพาะคอลัมน์ที่มีจริงใน DB
        valid_cols = ['emp_id', 'fname', 'lname', 'nickname', 'position', 'department', 
                      'salary', 'emp_type', 'id_card', 'start_date', 'status']
        
        # กรองเอาเฉพาะคอลัมน์ที่แมตช์และมีข้อมูล
        df_ready = df[df.columns.intersection(valid_cols)].where(pd.notnull(df), None)
        
        if 'emp_id' not in df_ready.columns:
            return 0, "ไม่พบคอลัมน์ 'รหัสพนักงาน' หรือ 'emp_id' ในไฟล์ Excel"

        success_count = 0
        
        with conn.cursor() as cursor:
            for index, row in df_ready.iterrows():
                emp_id = str(row['emp_id']).strip()
                if not emp_id: continue

                # สร้าง Dynamic Query สำหรับ UPDATE
                update_parts = []
                insert_cols = []
                insert_vals = []
                query_params = []
                
                for col in df_ready.columns:
                    val = row[col]
                    if col == 'emp_id': 
                        insert_cols.append(col)
                        insert_vals.append(val)
                        continue
                        
                    # เตรียมข้อมูลสำหรับ Insert
                    insert_cols.append(col)
                    insert_vals.append(val)
                    
                    # เตรียมข้อมูลสำหรับ Update (เฉพาะค่าที่ไม่ว่าง)
                    if val is not None:
                        update_parts.append(sql.SQL("{} = %s").format(sql.Identifier(col)))
                        query_params.append(val)

                # ถ้าไม่มีข้อมูลให้อัปเดตเลย (มีแค่ ID) ก็ข้าม
                if not update_parts: continue

                # เพิ่ม ID เข้าไปใน params สำหรับ WHERE clause ของ Update
                # แต่เดี๋ยวก่อน! เราใช้ ON CONFLICT เพื่อทำ Upsert ทีเดียวเลยดีกว่า
                
                # สร้าง SQL Upsert
                # INSERT INTO employees (...) VALUES (...) 
                # ON CONFLICT (emp_id) DO UPDATE SET col1=val1, col2=val2 ...
                
                insert_sql = sql.SQL("""
                    INSERT INTO employees ({fields})
                    VALUES ({values})
                    ON CONFLICT (emp_id) DO UPDATE
                    SET {updates}
                """).format(
                    fields=sql.SQL(', ').join(map(sql.Identifier, insert_cols)),
                    values=sql.SQL(', ').join(sql.Placeholder() * len(insert_cols)),
                    updates=sql.SQL(', ').join(
                        sql.SQL("{} = EXCLUDED.{}").format(sql.Identifier(col), sql.Identifier(col)) 
                        for col in insert_cols if col != 'emp_id'
                    )
                )
                
                # รวม Params (สำหรับ Insert values)
                cursor.execute(insert_sql, insert_vals)
                success_count += 1

            conn.commit()
            return success_count, f"อัปเดตข้อมูลสำเร็จ {success_count} รายการ"

    except Exception as e:
        print(f"Import Error: {e}")
        return 0, f"เกิดข้อผิดพลาด: {e}"
    finally:
        conn.close()

def get_leave_record_by_id(leave_id):
    try:
        conn = get_db_connection()
        if not conn: return None

        # เปลี่ยนมาใช้ DictCursor เพื่อให้ระบบดึงข้อมูลมา "ครบทุกคอลัมน์" อัตโนมัติ
        from psycopg2 import extras
        cursor = conn.cursor(cursor_factory=extras.DictCursor)
        
        cursor.execute("SELECT * FROM employee_leave_records WHERE leave_id = %s", (leave_id,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row) # คืนค่ากลับไปทั้งหมด (รวมถึง leave_date, เวลาเริ่ม-สิ้นสุด)
        return None
        
    except Exception as e:
        print(f"Error fetching leave record: {e}")
        return None
        
def update_leave_record(leave_id, leave_date, l_type, l_days, l_reason, start_time=None, end_time=None):
    try:
        conn = get_db_connection()
        if not conn: return False

        cursor = conn.cursor()
        # เพิ่ม leave_date เข้าไปในคำสั่ง SET
        sql = """
            UPDATE employee_leave_records 
            SET leave_date=%s, leave_type=%s, num_days=%s, reason=%s, 
                leave_start_time=%s, leave_end_time=%s 
            WHERE leave_id=%s
        """
        # เรียงตัวแปรให้ตรงกับคำสั่ง SQL
        cursor.execute(sql, (leave_date, l_type, l_days, l_reason, start_time, end_time, leave_id))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Error updating leave: {e}")
        return False

def delete_scan_logs_range(start_date, end_date):
    """ลบข้อมูลสแกนนิ้วในช่วงวันที่กำหนด (ใช้สำหรับโหมด Mass Edit/Replace)"""
    conn = get_db_connection()
    if not conn: return False
    
    try:
        with conn.cursor() as cursor:
            # ลบข้อมูลในช่วงวันที่นั้นๆ ทั้งหมด
            cursor.execute("""
                DELETE FROM time_attendance_logs 
                WHERE DATE(scan_timestamp) BETWEEN %s AND %s
            """, (start_date, end_date))
            conn.commit()
            print(f"Deleted logs from {start_date} to {end_date}")
            return True
    except Exception as e:
        conn.rollback()
        print(f"Delete Error: {e}")
        return False
    finally:
        if conn: conn.close()

def get_employee_leave_history(emp_id=None):
    """ดึงประวัติการลา (รองรับการกรองรายคน)"""
    conn = get_db_connection() #
    if not conn: return []
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            query = """
                SELECT l.leave_id, l.emp_id, e.fname, e.lname, l.leave_date, 
                       l.leave_type, l.num_days, l.reason
                FROM employee_leave_records l
                JOIN employees e ON l.emp_id = e.emp_id
                WHERE 1=1
            """
            params = []
            if emp_id:
                query += " AND l.emp_id = %s"
                params.append(str(emp_id))
            
            query += " ORDER BY l.leave_date DESC"
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()

def delete_leave_record_by_id(leave_id):
    try:
        conn = get_db_connection()
        if not conn: return False
        cursor = conn.cursor()
        # 👇 แก้ตรงนี้ด้วย
        cursor.execute("DELETE FROM employee_leave_records WHERE leave_id = %s", (leave_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Error deleting: {e}")
        return False

def update_daily_remark(emp_id, work_date, remark_text):
    """บันทึกหมายเหตุรายวันจากหน้า Popup"""
    conn = get_db_connection()
    if not conn: return False
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO employee_daily_records (emp_id, work_date, remark, status)
                VALUES (%s, %s, %s, 'ปกติ')
                ON CONFLICT (emp_id, work_date) DO UPDATE SET
                    remark = EXCLUDED.remark;
            """, (str(emp_id), work_date, remark_text))
            conn.commit()
            return True
    except Exception as e:
        print(f"Error updating remark: {e}")
        return False
    finally:
        if conn: conn.close()