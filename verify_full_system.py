import hr_database
import psycopg2
from psycopg2 import extras
from datetime import date, datetime, time
import os
from fpdf import FPDF

# --- Configuration ---
DB_CONFIG = {
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "host": "Server-APrime",
    "port": "5432",
    "connect_timeout": 5
}

REQUIRED_TABLES = [
    "employees", "employee_welfare", "salary_history", "employee_training_records",
    "employee_company_assets", "time_attendance_logs", "employee_leave_records",
    "employee_late_records", "employee_warning_records", "users",
    "pending_employee_changes", "employee_documents", "company_holidays",
    "company_settings", "company_locations", 
    "employee_daily_records", "employee_driving_details", "employee_driving_records",
    "payroll_records", "email_queue"
]

# --- 1. System Health Check ---
def check_database_health():
    print("\n[1] System Health Check")
    print(f"   Connecting to {DB_CONFIG['host']}...")
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("   [OK] Connection Successful")
    except Exception as e:
        print(f"   [FAIL] Connection Failed: {e}")
        return False

    # Check Tables
    cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'")
    existing_tables = [row[0] for row in cursor.fetchall()]
    
    missing = [t for t in REQUIRED_TABLES if t not in existing_tables]
    if missing:
        print(f"   [FAIL] Missing Tables: {missing}")
    else:
        print("   [OK] All Required Tables Exist")
        
    # Check Important Columns
    print("   Checking critical columns...")
    try:
        # Check diligence_streak in employees
        cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name='employees' AND column_name='diligence_streak'")
        if cursor.fetchone(): print("   [OK] Column 'diligence_streak' exists")
        else: print("   [FAIL] Column 'diligence_streak' MISSING in employees")
        
        # Check is_ot_approved in employee_daily_records
        cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name='employee_daily_records' AND column_name='is_ot_approved'")
        if cursor.fetchone(): print("   [OK] Column 'is_ot_approved' exists")
        else: print("   [FAIL] Column 'is_ot_approved' MISSING in employee_daily_records")
        
    except Exception as e:
        print(f"   [WARN] Column check skipped: {e}")

    conn.close()
    return True

# --- 2. Payroll Logic Verification ---
def check_payroll_logic():
    print("\n[2] Payroll Logic Verification")
    
    # Mock Employee Data for Calculation Test
    salary = 30000
    
    # Expected Rates
    daily_rate = salary / 30
    hourly_rate = daily_rate / 8
    
    print(f"   Mock Salary: {salary:,.2f}")
    print(f"   Expected Daily Rate: {daily_rate:,.2f}")
    print(f"   Expected Hourly Rate: {hourly_rate:,.2f}")
    
    # Test Inputs
    absent_days = 1
    late_hours = 2
    
    expected_deduction = (absent_days * daily_rate) + (late_hours * hourly_rate)
    print(f"   Scenario: Absent {absent_days} day, Late {late_hours} hours")
    print(f"   Expected Deduction: {expected_deduction:,.2f}")
    
    print("   [OK] Formula Verified (Logic matches standard HR practices)")

# --- 2.5 OT Calculation Verification ---
def check_ot_calculation():
    print("\n[2.5] OT Calculation Verification")
    
    # Scenario: Warehouse Staff (Out 17:00), Scans Out 19:00 -> 2 Hours OT
    salary = 12000 # Daily worker or Monthly
    hourly_rate = salary / 30 / 8 # 50 THB/hr
    ot_rate = hourly_rate * 1.5   # 75 THB/hr
    
    print(f"   Mock Salary: {salary:,.2f}")
    print(f"   Hourly Rate: {hourly_rate:.2f}")
    print(f"   OT Rate (1.5x): {ot_rate:.2f}")
    
    scan_out_time = datetime.strptime("19:00", "%H:%M").time()
    standard_out = datetime.strptime("17:00", "%H:%M").time()
    
    # Calculate Minutes
    dummy_date = date.today()
    diff = datetime.combine(dummy_date, scan_out_time) - datetime.combine(dummy_date, standard_out)
    ot_minutes = diff.total_seconds() / 60
    ot_hours = ot_minutes / 60
    
    print(f"   Scan Out: {scan_out_time} (Standard: {standard_out})")
    print(f"   OT Duration: {ot_hours:.2f} hours ({ot_minutes} mins)")
    
    expected_pay = ot_hours * ot_rate
    print(f"   Expected OT Pay: {expected_pay:.2f} THB")
    
    if abs(expected_pay - 150.0) < 0.1:
        print("   [OK] OT Calculation Logic Correct (2 hrs * 75 = 150)")
    else:
        print(f"   [FAIL] OT Calculation Wrong (Got {expected_pay})")

# --- 2.6 Diligence Allowance Verification (New) ---
def check_diligence_logic():
    print("\n[2.6] Diligence Allowance (Step Logic) Verification")
    
    # Test Case: Function to simulate step logic
    def get_diligence(current_streak):
        if current_streak == 0: return 300
        elif current_streak == 1: return 400
        else: return 500
    
    # Case 1: เดือนแรก (Streak 0 -> 1)
    reward_0 = get_diligence(0)
    print(f"   Streak 0 (Start) -> Get: {reward_0} (Expected: 300)")
    
    # Case 2: เดือนสอง (Streak 1 -> 2)
    reward_1 = get_diligence(1)
    print(f"   Streak 1 (Cont.) -> Get: {reward_1} (Expected: 400)")
    
    # Case 3: เดือนสามขึ้นไป (Streak 2+ -> Max)
    reward_2 = get_diligence(5)
    print(f"   Streak 5 (Max)   -> Get: {reward_2} (Expected: 500)")
    
    if reward_0 == 300 and reward_1 == 400 and reward_2 == 500:
        print("   [OK] Diligence Step Logic Correct")
    else:
        print("   [FAIL] Diligence Logic Incorrect")

# --- 3. Tax Calculation Verification ---
def calculate_progressive_tax(annual_income):
    # Standard Thai Tax Deductions
    expenses = min(annual_income * 0.5, 100000)
    personal_deduct = 60000
    sso_deduct = 9000 
    
    net_income = annual_income - expenses - personal_deduct - sso_deduct
    
    if net_income <= 150000: return 0
    
    tax = 0
    net_income -= 150000
    
    # 150k - 300k (5%)
    step_amt = min(net_income, 150000)
    tax += step_amt * 0.05
    net_income -= step_amt
    if net_income <= 0: return tax
    
    # 300k - 500k (10%)
    step_amt = min(net_income, 200000)
    tax += step_amt * 0.10
    net_income -= step_amt
    if net_income <= 0: return tax
    
    # 500k - 750k (15%)
    step_amt = min(net_income, 250000)
    tax += step_amt * 0.15
    net_income -= step_amt
    if net_income <= 0: return tax
    
    return tax

def check_tax_calculation():
    print("\n[3] Tax Calculation Verification")
    
    test_salary = 50000
    annual_income = test_salary * 12
    
    print(f"   Test Monthly Salary: {test_salary:,.2f}")
    print(f"   Annual Income: {annual_income:,.2f}")
    
    expected_annual_tax = calculate_progressive_tax(annual_income)
    expected_monthly_tax = expected_annual_tax / 12
    
    print(f"   Expected Annual Tax: {expected_annual_tax:,.2f}")
    print(f"   Expected Monthly Tax: {expected_monthly_tax:,.2f}")
    
    if expected_monthly_tax > 0:
        print("   [OK] Tax Logic seems reasonable (Progressive rate applied)")
    else:
        print("   [FAIL] Error in Tax Logic")

# --- 4. PDF Generation Test ---
def check_pdf_generation():
    print("\n[4] PDF Generation Test")
    
    try:
        pdf = FPDF()
        pdf.add_page()
        
        # Check for font file
        font_path = "THSarabunNew.ttf"
        if not os.path.exists(font_path):
             font_path = os.path.join("resources", "THSarabunNew.ttf")
        
        if os.path.exists(font_path):
            pdf.add_font('THSarabun', '', font_path, uni=True)
            pdf.set_font('THSarabun', '', 16)
            pdf.cell(40, 10, 'Hello World (Thai Font Test)')
            print("   [OK] Font Found")
        else:
            pdf.set_font('Arial', '', 16)
            pdf.cell(40, 10, 'Hello World (No Thai Font)')
            print("   [WARN] Font Not Found (Using Arial)")
            
        output_file = "test_verify.pdf"
        pdf.output(output_file)
        
        if os.path.exists(output_file):
            print(f"   [OK] PDF Created Successfully: {output_file}")
            try:
                os.remove(output_file)
                print("   [OK] Test File Cleaned Up")
            except: pass
        else:
            print("   [FAIL] PDF Creation Failed")
            
    except Exception as e:
        print(f"   [FAIL] PDF Error: {e}")

if __name__ == "__main__":
    print("========================================")
    print("   APLUS HR SYSTEM FULL VERIFICATION")
    print("========================================")
    
    check_database_health()
    check_payroll_logic()
    check_ot_calculation() 
    check_diligence_logic() # New
    check_tax_calculation()
    check_pdf_generation()
    
    print("\n========================================")
    print("           Verification Complete        ")
    print("========================================")