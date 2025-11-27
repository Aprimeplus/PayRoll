import os

file_path = r"c:\Users\Nitro V15\Desktop\PayRoll\hr_database.py"
print("Starting repair script v2...")

# The full block to insert
# Includes the end of add_company_holiday, all missing functions, and the start of get_dashboard_stats
replacement_code = r'''    except Exception as e:
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
        messagebox.showerror("DB Error (Add Scan)", f"ไม่สามารถเพิ่ม Log สแกนได้:\n{e}")
        return False
    finally:
        if conn: conn.close()

def process_attendance_summary(start_date, end_date):
    """
    ประมวลผลข้อมูลการเข้างาน (Attendance)
    - คำนวณ มาสาย (Late)
    - คำนวณ ขาดงาน (Absent)
    - คำนวณ OT (Overtime)
    """
    summary_report = []
    conn = get_db_connection()
    if not conn: return []
    
    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            # 1. ดึงข้อมูลพนักงานทั้งหมด
            cursor.execute("SELECT emp_id, work_location FROM employees WHERE status = 'Active' OR status IS NULL OR status = '' ORDER BY emp_id")
            employees = cursor.fetchall()
            
            # 2. ดึงวันหยุดบริษัท
            cursor.execute("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
            holidays = set(row[0] for row in cursor.fetchall())
            
            # 3. ดึงข้อมูลการลา
            cursor.execute("SELECT emp_id, leave_date, leave_type FROM employee_leave_records WHERE leave_date BETWEEN %s AND %s", (start_date, end_date))
            leaves = {}
            for row in cursor.fetchall():
                key = (row['emp_id'], row['leave_date'])
                leaves[key] = row['leave_type']

            # 4. ดึงข้อมูล Scan Logs
            cursor.execute("""
                SELECT emp_id, scan_timestamp 
                FROM time_attendance_logs 
                WHERE DATE(scan_timestamp) BETWEEN %s AND %s
                ORDER BY scan_timestamp ASC
            """, (start_date, end_date))
            all_logs = cursor.fetchall()
            
            logs_by_emp_date = {}
            for row in all_logs:
                e_id = row['emp_id']
                ts = row['scan_timestamp']
                d = ts.date()
                if (e_id, d) not in logs_by_emp_date:
                    logs_by_emp_date[(e_id, d)] = []
                logs_by_emp_date[(e_id, d)].append(ts)

            # 5. กำหนดกฎการเข้างาน
            WORK_RULES = {
                "สำนักงานใหญ่": { "standard_in": time(9,0), "standard_out": time(18,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
                "คลังสินค้า": { "standard_in": time(8,0), "standard_out": time(17,0), "tier_1_cutoff": time(8,30), "penalty_1_mins": 60, "penalty_2_mins": 120 },
                "default": { "standard_in": time(9,0), "standard_out": time(18,0), "tier_1_cutoff": time(9,30), "penalty_1_mins": 60, "penalty_2_mins": 120 }
            }

            # 6. Loop ประมวลผล
            current_date = start_date
            while current_date <= end_date:
                is_holiday = current_date in holidays
                is_weekend = current_date.weekday() == 6 # Sunday
                
                for emp in employees:
                    emp_id = emp['emp_id']
                    work_location = emp['work_location'] or "default"
                    if work_location not in WORK_RULES: work_location = "default"
                    rule = WORK_RULES[work_location]
                    
                    daily_logs = logs_by_emp_date.get((emp_id, current_date), [])
                    scan_in = daily_logs[0] if daily_logs else None
                    scan_out = daily_logs[-1] if len(daily_logs) > 1 else None
                    
                    is_late = False
                    late_minutes = 0
                    is_absent = False
                    status = "Normal"
                    ot_minutes = 0
                    
                    leave_type = leaves.get((emp_id, current_date))
                    
                    if leave_type:
                        status = f"Leave ({leave_type})"
                    elif is_holiday:
                        status = "Holiday"
                    elif is_weekend:
                        status = "Weekend"
                    elif not scan_in:
                        is_absent = True
                        status = "Absent"
                    else:
                        # มาทำงาน -> เช็คสาย
                        scan_in_time = scan_in.time()
                        if scan_in_time > rule['tier_1_cutoff']:
                            is_late = True
                            late_minutes = rule['penalty_2_mins']
                            status = "Late (>30m)"
                        elif scan_in_time > rule['standard_in']:
                            is_late = True
                            late_minutes = rule['penalty_1_mins']
                            status = "Late (<30m)"
                            
                    # คำนวณ OT
                    if scan_out and rule.get('standard_out'):
                        scan_out_time = scan_out.time()
                        std_out = rule['standard_out']
                        if scan_out_time > std_out:
                            dummy = datetime.today()
                            t1 = datetime.combine(dummy, scan_out_time)
                            t2 = datetime.combine(dummy, std_out)
                            diff = t1 - t2
                            ot_minutes = int(diff.total_seconds() / 60)

                    summary_report.append({
                        "emp_id": emp_id,
                        "date": current_date,
                        "scan_in": scan_in,
                        "scan_out": scan_out,
                        "is_late": is_late,
                        "late_minutes": late_minutes,
                        "is_absent": is_absent,
                        "status": status,
                        "ot_minutes": ot_minutes
                    })
                
                current_date += timedelta(days=1)

    except Exception as e:
        messagebox.showerror("DB Error (Process)", f"ไม่สามารถประมวลผลข้อมูลได้:\n{e}")
        return []
    finally:
        if conn: conn.close()
        
    return summary_report

def _calculate_social_security(base_salary):
    if base_salary < 1650:
        base_salary = 1650.0
    if base_salary > 15000:
        base_salary = 15000.0
    deduction = round(base_salary * 0.05) 
    return float(deduction)

def calculate_payroll_for_employee(emp_id, start_date, end_date, user_inputs=None):
    """
    (สมองหลัก Payroll V18.0 - รองรับค่าเที่ยวรถแบบ "เที่ยวฟรี" โดยดึงยอดเงินจริง)
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
        "debug_penalty_hours": 0.0, "debug_absent_days": 0.0,
        "debug_worked_days": 0.0 
    }

    conn = get_db_connection()
    if not conn: return result

    try:
        with conn.cursor(cursor_factory=extras.DictCursor) as cursor:
            
            # --- 1. ดึงข้อมูลพนักงาน ---
            emp_info = load_single_employee(emp_id)
            if not emp_info or not emp_info.get("salary"): return result

            salary_from_db = float(emp_info.get("salary", 0.0))
            emp_type = emp_info.get("emp_type", "")
            
            # ดึงค่าตำแหน่ง
            cursor.execute("SELECT position_allowance FROM salary_history WHERE emp_id = %s ORDER BY history_id DESC LIMIT 1", (emp_id,))
            pos_row = cursor.fetchone()
            position_allowance = float(pos_row[0]) if pos_row and pos_row[0] else 0.0
            result["position_allowance"] = position_allowance

            # --- 2. รับค่าจาก User Inputs ---
            manual_ot = float(user_inputs.get('ot', 0))
            result["commission"] = float(user_inputs.get('commission', 0))
            result["bonus"] = float(user_inputs.get('bonus', 0))
            result["other_income"] = float(user_inputs.get('other_income', 0))
            result["tax"] = float(user_inputs.get('tax', 0))
            result["provident_fund"] = float(user_inputs.get('provident_fund', 0))
            result["loan"] = float(user_inputs.get('loan', 0))
            result["other_deduct"] = float(user_inputs.get('other_deduct', 0))

            # --- 3. ดึงข้อมูลประกอบ ---
            work_location = emp_info.get('work_location', "default") or "default"
            if work_location not in WORK_RULES: work_location = "default"
            rules = WORK_RULES[work_location]
            
            try:
                cursor.execute("SELECT holiday_date FROM company_holidays WHERE holiday_date BETWEEN %s AND %s", (start_date, end_date))
                holiday_dates_set = {row['holiday_date'] for row in cursor.fetchall()}
            except: conn.rollback(); holiday_dates_set = set()
            
            try:
                cursor.execute("SELECT leave_date, leave_type, num_days, leave_start_time, leave_end_time FROM employee_leave_records WHERE emp_id = %s AND leave_date BETWEEN %s AND %s", (emp_id, start_date, end_date))
                leave_records_dict = {row['leave_date']: {'type': row['leave_type'], 'days': float(row['num_days'] or 1.0), 'start': row['leave_start_time'], 'end': row['leave_end_time']} for row in cursor.fetchall()}
            except: conn.rollback(); leave_records_dict = {}

            cursor.execute("SELECT scan_timestamp FROM time_attendance_logs WHERE emp_id = %s AND scan_timestamp::date BETWEEN %s AND %s ORDER BY scan_timestamp", (emp_id, start_date, end_date))
            scans = cursor.fetchall()
            scans_by_date = {}
            for scan in scans:
                d = scan['scan_timestamp'].date()
                if d not in scans_by_date: scans_by_date[d] = []
                scans_by_date[d].append(scan['scan_timestamp'])

            # (!!! SQL สำคัญ: ดึง total_amount มาด้วย !!!)
            cursor.execute("""
                SELECT work_date, status, ot_hours, trip_pickup, trip_crane, total_amount 
                FROM employee_daily_records 
                WHERE emp_id = %s AND work_date BETWEEN %s AND %s
            """, (emp_id, start_date, end_date))
            
            daily_records_dict = {}
            for row in cursor.fetchall():
                 daily_records_dict[row['work_date']] = dict(row)
                 
            # --- 4. วนลูปคำนวณ ---
            total_penalty_minutes = 0.0
            total_absent_days = 0.0
            actual_worked_days = 0.0 
            auto_driving_allowance = 0.0 
            
            all_dates = pd.date_range(start_date, end_date, freq='D')

            for current_date in all_dates:
                c_date = current_date.date()
                is_holiday = (current_date.weekday() == 6 or c_date in holiday_dates_set)
                
                scans_today = scans_by_date.get(c_date)
                leave_info = leave_records_dict.get(c_date)
                daily_rec = daily_records_dict.get(c_date)

                # (1) คำนวณค่าเที่ยว (!!! แก้ไข: ใช้ยอดเงินจริงจาก total_amount !!!)
                # เพราะถ้าใช้ (pickup*50) มันจะไม่รู้ว่าอันไหนฟรี
                if daily_rec:
                    trip_money = float(daily_rec.get('total_amount', 0) or 0)
                    auto_driving_allowance += trip_money
                
                if is_holiday: continue

                is_present = False
                if scans_today: is_present = True
                if daily_rec: is_present = True 
                if leave_info: is_present = True 

                if is_present:
                    actual_worked_days += 1.0
                
                effective_start = rules["standard_in"]
                effective_tier1 = rules["tier_1_cutoff"]
                num_leave = 0.0
                leave_type_today = None
                
                if leave_info:
                    num_leave = leave_info['days']
                    leave_type_today = leave_info.get('type')
                    if num_leave >= 1.0: continue 
                    if leave_info['start'] and leave_info['end']:
                         if leave_info['start'] <= effective_start <= leave_info['end']:
                             effective_start = leave_info['end']
                             effective_tier1 = (datetime.combine(datetime.today(), effective_start) + timedelta(minutes=30)).time()
                    elif num_leave < 1.0:
                        effective_start = (datetime.combine(datetime.today(), effective_start) + timedelta(hours=(rules["work_hours_per_day"] * num_leave))).time()
                        effective_tier1 = (datetime.combine(datetime.today(), effective_start) + timedelta(minutes=30)).time()

                if not scans_today:
                    if not daily_rec:
                        total_absent_days += (1.0 - num_leave)
                else:
                    t_in = min(scans_today).time()
                    t_out = max(scans_today).time()
                    total_dur = 0
                    if t_in != t_out:
                         dur_diff = datetime.combine(datetime.today(), t_out) - datetime.combine(datetime.today(), t_in)
                         total_dur = dur_diff.total_seconds() / 60
                    
                    break_deduct = 0
                    if (t_in < rules["break_end"] and t_out > rules["break_start"]):
                         ov_start = max(t_in, rules["break_start"])
                         ov_end = min(t_out, rules["break_end"])
                         break_deduct = (datetime.combine(datetime.today(), ov_end) - datetime.combine(datetime.today(), ov_start)).total_seconds() / 60
                         if break_deduct < 0: break_deduct = 0
                    actual_work = total_dur - break_deduct
                    
                    if t_in > effective_start:
                        if effective_tier1 and t_in > effective_tier1: total_penalty_minutes += rules["penalty_2_mins"]
                        else: total_penalty_minutes += rules["penalty_1_mins"]
                    elif (actual_work < rules["required_duration_minutes"]) and (leave_type_today != "ลากิจ"):
                        total_penalty_minutes += rules["penalty_1_mins"]

            # --- 5. สรุปรายได้ ---
            penalty_hours = total_penalty_minutes / 60.0
            result["debug_penalty_hours"] = penalty_hours
            result["debug_absent_days"] = total_absent_days
            result["debug_worked_days"] = actual_worked_days
            
            # บันทึกค่าเที่ยวรวม (ยอดเงินจริง)
            result["driving_allowance"] = auto_driving_allowance
            
            calculated_base_salary = 0.0
            late_deduction_amt = 0.0
            
            if salary_from_db > 0:
                if "รายวัน" in emp_type or "daily" in str(emp_type).lower():
                    # รายวัน
                    calculated_base_salary = salary_from_db * actual_worked_days
                    daily_rate = salary_from_db
                    hourly_rate = daily_rate / 8.0
                    late_deduction_amt = (hourly_rate * penalty_hours)
                    
                else:
                    # รายเดือน
                    calculated_base_salary = salary_from_db
                    daily_rate = salary_from_db / 30.0
                    hourly_rate = daily_rate / 8.0
                    late_deduction_amt = (daily_rate * total_absent_days) + (hourly_rate * penalty_hours)
            
            result["base_salary"] = calculated_base_salary
            result["late_deduct"] = late_deduction_amt
            result["ot"] = manual_ot 
            
            # --- 6. ประกันสังคม & สรุป ---
            sso_base = calculated_base_salary
            if sso_base < 1650: sso_base = 1650
            if sso_base > 15000: sso_base = 15000
            result["sso"] = round(sso_base * 0.05)

            result["total_income"] = (
                result["base_salary"] + position_allowance + 
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
'''

try:
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    print(f"Read {len(lines)} lines from {file_path}")

    # Find the start point: "    except Exception as e:" inside add_company_holiday
    # But wait, there are many "except Exception as e:".
    # I need to find the one that is followed by "        conn.rollback()" and then "    if not conn: return {}"
    
    start_index = -1
    for i in range(len(lines) - 2):
        if "    except Exception as e:" in lines[i] and \
           "        conn.rollback()" in lines[i+1] and \
           "    if not conn: return {}" in lines[i+2]:
            start_index = i
            print(f"Found start marker at line {i}")
            break
            
    # The end point is "    if not conn: return {}" itself (inclusive replacement)
    # Actually, I want to replace from line i (except) to line i+2 (if not conn)
    # And replace it with my big block which ends with "    if not conn: return {}"
    
    if start_index != -1:
        # We replace lines start_index to start_index + 3 (exclusive)
        # i.e. lines[i], lines[i+1], lines[i+2] are replaced.
        
        part1 = lines[:start_index]
        part3 = lines[start_index+3:] # Skip the 3 corrupted lines
        
        new_content = "".join(part1) + replacement_code + "\n" + "".join(part3)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
            
        print("Successfully repaired hr_database.py (v2)")
    else:
        print("Could not find the specific corruption pattern.")
        # Debug
        for i in range(len(lines) - 2):
            if "if not conn: return {}" in lines[i+2]:
                print(f"Near match at {i}: {lines[i].strip()} / {lines[i+1].strip()} / {lines[i+2].strip()}")

except Exception as e:
    print(f"Error: {e}")
