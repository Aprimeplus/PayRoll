# (ไฟล์: fix_employees_only.py)
import psycopg2
import sys

# ตั้งค่า Connection
DB_CONFIG = {
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "host": "192.168.1.51",
    "port": "5432"
}

def run_fix():
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        conn.autocommit = True # (สำคัญ: สั่งให้บันทึกทันทีทีละคำสั่ง)
        cursor = conn.cursor()
        
        print("--- 🚀 เริ่มซ่อมแซมตาราง Employees ---")
        
        # รายชื่อคอลัมน์วันที่ในตาราง employees
        cols_to_fix = [
            "birth_date", "start_date", "termination_date", 
            "sso_start_date", "sso_end_date", 
            "sso_start_action_date", "sso_end_action_date",
            "probation_end_date"
        ]

        for col in cols_to_fix:
            print(f"กำลังแปลง {col} เป็น DATE ... ", end="")
            try:
                # 1. แปลงจาก TEXT เป็น DATE (และจัดการค่าว่าง)
                sql = f"""
                    ALTER TABLE employees 
                    ALTER COLUMN {col} TYPE DATE 
                    USING CASE 
                        WHEN TRIM({col}) = '' OR {col} IS NULL THEN NULL
                        ELSE to_date(TRIM({col}), 'DD/MM/YYYY') - interval '543 years'
                    END
                """
                cursor.execute(sql)
                print("✅ สำเร็จ!")
            except Exception as e:
                if "cannot be cast automatically" in str(e) or "column" in str(e):
                     # ถ้ามันเป็น Date อยู่แล้ว หรือแปลงไม่ได้
                     print(f"⚠️ (ข้าม: {e})")
                else:
                     print(f"❌ Error: {e}")

        print("\n✅ เสร็จสิ้น! ลองเข้าโปรแกรมใหม่อีกครั้งครับ")
        
    except Exception as e:
        print(f"\n❌ การเชื่อมต่อล้มเหลว: {e}")
    finally:
        if 'conn' in locals() and conn: conn.close()

if __name__ == "__main__":
    run_fix()