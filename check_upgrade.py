import psycopg2

# ตั้งค่า Connection
conn = psycopg2.connect(
    dbname="aplus_com_test", 
    user="app_user", 
    password="cailfornia123", 
    host="Server-APrime", 
    port="5432"
)
cursor = conn.cursor()

print("--- 🔍 ตรวจสอบโครงสร้างฐานข้อมูล ---")

# เช็กคอลัมน์ start_date ในตาราง employees
cursor.execute("""
    SELECT column_name, data_type 
    FROM information_schema.columns 
    WHERE table_name = 'employees' AND column_name = 'start_date';
""")
col_info = cursor.fetchone()
print(f"1. คอลัมน์ 'start_date' เป็นชนิด:  {col_info[1].upper()}") 
# (ต้องขึ้นว่า DATE)

# เช็กข้อมูลจริง 1 แถว
cursor.execute("SELECT emp_id, start_date FROM employees LIMIT 1")
data = cursor.fetchone()
print(f"2. ตัวอย่างข้อมูลดิบใน DB:      {data[1]} (Type: {type(data[1])})")
# (ต้องเป็น 20xx-xx-xx และ Type เป็น datetime.date)

conn.close()
print("\n----------------------------------")
if col_info[1].upper() == 'DATE':
    print("✅ ยืนยัน: ระบบอัปเกรดเป็น DATE สมบูรณ์แล้ว!")
else:
    print("❌ ยังไม่ได้อัปเกรด (ยังเป็น TEXT อยู่)")