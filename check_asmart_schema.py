import psycopg2

# การตั้งค่าเชื่อมต่อ A+ Smart (ตามไฟล์ migrate_db.py)
ASMART_CONFIG = {
    "host": "192.168.1.51",  # IP ของเครื่อง A+ Smart
    "dbname": "aplus_com_test",
    "user": "app_user",
    "password": "cailfornia123",
    "port": "5432"
}

def check_columns():
    try:
        print(f"📡 กำลังเชื่อมต่อ A+ Smart ที่ {ASMART_CONFIG['host']}...")
        conn = psycopg2.connect(**ASMART_CONFIG)
        cursor = conn.cursor()
        
        # ดึงรายชื่อคอลัมน์จากตาราง commissions
        cursor.execute("""
            SELECT column_name, data_type 
            FROM information_schema.columns 
            WHERE table_name = 'commissions';
        """)
        columns = cursor.fetchall()
        
        print("\n✅ พบตาราง 'commissions' มีคอลัมน์ดังนี้:")
        print("-----------------------------------------")
        for col in columns:
            print(f" - {col[0]} ({col[1]})")
        print("-----------------------------------------")
        
        conn.close()
        
    except Exception as e:
        print(f"❌ เชื่อมต่อไม่ได้: {e}")

if __name__ == "__main__":
    check_columns()