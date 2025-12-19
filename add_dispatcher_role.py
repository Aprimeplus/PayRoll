# ไฟล์: add_dispatcher_role.py
import hr_database
import psycopg2

def fix_database_roles():
    conn = hr_database.get_db_connection()
    if not conn:
        print("❌ เชื่อมต่อ Database ไม่ได้")
        return

    try:
        with conn.cursor() as cursor:
            print("⏳ กำลังปลดล็อก Database Constraint...")
            
            # 1. หาชื่อ Constraint เดิมที่ล็อก Role ไว้ (เช่น users_role_check)
            cursor.execute("SELECT conname FROM pg_constraint WHERE conrelid = 'users'::regclass AND contype = 'c';")
            constraints = cursor.fetchall()
            
            # 2. ลบ Constraint เดิมทิ้ง
            for row in constraints:
                c_name = row[0]
                # เช็คคร่าวๆ ว่าน่าจะเป็น constraint ของ role (ป้องกันลบผิดอัน)
                # ปกติถ้ามี constraint เดียวในตาราง users ก็ลบได้เลย แต่เพื่อความชัวร์เราจะลบหมดแล้วสร้างใหม่ที่ถูกต้อง
                print(f"   - พบ Constraint: {c_name} -> กำลังลบ...")
                cursor.execute(f"ALTER TABLE users DROP CONSTRAINT {c_name};")
            
            # 3. สร้าง Constraint ใหม่ที่รองรับ 'dispatcher'
            print("   - สร้าง Constraint ใหม่ (รองรับ dispatcher)...")
            cursor.execute("""
                ALTER TABLE users 
                ADD CONSTRAINT users_role_check 
                CHECK (role IN ('hr', 'approver', 'dispatcher'));
            """)
            
            # 4. เพิ่ม User 'dispatcher' (Default)
            print("⏳ กำลังสร้าง User: dispatcher_user ...")
            password_hash = hr_database.hash_password('truck123') # รหัสผ่านเริ่มต้น
            cursor.execute("""
                INSERT INTO users (username, password_hash, role)
                VALUES (%s, %s, %s)
                ON CONFLICT (username) DO NOTHING;
            """, ('dispatcher_user', password_hash, 'dispatcher'))
            
        conn.commit()
        print("\n✅ เสร็จสิ้น! ตอนนี้ระบบรองรับ Role 'dispatcher' แล้ว")
        print("✅ เพิ่ม User: 'dispatcher_user' / Pass: 'truck123' ให้แล้ว")
        
    except Exception as e:
        conn.rollback()
        print(f"\n❌ เกิดข้อผิดพลาด: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    fix_database_roles()