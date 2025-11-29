import tkinter as tk
from tkinter import ttk
import hr_database

class AuditLogViewer(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("📝 บันทึกการแก้ไขข้อมูล (Audit Trail)")
        self.geometry("1000x600")
        
        self._build_ui()
        self._load_logs()
        
    def _build_ui(self):
        # Header
        ttk.Label(self, text="ประวัติการแก้ไขข้อมูลสำคัญ (Sensitive Data Changes)", 
                  font=("Segoe UI", 14, "bold"), foreground="#c0392b").pack(pady=10)
        
        # Table
        cols = ("time", "actor", "action", "emp", "field", "old", "new")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=20)
        
        self.tree.heading("time", text="เวลาที่แก้")
        self.tree.heading("actor", text="ผู้แก้ไข")
        self.tree.heading("action", text="การกระทำ")
        self.tree.heading("emp", text="พนักงานที่ถูกแก้")
        self.tree.heading("field", text="ข้อมูลที่แก้")
        self.tree.heading("old", text="ค่าเดิม")
        self.tree.heading("new", text="ค่าใหม่")
        
        self.tree.column("time", width=150)
        self.tree.column("actor", width=100)
        self.tree.column("action", width=80)
        self.tree.column("emp", width=150)
        self.tree.column("field", width=120)
        self.tree.column("old", width=150)
        self.tree.column("new", width=150)
        
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Scrollbar
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        sb.place(relx=1, rely=0, relheight=1, anchor="ne")
        self.tree.configure(yscrollcommand=sb.set)
        
    def _load_logs(self):
        conn = hr_database.get_db_connection()
        if not conn: return
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    SELECT action_time, actor_name, action_type, target_emp_name, 
                           changed_field, old_value, new_value 
                    FROM audit_logs 
                    ORDER BY action_time DESC
                    LIMIT 100
                """)
                rows = cursor.fetchall()
                for row in rows:
                    # จัดรูปแบบเวลาให้สวยงาม
                    ts = row[0].strftime("%d/%m/%Y %H:%M")
                    self.tree.insert("", "end", values=(ts, *row[1:]))
        finally:
            conn.close()
if __name__ == "__main__":
    # จำลองการทำงานเมื่อรันไฟล์นี้โดดๆ
    root = tk.Tk()
    root.withdraw() # ซ่อนหน้าต่างแม่ตัวเปล่าๆ ไว้

    try:
        # เรียกเปิดหน้าต่าง Audit Log
        app = AuditLogViewer(root)
        
        # สั่งให้ปิดโปรแกรมจริงเมื่อปิดหน้าต่างนี้
        app.protocol("WM_DELETE_WINDOW", lambda: root.destroy())
        
        root.mainloop()
    except Exception as e:
        print(f"Error: {e}")
# วิธีใช้: เรียก AuditLogViewer(self) จากหน้า Main หรือ Dashboard