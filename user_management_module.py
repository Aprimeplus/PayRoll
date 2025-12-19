# (ไฟล์: user_management_module.py)

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import hr_database

class UserManagementModule(ttk.Frame):
    def __init__(self, parent, controller, current_user):
        super().__init__(parent)
        self.controller = controller
        self.current_user = current_user
        
        self._build_ui()
        self._load_users()

    def _build_ui(self):
        # Header
        header_frame = ttk.Frame(self, padding=20)
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text="🔐 จัดการผู้ใช้งานระบบ (User Management)", 
                  style="Header.TLabel").pack(side="left")

        # Content (Split 2 Sides)
        content = ttk.Frame(self, padding=(20, 0))
        content.pack(fill="both", expand=True)

        # --- Left: User List ---
        list_frame = ttk.LabelFrame(content, text=" รายชื่อผู้ใช้งาน ", padding=10)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        columns = ("id", "username", "role")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        self.tree.heading("id", text="ID")
        self.tree.heading("username", text="Username")
        self.tree.heading("role", text="Role (สิทธิ์)")
        
        self.tree.column("id", width=50, anchor="center")
        self.tree.column("username", width=150, anchor="w")
        self.tree.column("role", width=100, anchor="center")
        
        self.tree.pack(fill="both", expand=True)

        # ปุ่มจัดการ (ใต้ตาราง)
        btn_frame = ttk.Frame(list_frame, padding=(0, 10))
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="🔑 รีเซ็ตรหัสผ่าน", command=self._reset_password).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🗑️ ลบผู้ใช้", command=self._delete_user).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🔄 โหลดข้อมูล", command=self._load_users).pack(side="right", padx=5)

        # --- Right: Add New User Form ---
        form_frame = ttk.LabelFrame(content, text=" เพิ่มผู้ใช้ใหม่ ", padding=20)
        form_frame.pack(side="right", fill="y", padx=(10, 0))

        ttk.Label(form_frame, text="Username:").pack(anchor="w", pady=5)
        self.new_user_entry = ttk.Entry(form_frame, width=25)
        self.new_user_entry.pack(anchor="w")

        ttk.Label(form_frame, text="Password:").pack(anchor="w", pady=5)
        self.new_pass_entry = ttk.Entry(form_frame, width=25, show="*")
        self.new_pass_entry.pack(anchor="w")

        ttk.Label(form_frame, text="Role (สิทธิ์):").pack(anchor="w", pady=5)
        
        # --- (แก้ไขตรงนี้: เพิ่ม 'dispatcher') ---
        self.role_combo = ttk.Combobox(form_frame, values=["hr", "approver", "dispatcher"], state="readonly", width=22)
        # -------------------------------------
        
        self.role_combo.set("hr")
        self.role_combo.pack(anchor="w")

        ttk.Button(form_frame, text="💾 สร้างผู้ใช้ใหม่", 
                   command=self._add_user, style="Success.TButton").pack(fill="x", pady=20)

    def _load_users(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        users = hr_database.get_all_users()
        for u in users:
            self.tree.insert("", "end", values=(u['user_id'], u['username'], u['role']))

    def _add_user(self):
        username = self.new_user_entry.get().strip()
        password = self.new_pass_entry.get().strip()
        role = self.role_combo.get()

        if not username or not password:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก Username และ Password")
            return

        if messagebox.askyesno("ยืนยัน", f"ต้องการสร้างผู้ใช้ '{username}' ({role}) ใช่หรือไม่?"):
            success = hr_database.add_new_user(username, password, role)
            if success:
                messagebox.showinfo("สำเร็จ", "สร้างผู้ใช้ใหม่เรียบร้อยแล้ว")
                self.new_user_entry.delete(0, tk.END)
                self.new_pass_entry.delete(0, tk.END)
                self._load_users()

    def _delete_user(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("เลือกรายการ", "กรุณาเลือกผู้ใช้ที่ต้องการลบ")
            return
        
        item = self.tree.item(selection[0])
        user_id = item['values'][0]
        username = item['values'][1]

        # ป้องกันการลบตัวเอง
        if str(user_id) == str(self.current_user['user_id']):
            messagebox.showerror("ห้ามทำ", "คุณไม่สามารถลบ Account ของตัวเองได้")
            return

        if messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบผู้ใช้ '{username}' ถาวรหรือไม่?"):
            if hr_database.delete_user(user_id):
                messagebox.showinfo("สำเร็จ", "ลบผู้ใช้เรียบร้อย")
                self._load_users()

    def _reset_password(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("เลือกรายการ", "กรุณาเลือกผู้ใช้ที่ต้องการรีเซ็ตรหัส")
            return
        
        item = self.tree.item(selection[0])
        user_id = item['values'][0]
        username = item['values'][1]

        new_pass = simpledialog.askstring("เปลี่ยนรหัสผ่าน", f"ตั้งรหัสผ่านใหม่สำหรับ '{username}':", show='*')
        
        if new_pass:
            if hr_database.update_user_password(user_id, new_pass):
                messagebox.showinfo("สำเร็จ", f"เปลี่ยนรหัสผ่านให้ '{username}' เรียบร้อยแล้ว")