import tkinter as tk
from tkinter import ttk, messagebox
import hr_database # Import เพื่อเรียกใช้ authenticate_user

class LoginScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent # เก็บ parent (MainApp) ไว้
        self.title("Login - Employee Management")
        self.geometry("350x200")
        self.resizable(False, False)
        self.grab_set() # ทำให้ focus อยู่ที่หน้าต่างนี้

        self.user_info = None # ตัวแปรเก็บข้อมูล user หลัง login สำเร็จ

        # --- UI Elements ---
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text="Username:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.username_entry = ttk.Entry(main_frame, width=30, font=("Segoe UI", 10))
        self.username_entry.grid(row=0, column=1, pady=5)

        ttk.Label(main_frame, text="Password:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.password_entry = ttk.Entry(main_frame, width=30, show="*", font=("Segoe UI", 10))
        self.password_entry.grid(row=1, column=1, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=15)
        ttk.Button(button_frame, text="Login", command=self._perform_login, width=15).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.destroy, width=15).pack(side="left", padx=5)

        self.username_entry.focus()
        # Bind Enter key to login
        self.bind("<Return>", lambda event: self._perform_login())

        # ทำให้หน้าต่างรอจนกว่าจะปิด
        self.wait_window(self)

    def _perform_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if not username or not password:
            messagebox.showwarning("Input Required", "Please enter both username and password.")
            return

        user_data = hr_database.authenticate_user(username, password)

        if user_data:
            self.user_info = user_data # เก็บข้อมูล user
            self.destroy() # ปิดหน้าต่าง login
        else:
            messagebox.showerror("Login Failed", "Invalid username or password.")
            self.password_entry.delete(0, tk.END) 