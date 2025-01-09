import tkinter as tk
from tkinter import messagebox

def submit_input():
    name = name_entry.get()
    email = email_entry.get()
    if name and email:
        messagebox.showinfo("Input Received", f"Name: {name}\nEmail: {email}")
    else:
        messagebox.showerror("Error", "Please fill in all fields!")

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("ระบบจัดสลาก L6")
root.geometry("600x500")

# สร้าง Label และช่องกรอกสำหรับชื่อ
tk.Label(root, text="Name:").pack(pady=5)
name_entry = tk.Entry(root, width=30)
name_entry.pack(pady=5)

# สร้าง Label และช่องกรอกสำหรับอีเมล
tk.Label(root, text="Email:").pack(pady=5)
email_entry = tk.Entry(root, width=30)
email_entry.pack(pady=5)

# สร้างปุ่ม Submit
submit_button = tk.Button(root, text="Submit", command=submit_input)
submit_button.pack(pady=10)

root.mainloop()