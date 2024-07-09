import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from GUI.MainFormGUI import MainFormGUI

class LoginGUI:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Login")
        self.root.geometry("875x538")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.bg_image = Image.open("Images\\Bees.jpg")
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        self.bg_label = tk.Label(self.canvas, image=self.bg_photo)
        self.bg_label.place(x=0, y=0)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=415, y=29, width=430, height=444)
        self.panel = tk.Frame(self.root, bd=4, relief="solid")
        self.panel.place(x=415, y=29, width=430, height=444)
        self.lbl_login = tk.Label(self.panel, text="LOGIN", font=("cambria", 24, "bold"), fg="#FBA834")
        self.lbl_login.place(x=158, y=10)
        self.Title = tk.Label(self.canvas, text="SMART BEES", font=("cambria", 40, "bold"),fg="black",bg="#fffbc4")
        self.Title.place(x=55, y=220)
        self.lb1 = tk.Label(self.panel, text="UserName", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=55)
        self.tf1 = tk.Entry(self.panel)
        self.tf1.place(x=33, y=108, width=357, height=30)
        self.lb2 = tk.Label(self.panel, text="Password", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=33, y=171)
        self.tf2 = tk.Entry(self.panel, show="*")
        self.tf2.place(x=33, y=224, width=357, height=30)
        self.btn1 = tk.Button(self.panel, text="Login", font=("cambria", 14), command=self.login, width=20, bg="#FBA834",fg="white" )
        self.btn1.place(x=106, y=299)
    def login(self):
        username = self.tf1.get()
        password = self.tf2.get()
        if self.isValidCredentials(username, password):
            messagebox.showinfo("Success", "Đăng nhập thành công")
            self.root.destroy()
            # Add your logic to open the MainFormGUI here
            mainForm = MainFormGUI()
            mainForm.run()
            
        else:
            messagebox.showerror("Error", "Đăng nhập thất bại")

    def isValidCredentials(self, username, password):
        # Replace with your actual authentication logic
        return username == "admin" and password == "123"

    def run(self):
        self.root.mainloop()

