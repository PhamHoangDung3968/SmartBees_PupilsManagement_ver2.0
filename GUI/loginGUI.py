import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

from GUI.MainFormGUI import MainFormGUI

class LoginGUI:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Login")
        #self.root.geometry("875x538")
        
        # Center the window on the screen
        self.center_window(875,538)
        
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
        self.tf1 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=357, height=30)
        self.lb2 = tk.Label(self.panel, text="Password", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=33, y=171)
        '''
        self.tf2 = tk.Entry(self.panel, show="*", font=("cambria", 13, "bold"))
        self.tf2.place(x=33, y=224, width=357, height=30)
        '''
        # Entry cho mật khẩu với ký tự ẩn
        self.tf2 = tk.Entry(self.panel, show="*", font=("Cambria", 13, "bold"))
        self.tf2.place(x=33, y=224, width=357, height=30)

        # Tạo một biến để kiểm tra trạng thái của checkbox
        self.show_password_var = tk.BooleanVar()

        # Checkbutton để bật/tắt hiển thị mật khẩu
        self.show_password_cb = tk.Checkbutton(self.panel, text="Hiện mật khẩu", variable=self.show_password_var,
                                               onvalue=True, offvalue=False, command=self.toggle_password)
        self.show_password_cb.place(x=33, y=260)
        
        
        self.btn1 = tk.Button(self.panel, text="Login", font=("cambria", 14), command=self.login, width=20, bg="#FBA834",fg="white" )
        self.btn1.place(x=106, y=299)
        # Gắn sự kiện nhấn phím Enter với hàm Add_Class
        self.root.bind('<Return>', self.on_enter_key)
        
    # Hàm bật/tắt hiển thị mật khẩu
    def toggle_password(self):
        if self.show_password_var.get():
            self.tf2.config(show="")  # Hiển thị mật khẩu
        else:
            self.tf2.config(show="*")  # Ẩn mật khẩu
            
            
    def center_window(self, width, height):
        window_width = width
        window_height = height

        # Lấy kích thước màn hình
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Tính toán vị trí x và y để cửa sổ xuất hiện ở giữa màn hình
        position_x = int((screen_width / 2) - (window_width / 2))
        position_y = int((screen_height / 2) - (window_height / 2))

        # Đặt lại vị trí cho cửa sổ
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")


    def login(self):
        username = self.tf1.get()
        password = self.tf2.get()
        if self.isValidCredentials(username, password):
            self.root.destroy()
            # Add your logic to open the MainFormGUI here
            mainForm = MainFormGUI()
            mainForm.run()
            
        else:
            messagebox.showerror("Error", "Đăng nhập thất bại")

    def isValidCredentials(self, username, password):
        # Replace with your actual authentication logic
        return username == "admin" and password == "123"

    def on_enter_key(self, event):
        # Gọi cùng hàm xử lý khi phím Enter được nhấn
        self.login()
        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()