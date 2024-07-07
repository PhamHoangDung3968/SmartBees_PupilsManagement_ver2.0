from tkinter import *

def login():
    # Xử lý đăng nhập ở đây
    username = entry_username.get()
    password = entry_password.get()
    print(f"Đăng nhập với tên đăng nhập: {username} và mật khẩu: {password}")

window = Tk()
window.title("Đăng nhập")

label_username = Label(window, text="Tên đăng nhập:")
label_username.pack()
entry_username = Entry(window)
entry_username.pack()

label_password = Label(window, text="Mật khẩu:")
label_password.pack()
entry_password = Entry(window, show="*")
entry_password.pack()

button_login = Button(window, text="Đăng nhập", command=login)
button_login.pack()

window.mainloop()
