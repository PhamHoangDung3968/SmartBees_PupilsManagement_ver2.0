'''
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
'''


import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import gspread
import ezsheets
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
worksheet = sht.sheet1

values_list_Score = worksheet.get_all_values()[2:]  
result_list_Score = [row[22:24] for row in values_list_Score] 
lop = [row[25] for row in values_list_Score]
combined_data = result_list_Score.copy()
for i in range(len(combined_data)):
    combined_data[i].append(lop[i])
teacher = [row[40] for row in values_list_Score]
combined_data1 = result_list_Score.copy() 
for i in range(len(combined_data1)):
    combined_data1[i].append(teacher[i])
listen = [row[41] for row in values_list_Score]
combined_data2 = result_list_Score.copy() 
for i in range(len(combined_data2)):
    combined_data2[i].append(listen[i])
speak = [row[42] for row in values_list_Score]
combined_data3 = result_list_Score.copy() 
for i in range(len(combined_data3)):
    combined_data3[i].append(speak[i])
rw = [row[43] for row in values_list_Score]
combined_data4 = result_list_Score.copy() 
for i in range(len(combined_data4)):
    combined_data4[i].append(rw[i])
total = [row[44] for row in values_list_Score]
combined_data5 = result_list_Score.copy() 
for i in range(len(combined_data5)):
    combined_data5[i].append(total[i])
ps = [row[45] for row in values_list_Score]
combined_data6 = result_list_Score.copy() 
for i in range(len(combined_data6)):
    combined_data6[i].append(ps[i])



test = worksheet.get_all_values()
end_col = len([row[15] for row in test] )
print(end_col)