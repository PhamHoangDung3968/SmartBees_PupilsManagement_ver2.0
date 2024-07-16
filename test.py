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


# import tkinter as tk
# from tkinter import messagebox
# from tkinter import ttk
# import gspread
# import ezsheets
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
# from openpyxl.utils import get_column_letter

# gs = gspread.service_account("cre.json")
# sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# worksheet2 = sht.worksheet("sheet 2")
# values_list_Score = worksheet2.get_all_values()[2:]  
# result_list_Score = [row[:2] for row in values_list_Score] 
# lop = [row[3] for row in values_list_Score]
# combined_data = result_list_Score.copy()
# for i in range(len(combined_data)):
#     combined_data[i].append(lop[i])
# teacher = [row[18] for row in values_list_Score]
# combined_data1 = result_list_Score.copy() 
# for i in range(len(combined_data1)):
#     combined_data1[i].append(teacher[i])
# listen = [row[19] for row in values_list_Score]
# combined_data2 = result_list_Score.copy() 
# for i in range(len(combined_data2)):
#     combined_data2[i].append(listen[i])
# speak = [row[20] for row in values_list_Score]
# combined_data3 = result_list_Score.copy() 
# for i in range(len(combined_data3)):
#     combined_data3[i].append(speak[i])
# rw = [row[21] for row in values_list_Score]
# combined_data4 = result_list_Score.copy() 
# for i in range(len(combined_data4)):
#     combined_data4[i].append(rw[i])
# total = [row[22] for row in values_list_Score]
# combined_data5 = result_list_Score.copy() 
# for i in range(len(combined_data5)):
#     combined_data5[i].append(total[i])
# ps = [row[23] for row in values_list_Score]
# combined_data6 = result_list_Score.copy() 
# for i in range(len(combined_data6)):
#     combined_data6[i].append(ps[i])
# print(combined_data6)






# import gspread

# # Assuming your credentials file is named "cre.json"
# gs = gspread.service_account("cre.json")

# # Open the spreadsheet using the sheet key
# sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# worksheet = sht.sheet1

# # Get the current data in columns A to F (modify these as needed)
# existing_values = worksheet.get_all_values()

# # Prepare the data you want to add to specific columns
# new_data = [1, 2, 3, 4, 5]  # Assuming you want to add this data to columns B to F

# # Determine the starting row where you want to insert the new data
# # (adjust this based on your existing data)
# start_row = len(existing_values) + 1  # Insert after the last row

# # Efficiently update columns B to F (modify column indices as needed)
# column_index = 15  # Start from column B (index 1)
# for value in new_data:
#     worksheet.update_cell(start_row, column_index, value)
#     column_index += 1  # Move to the next column

# print("Data added to columns B to F!")



import gspread
# gs = gspread.service_account("cre.json")
# sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# worksheet = sht.worksheet("sheet 2")
# # Thêm một hàng mới vào cuối bảng tính
# test = worksheet.get_all_values()
# end_col = len([row[1] for row in test] )
# x= end_col-2+1
# new_row_values = [x, 2, 3, 4, 5]  # Giá trị của hàng mới
# worksheet.append_row(new_row_values, value_input_option='RAW')
# values_list_Book = worksheet.get_all_values()[2:]
# result_list_Book2 = [row[:5] for row in values_list_Book]

# print(result_list_Book2)



# import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# worksheet = sht.worksheet("sheet 3")
# # Thêm một hàng mới vào cuối bảng tính
# test = worksheet.get_all_values()
# end_col = len([row[1] for row in test] )
# x= end_col-2+1
# new_row_values = [x, 2, 3, 4, 5]  # Giá trị của hàng mới
# worksheet.append_row(new_row_values, value_input_option='RAW')
# values_list_Book = worksheet.get_all_values()[2:]
# result_list_Book2 = [row[:5] for row in values_list_Book]

# print(result_list_Book2)
worksheet3 = sht.worksheet("sheet 3")
test = worksheet3.get_all_values()[2:]
row_values =worksheet3.row_values(3)
print(row_values)


