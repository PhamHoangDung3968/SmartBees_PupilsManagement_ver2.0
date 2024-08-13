# '''
# from tkinter import *

# def login():
#     # Xử lý đăng nhập ở đây
#     username = entry_username.get()
#     password = entry_password.get()
#     print(f"Đăng nhập với tên đăng nhập: {username} và mật khẩu: {password}")

# window = Tk()
# window.title("Đăng nhập")

# label_username = Label(window, text="Tên đăng nhập:")
# label_username.pack()
# entry_username = Entry(window)
# entry_username.pack()

# label_password = Label(window, text="Mật khẩu:")
# label_password.pack()
# entry_password = Entry(window, show="*")
# entry_password.pack()

# button_login = Button(window, text="Đăng nhập", command=login)
# button_login.pack()

# window.mainloop()
# '''


# # import tkinter as tk
# # from tkinter import messagebox
# # from tkinter import ttk
# # import gspread
# # import ezsheets
# # from openpyxl import Workbook
# # from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
# # from openpyxl.utils import get_column_letter

# # gs = gspread.service_account("cre.json")
# # sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
# # worksheet2 = sht.worksheet("sheet 2")
# # values_list_Score = worksheet2.get_all_values()[2:]  
# # result_list_Score = [row[:2] for row in values_list_Score] 
# # lop = [row[3] for row in values_list_Score]
# # combined_data = result_list_Score.copy()
# # for i in range(len(combined_data)):
# #     combined_data[i].append(lop[i])
# # teacher = [row[18] for row in values_list_Score]
# # combined_data1 = result_list_Score.copy() 
# # for i in range(len(combined_data1)):
# #     combined_data1[i].append(teacher[i])
# # listen = [row[19] for row in values_list_Score]
# # combined_data2 = result_list_Score.copy() 
# # for i in range(len(combined_data2)):
# #     combined_data2[i].append(listen[i])
# # speak = [row[20] for row in values_list_Score]
# # combined_data3 = result_list_Score.copy() 
# # for i in range(len(combined_data3)):
# #     combined_data3[i].append(speak[i])
# # rw = [row[21] for row in values_list_Score]
# # combined_data4 = result_list_Score.copy() 
# # for i in range(len(combined_data4)):
# #     combined_data4[i].append(rw[i])
# # total = [row[22] for row in values_list_Score]
# # combined_data5 = result_list_Score.copy() 
# # for i in range(len(combined_data5)):
# #     combined_data5[i].append(total[i])
# # ps = [row[23] for row in values_list_Score]
# # combined_data6 = result_list_Score.copy() 
# # for i in range(len(combined_data6)):
# #     combined_data6[i].append(ps[i])
# # print(combined_data6)






# # import gspread

# # # Assuming your credentials file is named "cre.json"
# # gs = gspread.service_account("cre.json")

# # # Open the spreadsheet using the sheet key
# # sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
# # worksheet = sht.sheet1

# # # Get the current data in columns A to F (modify these as needed)
# # existing_values = worksheet.get_all_values()

# # # Prepare the data you want to add to specific columns
# # new_data = [1, 2, 3, 4, 5]  # Assuming you want to add this data to columns B to F

# # # Determine the starting row where you want to insert the new data
# # # (adjust this based on your existing data)
# # start_row = len(existing_values) + 1  # Insert after the last row

# # # Efficiently update columns B to F (modify column indices as needed)
# # column_index = 15  # Start from column B (index 1)
# # for value in new_data:
# #     worksheet.update_cell(start_row, column_index, value)
# #     column_index += 1  # Move to the next column

# # print("Data added to columns B to F!")



# # import gspread

# # gs = gspread.service_account("cre.json")
# # sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
# # worksheet1 = sht.worksheet("sheet 3")
# # new_values = ["haha", "hehe", "New Value 3", "New Value 4"]

# # # Update the second row
# # worksheet1.update(values=[new_values], range_name='A9:D9')


# # Tạo một dictionary để lưu trữ ánh xạ từ A-Z sang 1-26
# char_to_num = dict()
# for i, c in enumerate('ABCDEFGHIJKLMNOPQRSTUVWXYZ'):
#     char_to_num[c] = i + 1

# # Ví dụ sử dụng
# char = 'C'
# num = char_to_num[char]
# # print(f"Ký tự '{char}' được gán số {num}")


# # Tạo một danh sách các chữ cái từ A đến Z
# # Tạo một danh sách các chữ cái từ A đến Z
# letters = [chr(i) for i in range(65, 91)]

# # Giá trị của n
# n = 30

# # Tạo một từ điển để gán số từ 1 đến n thành các chữ cái
# mapping = {}
# for i in range(1, n + 1):
#     # Sử dụng phép chia dư để lặp lại các chữ cái từ A đến Z
#     mapping[i] = letters[(i - 1) % 26]

# # In kết quả

# # Kiểm tra số 3 tương ứng với chữ cái gì
# print(f"Số 3 tương ứng với chữ cái: {mapping[3]}")

import string

n = 50  # Tăng giá trị n lên 50
letters = string.ascii_uppercase

# Tạo một từ điển để gán số từ 1 đến n thành các chữ cái
mapping = {}
for i in range(1, n + 1):
    # Sử dụng phép chia dư để lặp lại các chữ cái từ A đến Z
    first_letter = letters[(i - 1) // 26 - 1] if i > 26 else ''
    second_letter = letters[(i - 1) % 26]
    mapping[i] = first_letter + second_letter

# Kiểm tra số 3 và số 50 tương ứng với chữ cái gì
print(f"Số 3 tương ứng với chữ cái: {mapping[3]}")
print(f"Số 50 tương ứng với chữ cái: {mapping[50]}")

