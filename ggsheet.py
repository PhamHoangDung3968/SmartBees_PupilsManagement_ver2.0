# import gspread

# gs = gspread.service_account("cre.json")
# sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# value = sht.sheet1.acell("B1").value
# # print(value)

# worksheet = sht.sheet1

# # Lấy dữ liệu từ ô A7 đến H15
# cell_range = "A7:H15"
# values = worksheet.get(cell_range)

# # In kết quả
# for row in values:
#     print(row)





# import gspread

# gs = gspread.service_account("cre.json")  # Assuming your credentials file is named "cre.json"
# sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
# worksheet = sht.sheet1

# # Get all values using `get_all_values()`
# all_values = worksheet.get_all_values()
# # Print the results
# for row in all_values:
#     print(row[:3])


import gspread

# Giả sử tệp chứng chỉ của bạn có tên là "cre.json"
gs = gspread.service_account("cre.json")

# Mở bảng tính bằng cách sử dụng khóa của nó
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")

# Lấy bảng tính đầu tiên (Sheet1)
worksheet = sht.sheet1

# Lấy tất cả giá trị từ cột A đến C
values_list = worksheet.get_all_values()
result_list = [row[:5] for row in values_list]

