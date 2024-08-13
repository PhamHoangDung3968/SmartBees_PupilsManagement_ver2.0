# # import gspread

# # gs = gspread.service_account("cre.json")
# # sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
# # value = sht.sheet1.acell("B1").value
# # # print(value)

# # worksheet = sht.sheet1

# # # Lấy dữ liệu từ ô A7 đến H15
# # cell_range = "A7:H15"
# # values = worksheet.get(cell_range)

# # # In kết quả
# # for row in values:
# #     print(row)


# '''
# import ezsheets
# ss = ezsheets.Spreadsheet("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
# ss.downloadAsExcel()
# '''


# import ezsheets
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
# from openpyxl.utils import get_column_letter
# from datetime import datetime

# # Function to get data from a specified range
# def get_data_from_range(sheet, start_row, end_row, start_col, end_col):
#     data = []
#     for row in range(start_row, end_row + 1):
#         row_data = []
#         for col in range(start_col, end_col + 1):
#             row_data.append(sheet[row, col])
#         data.append(row_data)
#     return data

# # Download the specific Google Sheet
# ss = ezsheets.Spreadsheet("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")

# # Specify the sheet, columns, and rows
# # lag
# sheet_name = 'sheet 1'  # Change this to the specific sheet name
# start_row = 1  # Skip the header row
# end_row = 13
# start_col = 3
# end_col = 18

# sheet = ss[sheet_name]
# data = get_data_from_range(sheet, start_row, end_row, start_col, end_col)

# # Create a new Excel file and write data to it
# wb = Workbook()
# ws = wb.active

# # Write headers and data to the new Excel sheet
# headers = [
#     "ID_BOOK", "CAMBRIDGE LEVEL", "PROGRESS", "MAIN BOOK", "SKILL BOOK 1",
#     "VOCAB BOOK", "SKILL BOOK 2", "SKILL BOOK 3", "SKILL BOOK 4",
#     "GRAMMAR BOOK", "TEST BOOK", "VIDEOS-MOVIES", "PICTURES-CARDS"
# ]

# # Define styles
# header_font = Font(bold=True, color="000000")
# header_alignment = Alignment(horizontal='center', vertical='center')
# thin_border = Border(
#     left=Side(style='thin'), 
#     right=Side(style='thin'), 
#     top=Side(style='thin'), 
#     bottom=Side(style='thin')
# )
# title_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
# title_font = Font(bold=True, size=14)

# # Write the title and apply styles
# title_cell = ws.cell(row=1, column=1, value="Quản Lý Sách")
# title_cell.font = title_font
# title_cell.fill = title_fill
# title_cell.alignment = header_alignment
# ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)

# # Write the headers and apply styles
# for col_idx, header in enumerate(headers, start=1):
#     cell = ws.cell(row=2, column=col_idx, value=header)
#     cell.font = header_font
#     cell.alignment = header_alignment
#     cell.border = thin_border

# # Write the data and apply borders
# for col_idx, col_data in enumerate(data, start=1):
#     for row_idx, value in enumerate(col_data, start=3):
#         cell = ws.cell(row=row_idx, column=col_idx, value=value)
#         cell.border = thin_border

# # Adjust column widths to fit the content
# for col in ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
#     max_length = 0
#     column = get_column_letter(col[0].column)  # Get the column name
#     for cell in col:
#         if cell.value is not None:
#             try:
#                 if len(str(cell.value)) > max_length:
#                     max_length = len(cell.value)
#             except:
#                 pass
#     adjusted_width = (max_length + 2)
#     ws.column_dimensions[column].width = adjusted_width

# # Generate unique file name with date and time
# current_time = datetime.now().strftime("%d-%m-%Y-%H-%M-%S")
# file_path = f'QLHS-{current_time}.xlsx'

# # Save the new Excel file
# wb.save(file_path)
# print(f"Data copied to {file_path}")
# gs = gspread.service_account("cre.json")
#     sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
#     worksheet = sht.sheet1

#     # Show data
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
worksheet2 = sht.worksheet("sheet 2")
# giai đoạn 1
values_list_Student = worksheet2.get_all_values()[2:]
result_list_Student = [row[:5] for row in values_list_Student]
tel = [row[7] for row in values_list_Student]
combined_data_student = result_list_Student.copy()
for i in range(len(combined_data_student)):
    combined_data_student[i].append(tel[i])
diachi = [row[8] for row in values_list_Student]
combined_data_student1 = result_list_Student.copy()
for i in range(len(combined_data_student1)):
    combined_data_student1[i].append(diachi[i])

print(combined_data_student1)

# lop = [row[3] for row in values_list_Score]
# combined_data = result_list_Score.copy()
# for i in range(len(combined_data)):
#     combined_data[i].append(lop[i])

# teacher = [row[21] for row in values_list_Score]
# combined_data1 = result_list_Score.copy()
# for i in range(len(combined_data1)):
#     combined_data1[i].append(teacher[i])

# listen1 = [row[25] for row in values_list_Score]
# combined_data2 = result_list_Score.copy()
# for i in range(len(combined_data2)):
#     combined_data2[i].append(listen1[i])

# speak1 = [row[26] for row in values_list_Score]
# combined_data3 = result_list_Score.copy()
# for i in range(len(combined_data3)):
#     combined_data3[i].append(speak1[i])

# rw1 = [row[27] for row in values_list_Score]
# combined_data4 = result_list_Score.copy()
# for i in range(len(combined_data4)):
#     combined_data4[i].append(rw1[i])

# total1 = [row[28] for row in values_list_Score]
# combined_data5 = result_list_Score.copy()
# for i in range(len(combined_data5)):
#     combined_data5[i].append(total1[i])

# ps1 = [row[29] for row in values_list_Score]
# combined_data6 = result_list_Score.copy()
# for i in range(len(combined_data6)):
#     combined_data6[i].append(ps1[i])

# #giai đoạn 2
# values_list_Score2 = worksheet2_1.get_all_values()[2:]
# result_list_Score2 = [row[:2] for row in values_list_Score2]
# lop2 = [row[3] for row in values_list_Score2]
# combined_data_2 = result_list_Score2.copy()
# for i in range(len(combined_data_2)):
#     combined_data_2[i].append(lop2[i])

# teacher2 = [row[21] for row in values_list_Score2]
# combined_data1_2 = result_list_Score2.copy()
# for i in range(len(combined_data1_2)):
#     combined_data1_2[i].append(teacher2[i])
    
# listen2 = [row[30] for row in values_list_Score2]
# combined_data2_2 = result_list_Score2.copy()
# for i in range(len(combined_data2_2)):
#     combined_data2_2[i].append(listen2[i])

# speak2 = [row[31] for row in values_list_Score2]
# combined_data3_2 = result_list_Score2.copy()
# for i in range(len(combined_data3_2)):
#     combined_data3_2[i].append(speak2[i])

# rw2 = [row[32] for row in values_list_Score2]
# combined_data4_2 = result_list_Score2.copy()
# for i in range(len(combined_data4_2)):
#     combined_data4_2[i].append(rw2[i])

# total2 = [row[33] for row in values_list_Score2]
# combined_data5_2 = result_list_Score2.copy()
# for i in range(len(combined_data5_2)):
#     combined_data5_2[i].append(total2[i])

# ps2 = [row[34] for row in values_list_Score2]
# combined_data6_2 = result_list_Score2.copy()
# for i in range(len(combined_data6_2)):
#     combined_data6_2[i].append(ps2[i])


# #giai đoạn 3
# values_list_Score3 = worksheet2_1.get_all_values()[2:]
# result_list_Score3 = [row[:2] for row in values_list_Score3]
# lop3 = [row[3] for row in values_list_Score3]
# combined_data_3 = result_list_Score3.copy()
# for i in range(len(combined_data_3)):
#     combined_data_3[i].append(lop3[i])

# teacher3 = [row[21] for row in values_list_Score3]
# combined_data1_3 = result_list_Score3.copy()
# for i in range(len(combined_data1_3)):
#     combined_data1_3[i].append(teacher3[i])
    
# listen3 = [row[35] for row in values_list_Score3]
# combined_data2_3 = result_list_Score3.copy()
# for i in range(len(combined_data2_3)):
#     combined_data2_3[i].append(listen3[i])

# speak3 = [row[36] for row in values_list_Score3]
# combined_data3_3 = result_list_Score3.copy()
# for i in range(len(combined_data3_3)):
#     combined_data3_3[i].append(speak3[i])

# rw3 = [row[37] for row in values_list_Score3]
# combined_data4_3 = result_list_Score3.copy()
# for i in range(len(combined_data4_3)):
#     combined_data4_3[i].append(rw3[i])

# total3 = [row[38] for row in values_list_Score3]
# combined_data5_3 = result_list_Score3.copy()
# for i in range(len(combined_data5_3)):
#     combined_data5_3[i].append(total3[i])

# ps3 = [row[39] for row in values_list_Score3]
# combined_data6_3 = result_list_Score3.copy()
# for i in range(len(combined_data6_3)):
#     combined_data6_3[i].append(ps3[i])


# #giai đoạn 4
# values_list_Score4 = worksheet2_1.get_all_values()[2:]
# result_list_Score4 = [row[:2] for row in values_list_Score4]
# lop4 = [row[3] for row in values_list_Score4]
# combined_data_4 = result_list_Score4.copy()
# for i in range(len(combined_data_4)):
#     combined_data_4[i].append(lop4[i])

# teacher4 = [row[21] for row in values_list_Score4]
# combined_data1_4 = result_list_Score4.copy()
# for i in range(len(combined_data1_4)):
#     combined_data1_4[i].append(teacher4[i])
    
# listen4 = [row[40] for row in values_list_Score4]
# combined_data2_4 = result_list_Score4.copy()
# for i in range(len(combined_data2_4)):
#     combined_data2_4[i].append(listen4[i])

# speak4 = [row[41] for row in values_list_Score4]
# combined_data3_4 = result_list_Score4.copy()
# for i in range(len(combined_data3_4)):
#     combined_data3_4[i].append(speak4[i])

# rw4 = [row[42] for row in values_list_Score4]
# combined_data4_4 = result_list_Score4.copy()
# for i in range(len(combined_data4_4)):
#     combined_data4_4[i].append(rw4[i])

# total4 = [row[43] for row in values_list_Score4]
# combined_data5_4 = result_list_Score4.copy()
# for i in range(len(combined_data5_4)):
#     combined_data5_4[i].append(total4[i])

# ps4 = [row[44] for row in values_list_Score4]
# combined_data6_4 = result_list_Score4.copy()
# for i in range(len(combined_data6_4)):
#     combined_data6_4[i].append(ps4[i])

# #giai đoạn 5
# values_list_Score5 = worksheet2_1.get_all_values()[2:]
# result_list_Score5 = [row[:2] for row in values_list_Score5]
# lop5 = [row[3] for row in values_list_Score5]
# combined_data_5 = result_list_Score5.copy()
# for i in range(len(combined_data_5)):
#     combined_data_5[i].append(lop5[i])

# teacher5 = [row[21] for row in values_list_Score5]
# combined_data1_5 = result_list_Score5.copy()
# for i in range(len(combined_data1_5)):
#     combined_data1_5[i].append(teacher5[i])
    
# listen5 = [row[45] for row in values_list_Score5]
# combined_data2_5 = result_list_Score5.copy()
# for i in range(len(combined_data2_5)):
#     combined_data2_5[i].append(listen5[i])

# speak5 = [row[46] for row in values_list_Score5]
# combined_data3_5 = result_list_Score5.copy()
# for i in range(len(combined_data3_5)):
#     combined_data3_5[i].append(speak5[i])

# rw5 = [row[47] for row in values_list_Score5]
# combined_data4_5 = result_list_Score5.copy()
# for i in range(len(combined_data4_5)):
#     combined_data4_5[i].append(rw5[i])

# total5 = [row[48] for row in values_list_Score5]
# combined_data5_5 = result_list_Score5.copy()
# for i in range(len(combined_data5_5)):
#     combined_data5_5[i].append(total5[i])

# ps5 = [row[49] for row in values_list_Score5]
# combined_data6_5 = result_list_Score5.copy()
# for i in range(len(combined_data6_5)):
#     combined_data6_5[i].append(ps5[i])
# print(combined_data6_5)