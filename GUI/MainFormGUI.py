import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import gspread
import ezsheets
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
from Add_NewClass import Add_NewClass



gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
worksheet = sht.sheet1
values_list_Book = worksheet.get_all_values()[2:]
result_list_Book = [row[:5] for row in values_list_Book]

values_list_Student = worksheet.get_all_values()[2:]
result_list_Student = [row[22:29] for row in values_list_Student]

values_list_Class = worksheet.get_all_values()[2:]
result_list_Class = [row[14:20] for row in values_list_Class]

values_list_Score = worksheet.get_all_values()[2:]
result_list_Score = [row[22:21] for row in values_list_Class]



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

class MainFormGUI:
    def __init__(self):
        self.root = tk.Tk()
        
        # Root window properties
        self.root.title("Main Form GUI")
        self.root.geometry("1097x700")
        self.root.configure(bg="#e0f7fa")
        
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#e6e6e6')
        self.style.configure('TButton', background='#cc0000', foreground='#cc0000', font=('Cambria', 12, 'bold'))
        
        self.style.configure('TLabel', background='#e6e6e6', foreground='#007acc', font=('Cambria', 12, 'bold'))
        
        self.style.configure('TEntry', background='#007acc', foreground='#007acc', font=('Cambria', 12))
        
        self.style.configure('TNotebook.Tab', font=('Cambria', 14, 'bold'), background='#007acc', foreground='#007acc')
        
        self.style.configure('TTreeview.Heading', font=('Cambria', 11, 'bold'), background='#007ACC', foreground='white')
        
        self.style.configure('TTreeview', font=('Cambria', 11), background='#f5f5f5', foreground='#333333')

        # Main content frame
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True)
        self.tab_control = ttk.Notebook(self.content_frame)
        self.tab_control.pack(fill="both", expand=True)
        
        # Class management tab
        self.create_class_management_tab()
        
        # Student management tab
        self.create_student_management_tab()
        
        # Score management tab
        self.create_score_management_tab()
        
        # Book management tab
        self.create_book_management_tab()

        # Logout button
        self.content_seach = ttk.Frame(self.root)
        self.content_seach.pack(fill="both", expand=True)
        btnDangXuat = ttk.Button(self.content_seach, text="Đăng xuất", width=25, command=self.dangxuat)
        btnDangXuat.pack(side="right", anchor="ne")

    def create_class_management_tab(self):
        self.class_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.class_management_tab, text="Quản lý lớp học")
        
        # Frame to hold the buttons
        button_frame = ttk.Frame(self.class_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        # Buttons
        btnAddNew = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_Class, width=25, style='TButton')
        btnXuatExcel = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel, width=25, style='TButton')
        
        btnAddNew.pack(side="right", padx=5, pady=5)
        btnXuatExcel.pack(side="right", padx=5, pady=5)
        
        table_columns = ["CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER"]
        self.table = ttk.Treeview(self.class_management_tab, columns=table_columns, show="headings", height=25)
        for col in table_columns:
            self.table.heading(col, text=col)
        for row in result_list_Class:
            self.table.insert("", "end", values=row)
        self.table.pack(fill="x")
        
        self.create_search_section(self.class_management_tab, "class")

    def create_student_management_tab(self):
        self.student_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.student_management_tab, text="Quản lý học sinh")
        
        button_frame = ttk.Frame(self.student_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew1 = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnXuatExcel1 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        
        btnAddNew1.pack(side="right", padx=5, pady=5)
        btnXuatExcel1.pack(side="right", padx=5, pady=5)
        
        table_columns1 = ["ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME"]
        self.table1 = ttk.Treeview(self.student_management_tab, columns=table_columns1, show="headings", height=25)
        for col in table_columns1:
            self.table1.heading(col, text=col)
        for row in result_list_Student:
            self.table1.insert("", "end", values=row)
        self.table1.pack(fill="x")
        tree_scroll_y1 = ttk.Scrollbar(self.student_management_tab, orient="vertical", command=self.table1.yview)
        tree_scroll_y1.pack(side="right", fill="y")
        self.table1.configure(yscrollcommand=tree_scroll_y1.set)
        self.create_search_section(self.student_management_tab, "student")

    def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        
        button_frame = ttk.Frame(self.score_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew2 = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel",command=self.XuatExcel12, width=25, style='TButton')
        
        btnAddNew2.pack(side="right", padx=5, pady=5)
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING & READING", "TOTAL GRADE", "PERCENT"]
        self.table2 = ttk.Treeview(self.score_management_tab, columns=table_columns2, show="headings", height=25)
        for col in table_columns2:
            self.table2.heading(col, text=col)
        for row in combined_data6:
            self.table2.insert("", "end", values=row)
        self.table2.pack(fill="x")
        
        tree_scrollx2 = ttk.Scrollbar(self.score_management_tab, orient="horizontal", command=self.table2.xview)
        tree_scrollx2.pack(fill="x")
        self.table2.configure(xscrollcommand=tree_scrollx2.set)
        
        self.create_search_section(self.score_management_tab, "score")

    def create_book_management_tab(self):
        self.book_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.book_management_tab, text="Quản lý sách")
        
        button_frame = ttk.Frame(self.book_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew3 = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnXuatExcel3 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel3, width=25, style='TButton')
        
        btnAddNew3.pack(side="right", padx=5, pady=5)
        btnXuatExcel3.pack(side="right", padx=5, pady=5)
        
        table_columns3 = ["ID", "CAMBRIDGE LEVEL", "BOOK NAME", "MAIN BOOK"]
        self.table3 = ttk.Treeview(self.book_management_tab, columns=table_columns3, show="headings", height=25)
        for col in table_columns3:
            self.table3.heading(col, text=col)
        for row in result_list_Book:
            self.table3.insert("", "end", values=row)
        self.table3.pack(fill="x")
        
        tree_scrollx3 = ttk.Scrollbar(self.book_management_tab, orient="horizontal", command=self.table3.xview)
        tree_scrollx3.pack(fill="x")
        self.table3.configure(xscrollcommand=tree_scrollx3.set)
        
        tree_scroll_y3 = ttk.Scrollbar(self.book_management_tab, orient="vertical", command=self.table3.yview)
        tree_scroll_y3.pack(side="right", fill="y")
        self.table3.configure(yscrollcommand=tree_scroll_y3.set)
        
        self.create_search_section(self.book_management_tab, "book")

    def create_search_section(self, tab, type_):
        if type_ == "class":
            fields = ["Giáo viên", "Phòng", "Lớp", "ID lớp"]
        elif type_ == "student":
            fields = ["Tên lớp", "Tên học sinh", "ID lớp"]
        elif type_ == "score":
            fields = ["Giáo viên", "Lớp", "Tên học sinh", "ID"]
        elif type_ == "book":
            fields = ["Tên sách", "CAMBRIDGE LEVEL", "ID"]
        
        for field in fields:
            lbl = ttk.Label(tab, text=f"Nhập {field}:", style='TLabel')
            lbl.pack(side="left", anchor="ne", padx=5, pady=5)
            tf = ttk.Entry(tab, width=25, style='TEntry')
            tf.pack(side="left", anchor="ne", ipady=3, padx=5, pady=5)
        
        btnSearch = ttk.Button(tab, text="Tìm kiếm", width=25, style='TButton')
        btnSearch.pack(side="left", anchor="ne", ipady=3, padx=5, pady=5)

    def run(self):
        self.root.mainloop()

    def dangxuat(self):
        self.root.destroy()
    
    def XuatExcel3(self):
        # Function to get data from a specified range
        def get_data_from_range(sheet, start_row, end_row, start_col, end_col):
            data = []
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    row_data.append(sheet[row, col])
                data.append(row_data)
            return data

        # Download the specific Google Sheet
        ss = ezsheets.Spreadsheet("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")

        # Specify the sheet, columns, and rows
        # lag
        sheet_name = 'sheet 1'  # Change this to the specific sheet name
        start_row = 1  # Skip the header row
        end_row = 13
        start_col = 3
        test = worksheet.get_all_values()
        end_col = len([row[1] for row in test] )

        sheet = ss[sheet_name]
        data = get_data_from_range(sheet, start_row, end_row, start_col, end_col)

        # Create a new Excel file and write data to it
        wb = Workbook()
        ws = wb.active

        # Write headers and data to the new Excel sheet
        headers = [
            "ID_BOOK", "CAMBRIDGE LEVEL", "PROGRESS", "MAIN BOOK", "SKILL BOOK 1",
            "VOCAB BOOK", "SKILL BOOK 2", "SKILL BOOK 3", "SKILL BOOK 4",
            "GRAMMAR BOOK", "TEST BOOK", "VIDEOS-MOVIES", "PICTURES-CARDS"
        ]

        # Define styles
        header_font = Font(bold=True, color="000000")
        header_alignment = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        title_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        title_font = Font(bold=True, size=14)

        # Write the title and apply styles
        title_cell = ws.cell(row=1, column=1, value="Quản Lý Sách")
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = header_alignment
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)

        # Write the headers and apply styles
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        # Write the data and apply borders
        for col_idx, col_data in enumerate(data, start=1):
            for row_idx, value in enumerate(col_data, start=3):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = thin_border

        # Adjust column widths to fit the content
        for col in ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            max_length = 0
            column = get_column_letter(col[0].column)  # Get the column name
            for cell in col:
                if cell.value is not None:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        # Generate unique file name with date and time
        current_time = datetime.now().strftime("%d-%m-%Y-%H-%M-%S")
        file_path = f'D:\\QLS-{current_time}.xlsx'

        # Save the new Excel file
        wb.save(file_path)
        messagebox.showinfo("Success", "Download the file successfully, please check your D drive!")

    def XuatExcel(self):
        # Function to get data from a specified range
        def get_data_from_range(sheet, start_row, end_row, start_col, end_col):
            data = []
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    row_data.append(sheet[row, col])
                data.append(row_data)
            return data

        # Download the specific Google Sheet
        ss = ezsheets.Spreadsheet("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")

        # Specify the sheet, columns, and rows
        # lag
        sheet_name = 'sheet 1'  # Change this to the specific sheet name
        start_row = 15  # Skip the header row
        end_row = 21
        start_col = 3
        test = worksheet.get_all_values()
        end_col = len([row[15] for row in test] )

        sheet = ss[sheet_name]
        data = get_data_from_range(sheet, start_row, end_row, start_col, end_col)

        # Create a new Excel file and write data to it
        wb = Workbook()
        ws = wb.active

        # Write headers and data to the new Excel sheet
        headers = [
            "CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER", "FOREIGN TEACHER"
        ]

        # Define styles
        header_font = Font(bold=True, color="000000")
        header_alignment = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        title_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        title_font = Font(bold=True, size=14)

        # Write the title and apply styles
        title_cell = ws.cell(row=1, column=1, value="Quản Lý Lớp Học")
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = header_alignment
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)

        # Write the headers and apply styles
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        # Write the data and apply borders
        for col_idx, col_data in enumerate(data, start=1):
            for row_idx, value in enumerate(col_data, start=3):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = thin_border

        # Adjust column widths to fit the content
        for col in ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            max_length = 0
            column = get_column_letter(col[0].column)  # Get the column name
            for cell in col:
                if cell.value is not None:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        # Generate unique file name with date and time
        current_time = datetime.now().strftime("%d-%m-%Y-%H-%M-%S")
        file_path = f'D:\\QLLH-{current_time}.xlsx'

        # Save the new Excel file
        wb.save(file_path)
        messagebox.showinfo("Success", "Download the file successfully, please check your D drive!")

    def XuatExcel12(self):
        # Function to get data from a specified range
        def get_data_from_range(sheet, start_row, end_row, start_col, end_col):
            data = []
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    row_data.append(sheet[row, col])
                data.append(row_data)
            return data

        # Download the specific Google Sheet
        ss = ezsheets.Spreadsheet("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")

        # Specify the sheet, columns, and rows
        # lag
        sheet_name = 'sheet 1'  # Change this to the specific sheet name
        start_row = 23  # Skip the header row
        end_row = 45
        start_col = 3
        test = worksheet.get_all_values()
        end_col = len([row[23] for row in test] )

        sheet = ss[sheet_name]
        data = get_data_from_range(sheet, start_row, end_row, start_col, end_col)

        # Create a new Excel file and write data to it
        wb = Workbook()
        ws = wb.active

        # Write headers and data to the new Excel sheet
        headers = [
            "ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME",	"ENROLCAMP",
            "MAIN CAMP", "TOTAL FEE", "MAIN FEE", "NEW COMER", "STARTING OFF MONTH", "STARTING QUIT MONTH", "CERTIFICATE",	
            "PUBLIC SCHOOL", "SUB TEL", "STARTING TRANSFER MONTH", "TEACHER", "LISTENING", "SPEAKING"
            "READING & WRITING", "TOTAL GRADE", "PERCENT"
        ]

        # Define styles
        header_font = Font(bold=True, color="000000")
        header_alignment = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        title_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        title_font = Font(bold=True, size=14)

        # Write the title and apply styles
        title_cell = ws.cell(row=1, column=1, value="Quản Lý Lớp Học")
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = header_alignment
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=23)

        # Write the headers and apply styles
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        # Write the data and apply borders
        for col_idx, col_data in enumerate(data, start=1):
            for row_idx, value in enumerate(col_data, start=3):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = thin_border

        # Adjust column widths to fit the content
        for col in ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            max_length = 0
            column = get_column_letter(col[0].column)  # Get the column name
            for cell in col:
                if cell.value is not None:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        # Generate unique file name with date and time
        current_time = datetime.now().strftime("%d-%m-%Y-%H-%M-%S")
        file_path = f'D:\\QLHS-{current_time}.xlsx'

        # Save the new Excel file
        wb.save(file_path)
        messagebox.showinfo("Success", "Download the file successfully, please check your D drive!")
        
    def AddGUI_Class(self):
        AddNewClass = Add_NewClass()
        AddNewClass.run()

if __name__ == "__main__":
    app = MainFormGUI()
    app.run()