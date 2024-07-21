import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import gspread
import ezsheets
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
from GUI.Add_NewClass import Add_NewClass
from GUI.Add_NewBook import Add_NewBook
from GUI.Add_NewStudent import Add_NewStudent


#connect to gg sheet
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
worksheet = sht.sheet1

#show data
values_list_Book = worksheet.get_all_values()[2:]
result_list_Book = [row[:5] for row in values_list_Book]
worksheet2 = sht.worksheet("sheet 2")
values_list_Student = worksheet2.get_all_values()[2:]
result_list_Student = [row[:7] for row in values_list_Student]
worksheet3 = sht.worksheet("sheet 3")
values_list_Class = worksheet3.get_all_values()[2:]
result_list_Class = [row[:7] for row in values_list_Class]
values_list_Score = worksheet.get_all_values()[2:]
worksheet2 = sht.worksheet("sheet 2")
values_list_Score = worksheet2.get_all_values()[2:]  
result_list_Score = [row[:2] for row in values_list_Score] 
lop = [row[3] for row in values_list_Score]
combined_data = result_list_Score.copy()
for i in range(len(combined_data)):
    combined_data[i].append(lop[i])
teacher = [row[18] for row in values_list_Score]
combined_data1 = result_list_Score.copy() 
for i in range(len(combined_data1)):
    combined_data1[i].append(teacher[i])
listen = [row[19] for row in values_list_Score]
combined_data2 = result_list_Score.copy() 
for i in range(len(combined_data2)):
    combined_data2[i].append(listen[i])
speak = [row[20] for row in values_list_Score]
combined_data3 = result_list_Score.copy() 
for i in range(len(combined_data3)):
    combined_data3[i].append(speak[i])
rw = [row[21] for row in values_list_Score]
combined_data4 = result_list_Score.copy() 
for i in range(len(combined_data4)):
    combined_data4[i].append(rw[i])
total = [row[22] for row in values_list_Score]
combined_data5 = result_list_Score.copy() 
for i in range(len(combined_data5)):
    combined_data5[i].append(total[i])
ps = [row[23] for row in values_list_Score]
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
        
        table_columns = ["CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER", "FOREIGN TEACHER"]
        self.table = ttk.Treeview(self.class_management_tab, columns=table_columns, show="headings", height=25)
        for col in table_columns:
            self.table.heading(col, text=col)
        for row in result_list_Class:
            self.table.insert("", "end", values=row)
        self.table.pack(fill="x")
        self.table.bind("<Double-1>", self.on_row_select)

        self.create_search_section(self.class_management_tab, "class")

    def create_student_management_tab(self):
        self.student_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.student_management_tab, text="Quản lý học sinh")
        
        button_frame = ttk.Frame(self.student_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew1 = ttk.Button(button_frame, text="Thêm mới",command=self.AddGUI_Student, width=25, style='TButton')
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
        self.table1.bind("<Double-1>", self.on_row_select1)

        tree_scrollx1 = ttk.Scrollbar(self.student_management_tab, orient="horizontal", command=self.table1.xview)
        tree_scrollx1.pack(fill="x")
        self.table1.configure(xscrollcommand=tree_scrollx1.set)
        self.create_search_section(self.student_management_tab, "student")

    def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        button_frame = ttk.Frame(self.score_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel",command=self.XuatExcel12, width=25, style='TButton')
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING & READING", "TOTAL GRADE", "PERCENT"]
        self.table2 = ttk.Treeview(self.score_management_tab, columns=table_columns2, show="headings", height=25)
        for col in table_columns2:
            self.table2.heading(col, text=col)
        for row in combined_data6:
            self.table2.insert("", "end", values=row)
        self.table2.pack(fill="x")
        self.table2.bind("<Double-1>", self.on_row_select2)

        tree_scrollx2 = ttk.Scrollbar(self.score_management_tab, orient="horizontal", command=self.table2.xview)
        tree_scrollx2.pack(fill="x")
        self.table2.configure(xscrollcommand=tree_scrollx2.set)
        
        self.create_search_section(self.score_management_tab, "score")

    def create_book_management_tab(self):
        self.book_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.book_management_tab, text="Quản lý sách")
        
        button_frame = ttk.Frame(self.book_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew3 = ttk.Button(button_frame, text="Thêm mới",command=self.AddGUI_Book, width=25, style='TButton')
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
        self.table3.bind("<Double-1>", self.on_row_select3)

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
    
    #Xuất Excel
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
        ss = ezsheets.Spreadsheet("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")

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
        ss = ezsheets.Spreadsheet("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")

        # Specify the sheet, columns, and rows
        # lag
        sheet_name = 'sheet 3'  # Change this to the specific sheet name
        start_row = 1  # Skip the header row
        end_row = 7
        start_col = 3
        test = worksheet3.get_all_values()
        end_col = len([row[1] for row in test] )

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
        ss = ezsheets.Spreadsheet("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")

        # Specify the sheet, columns, and rows
        # lag
        sheet_name = 'sheet 2'  # Change this to the specific sheet name
        start_row = 1  # Skip the header row
        end_row = 24
        start_col = 3
        test = worksheet2.get_all_values()
        end_col = len([row[1] for row in test] )

        sheet = ss[sheet_name]
        data = get_data_from_range(sheet, start_row, end_row, start_col, end_col)

        # Create a new Excel file and write data to it
        wb = Workbook()
        ws = wb.active

        # Write headers and data to the new Excel sheet
        headers = [
            "ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME",	"ENROLCAMP",
            "MAIN CAMP", "TOTAL FEE", "MAIN FEE", "NEW COMER", "STARTING OFF MONTH", "STARTING QUIT MONTH", "CERTIFICATE",	
            "PUBLIC SCHOOL", "SUB TEL", "STARTING TRANSFER MONTH", "TEACHER", "LISTENING", "SPEAKING",
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
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=24)

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

    def AddGUI_Book(self):
        AddNewBook = Add_NewBook()
        AddNewBook.run()
    
    def AddGUI_Student(self):
        AddNewStudent = Add_NewStudent()
        AddNewStudent.run()

    #select table
    def on_row_select(self, event):
        selected_item = self.table.selection()
        if selected_item:
            row_values = self.table.item(selected_item, "values")
            row_list = row_values[0]
            if row_list in worksheet3.col_values(1):
                vitribandau = "A"+str(worksheet3.find(row_values[0]).row)
                matched_row1 = worksheet3.find(row_values[0]).row

                count_values = len(worksheet3.row_values(matched_row1))
                row_data1 = worksheet3.row_values(matched_row1)
                if len(row_data1)<=6:
                    row_data1.extend([""] * (6 - len(row_data1) + 1))
                char_to_num = dict()
                letters = [chr(i) for i in range(65, 91)]
                n = 30
                mapping = {}
                for i in range(1, n + 1):
                    mapping[i] = letters[(i - 1) % 26]
                vitrisua = vitribandau+":"+mapping[count_values]+str(matched_row1)
            self.Edit_NewClass(row_data1,vitrisua)
            # print(row_data1)
        else:
            print("Value not found in the sheet.")


    def Edit_NewClass(self,row_data,vitrisua):
        self.rootClass = tk.Tk()
        self.rootClass.title("Edit class")
        self.rootClass.geometry("520x680")
        self.canvas1 = tk.Canvas(self.rootClass, width=self.rootClass.winfo_screenwidth(), height=self.rootClass.winfo_screenheight())
        self.canvas1.pack(fill=tk.BOTH, expand=True)
        self.panel1 = tk.Frame(self.canvas1, bd=4, relief="solid")
        self.panel1.place(x=10, y=10, width=500, height=650)
        self.lbl_addNewClass = tk.Label(self.panel1, text="Edit class", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewClass.place(x=180, y=10)
        self.lb1 = tk.Label(self.panel1, text="Main class", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=430, height=30)
        self.lb2 = tk.Label(self.panel1, text="STUDYING DAY", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=33, y=171)
        self.tf2 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf2.place(x=33, y=224, width=430, height=30)
        self.lb3 = tk.Label(self.panel1, text="STUDYING TIME", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=290)
        self.tf3 = tk.Entry(self.panel1, font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=340, width=200, height=30)
        self.lb4 = tk.Label(self.panel1, text="ROOM", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=260, y=290)
        self.tf4 = tk.Entry(self.panel1, font=("cambria", 13, "bold"))
        self.tf4.place(x=260, y=340, width=200, height=30)
        self.lb5 = tk.Label(self.panel1, text="TEACHER", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=400)
        self.tf5 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=445, width=430, height=30)
        self.lb6 = tk.Label(self.panel1, text="FOREIGN TEACHER", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=33, y=500)
        self.tf6 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf6.place(x=33, y=545, width=430, height=30)
        self.tf1.insert(0, row_data[1])
        self.tf2.insert(0, row_data[2])
        self.tf3.insert(0, row_data[3])
        self.tf4.insert(0, row_data[4])
        self.tf5.insert(0, row_data[5])
        self.tf6.insert(0, row_data[6])
        def chinhsua():
            name = self.tf1.get()
            day = self.tf2.get()
            time = self.tf3.get()
            room = int(self.tf4.get())
            teacher = self.tf5.get()
            fteacher = self.tf6.get()
            new_values = [int(row_data[0]),name,day,time,room,teacher,fteacher]
            # worksheet3.update(values=[new_values], range_name=vitrisua)
            try:
                worksheet3.update(values=[new_values], range_name=vitrisua)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")

            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel1, text="EDIT NEW",command=chinhsua, font=("cambria", 14, "bold"), width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=150, y=600)
    
        

    
        

    


    def on_row_select1(self, event):
        selected_item1 = self.table1.selection()
        if selected_item1:
            row_values1 = self.table1.item(selected_item1, "values")
            row_list1 = row_values1[0] 
            if row_list1 in worksheet2.col_values(1):
                matched_row = worksheet2.find(row_values1[0]).row
                row_data = worksheet2.row_values(matched_row)
            # print(row_data)
        else:
            print("Value not found in the sheet.")
    
    
    
    def on_row_select3(self, event):
        selected_item3 = self.table3.selection()
        if selected_item3:
            row_values3 = self.table3.item(selected_item3, "values")
            row_list3 = row_values3[0] 
            if row_list3 in worksheet.col_values(1):
                matched_row3 = worksheet.find(row_values3[0]).row
                row_data3 = worksheet.row_values(matched_row3)
            print(row_data3)
        else:
            print("Value not found in the sheet.")
    
    def on_row_select2(self, event):
        selected_item2 = self.table2.selection()
        if selected_item2:
            row_values2 = self.table2.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                matched_row2 = worksheet2.find(row_values2[0]).row
                row_data2 = worksheet2.row_values(matched_row2)
            print(row_data2)
        else:
            print("Value not found in the sheet.")


if __name__ == "__main__":
    app = MainFormGUI()
    app.run()