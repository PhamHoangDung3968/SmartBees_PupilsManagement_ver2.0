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


def initialize_globals():
    global gs, sht, worksheet, worksheet2, worksheet3
    global values_list_Book, result_list_Book
    global values_list_Student, result_list_Student
    global values_list_Class, result_list_Class
    global values_list_Score, result_list_Score
    global lop, combined_data
    global teacher, combined_data1
    global listen, combined_data2
    global speak, combined_data3
    global rw, combined_data4
    global total, combined_data5
    global ps, combined_data6

    # Connect to Google Sheets
    gs = gspread.service_account("cre.json")
    sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
    worksheet = sht.sheet1

    # Show data
    values_list_Book = worksheet.get_all_values()[2:]
    result_list_Book = [row[:5] for row in values_list_Book]

    worksheet2 = sht.worksheet("sheet 2")
    values_list_Student = worksheet2.get_all_values()[2:]
    result_list_Student = [row[:7] for row in values_list_Student]

    worksheet3 = sht.worksheet("sheet 3")
    values_list_Class = worksheet3.get_all_values()[2:]
    result_list_Class = [row[:7] for row in values_list_Class]

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

# Call the function to initialize the globals
initialize_globals()


from EXCEL.Excel_creating import Excel_Create

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
        self.style.configure('TEntry', background='#007acc', foreground='#007acc', font=('Cambria', 16,'bold'))
        self.style.configure('TNotebook.Tab', font=('Cambria', 14, 'bold'), background='#007acc', foreground='#007acc')
        self.style.configure('TTreeview.Heading', font=('Cambria', 20, 'bold'), background='#007ACC', foreground='white')
        self.style.configure('TTreeview', font=('Cambria', 20), background='#f5f5f5', foreground='#333333')

        # Main content frame
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True)
        self.tab_control = ttk.Notebook(self.content_frame)
        self.tab_control.pack(fill="both", expand=True)
        
        # Store entry widgets
        self.entries = {
            "class": [],
            "student": [],
            "score": [],
            "book": []
        }
        
        # Original data storage
        self.original_data_class = result_list_Class[:]
        self.original_data_student = result_list_Student[:]
        self.original_data_score = combined_data6[:]
        self.original_data_book = result_list_Book[:]
        
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
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("class"), width=25, style='TButton')
        
        btnAddNew.pack(side="right", padx=5, pady=5)
        btnXuatExcel.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        table_columns = ["CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER", "FOREIGN TEACHER"]
        self.table = ttk.Treeview(self.class_management_tab, columns=table_columns, show="headings", height=25)
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "CLASSNO": 50,
            "MAIN CLASS": 50,
            "STUDYING DAY": 50,
            "STUDYING TIME": 50,
            "ROOM": 50,
            "TEACHER": 200,
            "FOREIGN TEACHER": 200,
        }

        for col in table_columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình

        self.populate_table(self.table, self.original_data_class)
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
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("student"), width=25, style='TButton')
        
        btnAddNew1.pack(side="right", padx=5, pady=5)
        btnXuatExcel1.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        table_columns1 = ["ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME"]
        self.table1 = ttk.Treeview(self.student_management_tab, columns=table_columns1, show="headings", height=25)
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 50,
            "FULL NAME": 200,
            "BIRTHDAY (DOB)": 50,
            "MAIN CLASS": 50,
            "TEL": 50,
            "ADDRESS": 300,
            "PARENT NAME": 200,
        }

        for col in table_columns1:
            self.table1.heading(col, text=col)
            self.table1.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng cho từng cột

        
        self.populate_table(self.table1, self.original_data_student)
        self.table1.pack(fill="x")
        tree_scroll_y1 = ttk.Scrollbar(self.student_management_tab, orient="vertical", command=self.table1.yview)
        tree_scroll_y1.pack(side="right", fill="y")
        self.table1.configure(yscrollcommand=tree_scroll_y1.set)

        tree_scrollx1 = ttk.Scrollbar(self.student_management_tab, orient="horizontal", command=self.table1.xview)
        tree_scrollx1.pack(fill="x")
        self.table1.bind("<Double-1>", self.on_row_select1)

        self.table1.configure(xscrollcommand=tree_scrollx1.set)
        self.create_search_section(self.student_management_tab, "student")

    def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        button_frame = ttk.Frame(self.score_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("score"), width=25, style='TButton')
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING & READING", "TOTAL GRADE", "PERCENT"]
        self.table2 = ttk.Treeview(self.score_management_tab, columns=table_columns2, show="headings", height=25)
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 50,
            "FULL NAME": 200,
            "MAIN CLASS": 50,
            "TEACHER": 150,
            "LISTENING": 50,
            "SPEAKING": 50,
            "WRITING & READING": 50,
            "TOTAL GRADE": 50,
            "PERCENT": 50,
        }

        for col in table_columns2:
            self.table2.heading(col, text=col)
            self.table2.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng cho từng cột

        
        self.populate_table(self.table2, self.original_data_score)
        self.table2.pack(fill="x")

        tree_scrollx2 = ttk.Scrollbar(self.score_management_tab, orient="horizontal", command=self.table2.xview)
        tree_scrollx2.pack(fill="x")
        self.table2.bind("<Double-1>", self.on_row_select1)
        self.table2.configure(xscrollcommand=tree_scrollx2.set)
        
        self.create_search_section(self.score_management_tab, "score")

    def create_book_management_tab(self):
        self.book_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.book_management_tab, text="Quản lý sách")
        
        button_frame = ttk.Frame(self.book_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew3 = ttk.Button(button_frame, text="Thêm mới",command=self.AddGUI_Book, width=25, style='TButton')
        btnXuatExcel3 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel3, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("book"), width=25, style='TButton')
        
        btnAddNew3.pack(side="right", padx=5, pady=5)
        btnXuatExcel3.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        table_columns3 = ["ID", "CAMBRIDGE LEVEL", "BOOK NAME", "MAIN BOOK"]
        self.table3 = ttk.Treeview(self.book_management_tab, columns=table_columns3, show="headings", height=25)
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 50,
            "CAMBRIDGE LEVEL": 150,
            "BOOK NAME": 200,
            "MAIN BOOK": 100,
        }

        for col in table_columns3:
            self.table3.heading(col, text=col)
            self.table3.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng cho từng cột

        
        self.populate_table(self.table3, self.original_data_book)
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
            fields = ["TEACHER", "ROOM", "MAIN CLASS", "CLASSNO"]
        elif type_ == "student":
            fields = ["MAIN CLASS", "FULL NAME", "CLASSNO"]
        elif type_ == "score":
            fields = ["TEACHER", "MAIN CLASS", "FULL NAME", "ID"]
        elif type_ == "book":
            fields = ["BOOK NAME", "CAMBRIDGE LEVEL", "ID"]
        
        self.entries[type_] = {}  # Use a dictionary to store entry widgets
        for i, field in enumerate(fields):
            lbl_name = f"lbl{i+1}"
            tf_name = f"tf{i+1}"
            
            lbl = ttk.Label(tab, text=f"Nhập {field}:", style='TLabel')
            lbl.pack(side="left", anchor="ne", padx=5, pady=5)
            
            tf = ttk.Entry(tab, width=25, style='TEntry')
            tf.pack(side="left", anchor="ne", ipady=3, padx=5, pady=5)
            
            # Store the widgets in the dictionary
            self.entries[type_][lbl_name] = lbl
            self.entries[type_][tf_name] = tf

        btnSearch = ttk.Button(tab, text="Tìm kiếm", width=25, style='TButton', command=lambda t=type_: self.searching(t))
        btnSearch.pack(side="left", anchor="ne", ipady=3, padx=5, pady=5)
        
    def searching(self, type_):
        # Define the column mappings for class management
        column_mapping = {
            "class": {
                "TEACHER": 5,
                "ROOM": 4,
                "MAIN CLASS": 1,
                "CLASSNO": 0
            },
            "student": {
                "MAIN CLASS": 3,
                "FULL NAME": 1,
                "ID": 0
            },
            "score": {
                "TEACHER": 3,
                "MAIN CLASS": 2,
                "FULL NAME": 1,
                "ID": 0
            },
            "book": {
                "BOOK NAME": 2,
                "CAMBRIDGE LEVEL": 1,
                "ID": 0,
            }
        }
        
        search_criteria = {key: entry.get().lower() for key, entry in self.entries[type_].items() if 'tf' in key}
        matching_rows = []

        # Adjust the column_mapping based on type_
        mapping = column_mapping[type_]

        if type_ == "class":
            data_source = self.original_data_class
            table = self.table
        elif type_ == "student":
            data_source = self.original_data_student
            table = self.table1
        elif type_ == "score":
            data_source = self.original_data_score
            table = self.table2
        elif type_ == "book":
            data_source = self.original_data_book
            table = self.table3

        for row in data_source:
            if all(search_criteria[f'tf{index+1}'] in row[mapping[field]].lower() for index, field in enumerate(mapping)):
                matching_rows.append(row)
        
        self.populate_table(table, matching_rows)
        
        if not matching_rows:
            messagebox.showinfo("Thông báo", "Không tìm thấy dữ liệu khớp với các tiêu chí tìm kiếm đã nhập")

    def populate_table(self, table, data):
        for row in table.get_children():
            table.delete(row)
        for row in data:
            table.insert("", "end", values=row)


    def reload_tab(self, type_):
        initialize_globals()  # Refresh the global data
        self.update_original_data()  # Update the instance variables with new data
        if type_ == "class":
            self.populate_table(self.table, self.original_data_class)
        elif type_ == "student":
            self.populate_table(self.table1, self.original_data_student)
        elif type_ == "score":
            self.populate_table(self.table2, self.original_data_score)
        elif type_ == "book":
            self.populate_table(self.table3, self.original_data_book)

        # Reset search entries
        for entry in self.entries[type_].values():
            if isinstance(entry, ttk.Entry):
                entry.delete(0, tk.END)


    def update_original_data(self):
        self.original_data_class = result_list_Class
        self.original_data_student = result_list_Student
        self.original_data_score = result_list_Score
        self.original_data_book = result_list_Book

    
    def run(self):
        self.root.mainloop()

    def dangxuat(self):
        self.root.destroy()
    
    
    def XuatExcel(self):
        Xuat1 = Excel_Create()
        Xuat1.XuatExcel()
        
    def XuatExcel12(self):
        Xuat2 = Excel_Create()
        Xuat2.XuatExcel12()
    
    def XuatExcel3(self):
        Xuat3 = Excel_Create()
        Xuat3.XuatExcel3()
        
      
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

    #edit class
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
                row_data2 = len(row_data1)
                letters = [chr(i) for i in range(65, 91)]
                n = 30
                mapping = {}
                for i in range(1, n + 1):
                    mapping[i] = letters[(i - 1) % 26]
                vitrisua = vitribandau+":"+mapping[row_data2]+str(matched_row1)
            self.Edit_NewClass(row_data1,vitrisua)
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
            try:
                worksheet3.update(values=[new_values], range_name=vitrisua)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")
                self.rootClass.destroy()
                self.reload_tab("class")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel1, text="EDIT NEW",command=chinhsua, font=("cambria", 14, "bold"), width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=150, y=600)




    #edit student and point
    def on_row_select1(self, event):
        selected_item1 = self.table1.selection()
        if selected_item1:
            row_values2 = self.table1.item(selected_item1, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row

                # count_values3 = len(worksheet.row_values(matched_row3))
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=23:
                    row_data2.extend([""] * (23 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                letters2 = [chr(i) for i in range(65, 91)]
                n2 = 30
                mapping2 = {}
                for i in range(1, n2 + 1):
                    mapping2[i] = letters2[(i - 1) % 26]
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(row_data1)
            self.Edit_NewStudent(row_data2,vitrisua2)

        else:
            print("Value not found in the sheet.")
    
    def on_combobox_select(self, event = None):
            selected_value = self.tf13.get()
            return selected_value
             
    def Edit_NewStudent(self,row_data3,vitrisua3):
        self.rootStudent = tk.Tk()
        self.rootStudent.title("Edit student and point")
        self.rootStudent.geometry("1300x680")
        self.canvas2 = tk.Canvas(self.rootStudent, width=self.root.winfo_screenwidth(), height=self.rootStudent.winfo_screenheight())
        self.canvas2.pack(fill=tk.BOTH, expand=True)
        self.panel2 = tk.Frame(self.canvas2, bd=4, relief="solid")
        self.panel2.place(x=10, y=10, width=1275, height=650)
        self.lbl_EditNewStudent = tk.Label(self.panel2, text="Edit student and point", font=("cambria", 24, "bold"), fg="black")
        self.lbl_EditNewStudent.place(x=450, y=10)
        self.lb1 = tk.Label(self.panel2, text="Full name", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel2, text="Birthday (DOB)", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        self.tf2 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf2.place(x=370, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel2, text="Address", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel2, text="Starting off month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=370, y=160)
        self.tf4 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf4.place(x=370, y=213, width=300, height=30)

        self.lb5 = tk.Label(self.panel2, text="Public school", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=300, height=30)

        self.lb6 = tk.Label(self.panel2, text="Starting transfer month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=370, y=270)
        self.tf6 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf6.place(x=370, y=320, width=300, height=30)
        
        self.lb7 = tk.Label(self.panel2, text="Parent name", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=430, height=30)

        self.lb8 = tk.Label(self.panel2, text="New Comer", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=370, y=375)
        self.tf8 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf8.place(x=370, y=420, width=300, height=30)

        self.lb9 = tk.Label(self.panel2, text="Tel", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=130, height=30)

        self.lb10 = tk.Label(self.panel2, text="Enrolcamp", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb10.place(x=200, y=480)
        self.tf10 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf10.place(x=200, y=530, width=130, height=30)

        self.lb11 = tk.Label(self.panel2, text="Main camp", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb11.place(x=370, y=480)
        self.tf11 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf11.place(x=370, y=530, width=130, height=30)

        self.lb12 = tk.Label(self.panel2, text="Total fee", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb12.place(x=540, y=480)
        self.tf12 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf12.place(x=540, y=530, width=130, height=30)



        self.lb13 = tk.Label(self.panel2, text="Main class", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb13.place(x=700, y=60)
        self.tf13 = ttk.Combobox(self.panel2, font=("cambria", 13, "bold"))
        values_list_Class13 = worksheet3.get_all_values()[2:]
        result_list_Class13 = [row[1] for row in values_list_Class13]
        self.tf13['values'] = result_list_Class13
        self.tf13.current(0)
        self.tf13.place(x=700, y=108, width=240, height=30)
        
        # Ràng buộc sự kiện chọn giá trị trong combobox với hàm xử lý sự kiện
        self.tf13.bind("<<ComboboxSelected>>", self.on_combobox_select)
        

        self.lb15 = tk.Label(self.panel2, text="Starting quit month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb15.place(x=700, y=160)
        self.tf15 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf15.place(x=700, y=213, width=240, height=30)

        self.lb16 = tk.Label(self.panel2, text="Teacher", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb16.place(x=700, y=270)
        self.tf16 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf16.place(x=700, y=320, width=240, height=30)

        self.lb17 = tk.Label(self.panel2, text="Sub tel", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb17.place(x=700, y=375)
        self.tf17 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf17.place(x=700, y=420, width=240, height=30)


        self.lb18 = tk.Label(self.panel2, text="Main fee", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb18.place(x=700, y=480)
        self.tf18 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf18.place(x=700, y=530, width=130, height=30)


        self.lb19 = tk.Label(self.panel2, text="Certificate", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb19.place(x=850, y=480)
        self.tf19 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf19.place(x=850, y=530, width=130, height=30)


        self.lb20 = tk.Label(self.panel2, text="Reading & Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb20.place(x=1000, y=160)
        self.tf20 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf20.place(x=1050, y=213, width=100, height=30)

        self.lb21 = tk.Label(self.panel2, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb21.place(x=1050, y=270)
        self.tf21 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf21.place(x=1050, y=320, width=100, height=30)

        self.lb22 = tk.Label(self.panel2, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb22.place(x=1050, y=375)
        self.tf22 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf22.place(x=1050, y=420, width=100, height=30)

        
        self.tf1.insert(0, row_data3[1])
        self.tf2.insert(0, row_data3[2])
        self.tf13.insert(0, row_data3[3])
        self.tf9.insert(0, row_data3[4])
        self.tf3.insert(0, row_data3[5])
        self.tf7.insert(0, row_data3[6])
        self.tf10.insert(0, row_data3[7])
        self.tf11.insert(0, row_data3[8])
        self.tf12.insert(0, row_data3[9])
        self.tf18.insert(0, row_data3[10])
        self.tf8.insert(0, row_data3[11])
        self.tf4.insert(0, row_data3[12])
        self.tf15.insert(0, row_data3[13])
        self.tf19.insert(0, row_data3[14])
        self.tf5.insert(0, row_data3[15])
        self.tf17.insert(0, row_data3[16])
        self.tf6.insert(0, row_data3[17])
        self.tf16.insert(0, row_data3[18])
        self.tf21.insert(0, row_data3[19])
        self.tf22.insert(0, row_data3[20])
        self.tf20.insert(0, row_data3[21]) 
          
        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf10.get()
            a3 = self.on_combobox_select()
            a4 = self.tf9.get()
            a5 = self.tf3.get()
            a6 = self.tf7.get()
            a7 = self.tf10.get()
            a8 = self.tf11.get()
            a9 = self.tf12.get()
            a10 = self.tf18.get()
            a11 = self.tf8.get()
            a12 = self.tf4.get()
            a13 = self.tf15.get()
            a14 = self.tf19.get()
            a15 = self.tf5.get()
            a16 = self.tf17.get()
            a17 = self.tf6.get()
            a18 = self.tf16.get()
            
            try:
                a21 = float(self.tf20.get())
            except ValueError:
                a21 = 0
            try:
                a19 = float(self.tf21.get())
            except ValueError:
                a19 = 0
            try:
                a20 = float(self.tf22.get())
            except ValueError:
                a20 = 0
            total = a21 + a19 + a20
            percent = str(round((total/15)*100,2))+"%"
            

            new_values3 = [int(row_data3[0]),a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,total,percent]
            # worksheet3.update(values=[new_values], range_name=vitrisua)
            try:
                worksheet2.update(values=[new_values3], range_name=vitrisua3)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")
                self.rootStudent.destroy()
                self.reload_tab("student")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        
        self.btn1 = tk.Button(self.panel2, text="SUBMIT", font=("cambria", 14, "bold"),command=chinhsua, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=520, y=600)
    




    
    #edit book
    def on_row_select3(self, event):
        selected_item3 = self.table3.selection()
        if selected_item3:
            row_values3 = self.table3.item(selected_item3, "values")
            row_list3 = row_values3[0] 
            if row_list3 in worksheet.col_values(1):
                vitribandau3 = "A"+str(worksheet.find(row_values3[0]).row)
                matched_row3 = worksheet.find(row_values3[0]).row

                # count_values3 = len(worksheet.row_values(matched_row3))
                row_data3 = worksheet.row_values(matched_row3)
                if len(row_data3)<=12:
                    row_data3.extend([""] * (12 - len(row_data3) + 1))
                char_to_num = dict()
                count_values3 = len(row_data3)
                letters3 = [chr(i) for i in range(65, 91)]
                n3 = 30
                mapping3 = {}
                for i in range(1, n3 + 1):
                    mapping3[i] = letters3[(i - 1) % 26]
                vitrisua3 = vitribandau3+":"+mapping3[count_values3]+str(matched_row3)
            self.Edit_NewBook(row_data3,vitrisua3)
        else:
            print("Value not found in the sheet.")
    
    def Edit_NewBook(self,row_data3,vitrisua3):
        self.rootBook = tk.Tk()
        self.rootBook.title("Edit book")
        self.rootBook.geometry("1020x680")
        self.canvas3 = tk.Canvas(self.rootBook, width=self.rootBook.winfo_screenwidth(), height=self.rootBook.winfo_screenheight())
        self.canvas3.pack(fill=tk.BOTH, expand=True)
        self.panel3 = tk.Frame(self.canvas3, bd=4, relief="solid")
        self.panel3.place(x=10, y=10, width=1000, height=650)
        self.lbl_editNewBook = tk.Label(self.panel3, text="Edit book", font=("cambria", 24, "bold"), fg="black")
        self.lbl_editNewBook.place(x=300, y=10)
        self.lb1 = tk.Label(self.panel3, text="Cambridge level", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=430, height=30)

        self.lb2 = tk.Label(self.panel3, text="Main book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=530, y=60)
        self.tf2 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf2.place(x=530, y=108, width=430, height=30)


        self.lb3 = tk.Label(self.panel3, text="Skill book 1", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=430, height=30)

        self.lb4 = tk.Label(self.panel3, text="Skill book 2", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=530, y=160)
        self.tf4 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf4.place(x=530, y=213, width=430, height=30)

        self.lb5 = tk.Label(self.panel3, text="Skill book 3", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=430, height=30)

        self.lb6 = tk.Label(self.panel3, text="Skill book 4", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=530, y=270)
        self.tf6 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf6.place(x=530, y=320, width=430, height=30)

        self.lb7 = tk.Label(self.panel3, text="Vocab book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=430, height=30)

        self.lb8 = tk.Label(self.panel3, text="Grammar book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=530, y=375)
        self.tf8 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf8.place(x=530, y=420, width=430, height=30)

        self.lb9 = tk.Label(self.panel3, text="Test book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel3, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=200, height=30)

        self.lb10 = tk.Label(self.panel3, text="Progress", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=260, y=480)
        self.tf10 = tk.Entry(self.panel3, font=("cambria", 13, "bold"))
        self.tf10.place(x=260, y=530, width=200, height=30)

        self.lb11 = tk.Label(self.panel3, text="Videos-Movies", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb11.place(x=530, y=480)
        self.tf11 = tk.Entry(self.panel3, font=("cambria", 13, "bold"))
        self.tf11.place(x=530, y=530, width=200, height=30)

        self.lb12 = tk.Label(self.panel3, text="Pictures-Cards", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb12.place(x=750, y=480)
        self.tf12 = tk.Entry(self.panel3, font=("cambria", 13, "bold"))
        self.tf12.place(x=750, y=530, width=200, height=30)

        self.tf1.insert(0, row_data3[1])
        self.tf10.insert(0, row_data3[2])
        self.tf2.insert(0, row_data3[3])
        self.tf3.insert(0, row_data3[4])
        self.tf7.insert(0, row_data3[5])
        self.tf4.insert(0, row_data3[6])
        self.tf5.insert(0, row_data3[7])
        self.tf6.insert(0, row_data3[8])
        self.tf8.insert(0, row_data3[9])
        self.tf9.insert(0, row_data3[10])
        self.tf11.insert(0, row_data3[11])
        self.tf12.insert(0, row_data3[12])
        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf10.get()
            a3 = self.tf2.get()
            a4 = self.tf3.get()
            a5 = self.tf7.get()
            a6 = self.tf4.get()
            a7 = self.tf5.get()
            a8 = self.tf6.get()
            a9 = self.tf8.get()
            a10 = self.tf9.get()
            a11 = self.tf11.get()
            a12 = self.tf12.get()
            

            new_values3 = [int(row_data3[0]),a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12]
            # worksheet3.update(values=[new_values], range_name=vitrisua)
            try:
                worksheet.update(values=[new_values3], range_name=vitrisua3)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")
                self.rootBook.destroy()
                self.reload_tab("book")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel3, text="Edit", font=("cambria", 14, "bold"),command=chinhsua, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=400, y=600)
    
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