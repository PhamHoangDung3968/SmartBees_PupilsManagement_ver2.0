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
        self.style.configure('TEntry', background='#007acc', foreground='#007acc', font=('Cambria', 12))
        self.style.configure('TNotebook.Tab', font=('Cambria', 14, 'bold'), background='#007acc', foreground='#007acc')
        self.style.configure('TTreeview.Heading', font=('Cambria', 11, 'bold'), background='#007ACC', foreground='white')
        self.style.configure('TTreeview', font=('Cambria', 11), background='#f5f5f5', foreground='#333333')

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
        for col in table_columns:
            self.table.heading(col, text=col)
        self.populate_table(self.table, self.original_data_class)
        self.table.pack(fill="x")

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
        for col in table_columns1:
            self.table1.heading(col, text=col)
        self.populate_table(self.table1, self.original_data_student)
        self.table1.pack(fill="x")
        tree_scroll_y1 = ttk.Scrollbar(self.student_management_tab, orient="vertical", command=self.table1.yview)
        tree_scroll_y1.pack(side="right", fill="y")
        self.table1.configure(yscrollcommand=tree_scroll_y1.set)

        tree_scrollx1 = ttk.Scrollbar(self.student_management_tab, orient="horizontal", command=self.table1.xview)
        tree_scrollx1.pack(fill="x")
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
        for col in table_columns2:
            self.table2.heading(col, text=col)
        self.populate_table(self.table2, self.original_data_score)
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
        
        btnAddNew3 = ttk.Button(button_frame, text="Thêm mới",command=self.AddGUI_Book, width=25, style='TButton')
        btnXuatExcel3 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel3, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("book"), width=25, style='TButton')
        
        btnAddNew3.pack(side="right", padx=5, pady=5)
        btnXuatExcel3.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        table_columns3 = ["ID", "CAMBRIDGE LEVEL", "BOOK NAME", "MAIN BOOK"]
        self.table3 = ttk.Treeview(self.book_management_tab, columns=table_columns3, show="headings", height=25)
        for col in table_columns3:
            self.table3.heading(col, text=col)
        self.populate_table(self.table3, self.original_data_book)
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
    


if __name__ == "__main__":
    app = MainFormGUI()
    app.run()