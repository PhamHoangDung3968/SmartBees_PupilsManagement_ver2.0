
import tkinter as tk
from tkinter import ttk
import gspread

# Google Sheets setup
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
worksheet = sht.sheet1
values_list = worksheet.get_all_values()[2:]
result_list_Book = [row[:5] for row in values_list]

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
        btnAddNew = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnInPDF = ttk.Button(button_frame, text="In PDF", width=25, style='TButton')
        btnXuatExcel = ttk.Button(button_frame, text="Xuất excel", width=25, style='TButton')
        
        btnAddNew.pack(side="right", padx=5, pady=5)
        btnInPDF.pack(side="right", padx=5, pady=5)
        btnXuatExcel.pack(side="right", padx=5, pady=5)
        
        table_columns = ["CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER"]
        self.table = ttk.Treeview(self.class_management_tab, columns=table_columns, show="headings", height=25)
        for col in table_columns:
            self.table.heading(col, text=col)
        self.table.pack(fill="x")
        
        self.create_search_section(self.class_management_tab, "class")

    def create_student_management_tab(self):
        self.student_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.student_management_tab, text="Quản lý học sinh")
        
        button_frame = ttk.Frame(self.student_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew1 = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnInPDF1 = ttk.Button(button_frame, text="In PDF", width=25, style='TButton')
        btnXuatExcel1 = ttk.Button(button_frame, text="Xuất excel", width=25, style='TButton')
        
        btnAddNew1.pack(side="right", padx=5, pady=5)
        btnInPDF1.pack(side="right", padx=5, pady=5)
        btnXuatExcel1.pack(side="right", padx=5, pady=5)
        
        table_columns1 = ["ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME"]
        self.table1 = ttk.Treeview(self.student_management_tab, columns=table_columns1, show="headings", height=25)
        for col in table_columns1:
            self.table1.heading(col, text=col)
        self.table1.pack(fill="x")
        
        self.create_search_section(self.student_management_tab, "student")

    def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        
        button_frame = ttk.Frame(self.score_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew2 = ttk.Button(button_frame, text="Thêm mới", width=25, style='TButton')
        btnInPDF2 = ttk.Button(button_frame, text="In PDF", width=25, style='TButton')
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel", width=25, style='TButton')
        
        btnAddNew2.pack(side="right", padx=5, pady=5)
        btnInPDF2.pack(side="right", padx=5, pady=5)
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING", "READING", "TOTAL GRADE"]
        self.table2 = ttk.Treeview(self.score_management_tab, columns=table_columns2, show="headings", height=25)
        for col in table_columns2:
            self.table2.heading(col, text=col)
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
        btnInPDF3 = ttk.Button(button_frame, text="In PDF", width=25, style='TButton')
        btnXuatExcel3 = ttk.Button(button_frame, text="Xuất excel", width=25, style='TButton')
        
        btnAddNew3.pack(side="right", padx=5, pady=5)
        btnInPDF3.pack(side="right", padx=5, pady=5)
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

if __name__ == "__main__":
    app = MainFormGUI()
    app.run()

