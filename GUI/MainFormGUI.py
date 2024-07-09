import tkinter as tk
from tkinter import ttk
import gspread

gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
worksheet = sht.sheet1
values_list = worksheet.get_all_values()[2:]
result_list_Book = [row[:5] for row in values_list]

class MainFormGUI:
    def __init__(self):
        self.root = tk.Tk()
        #Chỗ này để trang trí
        style = ttk.Style()
        # style.configure('TButton', font=('cambria', 11, 'bold'))
        # style.configure('TTreeview', font=('cambria', 11, 'bold'))
        # Tạo thanh cuộn ngang
        self.root.title("Main Form GUI")
        self.root.geometry("1097x700")
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True)
        self.tab_control = ttk.Notebook(self.content_frame)
        self.tab_control.pack(fill="both", expand=True)
        self.class_management_tab = ttk.Frame(self.tab_control)

        #Quản lý lớp học
        self.tab_control.add(self.class_management_tab, text="Quản lý lớp học")
        btnAddNew = ttk.Button(self.class_management_tab, text="Thêm mới", width=25)
        btnInPDF = ttk.Button(self.class_management_tab, text="In PDF", width=25)
        btnXuatExcel = ttk.Button(self.class_management_tab, text="Xuất excel", width=25)
        btnAddNew.pack(side="top", anchor="ne",ipady=5)
        btnInPDF.pack(side="top", anchor="ne",ipady=5)
        btnXuatExcel.pack(side="top", anchor="ne",ipady=5)
        table_columns = ["CLASSNO", "MAIN CLASS", "STUDYING DAY", "STUDYING TIME", "ROOM", "TEACHER"]
        table_data = [[None] * len(table_columns)]
        self.table = ttk.Treeview(self.class_management_tab, columns=table_columns, show="headings",height=25)
        for col in table_columns:
            self.table.heading(col, text=col)
        for row in table_data:
            self.table.insert("", "end", values=row)
        self.table.pack(fill="x")
        btnSearch_Class = ttk.Button(self.class_management_tab, text="Tìm kiếm", width=25)
        tfSearch_GV = ttk.Entry(self.class_management_tab, width=25 )
        lblGV = ttk.Label(self.class_management_tab, text="Nhập Tên giáo viên:")
        tfSearch_Phong = ttk.Entry(self.class_management_tab, width=25 )
        lblPhong = ttk.Label(self.class_management_tab, text="Nhập phòng:")
        tfSearch_Lop = ttk.Entry(self.class_management_tab, width=25 )
        lblLop = ttk.Label(self.class_management_tab, text="Nhập lớp:")
        tfSearch_IDLop = ttk.Entry(self.class_management_tab, width=25 )
        lblIDLop = ttk.Label(self.class_management_tab, text="Nhập ID lớp:")
        btnSearch_Class.pack(side="right", anchor="ne",ipady=3)
        tfSearch_GV.pack(side="right", anchor="ne", ipady=3)
        lblGV.pack(side="right", anchor="ne",ipady=3)
        tfSearch_Phong.pack(side="right", anchor="ne", ipady=3)
        lblPhong.pack(side="right", anchor="ne",ipady=3)
        tfSearch_Lop.pack(side="right", anchor="ne", ipady=3)
        lblLop.pack(side="right", anchor="ne",ipady=3)
        tfSearch_IDLop.pack(side="right", anchor="ne", ipady=3)
        lblIDLop.pack(side="right", anchor="ne",ipady=3)
        self.content_seach = ttk.Frame(self.root)
        self.content_seach.pack(fill="both", expand=True)
        btnDangXuat = ttk.Button(self.content_seach, text="Đăng xuất", width=25, command=self.dangxuat)
        btnDangXuat.pack(side="right", anchor="ne")
        



        #Quản lý học sinh
        self.class_management_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.class_management_tab, text="Quản lý học sinh")
        btnAddNew1 = ttk.Button(self.class_management_tab, text="Thêm mới", width=25)
        btnInPDF1 = ttk.Button(self.class_management_tab, text="In PDF", width=25)
        btnXuatExcel1 = ttk.Button(self.class_management_tab, text="Xuất excel", width=25)
        btnAddNew1.pack(side="top", anchor="ne",ipady=5)
        btnInPDF1.pack(side="top", anchor="ne",ipady=5)
        btnXuatExcel1.pack(side="top", anchor="ne",ipady=5)
        table_columns1 = ["ID", "FULL NAME", "BIRTHDAY (DOB)", "MAIN CLASS", "TEL", "ADDRESS", "PARENT NAME"]
        table_data1 = [[None] * len(table_columns1)]
        self.table1 = ttk.Treeview(self.class_management_tab, columns=table_columns1, show="headings",height=25)
        for col in table_columns1:
            self.table1.heading(col, text=col)
        for row in table_data1:
            self.table1.insert("", "end", values=row)
        self.table1.pack(fill="x")
        btnSearch_Student = ttk.Button(self.class_management_tab, text="Tìm kiếm", width=25)
        tfSearch_Lop_Student = ttk.Entry(self.class_management_tab, width=25 )
        lblLop_Student = ttk.Label(self.class_management_tab, text="Nhập tên lớp")
        tfSearch_Ten_Student = ttk.Entry(self.class_management_tab, width=25 )
        lblTen_Student = ttk.Label(self.class_management_tab, text="Nhập tên học sinh")
        tfSearch_ID_Student = ttk.Entry(self.class_management_tab, width=25 )
        lblID_Student = ttk.Label(self.class_management_tab, text="Nhập ID lớp:")
        btnSearch_Student.pack(side="right", anchor="ne")
        tfSearch_Lop_Student.pack(side="right", anchor="ne", ipady=3)
        lblLop_Student.pack(side="right", anchor="ne",ipady=3)
        tfSearch_Ten_Student.pack(side="right", anchor="ne", ipady=3)
        lblTen_Student.pack(side="right", anchor="ne",ipady=3)
        tfSearch_ID_Student.pack(side="right", anchor="ne", ipady=3)
        lblID_Student.pack(side="right", anchor="ne",ipady=3)







        #Quản lý điểm số
        self.class_management_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.class_management_tab, text="Quản lý điểm số")
        btnAddNew2 = ttk.Button(self.class_management_tab, text="Thêm mới", width=25)
        btnInPDF2 = ttk.Button(self.class_management_tab, text="In PDF", width=25)
        btnXuatExcel2 = ttk.Button(self.class_management_tab, text="Xuất excel", width=25)
        btnAddNew2.pack(side="top", anchor="ne",ipady=5)
        btnInPDF2.pack(side="top", anchor="ne",ipady=5)
        btnXuatExcel2.pack(side="top", anchor="ne",ipady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING", "READING", "TOTAL GRADE"]
        table_data2 = [[None] * len(table_columns2)]
        self.table2 = ttk.Treeview(self.class_management_tab, columns=table_columns2, show="headings",height=25)
        for col in table_columns2:
            self.table2.heading(col, text=col)
        for row in table_data2:
            self.table2.insert("", "end", values=row)
        self.table2.pack(fill="x")

        #scroll chuột
        tree_scrollx2 = ttk.Scrollbar(self.class_management_tab, orient="horizontal", command=self.table2.xview)
        tree_scrollx2.pack(fill="x")
        # Cấu hình liên kết thanh cuộn với Treeview
        self.table2.configure(xscrollcommand=tree_scrollx2.set)

#         # Tạo thanh cuộn dọc
#         tree_scroll_y = ttk.Scrollbar(self.class_management_tab, orient="vertical", command=self.table2.yview)
#         tree_scroll_y.pack(side="right", fill="y")

# # Cấu hình liên kết thanh cuộn với Treeview
#         self.table2.configure(yscrollcommand=tree_scroll_y.set)




        btnSearch_Score = ttk.Button(self.class_management_tab, text="Tìm kiếm", width=25)
        tfSearch_GV_Score = ttk.Entry(self.class_management_tab, width=25 )
        lblGV_Score = ttk.Label(self.class_management_tab, text="Nhập tên giáo viên")
        tfSearch_Lop_Score = ttk.Entry(self.class_management_tab, width=25 )
        lblLop_Scode = ttk.Label(self.class_management_tab, text="Nhập lớp")
        tfSearch_Ten_Score = ttk.Entry(self.class_management_tab, width=25 )
        lblTen_Score = ttk.Label(self.class_management_tab, text="Nhập tên học sinh")
        tfSearch_ID_Score = ttk.Entry(self.class_management_tab, width=25 )
        lblID_Score = ttk.Label(self.class_management_tab, text="Nhập ID")
        btnSearch_Score.pack(side="right", anchor="ne")
        tfSearch_GV_Score.pack(side="right", anchor="ne", ipady=3)
        lblGV_Score.pack(side="right", anchor="ne",ipady=3)
        tfSearch_Lop_Score.pack(side="right", anchor="ne", ipady=3)
        lblLop_Scode.pack(side="right", anchor="ne",ipady=3)
        tfSearch_Ten_Score.pack(side="right", anchor="ne", ipady=3)
        lblTen_Score.pack(side="right", anchor="ne",ipady=3)
        tfSearch_ID_Score.pack(side="right", anchor="ne", ipady=3)
        lblID_Score.pack(side="right", anchor="ne",ipady=3)

        
        #Quản lý sách 
        self.class_management_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.class_management_tab, text="Quản lý sách")
        btnAddNew3 = ttk.Button(self.class_management_tab, text="Thêm mới", width=25)
        btnInPDF3 = ttk.Button(self.class_management_tab, text="In PDF", width=25)
        btnXuatExcel3 = ttk.Button(self.class_management_tab, text="Xuất excel", width=25)
        btnAddNew3.pack(side="top", anchor="ne",ipady=5)
        btnInPDF3.pack(side="top", anchor="ne",ipady=5)
        btnXuatExcel3.pack(side="top", anchor="ne",ipady=5)
        
        table_columns3 = ["ID", "CAMBRIDE LEVER", "BOOK NAME", "MAIN BOOK"]
        table_data3 = result_list_Book
        self.table3 = ttk.Treeview(self.class_management_tab, columns=table_columns3, show="headings",height=25)
        for col in table_columns3:
            self.table3.heading(col, text=col)
        for row in table_data3:
            self.table3.insert("", "end", values=row)
        self.table3.pack(fill="x")

        #scroll chuột
        tree_scrollx3 = ttk.Scrollbar(self.class_management_tab, orient="horizontal", command=self.table3.xview)
        tree_scrollx3.pack(fill="x")
        # Cấu hình liên kết thanh cuộn với Treeview
        self.table3.configure(xscrollcommand=tree_scrollx3.set)
        # Tạo thanh cuộn dọc
        tree_scroll_y3 = ttk.Scrollbar(self.class_management_tab, orient="vertical", command=self.table3.yview)
        tree_scroll_y3.pack(side="right", fill="y")

# Cấu hình liên kết thanh cuộn với Treeview
        self.table3.configure(yscrollcommand=tree_scroll_y3.set)

        btnSearch_Book = ttk.Button(self.class_management_tab, text="Tìm kiếm", width=25)
        tfSearch_Ten_Book = ttk.Entry(self.class_management_tab, width=25 )
        lblTen_Book = ttk.Label(self.class_management_tab, text="Nhập tên sách")
        tfSearch_CL_Book = ttk.Entry(self.class_management_tab, width=25 )
        lblCLBook_Scode = ttk.Label(self.class_management_tab, text="Nhập CAMBRIDE LEVER")
        tfSearch_ID_Book = ttk.Entry(self.class_management_tab, width=25 )
        lblID_Book = ttk.Label(self.class_management_tab, text="Nhập ID")
        btnSearch_Book.pack(side="right", anchor="ne")
        tfSearch_Ten_Book.pack(side="right", anchor="ne", ipady=3)
        lblTen_Book.pack(side="right", anchor="ne",ipady=3)
        tfSearch_CL_Book.pack(side="right", anchor="ne", ipady=3)
        lblCLBook_Scode.pack(side="right", anchor="ne",ipady=3)
        tfSearch_ID_Book.pack(side="right", anchor="ne", ipady=3)
        lblID_Book.pack(side="right", anchor="ne",ipady=3)
        


    def run(self):
        self.root.mainloop()
    def dangxuat(self):
        self.root.destroy()

if __name__ == "__main__":
    app = MainFormGUI()
    app.run()

