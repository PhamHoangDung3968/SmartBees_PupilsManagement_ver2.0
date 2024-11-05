import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import DateEntry

import socket
import sys  # Đảm bảo đã import sys

import gspread
import ezsheets
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import string

from PDF.Score_PDF import create_file

from GUI.Add_NewClass import Add_NewClass
from GUI.Add_NewSach import Add_NewSach
from GUI.Add_NewStudent import Add_NewStudent
from GUI.Add_NewChangeClass import Add_NewChangeClass
from GUI.Add_NewReviewClass import Add_NewReviewClass




def initialize_globals():
    global gs, sht, worksheet, worksheet2, worksheet3, worksheet4, worksheet5
    global values_list_Book, result_list_Book
    global values_list_Student, result_list_Student
    global values_list_Class, result_list_Class
    global values_list_Score, result_list_Score
    global values_list_changeclass, result_list_changeclass
    global values_list_reviewclass, result_list_reviewclass
    global lop, combined_data
    global teacher, combined_data1
    global listen1, combined_data2
    global speak1, combined_data3
    global rw1, combined_data4
    global total1, combined_data5
    global ps1, combined_data6

    global tel, combined_data_student
    global diachi, combined_data_student1

    global values_list_Score2, result_list_Score2
    global listen2, combined_data2_2
    global speak2, combined_data3_2
    global rw2, combined_data4_2
    global total2, combined_data5_2
    global ps2, combined_data6_2

    global values_list_Score3, result_list_Score3
    global listen3, combined_data2_3
    global speak3, combined_data3_3
    global rw3, combined_data4_3
    global total3, combined_data5_3
    global ps3, combined_data6_3

    global values_list_Score4, result_list_Score4
    global listen4, combined_data2_4
    global speak4, combined_data3_4
    global rw4, combined_data4_4
    global total4, combined_data5_4
    global ps4, combined_data6_4

    global values_list_Score5, result_list_Score5
    global listen5, combined_data2_5
    global speak5, combined_data3_5
    global rw5, combined_data4_5
    global total5, combined_data5_5
    global ps5, combined_data6_5

    try:
        # Connect to Google Sheets
        gs = gspread.service_account("cre.json")
        sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
        worksheet = sht.sheet1

    except socket.gaierror:
        messagebox.showerror("Lỗi Kết Nối", "Lỗi DNS: Không thể tìm thấy địa chỉ máy chủ. Vui lòng kiểm tra kết nối Internet của bạn.")
        sys.exit("Chương trình dừng lại do không thể kết nối đến Google Sheets.")

    except Exception as ex:
        messagebox.showerror("Lỗi Không Xác Định", f"Có lỗi xảy ra: {ex}")
        sys.exit("Chương trình dừng lại do lỗi không xác định.")


    # Show data
    worksheet4 = sht.worksheet("sheet 4")
    values_list_reviewclass = worksheet4.get_all_values()[2:]
    result_list_reviewclass = [row[:6] for row in values_list_reviewclass]

    worksheet5 = sht.worksheet("sheet 5")
    values_list_changeclass = worksheet5.get_all_values()[2:]
    result_list_changeclass = [row[:9] for row in values_list_changeclass]

    values_list_Book = worksheet.get_all_values()[2:]
    result_list_Book = [row[:9] for row in values_list_Book]

    worksheet2 = sht.worksheet("sheet 2")
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
        

    worksheet3 = sht.worksheet("sheet 3")
    values_list_Class = worksheet3.get_all_values()[2:]
    result_list_Class = [row[:7] for row in values_list_Class]

    worksheet2 = sht.worksheet("sheet 2")
    # giai đoạn 1
    values_list_Score = worksheet2.get_all_values()[2:]
    result_list_Score = [row[:2] for row in values_list_Score]
    
    lop = [row[3] for row in values_list_Score]
    combined_data = result_list_Score.copy()
    for i in range(len(combined_data)):
        combined_data[i].append(lop[i])

    teacher = [row[21] for row in values_list_Score]
    combined_data1 = result_list_Score.copy()
    for i in range(len(combined_data1)):
        combined_data1[i].append(teacher[i])

    listen1 = [row[25] for row in values_list_Score]
    combined_data2 = result_list_Score.copy()
    for i in range(len(combined_data2)):
        combined_data2[i].append(listen1[i])

    speak1 = [row[26] for row in values_list_Score]
    combined_data3 = result_list_Score.copy()
    for i in range(len(combined_data3)):
        combined_data3[i].append(speak1[i])

    rw1 = [row[27] for row in values_list_Score]
    combined_data4 = result_list_Score.copy()
    for i in range(len(combined_data4)):
        combined_data4[i].append(rw1[i])

    total1 = [row[28] for row in values_list_Score]
    combined_data5 = result_list_Score.copy()
    for i in range(len(combined_data5)):
        combined_data5[i].append(total1[i])

    ps1 = [row[29] for row in values_list_Score]
    combined_data6 = result_list_Score.copy()
    for i in range(len(combined_data6)):
        combined_data6[i].append(ps1[i])

    #giai đoạn 2
    values_list_Score2 = worksheet2.get_all_values()[2:]
    result_list_Score2 = [row[:2] for row in values_list_Score2]
    lop2 = [row[3] for row in values_list_Score2]
    combined_data_2 = result_list_Score2.copy()
    for i in range(len(combined_data_2)):
        combined_data_2[i].append(lop2[i])

    teacher2 = [row[21] for row in values_list_Score2]
    combined_data1_2 = result_list_Score2.copy()
    for i in range(len(combined_data1_2)):
        combined_data1_2[i].append(teacher2[i])
        
    listen2 = [row[30] for row in values_list_Score2]
    combined_data2_2 = result_list_Score2.copy()
    for i in range(len(combined_data2_2)):
        combined_data2_2[i].append(listen2[i])

    speak2 = [row[31] for row in values_list_Score2]
    combined_data3_2 = result_list_Score2.copy()
    for i in range(len(combined_data3_2)):
        combined_data3_2[i].append(speak2[i])

    rw2 = [row[32] for row in values_list_Score2]
    combined_data4_2 = result_list_Score2.copy()
    for i in range(len(combined_data4_2)):
        combined_data4_2[i].append(rw2[i])

    total2 = [row[33] for row in values_list_Score2]
    combined_data5_2 = result_list_Score2.copy()
    for i in range(len(combined_data5_2)):
        combined_data5_2[i].append(total2[i])

    ps2 = [row[34] for row in values_list_Score2]
    combined_data6_2 = result_list_Score2.copy()
    for i in range(len(combined_data6_2)):
        combined_data6_2[i].append(ps2[i])


    #giai đoạn 3
    values_list_Score3 = worksheet2.get_all_values()[2:]
    result_list_Score3 = [row[:2] for row in values_list_Score3]
    lop3 = [row[3] for row in values_list_Score3]
    combined_data_3 = result_list_Score3.copy()
    for i in range(len(combined_data_3)):
        combined_data_3[i].append(lop3[i])

    teacher3 = [row[21] for row in values_list_Score3]
    combined_data1_3 = result_list_Score3.copy()
    for i in range(len(combined_data1_3)):
        combined_data1_3[i].append(teacher3[i])
        
    listen3 = [row[35] for row in values_list_Score3]
    combined_data2_3 = result_list_Score3.copy()
    for i in range(len(combined_data2_3)):
        combined_data2_3[i].append(listen3[i])

    speak3 = [row[36] for row in values_list_Score3]
    combined_data3_3 = result_list_Score3.copy()
    for i in range(len(combined_data3_3)):
        combined_data3_3[i].append(speak3[i])

    rw3 = [row[37] for row in values_list_Score3]
    combined_data4_3 = result_list_Score3.copy()
    for i in range(len(combined_data4_3)):
        combined_data4_3[i].append(rw3[i])

    total3 = [row[38] for row in values_list_Score3]
    combined_data5_3 = result_list_Score3.copy()
    for i in range(len(combined_data5_3)):
        combined_data5_3[i].append(total3[i])

    ps3 = [row[39] for row in values_list_Score3]
    combined_data6_3 = result_list_Score3.copy()
    for i in range(len(combined_data6_3)):
        combined_data6_3[i].append(ps3[i])


    #giai đoạn 4
    values_list_Score4 = worksheet2.get_all_values()[2:]
    result_list_Score4 = [row[:2] for row in values_list_Score4]
    lop4 = [row[3] for row in values_list_Score4]
    combined_data_4 = result_list_Score4.copy()
    for i in range(len(combined_data_4)):
        combined_data_4[i].append(lop4[i])

    teacher4 = [row[21] for row in values_list_Score4]
    combined_data1_4 = result_list_Score4.copy()
    for i in range(len(combined_data1_4)):
        combined_data1_4[i].append(teacher4[i])
        
    listen4 = [row[40] for row in values_list_Score4]
    combined_data2_4 = result_list_Score4.copy()
    for i in range(len(combined_data2_4)):
        combined_data2_4[i].append(listen4[i])

    speak4 = [row[41] for row in values_list_Score4]
    combined_data3_4 = result_list_Score4.copy()
    for i in range(len(combined_data3_4)):
        combined_data3_4[i].append(speak4[i])

    rw4 = [row[42] for row in values_list_Score4]
    combined_data4_4 = result_list_Score4.copy()
    for i in range(len(combined_data4_4)):
        combined_data4_4[i].append(rw4[i])

    total4 = [row[43] for row in values_list_Score4]
    combined_data5_4 = result_list_Score4.copy()
    for i in range(len(combined_data5_4)):
        combined_data5_4[i].append(total4[i])

    ps4 = [row[44] for row in values_list_Score4]
    combined_data6_4 = result_list_Score4.copy()
    for i in range(len(combined_data6_4)):
        combined_data6_4[i].append(ps4[i])

    #giai đoạn 5
    values_list_Score5 = worksheet2.get_all_values()[2:]
    result_list_Score5 = [row[:2] for row in values_list_Score5]
    lop5 = [row[3] for row in values_list_Score5]
    combined_data_5 = result_list_Score5.copy()
    for i in range(len(combined_data_5)):
        combined_data_5[i].append(lop5[i])

    teacher5 = [row[21] for row in values_list_Score5]
    combined_data1_5 = result_list_Score5.copy()
    for i in range(len(combined_data1_5)):
        combined_data1_5[i].append(teacher5[i])
        
    listen5 = [row[45] for row in values_list_Score5]
    combined_data2_5 = result_list_Score5.copy()
    for i in range(len(combined_data2_5)):
        combined_data2_5[i].append(listen5[i])

    speak5 = [row[46] for row in values_list_Score5]
    combined_data3_5 = result_list_Score5.copy()
    for i in range(len(combined_data3_5)):
        combined_data3_5[i].append(speak5[i])

    rw5 = [row[47] for row in values_list_Score5]
    combined_data4_5 = result_list_Score5.copy()
    for i in range(len(combined_data4_5)):
        combined_data4_5[i].append(rw5[i])

    total5 = [row[48] for row in values_list_Score5]
    combined_data5_5 = result_list_Score5.copy()
    for i in range(len(combined_data5_5)):
        combined_data5_5[i].append(total5[i])

    ps5 = [row[49] for row in values_list_Score5]
    combined_data6_5 = result_list_Score5.copy()
    for i in range(len(combined_data6_5)):
        combined_data6_5[i].append(ps5[i])

# Call the function to initialize the globals
initialize_globals()


from EXCEL.Excel_creating import Excel_Create

class MainFormGUI:
    def __init__(self):
        self.root = tk.Tk()
        
        # Root window properties
        self.root.title("Main Form GUI")
        #self.root.geometry("1097x700")
        # Set the window to start maximized
        self.root.state('zoomed')    
        self.root.configure(bg="#e0f7fa")
        
        '''
        # Set full screen
        self.root.attributes('-fullscreen', True)
        '''

        self.style = ttk.Style()
        self.style.configure('TFrame', background='#e6e6e6')
        self.style.configure('TButton', background='#cc0000', foreground='#cc0000', font=('Cambria', 12, 'bold'))
        self.style.configure('TLabel', background='#e6e6e6', foreground='#007acc', font=('Cambria', 12, 'bold'))
        self.style.configure('TEntry', background='#007acc', foreground='#007acc', font=('Cambria', 14, 'bold'))
        self.style.configure('TNotebook.Tab', font=('Cambria', 14, 'bold'), background='#007acc', foreground='#007acc')
        self.style.configure('TTreeview.Heading', font=('Cambria', 14, 'bold'), background='#007ACC', foreground='white')
        self.style.configure('TTreeview', font=('Cambria', 14), background='#f5f5f5', foreground='#333333')

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
            "book": [],
            "changeclass": [],
            "reviewclass":[]
        }
        
        # Original data storage
        self.original_data_class = result_list_Class[:]
        self.original_data_student = combined_data_student1[:]
        self.original_data_score1 = combined_data6[:]
        self.original_data_score2 = combined_data6_2[:]
        self.original_data_score3 = combined_data6_3[:]
        self.original_data_score4 = combined_data6_4[:]
        self.original_data_score5 = combined_data6_5[:]
        self.original_data_book = result_list_Book[:]
        self.original_data_changeclass = result_list_changeclass[:]
        self.original_data_reviewclass = result_list_reviewclass[:]

        
        # Class management tab
        self.create_class_management_tab()
        
        # Student management tab
        self.create_student_management_tab()
        
        # Score management tab
        self.create_score_management_tab()
        
        # Book management tab
        self.create_book_management_tab()
        #Change class management tab
        self.create_changeclass_management_tab()

        #Review class management tab
        self.create_reviewclass_management_tab()
        
        # Logout button
        self.content_seach = ttk.Frame(self.root)
        self.content_seach.pack(fill="both", expand=True)
        btnDangXuat = ttk.Button(self.content_seach, text="Đăng xuất", width=25, command=self.dangxuat)
        btnDangXuat.pack(side="right", anchor="ne")
        

    def create_class_management_tab(self):
        self.class_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.class_management_tab, text="Quản lý lớp học")
        
        # Frame chứa các nút
        button_frame = ttk.Frame(self.class_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        # Thêm Label "Quản lý lớp học"
        label_frame = ttk.Label(button_frame, text="Quản lý lớp học", font=("Cambria", 12, "bold"))
        label_frame.pack(side="left", padx=20, pady=5)
        
        # Nút thêm mới, xuất excel và reload
        btnAddNew = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_Class, width=25, style='TButton')
        btnXuatExcel = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("class"), width=25, style='TButton')
        
        btnAddNew.pack(side="right", padx=5, pady=5)
        btnXuatExcel.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame = ttk.Frame(self.class_management_tab, style='TFrame')
        table_frame.pack(fill="both", expand=True)
    
        table_columns = ["Mã lớp", "Lớp chính", "Ngày học", "Thời gian học", "Phòng", "Giáo viên", "Giáo viên nước ngoài"]
    
        self.table = ttk.Treeview(table_frame, columns=table_columns, show="headings", height=28)
        
        # Cấu hình màu nền cho hàng xen kẽ
        self.table.tag_configure('oddrow', background="white")
        self.table.tag_configure('evenrow', background="#E8F6F3")
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "Mã lớp": 50,
            "Lớp chính": 50,
            "Ngày học": 50,
            "Thời gian học": 50,
            "Phòng": 50,
            "Giáo viên": 200,
            "Giáo viên nước ngoài": 200,
        }
        
        for col in table_columns:
            self.table.heading(col, text=col, command=lambda c=col: self.sort_table(self.table, c, False))
            self.table.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
        
        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.populate_table(self.table, self.original_data_class)
        self.table.bind("<Double-1>", self.on_row_select)
        self.create_search_section(self.class_management_tab, "class")


    def create_student_management_tab(self):
        self.student_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.student_management_tab, text="Quản lý học sinh")
        
        # Frame chứa các nút
        button_frame = ttk.Frame(self.student_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        # Thêm Label "Quản lý học sinh"
        label_frame = ttk.Label(button_frame, text="Quản lý học sinh", font=("Cambria", 12, "bold"))
        label_frame.pack(side="left", padx=20, pady=5)
        
        # Nút thêm mới, xuất excel và reload
        btnAddNew1 = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_Student, width=25, style='TButton')
        btnXuatExcel1 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("student"), width=25, style='TButton')
        
        btnAddNew1.pack(side="right", padx=5, pady=5)
        btnXuatExcel1.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame = ttk.Frame(self.student_management_tab, style='TFrame')
        table_frame.pack(fill="both", expand=True)
        
        table_columns1 = ["ID", "Họ và tên", "Ngày sinh", "Lớp chính", "Cấp độ hiện tại", "SĐT", "Địa chỉ"]
        self.table1 = ttk.Treeview(table_frame, columns=table_columns1, show="headings", height=28)
        
        # Cấu hình màu nền cho hàng xen kẽ
        self.table1.tag_configure('oddrow', background="white")
        self.table1.tag_configure('evenrow', background="#FDE8D7")
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 50,
            "Họ và tên": 200,
            "Ngày sinh": 100,
            "Lớp chính": 100,
            "Cấp độ hiện tại": 100,
            "SĐT": 100,
            "Địa chỉ": 250,
        }
        
        for col in table_columns1:
            self.table1.heading(col, text=col, command=lambda c=col: self.sort_table(self.table1, c, False))
            self.table1.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
        
        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y1 = ttk.Scrollbar(table_frame, orient="vertical", command=self.table1.yview)
        tree_scroll_x1 = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table1.xview)
        self.table1.configure(yscrollcommand=tree_scroll_y1.set, xscrollcommand=tree_scroll_x1.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table1.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y1.grid(row=0, column=1, sticky="ns")
        tree_scroll_x1.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.populate_table(self.table1, self.original_data_student)
        self.table1.bind("<Double-1>", self.on_row_select1)
        self.create_search_section(self.student_management_tab, "student")


    def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        
        self.inner_tab_control = ttk.Notebook(self.score_management_tab)
        self.inner_tab_control.pack(expand=1, fill="both")

        self.tab1 = ttk.Frame(self.inner_tab_control, style='TFrame')
        self.inner_tab_control.add(self.tab1, text="Giai đoạn 1")

        self.tab2 = ttk.Frame(self.inner_tab_control, style='TFrame')
        self.inner_tab_control.add(self.tab2, text="Giai đoạn 2")

        self.tab3 = ttk.Frame(self.inner_tab_control, style='TFrame')
        self.inner_tab_control.add(self.tab3, text="Giai đoạn 3")

        self.tab4 = ttk.Frame(self.inner_tab_control, style='TFrame')
        self.inner_tab_control.add(self.tab4, text="Giai đoạn 4")

        self.tab5 = ttk.Frame(self.inner_tab_control, style='TFrame')
        self.inner_tab_control.add(self.tab5, text="Giai đoạn 5")

        #tab 1
        button_frame = ttk.Frame(self.tab1, style='TFrame')
        button_frame.pack(side="top", fill="x")
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("score1"), width=25, style='TButton')
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        
        # Thêm Label "Giai đoạn 1"
        label_frame1 = ttk.Label(button_frame, text="Giai đoạn 1", font=("Cambria", 12, "bold"))
        label_frame1.pack(side="left", padx=20, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame = ttk.Frame(self.tab1, style='TFrame')
        table_frame.pack(fill="both", expand=True)
        
        
        table_columns2 = ["ID", "Họ và tên", "Lớp chính", "Giáo viên", "LISTENING", "SPEAKING", "WRITING & READING", "Tổng điểm", "Phần trăm"]
        column_widths = {
            "ID": 50,
            "Họ và tên": 200,
            "Lớp chính": 100,
            "Giáo viên": 150,
            "LISTENING": 100,
            "SPEAKING": 100,
            "WRITING & READING": 150,
            "Tổng điểm": 100,
            "Phần trăm": 100,
        }
        
        self.table2 = ttk.Treeview(table_frame, columns=table_columns2, show="headings", height=25)
        
        for col in table_columns2:
            self.table2.heading(col, text=col, command=lambda c=col: self.sort_table(self.table2, c, False))
            self.table2.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
            
        # Cấu hình màu nền cho hàng xen kẽ
        self.table2.tag_configure('oddrow', background="white")
        self.table2.tag_configure('evenrow', background="#D8CFE3")
        
        
        for col in table_columns2:
            self.table2.heading(col, text=col)
            self.table2.column(col, width=column_widths.get(col, 100), anchor=tk.W)

        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y1 = ttk.Scrollbar(table_frame, orient="vertical", command=self.table2.yview)
        tree_scroll_x1 = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table2.xview)
        self.table2.configure(yscrollcommand=tree_scroll_y1.set, xscrollcommand=tree_scroll_x1.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table2.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y1.grid(row=0, column=1, sticky="ns")
        tree_scroll_x1.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.populate_table(self.table2, self.original_data_score1)
        self.table2.bind("<Double-1>", self.on_row_select2)
        self.create_search_section(self.tab1, "score1")


        # tab 2
        button_frame_2 = ttk.Frame(self.tab2, style='TFrame')
        button_frame_2.pack(side="top", fill="x")
        btnXuatExcel2_2 = ttk.Button(button_frame_2, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload_2 = ttk.Button(button_frame_2, text="Reload", command=lambda: self.reload_tab("score2"), width=25, style='TButton')
        btnXuatExcel2_2.pack(side="right", padx=5, pady=5)
        btnReload_2.pack(side="right", padx=5, pady=5)
        
        # Thêm Label "Giai đoạn 2"
        label_frame2 = ttk.Label(button_frame_2, text="Giai đoạn 2", font=("Cambria", 12, "bold"))
        label_frame2.pack(side="left", padx=20, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame2 = ttk.Frame(self.tab2, style='TFrame')
        table_frame2.pack(fill="both", expand=True)

        table_columns2_2 = ["ID", "Họ và tên", "Lớp chính", "Giáo viên", "LISTENING", "SPEAKING", "WRITING & READING", "Tổng điểm", "Phần trăm"]
        
        self.table2_2 = ttk.Treeview(table_frame2, columns=table_columns2_2, show="headings", height=25)
        
        
        for col in table_columns2_2:
            self.table2_2.heading(col, text=col, command=lambda c=col: self.sort_table(self.table2_2, c, False))
            self.table2_2.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
            
            
        # Cấu hình màu nền cho hàng xen kẽ
        self.table2_2.tag_configure('oddrow', background="white")
        self.table2_2.tag_configure('evenrow', background="#D5E1EF")
        
        
        for col in table_columns2_2:
            self.table2_2.heading(col, text=col)
            self.table2_2.column(col, width=column_widths.get(col, 100), anchor=tk.W)

        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y2 = ttk.Scrollbar(table_frame2, orient="vertical", command=self.table2_2.yview)
        tree_scroll_x2 = ttk.Scrollbar(table_frame2, orient="horizontal", command=self.table2_2.xview)
        self.table2_2.configure(yscrollcommand=tree_scroll_y2.set, xscrollcommand=tree_scroll_x2.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table2_2.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y2.grid(row=0, column=1, sticky="ns")
        tree_scroll_x2.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame2.grid_rowconfigure(0, weight=1)
        table_frame2.grid_columnconfigure(0, weight=1)
        
        
        self.table2_2.bind("<Double-1>", self.on_row_select2_2)
        self.populate_table(self.table2_2, self.original_data_score2)
        self.create_search_section(self.tab2, "score2")


        #tab 3
        button_frame_3 = ttk.Frame(self.tab3, style='TFrame')
        button_frame_3.pack(side="top", fill="x")
        btnXuatExcel2_3 = ttk.Button(button_frame_3, text="Xuất excel",command=self.XuatExcel12, width=25, style='TButton')
        btnReload_3 = ttk.Button(button_frame_3, text="Reload", command=lambda: self.reload_tab("score3"), width=25, style='TButton')
        btnXuatExcel2_3.pack(side="right", padx=5, pady=5)
        btnReload_3.pack(side="right", padx=5, pady=5)
        
        # Thêm Label "Giai đoạn 3"
        label_frame3 = ttk.Label(button_frame_3, text="Giai đoạn 3", font=("Cambria", 12, "bold"))
        label_frame3.pack(side="left", padx=20, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame3 = ttk.Frame(self.tab3, style='TFrame')
        table_frame3.pack(fill="both", expand=True)

        table_columns2_3 = ["ID", "Họ và tên", "Lớp chính", "Giáo viên", "LISTENING", "SPEAKING", "WRITING & READING", "Tổng điểm", "Phần trăm"]
        self.table2_3 = ttk.Treeview(table_frame3, columns=table_columns2_3, show="headings", height=25)
        
        for col in table_columns2_3:
            self.table2_3.heading(col, text=col, command=lambda c=col: self.sort_table(self.table2_3, c, False))
            self.table2_3.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
            
            
        # Cấu hình màu nền cho hàng xen kẽ
        self.table2_3.tag_configure('oddrow', background="white")
        self.table2_3.tag_configure('evenrow', background="#BDD2F6")
        
        for col in table_columns2_3:
            self.table2_3.heading(col, text=col)
            self.table2_3.column(col, width=column_widths.get(col, 100), anchor=tk.W)

        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y3 = ttk.Scrollbar(table_frame3, orient="vertical", command=self.table2_3.yview)
        tree_scroll_x3 = ttk.Scrollbar(table_frame3, orient="horizontal", command=self.table2_3.xview)
        self.table2_3.configure(yscrollcommand=tree_scroll_y3.set, xscrollcommand=tree_scroll_x3.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table2_3.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y3.grid(row=0, column=1, sticky="ns")
        tree_scroll_x3.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame3.grid_rowconfigure(0, weight=1)
        table_frame3.grid_columnconfigure(0, weight=1)
        
        
        self.table2_3.bind("<Double-1>", self.on_row_select2_3)
        self.populate_table(self.table2_3, self.original_data_score3)
        self.create_search_section(self.tab3, "score3")


        #tab 4
        button_frame_4 = ttk.Frame(self.tab4, style='TFrame')
        button_frame_4.pack(side="top", fill="x")
        btnXuatExcel2_4 = ttk.Button(button_frame_4, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload_4 = ttk.Button(button_frame_4, text="Reload", command=lambda: self.reload_tab("score4"), width=25, style='TButton')
        btnXuatExcel2_4.pack(side="right", padx=5, pady=5)
        btnReload_4.pack(side="right", padx=5, pady=5)
        
        # Thêm Label "Giai đoạn 4"
        label_frame4 = ttk.Label(button_frame_4, text="Giai đoạn 4", font=("Cambria", 12, "bold"))
        label_frame4.pack(side="left", padx=20, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame4 = ttk.Frame(self.tab4, style='TFrame')
        table_frame4.pack(fill="both", expand=True)

        table_columns2_4 = ["ID", "Họ và tên", "Lớp chính", "Giáo viên", "LISTENING", "SPEAKING", "WRITING & READING", "Tổng điểm", "Phần trăm"]
        self.table2_4 = ttk.Treeview(table_frame4, columns=table_columns2_4, show="headings", height=25)
        
        for col in table_columns2_4:
            self.table2_4.heading(col, text=col, command=lambda c=col: self.sort_table(self.table2_4, c, False))
            self.table2_4.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
            
            
        # Cấu hình màu nền cho hàng xen kẽ
        self.table2_4.tag_configure('oddrow', background="white")
        self.table2_4.tag_configure('evenrow', background="#C2D69B")
        
        for col in table_columns2_4:
            self.table2_4.heading(col, text=col)
            self.table2_4.column(col, width=column_widths.get(col, 100), anchor=tk.W)


        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y4 = ttk.Scrollbar(table_frame4, orient="vertical", command=self.table2_4.yview)
        tree_scroll_x4 = ttk.Scrollbar(table_frame4, orient="horizontal", command=self.table2_4.xview)
        self.table2_4.configure(yscrollcommand=tree_scroll_y4.set, xscrollcommand=tree_scroll_x4.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table2_4.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y4.grid(row=0, column=1, sticky="ns")
        tree_scroll_x4.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame4.grid_rowconfigure(0, weight=1)
        table_frame4.grid_columnconfigure(0, weight=1)
      
        self.table2_4.bind("<Double-1>", self.on_row_select2_4)
        self.populate_table(self.table2_4, self.original_data_score4)
        self.create_search_section(self.tab4, "score4")


        # tab 5
        button_frame_5 = ttk.Frame(self.tab5, style='TFrame')
        button_frame_5.pack(side="top", fill="x")
        btnXuatExcel2_5 = ttk.Button(button_frame_5, text="Xuất excel", command=self.XuatExcel12, width=25, style='TButton')
        btnReload_5 = ttk.Button(button_frame_5, text="Reload", command=lambda: self.reload_tab("score5"), width=25, style='TButton')
        btnXuatExcel2_5.pack(side="right", padx=5, pady=5)
        btnReload_5.pack(side="right", padx=5, pady=5)
        
        # Thêm Label "Giai đoạn 1"
        label_frame5 = ttk.Label(button_frame_5, text="Giai đoạn 5", font=("Cambria", 12, "bold"))
        label_frame5.pack(side="left", padx=20, pady=5)

        # Tạo frame cho bảng và thanh cuộn
        table_frame5 = ttk.Frame(self.tab5, style='TFrame')
        table_frame5.pack(fill="both", expand=True)
        
        table_columns2_5 = ["ID", "Họ và tên", "Lớp chính", "Giáo viên", "LISTENING", "SPEAKING", "WRITING & READING", "Tổng điểm", "Phần trăm"]
        self.table2_5 = ttk.Treeview(table_frame5, columns=table_columns2_5, show="headings", height=25)
        
        for col in table_columns2_5:
            self.table2_5.heading(col, text=col, command=lambda c=col: self.sort_table(self.table2_5, c, False))
            self.table2_5.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
            
            
        # Cấu hình màu nền cho hàng xen kẽ
        self.table2_5.tag_configure('oddrow', background="white")
        self.table2_5.tag_configure('evenrow', background="#FDE8D7")
        
        for col in table_columns2_5:
            self.table2_5.heading(col, text=col)
            self.table2_5.column(col, width=column_widths.get(col, 100), anchor=tk.W)
     
        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y5 = ttk.Scrollbar(table_frame5, orient="vertical", command=self.table2_5.yview)
        tree_scroll_x5 = ttk.Scrollbar(table_frame5, orient="horizontal", command=self.table2_5.xview)
        self.table2_5.configure(yscrollcommand=tree_scroll_y5.set, xscrollcommand=tree_scroll_x5.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table2_5.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y5.grid(row=0, column=1, sticky="ns")
        tree_scroll_x5.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame5.grid_rowconfigure(0, weight=1)
        table_frame5.grid_columnconfigure(0, weight=1)
        
        self.table2_5.bind("<Double-1>", self.on_row_select2_5)
        self.populate_table(self.table2_5, self.original_data_score5)
        self.create_search_section(self.tab5, "score5")


    def create_book_management_tab(self):
        self.book_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.book_management_tab, text="Quản lý sách")
        
        button_frame = ttk.Frame(self.book_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        btnAddNew3 = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_Book, width=25, style='TButton')
        btnXuatExcel3 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel3, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("book"), width=25, style='TButton')
        
        btnAddNew3.pack(side="right", padx=5, pady=5)
        btnXuatExcel3.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        # Create a frame for the table and scrollbars
        table_frame = ttk.Frame(self.book_management_tab, style='TFrame')
        table_frame.pack(fill="both", expand=True)
        
        table_columns3 = ['ID', "Tên lớp chính", "Sách 1", "Sách 2","Sách 3",'Sách 4', 'Sách 5', 'Tên giáo viên', 'Tên giáo viên nước ngoài']
        self.table3 = ttk.Treeview(table_frame, columns=table_columns3, show="headings", height=28)
        column_widths = {
            "ID": 3,
            "Tên lớp chính": 20, 
            "Sách 1":30, 
            "Sách 2":30,
            "Sách 3":30,
            'Sách 4':30, 
            'Sách 5':30, 
            'Tên giáo viên': 60, 
            'Tên giáo viên nước ngoài': 60
        }
        # Add sorting capability
        self.table3 = ttk.Treeview(table_frame, columns=table_columns3, show="headings", height=25)
        
        for col in table_columns3:
            self.table3.heading(col, text=col, command=lambda c=col: self.sort_table(self.table3, c, False))
            self.table3.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
        
        # Scrollbars within the table frame
        tree_scroll_y3 = ttk.Scrollbar(table_frame, orient="vertical", command=self.table3.yview)
        tree_scrollx3 = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table3.xview)
        self.table3.configure(yscrollcommand=tree_scroll_y3.set, xscrollcommand=tree_scrollx3.set)

        # Pack the table and scrollbars within the table frame
        self.table3.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y3.grid(row=0, column=1, sticky="ns")
        tree_scrollx3.grid(row=1, column=0, sticky="ew")

        # Make the table frame expandable
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.populate_table(self.table3, self.original_data_book)
        self.table3.bind("<Double-1>", self.on_row_select3)
        
        self.create_search_section(self.book_management_tab, "book")
    def create_changeclass_management_tab(self):
        self.changeclass_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.changeclass_management_tab, text="Quản lý chuyển lớp")
        
        # Frame chứa các nút
        button_frame = ttk.Frame(self.changeclass_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        # Thêm Label "Quản lý lớp học"
        label_frame = ttk.Label(button_frame, text="Quản lý chuyển lớp", font=("Cambria", 12, "bold"))
        label_frame.pack(side="left", padx=20, pady=5)
        
        # Nút thêm mới, xuất excel và reload
        btnAddNew4 = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_ClassChange, width=25, style='TButton')
        btnXuatExcel4 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel4, width=25, style='TButton')
        btnReload4 = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("changeclass"), width=25, style='TButton')
        
        btnAddNew4.pack(side="right", padx=5, pady=5)
        btnXuatExcel4.pack(side="right", padx=5, pady=5)
        btnReload4.pack(side="right", padx=5, pady=5)


        
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame = ttk.Frame(self.changeclass_management_tab, style='TFrame')
        table_frame.pack(fill="both", expand=True)



    
        table_columns = ['ID', "Mã học sinh","Họ và tên", "SĐT", "Tên lớp chính","Tên lớp chuyển",'Lý do chuyển', 'Thời gian bắt đầu học']
        self.table4 = ttk.Treeview(table_frame, columns=table_columns, show="headings", height=28)
        
        # Cấu hình màu nền cho hàng xen kẽ
        self.table4.tag_configure('oddrow', background="white")
        self.table4.tag_configure('evenrow', background="#E8F6F3")
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 3,
            "Mã học sinh": 90,
            "Họ và tên": 120,
            "SĐT": 50,
            "Tên lớp chính": 170,
            "Tên lớp chuyển": 170,
            "Lý do chuyển": 100,
            "Thời gian bắt đầu học": 50,
        }
        
        for col in table_columns:
            self.table4.heading(col, text=col, command=lambda c=col: self.sort_table(self.table4, c, False))
            self.table4.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
        
        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.table4.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table4.xview)
        self.table4.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table4.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.populate_table(self.table4, self.original_data_changeclass)
        self.table4.bind("<Double-1>", self.on_row_select4)
        self.create_search_section(self.changeclass_management_tab, "changeclass")
    

    def create_reviewclass_management_tab(self):
        self.reviewclass_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.reviewclass_management_tab, text="Quản lý lớp ôn")
        
        # Frame chứa các nút
        button_frame = ttk.Frame(self.reviewclass_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        
        # Thêm Label "Quản lý học sinh"
        label_frame = ttk.Label(button_frame, text="Quản lý lớp ôn", font=("Cambria", 12, "bold"))
        label_frame.pack(side="left", padx=20, pady=5)
        
        # Nút thêm mới, xuất excel và reload
        btnAddNew1 = ttk.Button(button_frame, text="Thêm mới", command=self.AddGUI_ReviewClass, width=25, style='TButton')
        btnXuatExcel1 = ttk.Button(button_frame, text="Xuất excel", command=self.XuatExcel5, width=25, style='TButton')
        btnReload = ttk.Button(button_frame, text="Reload", command=lambda: self.reload_tab("reviewclass"), width=25, style='TButton')
        
        btnAddNew1.pack(side="right", padx=5, pady=5)
        btnXuatExcel1.pack(side="right", padx=5, pady=5)
        btnReload.pack(side="right", padx=5, pady=5)
        
        # Tạo frame cho bảng và thanh cuộn
        table_frame = ttk.Frame(self.reviewclass_management_tab, style='TFrame')
        table_frame.pack(fill="both", expand=True)
        
        table_columns1 = ['ID','Mã học sinh', "Họ và tên", "SĐT", "Tên lớp chính", "Tên lớp ôn"]
        self.table5 = ttk.Treeview(table_frame, columns=table_columns1, show="headings", height=28)
        
        # Cấu hình màu nền cho hàng xen kẽ
        self.table5.tag_configure('oddrow', background="white")
        self.table5.tag_configure('evenrow', background="#FDE8D7")
        
        # Đặt tiêu đề và độ rộng cho các cột
        column_widths = {
            "ID": 5,
            "Mã học sinh": 90,
            "Họ và tên": 150,
            "SĐT": 150,
            "Tên lớp chính": 90,
            "Tên lớp ôn": 90,
        }
        
        for col in table_columns1:
            self.table5.heading(col, text=col, command=lambda c=col: self.sort_table(self.table5, c, False))
            self.table5.column(col, width=column_widths.get(col, 100), anchor=tk.W)  # Đặt độ rộng theo cấu hình
        
        # Thêm thanh cuộn vào trong table_frame
        tree_scroll_y1 = ttk.Scrollbar(table_frame, orient="vertical", command=self.table5.yview)
        tree_scroll_x1 = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table5.xview)
        self.table5.configure(yscrollcommand=tree_scroll_y1.set, xscrollcommand=tree_scroll_x1.set)
        
        # Đặt bảng và thanh cuộn vào grid
        self.table5.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y1.grid(row=0, column=1, sticky="ns")
        tree_scroll_x1.grid(row=1, column=0, sticky="ew")
        
        # Đảm bảo frame của bảng có thể mở rộng
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.populate_table(self.table5, self.original_data_reviewclass)
        self.table5.bind("<Double-1>", self.on_row_select5)
        self.create_search_section(self.reviewclass_management_tab, "reviewclass")


    def sort_table(self, table, col, reverse):
        if col == 'ID' or col == 'CLASSNO':
            def convert(value):
                """Try to convert the value to an integer if possible, otherwise return the string."""
                try:
                    return int(value)
                except ValueError:
                    return value

            # Extract data from the table
            l = [(convert(table.set(k, col)), k) for k in table.get_children('')]

            # Sort based on the data type (numerical vs string) and reverse option
            l.sort(reverse=reverse)

            # Re-arrange the rows based on the sorted order
            for index, (val, k) in enumerate(l):
                table.move(k, '', index)

            # Update column heading to toggle sorting direction on click
            table.heading(col, command=lambda: self.sort_table(table, col, not reverse))
            
        else:
            l = [(table.set(k, col), k) for k in table.get_children('')]
            l.sort(reverse=reverse)

            for index, (val, k) in enumerate(l):
                table.move(k, '', index)

            table.heading(col, command=lambda: self.sort_table(table, col, not reverse))

        


    def create_search_section(self, tab, type_):
        if type_ == "class":
            fields = ["Giáo viên", "Phòng", "Lớp chính", "Mã lớp"]
        elif type_ == "student":
            fields = ["Lớp chính", "Họ và tên", "ID"]
        elif type_ == "score1":
            fields = ["Giáo viên", "Lớp chính", "Họ và tên", "ID"]
        elif type_ == "score2":
            fields = ["Giáo viên", "Lớp chính", "Họ và tên", "ID"]
        elif type_ == "score3":
            fields = ["Giáo viên", "Lớp chính", "Họ và tên", "ID"]
        elif type_ == "score4":
            fields = ["Giáo viên", "Lớp chính", "Họ và tên", "ID"]
        elif type_ == "score5":
            fields = ["Giáo viên", "Lớp chính", "Họ và tên", "ID"]
        elif type_ == "book":
            fields = ["Tên lớp chính", "Sách",'Tên giáo viên']
        elif type_ == "changeclass":
            fields = ["Họ và tên", "Mã học sinh",'Tên lớp chính', 'Tên lớp chuyển']
        elif type_ == "reviewclass":
            fields = ["Họ và tên","Mã học sinh", "Tên lớp chính", "Tên lớp ôn"]
            
        
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
                "Giáo viên": 5,
                "Phòng": 4,
                "Lớp chính": 1,
                "Mã lớp": 0
            },
            "student": {
                "Lớp chính": 3,
                "Họ và tên": 1,
                "ID": 0
            },
            "score1": {
                "Giáo viên": 3,
                "Lớp chính": 2,
                "Họ và tên": 1,
                "ID": 0
            },
             "score2": {
                "Giáo viên": 3,
                "Lớp chính": 2,
                "Họ và tên": 1,
                "ID": 0
            },
              "score3": {
                "Giáo viên": 3,
                "Lớp chính": 2,
                "Họ và tên": 1,
                "ID": 0
            },
               "score4": {
                "Giáo viên": 3,
                "Lớp chính": 2,
                "Họ và tên": 1,
                "ID": 0
            },
                "score5": {
                "Giáo viên": 3,
                "Lớp chính": 2,
                "Họ và tên": 1,
                "ID": 0
            },
            "book": {
                "Tên lớp chính": 1, 
                "Sách": [2, 3, 4, 5, 6],
                'Tên giáo viên': [7, 8],
            },
            "changeclass": {
                "Họ và tên": 2,
                "Mã học sinh": 1,
                "Tên lớp chính": 4,
                "Tên lớp chuyển": 5,
            },
            "reviewclass": {
                "Họ và tên": 2,
                "Mã học sinh": 1, 
                "Tên lớp chính": 4,
                "Tên lớp ôn": 5,
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
        elif type_ == "score1":
            data_source = self.original_data_score1
            table = self.table2
        elif type_ == "score2":
            data_source = self.original_data_score2
            table = self.table2_2
        elif type_ == "score3":
            data_source = self.original_data_score3
            table = self.table2_3
        elif type_ == "score4":
            data_source = self.original_data_score4
            table = self.table2_4
        elif type_ == "score5":
            data_source = self.original_data_score5
            table = self.table2_5
        elif type_ == "book":
            data_source = self.original_data_book
            table = self.table3
        elif type_ == "changeclass":
            data_source = self.original_data_changeclass
            table = self.table4
        elif type_ == "reviewclass":
            data_source = self.original_data_reviewclass
            table = self.table5

        # for row in data_source:
        #     if all(search_criteria[f'tf{index+1}'] in row[mapping[field]].lower() for index, field in enumerate(mapping)):
        #         matching_rows.append(row)
        for row in data_source:
            match = True
            for index, field in enumerate(mapping):
                if isinstance(mapping[field], list):
                    # If the mapping is a list, check all columns in the list
                    if not any(search_criteria[f'tf{index+1}'] in row[col].lower() for col in mapping[field]):
                        match = False
                        break
                else:
                    # If the mapping is a single column, check that column
                    if search_criteria[f'tf{index+1}'] not in row[mapping[field]].lower():
                        match = False
                        break
            if match:
                matching_rows.append(row)
        
        self.populate_table(table, matching_rows)
        
        if not matching_rows:
            messagebox.showinfo("Thông báo", "Không tìm thấy dữ liệu khớp với các tiêu chí tìm kiếm đã nhập")

    def populate_table(self, table, data):
        # Xóa tất cả các hàng hiện có trong bảng
        for row in table.get_children():
            table.delete(row)

        # Thêm dữ liệu mới vào bảng với màu sắc xen kẽ
        for i, row in enumerate(data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            table.insert("", "end", values=row, tags=(tag,))


    def reload_tab(self, type_):
        initialize_globals()  # Refresh the global data
        self.update_original_data()  # Update the instance variables with new data
        if type_ == "class":
            self.populate_table(self.table, self.original_data_class)
        elif type_ == "student":
            self.populate_table(self.table1, self.original_data_student)
        elif type_ == "score1":
            self.populate_table(self.table2, self.original_data_score1)
        elif type_ == "score2":
            self.populate_table(self.table2_2, self.original_data_score2)
        elif type_ == "score3":
            self.populate_table(self.table2_3, self.original_data_score3)
        elif type_ == "score4":
            self.populate_table(self.table2_4, self.original_data_score4)
        elif type_ == "score5":
            self.populate_table(self.table2_5, self.original_data_score5)
        elif type_ == "book":
            self.populate_table(self.table3, self.original_data_book)
        elif type_ == "changeclass":
            self.populate_table(self.table4, self.original_data_changeclass)
        elif type_ == "reviewclass":
            self.populate_table(self.table5, self.original_data_reviewclass)

        # Reset search entries
        for entry in self.entries[type_].values():
            if isinstance(entry, ttk.Entry):
                entry.delete(0, tk.END)

    def update_original_data(self):
        self.original_data_class = result_list_Class
        self.original_data_student = result_list_Student
        self.original_data_score1 = result_list_Score
        self.original_data_score2 = result_list_Score2
        self.original_data_score3 = result_list_Score3
        self.original_data_score4 = result_list_Score4
        self.original_data_score5 = result_list_Score5
        self.original_data_book = result_list_Book
        self.original_data_changeclass = result_list_changeclass
        self.original_data_reviewclass = result_list_reviewclass


    
    def run(self):
        self.root.mainloop()

    def dangxuat(self):
        self.root.destroy()
    
    
    def XuatExcel(self):
        Xuat1 = Excel_Create()
        Xuat1.XuatExcel()
    def XuatExcel4(self):
        Xuat4 = Excel_Create()
        Xuat4.XuatExcel4()
    def XuatExcel5(self):
        Xuat5 = Excel_Create()
        Xuat5.XuatExcel5()
        
    def XuatExcel12(self):
        Xuat2 = Excel_Create()
        Xuat2.XuatExcel12()
    
    def XuatExcel3(self):
        Xuat3 = Excel_Create()
        Xuat3.XuatExcel3()
        
        
    def AddGUI_Class(self):
        AddNewClass = Add_NewClass(self)
        AddNewClass.run()  # Chạy giao diện
    
    def AddGUI_Book(self):
        AddNewBook = Add_NewSach()
        AddNewBook.run()     
    
    def AddGUI_Student(self):
        AddNewStudent = Add_NewStudent(self)
        AddNewStudent.run()

    def AddGUI_ClassChange(self):
        AddNewChangeClass = Add_NewChangeClass(self)
        AddNewChangeClass.run()

    def AddGUI_ReviewClass(self):
        AddNewReviewClass = Add_NewReviewClass(self)
        AddNewReviewClass.run()
        

    def center_window(self, width, height, object):
        window_width = width
        window_height = height

        # Lấy kích thước màn hình
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Tính toán vị trí x và y để cửa sổ xuất hiện ở giữa màn hình
        position_x = int((screen_width / 2) - (window_width / 2))
        position_y = int((screen_height / 2) - (window_height / 2))

        # Đặt lại vị trí cho cửa sổ
        object.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

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
        #self.rootClass.geometry("520x680")
        self.center_window(520,680,self.rootClass)
        self.canvas1 = tk.Canvas(self.rootClass, width=self.rootClass.winfo_screenwidth(), height=self.rootClass.winfo_screenheight())
        self.canvas1.pack(fill=tk.BOTH, expand=True)
        self.panel1 = tk.Frame(self.canvas1, bd=4, relief="solid")
        self.panel1.place(x=10, y=10, width=500, height=650)
        self.lbl_addNewClass = tk.Label(self.panel1, text="Edit class", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewClass.place(x=180, y=10)
        self.lb1 = tk.Label(self.panel1, text="Lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=430, height=30)
        self.lb2 = tk.Label(self.panel1, text="Ngày học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=33, y=171)
        self.tf2 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf2.place(x=33, y=224, width=430, height=30)
        self.lb3 = tk.Label(self.panel1, text="Giờ học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=290)
        self.tf3 = tk.Entry(self.panel1, font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=340, width=200, height=30)
        self.lb4 = tk.Label(self.panel1, text="Phòng", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=260, y=290)
        self.tf4 = tk.Entry(self.panel1, font=("cambria", 13, "bold"))
        self.tf4.place(x=260, y=340, width=200, height=30)
        self.lb5 = tk.Label(self.panel1, text="Giáo viên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=400)
        self.tf5 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=445, width=430, height=30)
        self.lb6 = tk.Label(self.panel1, text="Giáo viên nước ngoài", font=("cambria", 18, "bold"), fg="#FBA834")
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
                self.rootClass.destroy()
                self.reload_tab("class")
                messagebox.showinfo("Thành công", "Cập nhật thành công!")

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
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                # mapping2 = {}
                # for i in range(1, n2 + 1):
                #     mapping2[i] = letters2[(i - 1) % 26]
                mapping2 = {}
                for i in range(1, n + 1):
                    # Sử dụng phép chia dư để lặp lại các chữ cái từ A đến Z
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.Edit_NewStudent(row_data2,vitrisua2)

        else:
            print("Value not found in the sheet.")
            
    def on_row_select2(self, event):
        selected_item2 = self.table2.selection()
        if selected_item2:
            row_values2 = self.table2.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row
                # row_data2 = worksheet2.row_values(matched_row2)
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                # mapping2 = {}
                # for i in range(1, n2 + 1):
                #     mapping2[i] = letters2[(i - 1) % 26]
                mapping2 = {}
                for i in range(1, n + 1):
                    # Sử dụng phép chia dư để lặp lại các chữ cái từ A đến Z
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.editScore(row_data2,vitrisua2)
        else:
            print("Value not found in the sheet.")
            
    def editScore(self,row_data4,vitrisua4):
        self.rootScore = tk.Tk()
        self.rootScore.title("Edit Score")
        #self.rootScore.geometry("1000x700")
        self.center_window(1000,700,self.rootScore)
        self.canvas4 = tk.Canvas(self.rootScore, width=self.rootScore.winfo_screenwidth(), height=self.rootScore.winfo_screenheight())
        self.canvas4.pack(fill=tk.BOTH, expand=True)
        self.panel4 = tk.Frame(self.canvas4, bd=4, relief="solid")
        self.panel4.place(x=10, y=10, width=980, height=670)
        self.lbl_editScore = tk.Label(self.panel4, text="Edit Score", font=("cambria", 24, "bold"), fg="black")
        self.lbl_editScore.place(x=450, y=10)
        self.tfname = tk.Entry(self.panel4,font=("cambria", 16, "bold"), state='readonly')
        self.tfname.place(x=33, y=10, width=300, height=30)
        self.lb1 = tk.Label(self.panel4, text="Exam invigilator", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        
        self.lb2 = tk.Label(self.panel4, text="Exam day", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        
        '''
        self.tf2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf2.place(x=370, y=108, width=300, height=30)
        '''
        # Sử dụng DateEntry thay vì Entry để chọn ngày
        self.tf2 = DateEntry(self.panel4, font=("Cambria", 13, "bold"), date_pattern='dd/mm/yyyy', 
                             background='darkblue', foreground='white', borderwidth=2)
        self.tf2.place(x=370, y=108, width=300, height=30)


        self.lb13 = tk.Label(self.panel4, text="Exam time", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb13.place(x=700, y=60)
        self.tf13 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf13.place(x=700, y=108, width=240, height=30)


        #giai đoạn 1
        self.lb3 = tk.Label(self.panel4, text="Giai đoạn 1:", font=("cambria", 18, "bold"))
        self.lb3.place(x=33, y=160)
        self.lb4 = tk.Label(self.panel4, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=200, y=160)
        self.lbLis1 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbLis1.place(x=265, y=213)
        self.tf1_1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1_1.place(x=200, y=213, width=65, height=30)
        self.lb31 = tk.Label(self.panel4, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb31.place(x=370, y=160)
        self.lbSpeak1 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbSpeak1.place(x=435, y=213)
        self.tf1_2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1_2.place(x=370, y=213, width=65, height=30)
        self.lb15 = tk.Label(self.panel4, text="Reading and Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb15.place(x=540, y=160)
        self.lbRW1 = tk.Label(self.panel4, text="/25", font=("cambria", 18, "bold"))
        self.lbRW1.place(x=605, y=213)
        self.tf1_3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1_3.place(x=540, y=213, width=65, height=30)
        

        #Giai đoạn 2
        self.lb5 = tk.Label(self.panel4, text="Giai đoạn 2: ", font=("cambria", 18, "bold"))
        self.lb5.place(x=33, y=270)
        self.lb6 = tk.Label(self.panel4, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=200, y=270)
        self.lbLis2 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbLis2.place(x=265, y=320)
        self.tf2_1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf2_1.place(x=200, y=320, width=65, height=30)

        self.lb32 = tk.Label(self.panel4, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb32.place(x=370, y=270)
        self.lbSpeak2 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbSpeak2.place(x=435, y=320)
        self.tf2_2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf2_2.place(x=370, y=320, width=65, height=30)

        self.lb16 = tk.Label(self.panel4, text="Reading and Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb16.place(x=540, y=270)
        self.lbRW2 = tk.Label(self.panel4, text="/25", font=("cambria", 18, "bold"))
        self.lbRW2.place(x=605, y=320)
        self.tf2_3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf2_3.place(x=540, y=320, width=65, height=30)
        
        #giai đoạn 3
        self.lb7 = tk.Label(self.panel4, text="Giai đoạn 3:", font=("cambria", 18, "bold"))
        self.lb7.place(x=33, y=375)
        self.lb8 = tk.Label(self.panel4, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=200, y=375)
        self.lbLis3 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbLis3.place(x=265, y=420)
        self.tf3_1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf3_1.place(x=200, y=420, width=65, height=30)

        self.lb30 = tk.Label(self.panel4, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb30.place(x=370, y=375)
        self.lbSpeak3 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbSpeak3.place(x=435, y=420)
        self.tf3_2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf3_2.place(x=370, y=420, width=65, height=30)

        self.lb17 = tk.Label(self.panel4, text="Reading and Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb17.place(x=540, y=375)
        self.lbRW3 = tk.Label(self.panel4, text="/25", font=("cambria", 18, "bold"))
        self.lbRW3.place(x=605, y=420)
        self.tf3_3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf3_3.place(x=540, y=420, width=65, height=30)
        
        #Giai đoạn 4
        self.lb9 = tk.Label(self.panel4, text="Giai đoạn 4", font=("cambria", 18, "bold"))
        self.lb9.place(x=33, y=480)
        
        self.lb10 = tk.Label(self.panel4, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=200, y=480)
        self.lbLis4 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbLis4.place(x=265, y=520)
        self.tf4_1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf4_1.place(x=200, y=520, width=65, height=30)
        
        self.lb11 = tk.Label(self.panel4, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb11.place(x=370, y=480)
        self.lbSpeak4 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbSpeak4.place(x=435, y=520)
        self.tf4_2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf4_2.place(x=370, y=520, width=65, height=30)
        
        self.lb12 = tk.Label(self.panel4, text="Reading and Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb12.place(x=540, y=480)
        self.lbRW4 = tk.Label(self.panel4, text="/25", font=("cambria", 18, "bold"))
        self.lbRW4.place(x=605, y=520)
        self.tf4_3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf4_3.place(x=540, y=520, width=65, height=30)
        
        #Giai đoạn 5
        self.lb33 = tk.Label(self.panel4, text="Giai đoạn 5", font=("cambria", 18, "bold"))
        self.lb33.place(x=33, y=580)
        
        self.lb34 = tk.Label(self.panel4, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb34.place(x=200, y=580)
        self.lbLis5 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbLis5.place(x=265, y=620)
        self.tf5_1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf5_1.place(x=200, y=620, width=65, height=30)

        self.lb35 = tk.Label(self.panel4, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb35.place(x=370, y=580)
        self.lbSpeak5 = tk.Label(self.panel4, text="/20", font=("cambria", 18, "bold"))
        self.lbSpeak5.place(x=435, y=620)
        self.tf5_2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf5_2.place(x=370, y=620, width=65, height=30)
        
        self.lb36 = tk.Label(self.panel4, text="Reading and Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb36.place(x=540, y=580)
        self.lbRW5 = tk.Label(self.panel4, text="/25", font=("cambria", 18, "bold"))
        self.lbRW5.place(x=605, y=620)
        self.tf5_3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf5_3.place(x=540, y=620, width=65, height=30)


        self.tfname.config(state='normal')
        self.tfname.insert(0, "Name: "+row_data4[1])
        self.tfname.config(state='readonly')
        self.tf1.insert(0, row_data4[24])
        
        self.tf2.delete(0, tk.END)  # Xóa mọi giá trị trong entry
        self.tf2.insert(0, row_data4[22])
        
        self.tf13.insert(0, row_data4[23])
        self.tf1_1.insert(0, row_data4[25])
        self.tf1_2.insert(0, row_data4[26])
        self.tf1_3.insert(0, row_data4[27])
        self.tf2_1.insert(0, row_data4[30])
        self.tf2_2.insert(0, row_data4[31])
        self.tf2_3.insert(0, row_data4[32])
        self.tf3_1.insert(0, row_data4[35])
        self.tf3_2.insert(0, row_data4[36])
        self.tf3_3.insert(0, row_data4[37])
        self.tf4_1.insert(0, row_data4[40])
        self.tf4_2.insert(0, row_data4[41])
        self.tf4_3.insert(0, row_data4[42])
        self.tf5_1.insert(0, row_data4[45])
        self.tf5_2.insert(0, row_data4[46])
        self.tf5_3.insert(0, row_data4[47])
        
        def get_value():
            global a1, a2, a3
            global a4, a5, a6,total_giaidoan1,percent_giaidoan1
            global a7, a8, a9, total_giaidoan2,percent_giaidoan2
            global a10, a11, a12, total_giaidoan3,percent_giaidoan3
            global a13, a14, a15 ,total_giaidoan4,percent_giaidoan4
            global a16, a17, a18, total_giaidoan5,percent_giaidoan5
            
            a1 = self.tf1.get()
            a2 = self.tf2.get()
            a3 = self.tf13.get()
            #giai đoạn 1
            try:
                a4 = int(self.tf1_1.get())
            except ValueError:
                a4 = 0
            try:
                a5 = int(self.tf1_2.get())
            except ValueError:
                a5 = 0
            try:
                a6 = int(self.tf1_3.get())
            except ValueError:
                a6 = 0
            total_giaidoan1 = a4 + a5 + a6
            percent_giaidoan1 = str(round((total_giaidoan1/65)*100))+"%"

            #giai đoạn 2
            try:
                a7 = int(self.tf2_1.get())
            except ValueError:
                a7 = 0
            try:
                a8 = int(self.tf2_2.get())
            except ValueError:
                a8 = 0
            try:
                a9 = int(self.tf2_3.get())
            except ValueError:
                a9 = 0
            total_giaidoan2 = a7 + a8 + a9
            percent_giaidoan2 = str(round((total_giaidoan2/65)*100))+"%"

            #giai đoạn 3
            try:
                a10 = int(self.tf3_1.get())
            except ValueError:
                a10 = 0
            try:
                a11 = int(self.tf3_2.get())
            except ValueError:
                a11 = 0
            try:
                a12 = int(self.tf3_3.get())
            except ValueError:
                a12 = 0
            total_giaidoan3 = a10 + a11 + a12
            percent_giaidoan3 = str(round((total_giaidoan3/65)*100))+"%"

            #giai đoạn 4
            try:
                a13 = int(self.tf4_1.get())
            except ValueError:
                a13 = 0
            try:
                a14 = int(self.tf4_2.get())
            except ValueError:
                a14 = 0
            try:
                a15 = int(self.tf4_3.get())
            except ValueError:
                a15 = 0
            total_giaidoan4 = a13 + a14 + a15
            percent_giaidoan4 = str(round((total_giaidoan4/65)*100))+"%"

            #giai đoạn 5
            try:
                a16 = int(self.tf5_1.get())
            except ValueError:
                a16 = 0
            try:
                a17 = int(self.tf5_2.get())
            except ValueError:
                a17 = 0
            try:
                a18 = int(self.tf5_3.get())
            except ValueError:
                a18 = 0
            total_giaidoan5 = a16 + a17 + a18
            percent_giaidoan5 = str(round((total_giaidoan5/65)*100))+"%"
            
            
        def chinhsua():
            get_value()

            new_values4 = [int(row_data4[0]),row_data4[1],row_data4[2],row_data4[3],row_data4[4],row_data4[5],row_data4[6],row_data4[7],row_data4[8],row_data4[9],row_data4[10],row_data4[11],row_data4[12],row_data4[13],row_data4[14],row_data4[15],row_data4[16],row_data4[17],row_data4[18],row_data4[19],row_data4[20],row_data4[21], 
                           a2, a3, a1, a4, 
                           a5,a6,total_giaidoan1,percent_giaidoan1,
                           a7,a8,a9,total_giaidoan2,percent_giaidoan2,
                           a10,a11,a12,total_giaidoan3,percent_giaidoan3,
                           a13,a14,a15,total_giaidoan4,percent_giaidoan4,
                           a16,a17,a18,total_giaidoan5,percent_giaidoan5]

            try:
                worksheet2.update(values=[new_values4], range_name=vitrisua4)
                self.rootScore.destroy()
                self.reload_tab("score1")
                self.reload_tab("score2")
                self.reload_tab("score3")
                self.reload_tab("score4")
                self.reload_tab("score5")

                messagebox.showinfo("Thành công", "Cập nhật thành công!")


            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")            
                

        def stage_definition(level, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18):
            result = None  # Biến lưu trữ tạm thời kết quả
            
            if (a4 is not None and a4 != 0) or (a5 is not None and a5 != 0) or (a6 is not None and a6 != 0):
                result = level + ' ' + '1/5'
                
            if (a7 is not None and a7 != 0) or (a8 is not None and a8 != 0) or (a9 is not None and a9 != 0):
                result = level + ' ' + '2/5'
                
            if (a10 is not None and a10 != 0) or (a11 is not None and a11 != 0) or (a12 is not None and a12 != 0):
                result = level + ' ' + '3/5'
                
            if (a13 is not None and a13 != 0) or (a14 is not None and a14 != 0) or (a15 is not None and a15 != 0):
                result = level + ' ' + '4/5'
                
            if (a16 is not None and a16 != 0) or (a17 is not None and a17 != 0) or (a18 is not None and a18 != 0):
                result = level + ' ' + '5/5'
            
            # Nếu có một if nào thỏa mãn thì trả về result, nếu không thì trả về 'Complete'
            return result if result else level + ' ' + 'Unidentified'


        
        def print_pdf():                
            try:
                # Lấy giá trị từ text box ở đây
                get_value()

                level = row_data4[4]
                address = 'Con Ong Thông Minh'
                stage = stage_definition(level, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18)
                exam_type = 'Camp'
                teacher = row_data4[21]
                study_date = row_data4[22]
                study_time = row_data4[23]
                name = row_data4[1]
                birth = row_data4[2]
                class_no = row_data4[3]
                
                '''
                def create_file(pdf, level, address, exam_date, stage, exam_type, exam_time, main_teacher, examiner_teacher, study_date, study_time, 
                 name, birth, class_no, 
                 stage1, listening1, reading1, speaking1, stage2, listening2, reading2, speaking2, stage3, listening3, reading3, speaking3, 
                 stage4, listening4, reading4, speaking4, stage5, listening5, reading5, speaking5, 
                 stage_no1, stage_no2, stage_no3, stage_no4, stage_no5):
                '''
                
                create_file('pdf', level, address, a2, stage, exam_type, a3, teacher , a1, study_date, study_time,
                            name, birth, class_no,
                            'Giai đoạn 1', a4, a5, a6, 'Giai đoạn 2', a7, a8, a9, 'Giai đoạn 3', a10, a11, a12, 
                            'Giai đoạn 4', a13, a14, a15, 'Giai đoạn 5', a16, a17, a18,
                            1,2,3,4,5)
                '''
                create_file('pdf', 'M', 'C.HB406','11/09/2024','23','Reading','15h30','Nguyễn Trung Sơn','Phạm Hoàng Dũng','M-W-F','19h20', 
                 'Phạm Hoàng Dũng', '15/06/2003', '8A11', 
                 'Giai đoạn 1', 10, 10, 10, 'Giai đoạn 2', 10, 20, 24, 'Giai đoạn 3', 10, 20, 20, 
                 'Giai đoạn 4', 10, 10, 25, 'Giai đoạn 5', 10, 20, 25,
                 1,2,3,4,5)
                '''
                
                messagebox.showinfo("Thành công", "In file PDF thành công!")
            
            except Exception as e:
                messagebox.showerror("Lỗi", f"In file PDF thất bại: {e}")
    
            
        self.btn1 = tk.Button(self.panel4, text="SUBMIT", font=("cambria", 14, "bold"),command=chinhsua ,width=12, bg="#FBA834",fg="black" )
        self.btn1.place(x=800, y=330)
        
        self.btn2 = tk.Button(self.panel4, text="PDF", font=("cambria", 14, "bold"),command=print_pdf ,width=12, bg="#FF0000",fg="white" )
        self.btn2.place(x=800, y=530)
        

    def on_row_select2_2(self, event):
        selected_item2 = self.table2_2.selection()
        if selected_item2:
            row_values2 = self.table2_2.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                mapping2 = {}
                for i in range(1, n + 1):
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.editScore(row_data2,vitrisua2)
        else:
            print("Value not found in the sheet.")
    def on_row_select2_3(self, event):
        selected_item2 = self.table2_3.selection()
        if selected_item2:
            row_values2 = self.table2_3.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                mapping2 = {}
                for i in range(1, n + 1):
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.editScore(row_data2,vitrisua2)
        else:
            print("Value not found in the sheet.")
            
    def on_row_select2_4(self, event):
        selected_item2 = self.table2_4.selection()
        if selected_item2:
            row_values2 = self.table2_4.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                mapping2 = {}
                for i in range(1, n + 1):
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.editScore(row_data2,vitrisua2)
        else:
            print("Value not found in the sheet.")
            
    def on_row_select2_5(self, event):
        selected_item2 = self.table2_5.selection()
        if selected_item2:
            row_values2 = self.table2_5.item(selected_item2, "values")
            row_list2 = row_values2[0] 
            if row_list2 in worksheet2.col_values(1):
                vitribandau2 = "A"+str(worksheet2.find(row_values2[0]).row)
                matched_row2 = worksheet2.find(row_values2[0]).row
                row_data2 = worksheet2.row_values(matched_row2)
                if len(row_data2)<=49:
                    row_data2.extend([""] * (49 - len(row_data2) + 1))
                char_to_num = dict()
                count_values2 = len(row_data2)
                n = 50  # Tăng giá trị n lên 50
                letters2 = string.ascii_uppercase
                mapping2 = {}
                for i in range(1, n + 1):
                    first_letter = letters2[(i - 1) // 26 - 1] if i > 26 else ''
                    second_letter = letters2[(i - 1) % 26]
                    mapping2[i] = first_letter + second_letter
                vitrisua2 = vitribandau2+":"+mapping2[count_values2]+str(matched_row2)
            # print(vitrisua2)
            self.editScore(row_data2,vitrisua2)
        else:
            print("Value not found in the sheet.")
            
    def on_combobox_select(self, event = None):
            selected_value = self.tf13.get()
            return selected_value
             
    def Edit_NewStudent(self,row_data3,vitrisua3):
        self.rootStudent = tk.Tk()
        self.rootStudent.title("Edit student and point")
        #self.rootStudent.geometry("1300x680")
        self.center_window(1300,680,self.rootStudent)
        self.canvas2 = tk.Canvas(self.rootStudent, width=self.root.winfo_screenwidth(), height=self.rootStudent.winfo_screenheight())
        self.canvas2.pack(fill=tk.BOTH, expand=True)
        self.panel2 = tk.Frame(self.canvas2, bd=4, relief="solid")
        self.panel2.place(x=10, y=10, width=1275, height=650)
        self.lbl_EditNewStudent = tk.Label(self.panel2, text="Edit student and point", font=("cambria", 24, "bold"), fg="black")
        self.lbl_EditNewStudent.place(x=450, y=10)
        self.lb1 = tk.Label(self.panel2, text="Họ và tên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        
        self.lb2 = tk.Label(self.panel2, text="Ngày sinh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        
        '''
        self.tf2 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf2.place(x=370, y=108, width=300, height=30)
        '''
        # Entry ngày sinh với DateEntry từ tkcalendar
        self.tf2 = DateEntry(self.panel2, font=("Cambria", 13, "bold"), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.tf2.place(x=370, y=108, width=300, height=30)
        

        self.lb3 = tk.Label(self.panel2, text="Địa chỉ", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel2, text="Tháng bắt đầu nghỉ", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=370, y=160)
        self.tf4 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf4.place(x=370, y=213, width=300, height=30)

        self.lb5 = tk.Label(self.panel2, text="Trường", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=300, height=30)

        self.lb6 = tk.Label(self.panel2, text="Tháng chuyển lớp", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=370, y=270)
        self.tf6 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf6.place(x=370, y=320, width=300, height=30)
        
        self.lb7 = tk.Label(self.panel2, text="Tên phụ huynh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=300, height=30)

        self.lb8 = tk.Label(self.panel2, text="New Comer", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=370, y=375)
        self.tf8 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf8.place(x=370, y=420, width=300, height=30)

        self.lb9 = tk.Label(self.panel2, text="SĐT", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=130, height=30)

        self.lb10 = tk.Label(self.panel2, text="Đăng ký cơ sở", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb10.place(x=200, y=480)
        self.tf10 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf10.place(x=200, y=530, width=130, height=30)

        self.lb11 = tk.Label(self.panel2, text="Cơ sở chính", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb11.place(x=370, y=480)
        self.tf11 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf11.place(x=370, y=530, width=130, height=30)

        self.lb12 = tk.Label(self.panel2, text="Tổng học phí", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb12.place(x=540, y=480)
        self.tf12 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf12.place(x=540, y=530, width=130, height=30)



        self.lb13 = tk.Label(self.panel2, text="Lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
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

        self.lb16 = tk.Label(self.panel2, text="Giáo viên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb16.place(x=700, y=270)
        self.tf16 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf16.place(x=700, y=320, width=240, height=30)

        self.lb17 = tk.Label(self.panel2, text="SĐT thêm", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb17.place(x=700, y=375)
        self.tf17 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf17.place(x=700, y=420, width=240, height=30)


        self.lb18 = tk.Label(self.panel2, text="Học phí chính", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb18.place(x=700, y=480)
        self.tf18 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf18.place(x=700, y=530, width=130, height=30)


        self.lb19 = tk.Label(self.panel2, text="Chứng chỉ", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb19.place(x=850, y=480)
        self.tf19 = tk.Entry(self.panel2, font=("cambria", 13, "bold"))
        self.tf19.place(x=850, y=530, width=130, height=30)


        self.lb20 = tk.Label(self.panel2, text="Cấp độ hiện tại", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb20.place(x=1000, y=160)
        self.tf20 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf20.place(x=1000, y=213, width=150, height=30)

        self.lb21 = tk.Label(self.panel2, text="Ngày học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb21.place(x=1000, y=270)
        self.tf21 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf21.place(x=1000, y=320, width=150, height=30)

        self.lb22 = tk.Label(self.panel2, text="Giờ học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb22.place(x=1000, y=375)
        self.tf22 = tk.Entry(self.panel2,font=("cambria", 13, "bold"))
        self.tf22.place(x=1000, y=420, width=150, height=30)

        
        self.tf1.insert(0, row_data3[1])
        
        self.tf2.delete(0, tk.END)  # Xóa mọi giá trị trong entry
        self.tf2.insert(0, row_data3[2])  # Chèn giá trị từ row_data3[2]

        self.tf13.delete(0, tk.END)  # Xóa mọi giá trị trong entry
        self.tf13.insert(0, row_data3[3])
        
        self.tf20.insert(0, row_data3[4])
        self.tf21.insert(0, row_data3[5])
        self.tf22.insert(0, row_data3[6])
        self.tf9.insert(0, row_data3[7])
        self.tf3.insert(0, row_data3[8])
        self.tf7.insert(0, row_data3[9])
        self.tf10.insert(0, row_data3[10])
        self.tf11.insert(0, row_data3[11])
        self.tf12.insert(0, row_data3[12])
        self.tf18.insert(0, row_data3[13])
        self.tf8.insert(0, row_data3[14])
        self.tf4.insert(0, row_data3[15])
        self.tf15.insert(0, row_data3[16])
        self.tf19.insert(0, row_data3[17])
        self.tf5.insert(0, row_data3[18])
        self.tf17.insert(0, row_data3[19])
        self.tf6.insert(0, row_data3[20])
        self.tf16.insert(0, row_data3[21]) 
          
        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf2.get()
            a3 = self.on_combobox_select()
            a4 = self.tf20.get()
            a5 = self.tf21.get()
            a6 = self.tf22.get()
            a7 = self.tf9.get()
            a8 = self.tf3.get()
            a9 = self.tf7.get()
            a10 = self.tf10.get()
            a11 = self.tf11.get()
            a12 = self.tf12.get()
            a13 = self.tf18.get()
            a14 = self.tf8.get()
            a15 = self.tf4.get()
            a16 = self.tf15.get()
            a17 = self.tf19.get()
            a18 = self.tf5.get()
            a19 = self.tf17.get()
            a20 = self.tf6.get()
            a21 = self.tf16.get()
            new_values3 = [int(row_data3[0]),a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,row_data3[22],row_data3[23],row_data3[24], row_data3[25], row_data3[26],row_data3[27],row_data3[28],row_data3[29],row_data3[30],row_data3[31],row_data3[32],row_data3[33],row_data3[34],row_data3[35],row_data3[36],row_data3[37],row_data3[38],row_data3[39],row_data3[40],row_data3[41],row_data3[42],row_data3[43],row_data3[44],row_data3[45],row_data3[46],row_data3[47],row_data3[48],row_data3[49]]
            try:
                worksheet2.update(values=[new_values3], range_name=vitrisua3)
                self.rootStudent.destroy()
                self.reload_tab("student")
                messagebox.showinfo("Thành công", "Cập nhật thành công!")

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
            self.Edit_NewSach(row_data3,vitrisua3)
        else:
            print("Value not found in the sheet.")
    
    def Edit_NewSach(self,row_data3,vitrisua3):
        self.rootSach = tk.Tk('EDIT')
        self.rootSach.title()
        self.rootSach.geometry("735x550")
        self.canvas4 = tk.Canvas(self.rootSach, width=self.rootSach.winfo_screenwidth(), height=self.rootSach.winfo_screenheight())
        self.canvas4.pack(fill=tk.BOTH, expand=True)
        self.panel4 = tk.Frame(self.canvas4, bd=4, relief="solid")
        self.panel4.place(x=10, y=10, width=710, height=530)
        self.lbl_addNewBook = tk.Label(self.panel4, text="Chỉnh sửa thông tin sách", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=190, y=10)
        self.lb1 = tk.Label(self.panel4, text="Tên lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel4, text="Sách 1", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=380, y=60)
        self.tf2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf2.place(x=380, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel4, text="Sách 2", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel4, text="Sách 3", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=380, y=160)
        self.tf4 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf4.place(x=380, y=213, width=300, height=30)

        self.lb9 = tk.Label(self.panel4, text="Sách 4", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=270)
        self.tf9 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=320, width=300, height=30)
        

        self.lb10 = tk.Label(self.panel4, text="Sách 5", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=380, y=270)
        self.tf10 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf10.place(x=380, y=320, width=300, height=30)

        self.lb11 = tk.Label(self.panel4, text="Giáo viên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb11.place(x=33, y=380)
        self.tf11 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf11.place(x=33, y=430, width=300, height=30)

        self.lb12 = tk.Label(self.panel4, text="Giáo viên nước ngoài", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb12.place(x=380, y=380)
        self.tf12 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf12.place(x=380, y=430, width=300, height=30)

        

        self.tf1.insert(0, row_data3[1])
        self.tf2.insert(0, row_data3[2])
        self.tf3.insert(0, row_data3[3])
        self.tf4.insert(0, row_data3[4])
        self.tf9.insert(0, row_data3[5])
        self.tf10.insert(0, row_data3[6])
        self.tf11.insert(0, row_data3[7])
        self.tf12.insert(0, row_data3[8])


        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf2.get()
            a3 = self.tf3.get()
            a4 = self.tf4.get()
            a9 = self.tf9.get()
            a10 = self.tf10.get()
            a11 = self.tf11.get()
            a12 = self.tf12.get()
            worksheet6 = sht.worksheet("sheet 1")
            test = worksheet6.get_all_values()[2:]
            existing_cls = [row[1].lower() for row in test]
            if a1 in existing_cls:
                messagebox.showerror("Error", "Lớp này đã có thông tin sách")
                self.tf1.delete(0, 'end')
            elif a1 == "":
                messagebox.showerror("Error", "Bạn chưa nhập tên lớp")
            elif a1 not in existing_cls:
                new_values3 = [int(row_data3[0]),a1, a2,a3,a4,a9,a10,a11,a12]
                # worksheet3.update(values=[new_values], range_name=vitrisua)
                try:
                    worksheet6.update(values=[new_values3], range_name=vitrisua3)
                    messagebox.showinfo("Thành công", "Cập nhật thành công!")
                    self.rootSach.destroy()

                except Exception as e:
                    messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel4, text="Xác nhận", font=("cambria", 14, "bold"),command=chinhsua, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=250, y=470)


    def on_row_select4(self, event):
        selected_item3 = self.table4.selection()
        if selected_item3:
            row_values3 = self.table4.item(selected_item3, "values")
            row_list3 = row_values3[0] 
            if row_list3 in worksheet5.col_values(1):
                vitribandau3 = "A"+str(worksheet5.find(row_values3[0]).row)
                matched_row3 = worksheet5.find(row_values3[0]).row

                # count_values3 = len(worksheet.row_values(matched_row3))
                row_data3 = worksheet5.row_values(matched_row3)
                if len(row_data3)<=7:
                    row_data3.extend([""] * (7 - len(row_data3) + 1))
                char_to_num = dict()
                count_values3 = len(row_data3)
                letters3 = [chr(i) for i in range(65, 91)]
                n3 = 30
                mapping3 = {}
                for i in range(1, n3 + 1):
                    mapping3[i] = letters3[(i - 1) % 26]
                vitrisua3 = vitribandau3+":"+mapping3[count_values3]+str(matched_row3)
            self.Edit_NewChangeClass(row_data3,vitrisua3)
        else:
            print("Value not found in the sheet.")

    def Edit_NewChangeClass(self,row_data3,vitrisua3):
        self.rootChangeClass = tk.Tk('EDIT')
        self.rootChangeClass.title()
        self.rootChangeClass.geometry("735x540")
        self.canvas1 = tk.Canvas(self.rootChangeClass, width=self.rootChangeClass.winfo_screenwidth(), height=self.rootChangeClass.winfo_screenheight())
        self.canvas1.pack(fill=tk.BOTH, expand=True)
        
        self.panel1 = tk.Frame(self.canvas1, bd=4, relief="solid")
        self.panel1.place(x=10, y=10, width=710, height=520)
        self.lbl_addNewBook = tk.Label(self.panel1, text="Sửa thông tin chuyển lớp", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=160, y=10)
        self.lb1 = tk.Label(self.panel1, text="Họ và tên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel1, text="Mã học sinh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=380, y=60)
        self.tf2 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf2.place(x=380, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel1, text="Tên lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel1, text="Tên lớp chuyển", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=380, y=160)
        self.tf4 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf4.place(x=380, y=213, width=300, height=30)

        self.lb9 = tk.Label(self.panel1, text="Lý do chuyển lớp", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=270)
        self.tf9 = tk.Text(self.panel1,font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=320, width=300, height=50)

        self.lb10 = tk.Label(self.panel1, text="Thời gian bắt đầu", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=380, y=270)
        self.tf10 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf10.place(x=380, y=320, width=300, height=30)
        
        self.lb5 = tk.Label(self.panel1, text="Số điện thoại", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=390)
        self.tf5 = tk.Entry(self.panel1,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=430, width=300, height=30)

        self.tf1.insert(0, row_data3[2])
        self.tf2.insert(0, row_data3[1])
        self.tf3.insert(0, row_data3[4])
        self.tf4.insert(0, row_data3[5])
        self.tf9.insert("1.0", row_data3[6])
        self.tf10.insert(0, row_data3[7])
        self.tf5.insert(0, row_data3[3])

        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf2.get()
            a3 = self.tf3.get()
            a4 = self.tf4.get()
            a9 = self.tf9.get("1.0", "end-1c")
            a10 = self.tf10.get()
            a5 = self.tf5.get()
            worksheet5 = sht.worksheet("sheet 5")
            test = worksheet5.get_all_values()[2:]
            existing_code = [row[1] for row in test]
            existing_name = [row2[2].lower() for row2 in test]
            # if a2 in existing_code or a1.lower() in existing_name:
            def custom_title(s):
                return ' '.join(word.capitalize() for word in s.split())
            output_string = custom_title(a1.lower())
            new_values3 = [int(row_data3[0]),a2,output_string.strip(),a5,a3,a4,a9,a10]
                # worksheet3.update(values=[new_values], range_name=vitrisua)
            try:
                worksheet5.update(values=[new_values3], range_name=vitrisua3)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")   
                self.rootChangeClass.destroy()

            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel1, text="Edit", font=("cambria", 14, "bold"),command=chinhsua, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=250, y=470)
    

    def on_row_select5(self, event):
        selected_item3 = self.table5.selection()
        if selected_item3:
            row_values3 = self.table5.item(selected_item3, "values")
            row_list3 = row_values3[0] 
            if row_list3 in worksheet4.col_values(1):
                vitribandau3 = "A"+str(worksheet4.find(row_values3[0]).row)
                matched_row3 = worksheet4.find(row_values3[0]).row

                # count_values3 = len(worksheet.row_values(matched_row3))
                row_data3 = worksheet4.row_values(matched_row3)
                if len(row_data3)<=5:
                    row_data3.extend([""] * (5 - len(row_data3) + 1))
                char_to_num = dict()
                count_values3 = len(row_data3)
                letters3 = [chr(i) for i in range(65, 91)]
                n3 = 30
                mapping3 = {}
                for i in range(1, n3 + 1):
                    mapping3[i] = letters3[(i - 1) % 26]
                vitrisua3 = vitribandau3+":"+mapping3[count_values3]+str(matched_row3)
            self.Edit_ReviewClass(row_data3,vitrisua3)
        else:
            print("Value not found in the sheet.")

    
    def Edit_ReviewClass(self,row_data3,vitrisua3):
        self.rootBook = tk.Tk('EDIT')
        self.rootBook.title()
        self.rootBook.geometry("735x430")
        self.canvas3 = tk.Canvas(self.rootBook, width=self.rootBook.winfo_screenwidth(), height=self.rootBook.winfo_screenheight())
        self.canvas3.pack(fill=tk.BOTH, expand=True)
        self.panel3 = tk.Frame(self.canvas3, bd=4, relief="solid")
        self.panel3.place(x=10, y=10, width=710, height=410)
        self.lbl_addNewBook = tk.Label(self.panel3, text="Chỉnh sửa thông tin lớp ôn", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=200, y=10)
        self.lb1 = tk.Label(self.panel3, text="Họ và tên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel3, text="Mã học sinh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=380, y=60)
        self.tf2 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf2.place(x=380, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel3, text="Tên lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel3, text="Tên lớp ôn", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=380, y=160)
        self.tf4 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf4.place(x=380, y=213, width=300, height=30)

        self.lb5 = tk.Label(self.panel3, text="Số điện thoại", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel3,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=315, width=300, height=30)


        self.tf1.insert(0, row_data3[2])
        self.tf2.insert(0, row_data3[1])
        self.tf3.insert(0, row_data3[4])
        self.tf4.insert(0, row_data3[5])
        self.tf5.insert(0, row_data3[3])
        
        def chinhsua():
            a1 = self.tf1.get()
            a2 = self.tf2.get()
            a3 = self.tf3.get()
            a4 = self.tf4.get()
            a5 = self.tf5.get()
            worksheet4 = sht.worksheet("sheet 4")
            test = worksheet4.get_all_values()[2:]
            existing_code = [row[1] for row in test]
            existing_name = [row2[2].lower() for row2 in test]
            # if a2 in existing_code or a1.lower() in existing_name:
            
            def custom_title(s):
                return ' '.join(word.capitalize() for word in s.split())
            output_string = custom_title(a1.lower())
            new_values3 = [int(row_data3[0]),a2,output_string.strip(),a5 ,a3,a4]
            # worksheet3.update(values=[new_values], range_name=vitrisua)
            try:
                worksheet4.update(values=[new_values3], range_name=vitrisua3)
                messagebox.showinfo("Thành công", "Cập nhật thành công!")
                self.rootBook.destroy()

            except Exception as e:
                messagebox.showerror("Lỗi", f"Cập nhật thất bại: {e}")
        self.btn1 = tk.Button(self.panel3, text="Xác nhận", font=("cambria", 14, "bold"),command=chinhsua, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=250, y=360)

if __name__ == "__main__":
    app = MainFormGUI()
    app.run()