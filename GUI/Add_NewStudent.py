import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter import ttk

import gspread

gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
worksheet3 = sht.worksheet("sheet 3")
values_list_Class = worksheet3.get_all_values()[2:]
result_list_Class = [row[1] for row in values_list_Class]

class Add_NewStudent:
    def __init__(self, parent):
        self.parent = parent  # Reference to MainFormGUI

        self.root = tk.Tk()
        self.root.title("Add new student manager")
        #self.root.geometry("1300x680")
        self.center_window(1300,680)
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=1275, height=650)
        self.lbl_addNewBook = tk.Label(self.panel, text="Add new student manager", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=450, y=10)
        self.lb1 = tk.Label(self.panel, text="Họ và tên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel, text="Ngày sinh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        '''
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=370, y=108, width=300, height=30)
        '''
        # Entry ngày sinh với DateEntry từ tkcalendar
        self.tf2 = DateEntry(self.panel, font=("Cambria", 13, "bold"), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.tf2.place(x=370, y=108, width=300, height=30)

        self.lb3 = tk.Label(self.panel, text="Địa chỉ", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel, text="Tháng bắt đầu nghỉ", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=370, y=160)
        self.tf4 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf4.place(x=370, y=213, width=300, height=30)

        self.lb5 = tk.Label(self.panel, text="Trường", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=300, height=30)

        self.lb6 = tk.Label(self.panel, text="Tháng chuyển lớp", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=370, y=270)
        self.tf6 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf6.place(x=370, y=320, width=300, height=30)
        
        self.lb7 = tk.Label(self.panel, text="Tên phụ huynh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=300, height=30)

        self.lb8 = tk.Label(self.panel, text="New Comer", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=370, y=375)
        self.tf8 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf8.place(x=370, y=420, width=300, height=30)

        self.lb9 = tk.Label(self.panel, text="SĐT", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=130, height=30)

        self.lb10 = tk.Label(self.panel, text="Đăng ký cơ sở", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb10.place(x=200, y=480)
        self.tf10 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf10.place(x=200, y=530, width=130, height=30)

        self.lb11 = tk.Label(self.panel, text="Cơ sở chính", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb11.place(x=370, y=480)
        self.tf11 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf11.place(x=370, y=530, width=130, height=30)

        self.lb12 = tk.Label(self.panel, text="Tổng học phí", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb12.place(x=540, y=480)
        self.tf12 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf12.place(x=540, y=530, width=130, height=30)



        self.lb13 = tk.Label(self.panel, text="Lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb13.place(x=700, y=60)
        self.tf13 = ttk.Combobox(self.panel, font=("cambria", 13, "bold"))
        self.tf13['values'] = result_list_Class
        self.tf13.current(0)
        self.tf13.place(x=700, y=108, width=240, height=30)

        self.lb15 = tk.Label(self.panel, text="Starting quit month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb15.place(x=700, y=160)
        self.tf15 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf15.place(x=700, y=213, width=240, height=30)

        self.lb16 = tk.Label(self.panel, text="Giáo viên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb16.place(x=700, y=270)
        self.tf16 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf16.place(x=700, y=320, width=240, height=30)

        self.lb17 = tk.Label(self.panel, text="SĐT thêm", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb17.place(x=700, y=375)
        self.tf17 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf17.place(x=700, y=420, width=240, height=30)


        self.lb18 = tk.Label(self.panel, text="Học phí chính", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb18.place(x=700, y=480)
        self.tf18 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf18.place(x=700, y=530, width=130, height=30)


        self.lb19 = tk.Label(self.panel, text="Chứng chỉ", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb19.place(x=850, y=480)
        self.tf19 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf19.place(x=850, y=530, width=130, height=30)


        self.lb20 = tk.Label(self.panel, text="Cấp độ hiện tại", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb20.place(x=1000, y=160)
        self.tf20 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf20.place(x=1000, y=213, width=150, height=30)

        self.lb21 = tk.Label(self.panel, text="Ngày học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb21.place(x=1000, y=270)
        self.tf21 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf21.place(x=1000, y=320, width=150, height=30)

        self.lb22 = tk.Label(self.panel, text="Thời gian học", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb22.place(x=1000, y=375)
        self.tf22 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf22.place(x=1000, y=420, width=150, height=30)

        
        # self.btn1 = tk.Button(self.panel, text="ADD NEW", font=("cambria", 14, "bold"), width=20, bg="#FBA834",fg="black" )
        self.btn1 = tk.Button(self.panel, text="Thêm mới", font=("cambria", 14, "bold"),command=self.Add_Student, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=520, y=600)
        # Gắn sự kiện nhấn phím Enter với hàm Add_Student
        self.root.bind('<Return>', self.on_enter_key)

    def Add_Student(self):
        worksheet2 = sht.worksheet("sheet 2")
        test = worksheet2.get_all_values()
        values = worksheet2.col_values(1)[2:]
        max_value = max(list(map(int, values))) + 1  
        fn = self.tf1.get()
        b = self.tf2.get()
        stc = self.tf7.get()
        a = self.tf3.get()
        som = self.tf4.get()
        ps = self.tf5.get()
        stm = self.tf6.get()
        nc = self.tf8.get()
        t = self.tf9.get()
        e = self.tf10.get()
        mc = self.tf11.get()
        tf = self.tf12.get()
        mcla = self.tf13.get()
        sqm = self.tf15.get()
        tea = self.tf16.get()
        st = self.tf17.get()
        mf = self.tf18.get()
        c = self.tf19.get()
        cl = self.tf20.get()
        sd = self.tf21.get()
        hours = self.tf22.get()
        # try:
        #     rw = float(self.tf20.get())
        # except ValueError:
        #     rw = 0
        # try:
        #     lis = float(self.tf21.get())
        # except ValueError:
        #     lis = 0

        # try:
        #     spe = float(self.tf22.get())
        # except ValueError:
        #     spe = 0
        # total = rw + lis + spe
        # percent = str(round((total/15)*100,2))+"%"
        
        existing_fn = [row[:3] for row in test]
        if fn in existing_fn:
            messagebox.showerror("Error", "Học sinh này đã được lưu")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')

        elif fn == "":
            messagebox.showerror("Error", "Bạn chưa nhập cấp độ")
        else:
            new_row_values = [max_value, fn, b, mcla, cl, sd, hours, t, a, stc, e, mc, tf, mf, nc, som, sqm, c, ps, st, stm, tea] + [''] * 28
            worksheet2.append_row(new_row_values, value_input_option='RAW')
            
            # Call reload_tab of MainFormGUI to refresh the interface
            self.parent.reload_tab(type_="student")
            
            messagebox.showinfo("Success", "Lưu thành công!")
            
            
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')
            self.tf3.delete(0, 'end')
            self.tf4.delete(0, 'end')
            self.tf5.delete(0, 'end')
            self.tf6.delete(0, 'end')
            self.tf7.delete(0, 'end')
            self.tf8.delete(0, 'end')
            self.tf9.delete(0, 'end')
            self.tf10.delete(0, 'end')
            self.tf11.delete(0, 'end')
            self.tf12.delete(0, 'end')
            self.tf13.delete(0, 'end')
            self.tf15.delete(0, 'end')
            self.tf16.delete(0, 'end')
            self.tf17.delete(0, 'end')
            self.tf18.delete(0, 'end')
            self.tf19.delete(0, 'end')
            self.tf20.delete(0, 'end')
            self.tf21.delete(0, 'end')
            self.tf22.delete(0, 'end')
            
    
    def on_enter_key(self, event):
        # Gọi cùng hàm xử lý khi phím Enter được nhấn
        self.Add_Student()
            
    def center_window(self, width, height):
        window_width = width
        window_height = height

        # Lấy kích thước màn hình
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Tính toán vị trí x và y để cửa sổ xuất hiện ở giữa màn hình
        position_x = int((screen_width / 2) - (window_width / 2))
        position_y = int((screen_height / 2) - (window_height / 2))

        # Đặt lại vị trí cho cửa sổ
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    
        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewStudent()
    app.run()
