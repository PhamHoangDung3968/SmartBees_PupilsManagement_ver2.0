import tkinter as tk
from tkinter import messagebox
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
class Add_NewSach:
    def __init__(self):
        self.root = tk.Tk('ADD')
        self.root.title()
        self.root.geometry("735x550")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=710, height=530)
        self.lbl_addNewBook = tk.Label(self.panel, text="Thêm thông tin sách", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=190, y=10)
        self.lb1 = tk.Label(self.panel, text="Tên lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel, text="Sách 1", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=380, y=60)
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=380, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel, text="Sách 2", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel, text="Sách 3", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=380, y=160)
        self.tf4 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf4.place(x=380, y=213, width=300, height=30)

        self.lb9 = tk.Label(self.panel, text="Sách 4", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=270)
        self.tf9 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=320, width=300, height=30)
        

        self.lb10 = tk.Label(self.panel, text="Sách 5", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=380, y=270)
        self.tf10 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf10.place(x=380, y=320, width=300, height=30)

        self.lb11 = tk.Label(self.panel, text="Giáo viên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb11.place(x=33, y=380)
        self.tf11 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf11.place(x=33, y=430, width=300, height=30)

        self.lb12 = tk.Label(self.panel, text="Giáo viên nước ngoài", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb12.place(x=380, y=380)
        self.tf12 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf12.place(x=380, y=430, width=300, height=30)

        self.btn1 = tk.Button(self.panel, text="Xác nhận", font=("cambria", 14, "bold"),command=self.Add_Sach, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=250, y=470)

    def Add_Sach(self):
        worksheet6 = sht.worksheet("sheet 1")
        test = worksheet6.get_all_values()[2:]
        values = worksheet6.col_values(1)[2:]
        max_value = max(list(map(int, values))) + 1  
        mainclass =  self.tf1.get()
        book1 = self.tf2.get()
        book2 = self.tf3.get()
        book3 = self.tf4.get()
        book4 = self.tf9.get()
        book5 = self.tf10.get()
        teacher = self.tf11.get()
        fteacher = self.tf12.get()
        existing_cls = [row[1].lower() for row in test]
        if mainclass in existing_cls:
            messagebox.showerror("Error", "Lớp này đã có thông tin sách")
            self.tf1.delete(0, 'end')
        elif mainclass == "":
            messagebox.showerror("Error", "Bạn chưa nhập tên lớp")
        elif mainclass not in existing_cls:
            new_row_values = [max_value,mainclass,book1,book2,book3,book4,book5,teacher,fteacher]
            worksheet6.append_row(new_row_values, value_input_option='RAW')
            messagebox.showinfo("Success", "Lưu thành công!")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')
            self.tf3.delete(0, 'end')
            self.tf4.delete(0, 'end')
            self.tf9.delete(0, "end")
            self.tf10.delete(0, 'end')
            self.tf11.delete(0, 'end')
            self.tf12.delete(0, 'end')
        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewSach()
    app.run()
