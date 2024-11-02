import tkinter as tk
from tkinter import messagebox
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
class Add_NewChangeClass:
    def __init__(self,parent):
        self.parent = parent  # Reference to MainFormGUI

        self.root = tk.Tk('ADD')
        self.root.title()
        self.root.geometry("735x540")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=710, height=520)
        self.lbl_addNewBook = tk.Label(self.panel, text="Thêm thông tin chuyển lớp", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=160, y=10)
        self.lb1 = tk.Label(self.panel, text="Họ và tên", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel, text="Mã học sinh", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=380, y=60)
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=380, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel, text="Tên lớp chính", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel, text="Tên lớp chuyển", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=380, y=160)
        self.tf4 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf4.place(x=380, y=213, width=300, height=30)

        self.lb9 = tk.Label(self.panel, text="Lý do chuyển lớp", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=270)
        self.tf9 = tk.Text(self.panel,font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=320, width=300, height=50)

        self.lb10 = tk.Label(self.panel, text="Thời gian bắt đầu", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=380, y=270)
        self.tf10 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf10.place(x=380, y=320, width=300, height=30)

        self.lb5 = tk.Label(self.panel, text="Số điện thoại", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=390)
        self.tf5 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=430, width=300, height=30)
        

        self.btn1 = tk.Button(self.panel, text="Xác nhận", font=("cambria", 14, "bold"),command=self.Add_ChangeClass, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=250, y=470)

    def Add_ChangeClass(self):
        worksheet5 = sht.worksheet("sheet 5")
        test = worksheet5.get_all_values()[2:]
        values = worksheet5.col_values(1)[2:]
        max_value = max(list(map(int, values))) + 1  
        cl =  self.tf1.get().lower()
        mb = self.tf2.get()
        sk1 = self.tf3.get()
        sk2 = self.tf4.get()
        reason = self.tf9.get("1.0", "end-1c")
        sd = self.tf10.get()
        tl = self.tf5.get()
        existing_name = [row2[2].lower() for row2 in test]
        existing_code = [row[1] for row in test]
        # if mb in existing_code or cl.lower() in existing_name:
        if cl.lower() in existing_name:
            messagebox.showerror("Error", "Người này đã được lưu")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')
        # elif cl == "" or mb == '':
        elif cl == "":
            messagebox.showerror("Error", "Bạn chưa nhập tên hoặc mã")
        # elif mb not in existing_code or cl.lower() not in existing_name:
        elif cl.lower() not in existing_name:
            def custom_title(s):
                return ' '.join(word.capitalize() for word in s.split())
            output_string = custom_title(cl)
            new_row_values = [max_value,mb,output_string,tl,sk1,sk2,reason,sd]
            worksheet5.append_row(new_row_values, value_input_option='RAW')
            messagebox.showinfo("Success", "Lưu thành công!")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')
            self.tf3.delete(0, 'end')
            self.tf4.delete(0, 'end')
            self.tf9.delete("1.0", "end")
            self.tf10.delete(0, 'end')
            self.tf5.delete(0, 'end')
        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewChangeClass()
    app.run()
