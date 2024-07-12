import tkinter as tk
from tkinter import messagebox
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
class Add_NewClass:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Add new class manager")
        self.root.geometry("520x680")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=500, height=650)
        self.lbl_addNewClass = tk.Label(self.panel, text="Add new class manager", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewClass.place(x=80, y=10)
        self.lb1 = tk.Label(self.panel, text="Main class", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=430, height=30)
        self.lb2 = tk.Label(self.panel, text="STUDYING DAY", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=33, y=171)
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=33, y=224, width=430, height=30)
        self.lb3 = tk.Label(self.panel, text="STUDYING TIME", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=290)
        self.tf3 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=340, width=200, height=30)
        self.lb4 = tk.Label(self.panel, text="ROOM", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=260, y=290)
        self.tf4 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf4.place(x=260, y=340, width=200, height=30)
        self.lb5 = tk.Label(self.panel, text="TEACHER", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=400)
        self.tf5 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=445, width=430, height=30)
        self.lb6 = tk.Label(self.panel, text="FOREIGN TEACHER", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=33, y=500)
        self.tf6 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf6.place(x=33, y=545, width=430, height=30)
        self.btn1 = tk.Button(self.panel, text="ADD NEW", font=("cambria", 14, "bold"),command=self.Add_book, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=150, y=600)

    def Add_book(self):
        worksheet3 = sht.worksheet("sheet 3")
        test = worksheet3.get_all_values()
        end_col = len([row[1] for row in test] )
        x= end_col-2+1
        name = self.tf1.get()
        day = self.tf2.get()
        time = self.tf3.get()
        room = int(self.tf4.get())
        teacher = self.tf5.get()
        fteacher = self.tf6.get()
        existing_names = [row[1] for row in test]
        if name in existing_names:
            messagebox.showerror("Error", "Lớp này đã được lưu")
        elif name == "":
            messagebox.showerror("Error", "Bạn chưa nhập tên lớp")
        else:
            new_row_values = [x,name,day,time,room,teacher,fteacher]
            worksheet3.append_row(new_row_values, value_input_option='RAW')
            messagebox.showinfo("Success", "Lưu thành công!")

        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewClass()
    app.run()
