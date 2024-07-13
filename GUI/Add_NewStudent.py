import tkinter as tk
from tkinter import messagebox
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1RPL8Tv_JctB7icajUTBoEq1lMO8XYb3sxySdGHJGgvY")
class Add_NewStudent:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Add new student manager")
        self.root.geometry("1300x680")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=1275, height=650)
        self.lbl_addNewBook = tk.Label(self.panel, text="Add new student manager", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=450, y=10)
        self.lb1 = tk.Label(self.panel, text="Full name", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel, text="Birthday (DOB)", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=370, y=108, width=300, height=30)


        self.lb3 = tk.Label(self.panel, text="Address", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=300, height=30)

        self.lb4 = tk.Label(self.panel, text="Starting off month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=370, y=160)
        self.tf4 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf4.place(x=370, y=213, width=300, height=30)

        self.lb5 = tk.Label(self.panel, text="Public school", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=300, height=30)

        self.lb6 = tk.Label(self.panel, text="Starting transfer month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=370, y=270)
        self.tf6 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf6.place(x=370, y=320, width=300, height=30)

        self.lb7 = tk.Label(self.panel, text="Vocab book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=300, height=30)

        self.lb8 = tk.Label(self.panel, text="New Comer", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=370, y=375)
        self.tf8 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf8.place(x=370, y=420, width=300, height=30)

        self.lb9 = tk.Label(self.panel, text="Tel", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=130, height=30)

        self.lb10 = tk.Label(self.panel, text="Enrolcamp", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb10.place(x=200, y=480)
        self.tf10 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf10.place(x=200, y=530, width=130, height=30)

        self.lb11 = tk.Label(self.panel, text="Main camp", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb11.place(x=370, y=480)
        self.tf11 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf11.place(x=370, y=530, width=130, height=30)

        self.lb12 = tk.Label(self.panel, text="Total fee", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb12.place(x=540, y=480)
        self.tf12 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf12.place(x=540, y=530, width=130, height=30)



        self.lb13 = tk.Label(self.panel, text="Main class", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb13.place(x=700, y=60)
        self.tf13 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf13.place(x=700, y=108, width=240, height=30)

        self.lb14 = tk.Label(self.panel, text="Parent name", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb14.place(x=1000, y=60)
        self.tf14 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf14.place(x=1000, y=108, width=240, height=30)


        self.lb15 = tk.Label(self.panel, text="Starting quit month", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb15.place(x=700, y=160)
        self.tf15 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf15.place(x=700, y=213, width=240, height=30)

        self.lb16 = tk.Label(self.panel, text="Teacher", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb16.place(x=700, y=270)
        self.tf16 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf16.place(x=700, y=320, width=240, height=30)

        self.lb17 = tk.Label(self.panel, text="Sub tel", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb17.place(x=700, y=375)
        self.tf17 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf17.place(x=700, y=420, width=240, height=30)


        self.lb18 = tk.Label(self.panel, text="Main fee", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb18.place(x=700, y=480)
        self.tf18 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf18.place(x=700, y=530, width=130, height=30)


        self.lb19 = tk.Label(self.panel, text="Certificate", font=("cambria", 16, "bold"), fg="#FBA834")
        self.lb19.place(x=850, y=480)
        self.tf19 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf19.place(x=850, y=530, width=130, height=30)


        self.lb20 = tk.Label(self.panel, text="Reading & Writing", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb20.place(x=1000, y=160)
        self.tf20 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf20.place(x=1050, y=213, width=100, height=30)

        self.lb21 = tk.Label(self.panel, text="Listening", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb21.place(x=1050, y=270)
        self.tf21 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf21.place(x=1050, y=320, width=100, height=30)

        self.lb22 = tk.Label(self.panel, text="Speaking", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb22.place(x=1050, y=375)
        self.tf22 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf22.place(x=1050, y=420, width=100, height=30)

        

        self.btn1 = tk.Button(self.panel, text="ADD NEW", font=("cambria", 14, "bold"), width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=520, y=600)

    # def Add_Class(self):
    #     worksheet3 = sht.worksheet("sheet 3")
    #     test = worksheet3.get_all_values()
    #     end_col = len([row[1] for row in test] )
    #     x= end_col-2+1
    #     name = self.tf1.get()
    #     day = self.tf2.get()
    #     time = self.tf3.get()
    #     room = int(self.tf4.get())
    #     teacher = self.tf5.get()
    #     fteacher = self.tf6.get()
    #     existing_names = [row[1] for row in test]
    #     if name in existing_names:
    #         messagebox.showerror("Error", "Lớp này đã được lưu")
    #     elif name == "":
    #         messagebox.showerror("Error", "Bạn chưa nhập tên lớp")
    #     else:
    #         new_row_values = [x,name,day,time,room,teacher,fteacher]
    #         worksheet3.append_row(new_row_values, value_input_option='RAW')
    #         messagebox.showinfo("Success", "Lưu thành công!")

        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewStudent()
    app.run()
