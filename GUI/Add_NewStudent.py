import tkinter as tk
from tkinter import messagebox
import gspread
from tkinter import ttk
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
worksheet3 = sht.worksheet("sheet 3")
values_list_Class = worksheet3.get_all_values()[2:]
result_list_Class = [row[1] for row in values_list_Class]
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
        
        self.lb7 = tk.Label(self.panel, text="Parent name", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=430, height=30)

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
        self.tf13 = ttk.Combobox(self.panel, font=("cambria", 13, "bold"))
        self.tf13['values'] = result_list_Class
        self.tf13.current(0)
        self.tf13.place(x=700, y=108, width=240, height=30)

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

        

        self.btn1 = tk.Button(self.panel, text="ADD NEW", font=("cambria", 14, "bold"),command=self.Add_Student, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=520, y=600)

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
        rw = float(self.tf20.get())
        lis = float(self.tf21.get())
        spe = float(self.tf22.get())
        total = rw + lis + spe
        percent = str(round((total/15)*100,2))+"%"
        
        existing_fn = [row[:3] for row in test]
        if fn and b in existing_fn:
            messagebox.showerror("Error", "Học sinh này đã được lưu")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')

        elif fn == "" or b =="":
            messagebox.showerror("Error", "Bạn chưa nhập cấp độ")
        else:
            new_row_values = [max_value ,fn, b, mcla, t, a, stc, e, mc, tf,mf, nc, som,sqm,c, ps,st, stm, tea,lis,spe, rw,total, percent]
            worksheet2.append_row(new_row_values, value_input_option='RAW')
            messagebox.showinfo("Success", "Lưu thành công!")
            self.tf1.delete(0, 'end')
            self.tf2.delete(0, 'end')
            self.tf3.delete(0, 'end')
            self.tf4.delete(0, 'end')
            self.tf5.delete(0, 'end')
            self.tf6.delete(0, 'end')
            self.tf8.delete(0, 'end')
            self.tf9.delete(0, 'end')
            self.tf10.delete(0, 'end')
            self.tf11.delete(0, 'end')
            self.tf12.delete(0, 'end')

            self.tf13.delete(0, 'end')
            self.tf14.delete(0, 'end')
            self.tf15.delete(0, 'end')
            self.tf16.delete(0, 'end')
            self.tf17.delete(0, 'end')
            self.tf18.delete(0, 'end')
            self.tf19.delete(0, 'end')
            self.tf20.delete(0, 'end')
            self.tf21.delete(0, 'end')
            self.tf22.delete(0, 'end')

        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewStudent()
    app.run()
