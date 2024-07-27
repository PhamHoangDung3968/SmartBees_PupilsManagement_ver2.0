import tkinter as tk
from tkinter import messagebox
import gspread
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
class Add_NewBook:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Add new book")
        self.root.geometry("1020x680")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=1000, height=650)
        self.lbl_addNewBook = tk.Label(self.panel, text="ADD NEW BOOK", font=("cambria", 24, "bold"), fg="black")
        self.lbl_addNewBook.place(x=380, y=10)
        self.lb1 = tk.Label(self.panel, text="Cambridge level", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=430, height=30)

        self.lb2 = tk.Label(self.panel, text="Main book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=530, y=60)
        self.tf2 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf2.place(x=530, y=108, width=430, height=30)


        self.lb3 = tk.Label(self.panel, text="Skill book 1", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb3.place(x=33, y=160)
        self.tf3 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf3.place(x=33, y=213, width=430, height=30)

        self.lb4 = tk.Label(self.panel, text="Skill book 2", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb4.place(x=530, y=160)
        self.tf4 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf4.place(x=530, y=213, width=430, height=30)

        self.lb5 = tk.Label(self.panel, text="Skill book 3", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb5.place(x=33, y=270)
        self.tf5 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf5.place(x=33, y=320, width=430, height=30)

        self.lb6 = tk.Label(self.panel, text="Skill book 4", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb6.place(x=530, y=270)
        self.tf6 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf6.place(x=530, y=320, width=430, height=30)

        self.lb7 = tk.Label(self.panel, text="Vocab book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb7.place(x=33, y=375)
        self.tf7 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf7.place(x=33, y=420, width=430, height=30)

        self.lb8 = tk.Label(self.panel, text="Grammar book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb8.place(x=530, y=375)
        self.tf8 = tk.Entry(self.panel,font=("cambria", 13, "bold"))
        self.tf8.place(x=530, y=420, width=430, height=30)

        self.lb9 = tk.Label(self.panel, text="Test book", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb9.place(x=33, y=480)
        self.tf9 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf9.place(x=33, y=530, width=200, height=30)

        self.lb10 = tk.Label(self.panel, text="Progress", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb10.place(x=260, y=480)
        self.tf10 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf10.place(x=260, y=530, width=200, height=30)

        self.lb11 = tk.Label(self.panel, text="Videos-Movies", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb11.place(x=530, y=480)
        self.tf11 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf11.place(x=530, y=530, width=200, height=30)

        self.lb12 = tk.Label(self.panel, text="Pictures-Cards", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb12.place(x=750, y=480)
        self.tf12 = tk.Entry(self.panel, font=("cambria", 13, "bold"))
        self.tf12.place(x=750, y=530, width=200, height=30)


        self.btn1 = tk.Button(self.panel, text="SUBMIT", font=("cambria", 14, "bold"),command=self.Add_Book, width=20, bg="#FBA834",fg="black" )
        self.btn1.place(x=400, y=600)

    def Add_Book(self):
        worksheet1 = sht.worksheet("sheet 1")
        test = worksheet1.get_all_values()
        values = worksheet1.col_values(1)[2:]
        max_value = max(list(map(int, values))) + 1  
        cl = self.tf1.get()
        mb = self.tf2.get()
        sk1 = self.tf3.get()
        sk2 = self.tf4.get()
        sk3 = self.tf5.get()
        sk4 = self.tf6.get()
        vb = self.tf7.get()
        gb = self.tf8.get()
        tb = self.tf9.get()
        p = self.tf10.get()
        vm = self.tf11.get()
        pc = self.tf12.get()
        existing_cls = [row[1] for row in test]
        if cl in existing_cls:
            messagebox.showerror("Error", "Cấp độ này đã được lưu")
            self.tf1.delete(0, 'end')
        elif cl == "":
            messagebox.showerror("Error", "Bạn chưa nhập cấp độ")
        else:
            new_row_values = [max_value,cl,mb,sk1,sk2,sk3,sk4,vb,gb,tb,p,vm,pc]
            worksheet1.append_row(new_row_values, value_input_option='RAW')
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

        
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewBook()
    app.run()
