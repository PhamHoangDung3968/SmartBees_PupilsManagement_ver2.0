import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

class Add_NewClass:
    def __init__(self):
        
        self.root = tk.Tk()
        self.root.title("Add new class manager")
        self.root.geometry("520x680")
        self.canvas = tk.Canvas(self.root, width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Frame(self.canvas, bd=4, relief="solid")
        self.panel.place(x=10, y=10, width=500, height=650)
        self.lbl_addNewClass = tk.Label(self.panel, text="Add new class manager", font=("cambria", 24, "bold"), fg="#FBA834")
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
        self.btn1 = tk.Button(self.panel, text="ADD NEW", font=("cambria", 14), width=20, bg="#FBA834",fg="white" )
        self.btn1.place(x=150, y=600)
    
    def run(self):
        self.root.resizable(False, False)
        self.root.mainloop()

if __name__ == "__main__":
    app = Add_NewClass()
    app.run()
