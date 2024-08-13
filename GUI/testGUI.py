import tkinter as tk
from tkinter import messagebox
import gspread
from tkinter import ttk
gs = gspread.service_account("cre.json")
sht = gs.open_by_key("1tTAZapKjFJ21FYJGoEZBaIYRmHWv2LmW_G4lwZ2pOUE")
worksheet3 = sht.worksheet("sheet 3")
values_list_Class = worksheet3.get_all_values()[2:]
result_list_Class = [row[1] for row in values_list_Class]
class testGUI:
    def __init__(self):
        self.rootScore = tk.Tk()
        self.rootScore.title("Edit Score")
        self.rootScore.geometry("1000x700")
        self.canvas4 = tk.Canvas(self.rootScore, width=self.rootScore.winfo_screenwidth(), height=self.rootScore.winfo_screenheight())
        self.canvas4.pack(fill=tk.BOTH, expand=True)
        self.panel4 = tk.Frame(self.canvas4, bd=4, relief="solid")
        self.panel4.place(x=10, y=10, width=980, height=670)
        self.lbl_editScore = tk.Label(self.panel4, text="Edit Score", font=("cambria", 24, "bold"), fg="black")
        self.lbl_editScore.place(x=450, y=10)
        self.tfname = tk.Entry(self.panel4,font=("cambria", 13, "bold"), state='readonly')
        self.tfname.place(x=33, y=10, width=300, height=30)

        self.lb1 = tk.Label(self.panel4, text="Exam invigilator", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb1.place(x=33, y=60)
        self.tf1 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
        self.tf1.place(x=33, y=108, width=300, height=30)

        self.lb2 = tk.Label(self.panel4, text="Exam day", font=("cambria", 18, "bold"), fg="#FBA834")
        self.lb2.place(x=370, y=60)
        self.tf2 = tk.Entry(self.panel4,font=("cambria", 13, "bold"))
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


        

        self.btn1 = tk.Button(self.panel4, text="SUBMIT", font=("cambria", 14, "bold"), width=12, bg="#FBA834",fg="black" )
        self.btn1.place(x=800, y=330)
        
    def run(self):
        self.rootScore.resizable(False, False)
        self.rootScore.mainloop()

if __name__ == "__main__":
    app = testGUI()
    app.run()
