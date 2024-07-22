def create_score_management_tab(self):
        self.score_management_tab = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.score_management_tab, text="Quản lý điểm số")
        button_frame = ttk.Frame(self.score_management_tab, style='TFrame')
        button_frame.pack(side="top", fill="x")
        btnXuatExcel2 = ttk.Button(button_frame, text="Xuất excel",command=self.XuatExcel12, width=25, style='TButton')
        btnXuatExcel2.pack(side="right", padx=5, pady=5)
        
        table_columns2 = ["ID", "FULL NAME", "MAIN CLASS", "TEACHER", "LISTENING", "SPEAKING", "WRITING & READING", "TOTAL GRADE", "PERCENT"]
        self.table2 = ttk.Treeview(self.score_management_tab, columns=table_columns2, show="headings", height=25)
        for col in table_columns2:
            self.table2.heading(col, text=col)
        for row in combined_data6:
            self.table2.insert("", "end", values=row)
        self.table2.pack(fill="x")
        self.table2.bind("<Double-1>", self.on_row_select2)

        tree_scrollx2 = ttk.Scrollbar(self.score_management_tab, orient="horizontal", command=self.table2.xview)
        tree_scrollx2.pack(fill="x")
        self.table2.configure(xscrollcommand=tree_scrollx2.set)
        
        self.create_search_section(self.score_management_tab, "score")