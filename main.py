import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk

class ExcelCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 文件比較工具")
        self.root.geometry("800x800")


        self.frame_source = tk.Frame(self.root,borderwidth=2, relief="ridge")

        self.file_label_source = tk.Label(self.frame_source, text="選擇來源文件")
        self.file_label_source.pack(anchor="w")
        self.file_button_source = tk.Button(self.frame_source, text="瀏覽", command=self.open_source_fils)
        self.file_button_source.pack(anchor="w")
        self.column_label_source = tk.Label(self.frame_source, text="選擇來源文件要比較的欄位")
        self.column_label_source.pack(anchor="w")

        self.column_selection1 = ttk.Combobox(self.frame_source)
        self.column_selection1.pack(anchor="w")
        self.frame_source.pack(fill="x", padx=10, pady=10)


        self.frame_filter = tk.Frame(self.root,borderwidth=2, relief="ridge")
        self.frame_filter.pack(fill="x", padx=10, pady=10)

        self.file_label = tk.Label(self.frame_filter, text="選擇用來過濾的文件")
        self.file_label.pack(anchor="w")
        self.file_button = tk.Button(self.frame_filter, text="瀏覽", command=self.open_files)
        self.file_button.pack(anchor="w")


        self.column_label = tk.Label(self.frame_filter, text="請選擇要比較的欄位")
        self.column_label.pack(anchor="w")

        self.column_selections = []

        self.result_label = tk.Label(self.frame_filter, text="")
        self.result_label.pack(anchor="w")

        self.compare_button = tk.Button(self.root, text="排除重複值", command=self.exclude_data)
        self.compare_button.pack(anchor="w")

        self.compare_button = tk.Button(self.root, text="找尋重複值", command=self.include_data)
        self.compare_button.pack(anchor="w")

        self.compare_button=[]
        

        self.files_data = []
        self.file_paths = []

    def open_source_fils(self):
        file_types = [("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
        file_paths = filedialog.askopenfilenames(filetypes=file_types)
        if file_paths:
            self.source_file_path = file_paths[0]
            self.source_file_data = {}
            
            if self.source_file_path.endswith(".csv"):
                data = pd.read_csv(self.source_file_path)
            else:
                data = pd.read_excel(self.source_file_path)
            self.source_file_data=data
            self.populate_column_selection(self.column_selection1, data)

    def open_files(self):
        file_types = [("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
        file_paths = filedialog.askopenfilenames(filetypes=file_types)
        if file_paths:
            self.result_label.config(text="已選擇 {} 個文件".format(len(file_paths)))

            for file_path in file_paths:
                if file_path.endswith(".csv"):
                    data = pd.read_csv(file_path)
                else:
                    data = pd.read_excel(file_path)
                self.files_data.append(data)
                self.file_paths.append(file_path)

                paths = file_path.split("/")
                column_label = tk.Label(self.frame_filter, text=f"選擇 {paths[len(paths)-1]} 的比較欄位")
                column_label.pack(anchor="w")
                column_selection = ttk.Combobox(self.frame_filter)
                column_selection.pack(anchor="w")

                self.column_selections.append(column_selection)
                self.populate_column_selection(column_selection, data)

                # self.remove_button = tk.Button(self.frame_filter, text="移除", command=lambda path=file_path: self.remove_file(file_path))
                # self.remove_button.pack(side="left")
            
    def remove_file(self, file_path):
        index = self.file_paths.index(file_path)
        if index >= 0:
            self.files_data.pop(index)
            self.file_paths.pop(index)
            self.column_selections.pop(index).destroy()
            
    def populate_column_selection(self, combobox, data):
        if data is not None:
            columns = data.columns.tolist()
            combobox['values'] = columns
            if len(columns) > 0:
                combobox.set(columns[0])

    def exclude_data(self):
        if not self.files_data:
            self.result_label.config(text="請選擇要比較的文件")
            return

        selected_columns = [column_selection.get() for column_selection in self.column_selections]

        if not all(selected_columns):
            self.result_label.config(text="請為每個文件選擇要比較的欄位")
            return

        try:
            for i, data in enumerate(self.files_data):
                selected_column = selected_columns[i]
                source_column = self.column_selection1.get()
                self.source_file_data=self.source_file_data[~self.source_file_data[source_column].isin(data[selected_column])]

            self.save_file(self.source_file_data,"exclude_result")
            
        except Exception as e:
            self.result_label.config(text="錯誤：" + str(e))
    def include_data(self):
        if not self.files_data:
            self.result_label.config(text="請選擇要比較的文件")
            return

        selected_columns = [column_selection.get() for column_selection in self.column_selections]

        if not all(selected_columns):
            self.result_label.config(text="請為每個文件選擇要比較的欄位")
            return

        try:
            for i, data in enumerate(self.files_data):
                selected_column = selected_columns[i]
                source_column = self.column_selection1.get()
                self.source_file_data=self.source_file_data[self.source_file_data[source_column].isin(data[selected_column])]
            self.save_file(self.source_file_data,"include_result")

        except Exception as e:
            self.result_label.config(text="錯誤：" + str(e))

    def save_file(self, data,default_name):
        if self.source_file_path.endswith(".csv"):
            file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")], initialfile=default_name+".csv")
            data.to_csv(file_path, index=False)
        else:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("excel Files", "*.xlsx")], initialfile=default_name+".xlsx")
            data.to_excel(file_path, index=False)
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()
