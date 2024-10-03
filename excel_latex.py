import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

import openpyxl
import pyperclip

from convert_to_latex import LatexConvert

class ExcelToLatexConverter:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.drop_target_register(DND_FILES)
        self.root.title("Excel to LaTeX Converter")
        self.root.geometry("1000x600")

        self.table = None
        self.latex_text = None
        self.data = None  # 添加这行来存储当前数据

        self.create_ui()
        
        # 添加拖放支持
        self.root.dnd_bind('<<Drop>>', self.drop_import)

    def create_ui(self):
        # Create menu
        menu = tk.Menu(self.root)
        self.root.config(menu=menu)

        file_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Excel", command=self.import_excel)
        file_menu.add_command(label="Transpose Data", command=self.transpose_data)  # 添加新的菜单项

        # Create a horizontal PanedWindow
        paned_window = tk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # Left frame for table
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame)
        paned_window.paneconfig(left_frame, stretch="always", width=200)

        # Create table to display Excel data
        self.table = ttk.Treeview(left_frame)
        self.table.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.table["columns"] = []
        self.table["show"] = "headings"

        # Add scrollbars to the table
        vsb = ttk.Scrollbar(left_frame, orient="vertical", command=self.table.yview)
        hsb = ttk.Scrollbar(left_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        # Right frame for LaTeX output
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame)

        # Create text widget to display LaTeX code
        self.latex_text = tk.Text(right_frame, height=10)
        self.latex_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Create copy button
        copy_button = tk.Button(right_frame, text="Copy LaTeX", command=self.copy_latex)
        copy_button.pack(pady=10)

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.process_excel_file(file_path)

    def process_excel_file(self, file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            self.data = [[cell.value for cell in row] for row in sheet.iter_rows()]
            self.display_data(self.data)
            latex_code = LatexConvert.convert_to_latex_code(self.data)
            self.display_latex(latex_code)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel file: {str(e)}")

    def transpose_data(self):
        if self.data is None:
            messagebox.showwarning("Warning", "Please import data first.")
            return
        
        # 转置数据
        transposed_data = [list(row) for row in zip(*self.data)]
        self.data = transposed_data
        
        # 更新显示
        self.display_data(self.data)
        latex_code = LatexConvert.convert_to_latex_code(self.data)
        self.display_latex(latex_code)

    def display_data(self, data):
        self.table.delete(*self.table.get_children())

        # Clear existing columns
        for col in self.table["columns"]:
            self.table.heading(col, text="")

        # Set up new columns based on the data
        num_columns = len(data[0])
        self.table["columns"] = [f"col{i}" for i in range(num_columns)]

        for i, heading in enumerate(data[0]):
            self.table.heading(f"col{i}", text=heading)
            self.table.column(f"col{i}", width=100)  # Set a default width

        # Insert data rows
        for row in data[1:]:
            self.table.insert("", "end", values=row)

    def display_latex(self, latex_code):
        self.latex_text.delete(1.0, tk.END)
        self.latex_text.insert(tk.END, latex_code)

    def copy_latex(self):
        latex_code = self.latex_text.get(1.0, tk.END)
        pyperclip.copy(latex_code)
        messagebox.showinfo("Success", "LaTeX code copied to clipboard!")

    def drop_import(self, event):
        file_path = event.data
        if file_path.endswith(('.xlsx', '.xls')):
            self.process_excel_file(file_path)
        else:
            messagebox.showerror("Error", "Please drop a valid Excel file (.xlsx or .xls)")

    def run(self):
        self.root.mainloop()


def main():
    app = ExcelToLatexConverter()
    app.run()


if __name__ == "__main__":
    main()
