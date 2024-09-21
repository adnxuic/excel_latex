import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import openpyxl
import pyperclip

def import_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            data = [[cell.value for cell in row] for row in sheet.iter_rows()]
            display_data(data)
            latex_code = convert_to_latex(data)
            display_latex(latex_code)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel file: {str(e)}")

def display_data(data):
    table.delete(*table.get_children())
    for row in data:
        table.insert("", "end", values=row)

def convert_to_latex(data):
    num_columns = len(data[0])
    num_rows = len(data)
    
    latex = "\\begin{table}[htbp]\n"
    latex += "\t\\centering\n"
    latex += "\t\t\\caption{表名}\n"
    latex += "\t\t\\vspace{1em}\n"
    latex += f"\t\t\\resizebox{{{0.1 * num_columns}\\textwidth}}{{{2.5 * num_rows:.1f}mm}}\n"
    latex += "\t\t{\n"
    latex += f"\t\t\\begin{{tabular}}{{{'c' * num_columns}}}\n"

    latex += "\t\t\\toprule\n"
    
    for i, row in enumerate(data):
        latex += "\t\t" + " & ".join(str(cell) for cell in row) + " \\\\\n"
        if i == 0:
            latex += "\t\t\\midrule\n"
    
    latex += "\t\t\\bottomrule\n"
    latex += "\t\t\\end{tabular}\n"
    latex += "\t\t}\n"
    latex += "\t\t\\label{tab:data}\n"
    latex += "\\end{table}"
    return latex

def display_latex(latex_code):
    latex_text.delete(1.0, tk.END)
    latex_text.insert(tk.END, latex_code)

def copy_latex():
    latex_code = latex_text.get(1.0, tk.END)
    pyperclip.copy(latex_code)
    messagebox.showinfo("Success", "LaTeX code copied to clipboard!")

# Create main window
root = tk.Tk()
root.title("Excel to LaTeX Converter")
root.geometry("1000x600")

# Create menu
menu = tk.Menu(root)
root.config(menu=menu)

file_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Import Excel", command=import_excel)

# Create a horizontal PanedWindow
paned_window = tk.PanedWindow(root, orient=tk.HORIZONTAL)
paned_window.pack(fill=tk.BOTH, expand=True)

# Left frame for table
left_frame = ttk.Frame(paned_window)
paned_window.add(left_frame)

# Set weight of the left frame
paned_window.paneconfig(left_frame, stretch="always", width=200)  # Set initial width to 200 pixels

# Create table to display Excel data
table = ttk.Treeview(left_frame)
table.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Add scrollbars to the table
vsb = ttk.Scrollbar(left_frame, orient="vertical", command=table.yview)
hsb = ttk.Scrollbar(left_frame, orient="horizontal", command=table.xview)
table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
vsb.pack(side="right", fill="y")
hsb.pack(side="bottom", fill="x")

# Right frame for LaTeX output
right_frame = ttk.Frame(paned_window)
paned_window.add(right_frame)

# Create text widget to display LaTeX code
latex_text = tk.Text(right_frame, height=10)
latex_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Create copy button
copy_button = tk.Button(right_frame, text="Copy LaTeX", command=copy_latex)
copy_button.pack(pady=10)

root.mainloop()
