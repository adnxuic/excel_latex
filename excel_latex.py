import webview
import json
import openpyxl
import pyperclip
from convert_to_latex import LatexConvert
import os
import io
import base64

class ExcelToLatexConverter:
    def __init__(self):
        self.data = None

    def import_excel(self, file_data):
        try:
            # 解码 base64 编码的文件数据
            decoded_data = base64.b64decode(file_data['data'])
            # 将解码后的数据转换为 BytesIO 对象
            excel_data = io.BytesIO(decoded_data)
            wb = openpyxl.load_workbook(excel_data)
            sheet = wb.active
            self.data = [[cell.value for cell in row] for row in sheet.iter_rows()]
            return json.dumps(self.data)
        except Exception as e:
            print(f"Error importing Excel: {str(e)}")  # 打印错误信息
            return json.dumps({"error": str(e)})

    def transpose_data(self):
        if self.data is None:
            return json.dumps({"error": "Please import data first."})
        
        self.data = [list(row) for row in zip(*self.data)]
        return json.dumps(self.data)

    def convert_to_latex(self):
        if self.data is None:
            return json.dumps({"error": "Please import data first."})
        
        latex_code = LatexConvert.convert_to_latex_code(self.data)
        return json.dumps({"latex": latex_code})

    def copy_to_clipboard(self, text):
        pyperclip.copy(text)
        return json.dumps({"success": True})

def main():
    converter = ExcelToLatexConverter()
    
    # Get the directory of the current script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Construct the path to the HTML file
    html_path = os.path.join(current_dir, 'index.html')

    # 创建一个更大的窗口
    webview.create_window("Excel to LaTeX Converter", url=html_path, js_api=converter, width=1000, height=800, min_size=(1000, 800))
    webview.start()

if __name__ == "__main__":
    main()
