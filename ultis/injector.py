import os
import win32com.client
import pythoncom
from pathlib import Path

class VBAInjector:
    def __init__(self):
        self.input_file = None

    def inject_vba_to_office(self, input_path, output_path, vba_code):
        pythoncom.CoInitialize()
        self.input_file = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)

        # Đảm bảo folder output tồn tại
        output_folder = os.path.dirname(output_path)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        suffix = Path(self.input_file).suffix.lower()

        if suffix == '.docx':
            self.embed_word_vba(vba_code, output_path)
        elif suffix == '.xlsx':
            self.embed_excel_vba(vba_code, output_path)
        else:
            raise ValueError(f"Không hỗ trợ file: {self.input_file}")

        pythoncom.CoUninitialize()

    def embed_word_vba(self, vba_code, output_file):
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(self.input_file)
            this_document = doc.VBProject.VBComponents("ThisDocument")
            this_document.CodeModule.AddFromString(vba_code)
            doc.SaveAs(output_file, FileFormat=13)  # 13 = wdFormatXMLDocumentMacroEnabled
            doc.Close()
        finally:
            word.Quit()

    def embed_excel_vba(self, vba_code, output_file):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(self.input_file)
            this_workbook = wb.VBProject.VBComponents("ThisWorkbook")
            this_workbook.CodeModule.AddFromString(vba_code)
            wb.SaveAs(output_file, FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
            wb.Close()
        finally:
            excel.Quit()
