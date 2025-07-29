import sys
import os
import time
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit, QFileDialog,
    QLabel, QHBoxLayout, QProgressBar, QMessageBox, QLineEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# ----------------- Excel File Validation Logic -----------------
def check_excel_file_advanced(file_path):
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        filename = os.path.basename(file_path)
        results = []

        required_sheets = {'表紙', 'テスト項目'}
        missing_sheets = required_sheets - set(wb.sheetnames)
        if missing_sheets:
            return f"\n[ERROR] {filename}:\nMissing sheet(s): {', '.join(missing_sheets)}"

        ws = wb['表紙']
        p24_value = ws['P24'].value
        if p24_value is None or str(p24_value).strip() == "":
            wb.close()
            return f"\n[ERROR] {filename}:\nMissing \"Confirm by\"."
        results.append(f" Confirm by: {str(p24_value).strip()}")

        ws2 = wb['テスト項目']
        error_rows = []
        checked_rows = []
        status_addr = ""
        status_col = 0

        for cell in ws2[3]:
            if cell.value == "確認":
                status_addr = cell.coordinate[:2]
                status_col = cell.column
                break

        max_rows_to_check = 1000
        consecutive_empty_limit = 10
        consecutive_empty = 0

        for row in range(5, max_rows_to_check + 1):
            if consecutive_empty >= consecutive_empty_limit:
                break
            b_cell = ws2.cell(row=row, column=2)
            b_value = b_cell.value
            if b_value is not None and str(b_value).strip():
                consecutive_empty = 0
                checked_rows.append(row)
                status_cell = ws2.cell(row=row, column=status_col)
                status_value = status_cell.value
                if status_value is None or str(status_value).strip().upper() != "OK":
                    error_rows.append(
                        f"Row {row} (B{row}='{str(b_value).strip()}', "
                        f"{status_addr}{row}='{str(status_value or '').strip()}')"
                    )
            else:
                consecutive_empty += 1

        if not checked_rows:
            results.append(" Did not find any Test case in sheet テスト項目")
        else:
            results.append(f" Checked {len(checked_rows)} Test case(s):")
            if error_rows:
                results.append(f" {len(error_rows)} rows failed validation:")
                for error_row in error_rows[:5]:
                    results.append(f"  - {error_row}")
                if len(error_rows) > 5:
                    results.append(f"  ... and {len(error_rows) - 5} more errors")
                wb.close()
                return f"\n[ERROR] {filename}:\n" + "\n".join(results)
            else:
                results.append(" All 確認 rows contain value 'OK'")

        wb.close()
        return f"\n[SUCCESS] {filename}:\n" + "\n".join(results)
    except Exception as e:
        return f"\n[ERROR] {os.path.basename(file_path)}: {str(e)}"

def find_excel_files_recursive(folder_path):
    excel_files = []
    excel_extensions = ('.xlsx', '.xlsm', '.xls')
    for root_dir, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(excel_extensions):
                excel_files.append(os.path.join(root_dir, file))
    return excel_files

# ----------------- Worker Thread -----------------
class ExcelCheckWorker(QThread):
    progress_changed = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        files = find_excel_files_recursive(self.folder_path)
        total = len(files)
        if not files:
            self.log_message.emit("No Excel files found in the selected folder.")
            self.finished_signal.emit()
            return

        self.log_message.emit(f"Found {total} Excel files. Starting validation...\n")
        success_count, error_count = 0, 0

        for i, file_path in enumerate(files, 1):
            # Get relative path for logs
            relative_path = os.path.relpath(file_path, self.folder_path)

            result = check_excel_file_advanced(file_path)

            # Replace file name with relative path in result
            filename = os.path.basename(file_path)
            result = result.replace(f"{filename}:", f"{relative_path}:")

            if result.startswith("\n[SUCCESS]"):
                success_count += 1
            else:
                error_count += 1

            self.log_message.emit(result)
            self.progress_changed.emit(int((i / total) * 100))
            time.sleep(0.05)

        self.log_message.emit(f"\nCompleted: {total} files ({success_count} passed, {error_count} failed)")
        self.finished_signal.emit()


# ----------------- PyQt Main Window -----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Validator")
        self.setGeometry(100, 100, 600, 400)

        # Layouts
        main_layout = QVBoxLayout()
        input_layout = QHBoxLayout()
        button_layout = QHBoxLayout()

        # Folder input bar
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("Paste or type folder path here...")
        self.folder_input.textChanged.connect(self.on_folder_input_change)

        # Select button
        self.btn_select = QPushButton("Browse")
        self.btn_select.clicked.connect(self.select_folder)

        input_layout.addWidget(self.folder_input)
        input_layout.addWidget(self.btn_select)

        # Execute button
        self.btn_execute = QPushButton("Execute")
        self.btn_execute.setEnabled(False)
        self.btn_execute.clicked.connect(self.start_execution)

        button_layout.addWidget(self.btn_execute)

        # Log box and progress bar
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setValue(0)

        # Add widgets
        main_layout.addLayout(input_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.log_box)
        main_layout.addWidget(self.progress_bar)
        self.setLayout(main_layout)

        self.worker = None

    def on_folder_input_change(self, text):
        if os.path.isdir(text.strip()):
            self.btn_execute.setEnabled(True)
        else:
            self.btn_execute.setEnabled(False)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_input.setText(folder)

    def start_execution(self):
        folder_path = self.folder_input.text().strip()
        if not os.path.isdir(folder_path):
            QMessageBox.warning(self, "Invalid Folder", "Please provide a valid folder path.")
            return

        self.progress_bar.setValue(0)
        self.log_box.clear()
        self.btn_execute.setEnabled(False)

        self.worker = ExcelCheckWorker(folder_path)
        self.worker.progress_changed.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.log_box.append)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self):
        self.btn_execute.setEnabled(True)
        self.log_box.append("\nValidation completed.\n")

# ----------------- Entry Point -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
