import sys
import os
import time
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QFileDialog,
    QLabel,
    QHBoxLayout,
    QProgressBar,
    QMessageBox,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor
import subprocess


# --- Helper Functions ---
def check_valid_filename(file_path):
    category_prefix_map = {
        "BO-API": "共通書店システムのオンプレミス化対応_単体テスト報告書_BO",
        "BO-WEB": "共通書店システムのオンプレミス化対応_単体テスト報告書_BO",
        "QUEUE-API": "共通書店システムのオンプレミス化対応_単体テスト報告書_BO",
        "EXT-API": "共通書店システムのオンプレミス化対応_単体テスト報告書_EXT",
        "DB-Functions": "共通書店システムのオンプレミス化対応_単体テスト報告書_DB",
        "DB-Packages": "共通書店システムのオンプレミス化対応_単体テスト報告書_DB",
        "DB-Sequences": "共通書店システムのオンプレミス化対応_単体テスト報告書_DB",
        "DB-Tables": "共通書店システムのオンプレミス化対応_単体テスト報告書_DB",
    }

    filename = os.path.basename(file_path)
    parts = os.path.normpath(file_path).split(os.sep)

    for folder_name, expected_prefix in category_prefix_map.items():
        if folder_name in parts:
            if not filename.startswith(expected_prefix):
                return f"Invalid filename for '{folder_name}'"
            break  # Stop checking after first match

    return None


def check_invalid_sheet(wb, invalid_sheets={"HOW TO TEST"}):
    for sheet in invalid_sheets:
        if sheet in wb.sheetnames:
            return f"Contains invalid sheet: {sheet}"
    return None


def check_required_sheets(wb, required_sheets={"表紙", "テスト項目"}):
    for sheet in required_sheets:
        if sheet not in wb.sheetnames:
            return f"Missing required sheet: {sheet}"
    return None


def check_confirm_by(wb):
    if "表紙" not in wb.sheetnames:
        return None
    ws = wb["表紙"]
    p24_value = ws["P24"].value
    if p24_value is None or str(p24_value).strip() == "":
        return "Missing Confirm"
    return None


def find_column_indexes(ws, headers=("確認", "参考"), header_row=3):
    col_indexes = {}
    for cell in ws[header_row]:
        if cell.value in headers:
            col_indexes[cell.value] = cell.column
    return col_indexes


def check_status_in_test_items(wb, max_rows=1000, empty_limit=10):
    if "テスト項目" not in wb.sheetnames:
        return None
    ws = wb["テスト項目"]
    col_indexes = find_column_indexes(ws)
    status_col = col_indexes.get("確認")

    if not status_col:
        return "Column '確認' not found"

    error_rows = []
    consecutive_empty = 0

    for row in range(5, max_rows + 1):
        if consecutive_empty >= empty_limit:
            break
        b_cell = ws.cell(row=row, column=2)
        b_value = b_cell.value
        if b_value and str(b_value).strip():
            consecutive_empty = 0
            status_cell = ws.cell(row=row, column=status_col)
            status_value = status_cell.value
            if status_value is None or str(status_value).strip().upper() != "OK":
                error_rows.append(f"{str(b_value).strip()}")
        else:
            consecutive_empty += 1

    if error_rows:
        return f"{len(error_rows)} TC(s) != 'OK': " + "; ".join(error_rows)
    return None


# --- Main Function ---
def check_excel_file_advanced(file_path):
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        error_messages = []

        # Run each check
        err = check_valid_filename(file_path)
        if err:
            error_messages.append(err)

        err = check_required_sheets(wb)
        if err:
            error_messages.append(err)

        err = check_confirm_by(wb)
        if err:
            error_messages.append(err)

        err = check_status_in_test_items(wb)
        if err:
            error_messages.append(err)

        err = check_invalid_sheet(wb)
        if err:
            error_messages.append(err)

        wb.close()

        # Return aggregated results
        if error_messages:
            return "ERROR", ", ".join(error_messages)
        return "OK", ""

    except Exception as e:
        return "ERROR", str(e)


def find_excel_files_recursive(folder_path):
    excel_files = []
    excel_extensions = (".xlsx", ".xlsm", ".xls")
    for root_dir, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(excel_extensions):
                excel_files.append(os.path.join(root_dir, file))
    return excel_files


# ----------------- Worker Thread -----------------
class ExcelCheckWorker(QThread):
    progress_changed = pyqtSignal(int)
    file_result = pyqtSignal(
        str, str, str, str
    )  # prefix_path, relative_path, status, error
    finished_signal = pyqtSignal()

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        files = find_excel_files_recursive(self.folder_path)
        total = len(files)
        if not files:
            self.file_result.emit(self.folder_path, "", "INFO", "No Excel files found.")
            self.finished_signal.emit()
            return

        for i, file_path in enumerate(files, 1):
            relative_path = os.path.relpath(file_path, self.folder_path)
            status, error_msg = check_excel_file_advanced(file_path)
            self.file_result.emit(self.folder_path, relative_path, status, error_msg)
            self.progress_changed.emit(int((i / total) * 100))
            time.sleep(0.05)

        self.finished_signal.emit()


# ----------------- PyQt Main Window -----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Checker")
        self.setGeometry(100, 100, 800, 500)

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

        # Table widget
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            ["Prefix Path", "Relative Path", "Status", "Errors"]
        )
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.itemDoubleClicked.connect(self.open_selected_file)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setValue(0)

        # Layout add
        main_layout.addLayout(input_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.table)
        main_layout.addWidget(self.progress_bar)
        self.setLayout(main_layout)

        self.worker = None

    def on_folder_input_change(self, text):
        self.btn_execute.setEnabled(os.path.isdir(text.strip()))

    def select_folder(self):
        current_path = self.folder_input.text().strip()

        if os.path.isdir(current_path):
            start_dir = current_path
        else:
            start_dir = os.path.expanduser("~")

        folder = QFileDialog.getExistingDirectory(self, "Select Folder", start_dir)
        if folder:
            self.folder_input.setText(folder)

    def start_execution(self):
        folder_path = self.folder_input.text().strip()
        if not os.path.isdir(folder_path):
            QMessageBox.warning(
                self, "Invalid Folder", "Please provide a valid folder path."
            )
            return

        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        self.btn_execute.setEnabled(False)

        self.worker = ExcelCheckWorker(folder_path)
        self.worker.progress_changed.connect(self.progress_bar.setValue)
        self.worker.file_result.connect(self.add_table_row)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def add_table_row(self, prefix_path, path, status, error):
        self.table.setSortingEnabled(False)
        row = self.table.rowCount()
        self.table.insertRow(row)
        correct_path = path.replace("\\", "/")
        prefix_item = QTableWidgetItem(prefix_path)
        status_item = QTableWidgetItem(status)
        path_item = QTableWidgetItem(correct_path)
        error_item = QTableWidgetItem(error)
        if status == "OK":
            status_item.setForeground(QColor("green"))
        elif status == "ERROR":
            status_item.setForeground(QColor("red"))
        self.table.setItem(row, 0, prefix_item)
        self.table.setItem(row, 1, path_item)
        self.table.setItem(row, 2, status_item)
        self.table.setItem(row, 3, error_item)
        self.table.setSortingEnabled(True)

    def open_selected_file(self, item):
        row = item.row()
        path = os.path.join(self.folder_input.text(), self.table.item(row, 1).text())
        if os.path.exists(path):
            if os.name == "nt":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        else:
            QMessageBox.warning(self, "File Not Found", f"File not found: {path}")

    def on_finished(self):
        self.btn_execute.setEnabled(True)
        QMessageBox.information(self, "Done", "Check completed.")


# ----------------- Entry Point -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
