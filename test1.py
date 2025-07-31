import sys
import os
import time
import pandas as pd
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
    QCheckBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
from datetime import datetime


def load_config(config_path="config.json"):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()

# --- Constants ---
CATEGORY_PREFIX_MAP = CONFIG["category_prefix_map"]
INVALID_SHEETS = CONFIG["invalid_sheets"]
REQUIRED_SHEETS = CONFIG["required_sheets"]
EXCEL_EXTENSIONS = tuple(CONFIG["excel_extensions"])
INVALID_CHARS = CONFIG["invalid_chars"]
INVALID_TEXT = CONFIG["invalid_text"]


# --- Helper Functions ---
def check_invalid_text(wb):
    for sheet in wb.sheet_names:
        df = wb.parse(sheet)
        for col in df.columns:
            if df[col].dtype == object:  # Check only string columns
                for cell in df[col]:
                    if isinstance(cell, str) and any(text in cell for text in INVALID_TEXT):
                        return f"Contains invalid text in sheet '{sheet}'"
    return None


def column_letter(col_idx):
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def check_contains_vietnamese_characters(wb):
    vietnamese_chars = set(INVALID_CHARS)
    for sheet in wb.sheet_names:
        df = wb.parse(sheet)
        for col in df.columns:
            if df[col].dtype == object:  # Check only string columns
                for idx, cell in enumerate(df[col]):
                    if isinstance(cell, str) and any(char in vietnamese_chars for char in cell):
                        col_letter = column_letter(df.columns.get_loc(col) + 1)
                        row_num = idx + 2  # +1 for 1-based index, +1 for header row
                        cell_preview = cell[:50] + "..." if len(cell) > 50 else cell
                        return (
                            f"Contains Vietnamese characters at {sheet}!{col_letter}{row_num} "
                            f"(value: '{cell_preview}')"
                        )
    return None


def check_valid_filename(file_path):
    filename = os.path.basename(file_path)
    parts = os.path.normpath(file_path).split(os.sep)

    for folder_name, expected_prefix in CATEGORY_PREFIX_MAP.items():
        if folder_name in parts:
            if not filename.startswith(expected_prefix):
                return f"Invalid filename for '{folder_name}'"
            break
    return None


def check_invalid_sheet(wb):
    for sheet in INVALID_SHEETS:
        if sheet in wb.sheet_names:
            return f"Contains invalid sheet: {sheet}"
    return None


def check_required_sheets(wb):
    for sheet in REQUIRED_SHEETS:
        if sheet not in wb.sheet_names:
            return f"Missing required sheet: {sheet}"
    return None


def check_confirm_by(wb):
    if "表紙" not in wb.sheet_names:
        return None
    df = wb.parse("表紙")
    # Assuming P24 is column P (16th column), row 24 (0-based index 23)
    if len(df) < 24 or df.iloc[23, 15] is None or pd.isna(df.iloc[23, 15]):
        return "Missing Confirm"
    return None


def find_column_indexes(df, headers=("確認", "参考")):
    return {col: idx for idx, col in enumerate(df.columns) if col in headers}


def check_status_in_test_items(wb, max_rows=1000, empty_limit=10):
    if "テスト項目" not in wb.sheet_names:
        return None

    df = wb.parse("テスト項目", header=2)  # Assuming headers are in row 3 (0-based index 2)
    col_indexes = find_column_indexes(df)
    if "確認" not in col_indexes:
        return "Column '確認' not found"

    status_col = col_indexes["確認"]
    error_rows = []
    consecutive_empty = 0

    for idx, row in df.iterrows():
        if consecutive_empty >= empty_limit:
            break

        b_value = row.iloc[1]  # Column B is index 1
        if not pd.isna(b_value) and str(b_value).strip():
            consecutive_empty = 0
            status_value = row.iloc[status_col]
            if pd.isna(status_value) or str(status_value).strip().upper() != "OK":
                error_rows.append(str(b_value).strip())
        else:
            consecutive_empty += 1

    return (
        f"{len(error_rows)} TC(s) != 'OK': " + "; ".join(error_rows)
        if error_rows
        else None
    )


# --- Main Function ---
def check_excel_file_advanced(file_path, options):
    try:
        error_messages = []
        wb = pd.ExcelFile(file_path)

        if options.get("check_filename_prefix", True):
            if err := check_valid_filename(file_path):
                error_messages.append(err)

        if err := check_required_sheets(wb):
            error_messages.append(err)

        if options.get("check_invalid_sheets", True):
            if err := check_invalid_sheet(wb):
                error_messages.append(err)

        if options.get("check_confirm_cell", True):
            if err := check_confirm_by(wb):
                error_messages.append(err)

        if options.get("check_testcase_status", True):
            if err := check_status_in_test_items(wb):
                error_messages.append(err)

        if options.get("check_contains_vietnamese_characters", True):
            if err := check_contains_vietnamese_characters(wb):
                error_messages.append(err)

        if options.get("check_invalid_text", True):
            if err := check_invalid_text(wb):
                error_messages.append(err)

        return ("ERROR", ", ".join(error_messages)) if error_messages else ("OK", "")

    except Exception as e:
        return "ERROR", str(e)


def find_excel_files_recursive(folder_path):
    excel_files = []
    for root_dir, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(EXCEL_EXTENSIONS):
                excel_files.append(os.path.join(root_dir, file))
    return excel_files


# ----------------- Worker Thread -----------------
class ExcelCheckWorker(QThread):
    progress_changed = pyqtSignal(int)
    file_result = pyqtSignal(
        str, str, str, str
    )  # prefix_path, relative_path, status, error
    finished_signal = pyqtSignal()

    def __init__(self, folder_path, options, max_workers=4):
        super().__init__()
        self.folder_path = folder_path
        self.options = options
        self.max_workers = max_workers
        self._is_running = True

    def run(self):
        files = find_excel_files_recursive(self.folder_path)
        total = len(files)
        if not files:
            self.file_result.emit(self.folder_path, "", "INFO", "No Excel files found.")
            self.finished_signal.emit()
            return

        processed = 0
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(check_excel_file_advanced, file, self.options): file
                for file in files
            }

            for future in as_completed(futures):
                if not self._is_running:
                    break

                file_path = futures[future]
                relative_path = os.path.relpath(file_path, self.folder_path)
                status, error_msg = future.result()

                self.file_result.emit(
                    self.folder_path, relative_path, status, error_msg
                )

                processed += 1
                self.progress_changed.emit(int((processed / total) * 100))

        self.finished_signal.emit()

    def stop(self):
        self._is_running = False


# ----------------- PyQt Main Window -----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Checker")
        self.setGeometry(100, 100, 1000, 600)
        self.worker = None

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # Input section
        input_layout = QHBoxLayout()
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("Paste or type folder path here...")
        self.folder_input.textChanged.connect(self.on_folder_input_change)

        self.btn_select = QPushButton("Browse")
        self.btn_select.clicked.connect(self.select_folder)

        input_layout.addWidget(self.folder_input)
        input_layout.addWidget(self.btn_select)

        # Options section
        option_layout = QVBoxLayout()
        self.confirm_cell_cb = QCheckBox("1. Check confirm")
        self.testcase_status_cb = QCheckBox("2. Check test case status")
        self.filename_check_cb = QCheckBox("3. Check filename prefix")
        self.sheet_check_cb = QCheckBox("4. Check contains invalid sheets")
        self.check_contains_vietnamese_characters_cb = QCheckBox(
            "5. Check contains Vietnamese characters for JP files"
        )
        self.check_invalid_text_cb = QCheckBox("6. Check contains invalid text")

        # Set defaults
        self.confirm_cell_cb.setChecked(True)
        self.testcase_status_cb.setChecked(True)
        self.filename_check_cb.setChecked(True)
        self.sheet_check_cb.setChecked(True)
        self.check_contains_vietnamese_characters_cb.setChecked(False)
        self.check_invalid_text_cb.setChecked(False)

        option_layout.addWidget(self.confirm_cell_cb)
        option_layout.addWidget(self.testcase_status_cb)
        option_layout.addWidget(self.filename_check_cb)
        option_layout.addWidget(self.sheet_check_cb)
        option_layout.addWidget(self.check_contains_vietnamese_characters_cb)
        option_layout.addWidget(self.check_invalid_text_cb)

        # Button section
        button_layout = QHBoxLayout()
        self.btn_execute = QPushButton("Execute")
        self.btn_execute.setEnabled(False)
        self.btn_execute.clicked.connect(self.start_execution)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_execution)
        self.btn_export = QPushButton("Export Results")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export_results)
        button_layout.addWidget(self.btn_export)

        button_layout.addWidget(self.btn_execute)
        button_layout.addWidget(self.btn_stop)

        # Table widget
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            ["Prefix Path", "Relative Path", "Status", "Errors"]
        )
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.itemDoubleClicked.connect(self.open_selected_file)
        self.table.setSortingEnabled(True)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setValue(0)

        # Status label
        config_info_label = QLabel(
            "Note: Case 3, 4, 5, and 6 are configurable.\n"
            "You can change their rules in the 'config.json' file located in the tool's directory."
        )
        config_info_label.setStyleSheet("color: gray; font-size: 11px;")
        config_info_label.setWordWrap(True)

        self.status_label = QLabel("Ready")

        # Assemble main layout
        main_layout.addLayout(input_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(config_info_label)
        main_layout.addLayout(option_layout)
        main_layout.addWidget(self.table)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)

    def on_folder_input_change(self, text):
        self.btn_execute.setEnabled(os.path.isdir(text.strip()))

    def select_folder(self):
        current_path = self.folder_input.text().strip()
        start_dir = (
            current_path if os.path.isdir(current_path) else os.path.expanduser("~")
        )

        if folder := QFileDialog.getExistingDirectory(self, "Select Folder", start_dir):
            self.folder_input.setText(folder)

    def start_execution(self):
        self.btn_export.setEnabled(False)
        folder_path = self.folder_input.text().strip()
        if not os.path.isdir(folder_path):
            QMessageBox.warning(
                self, "Invalid Folder", "Please provide a valid folder path."
            )
            return

        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        self.btn_execute.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.status_label.setText("Processing...")

        options = {
            "check_invalid_sheets": self.sheet_check_cb.isChecked(),
            "check_filename_prefix": self.filename_check_cb.isChecked(),
            "check_confirm_cell": self.confirm_cell_cb.isChecked(),
            "check_testcase_status": self.testcase_status_cb.isChecked(),
            "check_contains_vietnamese_characters": self.check_contains_vietnamese_characters_cb.isChecked(),
            "check_invalid_text": self.check_invalid_text_cb.isChecked(),
        }

        load_config()
        self.worker = ExcelCheckWorker(folder_path, options)
        self.worker.progress_changed.connect(self.progress_bar.setValue)
        self.worker.file_result.connect(self.add_table_row)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def stop_execution(self):
        if self.worker:
            self.worker.stop()
            self.worker.wait()
            self.status_label.setText("Process stopped by user")
            self.btn_execute.setEnabled(True)
            self.btn_stop.setEnabled(False)

    def add_table_row(self, prefix_path, path, status, error):
        row = self.table.rowCount()
        self.table.insertRow(row)

        items = [
            QTableWidgetItem(prefix_path),
            QTableWidgetItem(path.replace("\\", "/")),
            QTableWidgetItem(status),
            QTableWidgetItem(error),
        ]

        if status == "OK":
            items[2].setForeground(QColor("green"))
        elif status == "ERROR":
            items[2].setForeground(QColor("red"))

        for col, item in enumerate(items):
            self.table.setItem(row, col, item)

        if row == 0:
            self.btn_export.setEnabled(True)

    def open_selected_file(self, item):
        row = item.row()
        path = os.path.join(self.folder_input.text(), self.table.item(row, 1).text())

        if os.path.exists(path):
            try:
                modifiers = QApplication.keyboardModifiers()
                if modifiers == Qt.ControlModifier:
                    folder_path = os.path.dirname(path)
                    if os.name == "nt":
                        os.startfile(folder_path)
                    elif sys.platform == "darwin":
                        subprocess.call(["open", folder_path])
                    else:
                        subprocess.call(["xdg-open", folder_path])
                else:
                    # Open the file
                    if os.name == "nt":
                        os.startfile(path)
                    elif sys.platform == "darwin":
                        subprocess.call(["open", path])
                    else:
                        subprocess.call(["xdg-open", path])
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open: {str(e)}")
        else:
            QMessageBox.warning(self, "File Not Found", f"File not found: {path}")

    def on_finished(self):
        self.btn_execute.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.status_label.setText("Process completed")
        QMessageBox.information(self, "Done", "Check completed.")

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
        event.accept()

    def export_results(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "There are no results to export.")
            return

        default_name = (
            f"Excel_Check_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        # Get save file path
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", default_name, "Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path:
            return  # User cancelled

        # Ensure .xlsx extension
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            # Create a DataFrame from the table data
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)

            df = pd.DataFrame(data, columns=["Prefix Path", "Relative Path", "Status", "Errors"])

            # Export to Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Check Results')
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Check Results']
            
            # Add a header format
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1,
                'align': 'center'
            })
            
            # Write the column headers with the defined format
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                max_len = max((
                    df[col].astype(str).map(len).max(),  # Max length in column
                    len(col)  # Length of column header
                )) + 2  # Add a little extra space
                worksheet.set_column(i, i, max_len)
            
            writer.close()
            
            QMessageBox.information(
                self, "Success", f"Results exported to:\n{file_path}"
            )

        except Exception as e:
            QMessageBox.critical(
                self, "Export Error", f"Failed to export results:\n{str(e)}"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())