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
    QCheckBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QColor
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed, CancelledError
import json
from datetime import datetime
import re
from threading import Event
import xlwings as xw
import zipfile
import xml.etree.ElementTree as ET


def load_config(config_path="config.json"):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()

# --- Constants ---
CATEGORY_PREFIX_MAP = CONFIG["category_prefix_map"]
INVALID_SHEETS = set(CONFIG["invalid_sheets"])
REQUIRED_SHEETS = set(CONFIG["required_sheets"])
EXCEL_EXTENSIONS = tuple(CONFIG["excel_extensions"])
INVALID_CHARS = set(CONFIG["invalid_chars"])
INVALID_TEXT = set(CONFIG["invalid_text"])


# --- Helper Functions v2 ---
def get_shared_strings(zip_ref):
    try:
        with zip_ref.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            strings = []
            for si in root.findall('a:si', ns):
                text_parts = [t.text for t in si.findall('.//a:t', ns) if t.text]
                strings.append(''.join(text_parts))
            return strings
    except KeyError:
        return []

def get_sheet_names(zip_ref):
    with zip_ref.open('xl/workbook.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        return [sheet.attrib['name'] for sheet in root.findall('.//a:sheet', ns)]

def read_cells_from_sheet(zip_ref, sheet_filename, shared_strings):
    with zip_ref.open(sheet_filename) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        values = []

        for c in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            cell_type = c.attrib.get('t')
            v = c.find('a:v', ns)
            if v is not None:
                value = v.text
                if cell_type == 's':
                    value = shared_strings[int(value)]
                values.append(value)
        return values

def check_invalid_text_zip(cell_values, invalid_text_set):
    for value in cell_values:
        if isinstance(value, str) and any(t in value for t in invalid_text_set):
            return f"Contains invalid text: {value}"
    return None

def check_contains_vn_chars_zip(cell_values, invalid_chars):
    pattern = re.compile(f"[{''.join(re.escape(c) for c in invalid_chars)}]")
    for value in cell_values:
        if isinstance(value, str) and pattern.search(value):
            return f"Contains Vietnamese character: {value}"
    return None

def check_incorrect_textbox(zip_ref):
    try:
        with zip_ref.open("xl/drawings/drawing1.xml") as f:
            tree = ET.parse(f)
            ns = {
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
            }
            root = tree.getroot()
            texts = []
            for txBody in root.findall(".//xdr:txBody", ns):
                for p in txBody.findall(".//a:p", ns):
                    text_parts = [t.text for t in p.findall(".//a:t", ns) if t.text]
                    if text_parts:
                        texts.append("".join(text_parts))
            for text in texts:
                if not text or "API" in text:
                    return f"Incorrect TextBox content: '{text}'"
    except KeyError:
        pass
    return None

def check_excel_file_advanced_zip(file_path, options, stop_event=None):
    if stop_event and stop_event.is_set():
        return "CANCELLED", "Stopped by user"

    try:
        error_messages = []

        with zipfile.ZipFile(file_path, "r") as zip_ref:
            shared_strings = get_shared_strings(zip_ref)
            sheet_names = get_sheet_names(zip_ref)
            sheet_files = [f for f in zip_ref.namelist() if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]

            # --- filename prefix check ---
            if options.get("check_filename_prefix", True):
                err = check_valid_filename(file_path)
                if err:
                    error_messages.append(err)

            # --- invalid sheets ---
            if options.get("check_invalid_sheets", True):
                for sheet in INVALID_SHEETS:
                    if sheet in sheet_names:
                        error_messages.append(f"Contains invalid sheet: {sheet}")

            # --- required sheets ---
            if options.get("check_required_sheets", True):
                for sheet in REQUIRED_SHEETS:
                    if sheet not in sheet_names:
                        error_messages.append(f"Missing required sheet: {sheet}")

            # --- check cell contents ---
            for sheet_file in sheet_files:
                if stop_event and stop_event.is_set():
                    return "CANCELLED", "Stopped by user"
                cell_values = read_cells_from_sheet(zip_ref, sheet_file, shared_strings)

                if options.get("check_invalid_text", True):
                    err = check_invalid_text_zip(cell_values, INVALID_TEXT)
                    if err:
                        error_messages.append(err)

                if options.get("check_contains_vietnamese_characters", True):
                    err = check_contains_vn_chars_zip(cell_values, INVALID_CHARS)
                    if err:
                        error_messages.append(err)

            # --- check text box ---
            if options.get("check_incorrect_tb_content", True):
                err = check_incorrect_textbox(zip_ref)
                if err:
                    error_messages.append(err)

        return ("ERROR", ", ".join(error_messages)) if error_messages else ("OK", "")

    except Exception as e:
        return "ERROR", str(e)

# --- Helper Functions ---
def check_invalid_text(wb, stop_event=None):
    if stop_event and stop_event.is_set():
        return "CANCELLED", "Process was cancelled by user."
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        if stop_event and stop_event.is_set():
            return "CANCELLED", "Process was cancelled by user."
        for row in ws.iter_rows(values_only=True):
            if stop_event and stop_event.is_set():
                return "CANCELLED", "Process was cancelled by user."
            for cell in row:
                if isinstance(cell, str) and any(text in cell for text in INVALID_TEXT):
                    return f"Contains invalid text in sheet '{sheet}'"
    return None


def check_contains_vietnamese_characters(wb, stop_event=None):
    results = ""
    if stop_event and stop_event.is_set():
        return results
    pattern = re.compile(f"[{''.join(re.escape(c) for c in INVALID_CHARS)}]")
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        if stop_event and stop_event.is_set():
            return results
        for row in ws.iter_rows():
            if stop_event and stop_event.is_set():
                return results
            for cell in row:
                if stop_event and stop_event.is_set():
                    return results
                if cell.value and isinstance(cell.value, str):
                    if pattern.search(cell.value):
                        print(ws.title, cell.value, cell.coordinate)
                        results += (
                            f"sheet: {ws.title}, "
                            f"cell: {cell.coordinate}, "
                            f"value: {cell.value}; "
                        )
                        break
    return results


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
        if sheet in wb.sheetnames:
            return f"Contains invalid sheet: {sheet}"
    return None


def check_required_sheets(wb):
    for sheet in REQUIRED_SHEETS:
        if sheet not in wb.sheetnames:
            return f"Missing required sheet: {sheet}"
    return None


def check_confirm_by(wb):
    if "表紙" not in wb.sheetnames:
        return f"Missing required sheet: '表紙'"
    ws = wb["表紙"]
    for row in ws.iter_rows(
        min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for cell in row:
            if cell.value == "確認":
                if ws.cell(row=cell.row + 1, column=cell.column).value is None:
                    return "Missing Confirm"
                else:
                    return None


def find_column_indexes(ws, headers=("確認", "参考"), header_rows=(3, 4)):
    found = {}
    for row in header_rows:
        for cell in ws[row]:
            if cell.value in headers and cell.value not in found:
                found[cell.value] = cell.column
        if len(found) == len(headers):
            break
    return found


def check_status_in_test_items(wb, max_rows=1000, empty_limit=10):
    if "テスト項目" not in wb.sheetnames:
        return f"Missing required sheet: 'テスト項目"

    ws = wb["テスト項目"]
    col_indexes = find_column_indexes(ws)
    if "確認" not in col_indexes:
        return "Column '確認' not found"

    status_col = col_indexes["確認"]
    error_rows = []
    consecutive_empty = 0

    for row in range(5, max_rows + 1):
        if consecutive_empty >= empty_limit:
            break

        b_cell = ws.cell(row=row, column=2)
        if b_cell.value and str(b_cell.value).strip():
            consecutive_empty = 0
            status_value = ws.cell(row=row, column=status_col).value
            if status_value is None or str(status_value).strip().upper() != "OK":
                error_rows.append(str(b_cell.value).strip())
        else:
            consecutive_empty += 1

    return (
        f"{len(error_rows)} TC(s) != 'OK': " + "; ".join(error_rows)
        if error_rows
        else None
    )


def check_incorrect_tb_content(wb, file_path):
    print(file_path)
    app = None
    try:
        app = xw.App(visible=False)
        xlwb = app.books.open(file_path, read_only=True)
        error_msgs = []
        try:
            shape = next(
                (s for s in xlwb.sheets[0].shapes if s.name == "Text Box 1"), None
            )
            if shape:
                text = ""
                try:
                    text = shape.text
                except Exception:
                    try:
                        text = shape.api.TextFrame2.TextRange.Text
                    except Exception:
                        text = ""
                if not text or "API" in text:
                    error_msgs.append(
                        f"Sheet '{xlwb.sheets[0].name}': 'Text Box 1' incorrect content: '{text}'"
                    )
        except Exception as e:
            error_msgs.append(
                f"Sheet '{xlwb.sheets[0].name}': Error reading Text Box 1 ({e})"
            )
        xlwb.close()
        if app:
            app.quit()
        return "; ".join(error_msgs) if error_msgs else None
    except Exception as e:
        if app:
            app.quit()
        return f"Error opening file with xlwings: {e}"


# --- Main Function ---
def check_excel_file_advanced(file_path, options, stop_event=None):
    if stop_event and stop_event.is_set():
        return "CANCELLED", "Process was cancelled by user."
    try:
        error_messages = []
        wb = load_workbook(file_path, data_only=True, read_only=True)

        if stop_event and stop_event.is_set():
            wb.close()
            return "CANCELLED", "Process was cancelled by user."

        if options.get("check_filename_prefix", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_valid_filename(file_path):
                error_messages.append(err)

        if options.get("check_invalid_sheets", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_invalid_sheet(wb):
                error_messages.append(err)

        if options.get("check_required_sheets", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_required_sheets(wb):
                error_messages.append(err)

        if options.get("check_confirm_cell", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_confirm_by(wb):
                error_messages.append(err)

        if options.get("check_testcase_status", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_status_in_test_items(wb):
                error_messages.append(err)

        if options.get("check_contains_vietnamese_characters", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_contains_vietnamese_characters(wb, stop_event):
                error_messages.append(err)

        if options.get("check_invalid_text", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_invalid_text(wb):
                error_messages.append(err)

        if options.get("check_incorrect_tb_content", True):
            if stop_event and stop_event.is_set():
                wb.close()
                return "CANCELLED", "Process was cancelled by user."
            if err := check_incorrect_tb_content(wb, file_path):
                error_messages.append(err)

        wb.close()
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
        self._stop_event = Event()

    def run(self):
        files = find_excel_files_recursive(self.folder_path)
        total = len(files)
        if not files:
            self.file_result.emit(self.folder_path, "", "INFO", "No Excel files found.")
            self.finished_signal.emit()
            return

        processed = 0
        chunk_size = 4

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            for i in range(0, total, chunk_size):
                if self._stop_event.is_set():
                    break

                current_chunk = files[i : i + chunk_size]
                chunk_futures = {
                    executor.submit(
                        check_excel_file_advanced_zip, file, self.options, self._stop_event
                    ): file
                    for file in current_chunk
                }

                for future in as_completed(chunk_futures):
                    if self._stop_event.is_set():
                        break

                    file_path = chunk_futures[future]
                    try:
                        relative_path = os.path.relpath(file_path, self.folder_path)
                        status, error_msg = future.result()
                        self.file_result.emit(
                            self.folder_path, relative_path, status, error_msg
                        )
                    except Exception as e:
                        self.file_result.emit(
                            self.folder_path,
                            os.path.relpath(file_path, self.folder_path),
                            "ERROR",
                            str(e),
                        )

                    processed += 1
                    self.progress_changed.emit(int((processed / total) * 100))

        self.finished_signal.emit()

    def stop(self):
        self._stop_event.set()


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
        self.sheet_req_check_cb = QCheckBox("2. Check required sheets")
        self.testcase_status_cb = QCheckBox("3. Check test case status")
        self.filename_check_cb = QCheckBox("4. Check filename prefix")
        self.sheet_check_cb = QCheckBox("5. Check contains invalid sheets")
        self.check_contains_vietnamese_characters_cb = QCheckBox(
            "6. Check contains Vietnamese characters for JP files"
        )
        self.check_invalid_text_cb = QCheckBox("7. Check contains invalid text")
        self.check_incorrect_tb_content_cb = QCheckBox(
            "8. Check incorrect 'Text Box 1' content"
        )

        # Set defaults
        self.confirm_cell_cb.setChecked(False)
        self.testcase_status_cb.setChecked(False)
        self.filename_check_cb.setChecked(False)
        self.sheet_req_check_cb.setChecked(False)
        self.sheet_check_cb.setChecked(False)
        self.check_contains_vietnamese_characters_cb.setChecked(False)
        self.check_invalid_text_cb.setChecked(False)
        self.check_incorrect_tb_content_cb.setChecked(False)

        option_layout.addWidget(self.confirm_cell_cb)
        option_layout.addWidget(self.sheet_req_check_cb)
        option_layout.addWidget(self.testcase_status_cb)
        option_layout.addWidget(self.filename_check_cb)
        option_layout.addWidget(self.sheet_check_cb)
        option_layout.addWidget(self.check_contains_vietnamese_characters_cb)
        option_layout.addWidget(self.check_invalid_text_cb)
        option_layout.addWidget(self.check_incorrect_tb_content_cb)

        # Button section
        button_layout = QHBoxLayout()
        self.btn_execute = QPushButton("Execute")
        self.btn_execute.setEnabled(False)
        self.btn_execute.clicked.connect(self.start_execution)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_execution)
        self.btn_export = QPushButton("Export results to Excel")
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
            "Note: Case 2, 3, 4, 5, and 6 are configurable.\n"
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
            "check_required_sheets": self.sheet_req_check_cb.isChecked(),
            "check_confirm_cell": self.confirm_cell_cb.isChecked(),
            "check_testcase_status": self.testcase_status_cb.isChecked(),
            "check_contains_vietnamese_characters": self.check_contains_vietnamese_characters_cb.isChecked(),
            "check_invalid_text": self.check_invalid_text_cb.isChecked(),
            "check_incorrect_tb_content": self.check_incorrect_tb_content_cb.isChecked(),
        }

        load_config()
        self.worker = ExcelCheckWorker(folder_path, options)
        self.worker.progress_changed.connect(self.progress_bar.setValue)
        self.worker.file_result.connect(self.add_table_row)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def stop_execution(self):
        if self.worker:
            self.btn_stop.setText("Stopping...")
            self.btn_stop.setEnabled(False)
            self.btn_execute.setEnabled(False)
            self.btn_export.setEnabled(False)
            self.progress_bar.setValue(0)
            self.status_label.setText("Process stopped by user... ")
            QApplication.processEvents()
            self.worker.stop()
            self.btn_stop.setText("Stop")
            self.btn_execute.setEnabled(True)

    def add_table_row(self, prefix_path, path, status, error):
        # self.status_label.setText(f"Processing: {path}")
        row = self.table.rowCount()
        self.table.insertRow(row)

        items = [
            QTableWidgetItem(prefix_path.replace("/", "\\")),
            QTableWidgetItem(path),
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
        path = os.path.join(
            self.table.item(row, 0).text(), self.table.item(row, 1).text()
        )
        print(f"Opening file: {path}")

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

        if self.worker._stop_event.is_set():
            self.status_label.setText("Process stopped by user")
            QMessageBox.information(self, "Stopped", "Process was stopped by user.")
        else:
            total_files = self.table.rowCount()
            error_count = sum(
                1
                for row in range(total_files)
                if self.table.item(row, 2).text() == "ERROR"
            )
            ok_count = total_files - error_count
            summary = f"Check completed.\nTotal files: {total_files}\nOK: {ok_count}\nErrors: {error_count}"
            self.status_label.setText("Process completed")
            QMessageBox.information(self, "Done", summary)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
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
            from openpyxl import Workbook
            from openpyxl.styles import Font, Color, Alignment

            wb = Workbook()
            ws = wb.active
            ws.title = "Check Results"

            # Write headers
            headers = ["Prefix Path", "Relative Path", "Status", "Errors"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")

            # Write data
            for row in range(self.table.rowCount()):
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    ws.cell(
                        row=row + 2, column=col + 1, value=item.text() if item else ""
                    )

            # Auto-size columns
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(file_path)
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
