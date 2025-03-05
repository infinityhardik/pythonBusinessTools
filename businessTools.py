#!/usr/bin/env python3
"""
Business Tools – A unified PDF processing application

This application offers three main operations:
  1. Split PDF by Party Name – splits a PDF file into segments based on a party name,
     where the party name is detected dynamically using a keyword and line offset.
  2. Extract PDF by Order ID – splits a PDF into segments based on a combination
     of party name and order ID extraction using dynamic options.
  3. Extract Text from PDF – extracts text from PDFs (using OCR as a fallback),
     saves the output as a text file, and prints the extracted text in the log dialog.

The dynamic extraction options (keywords, offsets, regex patterns) are configurable
from the GUI so you can modify the extraction logic without changing the code.
"""

import sys
import os
import re
import io

import fitz  # PyMuPDF for PDF processing
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QPushButton,
    QFileDialog,
    QLabel,
    QComboBox,
    QLineEdit,
    QTextEdit,
    QMessageBox,
    QGroupBox,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PIL import Image
import pytesseract

# ----------------------- Utility Functions -----------------------


def extract_party_name_dynamic(text, party_keyword="M/s.", offset=-1):
    """
    Extracts the party name dynamically based on a keyword and a line offset.
    Default: look for "M/s." and take the previous line (offset = -1).
    """
    lines = text.splitlines()
    party_name = None
    for i in range(len(lines)):
        if party_keyword in lines[i]:
            idx = i + offset
            if 0 <= idx < len(lines):
                party_name = lines[idx].strip()
                break
    return party_name


def extract_order_id_dynamic(text, order_keyword="ID :", regex_pattern=r"ID : (\d+)"):
    """
    Extracts the order ID dynamically.
    Default: looks for "ID :" and uses a regex to capture numeric order ID.
    """
    lines = text.splitlines()
    order_id = None
    for line in lines:
        if order_keyword in line:
            match = re.search(regex_pattern, line)
            if match:
                order_id = match.group(1)
                break
    return order_id


def extract_text_from_pdf(pdf_path):
    """
    Extracts text from a PDF. Uses direct text extraction if available;
    otherwise, uses OCR on the rendered page image.
    """
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_text = page.get_text("text")
            if page_text.strip():
                text += page_text
            else:
                # Fallback to OCR if no text is extracted
                pix = page.get_pixmap()
                img = Image.open(io.BytesIO(pix.tobytes()))
                ocr_text = pytesseract.image_to_string(img)
                text += ocr_text
        doc.close()
    except Exception as e:
        text = f"Error processing {pdf_path}: {str(e)}"
    return text


def sanitize_filename(filename):
    """
    Removes invalid characters for Windows filenames.
    """
    return re.sub(r'[<>:"/\\|?*]', "", filename)


# ----------------------- PDF Processing Functions -----------------------


def split_pdf_by_party_name(pdf_path, output_directory, party_keyword, party_offset):
    """
    Splits the input PDF into segments based on the party name.
    The party name is extracted using the dynamic parameters.
    """
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    current_party = None
    current_start_page = 0
    splits = []  # List of tuples: (start_page, end_page, party_name)

    for page_num in range(num_pages):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        party = extract_party_name_dynamic(text, party_keyword, party_offset)
        if party:
            if current_party is not None:
                splits.append((current_start_page, page_num - 1, current_party))
            current_party = party
            current_start_page = page_num

    if current_party:
        splits.append((current_start_page, num_pages - 1, current_party))

    for start_page, end_page, party in splits:
        new_pdf = fitz.open()
        for i in range(start_page, end_page + 1):
            new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        sanitized_party = sanitize_filename(party)
        output_filename = f"{sanitized_party}.pdf"
        output_path = os.path.join(output_directory, output_filename)
        new_pdf.save(output_path)
        new_pdf.close()
    doc.close()
    return f"Processed (Split by Party Name): {os.path.basename(pdf_path)}"


def extract_pdf_by_order_id(
    pdf_path, output_directory, party_keyword, party_offset, order_keyword, order_regex
):
    """
    Splits the PDF into segments based on both the party name and order ID.
    Uses dynamic parameters for both extractions.
    """
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    current_party = None
    current_order_id = None
    current_start_page = 0
    splits = []  # List of tuples: (start_page, end_page, party_name, order_id)

    for page_num in range(num_pages):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        party = extract_party_name_dynamic(text, party_keyword, party_offset)
        order_id = extract_order_id_dynamic(text, order_keyword, order_regex)
        if party and order_id:
            if current_party is not None:
                splits.append(
                    (current_start_page, page_num - 1, current_party, current_order_id)
                )
            current_party = party
            current_order_id = order_id
            current_start_page = page_num

    if current_party:
        splits.append(
            (current_start_page, num_pages - 1, current_party, current_order_id)
        )

    for start_page, end_page, party, order_id in splits:
        new_pdf = fitz.open()
        for i in range(start_page, end_page + 1):
            new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        combined_name = sanitize_filename(f"{party}_{order_id}")
        output_filename = f"{combined_name}.pdf"
        output_path = os.path.join(output_directory, output_filename)
        new_pdf.save(output_path)
        new_pdf.close()
    doc.close()
    return f"Processed (Extract by Order ID): {os.path.basename(pdf_path)}"


def process_text_extraction(pdf_path, output_directory):
    """
    Extracts text from the PDF, saves it as a .txt file, and returns the extracted text.
    """
    text = extract_text_from_pdf(pdf_path)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    output_filename = f"{sanitize_filename(base)}.txt"
    output_path = os.path.join(output_directory, output_filename)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    return f"Processed (Text Extraction): {os.path.basename(pdf_path)}", text


# ----------------------- Directory Processing Wrappers -----------------------


def process_split_operation(input_dir, output_dir, party_keyword, party_offset):
    messages = []
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            msg = split_pdf_by_party_name(
                pdf_path, output_dir, party_keyword, party_offset
            )
            messages.append(msg)
    return messages


def process_extract_operation(
    input_dir, output_dir, party_keyword, party_offset, order_keyword, order_regex
):
    messages = []
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            msg = extract_pdf_by_order_id(
                pdf_path,
                output_dir,
                party_keyword,
                party_offset,
                order_keyword,
                order_regex,
            )
            messages.append(msg)
    return messages


def process_text_extraction_operation(input_dir, output_dir):
    results = []
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            msg, text = process_text_extraction(pdf_path, output_dir)
            results.append((msg, text))
    return results


# ----------------------- GUI Application -----------------------


class BusinessToolsUI(QWidget):
    def __init__(self):
        super().__init__()
        self.default_party_keyword = "M/s."
        self.default_party_offset = "-1"
        self.default_order_keyword = "ID :"
        self.default_order_regex = r"ID : (\d+)"
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Business Tools")
        self.resize(750, 650)
        # Set a global font size for better readability.
        self.setStyleSheet("QWidget { font-size: 12pt; }")

        main_layout = QVBoxLayout()

        # -------------------- Directories Group --------------------
        dir_group = QGroupBox("Directories")
        dir_layout = QFormLayout()

        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("Select input folder...")
        btn_input = QPushButton("Browse")
        btn_input.clicked.connect(self.browse_input)
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_line)
        input_layout.addWidget(btn_input)
        dir_layout.addRow(QLabel("Input Directory:"), input_layout)

        self.output_line = QLineEdit()
        self.output_line.setPlaceholderText("Select output folder...")
        btn_output = QPushButton("Browse")
        btn_output.clicked.connect(self.browse_output)
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_line)
        output_layout.addWidget(btn_output)
        dir_layout.addRow(QLabel("Output Directory:"), output_layout)
        dir_group.setLayout(dir_layout)
        main_layout.addWidget(dir_group)

        # -------------------- Operation Selection Group --------------------
        op_group = QGroupBox("Operation")
        op_layout = QHBoxLayout()
        self.operation_combo = QComboBox()
        self.operation_combo.addItems(
            ["Split by Party Name", "Extract by Order ID", "Extract Text from PDF"]
        )
        op_layout.addWidget(QLabel("Select Operation:"))
        op_layout.addWidget(self.operation_combo)
        op_group.setLayout(op_layout)
        main_layout.addWidget(op_group)

        # -------------------- Dynamic Extraction Options Group --------------------
        dynamic_group = QGroupBox("Dynamic Extraction Options")
        dyn_layout = QFormLayout()
        # Party Name Extraction Options
        self.party_keyword_edit = QLineEdit(self.default_party_keyword)
        self.party_offset_edit = QLineEdit(self.default_party_offset)
        dyn_layout.addRow(QLabel("Party Keyword:"), self.party_keyword_edit)
        dyn_layout.addRow(
            QLabel("Party Offset (line relative to keyword):"), self.party_offset_edit
        )
        # Order ID Extraction Options
        self.order_keyword_edit = QLineEdit(self.default_order_keyword)
        self.order_regex_edit = QLineEdit(self.default_order_regex)
        dyn_layout.addRow(QLabel("Order ID Keyword:"), self.order_keyword_edit)
        dyn_layout.addRow(QLabel("Order ID Regex Pattern:"), self.order_regex_edit)
        # Reset Button for Extraction Options
        btn_reset = QPushButton("Reset Extraction Options to Default")
        btn_reset.clicked.connect(self.reset_extraction_options)
        dyn_layout.addRow(btn_reset)
        dynamic_group.setLayout(dyn_layout)
        main_layout.addWidget(dynamic_group)

        # -------------------- Log Output --------------------
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        main_layout.addWidget(QLabel("Log Output:"))
        main_layout.addWidget(self.log_output)

        # -------------------- Start Button --------------------
        self.start_btn = QPushButton("Start Processing")
        self.start_btn.clicked.connect(self.start_processing)
        main_layout.addWidget(self.start_btn)

        self.setLayout(main_layout)

    def browse_input(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.input_line.setText(folder)

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_line.setText(folder)

    def reset_extraction_options(self):
        """
        Resets the dynamic extraction options to their default values.
        """
        self.party_keyword_edit.setText(self.default_party_keyword)
        self.party_offset_edit.setText(self.default_party_offset)
        self.order_keyword_edit.setText(self.default_order_keyword)
        self.order_regex_edit.setText(self.default_order_regex)
        self.log_message("Extraction options reset to default.")

    def log_message(self, message):
        self.log_output.append(message)

    def start_processing(self):
        # Get directory paths
        input_dir = self.input_line.text().strip()
        output_dir = self.output_line.text().strip()
        if not input_dir or not output_dir:
            QMessageBox.warning(
                self, "Error", "Please select both input and output directories."
            )
            return

        os.makedirs(output_dir, exist_ok=True)

        # Get dynamic extraction options
        party_keyword = self.party_keyword_edit.text().strip()
        try:
            party_offset = int(self.party_offset_edit.text().strip())
        except ValueError:
            QMessageBox.warning(self, "Error", "Party Offset must be an integer.")
            return
        order_keyword = self.order_keyword_edit.text().strip()
        order_regex = self.order_regex_edit.text().strip()

        operation = self.operation_combo.currentText()
        self.log_output.clear()
        self.log_message("Processing started...")

        try:
            if operation == "Split by Party Name":
                messages = process_split_operation(
                    input_dir, output_dir, party_keyword, party_offset
                )
                for msg in messages:
                    self.log_message(msg)
            elif operation == "Extract by Order ID":
                messages = process_extract_operation(
                    input_dir,
                    output_dir,
                    party_keyword,
                    party_offset,
                    order_keyword,
                    order_regex,
                )
                for msg in messages:
                    self.log_message(msg)
            elif operation == "Extract Text from PDF":
                results = process_text_extraction_operation(input_dir, output_dir)
                for msg, text in results:
                    self.log_message(msg)
                    self.log_message("Extracted Text:")
                    self.log_message(text)
                    self.log_message("-" * 40)
            else:
                self.log_message("Unknown operation selected.")
            self.log_message("Processing completed.")
        except Exception as e:
            self.log_message(f"Error during processing: {str(e)}")
            QMessageBox.critical(self, "Processing Error", str(e))


# ----------------------- Main Execution -----------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BusinessToolsUI()
    window.show()
    sys.exit(app.exec_())
