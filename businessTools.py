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
Additionally, you can now choose whether the input is a single file or a folder.
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


def extract_order_id_dynamic(text, order_keyword="ID :", regex_pattern=r"ID :\s*(\d+)"):
    """
    Extracts the order ID dynamically.
    This function first checks the line that contains the order keyword.
    If a numeric order ID is not found on that line, it will check the next line.
    """
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if order_keyword in line:
            # Try to extract number from the same line
            match = re.search(regex_pattern, line)
            if match:
                return match.group(1)
            # Otherwise, check the next non-empty line for a numeric value.
            j = i + 1
            while j < len(lines):
                candidate = lines[j].strip()
                # Skip if the candidate is just a colon or similar punctuation.
                if candidate == ":":
                    j += 1
                    continue
                if candidate:
                    if candidate.isdigit():
                        return candidate
                    match = re.search(r"(\d+)", candidate)
                    if match:
                        return match.group(1)
                    break
                j += 1
    return None


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
        # Defaults
        self.default_party_keyword = "M/s."
        self.default_party_offset = "-1"
        self.default_order_keyword = "ID :"
        self.default_order_regex = r"ID :\s*(\d+)"
        # Invoice defaults for order extraction
        self.invoice_order_keyword = "Invoice No."
        self.invoice_order_regex = r"Invoice No\.\s*:\s*(\d+)"

        # Define extraction profiles
        # "Custom" means no preset – user input is preserved.
        self.extraction_profiles = {
            "Defaults": {
                "party_keyword": self.default_party_keyword,
                "party_offset": self.default_party_offset,
                "order_keyword": self.default_order_keyword,
                "order_regex": self.default_order_regex,
            },
            "Invoice Defaults": {
                "party_keyword": self.default_party_keyword,
                "party_offset": self.default_party_offset,
                "order_keyword": self.invoice_order_keyword,
                "order_regex": self.invoice_order_regex,
            },
            "Custom": None,
        }
        # Set the default profile to "Defaults"
        self.current_profile = "Defaults"
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Business Tools")
        self.resize(750, 680)
        # Increase global font size for better readability.
        self.setStyleSheet("QWidget { font-size: 12pt; }")

        main_layout = QVBoxLayout()

        # -------------------- Directories Group --------------------
        dir_group = QGroupBox("Directories")
        dir_layout = QFormLayout()

        # Input Type Selection (File or Folder)
        self.input_type_combo = QComboBox()
        self.input_type_combo.addItems(["File", "Folder"])
        dir_layout.addRow(QLabel("Input Type:"), self.input_type_combo)

        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("Select input path...")
        btn_input = QPushButton("Browse")
        btn_input.clicked.connect(self.browse_input)
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_line)
        input_layout.addWidget(btn_input)
        dir_layout.addRow(QLabel("Input Path:"), input_layout)

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

        # Extraction Profile Selection
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(list(self.extraction_profiles.keys()))
        self.profile_combo.currentIndexChanged.connect(self.profile_changed)
        dyn_layout.addRow(QLabel("Extraction Profile:"), self.profile_combo)

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

        # Connect signals to detect manual modifications.
        self.party_keyword_edit.textChanged.connect(self.dynamic_fields_modified)
        self.party_offset_edit.textChanged.connect(self.dynamic_fields_modified)
        self.order_keyword_edit.textChanged.connect(self.dynamic_fields_modified)
        self.order_regex_edit.textChanged.connect(self.dynamic_fields_modified)

        # Reset Button for Extraction Options
        btn_reset = QPushButton("Reset Extraction Options to Profile Default")
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

    def profile_changed(self):
        profile = self.profile_combo.currentText()
        self.current_profile = profile
        defaults = self.extraction_profiles.get(profile)
        if defaults is not None:
            # Block signals during programmatic updates
            self.party_keyword_edit.blockSignals(True)
            self.party_offset_edit.blockSignals(True)
            self.order_keyword_edit.blockSignals(True)
            self.order_regex_edit.blockSignals(True)
            self.party_keyword_edit.setText(defaults["party_keyword"])
            self.party_offset_edit.setText(defaults["party_offset"])
            self.order_keyword_edit.setText(defaults["order_keyword"])
            self.order_regex_edit.setText(defaults["order_regex"])
            self.party_keyword_edit.blockSignals(False)
            self.party_offset_edit.blockSignals(False)
            self.order_keyword_edit.blockSignals(False)
            self.order_regex_edit.blockSignals(False)
            self.log_message(
                f"Profile changed to '{profile}'. Extraction options set to default values."
            )
        else:
            self.log_message(
                "Profile changed to 'Custom'. You can now enter your own extraction options."
            )

    def dynamic_fields_modified(self):
        # If the current profile is not already Custom, check for changes.
        if self.current_profile != "Custom":
            current_defaults = self.extraction_profiles[self.current_profile]
            if (
                self.party_keyword_edit.text() != current_defaults["party_keyword"]
                or self.party_offset_edit.text() != current_defaults["party_offset"]
                or self.order_keyword_edit.text() != current_defaults["order_keyword"]
                or self.order_regex_edit.text() != current_defaults["order_regex"]
            ):
                # Switch to Custom profile without triggering the profile_changed slot recursively.
                self.profile_combo.blockSignals(True)
                self.profile_combo.setCurrentText("Custom")
                self.current_profile = "Custom"
                self.profile_combo.blockSignals(False)
                self.log_message(
                    "Extraction fields modified. Profile switched to 'Custom'."
                )

    def browse_input(self):
        input_type = self.input_type_combo.currentText()
        if input_type == "File":
            filename, _ = QFileDialog.getOpenFileName(
                self, "Select PDF File", "", "PDF Files (*.pdf)"
            )
            if filename:
                self.input_line.setText(filename)
        else:  # Folder
            folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
            if folder:
                self.input_line.setText(folder)

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_line.setText(folder)

    def reset_extraction_options(self):
        """Resets the dynamic extraction options to the defaults of the current profile."""
        profile = self.profile_combo.currentText()
        defaults = self.extraction_profiles.get(profile)
        if defaults is not None:
            self.party_keyword_edit.blockSignals(True)
            self.party_offset_edit.blockSignals(True)
            self.order_keyword_edit.blockSignals(True)
            self.order_regex_edit.blockSignals(True)
            self.party_keyword_edit.setText(defaults["party_keyword"])
            self.party_offset_edit.setText(defaults["party_offset"])
            self.order_keyword_edit.setText(defaults["order_keyword"])
            self.order_regex_edit.setText(defaults["order_regex"])
            self.party_keyword_edit.blockSignals(False)
            self.party_offset_edit.blockSignals(False)
            self.order_keyword_edit.blockSignals(False)
            self.order_regex_edit.blockSignals(False)
            self.log_message(f"Extraction options reset to '{profile}' defaults.")
        else:
            self.log_message("Custom extraction options remain unchanged.")

    def log_message(self, message):
        self.log_output.append(message)

    def start_processing(self):
        # Do not clear previous log messages; instead, add a separator
        if self.log_output.toPlainText():
            self.log_message("----------------------------------------")
        self.log_message("Processing started...")

        # Get input and output paths
        input_path = self.input_line.text().strip()
        output_dir = self.output_line.text().strip()
        if not input_path or not output_dir:
            QMessageBox.warning(
                self,
                "Error",
                "Please select both an input path and an output directory.",
            )
            return

        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        # Capture current files in the output directory
        initial_output_files = set(os.listdir(output_dir))

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
        input_type = self.input_type_combo.currentText()

        try:
            if input_type == "File":
                if not os.path.isfile(input_path):
                    QMessageBox.warning(
                        self, "Error", "Please select a valid file as input."
                    )
                    return
                if operation == "Split by Party Name":
                    msg = split_pdf_by_party_name(
                        input_path, output_dir, party_keyword, party_offset
                    )
                    self.log_message(msg)
                elif operation == "Extract by Order ID":
                    msg = extract_pdf_by_order_id(
                        input_path,
                        output_dir,
                        party_keyword,
                        party_offset,
                        order_keyword,
                        order_regex,
                    )
                    self.log_message(msg)
                elif operation == "Extract Text from PDF":
                    msg, text = process_text_extraction(input_path, output_dir)
                    self.log_message(msg)
                    self.log_message("Extracted Text:")
                    self.log_message(text)
                    self.log_message("-" * 40)
                else:
                    self.log_message("Unknown operation selected.")
            else:  # Folder
                if not os.path.isdir(input_path):
                    QMessageBox.warning(
                        self, "Error", "Please select a valid folder as input."
                    )
                    return
                if operation == "Split by Party Name":
                    messages = process_split_operation(
                        input_path, output_dir, party_keyword, party_offset
                    )
                    for msg in messages:
                        self.log_message(msg)
                elif operation == "Extract by Order ID":
                    messages = process_extract_operation(
                        input_path,
                        output_dir,
                        party_keyword,
                        party_offset,
                        order_keyword,
                        order_regex,
                    )
                    for msg in messages:
                        self.log_message(msg)
                elif operation == "Extract Text from PDF":
                    results = process_text_extraction_operation(input_path, output_dir)
                    for msg, text in results:
                        self.log_message(msg)
                        self.log_message("Extracted Text:")
                        self.log_message(text)
                        self.log_message("-" * 40)
                else:
                    self.log_message("Unknown operation selected.")

            # After processing, compute summary.
            if input_type == "File":
                input_count = 1
            else:
                input_count = len(
                    [f for f in os.listdir(input_path) if f.lower().endswith(".pdf")]
                )
            final_output_files = set(os.listdir(output_dir))
            new_output_files = final_output_files - initial_output_files

            self.log_message("Processing completed.")
            self.log_message(f"Total Number of Files Input: {input_count}")
            self.log_message(f"Total Number of Files Output: {len(new_output_files)}")
            self.log_message("----------------------------------------")
        except Exception as e:
            self.log_message(f"Error during processing: {str(e)}")
            QMessageBox.critical(self, "Processing Error", str(e))


# ----------------------- Main Execution -----------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BusinessToolsUI()
    window.show()
    sys.exit(app.exec_())
