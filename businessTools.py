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

import pymupdf
from PyQt6.QtWidgets import (
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


def extract_order_id_dynamic(text, order_keyword="ID :", regex_pattern=r"ID\s*:\s*(?:\n\s*Date\s*:[^\n]*)?\s*\n\s*([A-Za-z0-9][A-Za-z0-9._/-]*)"):
    """
    Extracts an order ID using the configured keyword and regex.

    Some source PDFs place the identifier on the line after the ID label and
    put a Date line between the label and identifier, for example::

        ID :
        Date : 02/10/2026
        R/349

    The supplied regex is therefore matched against the complete page text,
    including newlines. A small line-based fallback is retained for custom
    profiles whose regex targets only a same-line or nearby value.
    """
    try:
        match = re.search(regex_pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        if match:
            value = match.group(1).strip()
            if value:
                return value
    except (re.error, IndexError):
        # A user-configured regex may be invalid or may not contain a capture
        # group. Fall back to the line-based logic below rather than breaking
        # the whole PDF operation.
        pass

    lines = text.splitlines()
    for i, line in enumerate(lines):
        if order_keyword.lower() not in line.lower():
            continue

        # Try the configured regex against the keyword line first.
        try:
            match = re.search(regex_pattern, line, flags=re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                if value:
                    return value
        except (re.error, IndexError):
            pass

        # Then inspect nearby non-empty lines. Skip common date labels so a
        # date value is never mistaken for the order ID.
        for j in range(i + 1, min(i + 8, len(lines))):
            candidate = lines[j].strip()
            if not candidate or candidate == ":":
                continue
            if re.match(r"^Date\s*:", candidate, flags=re.IGNORECASE):
                continue

            # Prefer a complete order-ID token such as R/349, R-349 or 349.
            token = re.search(r"\b[A-Za-z]+\s*[/_-]\s*\d+\b|\b\d+\b", candidate)
            if token:
                return re.sub(r"\s+", "", token.group(0))

    return None


def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF, falling back to OCR for image-only pages.

    Text is accumulated in a list rather than repeatedly concatenating strings,
    which avoids unnecessary intermediate string allocations on larger PDFs.
    """
    text_chunks = []
    try:
        with pymupdf.open(pdf_path) as doc:
            for page in doc:
                page_text = page.get_text("text")
                if page_text.strip():
                    text_chunks.append(page_text)
                    continue

                # Fallback to OCR if no text is embedded on the page.
                pix = page.get_pixmap()
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                text_chunks.append(pytesseract.image_to_string(img))
    except Exception as e:
        return f"Error processing {pdf_path}: {e}"

    return "".join(text_chunks)


def sanitize_filename(filename):
    """
    Removes invalid characters for Windows filenames, including square brackets.
    """
    # The regex removes: < > : " / \ | ? * [ ]
    return re.sub(r'[<>:"/\\|?*\[\]]+', "", filename)


# ----------------------- PDF Processing Functions -----------------------


def _group_page_numbers_by_key(doc, party_keyword, party_offset, order_keyword=None, order_regex=None):
    """Collect pages by their extracted party or party/order key.

    Pages are grouped globally rather than only when the same key appears on
    consecutive pages. This prevents a repeated party later in the document
    from overwriting an earlier output file.
    """
    groups = {}
    last_key = None

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        party = extract_party_name_dynamic(text, party_keyword, party_offset)

        if order_keyword is None:
            key = party
        else:
            order_id = extract_order_id_dynamic(text, order_keyword, order_regex)
            key = (party, order_id) if party and order_id else None

        if key is not None:
            groups.setdefault(key, []).append(page_num)
            last_key = key
        elif last_key is not None:
            # Preserve the existing behavior for continuation pages that do
            # not repeat the identifying text.
            groups.setdefault(last_key, []).append(page_num)

    return groups


def _insert_page_numbers(dest_doc, source_doc, page_numbers):
    """Copy possibly non-contiguous source pages using contiguous ranges."""
    if not page_numbers:
        return

    start = previous = page_numbers[0]
    for page_num in page_numbers[1:]:
        if page_num == previous + 1:
            previous = page_num
            continue

        dest_doc.insert_pdf(source_doc, from_page=start, to_page=previous)
        start = previous = page_num

    dest_doc.insert_pdf(source_doc, from_page=start, to_page=previous)


def _unique_output_path(output_directory, filename):
    """Return a collision-free output path without overwriting another group."""
    base, extension = os.path.splitext(filename)
    candidate = os.path.join(output_directory, filename)
    counter = 2
    while os.path.exists(candidate):
        candidate = os.path.join(output_directory, f"{base} ({counter}){extension}")
        counter += 1
    return candidate


def split_pdf_by_party_name(pdf_path, output_directory, party_keyword, party_offset):
    """
    Splits the input PDF by party name.

    All pages belonging to the same party are collected into one output PDF,
    even when that party appears again later in the source document.
    """
    with pymupdf.open(pdf_path) as doc:
        groups = _group_page_numbers_by_key(
            doc, party_keyword, party_offset
        )

        for party, page_numbers in groups.items():
            if not party:
                continue
            new_pdf = pymupdf.open()
            _insert_page_numbers(new_pdf, doc, page_numbers)
            sanitized_party = sanitize_filename(party).strip()
            output_filename = f"{sanitized_party}.pdf"
            output_path = _unique_output_path(output_directory, output_filename)
            new_pdf.save(output_path)
            new_pdf.close()

    return f"Processed (Split by Party Name): {os.path.basename(pdf_path)}"


def extract_pdf_by_order_id(
    pdf_path, output_directory, party_keyword, party_offset, order_keyword, order_regex
):
    """
    Splits the PDF into PDFs grouped by party name + order ID.

    Repeated occurrences of the same party/order ID anywhere in the source are
    combined into one output file, while different order IDs remain separate.
    """
    with pymupdf.open(pdf_path) as doc:
        groups = _group_page_numbers_by_key(
            doc,
            party_keyword,
            party_offset,
            order_keyword,
            order_regex,
        )

        for key, page_numbers in groups.items():
            if not isinstance(key, tuple) or len(key) != 2:
                continue
            party, order_id = key
            if not party or not order_id:
                continue

            new_pdf = pymupdf.open()
            _insert_page_numbers(new_pdf, doc, page_numbers)
            combined_name = sanitize_filename(f"{party}_{order_id}")
            output_filename = f"{combined_name}.pdf"
            output_path = _unique_output_path(output_directory, output_filename)
            new_pdf.save(output_path)
            new_pdf.close()

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
        self.default_order_regex = r"ID\s*:\s*(?:\n\s*Date\s*:[^\n]*)?\s*\n\s*([A-Za-z0-9][A-Za-z0-9._/-]*)"
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
    sys.exit(app.exec())
