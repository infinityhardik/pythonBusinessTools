# Description: This script is the main script that will be run to start the GUI for the Business Tools.
# The script will create a simple GUI using PyQt5 that will allow the user to select an input folder, an output folder, and a processing type.
# The processing types will include "Split by Party Name" and "Extract by Order ID".
# When the user clicks the "Start Processing" button, the selected processing type will be executed using the input and output folders selected by the user.
# The processing will be done by running the appropriate Python script using the os.system() function.
# The GUI will display a status message indicating the progress of the processing.
# The user can select the input and output folders using the "Browse Input Folder" and "Browse Output Folder" buttons, respectively.
# The user can select the processing type using a drop-down list.

import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QComboBox

class PDFProcessorUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("PDF Processing Tool")
        self.setGeometry(100, 100, 500, 300)
        
        layout = QVBoxLayout()
        
        self.label_input = QLabel("Select Input Directory:")
        layout.addWidget(self.label_input)
        
        self.btn_input = QPushButton("Browse Input Folder")
        self.btn_input.clicked.connect(self.select_input_folder)
        layout.addWidget(self.btn_input)
        
        self.label_output = QLabel("Select Output Directory:")
        layout.addWidget(self.label_output)
        
        self.btn_output = QPushButton("Browse Output Folder")
        self.btn_output.clicked.connect(self.select_output_folder)
        layout.addWidget(self.btn_output)
        
        self.label_service = QLabel("Select Processing Type:")
        layout.addWidget(self.label_service)
        
        self.combo_service = QComboBox()
        self.combo_service.addItems(["Split by Party Name", "Extract by Order ID"])
        layout.addWidget(self.combo_service)
        
        self.btn_start = QPushButton("Start Processing")
        self.btn_start.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_start)
        
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
    
    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.label_input.setText(f"Input: {folder}")
            self.input_folder = folder
    
    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.label_output.setText(f"Output: {folder}")
            self.output_folder = folder
    
    def start_processing(self):
        if not hasattr(self, 'input_folder') or not hasattr(self, 'output_folder'):
            self.status_label.setText("Please select input and output folders.")
            return
        
        service = self.combo_service.currentText()
        self.status_label.setText(f"Processing PDFs using: {service}...")
        
        if service == "Split by Party Name":
            self.run_split_by_party()
        elif service == "Extract by Order ID":
            self.run_extract_by_order_id()
        
    def run_split_by_party(self):
        os.system(f'python singlePDF_MultiExtraction_PartyName.py "{self.input_folder}" "{self.output_folder}"')
        self.status_label.setText("Splitting completed!")
    
    def run_extract_by_order_id(self):
        os.system(f'python PDFsIndividualFileExtraction_PartyName_OrderID.py "{self.input_folder}" "{self.output_folder}"')
        self.status_label.setText("Extraction completed!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFProcessorUI()
    window.show()
    sys.exit(app.exec_())