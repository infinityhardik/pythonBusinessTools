# Individual File Extraction based on Order ID and Party Name
import os
import fitz  # PyMuPDF
import re

def sanitize_name(filename):
    # Remove invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized_filename = re.sub(invalid_chars, '_', filename)
    return sanitized_filename

def extract_party_name_and_order_id(text):
    # Extract lines from the text
    lines = text.split('\n')
    
    party_name = None
    order_id = None

    # Loop through lines to find the line above "M/s." for the party name
    for i in range(1, len(lines)):
        if "M/s" in lines[i]:
            party_name = lines[i - 1].strip()
        if "ID :" in lines[i]:
            order_id_match = re.search(r"ID : (\d+)", lines[i])
            if order_id_match:
                order_id = order_id_match.group(1)

    if party_name and order_id:
        # Sanitize Party Name & Strip it
        party_name = sanitize_name(party_name) 
        party_name = party_name.strip()
        raw_filename = f"{party_name}_{order_id}"
        sanitized_filename = sanitize_name(raw_filename)
        return sanitized_filename
    return None

def demerge_and_rename_pdf(pdf_path, output_directory):
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    
    current_party_name = None
    current_start_page = 0
    splits = []

    # Extract text and determine split points
    for page_num in range(num_pages):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        
        new_party_name = extract_party_name_and_order_id(text)
        if new_party_name:
            if current_party_name:
                splits.append((current_start_page, page_num - 1, current_party_name))
            current_party_name = new_party_name
            current_start_page = page_num

    # Add the last segment
    if current_party_name:
        splits.append((current_start_page, num_pages - 1, current_party_name))

    # Split and rename the files
    for start_page, end_page, party_name in splits:
        new_pdf = fitz.open()
        for i in range(start_page, end_page + 1):
            new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        
        output_filename = f"{party_name}.pdf"
        output_path = os.path.join(output_directory, output_filename)
        new_pdf.save(output_path)
        new_pdf.close()
        print(f"Created: {output_path}")

def process_pdf_directory(directory_path, output_directory):
    for filename in os.listdir(directory_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(directory_path, filename)
            print(f"Processing file: {file_path}")
            demerge_and_rename_pdf(file_path, output_directory)

# Directory containing the PDF files to be processed
pdf_directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/"

# Directory where the split and renamed PDF files will be saved
output_directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/Output"

# Ensure the output directory exists
os.makedirs(output_directory_path, exist_ok=True)

# Process all PDF files in the specified directory
process_pdf_directory(pdf_directory_path, output_directory_path)
