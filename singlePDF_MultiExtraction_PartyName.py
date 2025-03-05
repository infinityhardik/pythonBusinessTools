import os
import fitz  # PyMuPDF
import re
from collections import defaultdict

def sanitize_filename(filename):
    # Remove invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized_filename = re.sub(invalid_chars, '', filename)
    return sanitized_filename

def extract_party_name(text):
    # Extract lines from the text
    lines = text.split('\n')
    
    party_name = None

    # Loop through lines to find the line above "M/s." for the party name
    for i in range(1, len(lines)):
        if "M/s." in lines[i]:
            party_name = lines[i - 1].strip()  # Trim the party name
            break

    return sanitize_filename(party_name).strip() if party_name else None

def split_pdf_by_party_name(pdf_path):
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    
    current_party_name = None
    current_start_page = 0
    splits = defaultdict(list)

    # Extract text and determine split points
    for page_num in range(num_pages):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        
        new_party_name = extract_party_name(text)
        if new_party_name:
            if current_party_name:
                splits[current_party_name].append((current_start_page, page_num - 1))
            current_party_name = new_party_name
            current_start_page = page_num

    # Add the last segment
    if current_party_name:
        splits[current_party_name].append((current_start_page, num_pages - 1))

    return splits, doc

def save_split_pdfs(splits, doc, output_directory):
    for party_name, ranges in splits.items():
        new_pdf = fitz.open()
        for start_page, end_page in ranges:
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
            splits, doc = split_pdf_by_party_name(file_path)
            save_split_pdfs(splits, doc, output_directory)

# Directory containing the PDF files to be processed
pdf_directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/"
# Directory where the split and renamed PDF files will be saved
output_directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/Output"

# Ensure the output directory exists
os.makedirs(output_directory_path, exist_ok=True)

# Process all PDF files in the specified directory
process_pdf_directory(pdf_directory_path, output_directory_path)
