import os
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import re

def extract_text_from_pdf(pdf_path):
    try:
        text = ""
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text += page.get_text("text")  # Try to extract text directly
        return text
    except Exception as e:
        print(f"Error extracting text from {pdf_path}: {e}")
        return ""

def extract_text_from_image_pdf(pdf_path):
    try:
        text = ""
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.open(io.BytesIO(pix.tobytes()))
            text += pytesseract.image_to_string(img)  # Extract text using OCR
        return text
    except Exception as e:
        print(f"Error extracting text with OCR from {pdf_path}: {e}")
        return ""

def sanitize_name(filename):
    # Remove invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized_filename = re.sub(invalid_chars, '', filename)
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

def rename_pdf_file(old_path, new_name):
    directory = os.path.dirname(old_path)
    new_path = os.path.join(directory, new_name + ".pdf")
    os.rename(old_path, new_path)
    return new_path

def extract_text_from_directory(directory):
    text_data = {}
    for filename in os.listdir(directory):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(directory, filename)
            print(f"Processing file: {file_path}")
            text = extract_text_from_pdf(file_path)
            if not text.strip():  # If text is empty, use OCR
                print(f"Direct text extraction failed for {file_path}, using OCR...")
                text = extract_text_from_image_pdf(file_path)
            
            text_data[filename] = text
            party_name_and_order_id = extract_party_name_and_order_id(text)
            if party_name_and_order_id:
                new_file_path = rename_pdf_file(file_path, party_name_and_order_id)
                print(f"Renamed file to: {new_file_path}")
            else:
                print(f"Party name or order ID not found in {file_path}")
    return text_data

# Directory path containing PDF files
directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/"
text_data = extract_text_from_directory(directory_path)

# Print extracted text from each PDF file
for filename, text in text_data.items():
    print(f"Text from {filename}:\n{text}\n{'='*40}\n")

# Now text_data is a dictionary where keys are filenames and values are the extracted text
