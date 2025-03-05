import os
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

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
            print(f"Extracted text from {filename}: {text[:100]}...")  # Show first 100 characters as a preview
    return text_data

# Directory path containing PDF files
directory_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/"  # Change this to your directory path
text_data = extract_text_from_directory(directory_path)

# Print extracted text from each PDF file
for filename, text in text_data.items():
    print(f"Text from {filename}:\n{text}\n{'='*40}\n")

# Now text_data is a dictionary where keys are filenames and values are the extracted text