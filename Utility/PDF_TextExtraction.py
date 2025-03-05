import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

def extract_text_from_pdf(pdf_path):
    text = ""
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text("text")  # Try to extract text directly
    return text

def extract_text_from_image_pdf(pdf_path):
    text = ""
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes()))
        text += pytesseract.image_to_string(img)  # Extract text using OCR
    return text

# Path to the uploaded PDF file
pdf_path = "C:/Users/Hardik Bhaavani/Desktop/Python Project/files/1.PDF"

# Extract text directly from the PDF
text = extract_text_from_pdf(pdf_path)
if not text.strip():  # If text is empty, use OCR
    text = extract_text_from_image_pdf(pdf_path)

print(text)
