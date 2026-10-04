# pythonBusinessTools

A Windows desktop utility for splitting PDFs by party/order information and extracting PDF text.

## Python version

Use **Python 3.14.x** (64-bit Windows is recommended).

## Install Python dependencies

Open a terminal in this project folder and run:

```powershell
py -3.14 -m pip install --upgrade pip
py -3.14 -m pip install -r requirements.txt
```

### Important: do not install `fitz`

The project uses **PyMuPDF**, whose supported import is now `pymupdf`.
The separate PyPI package named `fitz` is obsolete and unrelated to PyMuPDF. Installing it can break PyMuPDF imports.

If `fitz` was previously installed, remove it and reinstall PyMuPDF:

```powershell
py -3.14 -m pip uninstall -y fitz
py -3.14 -m pip install --upgrade --force-reinstall PyMuPDF
```

## OCR support

Direct PDF text extraction works without Tesseract. OCR is used as a fallback for image-only PDFs.

To use OCR, install the **Tesseract OCR application** separately and make sure `tesseract.exe` is available on `PATH`.

`pytesseract` is only the Python wrapper; it does not contain the Tesseract OCR engine.

## Run the application

```powershell
py -3.14 businessTools.py
```

Or use `PDFTool.bat` from the project folder.

## Main operations

- **Split by Party Name** – groups consecutive pages by the detected party name.
- **Extract by Order ID** – groups consecutive pages by party name + order ID.
- **Extract Text from PDF** – extracts text and falls back to OCR when a page has no direct text.
