import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.shared import Inches
import pytesseract
import os
import tempfile

def pdf_to_word_image_based(pdf_path, docx_path, dpi=300):
    pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
    
    pdf = fitz.open(pdf_path)
    doc = Document()

    with tempfile.TemporaryDirectory() as temp_dir:
        for i in range(len(pdf)):
            page = pdf[i]
            # Render PDF page to image
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            image_path = os.path.join(temp_dir, f"page_{i}.png")
            pix.save(image_path)
            
            # Add full-page image to Word
            doc.add_picture(image_path, width=Inches(6.5))  # Adjust width for A4 layout
            
            # (Optional) OCR text as separate paragraph
            ocr_text = pytesseract.image_to_string(Image.open(image_path))
            # doc.add_paragraph(ocr_text)  # Uncomment to include OCR text as selectable text
            
            # Add page break
            doc.add_page_break()

    doc.save(docx_path)
    print(f"Word file saved to: {docx_path}")
    # Example usage
    pdf_to_word_image_based("C:/Users/Dnyanesh/Python/PDF_to_Word_With_OCR/Sample.pdf", "C:\Users\Dnyanesh\Python\PDF_to_Word_With_OCR\output.docx")
