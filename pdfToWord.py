import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
import tempfile
import os


# Ensure you have the required libraries installed:
# pip install PyMuPDF pytesseract Pillow python-docx

def pdf_to_word_with_ocr(pdf_path, docx_path, dpi=300):
    """
    Convert a PDF to a Word document with OCR for scanned pages.
    
    Args:
        pdf_path (str): Path to the input PDF file
        docx_path (str): Path to save the output Word document
        dpi (int): Resolution for image rendering (default: 300)
    """
    # Configure Tesseract path if necessary (uncomment below line and modify path if needed)
    pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
    
    pdf = fitz.open(pdf_path)
    doc = Document()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            text = page.get_text().strip()
            
            # Add extracted text if available
            if text:
                doc.add_paragraph(text)
            
            # Extract and save embedded images
            image_list = page.get_images(full=True)
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = pdf.extract_image(xref)
                image_data = base_image["image"]
                image_ext = base_image["ext"]
                image_path = os.path.join(temp_dir, f"img_{page_num}_{img_index}.{image_ext}")
                
                with open(image_path, "wb") as f:
                    f.write(image_data)
                
                # Add image to Word document
                try:
                    doc.add_picture(image_path)
                except Exception as e:
                    print(f"Failed to add image: {e}")
            
            # Perform OCR if no text was found
            if not text:
                # Render page as image
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                pix = page.get_pixmap(matrix=mat)
                image_path = os.path.join(temp_dir, f"page_{page_num}.png")
                pix.save(image_path)
                
                # Perform OCR using Tesseract
                try:
                    ocr_text = pytesseract.image_to_string(Image.open(image_path))
                    doc.add_paragraph(ocr_text)
                except Exception as e:
                    print(f"OCR failed for page {page_num + 1}: {e}")
                    doc.add_paragraph(f"[OCR failed for page {page_num + 1}]")
    
    # Save the Word document
    doc.save(docx_path)
    print(f"Successfully converted PDF to Word document: {docx_path}")

# Example usage
if __name__ == "__main__":
    input_pdf = "C:/Users/Dnyanesh/Python/PDF_to_Word_With_OCR/Sample.pdf"
    output_word = "C:/Users/Dnyanesh/Python/PDF_to_Word_With_OCR/output.docx"
    pdf_to_word_with_ocr(input_pdf, output_word)