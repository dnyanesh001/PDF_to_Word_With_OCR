from pdf2docx import Converter

def convert_pdf_to_docx(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)  # full document
    cv.close()

# Example usage
convert_pdf_to_docx("C:/Users/Dnyanesh/Python/PDF_to_Word_With_OCR/Sample.pdf", "C:/Users/Dnyanesh/Python/PDF_to_Word_With_OCR/output.docx")
