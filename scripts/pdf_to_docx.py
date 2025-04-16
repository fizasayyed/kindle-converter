from pdf2docx import Converter
import time

def pdf_to_docx(input_path, output_path):
    start_time = time.time()

    # Convert PDF to DOCX
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None, multi_processing=True)
    cv.close()

    end_time = time.time()
    print(f"Conversion completed in {end_time - start_time:.2f} seconds")

if __name__ == "__main__":
    input_pdf = "/input/input.pdf"  # Replace with your PDF path
    output_docx = "output.docx"  # Replace with desired output path

    pdf_to_docx(input_pdf, output_docx)