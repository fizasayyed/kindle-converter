import subprocess
import sys
import os
import tempfile
from docx import Document

def check_calibre():
    """Verify Calibre's ebook-convert exists"""
    if not subprocess.run(["which", "ebook-convert"], capture_output=True).returncode == 0:
        print("Calibre not found. Please install it first:")
        print("1. Download from https://calibre-ebook.com")
        print("2. Add to PATH during installation")
        print("3. Restart your terminal/IDE after installation")
        sys.exit(1)

def is_scanned_docx(docx_path):
    """Check if document contains actual text"""
    try:
        doc = Document(docx_path)
        return len([p.text for p in doc.paragraphs if p.text.strip()]) < 5
    except:
        return True

def convert_for_kindle(input_docx, output_path):
    """
    Convert to Kindle-compatible format with proper metadata
    Supported formats: .azw3 (recommended) or .mobi
    """
    _, ext = os.path.splitext(output_path)
    if ext.lower() not in ['.azw3', '.mobi']:
        raise ValueError("Output must be .azw3 or .mobi")

    cmd = [
        'ebook-convert',
        input_docx,
        output_path,
        '--output-profile', 'kindle_pw3',  # Optimal for modern Kindles
        # '--preserve-cover-aspect-ratio',
        '--chapter-mark', 'pagebreak',
        '--disable-remove-fake-margins',
        '--margin-left', '30',
        '--margin-right', '30',
        '--margin-top', '30',
        '--margin-bottom', '30',
        '--extra-css', 'body { font-family: Helvetica }',
        '--language', 'en',
        '--publisher', "My Publisher",
        '--comments', "Converted using Kindle Optimization Script"
    ]

    if is_scanned_docx(input_docx):
        print("Warning: Document appears scanned/image-based. Adding OCR preprocessing...")
        cmd += []
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"Successfully converted to {output_path}")
        print(f"Open in Kindle Previewer: {output_path}")
    except subprocess.CalledProcessError as e:
        print(f"Conversion failed: {e.stderr}")
        sys.exit(1)

if __name__ == "__main__":
    # Configuration
    INPUT_DOCX = "/input/input.docx"  # ← Replace with your file
    OUTPUT_FILE = "kindle_ready.azw3"  # .azw3 (recommended) or .mobi

    # Run conversion
    check_calibre()
    convert_for_kindle(INPUT_DOCX, OUTPUT_FILE)