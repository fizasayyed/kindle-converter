# Book Conversion Scripts

## Overview
I was trying to obtain the book "Communing with the Divine" in EPUB, MOBI, or AZW format for Kindle but was unsuccessful in finding them. To solve this, I have written several scripts that facilitate the conversion of PDF or DOCX files into the required formats.

## Scripts

| Script Name                          | Description                                                                 | How to Run                          | Input/Output                                                                 |
|---------------------------------------|-----------------------------------------------------------------------------|-------------------------------------|------------------------------------------------------------------------------|
| `pdf_to_docx.py`                     | Converts a PDF file to a DOCX format.                                     | ```bash python pdf_to_docx.py ```  | Input: Specify the input PDF file path in the script. <br> Output: Generates a DOCX file. |
| `docx_to_mobi.py`                    | Converts a DOCX file to MOBI format using Calibre.                        | ```bash python docx_to_mobi.py ``` | Input: Specify the input DOCX file path in the script. <br> Output: Generates a MOBI file. |
| `docx_to_epub.py`                    | Converts a DOCX file to EPUB format using Calibre.                        | ```bash python docx_to_epub.py ``` | Input: Specify the input DOCX file path in the script. <br> Output: Generates an EPUB file. |
| `docx_image_text_extractor.py`       | Extracts text from images in a DOCX file using OCR and creates a new DOCX file with the extracted text. | ```bash python docx_image_text_extractor.py ``` | Input: Specify the input DOCX file path in the script. <br> Output: Generates a new DOCX file with text extracted from images. |

## Best Shot Approach
1. **Convert PDF to DOCX**: Use `pdf_to_docx.py` to convert your PDF file to DOCX.
2. **Extract Text from DOCX**: If your DOCX contains images, use `docx_image_text_extractor.py` to extract text from those images.
3. **Preview in Kindle Previewer 3**: Open the generated DOCX file in Kindle Previewer 3 to check the formatting and layout.
4. **Export as Needed Format**: Use Kindle Previewer 3 to export the file in your desired format (MOBI, EPUB, etc.).

## Truly for Learning Purpose

This repository is intended for educational purposes only. Users are encouraged to learn from the code and modify it for personal use. However, copying, redistributing, or using this work for commercial purposes without permission is strictly prohibited and may result in legal action.

Please respect the original work and its intent.

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## Usage Instructions
- Ensure you have Calibre installed and added to your PATH for the conversion scripts that require it.
