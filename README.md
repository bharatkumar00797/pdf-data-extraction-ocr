# PDF Data Extraction & OCR Comparison Tool

Python-based project to extract data from scanned PDF annual reports using OCR, convert content into Word and Excel formats, and verify data accuracy by comparing multiple scanned PDFs against a compiled Excel-to-PDF file.

## Features

- Extract text from normal (text-based) PDFs
- Perform OCR on scanned / image-based PDFs
- Convert extracted data into Word (`.docx`) and Excel (`.xlsx`)
- Compare one Excel-converted PDF against 15 original scanned annual reports
- Generate comparison reports showing matched and unmatched data points

## Tech Stack

- Python 3.12
- `pdfplumber` – text & table extraction
- `pdf2image` + Tesseract OCR – scanned document processing
- `python-docx` – Word document generation
- `openpyxl` – Excel file generation
- `pytesseract` – OCR interface

## Project Structure

| File | Description |
|------|-------------|
| `extract_simple.py` | Basic text extraction from PDFs into Word + Excel |
| `extract_ocr.py` | OCR extraction from scanned PDFs |
| `compare_excel_vs_15pdfs.py` | Full comparison between Excel PDF and 15 scanned reports |
| `requirements.txt` | Python dependencies |

## Setup

1. Create and activate a virtual environment:
```bash
python -m venv myenv
myenv\Scripts\activate          # Windows
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Install system dependencies:
   - **Tesseract OCR**: Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
   - **Poppler**: Required by `pdf2image` (Windows builds available on GitHub)

4. Update the file paths inside the scripts (PDF folders, Tesseract path, Poppler path).

## Usage

```bash
# Simple text extraction
python extract_simple.py

# OCR on scanned PDFs
python extract_ocr.py

# Full comparison against 15 reports
python compare_excel_vs_15pdfs.py
```

## Purpose

This project was built while working with real hospital annual reports.  
The goal was to extract structured data from scanned documents and verify the accuracy of compiled datasets.

## Notes

- OCR accuracy depends on scan quality. Results are approximate.
- Processing 15 multi-page scanned PDFs can take significant time on the first run.
- Intermediate OCR text files are saved so subsequent runs are much faster.

## Future Improvements

- Better table detection from scanned pages
- Improved OCR preprocessing (deskew, contrast, etc.)
- Automated data cleaning and structuring
- Support for multiple languages
