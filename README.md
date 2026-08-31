# PDF Data Extraction & Accuracy Analysis

This repository contains a set of Python tools developed for extracting data from scanned PDF documents and measuring the accuracy of compiled datasets against original source files.

The project was created while working with a collection of hospital annual reports. Most of the source files were image-based scans, which required OCR processing before any meaningful comparison could be performed.

---

## Overview

The tools support the following workflow:

1. Extract text from standard (text-based) PDFs
2. Perform OCR on scanned PDF pages
3. Convert extracted content into Word and Excel formats
4. Compare a compiled Excel/PDF dataset against the original 15 source reports
5. Calculate match rate and overall data accuracy
6. Generate detailed difference reports

---

## Scripts

| Script | Description |
|--------|-------------|
| `extract_simple.py` | Extracts text from PDFs that contain selectable text and exports results to Word and Excel |
| `extract_ocr.py` | Converts scanned PDF pages into images and applies OCR |
| `compare_excel_vs_15pdfs.py` | Performs OCR on multiple reports and compares extracted numbers against a compiled file |
| `data_difference_checker.py` | Compares character counts, word counts, and unique numbers between an original text source and a processed file |
| `full_15pdf_accuracy_checker.py` | Starts from the original 15 PDF files, extracts all unique numbers, and calculates match rate against the target Excel/PDF |

---

## Requirements

```bash
pip install -r requirements.txt
```

System dependencies:

- **Tesseract OCR** — [Windows installer](https://github.com/UB-Mannheim/tesseract/wiki)
- **Poppler** — required by `pdf2image` for converting PDF pages to images

After installation, update the following paths inside the relevant scripts:

```python
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH   = r"C:\poppler\Library\bin"
```

Also update the input folder and file paths to match your local directory structure.

---

## Usage

```bash
# Basic text extraction
python extract_simple.py

# OCR extraction from scanned PDFs
python extract_ocr.py

# Compare compiled file against previously extracted text
python data_difference_checker.py

# Full accuracy check starting from the original 15 PDFs
python full_15pdf_accuracy_checker.py
```

---

## Accuracy Calculation

The primary accuracy metric is based on unique numerical values:

```
Accuracy = (Matched Numbers / Total Unique Numbers in Source) × 100
```

Reports include:

- Total unique numbers found in the source files
- Numbers successfully matched in the target file
- Numbers missing from the target file
- Numbers present only in the target file
- Overall match rate / accuracy percentage

---

## Notes

- OCR quality depends heavily on the resolution and clarity of the original scans.
- The first run of any OCR-based script can take a significant amount of time.
- Intermediate OCR text files are saved locally so subsequent runs are much faster.
- These scripts were written for a specific set of annual reports and will require path updates for use with other documents.

---

## Future Improvements

- Image preprocessing (deskew, contrast enhancement) before OCR
- More robust table structure detection
- Configurable paths through an external settings file
- Per-document accuracy breakdown
