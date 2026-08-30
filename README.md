# PDF Data Extraction & OCR Tools

A small collection of Python scripts I built to work with scanned PDF reports.

I needed to pull data out of a bunch of old annual reports (most of them were just scanned images) and then check how accurate a compiled Excel version was against the originals. These scripts helped me do that.

---

## What it does

- Extract text from normal PDFs
- Run OCR on scanned/image-based PDFs
- Save the extracted text into Word and Excel files
- Compare numbers from one clean PDF against text extracted from 15 scanned reports

---

## Files

| Script | What it does |
|--------|--------------|
| `extract_simple.py` | Basic text extraction (works only if the PDF has selectable text) |
| `extract_ocr.py` | Converts PDF pages to images and runs OCR |
| `compare_excel_vs_15pdfs.py` | OCRs all 15 reports and checks which numbers from the Excel PDF appear in them |

---

## Requirements

```bash
pip install -r requirements.txt
```

You also need two things installed on your system:

1. **Tesseract OCR**  
   Download the Windows installer from here:  
   https://github.com/UB-Mannheim/tesseract/wiki

2. **Poppler** (needed by pdf2image)  
   Get the Windows build and extract it somewhere simple like `C:\poppler`

After installing, open the scripts and update these two lines:

```python
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH   = r"C:\poppler\Library\bin"
```

Also change the folder paths at the top of each script to point to your own PDF files.

---

## How to run

```bash
# Simple extraction (fast)
python extract_simple.py

# OCR version (much slower)
python extract_ocr.py

# Full comparison
python compare_excel_vs_15pdfs.py
```

The comparison script saves OCR results locally, so the second time you run it will be a lot faster.

---

## Notes

- OCR is never perfect. Some numbers will be misread, especially if the scan quality is bad.
- Running OCR on 15 full reports takes time. Be patient on the first run.
- These scripts were written for a specific set of hospital annual reports, so you will need to change the file paths.

---

## Possible improvements

- Better image preprocessing before OCR
- Detecting tables more reliably
- Cleaning up the extracted numbers automatically
- Making the paths configurable through a config file instead of hardcoding them
