# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

- Planned: Image preprocessing before OCR
- Planned: Configurable paths via external settings file
- Planned: Per-document accuracy breakdown

## [2026-08-31]

### Added
- `data_difference_checker.py` — Compares character counts, word counts, and unique numbers between original and processed files
- `full_15pdf_accuracy_checker.py` — Calculates match rate and accuracy starting from the original 15 PDF files
- Accuracy reporting in both text and Excel formats

### Changed
- Improved README documentation
- Clarified accuracy calculation method

## [2026-08-30]

### Added
- Initial project structure
- `extract_simple.py` — Basic text extraction from PDFs
- `extract_ocr.py` — OCR extraction for scanned PDFs
- `compare_excel_vs_15pdfs.py` — Comparison between compiled file and multiple source reports
- Project README and requirements file
