# PDF Crash Report to Excel Extractor

## Overview

This tool extracts structured crash report information from PDF files and converts the extracted data into a clean and formatted Excel sheet.

It is designed for accident and crash investigation reports where the information is stored in semi-structured PDF documents. The extraction process uses text parsing and regular expressions to identify important fields and organize them into tabular format.

The generated Excel file can be used for:
- Data analysis
- Dashboard creation
- Reporting
- Research work
- Machine learning pipelines
- Large scale accident data digitization

---

## Features

- Batch processing of multiple PDFs
- Structured Excel output
- Automatic schema validation
- Duplicate accident detection
- Missing value normalization
- Clean and formatted Excel generation
- Automatic column width adjustment
- Freeze header row support
- Error handling during processing

---

## Technologies Used

- Python
- pdfplumber
- pandas
- openpyxl
- regular expressions (regex)

---

## Project Structure

```text
Tasks/
│
├── TASK_A1/
│   │
│   ├── output/
│   │   └── crash_reports.xlsx
│   │
│   ├── pdfs/
│   │   └── report1.pdf
│   │
│   ├── pdf_to_excel.py
│   ├── requirements.txt
│   └── README.md
│
├── TASK_B2/
```

---

## Installation

Install Python 3.10 or later.

Install required libraries using:

```bash
pip install -r requirements.txt
```

---

## Required Libraries

The following libraries are required:

```text
pdfplumber
pandas
openpyxl
numpy
```

---

## How to Run the Tool

Place all crash report PDFs inside the `pdfs` folder.

Run the script using:

```bash
python pdf_to_excel.py
```

The tool will:
1. Read all PDF files from the `pdfs` folder
2. Extract structured fields using regex patterns
3. Clean and normalize values
4. Remove duplicate accident records
5. Generate a formatted Excel sheet

---

## Output

The generated Excel file will be saved inside:

```text
output/crash_reports.xlsx
```

The output contains:
- Structured accident information
- Vehicle details
- Driver details
- Road details
- Casualty information
- Mechanical inspection data
- Safety system information

---

## Missing Value Handling

The following values are automatically converted into empty values during processing:

- NA
- N/A
- null
- None
- blank strings

This improves consistency and makes downstream analysis easier.

---

## Duplicate Detection

Duplicate accident records are automatically removed using:
- Accident ID

Only the first occurrence is retained.

---

## Excel Formatting

The generated Excel sheet includes:
- Light grey header row
- Black header text
- Borders for all cells
- Auto-sized columns
- Frozen header row

The formatting is intentionally kept simple and professional for readability.

---

## Notes

This tool currently works best on digitally generated PDFs where text can be selected normally.

For scanned image PDFs, OCR support can be added later if required.

---

## Limitations

- Extraction accuracy depends on PDF structure consistency
- Different report templates may require regex adjustments
- Multi-vehicle and multi-person records are currently flattened into single rows

---

## Future Improvements

Possible future upgrades include:
- OCR support for scanned PDFs
- Streamlit web interface
- Database integration
- JSON export
- Confidence scoring
- Multi-row vehicle extraction
- AI-assisted correction pipeline

