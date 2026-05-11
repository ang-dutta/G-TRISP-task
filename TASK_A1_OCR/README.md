# PDF Crash Report to Excel Converter -- OCR Pipeline

## Overview

This tool converts crash report PDFs into a structured Excel file using Optical Character Recognition (OCR). Unlike the `TASK_A1` version which relies on embedded text inside digital PDFs, this pipeline renders every PDF page into a high-resolution image and applies Tesseract OCR to read the content. This makes it capable of processing both digitally generated and scanned PDFs.

---

## How It Differs from TASK_A1

| Aspect | TASK_A1 (pdfplumber) | TASK_A1_OCR (Tesseract) |
|---|---|---|
| Text source | Embedded PDF text | Rendered page image |
| Works on scanned PDFs | No | Yes |
| Requires Tesseract | No | Yes |
| Image preprocessing | No | Yes (denoise, binarise, deskew) |
| Speed | Fast | Slower (image rendering + OCR) |
| Accuracy on digital PDFs | Higher | Slightly lower (OCR errors possible) |
| Accuracy on scanned PDFs | Fails | Works |

---

## System Requirements

### 1. Install Tesseract OCR

Tesseract is an external binary that must be installed on your machine before running this tool.

**Windows:**

Download the installer from the official repository:

```
https://github.com/UB-Mannheim/tesseract/wiki
```

Run the installer. The default installation path is:

```
C:\Program Files\Tesseract-OCR\tesseract.exe
```

During installation, make sure to select "English" language data (selected by default).

After installation, either:
- Add `C:\Program Files\Tesseract-OCR` to your system `PATH`, or
- Open `ocr_to_excel.py` and set `TESSERACT_CMD` at the top:

```python
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

**Linux:**

```bash
sudo apt-get install tesseract-ocr
```

**macOS:**

```bash
brew install tesseract
```

---

## Python Installation

Install Python 3.10 or later, then install dependencies:

```bash
pip install -r requirements.txt
```

No additional system library (such as Poppler) is required. The PDF-to-image rendering is handled by PyMuPDF which is a self-contained Python package.

---

## Project Structure

```text
TASK_A1_OCR/
|-- ocr_to_excel.py       Main OCR pipeline script
|-- requirements.txt      Python dependencies
|-- README.md             This file
|-- pdfs/                 Place input PDF crash reports here
|   +-- report1.pdf
+-- output/               Generated Excel output (created automatically)
    +-- crash_reports_ocr.xlsx
```

---

## How to Run

1. Place all crash report PDF files inside the `pdfs/` folder.

2. Open `ocr_to_excel.py` and set `TESSERACT_CMD` if Tesseract is not on your PATH.

3. Run the script:

```bash
python ocr_to_excel.py
```

The tool will process all PDFs found in `pdfs/`, extract structured fields, and save the result to `output/crash_reports_ocr.xlsx`.

---

## OCR Pipeline Steps

For each PDF file, the tool performs the following steps:

1. **Rasterisation** -- Each page is rendered to a high-resolution image at 300 DPI using PyMuPDF (no Poppler needed).
2. **Grayscale conversion** -- Colour images are converted to grayscale to reduce noise channels.
3. **Gaussian blur** -- A 3x3 blur reduces scanner grain and JPEG compression artefacts.
4. **Otsu binarisation** -- The image is converted to pure black and white using Otsu's automatic threshold selection. This separates text from background reliably.
5. **Deskew** -- The minimum-area bounding rectangle of text pixels is computed to detect and correct page tilt (up to approximately 15 degrees).
6. **Tesseract OCR** -- The cleaned binary image is passed to Tesseract with PSM 6 (uniform text block) and the LSTM neural network engine.
7. **Field extraction** -- The same regular expression patterns used in TASK_A1 are applied to the OCR output text.
8. **Excel export** -- All records are compiled into one DataFrame and saved as a formatted Excel file.

---

## Output Format

The output file `output/crash_reports_ocr.xlsx` contains:
- One row per PDF crash report.
- One column per extracted field.
- `source_pdf` as the first column, showing which file each row came from.

### Extracted Field Categories

- Case and FIR information
- Accident details (date, time, location, severity, collision type)
- Casualty counts (driver, passenger, pedestrian -- killed, injured)
- Vehicle details (registration, type, class, make, insurance)
- Driver details (name, age, license, seatbelt, impairment)
- Road characteristics (classification, surface, junction, speed limit)
- Vehicle safety and mechanical condition

---

## Tuning OCR Accuracy

Two parameters in `ocr_to_excel.py` control OCR behaviour:

| Parameter | Default | Effect |
|---|---|---|
| `RENDER_DPI` | 300 | Increase to 400 for dense or small-font PDFs. Higher DPI improves accuracy but increases processing time. |
| `PSM_MODE` | 6 | Try PSM 3 (fully automatic) if the page layout is complex or multi-column. |

---

## Missing Value Handling

The following values extracted by OCR are automatically normalised to empty cells in the Excel output:

- `NA`, `N/A`, `na`, `null`, `None`, blank strings, `-`

---

## Limitations

- OCR accuracy is dependent on PDF scan quality. Blurry, skewed, or low-resolution scans will produce more extraction errors.
- Multi-vehicle crash records are flattened to a single row (only the first vehicle's details are captured per report).
- Handwritten content in PDFs cannot be read by Tesseract without specialist handwriting OCR models.
