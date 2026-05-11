# G-TRISP Project Tasks

This directory contains the implementations for the G-TRISP project assignments. Each task is self-contained within its own folder and includes its own detailed `README.md` with setup instructions, pipeline explanations, and execution steps.

## Directory Structure

```text
Tasks/
│
├── TASK_A1/
│   # PDF Crash Report to Structured Excel Conversion
│   # Uses native text extraction (pdfplumber) for digitally generated PDF reports.
│   # Highly accurate and fast, but requires digital PDFs.
│
├── TASK_A1_OCR/
│   # PDF Crash Report to Structured Excel Conversion (OCR Version)
│   # Uses a hybrid approach with PyMuPDF and Tesseract OCR.
│   # Designed to handle scanned or legacy PDF reports lacking embedded text.
│
└── TASK_B1/
    # Vehicle Detection, Tracking, and Trajectory Extraction
    # Uses YOLOv8s and ByteTrack to analyze traffic video segments.
    # Features custom logic to filter out stationary false positives (e.g., bus stands).
```

## How to Run

Navigate into the specific task folder you wish to execute. Example:

```bash
cd TASK_B1
pip install -r requirements.txt
python run.py
```

*Please refer to the `README.md` file located inside each respective task folder for complete, task-specific instructions.*
