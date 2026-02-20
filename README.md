# Automated PDF-to-Excel Data Extraction Pipeline (with OCR Fallback)

##  Overview
This project automates the extraction of structured information from a collection of institutional PDF documents and compiles the results into a clean, analysis-ready Excel dataset.

The solution is designed to handle both:
-  **Text-based PDFs** (machine-readable)
-  **Scanned/Image PDFs** (non-readable)

When direct text extraction fails, the system automatically switches to **OCR (Optical Character Recognition)** to recover the content — ensuring reliable extraction regardless of document format.

---

## Problem Statement
Manually extracting contact and institutional details from hundreds of PDFs is:

- Time-consuming  
- Error-prone  
- Difficult to scale  
- Inconsistent due to mixed document formats  

This project converts semi-structured PDF documents into a structured dataset automatically.

---

## How It Works

### Step 1 — Batch Processing
- Reads all PDFs from a specified folder.
- Iterates through each file automatically.

### Step 2 — Attempt Direct Text Extraction
Uses **PyPDF2** to extract text from:
- Digital / machine-generated PDFs.

If text is successfully extracted → proceed to parsing.

---

### Step 3 — OCR Fallback (When PDF is Scanned)
If no text is detected:

1. Extract images from the PDF using **PyMuPDF (fitz)**  
2. Convert images to readable text using **Tesseract OCR**  
3. Clean and process the OCR output  

This ensures scanned documents are handled seamlessly.

---

### Step 4 — Data Parsing
The script extracts key attributes such as:

- College Name  
- City  
- Email ID  
- Phone Number  
- Website  
- Institution Type  

Text is cleaned to remove OCR noise and formatting artifacts.

---

### Step 5 — Structured Output
All extracted data is written into an Excel workbook using **openpyxl**, producing a clean tabular dataset ready for analysis.

---

## Technology Stack

| Tool | Purpose |
|------|---------|
Python | Core automation |
PyPDF2 | Text extraction from digital PDFs |
PyMuPDF (fitz) | Image extraction from scanned PDFs |
Tesseract OCR | Optical Character Recognition |
Pillow (PIL) | Image preprocessing |
openpyxl | Excel generation |
Pathlib | File system automation |

---

## Key Features

✔ Handles mixed PDF types automatically  
✔ Intelligent fallback from text extraction → OCR  
✔ Batch processing for large datasets  
✔ Converts unstructured documents into structured data  
✔ Eliminates manual transcription work  
✔ Designed for scalability and reuse  

---

## Example Output

| File Name | College Name | City | Email | Phone | Website | Type |
|-----------|--------------|------|-------|-------|---------|------|

---

## Use Cases

- Digitizing legacy document archives  
- Creating institutional directories  
- Preparing datasets for analytics/reporting  
- Automating administrative record consolidation  
- Document ingestion pipelines in data engineering workflows  

---

## What This Project Demonstrates

This project showcases practical skills in:

- Document data engineering  
- OCR integration in automation pipelines  
- Handling unstructured/semi-structured data  
- Real-world data cleaning challenges  
- Building resilient extraction workflows  
- Transforming legacy data into usable formats  

---

## Potential Enhancements

- Add regex-based parsing for improved accuracy  
- Store outputs in SQL database instead of Excel  
- Introduce multiprocessing for faster execution  
- Add validation for email/URL fields  
- Build a dashboard (Power BI / Streamlit) on extracted data  
- Package as a CLI tool for reusable deployments  

---

## ⚠️ Requirements

Install dependencies before running:

```bash
pip install PyPDF2 pymupdf pytesseract pillow openpyxl
