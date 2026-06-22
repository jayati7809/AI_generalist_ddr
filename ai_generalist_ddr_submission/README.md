# AI-Powered DDR Report Generation System

## Overview

This project is an AI-powered DDR (Detailed Diagnostic Report) Generation System designed to automate the process of converting raw Inspection Reports and Thermal Reports into structured, client-ready diagnostic reports.

The system processes technical inspection documents, extracts observations, thermal findings, and relevant images, and generates a comprehensive DDR report in both DOCX and PDF formats. It also extracts thermal readings into CSV format for traceability and future analysis.

---

## Objective

The objective of this project is to eliminate manual report preparation by automatically analyzing inspection documents and generating professional diagnostic reports while preserving accuracy, consistency, and transparency.

---

## Features

* Automated PDF processing
* Inspection report text extraction
* Thermal report text extraction
* Thermal reading extraction
* Image extraction from source documents
* Observation consolidation
* Structured DDR report generation
* DOCX report generation
* PDF report generation
* Missing information identification
* Traceable thermal data export

---

## DDR Report Structure

The generated report contains:

1. Property Issue Summary
2. Area-wise Observations
3. Probable Root Cause
4. Severity Assessment
5. Recommended Actions
6. Additional Notes
7. Missing or Unclear Information

---

## Project Architecture

```text
Inspection Report PDF
            +
Thermal Report PDF
            │
            ▼
      PDF Processing
            │
            ▼
 Text & Image Extraction
            │
            ▼
 Thermal Data Processing
            │
            ▼
 Observation Consolidation
            │
            ▼
     DDR Generation
            │
            ▼
  DOCX + PDF Outputs
```

---

## Project Structure

```text
ai_generalist_ddr_submission/
│
├── input/
│   ├── Sample Report.pdf
│   └── Thermal Images.pdf
│
├── outputs/
│   ├── DDR_Report.docx
│   ├── DDR_Report.pdf
│   ├── inspection_text.txt
│   ├── thermal_text.txt
│   └── thermal_readings_extracted.csv
│
├── output_images/
│
├── rendered/
│
├── src/
│   └── main.py
│
├── main.py
├── requirements.txt
├── README.md
└── Loom_Video_Script.txt
```

---

## Technologies Used

| Technology  | Purpose                       |
| ----------- | ----------------------------- |
| Python      | Core implementation           |
| PyMuPDF     | PDF processing and extraction |
| python-docx | Word document generation      |
| ReportLab   | PDF generation                |
| CSV         | Thermal data export           |

---

## Installation

Clone the repository:

```bash
git clone https://github.com/jayati7809/AI_generalist_ddr.git
cd AI_generalist_ddr
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

Run the application:

```bash
python main.py
```

Generated outputs:

```text
outputs/
├── DDR_Report.docx
├── DDR_Report.pdf
├── inspection_text.txt
├── thermal_text.txt
└── thermal_readings_extracted.csv
```

---

## Sample Workflow

```text
Input Documents
│
├── Inspection Report
└── Thermal Report
      │
      ▼
Automated Processing
      │
      ▼
Information Extraction
      │
      ▼
DDR Report Generation
      │
      ▼
Client-Ready Outputs
```

---

## Limitations

* Optimized for report formats similar to the provided sample documents.
* Rule-based report generation approach.
* Limited support for significantly different PDF layouts.

---

## Future Enhancements

* Multi-format inspection report support
* LLM-powered reasoning and summarization
* Automated severity scoring
* Web-based dashboard
* Batch report processing
* Advanced image classification and annotation



