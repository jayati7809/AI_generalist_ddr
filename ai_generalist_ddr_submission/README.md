# AI Generalist - Applied AI Builder DDR Report Generator

This is a VS Code / GitHub-ready project for the **Applied AI Builder (DDR Report Generation)** assignment.

It reads the two provided PDFs and generates a structured DDR report using only the data inside those PDFs.

## Input PDFs

Keep these files inside the `input/` folder:

- `Sample Report.pdf`
- `Thermal Images.pdf`

## What the system generates

After running the project, these files are created:

- `outputs/DDR_Report.docx` - editable DDR report
- `outputs/DDR_Report.pdf` - final PDF DDR report
- `outputs/thermal_readings_extracted.csv` - thermal metadata extracted from the thermal PDF
- `outputs/inspection_text.txt` - extracted text from the inspection PDF
- `outputs/thermal_text.txt` - extracted text from the thermal PDF
- `output_images/inspection_pages/` - rendered images from the inspection PDF
- `output_images/thermal_pages/` - rendered images from the thermal PDF

## How to run in VS Code

```bash
pip install -r requirements.txt
python src/main.py
```

## Workflow

```text
Sample Report.pdf
        +
Thermal Images.pdf
        ↓
Extract text
        ↓
Extract/render source images
        ↓
Extract thermal readings
        ↓
Build structured DDR data
        ↓
Generate DOCX + PDF report
```

## DDR sections included

1. Property Issue Summary
2. Area-wise Observations
3. Probable Root Cause
4. Severity Assessment with reasoning
5. Recommended Actions
6. Additional Notes
7. Missing or Unclear Information

## Strict rules followed

- No internet data used.
- No outside assumptions added.
- Missing details are written as `Not Available`.
- Conflicting/unclear information is explicitly mentioned.
- Images are taken directly from the uploaded PDF pages.
- The system is designed to run again from the same two PDFs.

## GitHub submission suggestion

Upload this full folder to GitHub and share the repository link. Also include the final `outputs/DDR_Report.pdf`, `outputs/DDR_Report.docx`, and Loom video link in your Google Drive submission folder.
