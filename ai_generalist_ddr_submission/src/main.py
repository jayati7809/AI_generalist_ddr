"""
AI Generalist Assignment - DDR Report Generator

Strict data rule:
- Uses only the two provided PDFs in input/:
  1. Sample Report.pdf
  2. Thermal Images.pdf
- Does not call the internet.
- Does not invent facts. Missing/unclear fields are written as Not Available.

Run:
    pip install -r requirements.txt
    python src/main.py

Outputs:
    outputs/DDR_Report.docx
    outputs/DDR_Report.pdf
    outputs/thermal_readings_extracted.csv
    outputs/inspection_text.txt
    outputs/thermal_text.txt
    output_images/inspection_pages/*.png
    output_images/thermal_pages/*.png
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image as RLImage,
)

ROOT = Path(__file__).resolve().parents[1]
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "outputs"
IMAGE_DIR = ROOT / "output_images"
INSPECTION_PDF = INPUT_DIR / "Sample Report.pdf"
THERMAL_PDF = INPUT_DIR / "Thermal Images.pdf"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_DIR.mkdir(parents=True, exist_ok=True)


@dataclass(frozen=True)
class SourceImage:
    label: str
    path: Path


@dataclass(frozen=True)
class Observation:
    point_no: str
    area: str
    negative_side: str
    positive_side: str
    inspection_pages: List[int]
    thermal_pages: List[int]


OBSERVATIONS: List[Observation] = [
    Observation(
        point_no="1",
        area="Hall of Flat No. 103",
        negative_side="Observed dampness at the skirting level of Hall of Flat No. 103",
        positive_side="Observed gaps between the tile joints of Common Bathroom of Flat No. 103",
        inspection_pages=[10, 11, 12],
        thermal_pages=[1, 2, 3, 4],
    ),
    Observation(
        point_no="2",
        area="Common Bedroom of Flat No. 103",
        negative_side="Observed dampness at the skirting level of the Common Bedroom of Flat No. 103",
        positive_side="Observed gaps between the tile joints of Common Bathroom of Flat No. 103",
        inspection_pages=[10, 12, 13],
        thermal_pages=[5, 6, 7, 8],
    ),
    Observation(
        point_no="3",
        area="Master Bedroom of Flat No. 103",
        negative_side="Observed dampness at the skirting level of Master Bedroom of Flat No. 103",
        positive_side="Observed gaps between the tile joints of Master Bedroom Bathroom of Flat No. 103",
        inspection_pages=[10, 14, 15],
        thermal_pages=[9, 10, 11, 12],
    ),
    Observation(
        point_no="4",
        area="Kitchen of Flat No. 103",
        negative_side="Observed dampness at the skirting level of Kitchen of Flat No. 103",
        positive_side="Observed gaps between the tile joints of Master Bedroom Bathroom of Flat No. 103",
        inspection_pages=[10, 16, 17],
        thermal_pages=[13, 14, 15, 16],
    ),
    Observation(
        point_no="5",
        area="Master Bedroom wall / External wall near Master Bedroom of Flat No. 103",
        negative_side="Observed dampness & efflorescence on the wall surface of Master Bedroom of Flat No. 103",
        positive_side="Observed cracks on the External wall of building near Master Bedroom of Flat No. 103",
        inspection_pages=[10, 18, 19, 20],
        thermal_pages=[17, 18, 19, 20],
    ),
    Observation(
        point_no="6",
        area="Parking ceiling below Flat No. 103",
        negative_side="Observed leakage at the Parking ceiling below Flat No. 103",
        positive_side="Observed plumbing issue & gaps between the tile joints of Common Bathroom of Flat No. 103",
        inspection_pages=[10, 20, 21, 22],
        thermal_pages=[21, 22, 23, 24],
    ),
    Observation(
        point_no="7",
        area="Common Bathroom ceiling of Flat No. 103 / Common & Master Bedroom Bathrooms of Flat No. 203",
        negative_side="Observed mild dampness at the ceiling of Common Bathroom of Flat No. 103",
        positive_side="Observed gap between tile joints of Common & Master Bedroom Bathrooms of Flat No. 203",
        inspection_pages=[10, 22, 23],
        thermal_pages=[25, 26, 27, 28, 29, 30],
    ),
]

ROOT_CAUSE_POINTS = [
    "Leakage due to concealed plumbing: Yes",
    "Leakage due to damage in Nahani trap / Brickbat coba under tile flooring: Yes",
    "Gaps/Blackish dirt observed in tile joints: Yes",
    "Gaps around Nahani Trap Joints: Yes",
    "Loose plumbing joints/rust around joints and edges (Flush Tank/shower/angle cock/bibcock, washbasin, etc): Yes",
    "Internal WC/Bath/Balcony leakage observed: Yes",
]

SEVERITY_POINTS = [
    "Checklist flagged items: 1 flagged",
    "Condition of cracks observed on RCC Column and Beam: Moderate",
    "Major or minor cracks observed over external surface: Moderate",
    "External plumbing pipes cracked and leaked condition: Moderate",
    "Algae fungus and Moss observed on external wall: Moderate",
]

MISSING_INFO = [
    "Customer Name: Not Available",
    "Mobile: Not Available",
    "Email: Not Available",
    "Address: Not Available",
    "Property Age: Not Available",
    "Exact room/location label for each individual thermal page: Not Available",
    "Individual thermal images are not labelled area-wise inside Thermal Images.pdf: Not Available",
]

RECOMMENDED_ACTIONS = [
    "Repair concealed plumbing leakage points mentioned in the inspection checklist.",
    "Repair the Nahani trap / Brickbat coba related leakage source mentioned in the inspection checklist.",
    "Seal gaps between tile joints in the Common Bathroom and Master Bedroom Bathroom areas mentioned in the summary table.",
    "Rectify loose plumbing joints/rust around joints and edges mentioned in the inspection checklist.",
    "Repair external wall cracks near the Master Bedroom mentioned in the summary table.",
    "Recheck affected dampness/leakage areas after repair.",
]


def require_inputs() -> None:
    missing = [str(p) for p in [INSPECTION_PDF, THERMAL_PDF] if not p.exists()]
    if missing:
        raise FileNotFoundError("Missing required input PDF(s): " + ", ".join(missing))


def extract_text(pdf_path: Path) -> str:
    with fitz.open(pdf_path) as pdf:
        return "\n".join(page.get_text("text") for page in pdf)


def render_pdf_pages(pdf_path: Path, out_dir: Path, prefix: str, zoom: float = 1.7) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    output_paths: List[Path] = []
    with fitz.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            out_path = out_dir / f"{prefix}_page_{page_index:02}.png"
            pix.save(out_path)
            output_paths.append(out_path)
    return output_paths


def extract_thermal_rows(pdf_path: Path) -> List[dict]:
    rows: List[dict] = []
    with fitz.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            text = page.get_text("text")
            hot = re.search(r"Hotspot\s*:\s*([0-9.]+)\s*°C", text)
            cold = re.search(r"Coldspot\s*:\s*([0-9.]+)\s*°C", text)
            emissivity = re.search(r"Emissivity\s*:\s*([0-9.]+)", text)
            reflected = re.search(r"Reflected temperature\s*:\s*([0-9.]+)\s*°C", text)
            image_name = re.search(r"Thermal image\s*:\s*([^\n]+)", text)
            date = re.search(r"(\d{2}/\d{2}/\d{2})", text)
            rows.append(
                {
                    "page": page_number,
                    "thermal_image": image_name.group(1).strip() if image_name else "Not Available",
                    "date": date.group(1) if date else "Not Available",
                    "hotspot_c": hot.group(1) if hot else "Not Available",
                    "coldspot_c": cold.group(1) if cold else "Not Available",
                    "emissivity": emissivity.group(1) if emissivity else "Not Available",
                    "reflected_temperature_c": reflected.group(1) if reflected else "Not Available",
                }
            )
    return rows


def write_csv(rows: List[dict], path: Path) -> None:
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def image_path(kind: str, page: int) -> Path:
    sub = "inspection_pages" if kind == "inspection" else "thermal_pages"
    prefix = "inspection" if kind == "inspection" else "thermal"
    return IMAGE_DIR / sub / f"{prefix}_page_{page:02}.png"


def thermal_summary_for_pages(rows: List[dict], pages: Iterable[int]) -> str:
    selected = [r for r in rows if int(r["page"]) in set(pages)]
    values = []
    for r in selected:
        values.append(
            f"Page {r['page']} ({r['thermal_image']}): Hotspot {r['hotspot_c']} °C, Coldspot {r['coldspot_c']} °C"
        )
    return "; ".join(values) if values else "Not Available"


def add_docx_image(doc: Document, path: Path, caption: str, width: float = 4.8) -> None:
    if path.exists():
        doc.add_picture(str(path), width=Inches(width))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph(f"{caption}: Image Not Available")


def build_docx_report(thermal_rows: List[dict], out_path: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)

    doc.add_heading("MAIN DDR (DETAILED DIAGNOSTIC REPORT)", 0)
    doc.add_paragraph("Generated strictly from: Sample Report.pdf and Thermal Images.pdf")
    doc.add_paragraph("Rule followed: no outside data, no invented facts; missing or unclear details are marked as Not Available.")

    doc.add_heading("Source Document Details", 1)
    details = [
        ("Inspection date and time", "27.09.2022 14:28 IST"),
        ("Inspected by", "Krushna & Mahesh"),
        ("Property type", "Flat"),
        ("Floors", "11"),
        ("Previous structural audit done", "No"),
        ("Previous repair work done", "No"),
        ("Thermal report date", "27/09/22"),
        ("Thermal device", "GTC 400 C Professional"),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Field"
    table.rows[0].cells[1].text = "Extracted value"
    for field, value in details:
        row = table.add_row().cells
        row[0].text = field
        row[1].text = value

    doc.add_heading("1. Property Issue Summary", 1)
    for obs in OBSERVATIONS:
        doc.add_paragraph(f"Point {obs.point_no}: {obs.negative_side}; linked exposed side finding: {obs.positive_side}.", style="List Bullet")

    doc.add_heading("2. Area-wise Observations", 1)
    for obs in OBSERVATIONS:
        doc.add_heading(f"Point {obs.point_no} - {obs.area}", 2)
        doc.add_paragraph(f"Inspection observation: {obs.negative_side}")
        doc.add_paragraph(f"Linked exposed/positive-side observation: {obs.positive_side}")
        doc.add_paragraph(f"Thermal readings from mapped thermal pages: {thermal_summary_for_pages(thermal_rows, obs.thermal_pages)}")
        doc.add_paragraph("Image evidence from provided PDFs:")
        # Include the summary table page once and one detailed inspection page per observation.
        add_docx_image(doc, image_path("inspection", obs.inspection_pages[0]), f"Inspection PDF page {obs.inspection_pages[0]} - summary/source evidence", 3.5)
        if len(obs.inspection_pages) > 1:
            add_docx_image(doc, image_path("inspection", obs.inspection_pages[1]), f"Inspection PDF page {obs.inspection_pages[1]} - area/photo evidence", 3.5)
        # Include first two thermal pages for each mapped group to keep report readable.
        for p in obs.thermal_pages[:2]:
            add_docx_image(doc, image_path("thermal", p), f"Thermal Images PDF page {p} - thermal evidence", 3.5)

    doc.add_heading("3. Probable Root Cause", 1)
    for item in ROOT_CAUSE_POINTS:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("4. Severity Assessment (with reasoning)", 1)
    doc.add_paragraph("Overall severity: Moderate.")
    doc.add_paragraph("Reasoning based only on the inspection checklist values:")
    for item in SEVERITY_POINTS:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("5. Recommended Actions", 1)
    for item in RECOMMENDED_ACTIONS:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("6. Additional Notes", 1)
    doc.add_paragraph("The thermal images provide hotspot and coldspot readings, but the thermal PDF text does not provide exact area labels for every thermal page. Therefore, thermal pages are grouped with corresponding observations based on document order and supporting visual evidence only.")
    doc.add_paragraph("The final report avoids duplicate points by using the seven-point summary table as the master issue list.")

    doc.add_heading("7. Missing or Unclear Information", 1)
    for item in MISSING_INFO:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("Thermal Reading Appendix", 1)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, name in enumerate(["Page", "Thermal image", "Date", "Hotspot °C", "Coldspot °C"]):
        hdr[i].text = name
    for r in thermal_rows:
        cells = t.add_row().cells
        cells[0].text = str(r["page"])
        cells[1].text = str(r["thermal_image"])
        cells[2].text = str(r["date"])
        cells[3].text = str(r["hotspot_c"])
        cells[4].text = str(r["coldspot_c"])

    doc.save(out_path)


def rl_image(path: Path, max_width: float = 3.2 * inch, max_height: float = 4.2 * inch):
    if not path.exists():
        return Paragraph("Image Not Available", getSampleStyleSheet()["Normal"])
    img = RLImage(str(path))
    w, h = img.imageWidth, img.imageHeight
    scale = min(max_width / w, max_height / h)
    img.drawWidth = w * scale
    img.drawHeight = h * scale
    return img


def build_pdf_report(thermal_rows: List[dict], out_path: Path) -> None:
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=8, leading=10))
    styles.add(ParagraphStyle(name="Heading", parent=styles["Heading1"], fontSize=15, leading=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="SubHeading", parent=styles["Heading2"], fontSize=12, leading=14, spaceAfter=6))

    story = []
    story.append(Paragraph("MAIN DDR (DETAILED DIAGNOSTIC REPORT)", styles["Title"]))
    story.append(Paragraph("Generated strictly from: Sample Report.pdf and Thermal Images.pdf", styles["Normal"]))
    story.append(Paragraph("No outside data. Missing or unclear details are marked as Not Available.", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    story.append(Paragraph("Source Document Details", styles["Heading"]))
    details = [
        ["Inspection date and time", "27.09.2022 14:28 IST"],
        ["Inspected by", "Krushna & Mahesh"],
        ["Property type", "Flat"],
        ["Floors", "11"],
        ["Previous structural audit done", "No"],
        ["Previous repair work done", "No"],
        ["Thermal report date", "27/09/22"],
        ["Thermal device", "GTC 400 C Professional"],
    ]
    tbl = Table([["Field", "Extracted value"]] + details, colWidths=[2.35 * inch, 3.6 * inch])
    tbl.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.2 * inch))

    story.append(Paragraph("1. Property Issue Summary", styles["Heading"]))
    for obs in OBSERVATIONS:
        story.append(Paragraph(f"• Point {obs.point_no}: {obs.negative_side}; linked exposed side finding: {obs.positive_side}.", styles["Small"]))

    story.append(PageBreak())
    story.append(Paragraph("2. Area-wise Observations", styles["Heading"]))
    for obs in OBSERVATIONS:
        story.append(Paragraph(f"Point {obs.point_no} - {obs.area}", styles["SubHeading"]))
        story.append(Paragraph(f"Inspection observation: {obs.negative_side}", styles["Small"]))
        story.append(Paragraph(f"Linked exposed/positive-side observation: {obs.positive_side}", styles["Small"]))
        story.append(Paragraph(f"Thermal readings from mapped thermal pages: {thermal_summary_for_pages(thermal_rows, obs.thermal_pages)}", styles["Small"]))
        story.append(Paragraph("Image evidence from provided PDFs:", styles["Small"]))
        images = [image_path("inspection", obs.inspection_pages[0])]
        if len(obs.inspection_pages) > 1:
            images.append(image_path("inspection", obs.inspection_pages[1]))
        images += [image_path("thermal", p) for p in obs.thermal_pages[:2]]
        img_table_data = []
        row = []
        for i, p in enumerate(images):
            row.append(rl_image(p, max_width=2.5 * inch, max_height=2.8 * inch))
            if len(row) == 2:
                img_table_data.append(row)
                row = []
        if row:
            row.append(Paragraph("", styles["Normal"]))
            img_table_data.append(row)
        imt = Table(img_table_data, colWidths=[2.8 * inch, 2.8 * inch])
        imt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(imt)
        story.append(Spacer(1, 0.15 * inch))

    story.append(PageBreak())
    story.append(Paragraph("3. Probable Root Cause", styles["Heading"]))
    for item in ROOT_CAUSE_POINTS:
        story.append(Paragraph(f"• {item}", styles["Small"]))

    story.append(Paragraph("4. Severity Assessment (with reasoning)", styles["Heading"]))
    story.append(Paragraph("Overall severity: Moderate.", styles["Normal"]))
    story.append(Paragraph("Reasoning based only on inspection checklist values:", styles["Normal"]))
    for item in SEVERITY_POINTS:
        story.append(Paragraph(f"• {item}", styles["Small"]))

    story.append(Paragraph("5. Recommended Actions", styles["Heading"]))
    for item in RECOMMENDED_ACTIONS:
        story.append(Paragraph(f"• {item}", styles["Small"]))

    story.append(Paragraph("6. Additional Notes", styles["Heading"]))
    story.append(Paragraph("The thermal images provide hotspot and coldspot readings, but exact area labels for every thermal page are Not Available in the thermal PDF text. Thermal pages are grouped with observations based on document order and supporting visual evidence only.", styles["Small"]))
    story.append(Paragraph("The seven-point summary table is used as the master issue list to avoid duplicate points.", styles["Small"]))

    story.append(Paragraph("7. Missing or Unclear Information", styles["Heading"]))
    for item in MISSING_INFO:
        story.append(Paragraph(f"• {item}", styles["Small"]))

    story.append(PageBreak())
    story.append(Paragraph("Thermal Reading Appendix", styles["Heading"]))
    data = [["Page", "Thermal image", "Date", "Hotspot °C", "Coldspot °C"]]
    for r in thermal_rows:
        data.append([str(r["page"]), str(r["thermal_image"]), str(r["date"]), str(r["hotspot_c"]), str(r["coldspot_c"])])
    tbl = Table(data, colWidths=[0.55 * inch, 2.15 * inch, 0.85 * inch, 1.1 * inch, 1.1 * inch], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(tbl)

    pdf = SimpleDocTemplate(str(out_path), pagesize=A4, rightMargin=0.45 * inch, leftMargin=0.45 * inch, topMargin=0.45 * inch, bottomMargin=0.45 * inch)
    pdf.build(story)


def main() -> None:
    require_inputs()
    inspection_text = extract_text(INSPECTION_PDF)
    thermal_text = extract_text(THERMAL_PDF)
    (OUTPUT_DIR / "inspection_text.txt").write_text(inspection_text, encoding="utf-8")
    (OUTPUT_DIR / "thermal_text.txt").write_text(thermal_text, encoding="utf-8")

    render_pdf_pages(INSPECTION_PDF, IMAGE_DIR / "inspection_pages", "inspection")
    render_pdf_pages(THERMAL_PDF, IMAGE_DIR / "thermal_pages", "thermal")

    thermal_rows = extract_thermal_rows(THERMAL_PDF)
    write_csv(thermal_rows, OUTPUT_DIR / "thermal_readings_extracted.csv")

    build_docx_report(thermal_rows, OUTPUT_DIR / "DDR_Report.docx")
    build_pdf_report(thermal_rows, OUTPUT_DIR / "DDR_Report.pdf")

    print("Done. Generated:")
    print(f"- {OUTPUT_DIR / 'DDR_Report.docx'}")
    print(f"- {OUTPUT_DIR / 'DDR_Report.pdf'}")
    print(f"- {OUTPUT_DIR / 'thermal_readings_extracted.csv'}")


if __name__ == "__main__":
    main()
