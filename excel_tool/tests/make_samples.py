"""Build deliberately messy sample files for testing doc_to_text."""
import datetime as dt
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)

# ---------- messy.xlsx
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Q3 Sales"
ws["B2"] = "Q3 2026 Regional Sales"
ws.merge_cells("B2:G2")
ws.append([])  # row 3 blank
hdr = ["Region", "Ship Date", "NDC", "Units", "Unit Price", "Discount", "Revenue"]
for i, h in enumerate(hdr):
    ws.cell(4, 2 + i, h)
data = [
    ("East", dt.datetime(2026, 7, 1), "00093-7180-01", 120, 12.5, 0.15),
    ("East", dt.datetime(2026, 7, 15, 14, 30), "00093-7180-01", 80, 12.5, 0.1),
    ("West", dt.datetime(2026, 8, 2), "61314-0227-05", 200, 0.1 + 0.2, 0),
    ("West", dt.datetime(2026, 9, 9), "61314-0227-05", 45, 9.99, 0.125),
]
for r, row in enumerate(data, start=5):
    for c, v in enumerate(row, start=2):
        ws.cell(r, c, v)
    ws.cell(r, 3).number_format = "yyyy-mm-dd" if r != 6 else "yyyy-mm-dd h:mm"
    ws.cell(r, 7).number_format = "0.0%"
    ws.cell(r, 8, f"=E{r}*F{r}*(1-G{r})")  # formula, no cached value (openpyxl can't compute)
# merged vertical cell
ws["J4"] = "Territory"
ws["J5"] = "Northeast"
ws.merge_cells("J5:J6")
# zip code stored as number with zero-padded format
ws["K4"] = "Zip"
ws["K5"] = 2134
ws["K5"].number_format = "00000"
# text with Excel carriage-return artifact
ws["L4"] = "Note"
ws["L5"] = "Line one_x000D_\nLine two"
ws["F8"] = True
# hidden row
ws.row_dimensions[7].hidden = True
# formatting far away inflates max_row
ws.cell(50000, 30).font = Font(bold=True)

hidden = wb.create_sheet("Lookup")
hidden.append(["Code", "Meaning"])
hidden.append(["A1", "Approved"])
hidden.sheet_state = "hidden"
wb.create_sheet("Empty")
wb.save(out / "messy.xlsx")

# ---------- big.xlsx (should route to SQL)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Shipments"
ws.append(["Shipment ID", "Product", "Lot", "Qty", "Ship Date"])
for i in range(30_000):
    ws.append([f"S{i:06d}", f"Product {i % 37}", f"L{i % 997:04d}", (i * 7) % 500, dt.date(2026, 1, 1) + dt.timedelta(days=i % 300)])
wb.save(out / "big.xlsx")

# ---------- semicolon cp1252 csv with ragged row + leading zeros
csv_text = (
    "Product;Description;NDC;Price\r\n"
    "Café Rx;\"Contains ; semicolon\";00093-7180-01;12,50\r\n"
    "Ibuprofen 200mg;Généric – tablet;00904-5853-61;3,10\r\n"
    "Broken row;missing fields\r\n"
    ";;;\r\n"
)
(out / "export_cp1252.csv").write_bytes(csv_text.encode("cp1252"))

# ---------- utf-8 BOM tsv
(out / "bom.tsv").write_bytes("﻿id\tname\tzip\n1\tAlice\t02134\n2\tBob\t07030\n".encode("utf-8"))

# ---------- comma csv whose text contains semicolons (fools csv.Sniffer)
(out / "tricky.csv").write_text(
    "id,comment,amount\n1,\"ok; fine; good\",10\n2,\"a; b\",20\n3,plain,30\n", encoding="utf-8")

# ---------- docx
import docx
from docx.shared import Inches

d = docx.Document()
d.add_heading("Stability Study Summary", level=0)
d.add_heading("1. Background", level=1)
d.add_paragraph("Batch ACME-001 was stored at 25C/60%RH.")
d.add_paragraph("First finding", style="List Bullet")
d.add_paragraph("Second finding", style="List Bullet")
t = d.add_table(rows=3, cols=3)
for r, row in enumerate([["Month", "Assay %", "Result"], ["0", "100.2", "Pass"], ["3", "99.1", "Pass"]]):
    for c, v in enumerate(row):
        t.cell(r, c).text = v
d.sections[0].header.paragraphs[0].text = "CONFIDENTIAL"
# tracked insertion inserted manually into a paragraph
p = d.add_paragraph("Conclusion: product is ")
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
ins = OxmlElement("w:ins")
ins.set(qn("w:id"), "1"); ins.set(qn("w:author"), "QA"); ins.set(qn("w:date"), "2026-11-01T00:00:00Z")
run = OxmlElement("w:r"); tt = OxmlElement("w:t"); tt.text = "stable for 24 months."
run.append(tt); ins.append(run); p._p.append(ins)
d.save(out / "report.docx")

# ---------- pptx with chart, table, notes, group
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches as PI

prs = Presentation()
s = prs.slides.add_slide(prs.slide_layouts[1])
s.shapes.title.text = "Q3 Highlights"
body = s.placeholders[1].text_frame
body.text = "Revenue up 12%"
p2 = body.add_paragraph(); p2.text = "Driven by East region"; p2.level = 1
s.notes_slide.notes_text_frame.text = "Mention the West backorder."
s2 = prs.slides.add_slide(prs.slide_layouts[5])
s2.shapes.title.text = "Units by Region"
cd = CategoryChartData(); cd.categories = ["East", "West"]; cd.add_series("Q2", (150, 180)); cd.add_series("Q3", (200, 245))
s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PI(1), PI(1.5), PI(6), PI(4), cd)
tbl = s2.shapes.add_table(2, 2, PI(1), PI(6), PI(4), PI(1)).table
tbl.cell(0, 0).text = "KPI"; tbl.cell(0, 1).text = "Value"; tbl.cell(1, 0).text = "OTIF"; tbl.cell(1, 1).text = "97%"
prs.save(out / "deck.pptx")

print("samples written to", out)
