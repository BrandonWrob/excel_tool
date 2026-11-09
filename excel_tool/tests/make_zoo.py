"""A workbook of awkward layouts, then recalculated by LibreOffice so formulas have saved
results like a real Excel save. Also writes two awkward CSVs."""
import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

out = Path(sys.argv[1]).resolve()
out.mkdir(parents=True, exist_ok=True)
wb = openpyxl.Workbook()

# 1. Two stacked tables, each with a title, a totals row and a footnote
ws = wb.active
ws.title = "Stacked"
ws["A1"] = "Monthly Shipments - East"
ws.append([])
ws.append(["Month", "Units", "Revenue"])
for m, u, r in (("Jul", 120, 1500.0), ("Aug", 80, 2400.0), ("Sep", 200, 2500.0)):
    ws.append([m, u, r])
ws.append(["Total", "=SUM(B4:B6)", "=SUM(C4:C6)"])
ws.append([]); ws.append([])
ws["A10"] = "Monthly Shipments - West"
ws["A11"], ws["B11"], ws["C11"] = "Month", "Units", "Revenue"
for i, (m, u, r) in enumerate((("Jul", 45, 1350.0), ("Aug", 150, 1875.0), ("Sep", 60, 1800.0)), start=12):
    ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"] = m, u, r
ws["A16"] = "* Revenue in USD; excludes returns."
for r in range(4, 16):
    ws[f"C{r}"].number_format = '"$"#,##0.00'

# 2. Two tables side by side with a two-column gap
ws = wb.create_sheet("SideBySide")
for row in (["Lot", "Qty", "Site"], ["L001", 10, "RDU"], ["L002", 12, "RDU"], ["L003", 9, "CLT"],
            ["L004", 14, "CLT"], ["L005", 11, "RDU"]):
    ws.append(row)
for i, row in enumerate((["Site", "Manager"], ["RDU", "Ana Ruiz"], ["CLT", "Ben Okafor"]), start=1):
    ws.cell(i, 6, row[0]); ws.cell(i, 7, row[1])

# 3. Crosstab with a two-level header (merged years over quarters) and a totals row
ws = wb.create_sheet("Forecast")
ws["A1"] = "2025-2026 Unit Forecast"
ws["B3"] = 2025; ws.merge_cells("B3:E3")
ws["F3"] = 2026; ws.merge_cells("F3:I3")
ws["A4"] = "Product"
for j, q in enumerate(["Q1", "Q2", "Q3", "Q4"] * 2):
    ws.cell(4, 2 + j, q)
products = [("Amoxil", 100), ("Zestril", 80), ("Lipitor", 150), ("Norvasc", 60)]
for i, (p, base) in enumerate(products, start=5):
    ws.cell(i, 1, p)
    for j in range(8):
        ws.cell(i, 2 + j, base + 5 * j)
ws.cell(9, 1, "Total")
for j in range(8):
    col = openpyxl.utils.get_column_letter(2 + j)
    ws.cell(9, 2 + j, f"=SUM({col}5:{col}8)")

# 4. Batch record form (label: value pairs in two column pairs) with a results table below
ws = wb.create_sheet("BatchRecord")
form = [("Batch Number:", "ACME-0042", "Manufacture Date:", dt.date(2026, 10, 1)),
        ("Product:", "Amoxicillin 500mg", "Expiry Date:", dt.date(2028, 10, 1)),
        ("Batch Size (kg):", 250, "Approved By:", "J. Smith")]
for i, (a, b, c, d) in enumerate(form, start=2):
    ws.cell(i, 2, a); ws.cell(i, 3, b); ws.cell(i, 5, c); ws.cell(i, 6, d)
    for col in (3, 6):
        if isinstance(ws.cell(i, col).value, dt.date):
            ws.cell(i, col).number_format = "yyyy-mm-dd"
ws["B7"], ws["C7"], ws["D7"], ws["E7"] = "Test", "Spec", "Result", "Pass?"
for i, row in enumerate([("Assay %", "95-105", 99.2, "Yes"), ("Dissolution %", ">=80", 86.0, "Yes"),
                         ("Water %", "<=1.0", 1.3, "No")], start=8):
    for j, v in enumerate(row):
        ws.cell(i, 2 + j, v)

# 5. Group label rows, subtotals and a grand total
ws = wb.create_sheet("Grouped")
ws.append(["Region", "Product", "Units", "Revenue"])
r = 2
for region, items in (("East", [("Amoxil", 120, 1500), ("Zestril", 80, 2400)]),
                      ("West", [("Amoxil", 200, 2500), ("Zestril", 45, 1350), ("Lipitor", 30, 900)])):
    ws.cell(r, 1, region); r += 1
    first = r
    for prod, u, rev in items:
        ws.cell(r, 2, prod); ws.cell(r, 3, u); ws.cell(r, 4, rev); r += 1
    ws.cell(r, 1, f"Subtotal {region}")
    ws.cell(r, 3, f"=SUM(C{first}:C{r - 1})"); ws.cell(r, 4, f"=SUM(D{first}:D{r - 1})"); r += 1
ws.cell(r, 1, "Grand Total"); ws.cell(r, 3, "=SUM(C3,C4,C7,C8,C9)"); ws.cell(r, 4, "=D5+D10")

# 6. Dirty data
ws = wb.create_sheet("DirtyData")
ws.append(["Item ID", "Qty", "Qty", None, "Received", "Unit Cost"])
rows = [("00123", 15, 3, "x", dt.date(2026, 9, 1), 2.5), ("00124", "18", 4, "", dt.date(2026, 9, 3), "#N/A"),
        ("00125", "TBD", 1, "", dt.date(2026, 9, 7), 3.75), ("00126", 22, 2, "y", "not received", 4.0),
        ("00127", "9", 5, "", dt.date(2026, 9, 9), 1.25)]
for row in rows:
    ws.append(list(row))
for r in range(2, 7):
    ws[f"E{r}"].number_format = "yyyy-mm-dd"

# 7. Filtered rows
ws = wb.create_sheet("Filtered")
ws.append(["Order", "Status", "Amount"])
for i in range(1, 9):
    ws.append([f"SO-{1000 + i}", "Open" if i % 3 else "Closed", i * 100])
ws.auto_filter.ref = "A1:C9"
for r in (4, 7):
    ws.row_dimensions[r].hidden = True

# 8. Excel table with a totals row
ws = wb.create_sheet("TableTotals")
ws.append(["SKU", "Units", "Value"])
for row in (["A-1", 5, 50], ["A-2", 7, 70], ["A-3", 2, 20]):
    ws.append(row)
ws.append(["Total", "=SUBTOTAL(109,B2:B4)", "=SUBTOTAL(109,C2:C4)"])
t = Table(displayName="Inventory", ref="A1:C5", totalsRowCount=1,
          tableStyleInfo=TableStyleInfo(name="TableStyleLight9"))
ws.add_table(t)

# 9. Sparse wide checklist, 10. no-header numeric block, 11. notes
ws = wb.create_sheet("Checklist")
ws.append(["Step", "Owner", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Notes"])
for i in range(2, 10):
    ws.cell(i, 1, f"Step {i - 1}")
    ws.cell(i, 2 + (i % 8), "x")
ws.cell(5, 10, "Needs QA sign-off")
ws = wb.create_sheet("RawNumbers")
for r in range(1, 6):
    ws.append([r * 1.5, r * 2, r * 3, r * 4])
ws = wb.create_sheet("Notes")
ws["A1"] = "This workbook tracks Q3 shipments for the East and West regions."
ws["A2"] = "Figures are preliminary until finance closes the quarter on Oct 15."
ws["A3"] = "Contact supply-chain@example.com with corrections."

base = out / "zoo_base.xlsx"
wb.save(base)
subprocess.run(["soffice", "-env:UserInstallation=file:///tmp/lo_zoo", "--headless", "--convert-to",
                "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out / "recalc"), str(base)],
               check=True, capture_output=True, timeout=120)
(out / "recalc" / "zoo_base.xlsx").replace(out / "zoo.xlsx")

# CSVs
(out / "report_export.csv").write_text(
    "Shipments Report,,\n"
    "Generated 2026-11-01,,\n"
    ",,\n"
    "Region,Units,Revenue\n"
    "East,400,\"6,400.00\"\n"
    "West,255,\"5,025.00\"\n"
    "Total,655,\"11,425.00\"\n"
    ",,\n"
    "Returns,,\n"
    "Region,Units,Reason\n"
    "East,3,Damaged\n"
    "West,1,Wrong item\n", encoding="utf-8")
(out / "eu_prices.csv").write_bytes(
    "Produkt;Preis;Menge;NDC\nAmoxil;12,50;100;00093-7180-01\nZestril;30,00;45;00093-1111-01\n"
    "Lipitor;1.234,75;12;00071-0155-23\nSumme;1.277,25;157;\n".encode("cp1252"))
print("zoo written")
