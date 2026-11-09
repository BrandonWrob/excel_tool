"""Build an Excel file with a pivot table, chart, conditional formatting, comments, highlights,
currency formats, named range and an Excel Table. Base content via openpyxl, then LibreOffice
(UNO) adds the pivot + chart and computes formula results, the way a real Excel save would."""
import os
import subprocess
import sys
import time
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

out = Path(sys.argv[1]).resolve()
base = out / "rich_base.xlsx"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Region", "Product", "Month", "Units", "Price", "Revenue", "Margin", "Status"])
rows = [
    ("East", "Amoxil", "Jul", 120, 12.5, 0.22, "Shipped"),
    ("East", "Zestril", "Aug", 80, 30.0, 0.08, "Pending"),
    ("West", "Amoxil", "Jul", 200, 12.5, 0.25, "Shipped"),
    ("West", "Zestril", "Sep", 45, 30.0, 0.05, "Backorder"),
    ("South", "Amoxil", "Aug", 150, 12.5, 0.18, "Shipped"),
    ("South", "Zestril", "Sep", 60, 30.0, 0.12, "Pending"),
]
for i, (reg, prod, mon, units, price, margin, status) in enumerate(rows, start=2):
    ws.append([reg, prod, mon, units, price, f"=D{i}*E{i}", margin, status])
    ws[f"E{i}"].number_format = '"$"#,##0.00'
    ws[f"F{i}"].number_format = '"$"#,##0.00'
    ws[f"G{i}"].number_format = "0%"
ws["E9"] = "Total"
ws["F9"] = "=SUM(F2:F7)"
ws["F9"].number_format = '"$"#,##0.00'
ws.conditional_formatting.add("G2:G7", CellIsRule(operator="lessThan", formula=["0.1"],
                                                  fill=PatternFill("solid", start_color="FFC7CE")))
yellow = PatternFill("solid", start_color="FFFF00")
ws["H3"].fill = yellow
ws["H7"].fill = yellow
ws["H5"].comment = Comment("Supplier delay, ETA Oct 15", "QA")
ws.add_table(Table(displayName="SalesTable", ref="A1:H7",
                   tableStyleInfo=TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)))
wb.defined_names["TotalRevenue"] = DefinedName("TotalRevenue", attr_text="Sales!$F$9")

summary = wb.create_sheet("Summary")
summary.append(["Region", "Units"])
for reg, u in (("East", 200), ("West", 245), ("South", 210)):
    summary.append([reg, u])
wb.save(base)

# ---- LibreOffice: add pivot + chart, recalc, save as xlsx and xlsm
port = 2087
proc = subprocess.Popen(["soffice", "-env:UserInstallation=file:///tmp/lo_uno", "--headless", "--invisible",
                         f"--accept=socket,host=localhost,port={port};urp;", "--norestore"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402
from com.sun.star.awt import Rectangle  # noqa: E402
from com.sun.star.table import CellAddress, CellRangeAddress  # noqa: E402

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(60):
    try:
        ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(0.5)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(base)), "_blank", 0, (pv("Hidden", True),))
sheets = doc.Sheets
sales = sheets.getByName("Sales")
sheets.insertNewByName("Pivot", 2)
pivot_sheet = sheets.getByName("Pivot")

src = CellRangeAddress()
src.Sheet, src.StartColumn, src.StartRow, src.EndColumn, src.EndRow = 0, 0, 0, 7, 6
tables = pivot_sheet.getDataPilotTables()
desc = tables.createDataPilotDescriptor()
desc.setSourceRange(src)
fields = desc.getDataPilotFields()
ROW = uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "ROW")
COLUMN = uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "COLUMN")
DATA = uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "DATA")
SUM = uno.Enum("com.sun.star.sheet.GeneralFunction", "SUM")
fields.getByIndex(0).Orientation = ROW       # Region
fields.getByIndex(1).Orientation = COLUMN    # Product
rev = fields.getByIndex(5)                   # Revenue
rev.Orientation = DATA
rev.Function = SUM
dest = CellAddress()
dest.Sheet, dest.Column, dest.Row = 2, 0, 0
tables.insertNewByName("RevenuePivot", dest, desc)

summ = sheets.getByName("Summary")
rng = CellRangeAddress()
rng.Sheet, rng.StartColumn, rng.StartRow, rng.EndColumn, rng.EndRow = 1, 0, 0, 1, 3
rect = Rectangle()
rect.X, rect.Y, rect.Width, rect.Height = 4000, 1000, 12000, 8000
summ.Charts.addNewByName("UnitsChart", rect, (rng,), True, True)
chart_doc = summ.Charts.getByName("UnitsChart").EmbeddedObject
chart_doc.HasMainTitle = True
chart_doc.Title.String = "Units by Region"

doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(str(out / "rich.xlsx")), (pv("FilterName", "Calc MS Excel 2007 XML"),))
doc.storeToURL(uno.systemPathToFileUrl(str(out / "rich.xlsm")), (pv("FilterName", "Calc MS Excel 2007 VBA XML"),))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print("wrote", out / "rich.xlsx", out / "rich.xlsm")
