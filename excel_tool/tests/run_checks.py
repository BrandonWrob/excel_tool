"""End-to-end checks: layouts are recognized, nothing is lost, SQL answers match, writes are refused."""
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from docreader import ConversionError, ReadOnlyDB, build_sqlite, convert, to_prompt_block  # noqa: E402

S = Path(sys.argv[1]).resolve()
failures = []


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        failures.append(msg)


def q(db, sql):
    return db.query(sql).splitlines()


tmp = Path(tempfile.mkdtemp())

print("zoo.xlsx")
r = convert(S / "zoo/zoo.xlsx")
t = r.text
db_path = build_sqlite(r, tmp / "zoo.sqlite")
db = ReadOnlyDB(db_path)
names = {x.name for x in r.tables}
check("Title: Monthly Shipments - East" in t and "Title: Monthly Shipments - West" in t, "stacked tables keep their titles")
check("7 (total),Total,400" in t, "stacked totals row marked")
check("A16: * Revenue in USD" in t, "footnote attached to the table above it")
check({"sidebyside_3", "sidebyside_4"} <= names, "side-by-side tables split into two")
check("B: 2025 / Q1" in t and "I: 2026 / Q4" in t, "two-level merged header combined")
check(q(db, "SELECT SUM(value) FROM forecast_long WHERE period LIKE '2026%' AND _row_type='data'")[1] == "2000.0",
      "crosstab long format answers per-year sum (2026 = 2000)")
check("- Batch Number: ACME-0042" in t and "- Expiry Date: 2028-10-01" in t, "form read as label: value pairs")
check("batchrecord_7" in names and "10,Water %,<=1.0,1.3,No" in t, "table under the form kept separately")
check(q(db, "SELECT _group, SUM(units) FROM grouped WHERE _row_type='data' GROUP BY _group ORDER BY _group")[1:]
      == ["East,200.0", "West,275.0"], "group labels flow into _group; subtotals excluded")
check("Column B: 2 number(s) stored as text" in t and "'TBD' (B4)" in t and "F3 #N/A" in t,
      "dirty data flagged (numbers as text, stray text, errors)")
check("Repeated header name(s) 'Qty'" in t and "No header text for column(s) D" in t, "header problems flagged")
check("00123" in t and q(db, "SELECT item_id FROM dirtydata WHERE _row=2")[1] == "00123", "leading zeros kept")
check("Row(s) 4, 7 are hidden" in t and "filter" in t, "filtered/hidden rows flagged, still included")
check(q(db, "SELECT COUNT(*) FROM filtered")[1] == "8", "hidden rows are in SQL too")
check("5 (total),Total,14,140" in t, "Excel table totals row marked")
check("J Notes=Needs QA sign-off" in t and t.count("Needs QA sign-off") == 1, "sparse table, no duplicate cells")
check("Contact supply-chain@example.com" in t, "free text notes kept")

print("read-only enforcement")
for bad in ("DROP TABLE grouped", "DELETE FROM grouped", "UPDATE grouped SET units = 0",
            "INSERT INTO grouped(product) VALUES ('x')", "ATTACH DATABASE '/tmp/x.db' AS x",
            "PRAGMA writable_schema = 1", "SELECT 1; DROP TABLE grouped",
            "WITH x AS (SELECT 1) DELETE FROM grouped", "CREATE TABLE evil(a)"):
    out = db.query(bad)
    check(out.startswith("Error"), f"refused: {bad[:45]}")
check(q(db, "SELECT COUNT(*) FROM grouped")[1] == "8", "data unchanged after write attempts")
t0 = time.monotonic()
slow = db.query("WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM c) SELECT COUNT(*) FROM c")
check(slow.startswith("Error") and time.monotonic() - t0 < 10, "runaway query stopped by timeout")
db.close()

print("rich.xlsx (pivot, chart, formatting)")
r = convert(S / "rich.xlsx")
check("9 (total),,,,,Total" in r.text, "totals row under an Excel table attached to it")
check("Pivot table 'RevenuePivot'" in r.text and "Sum of Revenue" in r.text, "pivot definition")
check("Column D ('Product / Total Result') totals the other columns" in r.text, "pivot total column flagged")
check("currently matching: G3, G5" in r.text, "conditional formatting evaluated")
check("data saved with the chart" in r.text, "chart data")

print("CSV")
r = convert(S / "zoo/report_export.csv")
check("Title: Shipments Report" in r.text or "Shipments Report" in r.text, "CSV title kept")
check("(total),Total,655" in r.text, "CSV totals row marked")
check(len(r.tables) == 2, "two CSV tables separated by a blank line")
r = convert(S / "zoo/eu_prices.csv")
db = ReadOnlyDB(build_sqlite(r, tmp / "eu.sqlite"))
check(any("decimal comma" in w for w in r.warnings), "decimal comma detected")
check(q(db, f"SELECT SUM(preis) FROM {r.tables[0].name} WHERE _row_type='data'")[1] == "1277.25",
      "decimal-comma values summed correctly (1277.25), 'Summe' row excluded")
check("00093-7180-01" in r.text, "NDC codes intact")
db.close()

print("earlier samples")
for f, expect in (("messy.xlsx", "00093-7180-01"), ("messy.xls", "Northeast"), ("export_cp1252.csv", "Café"),
                  ("bom.tsv", "02134"), ("tricky.csv", "ok; fine; good"), ("report.docx", "stable for 24 months"),
                  ("deck.pptx", "Speaker notes"), ("report.pdf", "Stability Study"), ("rich.xlsm", "RevenuePivot"),
                  ("macro.xlsm", "VBA macros")):
    try:
        r = convert(S / f)
        check(expect in r.text, f"{f}: contains {expect!r}")
    except ConversionError as e:
        check(False, f"{f}: {e}")
r = convert(S / "messy.xlsx", recalc_formulas=True)
check("1275" in r.text, "recalc fills missing formula results")
try:
    convert(S / "renamed.xlsx")
    check(False, "encrypted file rejected")
except ConversionError:
    check(True, "encrypted/renamed file rejected with a clear message")

print("big data")
t0 = time.monotonic()
r = convert(S / "big.xlsx")
dt = time.monotonic() - t0
db = ReadOnlyDB(build_sqlite(r, tmp / "big.sqlite"))
check(r.strategy == "sql" and r.est_tokens < 20_000, f"30k rows summarized to ~{r.est_tokens:,} tokens ({dt:.1f}s)")
check(q(db, "SELECT COUNT(*) FROM big")[1] == "30000" if "big" in {x.name for x in r.tables}
      else q(db, f"SELECT COUNT(*) FROM {r.tables[0].name}")[1] == "30000", "all 30,000 rows in SQL")
check("Column summary" in r.text and "sum" in r.text, "summary has column statistics")
db.close()

block = to_prompt_block(convert(S / "zoo/zoo.xlsx"))
check(block.count("<content>") == 1 and "SQL tables" in block, "prompt block has content and SQL catalog")

print(f"\n{len(failures)} failure(s)")
sys.exit(1 if failures else 0)
