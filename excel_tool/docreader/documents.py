"""Word, PowerPoint, PDF, plain text and HTML readers."""
from __future__ import annotations

import csv
import io
import re
from html.parser import HTMLParser
from pathlib import Path

from .cells import classify_text, num_text
from .grid import GCell, Grid, _region, analyze_table, table_to_sql
from .models import ConversionError, ConversionResult, SqlTable
from .util import decode_bytes, guard_zip, ranges_text, sql_ident


def _stem(path: Path) -> str:
    return sql_ident(path.stem, "doc")


def _doc_table_sql(rows: list[list[str]], name: str, source: str) -> list[SqlTable]:
    """Run a Word/PowerPoint table through the same header/total detection as spreadsheets."""
    cells = {}
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            if v.strip():
                kind, raw = classify_text(v, from_csv=True)
                cells[(r, c)] = GCell(v, kind, raw)
    if len(cells) < 2:
        return []
    grid = Grid(name, cells, source="csv")
    tv = analyze_table(grid, _region(set(cells)))
    return table_to_sql(grid, tv, sql_ident(name, "table"), source)


# ---------------------------------------------------------------- word

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_NS = {"w": _W, "mc": _MC,
       "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
       "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


def _xp(expr: str):
    """Compiled XPath with our namespaces (python-docx's .xpath() doesn't know mc:)."""
    from lxml import etree
    return etree.XPath(expr, namespaces=_NS)


_X_TEXT = _xp(".//w:t | .//w:tab | .//w:br | .//w:cr")
_X_SKIP = _xp("ancestor::w:txbxContent | ancestor::mc:Fallback")
_X_TXBX = _xp(".//w:txbxContent[not(ancestor::mc:Fallback)]")
_X_PARAS = _xp(".//w:p")
_X_TRACKED = _xp(".//w:ins | .//w:del")
_X_PICS = _xp(".//pic:pic[not(ancestor::mc:Fallback)]")
_X_CHARTS = _xp(".//c:chart[not(ancestor::mc:Fallback)]")


def _para_text(p_el) -> str:
    """Paragraph text including tracked insertions, content controls and smart tags
    (python-docx's .text skips those). Deleted text (w:delText) is left out."""
    parts = []
    for el in _X_TEXT(p_el):
        if _X_SKIP(el):
            continue
        tag = el.tag.split("}")[1]
        if tag == "t":
            parts.append(el.text or "")
        elif tag == "tab":
            parts.append("\t")
        else:
            parts.append("\n")
    return "".join(parts)


def convert_docx(path: Path) -> ConversionResult:
    import docx
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph

    guard_zip(path)
    try:
        d = docx.Document(str(path))
    except Exception as e:  # noqa: BLE001
        raise ConversionError(f"Could not read {path.name} as a Word document: {e}") from e

    out, tables, warns = [], [], []
    for block in d.iter_inner_content():
        if isinstance(block, Paragraph):
            text = _para_text(block._p).strip()
            if not text:
                continue
            style = block.style.name if block.style is not None else ""
            m = re.match(r"Heading (\d)", style or "")
            if m:
                out.append("#" * min(int(m.group(1)) + 1, 6) + " " + text)
            elif style == "Title":
                out.append("# " + text)
            elif "List" in (style or "") or block._p.pPr is not None and block._p.pPr.numPr is not None:
                out.append("- " + text)
            else:
                out.append(text)
        elif isinstance(block, DocxTable):
            rows = [[_para_text_cell(c) for c in row.cells] for row in block.rows]
            rows = [r for r in rows if any(c.strip() for c in r)]
            if rows:
                out.append(_md_table(rows))
                tables += _doc_table_sql(rows, f"{_stem(path)}_table_{len(tables) + 1}", f"Word table {len(tables) + 1}")

    body = d.element.body
    boxes = []
    for tb in _X_TXBX(body):
        t = " ".join(_para_text(p).strip() for p in _X_PARAS(tb)).strip()
        if t and t not in boxes:
            boxes.append(t)
    if boxes:
        out.append("#### Text boxes\n" + "\n".join(f"- {b}" for b in boxes))

    if _X_TRACKED(body):
        warns.append("Document has tracked changes; text is shown as if all changes were accepted.")
    n_pic = len(_X_PICS(body))
    n_chart = len(_X_CHARTS(body))
    if n_pic or n_chart:
        warns.append(f"{n_pic} picture(s) and {n_chart} chart(s) were not included (text only). "
                     "If they matter, convert the file to PDF and send it inline instead.")
    if any(_section_has_text(s) for s in d.sections):
        warns.append("Headers/footers were not included.")
    return ConversionResult(path.name, "document", "\n\n".join(out), warnings=warns, tables=tables)




def _para_text_cell(cell) -> str:
    return " ".join(t for t in (_para_text(p._p).strip() for p in cell.paragraphs) if t)


def _section_has_text(section) -> bool:
    try:
        parts = section.header.paragraphs + section.footer.paragraphs
        return any(p.text.strip() for p in parts)
    except Exception:  # noqa: BLE001 - malformed header parts shouldn't fail the whole file
        return False


def _md_table(rows: list[list[str]]) -> str:
    width = max(len(r) for r in rows)
    rows = [[c.replace("|", "\\|").replace("\n", " ").strip() for c in r] + [""] * (width - len(r))
            for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


# ---------------------------------------------------------------- powerpoint

def convert_pptx(path: Path) -> ConversionResult:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    guard_zip(path)
    try:
        prs = Presentation(str(path))
    except Exception as e:  # noqa: BLE001
        raise ConversionError(f"Could not read {path.name} as a PowerPoint file: {e}") from e

    out, tables, warns = [], [], []
    pictures = 0

    def walk(shapes):
        for s in sorted(shapes, key=lambda s: ((s.top or 0), (s.left or 0))):
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from walk(s.shapes)
            else:
                yield s

    for i, slide in enumerate(prs.slides, 1):
        title_shape = slide.shapes.title
        title = title_shape.text_frame.text.strip() if title_shape is not None else ""
        lines = [f"## Slide {i}" + (f": {title.replace(chr(11), ' ')}" if title else "")]
        for s in walk(slide.shapes):
            if title_shape is not None and s.shape_id == title_shape.shape_id:
                continue
            if s.has_text_frame:
                for p in s.text_frame.paragraphs:
                    t = p.text.replace("\v", " ").strip()
                    if t:
                        lines.append("  " * p.level + "- " + t)
            if getattr(s, "has_table", False) and s.has_table:
                rows = [[c.text.replace("\v", " ").strip() for c in r.cells] for r in s.table.rows]
                if rows:
                    lines.append(_md_table(rows))
                    tables += _doc_table_sql(rows, f"{_stem(path)}_slide{i}_table", f"slide {i} table")
            if getattr(s, "has_chart", False) and s.has_chart:
                lines.append(_chart_text(s.chart))
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ""
            if notes:
                lines.append(f"Speaker notes: {notes}")
        out.append("\n".join(lines))

    if pictures:
        warns.append(f"{pictures} picture(s) were not included (text only).")
    return ConversionResult(path.name, "slides", "\n\n".join(out), warnings=warns, tables=tables,
                            meta={"slides": len(prs.slides)})


def _chart_text(chart) -> str:
    try:
        title = chart.chart_title.text_frame.text.strip() if chart.has_title else ""
        plot = chart.plots[0]
        cats = [str(c) for c in plot.categories]
        series = list(plot.series)
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\n")
        w.writerow(["category", *[s.name for s in series]])
        for j, cat in enumerate(cats):
            w.writerow([cat, *[num_text(s.values[j]) if j < len(s.values) and s.values[j] is not None else ""
                               for s in series]])
        return f"Chart data{': ' + title if title else ''}\n{buf.getvalue().rstrip()}"
    except Exception:  # noqa: BLE001 - chart XML varies a lot
        return "[Chart present; data could not be extracted]"


# ---------------------------------------------------------------- pdf

def convert_pdf(path: Path) -> ConversionResult:
    import pdfplumber

    out, warns, scanned = [], [], []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                t = (page.extract_text() or "").strip()
                if len(t) < 25:
                    scanned.append(i)
                out.append(f"## Page {i}\n{t}" if t else f"## Page {i}\n[no text layer]")
            n_pages = len(pdf.pages)
    except Exception as e:  # noqa: BLE001
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            raise ConversionError(f"{path.name} is password-protected.") from e
        raise ConversionError(f"Could not read {path.name} as a PDF: {e}") from e

    if scanned:
        pages = ranges_text(scanned)
        warns.append(f"Page(s) {pages} have little or no text layer (likely scanned or image-only); "
                     "they need OCR, e.g. Azure Document Intelligence, or send the PDF inline.")
    warns.append("Tables and charts in PDFs are extracted as plain text; layout may be lost.")
    return ConversionResult(path.name, "pdf", "\n\n".join(out), warnings=warns,
                            meta={"pages": n_pages, "scanned_pages": scanned})


# ---------------------------------------------------------------- text / html

def convert_text(path: Path) -> ConversionResult:
    text, enc = decode_bytes(path.read_bytes())
    warns = [] if enc in ("utf-8", "utf-8-sig", "ascii") else [f"File was not UTF-8; decoded as {enc}."]
    return ConversionResult(path.name, "text", text.replace("\r\n", "\n"), warnings=warns)


class _HTMLText(HTMLParser):
    BLOCK = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "table", "section"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")
        elif tag in ("td", "th"):
            self.parts.append(" | ")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript") and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def convert_html(path: Path) -> ConversionResult:
    raw, enc = decode_bytes(path.read_bytes())
    p = _HTMLText()
    p.feed(raw)
    text = re.sub(r"[ \t]+", " ", "".join(p.parts))
    text = re.sub(r"\n\s*\n+", "\n\n", text).strip()
    return ConversionResult(path.name, "text", text, warnings=["HTML converted to plain text; links and images dropped."])


