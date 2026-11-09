"""Excel reader: builds a Grid per sheet (values as Excel displays them, formulas kept,
merged cells resolved) and lists what a value grid alone would lose."""
from __future__ import annotations

import csv
import io
import re
import warnings as _warnings
import zipfile
from pathlib import Path

from .cells import classify_excel, num_text
from .grid import Block, Explicit, GCell, Grid, render_grid
from .models import ConversionError, ConversionResult, SqlTable, estimate_tokens
from .util import col_letter, guard_zip, ranges_text, via_libreoffice

MAX_MERGE_FILL_CELLS = 10_000

# ---------------------------------------------------------------- spreadsheet structure
# Everything a cell grid of cached values loses: formulas, pivot tables, charts,
# conditional formatting, highlights, comments, dropdowns, named ranges, tables, macros.

MAX_FORMULA_LINES = 150
MAX_LISTED_CELLS = 40

_QUOTED_SHEET = re.compile(r"'(?:[^']|'')+'!")
_STRING_LIT = re.compile(r'"(?:[^"]|"")*"')
_CELL_REF = re.compile(r"(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_(!])")


def _formula_text(v) -> str | None:
    if isinstance(v, str):
        return v if v.startswith("=") else None
    text = getattr(v, "text", None)  # ArrayFormula / DataTableFormula
    return str(text) if text else None


def _formula_pattern(formula: str, row: int, col: int) -> str:
    """Formula with relative references turned into offsets, so copies of the same formula
    down a column (=D2*E2, =D3*E3, ...) get the same key."""
    from openpyxl.utils import column_index_from_string

    protected: list[str] = []

    def protect(m):
        protected.append(m.group(0))
        return f"\x00{len(protected) - 1}\x00"

    s = _STRING_LIT.sub(protect, _QUOTED_SHEET.sub(protect, formula))

    def offset(m):
        cabs, letters, rabs, digits = m.groups()
        c, r = column_index_from_string(letters), int(digits)
        return (f"R{r}" if rabs else f"R[{r - row}]") + (f"C{c}" if cabs else f"C[{c - col}]")

    s = _CELL_REF.sub(offset, s)
    return re.sub(r"\x00(\d+)\x00", lambda m: protected[int(m.group(1))], s)


def _formula_lines(wf) -> list[str]:
    
    formulas = {k: f for k, c in wf._cells.items() if (f := _formula_text(c.value))}
    if not formulas:
        return []
    groups: dict[str, list[tuple[int, int]]] = {}
    for (r, c), f in formulas.items():
        groups.setdefault(_formula_pattern(f, r, c), []).append((r, c))

    blocks = []  # (min_row, min_col, max_row, max_col)
    for cells in groups.values():
        by_col: dict[int, list[int]] = {}
        for r, c in cells:
            by_col.setdefault(c, []).append(r)
        runs = []
        for c, rows in by_col.items():
            rows.sort()
            start = prev = rows[0]
            for r in rows[1:] + [None]:
                if r is not None and r == prev + 1:
                    prev = r
                    continue
                runs.append((start, prev, c))
                if r is not None:
                    start = prev = r
        runs.sort(key=lambda t: (t[0], t[1], t[2]))
        merged: list[list[int]] = []  # join side-by-side runs covering the same rows
        for start, end, c in runs:
            if merged and merged[-1][0] == start and merged[-1][2] == end and merged[-1][3] == c - 1:
                merged[-1][3] = c
            else:
                merged.append([start, c, end, c])
        blocks += [tuple(b) for b in merged]

    blocks.sort()
    lines = []
    for r0, c0, r1, c1 in blocks[:MAX_FORMULA_LINES]:
        ref = f"{col_letter(c0)}{r0}"
        if (r0, c0) != (r1, c1):
            ref += f":{col_letter(c1)}{r1}"
        lines.append(f"- {ref}: {formulas[(r0, c0)]}")
    if len(blocks) > MAX_FORMULA_LINES:
        lines.append(f"- ... {len(blocks) - MAX_FORMULA_LINES} more formula groups not listed")
    return lines


_PIVOT_FUNCS = {"sum": "Sum", "count": "Count", "average": "Average", "max": "Max", "min": "Min",
                "product": "Product", "countNums": "Count numbers", "stdDev": "StdDev",
                "stdDevp": "StdDevP", "var": "Var", "varp": "VarP"}


def _describe_pivot(p) -> str:
    names = []
    try:
        names = [f.name for f in p.cache.cacheFields]
    except Exception:  # noqa: BLE001
        pass

    def fname(i):
        if i == -2:
            return "(Values)"
        return names[i] if 0 <= i < len(names) else f"field {i}"

    src = "unknown source"
    try:
        ws_src = p.cache.cacheSource.worksheetSource
        if ws_src is not None:
            src = ws_src.name or (f"{ws_src.sheet}!{ws_src.ref}" if ws_src.sheet else ws_src.ref)
        else:
            src = f"{p.cache.cacheSource.type} source"
    except Exception:  # noqa: BLE001
        pass
    rows = ", ".join(fname(f.x) for f in (p.rowFields or [])) or "none"
    cols = ", ".join(fname(f.x) for f in (p.colFields or [])) or "none"
    filt = ", ".join(fname(f.fld) for f in (p.pageFields or [])) or "none"
    vals = ", ".join(f"{_PIVOT_FUNCS.get(d.subtotal or 'sum', d.subtotal)} of {fname(d.fld)}"
                     for d in (p.dataFields or [])) or "none"
    return (f"- Pivot table '{p.name}' at {p.location.ref}: source {src}; rows: {rows}; "
            f"columns: {cols}; values: {vals}; filters: {filt}. "
            "The cells it covers show its output as of the last refresh.")


_CHART_KINDS = {"lineChart": "line", "line3DChart": "3D line", "pieChart": "pie", "pie3DChart": "3D pie",
                "doughnutChart": "doughnut", "areaChart": "area", "area3DChart": "3D area",
                "scatterChart": "scatter", "radarChart": "radar", "bubbleChart": "bubble",
                "stockChart": "stock", "surfaceChart": "surface", "ofPieChart": "pie-of-pie"}


def _chart_title(ch) -> str:
    t = getattr(ch, "title", None)
    tx = getattr(t, "tx", None)
    if tx is None:
        return ""
    if tx.rich is not None:
        return "".join(r.t or "" for p in tx.rich.p for r in (p.r or [])).strip()
    if tx.strRef is not None:
        return tx.strRef.f or ""
    return ""


def _ref_and_cache(ds):
    """(reference formula, cached point values) for a chart data source."""
    if ds is None:
        return None, []
    for ref in (getattr(ds, "numRef", None), getattr(ds, "strRef", None)):
        if ref is not None:
            cache = getattr(ref, "numCache", None) or getattr(ref, "strCache", None)
            pts = {p.idx: p.v for p in (cache.pt if cache is not None else [])}
            return ref.f, [pts.get(i) for i in range(max(pts) + 1)] if pts else []
    lit = getattr(ds, "numLit", None) or getattr(ds, "strLit", None)
    if lit is not None:
        pts = {p.idx: p.v for p in lit.pt}
        return "(literal values)", [pts.get(i) for i in range(max(pts) + 1)] if pts else []
    return None, []


def _describe_chart(ch) -> str:
    
    tag = ch.tagname
    if tag in ("barChart", "bar3DChart"):
        kind = ("column" if getattr(ch, "type", "col") == "col" else "bar") + (" (3D)" if "3D" in tag else "")
    else:
        kind = _CHART_KINDS.get(tag, tag)
    title = _chart_title(ch)
    where = ""
    try:
        frm = ch.anchor._from
        where = f" placed at {col_letter(frm.col + 1)}{frm.row + 1}"
    except Exception:  # noqa: BLE001
        pass
    lines = [f"- {kind.capitalize()} chart{(' ' + repr(title)) if title else ''}{where}:"]
    table_cols, cats = [], []
    for s in ch.series:
        name = ""
        if s.tx is not None:
            if s.tx.strRef is not None:
                cache = s.tx.strRef.strCache
                name = cache.pt[0].v if cache is not None and cache.pt else (s.tx.strRef.f or "")
            elif s.tx.v:
                name = s.tx.v
        if tag in ("scatterChart", "bubbleChart"):
            cat_ref, cat_vals = _ref_and_cache(getattr(s, "xVal", None))
            val_ref, vals = _ref_and_cache(getattr(s, "yVal", None))
        else:
            cat_ref, cat_vals = _ref_and_cache(getattr(s, "cat", None))
            val_ref, vals = _ref_and_cache(getattr(s, "val", None))
        lines.append(f"  - series {name or '(unnamed)'!r}: values {val_ref or '?'}"
                     + (f", categories {cat_ref}" if cat_ref else ""))
        table_cols.append((name or f"series {len(table_cols) + 1}", vals))
        cats = cats or cat_vals
    n = max([len(v) for _, v in table_cols] + [len(cats), 0])
    if 0 < n <= 50:
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\n")
        w.writerow(["category", *[c for c, _ in table_cols]])
        for i in range(n):
            w.writerow([cats[i] if i < len(cats) and cats[i] is not None else i + 1,
                        *[num_text(float(v[i])) if i < len(v) and _is_number(v[i]) else (v[i] if i < len(v) else "")
                          for _, v in table_cols]])
        lines.append("  - data saved with the chart:\n" + "\n".join("    " + l for l in buf.getvalue().strip().split("\n")))
    return "\n".join(lines)


def _is_number(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


_COLOR_NAMES = {
    "red": (255, 0, 0), "light red": (255, 199, 206), "dark red": (156, 0, 6), "orange": (255, 192, 0),
    "light orange": (252, 228, 214), "yellow": (255, 255, 0), "light yellow": (255, 235, 156),
    "green": (0, 176, 80), "light green": (198, 239, 206), "dark green": (0, 97, 0),
    "blue": (0, 112, 192), "light blue": (189, 215, 238), "dark blue": (31, 56, 100),
    "purple": (112, 48, 160), "gray": (166, 166, 166), "light gray": (217, 217, 217),
    "white": (255, 255, 255), "black": (0, 0, 0),
}


def _color_name(color) -> str | None:
    """Readable name for an openpyxl Color, or None for automatic/no color."""
    if color is None:
        return None
    try:
        ctype = color.type
        if ctype == "rgb":
            if str(color.rgb) == "00000000":  # openpyxl's "not set"
                return None
            rgb = str(color.rgb)[-6:]
        elif ctype == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            if color.indexed in (64, 65):  # system foreground/background
                return None
            rgb = COLOR_INDEX[color.indexed][-6:]
        elif ctype == "theme":
            tint = f", tint {color.tint:+.2f}" if color.tint else ""
            return f"theme color {color.theme}{tint}"
        else:
            return None
        r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:  # noqa: BLE001
        return None
    name = min(_COLOR_NAMES, key=lambda k: sum((a - b) ** 2 for a, b in zip(_COLOR_NAMES[k], (r, g, b))))
    return f"{name} (#{rgb.upper()})"


def _fill_name(fill, *, differential: bool = False) -> str | None:
    """Fill color name. Conditional-format (differential) fills usually omit the pattern type
    and keep the color in bgColor, so they're read differently from normal cell fills."""
    if fill is None:
        return None
    if differential:
        name = _color_name(fill.bgColor) or _color_name(fill.fgColor)
    elif fill.fill_type in (None, "none"):
        return None
    else:
        name = (_color_name(fill.fgColor) if fill.fill_type == "solid" else None) or _color_name(fill.bgColor)
    return None if name and name.startswith("white") else name


def _cell_list(coords: list[str]) -> str:
    shown = ", ".join(coords[:MAX_LISTED_CELLS])
    return shown + (f" ... (+{len(coords) - MAX_LISTED_CELLS} more)" if len(coords) > MAX_LISTED_CELLS else "")


_CF_OPS = {"lessThan": "<", "lessThanOrEqual": "<=", "greaterThan": ">", "greaterThanOrEqual": ">=",
           "equal": "=", "notEqual": "<>", "between": "between", "notBetween": "not between"}


def _describe_cf(ws) -> list[str]:
    
    lines = []
    for cf in ws.conditional_formatting:
        ranges = str(cf.sqref)
        for rule in cf.rules:
            t = rule.type
            if t == "cellIs":
                op = _CF_OPS.get(rule.operator, rule.operator)
                args = " and ".join(f.lstrip("=") for f in (rule.formula or []))
                cond = f"value {op} {args}"
            elif t == "expression":
                cond = f"formula ={(rule.formula or ['?'])[0].lstrip('=')} is true (relative to the top-left cell)"
            elif t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
                cond = f"text {t} {rule.text!r}"
            elif t == "top10":
                cond = f"{'bottom' if rule.bottom else 'top'} {rule.rank}{'%' if rule.percent else ''}"
            elif t == "aboveAverage":
                cond = "above average" if rule.aboveAverage is not False else "below average"
            elif t in ("colorScale", "dataBar", "iconSet"):
                cond = {"colorScale": "color scale (shade varies with value)",
                        "dataBar": "data bars (bar length varies with value)",
                        "iconSet": "icon set (icon varies with value)"}[t]
            else:
                cond = t
            style = []
            dxf = rule.dxf
            if dxf is not None:
                fn = _fill_name(dxf.fill, differential=True) if dxf.fill is not None else None
                if fn:
                    style.append(f"{fn} fill")
                if dxf.font is not None and dxf.font.color is not None:
                    cn = _color_name(dxf.font.color)
                    if cn:
                        style.append(f"{cn} text")
                if dxf.font is not None and dxf.font.b:
                    style.append("bold")
            line = f"- {ranges}: {cond}" + (f" -> {', '.join(style)}" if style else "")

            if t == "cellIs" and rule.operator in _CF_OPS:  # list cells that match right now
                try:
                    bounds = [float(f) for f in rule.formula]
                except (TypeError, ValueError):
                    bounds = None
                if bounds:
                    hits = []
                    for rng in cf.sqref.ranges:
                        for r in range(rng.min_row, rng.max_row + 1):
                            for c in range(rng.min_col, rng.max_col + 1):
                                cell = ws._cells.get((r, c))
                                v = cell.value if cell is not None else None
                                if isinstance(v, (int, float)) and not isinstance(v, bool) and _cf_match(rule.operator, v, bounds):
                                    hits.append(f"{col_letter(c)}{r}")
                    line += f"; currently matching: {_cell_list(hits) if hits else 'none'}"
            lines.append(line)
    return lines


def _cf_match(op: str, v: float, b: list[float]) -> bool:
    if op == "between":
        return min(b) <= v <= max(b)
    if op == "notBetween":
        return not min(b) <= v <= max(b)
    x = b[0]
    return {"lessThan": v < x, "lessThanOrEqual": v <= x, "greaterThan": v > x,
            "greaterThanOrEqual": v >= x, "equal": v == x, "notEqual": v != x}[op]


def _describe_highlights(ws, skip_rows: set[int], col_bounds: tuple[int, int]) -> list[str]:
    """Manually colored cells and colored text, grouped by color. Solid blocks are reported as
    one range and fully colored rows as row numbers, so banding doesn't list every cell."""
    
    fills: dict[str, list[tuple[int, int]]] = {}
    fonts: dict[str, list[tuple[int, int]]] = {}
    for (r, c), cell in sorted(ws._cells.items()):
        if r in skip_rows:
            continue
        fn = _fill_name(cell.fill)
        if fn:
            fills.setdefault(fn, []).append((r, c))
        if cell.value is not None and cell.font is not None and cell.font.color is not None \
                and cell.font.color.type == "rgb":
            cn = _color_name(cell.font.color)
            if cn and not cn.startswith("black"):
                fonts.setdefault(cn, []).append((r, c))

    c0, c1 = col_bounds
    lines = []
    for label, groups in (("fill", fills), ("text color", fonts)):
        for name, cells in groups.items():
            rows = [r for r, _ in cells]
            cols = [c for _, c in cells]
            box = (max(rows) - min(rows) + 1) * (max(cols) - min(cols) + 1)
            if len(cells) > 1 and len(cells) >= 0.9 * box:
                ref = f"{col_letter(min(cols))}{min(rows)}:{col_letter(max(cols))}{max(rows)}"
                lines.append(f"- {name} {label}: {ref}")
                continue
            by_row: dict[int, set[int]] = {}
            for r, c in cells:
                by_row.setdefault(r, set()).add(c)
            full = sorted(r for r, cs in by_row.items() if c1 > c0 and set(range(c0, c1 + 1)) <= cs)
            rest = [f"{col_letter(c)}{r}" for r, c in cells if r not in full]
            parts = []
            if full:
                parts.append(f"whole rows {ranges_text(full)}")
            if rest:
                parts.append(_cell_list(rest))
            lines.append(f"- {name} {label}: {'; '.join(parts)}")
    return lines


def _clean_comment(text: str) -> str:
    if text.startswith("[Threaded comment]"):  # Excel's legacy copy of a threaded comment
        text = text.split("Comment:", 1)[-1]
    return " ".join(text.split())


def _sheet_structure(ws, wf, *, skip_rows: set[int], col_bounds: tuple[int, int]) -> str:
    sections = []

    formulas = _formula_lines(wf)
    if formulas:
        sections.append("Formulas (a range means the same formula is copied across it, with "
                        "references shifting for each cell):\n" + "\n".join(formulas))

    pivots = [_describe_pivot(p) for p in getattr(wf, "_pivots", [])]
    if pivots:
        sections.append("Pivot tables:\n" + "\n".join(pivots))

    charts = [_describe_chart(ch) for ch in getattr(wf, "_charts", [])]
    if charts:
        sections.append("Charts:\n" + "\n".join(charts))

    cf = _describe_cf(ws)
    if cf:
        sections.append("Conditional formatting:\n" + "\n".join(cf))

    hl = _describe_highlights(ws, skip_rows, col_bounds)
    if hl:
        sections.append("Highlighted cells:\n" + "\n".join(hl))

    comments = [f"- {c.coordinate}: {_clean_comment(c.comment.text)}"
                for _, c in sorted(ws._cells.items()) if c.comment is not None and c.comment.text]
    if comments:
        sections.append("Comments/notes:\n" + "\n".join(comments))

    dvs = []
    for dv in getattr(ws.data_validations, "dataValidation", []):
        if dv.type == "list" and dv.formula1:
            f = dv.formula1
            choices = f.strip('"') if f.startswith('"') else f"values from {f.lstrip('=')}"
            dvs.append(f"- {dv.sqref}: dropdown ({choices})")
        elif dv.type:
            rng = " and ".join(x for x in (dv.formula1, dv.formula2) if x)
            dvs.append(f"- {dv.sqref}: must be {dv.type} {dv.operator or 'between'} {rng}".rstrip())
    if dvs:
        sections.append("Data validation:\n" + "\n".join(dvs))

    return "\n\n".join(sections)


def _workbook_structure(wb, path: Path) -> list[str]:
    lines = []
    names = []
    try:
        for name, dn in wb.defined_names.items():
            if not name.startswith("_xlnm"):
                names.append(f"- {name} = {dn.attr_text}")
        for ws in wb.worksheets:
            for name, dn in getattr(ws, "defined_names", {}).items():
                if not name.startswith("_xlnm"):
                    names.append(f"- {name} (sheet '{ws.title}') = {dn.attr_text}")
    except Exception:  # noqa: BLE001
        pass
    if names:
        lines.append("Named ranges:\n" + "\n".join(names))
    with zipfile.ZipFile(path) as z:
        if any(n.lower().endswith("vbaproject.bin") for n in z.namelist()):
            lines.append("Macros: the workbook contains VBA macros. They were not run or included.")
    return lines




# ---------------------------------------------------------------- loading

def _has_formulas(path: Path) -> bool:
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                with z.open(name) as f:
                    tail = b""
                    while chunk := f.read(1 << 20):
                        buf = tail + chunk
                        if b"<f>" in buf or b"<f " in buf or b"<x:f" in buf:
                            return True
                        tail = buf[-8:]
    return False


def _sheet_grid(ws, wf, *, fill_merged_cells: bool) -> tuple[Grid, tuple[int, int, int, int] | None]:
    from openpyxl.utils import column_index_from_string, range_boundaries

    used = {k for k, c in ws._cells.items() if c.value is not None}
    used |= {k for k, c in wf._cells.items() if c.value is not None}
    if not used:
        return Grid(ws.title, {}), None

    cols_by_row: dict[int, set[int]] = {}
    for r, c in used:
        cols_by_row.setdefault(r, set()).add(c)
    merged_src: dict[tuple[int, int], tuple[int, int]] = {}
    if fill_merged_cells:
        for rng in ws.merged_cells.ranges:
            area = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
            if area > MAX_MERGE_FILL_CELLS or (rng.min_row, rng.min_col) not in used:
                continue
            if rng.max_col > rng.min_col and not any(
                    not rng.min_col <= col <= rng.max_col for col in cols_by_row.get(rng.min_row, ())):
                continue  # title banner alone on its row: keep it in its first cell only
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        merged_src[(r, c)] = (rng.min_row, rng.min_col)

    cells: dict[tuple[int, int], GCell] = {}
    for key in used | set(merged_src):
        src = merged_src.get(key, key)
        vcell, fcell = ws._cells.get(src), wf._cells.get(src)
        value = vcell.value if vcell is not None else None
        formula = _formula_text(fcell.value) if fcell is not None else None
        if value is None and formula:
            text, kind, raw = formula, "formula", None
        else:
            text, kind, raw = classify_excel(value, vcell.number_format if vcell is not None else None)
        if text != "":
            cells[key] = GCell(text, kind, raw, formula if src == key else None)

    hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
    hidden_cols = set()
    for letter, d in ws.column_dimensions.items():
        if d.hidden:
            lo = d.min or column_index_from_string(letter)
            hi = d.max or lo
            hidden_cols.update(range(lo, hi + 1))

    explicit = []
    for name in list(ws.tables.keys()):
        t = ws.tables[name]
        c0, r0, c1, r1 = range_boundaries(t.ref)
        explicit.append(Explicit("excel_table", name, r0, c0, r1, c1,
                                 header_rows=1 if t.headerRowCount is None else int(t.headerRowCount),
                                 total_rows=int(t.totalsRowCount or 0)))
    for p in getattr(wf, "_pivots", []):
        try:
            c0, r0, c1, r1 = range_boundaries(p.location.ref)
            first_data = int(p.location.firstDataRow or 1)
            explicit.append(Explicit("pivot", p.name, r0, c0, r1, c1, header_rows=first_data))
        except Exception:  # noqa: BLE001 - malformed pivot locations shouldn't stop the sheet
            pass

    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    bounds = (min(rows), min(cols), max(rows), max(cols)) if cells else None
    return Grid(ws.title, cells, "excel", hidden_rows, hidden_cols,
                ws.auto_filter.ref if ws.auto_filter is not None else None, explicit), bounds


def convert_xlsx(path: Path, *, include_hidden_sheets: bool = True, fill_merged_cells: bool = True,
                 recalc_formulas: bool = False, include_structure: bool = True,
                 max_inline_tokens: int = 20_000, table_prefix: str = "") -> ConversionResult:
    import openpyxl

    guard_zip(path)
    opts = dict(include_hidden_sheets=include_hidden_sheets, fill_merged_cells=fill_merged_cells,
                include_structure=include_structure, max_inline_tokens=max_inline_tokens,
                table_prefix=table_prefix)
    if recalc_formulas:
        return via_libreoffice(path, "xlsx", lambda p: convert_xlsx(p, recalc_formulas=False, **opts),
                               note="Formulas were recalculated with LibreOffice before reading.")

    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")  # openpyxl warns about extensions it doesn't support
        try:
            wb_vals = openpyxl.load_workbook(path, data_only=True)    # saved results
            # A second load is only needed to see formulas; skip it when there are none.
            wb_form = openpyxl.load_workbook(path, data_only=False) if _has_formulas(path) else wb_vals
        except Exception as e:  # noqa: BLE001 - openpyxl raises many types
            raise ConversionError(f"Could not read {path.name} as a spreadsheet: {e}") from e

    pieces: list[str | Block] = []
    sql: list[SqlTable] = []
    warns: list[str] = []
    idx = 1
    used_names: set[str] = set()
    uncached = 0

    for ws in wb_vals.worksheets:
        state = ws.sheet_state
        if state != "visible" and not include_hidden_sheets:
            warns.append(f"Skipped {state} sheet '{ws.title}'.")
            continue
        wf = wb_form[ws.title]
        grid, bounds = _sheet_grid(ws, wf, fill_merged_cells=fill_merged_cells)
        if bounds is None:
            pieces.append(f"### Sheet: {ws.title} ({state}) - empty")
            continue
        r0, c0, r1, c1 = bounds
        uncached += sum(1 for x in grid.cells.values() if x.kind == "formula" and x.formula)
        blocks, tables, idx, header_rows = render_grid(grid, sheet_label=ws.title, table_prefix=table_prefix,
                                                       start_index=idx, used_names=used_names)
        sql += tables
        head = f"### Sheet: {ws.title} ({state}) - used range {col_letter(c0)}{r0}:{col_letter(c1)}{r1}"
        if grid.autofilter:
            head += f"; filter on {grid.autofilter}"
        pieces.append(head)
        pieces += blocks
        if include_structure:
            structure = _sheet_structure(ws, wf, skip_rows=header_rows, col_bounds=(c0, c1))
            if structure:
                pieces.append(structure)
        if state != "visible":
            warns.append(f"Sheet '{ws.title}' is {state} in Excel but was included.")
        imgs = len(getattr(ws, "_images", []))
        if imgs:
            warns.append(f"Sheet '{ws.title}' has {imgs} image(s) that were not included.")

    if include_structure:
        wb_struct = _workbook_structure(wb_form, path)
        for cs in getattr(wb_form, "chartsheets", []):
            charts = [_describe_chart(ch) for ch in getattr(cs, "_charts", [])]
            if charts:
                wb_struct.append(f"Chart sheet '{cs.title}':\n" + "\n".join(charts))
        if wb_struct:
            pieces.append("### Workbook\n" + "\n".join(wb_struct))
    if uncached:
        warns.append(f"{uncached} formula cell(s) had no saved result (the file was probably written by "
                     "software other than Excel); their formula text is shown. "
                     "Convert with recalc_formulas=True to calculate them.")
    return assemble(path.name, "spreadsheet", pieces, sql, warns, max_inline_tokens,
                    meta={"sheets": len(wb_vals.worksheets)})


def assemble(filename: str, kind: str, pieces: list, sql: list[SqlTable], warns: list[str],
             max_inline_tokens: int, meta: dict | None = None) -> ConversionResult:
    """Join text pieces; if over budget, swap the largest tables for their summaries.
    Every row stays available in `sql` either way."""
    def text_of(p, use_summary):
        if isinstance(p, Block):
            return p.summary if use_summary.get(id(p)) and p.summary else p.text
        return p

    use_summary: dict[int, bool] = {}
    full = "\n\n".join(text_of(p, use_summary) for p in pieces)
    total = estimate_tokens(full)
    if total > max_inline_tokens:
        candidates = sorted((p for p in pieces if isinstance(p, Block) and p.summary),
                            key=lambda b: estimate_tokens(b.text) - estimate_tokens(b.summary), reverse=True)
        for b in candidates:
            if total <= max_inline_tokens:
                break
            use_summary[id(b)] = True
            total -= estimate_tokens(b.text) - estimate_tokens(b.summary)
    text = "\n\n".join(text_of(p, use_summary) for p in pieces)
    strategy = "sql" if use_summary else "inline"
    if use_summary:
        warns.append(f"The file is about {estimate_tokens(full):,} tokens, so {len(use_summary)} large table(s) "
                     "are summarized. All rows are in the SQL tables listed below.")
    return ConversionResult(filename, kind, text, full_text=full, warnings=warns, tables=sql,
                            strategy=strategy, meta=meta or {})
