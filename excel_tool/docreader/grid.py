"""Find and describe what's on a sheet: tables, forms, and notes. Nothing is changed or dropped;
this only decides how to present the cells so a model can read them correctly.

Used for Excel sheets and CSV files alike. Steps:
  1. regions   - Excel tables and pivot tables first, then blocks of connected cells
                 (blank rows/columns separate blocks), with merging rules for spacer columns,
                 blank rows inside a table, and a separated totals row
  2. attach    - one-line titles above a table and short notes/footnotes below it
  3. classify  - table, form (label: value pairs), or text
  4. tables    - title rows, multi-row headers (combined as "2025 / Q1"), total and subtotal
                 rows, group label rows, period columns (crosstabs), column types, data issues
  5. render    - CSV with sheet row numbers and column letters; large tables summarized with
                 column statistics; every row goes into SQL regardless
"""
from __future__ import annotations

import csv
import io
import re
from collections import Counter
from dataclasses import dataclass, field

from .cells import family, num_text
from .models import SqlTable, estimate_tokens
from .util import col_letter, ranges_text, sql_ident, unique_names

MAX_CELLS_FOR_DETECTION = 2_000_000
MAX_COMPONENTS_FOR_MERGING = 400
PROFILE_MIN_ROWS = 25          # tables with more data rows get column statistics
SUMMARY_PREVIEW_ROWS = 15
SPARSE_FILL_RATIO = 0.35       # below this (and wide), rows render as "column=value" records
MAX_LISTED = 8


# ------------------------------------------------------------------ data model

@dataclass
class GCell:
    text: str
    kind: str
    raw: object = None
    formula: str | None = None


@dataclass
class Explicit:
    """A region Excel itself defines: an Excel table (ListObject) or a pivot table."""
    kind: str
    name: str
    r0: int
    c0: int
    r1: int
    c1: int
    header_rows: int = 1
    total_rows: int = 0


@dataclass
class Grid:
    name: str
    cells: dict[tuple[int, int], GCell]
    source: str = "excel"                 # excel | csv
    hidden_rows: set[int] = field(default_factory=set)
    hidden_cols: set[int] = field(default_factory=set)
    autofilter: str | None = None
    explicit: list[Explicit] = field(default_factory=list)

    def get(self, r: int, c: int) -> GCell | None:
        x = self.cells.get((r, c))
        return x if x is not None and x.text != "" else None


@dataclass
class Region:
    r0: int
    c0: int
    r1: int
    c1: int
    cells: set[tuple[int, int]]
    source: str = "detected"              # detected | excel_table | pivot
    name: str = ""
    header_hint: int | None = None
    total_hint: int = 0
    kind: str = "table"                   # table | form | text
    title: str = ""
    title_ref: str = ""
    notes: list[str] = field(default_factory=list)

    @property
    def ref(self) -> str:
        a = f"{col_letter(self.c0)}{self.r0}"
        b = f"{col_letter(self.c1)}{self.r1}"
        return a if a == b else f"{a}:{b}"

    @property
    def height(self) -> int:
        return self.r1 - self.r0 + 1

    @property
    def width(self) -> int:
        return self.c1 - self.c0 + 1


def _region(cells: set[tuple[int, int]], **kw) -> Region:
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    return Region(min(rows), min(cols), max(rows), max(cols), set(cells), **kw)


# ------------------------------------------------------------------ 1-3. regions

def find_regions(grid: Grid) -> list[Region]:
    occupied = {k for k, c in grid.cells.items() if c.text != ""}
    if not occupied:
        return []
    if len(occupied) > MAX_CELLS_FOR_DETECTION:
        return [_region(occupied)]

    regions, claimed = [], set()
    for ex in grid.explicit:
        cells = {k for k in occupied if ex.r0 <= k[0] <= ex.r1 and ex.c0 <= k[1] <= ex.c1} - claimed
        if cells:
            regions.append(Region(ex.r0, ex.c0, ex.r1, ex.c1, cells, source=ex.kind, name=ex.name,
                                  header_hint=ex.header_rows, total_hint=ex.total_rows))
            claimed |= cells

    comps = [_region(c) for c in _components(occupied - claimed)]
    if len(comps) > MAX_COMPONENTS_FOR_MERGING:
        # Very scattered sheet (checklists, forms laid out freely): keep it as one sparse block.
        sparse = _region(occupied - claimed)
        sparse.kind = "sparse"
        return sorted(regions + [sparse], key=lambda g: (g.r0, g.c0))
    comps = _merge_components(grid, comps)
    for ex in regions:  # totals row(s) typed just under an Excel table belong to it
        for b in list(comps):
            gap = b.r0 - ex.r1 - 1
            if (0 <= gap <= 2 and b.height <= 3 and b.c0 >= ex.c0 and b.c1 <= ex.c1
                    and all(_row_has_total_marker(grid, r, b.c0, b.c1) for r in {r for r, _ in b.cells})):
                ex.cells |= b.cells
                ex.r1 = b.r1
                comps.remove(b)
    regions = _attach_titles_and_notes(grid, regions + comps)
    for reg in regions:
        if reg.source == "detected":
            reg.kind = _classify(grid, reg)
    return sorted(regions, key=lambda g: (g.r0, g.c0))


def _components(cells: set[tuple[int, int]]) -> list[set[tuple[int, int]]]:
    """Groups of cells touching each other, diagonals included."""
    remaining = set(cells)
    out = []
    while remaining:
        start = remaining.pop()
        comp, stack = {start}, [start]
        while stack:
            r, c = stack.pop()
            for dr in (-1, 0, 1):
                for dc in (-1, 0, 1):
                    n = (r + dr, c + dc)
                    if n in remaining:
                        remaining.remove(n)
                        comp.add(n)
                        stack.append(n)
        out.append(comp)
    return out


TOTAL_RE = re.compile(r"^\s*(grand\s+total|sub-?\s?totals?|totals?|sum|overall|summe|zwischensumme|gesamt\w*|"
                      r"somme|totale|totaal|suma|totalt)\b", re.I)
_VSUM_RE = re.compile(r"^=\s*(SUM|SUBTOTAL|AGGREGATE)\((?:\d+\s*,\s*)?\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)", re.I)


def _is_vertical_total(cell: GCell | None, r: int, c: int) -> bool:
    if cell is None or not cell.formula:
        return False
    m = _VSUM_RE.match(cell.formula)
    return bool(m) and m.group(2).upper() == m.group(4).upper() == col_letter(c) and int(m.group(5)) < r


def _row_has_total_marker(grid: Grid, r: int, c0: int, c1: int) -> bool:
    cells = [(c, grid.get(r, c)) for c in range(c0, c1 + 1)]
    texts = [x.text for _, x in cells if x is not None and family(x.kind) == "text"][:2]
    return any(TOTAL_RE.match(t) for t in texts) or any(_is_vertical_total(x, r, c) for c, x in cells)


def _row_families(grid: Grid, r: int, c0: int, c1: int) -> dict[int, str]:
    return {c: family(x.kind) for c in range(c0, c1 + 1) if (x := grid.get(r, c)) is not None}


def _should_merge(grid: Grid, a: Region, b: Region) -> bool:
    if a.source != "detected" or b.source != "detected":
        return False
    # Overlapping ranges (e.g. a lone note inside a sparse table's Notes column) are one block.
    if a.r0 <= b.r1 and b.r0 <= a.r1 and a.c0 <= b.c1 and b.c0 <= a.c1:
        return True
    # Extra columns past a single spacer column, starting on the same header row and of
    # similar length (a spacer column inside one table, not two tables side by side).
    if 0 <= b.c0 - a.c1 - 1 <= 1 and a.r0 == b.r0 and min(a.height, b.height) >= 0.5 * max(a.height, b.height):
        return True
    # Below a, within its columns, after at most two blank rows.
    gap = b.r0 - a.r1 - 1
    if 0 <= gap <= 2 and b.c0 >= a.c0 and b.c1 <= a.c1:
        if b.height <= 3 and all(_row_has_total_marker(grid, r, b.c0, b.c1) for r in range(b.r0, b.r1 + 1)
                                 if any(grid.get(r, c) for c in range(b.c0, b.c1 + 1))):
            return True  # separated totals row(s)
        if b.c0 == a.c0 and a.c1 - b.c1 <= 1 and a.height >= 2:
            last = _row_families(grid, a.r1, a.c0, a.c1)
            first = _row_families(grid, b.r0, b.c0, b.c1)
            shared = [c for c in first if c in last]
            same = sum(1 for c in shared if first[c] == last[c])
            typed_first = any(f not in ("text", "empty") for f in first.values())
            return len(shared) >= 2 and same >= 0.7 * len(shared) and typed_first
    return False


def _merge_components(grid: Grid, comps: list[Region]) -> list[Region]:
    changed = True
    while changed:
        changed = False
        comps.sort(key=lambda g: (g.r0, g.c0))
        for i, a in enumerate(comps):
            for j, b in enumerate(comps):
                if i != j and _should_merge(grid, a, b):
                    comps[i] = _region(a.cells | b.cells)
                    del comps[j]
                    changed = True
                    break
            if changed:
                break
    return comps


def _is_title_candidate(grid: Grid, reg: Region) -> bool:
    if reg.source != "detected" or reg.height != 1:
        return False
    xs = [grid.get(r, c) for r, c in reg.cells]
    return len(reg.cells) <= 2 and all(x is not None and family(x.kind) == "text" for x in xs)


def _is_note_candidate(grid: Grid, reg: Region) -> bool:
    if reg.source != "detected" or reg.height > 6 or reg.width > 3:
        return False
    xs = [grid.get(r, c) for r, c in reg.cells]
    return all(x is not None and family(x.kind) == "text" for x in xs)


def _attach_titles_and_notes(grid: Grid, regions: list[Region]) -> list[Region]:
    tables = [g for g in regions if g.source != "detected" or (g.height >= 2 and len(g.cells) >= 3)]
    table_ids = {id(g) for g in tables}
    used: set[int] = set()
    for t in sorted(tables, key=lambda g: (g.r0, g.c0)):
        # title: nearest one-line text block just above, roughly over the table
        best = None
        for g in regions:
            if id(g) in used or g is t or not _is_title_candidate(grid, g):
                continue
            gap = t.r0 - g.r1 - 1
            if 0 <= gap <= 2 and t.c0 - 1 <= g.c0 <= t.c1:
                if best is None or g.r1 > best.r1:
                    best = g
        if best is not None:
            used.add(id(best))
            t.title = " ".join(grid.get(r, c).text for r, c in sorted(best.cells))
            t.title_ref = best.ref
        # notes: short text blocks just below
        for g in sorted(regions, key=lambda g: g.r0):
            if id(g) in used or id(g) in table_ids or not _is_note_candidate(grid, g):
                continue
            gap = g.r0 - t.r1 - 1
            if 0 <= gap <= 3 and t.c0 <= g.c0 <= t.c1:
                used.add(id(g))
                for r, c in sorted(g.cells):
                    t.notes.append(f"{col_letter(c)}{r}: {grid.get(r, c).text}")
    return [g for g in regions if id(g) not in used]


def _classify(grid: Grid, reg: Region) -> str:
    xs = [grid.get(r, c) for r, c in reg.cells]
    if len(xs) == 1 or reg.height == 1 and all(family(x.kind) == "text" for x in xs):
        return "text"
    labels = sum(1 for x in xs if family(x.kind) == "text" and x.text.rstrip().endswith(":"))
    if labels >= 2 and labels >= 0.3 * len(xs):
        return "form"
    if reg.width == 2 and reg.height >= 3:
        left = [grid.get(r, reg.c0) for r in range(reg.r0, reg.r1 + 1)]
        right = [grid.get(r, reg.c1) for r in range(reg.r0, reg.r1 + 1)]
        if all(x is None or family(x.kind) == "text" for x in left):
            kinds = {family(x.kind) for x in right if x is not None} - {"empty", "unknown"}
            labels_unique = len({x.text for x in left if x}) == sum(1 for x in left if x)
            if labels_unique and len(kinds) >= 3:
                return "form"
    if reg.width == 1 and all(family(x.kind) == "text" for x in xs):
        return "text"
    return "table"


# ------------------------------------------------------------------ 4. tables

@dataclass
class TableView:
    region: Region
    cols: list[int]
    header_rows: list[int]
    names: list[str]
    body: list[int]
    row_type: dict[int, str]
    group_of: dict[int, str]
    col_family: dict[int, str]
    period_cols: list[int]
    issues: list[str]
    title: str
    col_kind: dict[int, str] = field(default_factory=dict)

    @property
    def data_rows(self) -> list[int]:
        return [r for r in self.body if self.row_type[r] == "data"]


_MONTHS = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"
_PERIOD_RE = re.compile(
    rf"^({_MONTHS}([\s\-'/]*\d{{2,4}})?|q[1-4]([\s\-'/]*(fy)?\d{{2,4}})?|(fy|cy)\s?'?\d{{2,4}}( q[1-4])?|"
    rf"(19|20)\d{{2}}([\s\-/](q[1-4]|\d{{1,2}}|{_MONTHS}))?|h[12]([\s\-]\d{{2,4}})?|\d{{4}}-\d{{2}}(-\d{{2}})?|"
    rf"(week|wk|month|period|p)\s?\d{{1,2}})$", re.I)


def _is_year(x: GCell | None) -> bool:
    return x is not None and isinstance(x.raw, (int, float)) and float(x.raw).is_integer() and 1990 <= x.raw <= 2100


def _detect_header_count(grid: Grid, rows: list[int], cols: list[int]) -> int:
    if len(rows) < 2:
        return 0
    sample = rows[max(1, len(rows) // 3):]
    typed: dict[int, str] = {}
    for c in cols:
        fams = Counter(family(x.kind) for r in sample if (x := grid.get(r, c)) is not None)
        fams.pop("unknown", None)
        fams.pop("error", None)
        if fams:
            fam, n = fams.most_common(1)[0]
            if fam in ("number", "date", "bool") and n >= 0.6 * sum(fams.values()):
                typed[c] = fam
    if not typed:
        return 1  # all-text table: first row is the header
    h = 0
    for i, r in enumerate(rows[:6]):
        present = [(c, x) for c in cols if (x := grid.get(r, c)) is not None]
        in_typed = [(c, x) for c, x in present if c in typed]
        has_text = any(family(x.kind) == "text" for _, x in present)
        if i > 0 and len(present) < 0.5 * len(cols):
            break  # a lone label under the header is a group label, not another header row
        # A row of years over numeric columns ("2025", "2026") is a header, not data.
        next_row = rows[i + 1] if i + 1 < len(rows) else None
        year_header = (bool(in_typed) and all(_is_year(x) for _, x in in_typed) and next_row is not None
                       and not all(_is_year(grid.get(next_row, c)) for c, _ in in_typed))
        looks_like_data = bool(in_typed) and sum(
            1 for c, x in in_typed if family(x.kind) == typed[c]) >= 0.5 * len(in_typed)
        if (looks_like_data or not has_text) and not year_header:
            break
        h += 1
    if h >= len(rows):
        h = 1
    return h


def _combine_headers(grid: Grid, header_rows: list[int], cols: list[int]) -> tuple[list[str], list[str]]:
    levels = []
    for i, r in enumerate(header_rows):
        vals = [(grid.get(r, c).text.strip() if grid.get(r, c) else "") for c in cols]
        if i < len(header_rows) - 1:  # upper levels span the columns to their right
            last = ""
            seen = False
            for j, v in enumerate(vals):
                if v:
                    last, seen = v, True
                elif seen:
                    vals[j] = last
        levels.append(vals)
    issues = []
    names = []
    for j, c in enumerate(cols):
        parts = []
        for lv in levels:
            v = " ".join(lv[j].split())
            if v and v not in parts:
                parts.append(v)
        names.append(" / ".join(parts))
    blanks = [col_letter(c) for c, n in zip(cols, names) if not n]
    if blanks and header_rows:
        issues.append(f"No header text for column(s) {', '.join(blanks)}; named by column letter.")
    names = [n or f"Column {col_letter(c)}" for c, n in zip(cols, names)]
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        issues.append(f"Repeated header name(s) {', '.join(repr(d) for d in dupes)}; "
                      "use the column letter to tell them apart.")
    return names, issues


def analyze_table(grid: Grid, reg: Region) -> TableView:
    rows = sorted({r for r, _ in reg.cells})
    cols = sorted({c for _, c in reg.cells})
    title = reg.title
    issues: list[str] = []

    if reg.source == "detected" and len(rows) >= 3 and len(cols) > 1:
        first = [grid.get(rows[0], c) for c in cols]
        vals = {x.text for x in first if x is not None}
        if len(vals) == 1 and all(x is None or family(x.kind) == "text" for x in first):
            title = (title + " / " if title else "") + vals.pop()
            rows = rows[1:]

    h = reg.header_hint if reg.header_hint is not None else _detect_header_count(grid, rows, cols)
    h = min(h, len(rows))
    header_rows, body = rows[:h], rows[h:]
    names, name_issues = _combine_headers(grid, header_rows, cols)
    issues += name_issues

    col_family: dict[int, str] = {}
    for c in cols:
        fams = Counter(family(x.kind) for r in body if (x := grid.get(r, c)) is not None)
        for k in ("unknown", "error"):
            fams.pop(k, None)
        col_family[c] = fams.most_common(1)[0][0] if fams else "empty"

    row_type: dict[int, str] = {}
    group_of: dict[int, str] = {}
    current_group = ""
    for r in body:
        present = [(c, grid.get(r, c)) for c in cols if grid.get(r, c) is not None]
        if reg.total_hint and r > reg.r1 - reg.total_hint:
            row_type[r] = "total"
        elif _row_has_total_marker(grid, r, cols[0], cols[-1]):
            row_type[r] = "total"
        elif (len(present) == 1 and len(cols) >= 3 and present[0][0] in cols[:2]
              and family(present[0][1].kind) == "text"
              and any(col_family[c] in ("number", "date") for c in cols if c != present[0][0])):
            row_type[r] = "group"  # a label alone on its row, e.g. "East region" above its rows
        else:
            row_type[r] = "data"
    groups = [r for r in body if row_type[r] == "group"]
    if body and body[-1] in groups:
        row_type[body[-1]] = "data"
        groups.pop()
    if len(groups) > 0.25 * len(body):  # too many: these are rows with missing values, not labels
        for r in groups:
            row_type[r] = "data"
        groups = []
    current_group = ""
    for r in body:
        if row_type[r] == "group":
            current_group = grid.get(r, next(c for c in cols if grid.get(r, c) is not None)).text
        elif current_group and row_type[r] == "data":
            group_of[r] = current_group

    col_kind: dict[int, str] = {}
    for c in cols:
        kinds = Counter(x.kind for r in body if (x := grid.get(r, c)) is not None and family(x.kind) == col_family[c])
        col_kind[c] = kinds.most_common(1)[0][0] if kinds else col_family[c]

    period_cols = []
    leaf = [(grid.get(header_rows[-1], c).text.strip() if header_rows and grid.get(header_rows[-1], c) else "")
            for c in cols]
    for c, label in zip(cols, leaf):
        cell = grid.get(header_rows[-1], c) if header_rows else None
        is_period = bool(label) and (_PERIOD_RE.match(label) or _is_year(cell) or (
            cell is not None and family(cell.kind) == "date"))
        if is_period and col_family[c] == "number":
            period_cols.append(c)
    if len(period_cols) < 3:
        period_cols = []
    total_cols = [c for c, label in zip(cols, leaf) if label and TOTAL_RE.match(label) and col_family[c] == "number"]
    for c in total_cols:
        issues.append(f"Column {col_letter(c)} ({names[cols.index(c)]!r}) totals the other columns; "
                      "leave it out when adding across columns.")
    period_cols = [c for c in period_cols if c not in total_cols]

    data_rows = [r for r in body if row_type[r] == "data"]
    for c in cols:
        cells = [(r, grid.get(r, c)) for r in data_rows]
        cells = [(r, x) for r, x in cells if x is not None]
        letter = col_letter(c)
        errs = [f"{letter}{r} {x.text}" for r, x in cells if x.kind == "error"]
        if errs:
            issues.append(f"Column {letter}: {len(errs)} error value(s): {', '.join(errs[:MAX_LISTED])}"
                          + (" ..." if len(errs) > MAX_LISTED else ""))
        numtext = [f"{letter}{r}" for r, x in cells if x.kind == "numtext"]
        if numtext:
            issues.append(f"Column {letter}: {len(numtext)} number(s) stored as text "
                          f"({', '.join(numtext[:MAX_LISTED])}{' ...' if len(numtext) > MAX_LISTED else ''}); "
                          "treated as numbers in SQL.")
        fam = col_family[c]
        if fam in ("number", "date"):
            odd = [f"{x.text!r} ({letter}{r})" for r, x in cells if family(x.kind) == "text"]
            if odd:
                issues.append(f"Column {letter} is mostly {fam}s but has text: {', '.join(odd[:MAX_LISTED])}"
                              + (" ..." if len(odd) > MAX_LISTED else ""))
        formulas = [f"{letter}{r}" for r, x in cells if x.kind == "formula"]
        if formulas:
            issues.append(f"Column {letter}: {len(formulas)} formula(s) have no saved result "
                          f"({', '.join(formulas[:MAX_LISTED])}).")
    hidden = sorted(r for r in body if r in grid.hidden_rows)
    if hidden:
        why = " (the sheet has a filter, so these may be filtered out)" if grid.autofilter else ""
        issues.append(f"Row(s) {ranges_text(hidden)} are hidden in Excel{why}; they are included here.")
    hidden_c = [col_letter(c) for c in cols if c in grid.hidden_cols]
    if hidden_c:
        issues.append(f"Column(s) {', '.join(hidden_c)} are hidden in Excel; they are included here.")

    return TableView(reg, cols, header_rows, names, body, row_type, group_of, col_family,
                     period_cols, issues, title, col_kind)


# ------------------------------------------------------------------ 5. render

def _row_label(tv: TableView, r: int) -> str:
    t = tv.row_type[r]
    return str(r) if t == "data" else f"{r} ({t})"


def _profile_lines(grid: Grid, tv: TableView) -> list[str]:
    lines = []
    data_rows = tv.data_rows
    for c, name in zip(tv.cols, tv.names):
        xs = [x for r in data_rows if (x := grid.get(r, c)) is not None]
        blank = len(data_rows) - len(xs)
        head = f"- {col_letter(c)} {name!r}:"
        fam = tv.col_family[c]
        if not xs:
            lines.append(f"{head} empty")
            continue
        if fam == "number":
            nums = [float(x.raw) for x in xs if isinstance(x.raw, (int, float)) and not isinstance(x.raw, bool)]
            kinds = Counter(x.kind for x in xs)
            kind = kinds.most_common(1)[0][0]
            fmt = (lambda v: num_text(round(v * 100, 10)) + "%") if kind == "percent" else (lambda v: num_text(round(v, 10)))
            if nums:
                total = f", sum {fmt(sum(nums))}" if kind != "percent" else ""
                desc = f"{kind}, min {fmt(min(nums))}, max {fmt(max(nums))}{total}"
            else:
                desc = kind
        elif fam == "date":
            texts = sorted(x.text for x in xs)
            desc = f"date, {texts[0]} to {texts[-1]}"
        else:
            counts = Counter(x.text for x in xs)
            if len(counts) == len(xs) and len(xs) > 12:
                ordered = sorted(counts)
                desc = f"{fam}, all {len(xs)} values unique ({ordered[0]} ... {ordered[-1]})"
            elif len(counts) <= 12:
                desc = f"{fam}, {len(counts)} distinct: " + ", ".join(f"{v} ({n})" for v, n in counts.most_common())
            else:
                top = ", ".join(f"{v} ({n})" for v, n in counts.most_common(5))
                desc = f"{fam}, {len(counts)} distinct; most common: {top}"
        lines.append(f"{head} {desc}" + (f"; {blank} blank" if blank else ""))
    return lines


def _csv_lines(rows: list[list[str]]) -> str:
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    return buf.getvalue().rstrip("\n")


def render_table(grid: Grid, tv: TableView, *, index: int, sql_name: str | None, summary: bool) -> str:
    reg = tv.region
    kind = {"excel_table": "Excel table", "pivot": "pivot table output", "detected": "table"}[reg.source]
    head = f"#### Table {index}" + (f" '{reg.name}'" if reg.name else "") + f" ({kind}) - range {reg.ref}"
    n_total = sum(1 for t in tv.row_type.values() if t == "total")
    n_group = sum(1 for t in tv.row_type.values() if t == "group")
    bits = [f"{len(tv.data_rows)} data rows"]
    if tv.header_rows:
        bits.insert(0, f"header row{'s' if len(tv.header_rows) > 1 else ''} {ranges_text(tv.header_rows)}")
    if n_total:
        bits.append(f"{n_total} total/subtotal row(s) marked (total)")
    if n_group:
        bits.append(f"{n_group} group label row(s) marked (group)")
    lines = [head + "; " + "; ".join(bits)]
    if tv.title:
        lines.append(f"Title: {tv.title}")
    if reg.source == "pivot":
        lines.append("This is a pivot table's output as of its last refresh; its definition is listed under Pivot tables.")
    if tv.period_cols:
        lines.append(f"Columns {col_letter(tv.period_cols[0])}-{col_letter(tv.period_cols[-1])} are periods "
                     "(a crosstab); the SQL copy also has a long format with one row per period.")

    header = ["row", *[f"{col_letter(c)}: {n}" for c, n in zip(tv.cols, tv.names)]]
    body_rows = tv.body if not summary else tv.body[:SUMMARY_PREVIEW_ROWS]
    filled = sum(1 for r in tv.body for c in tv.cols if grid.get(r, c) is not None)
    density = filled / max(1, len(tv.body) * len(tv.cols))
    if density < SPARSE_FILL_RATIO and len(tv.cols) >= 6:
        lines.append("(Sparse table: each row lists only its filled columns.)")
        for r in body_rows:
            parts = [f"{col_letter(c)} {n}={grid.get(r, c).text}" for c, n in zip(tv.cols, tv.names)
                     if grid.get(r, c) is not None]
            lines.append(f"row {_row_label(tv, r)}: " + "; ".join(parts))
    else:
        rows = [header] + [[_row_label(tv, r), *[(grid.get(r, c).text if grid.get(r, c) else "") for c in tv.cols]]
                           for r in body_rows]
        lines.append(_csv_lines(rows))
    if summary and len(tv.body) > SUMMARY_PREVIEW_ROWS:
        lines.append(f"... {len(tv.body) - SUMMARY_PREVIEW_ROWS} more rows not shown. "
                     f"All rows are in SQL table \"{sql_name}\".")
    if summary or len(tv.data_rows) >= PROFILE_MIN_ROWS:
        lines.append("Column summary (data rows only, totals excluded):")
        lines += _profile_lines(grid, tv)
    if tv.group_of:
        lines.append("Rows under a group label belong to that group (SQL column _group).")
    if tv.issues:
        lines.append("Data notes:")
        lines += [f"- {i}" for i in tv.issues]
    if reg.notes:
        lines.append("Notes below the table:")
        lines += [f"- {n}" for n in reg.notes]
    return "\n".join(lines)


def render_form(grid: Grid, reg: Region, index: int) -> str:
    lines = [f"#### Form/labels {index} - range {reg.ref} (label: value pairs)"]
    for r in range(reg.r0, reg.r1 + 1):
        row = [(c, grid.get(r, c)) for c in range(reg.c0, reg.c1 + 1) if grid.get(r, c) is not None]
        i = 0
        while i < len(row):
            c, x = row[i]
            ref = f"{col_letter(c)}{r}"
            if family(x.kind) == "text" and i + 1 < len(row):
                c2, y = row[i + 1]
                lines.append(f"- {x.text.rstrip(':').strip()}: {y.text}  [{ref} -> {col_letter(c2)}{r}]")
                i += 2
            else:
                lines.append(f"- {x.text}  [{ref}]")
                i += 1
    return "\n".join(lines)


def render_text(grid: Grid, reg: Region, label: str = "Text") -> str:
    lines = [f"#### {label} - range {reg.ref}"]
    for r, c in sorted(reg.cells):
        lines.append(f"- {col_letter(c)}{r}: {grid.get(r, c).text}")
    return "\n".join(lines)


# ------------------------------------------------------------------ SQL copies

def _sql_value(x: GCell | None, fam: str):
    if x is None or x.kind in ("error", "formula"):
        return None
    if fam == "number" and isinstance(x.raw, (int, float)) and not isinstance(x.raw, bool):
        return float(x.raw)
    if fam == "bool" and isinstance(x.raw, bool):
        return int(x.raw)
    if fam == "date" and x.raw is not None and hasattr(x.raw, "isoformat"):
        return x.text  # ISO text, sortable and usable with SQLite date functions
    return x.text


def table_to_sql(grid: Grid, tv: TableView, name: str, source: str) -> list[SqlTable]:
    types = {"number": "REAL", "bool": "INTEGER"}
    col_names = unique_names([sql_ident(n, f"col_{col_letter(c).lower()}") for c, n in zip(tv.cols, tv.names)])
    has_group = bool(tv.group_of)
    columns = [(n, types.get(tv.col_family[c], "TEXT")) for n, c in zip(col_names, tv.cols)]
    columns += [("_row", "INTEGER"), ("_row_type", "TEXT")] + ([("_group", "TEXT")] if has_group else [])
    labels = [n + (" [percent stored as a fraction: 0.22 = 22%]" if tv.col_kind.get(c) == "percent" else "")
              for c, n in zip(tv.cols, tv.names)] + ["sheet row number", "data or total"] + (["group label above the row"] if has_group else [])
    rows = []
    for r in tv.body:
        if tv.row_type[r] == "group":
            continue
        vals = [_sql_value(grid.get(r, c), tv.col_family[c]) for c in tv.cols]
        rows.append(tuple(vals + [r, tv.row_type[r]] + ([tv.group_of.get(r)] if has_group else [])))
    out = [SqlTable(name, columns, labels, rows, source=source)]

    if tv.period_cols:
        id_cols = [(n, c) for n, c in zip(col_names, tv.cols) if c not in tv.period_cols
                   and tv.col_family[c] in ("text", "code", "date")]
        leaf = {c: tv.names[tv.cols.index(c)] for c in tv.period_cols}  # full label, e.g. "2025 / Q1"
        long_cols = [(n, "TEXT") for n, _ in id_cols] + [("period", "TEXT"), ("value", "REAL"),
                                                         ("_row", "INTEGER"), ("_row_type", "TEXT")]
        long_rows = []
        for r in tv.body:
            if tv.row_type[r] == "group":
                continue
            ids = [_sql_value(grid.get(r, c), "text") for _, c in id_cols]
            for c in tv.period_cols:
                long_rows.append(tuple(ids + [leaf[c], _sql_value(grid.get(r, c), "number"), r, tv.row_type[r]]))
        out.append(SqlTable(f"{name}_long", long_cols, [tv.names[tv.cols.index(c)] for _, c in id_cols]
                            + ["period column header", "value", "sheet row number", "data or total"],
                            long_rows, source=source, note="one row per (row, period) from the crosstab"))
    return out


# ------------------------------------------------------------------ whole sheet

def _subgrid(grid: Grid, reg: Region) -> Grid:
    return Grid(grid.name, {k: grid.cells[k] for k in reg.cells}, grid.source, grid.hidden_rows,
                grid.hidden_cols, grid.autofilter, [])


@dataclass
class Block:
    text: str
    summary: str | None = None   # shorter version for when the full text doesn't fit
    rows: int = 0


def render_grid(grid: Grid, *, sheet_label: str, table_prefix: str, start_index: int = 1,
                used_names: set[str] | None = None) -> tuple[list[Block], list[SqlTable], int, set[int]]:
    """Blocks of text for one sheet, SQL tables holding every row, the next table index,
    and the sheet rows that are table headers."""
    regions = find_regions(grid)
    blocks: list[Block] = []
    sql: list[SqlTable] = []
    header_rows: set[int] = set()
    idx = start_index
    used_names = used_names if used_names is not None else set()
    for reg in regions:
        if reg.kind == "table":
            view = _subgrid(grid, reg)  # only this region's cells, so nothing renders twice
            tv = analyze_table(view, reg)
            header_rows.update(tv.header_rows)
            label = reg.name or (sheet_label + (f"_{idx}" if len(regions) > 1 else ""))
            base = sql_ident(f"{table_prefix}{label}", f"table_{idx}")
            name = base
            n = 2
            while name in used_names:
                name, n = f"{base}_{n}", n + 1
            used_names.add(name)
            source = f"{sheet_label!r} {reg.ref}" if grid.source == "excel" else reg.ref
            sql += table_to_sql(view, tv, name, source)
            full = render_table(view, tv, index=idx, sql_name=name, summary=False)
            short = render_table(view, tv, index=idx, sql_name=name, summary=True)
            blocks.append(Block(full, short if estimate_tokens(short) < estimate_tokens(full) else None,
                                rows=len(tv.body)))
        elif reg.kind == "form":
            blocks.append(Block(render_form(grid, reg, idx)))
        elif reg.kind == "sparse":
            blocks.append(Block(render_text(grid, reg, "Scattered cells")))
        else:
            blocks.append(Block(render_text(grid, reg)))
        idx += 1
    return blocks, sql, idx, header_rows
