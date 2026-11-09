"""CSV/TSV reader. Values are kept exactly as written (no type guessing on the text itself),
then the same layout analysis as Excel finds tables, titles, totals and notes."""
from __future__ import annotations

import csv
import io
import re
from collections import Counter
from pathlib import Path

from .cells import classify_text
from .excel import assemble
from .grid import GCell, Grid, render_grid
from .models import ConversionResult
from .util import decode_bytes

_EU_DECIMAL = re.compile(r"^-?\d{1,3}(\.\d{3})*,\d+$|^-?\d+,\d+$")


def detect_delimiter(text: str) -> str:
    """The delimiter giving the most consistent field count over the first lines.
    More predictable than csv.Sniffer, which punctuation inside text easily fools."""
    sample = text[:200_000]
    best, best_score = ",", -1.0
    for d in (",", ";", "\t", "|"):
        try:
            counts = [len(r) for _, r in zip(range(200), csv.reader(io.StringIO(sample, newline=""), delimiter=d))
                      if any(c.strip() for c in r)]
        except csv.Error:
            continue
        if not counts:
            continue
        mode, freq = Counter(counts).most_common(1)[0]
        if mode < 2:
            continue
        score = freq / len(counts) + mode * 1e-3
        if score > best_score:
            best, best_score = d, score
    return best


def convert_delimited(path: Path, *, max_inline_tokens: int = 20_000, table_prefix: str = "") -> ConversionResult:
    text, enc = decode_bytes(path.read_bytes())
    warns = []
    if enc not in ("utf-8", "utf-8-sig", "ascii"):
        warns.append(f"File was not UTF-8; decoded as {enc}.")
    delim = "\t" if path.suffix.lower() in (".tsv", ".tab") else detect_delimiter(text)
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))

    # Semicolon files usually come from European Excel, where "12,50" means 12.5.
    sample = [v.strip() for row in rows[:500] for v in row if v.strip()]
    decimal_comma = delim == ";" and sum(1 for v in sample if _EU_DECIMAL.match(v)) >= 3
    if decimal_comma:
        warns.append("Numbers use a decimal comma (e.g. 12,50 = 12.5); SQL values are converted.")

    lengths = Counter(len(r) for r in rows if any(c.strip() for c in r))
    if len(lengths) > 1:
        common = lengths.most_common(1)[0][0]
        odd = [i + 1 for i, r in enumerate(rows) if any(c.strip() for c in r) and len(r) != common]
        warns.append(f"{len(odd)} line(s) have a different number of fields than most "
                     f"(e.g. line {', '.join(map(str, odd[:5]))}); check column alignment there.")

    cells = {}
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            if v.strip():
                kind, raw = classify_text(v, from_csv=True, decimal_comma=decimal_comma)
                cells[(r, c)] = GCell(v, kind, raw)
    shown = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}.get(delim, repr(delim))
    grid = Grid(path.stem, cells, source="csv")
    blocks, sql, _, _ = render_grid(grid, sheet_label=path.stem, table_prefix=table_prefix)
    head = (f"### File: {path.name} ({shown}-delimited, {enc}). Row numbers count records from 1 (the line number unless a quoted value spans lines); "
            "column letters are as Excel would show them.")
    return assemble(path.name, "csv", [head, *blocks], sql, warns, max_inline_tokens,
                    meta={"encoding": enc, "delimiter": delim})
