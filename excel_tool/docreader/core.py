"""Entry points: convert() and to_prompt_block()."""
from __future__ import annotations

import html
import re
from pathlib import Path

from .models import ConversionError, ConversionResult, estimate_tokens
from .sqlstore import catalog_text
from .util import is_ole, via_libreoffice

DEFAULT_MAX_INLINE_TOKENS = 20_000
MAX_FILE_BYTES = 100 * 1024 * 1024
RETRIEVAL_PREVIEW_TOKENS = 4_000

SPREADSHEET_EXT = {".xlsx", ".xlsm"}
LEGACY_SHEET_EXT = {".xls", ".xlsb", ".ods"}
DELIMITED_EXT = {".csv", ".tsv", ".tab"}
DOCX_EXT = {".docx"}
LEGACY_DOC_EXT = {".doc", ".rtf", ".odt"}
PPTX_EXT = {".pptx"}
LEGACY_PPT_EXT = {".ppt", ".odp"}
PDF_EXT = {".pdf"}
TEXT_EXT = {".txt", ".md", ".json", ".xml", ".yaml", ".yml", ".log"}
HTML_EXT = {".html", ".htm"}


def convert(
    path: str | Path,
    *,
    max_inline_tokens: int = DEFAULT_MAX_INLINE_TOKENS,
    table_prefix: str = "",
    include_hidden_sheets: bool = True,
    fill_merged_cells: bool = True,
    recalc_formulas: bool = False,
    include_structure: bool = True,
) -> ConversionResult:
    """Read one file. Never modifies it. Raises ConversionError with a user-safe message.

    table_prefix: prepended to SQL table names, e.g. a short document id when several
    files share one database.
    """
    from .delimited import convert_delimited
    from .documents import convert_docx, convert_html, convert_pdf, convert_pptx, convert_text
    from .excel import convert_xlsx

    path = Path(path)
    if not path.is_file():
        raise ConversionError(f"File not found: {path.name}")
    size = path.stat().st_size
    if size == 0:
        raise ConversionError(f"{path.name} is empty.")
    if size > MAX_FILE_BYTES:
        raise ConversionError(f"{path.name} is larger than {MAX_FILE_BYTES // 2**20} MB.")

    ext = path.suffix.lower()
    xl = dict(include_hidden_sheets=include_hidden_sheets, fill_merged_cells=fill_merged_cells,
              recalc_formulas=recalc_formulas, include_structure=include_structure,
              max_inline_tokens=max_inline_tokens, table_prefix=table_prefix)

    if ext in SPREADSHEET_EXT:
        if is_ole(path):
            raise ConversionError(
                f"{path.name} is password-protected/encrypted (or an old .xls renamed to {ext}). "
                "Remove the password and save it as .xlsx again.")
        res = convert_xlsx(path, **xl)
    elif ext in LEGACY_SHEET_EXT:
        res = via_libreoffice(path, "xlsx", lambda p: convert_xlsx(p, **xl))
    elif ext in DELIMITED_EXT:
        res = convert_delimited(path, max_inline_tokens=max_inline_tokens, table_prefix=table_prefix)
    elif ext in DOCX_EXT:
        res = convert_docx(path)
    elif ext in LEGACY_DOC_EXT:
        res = via_libreoffice(path, "docx", convert_docx)
    elif ext in PPTX_EXT:
        res = convert_pptx(path)
    elif ext in LEGACY_PPT_EXT:
        res = via_libreoffice(path, "pptx", convert_pptx)
    elif ext in PDF_EXT:
        res = convert_pdf(path)
    elif ext in TEXT_EXT:
        res = convert_text(path)
    elif ext in HTML_EXT:
        res = convert_html(path)
    else:
        raise ConversionError(f"Unsupported file type: {ext or '(no extension)'}")

    res.filename = path.name
    if not res.full_text:
        res.full_text = res.text
    if res.kind not in ("spreadsheet", "csv") and estimate_tokens(res.text) > max_inline_tokens:
        res.strategy = "retrieval"
        cut = int(RETRIEVAL_PREVIEW_TOKENS * 3.2)
        res.text = res.full_text[:cut].rsplit("\n", 1)[0] + "\n[... truncated ...]"
        res.warnings.append(f"File is about {estimate_tokens(res.full_text):,} tokens; only the beginning is "
                            "included here. Search the document for other sections.")
    res.meta["est_tokens"] = estimate_tokens(res.text)
    return res


_TAG = re.compile(r"</?\s*(document|content|conversion_notes)\b", re.I)


def to_prompt_block(result: ConversionResult) -> str:
    """Wrap a conversion for the Claude message. Tag look-alikes inside the file are escaped so
    the content can't close the wrapper and pose as instructions outside it."""
    def esc(s: str) -> str:
        return _TAG.sub(lambda m: html.escape(m.group(0)), s)

    name = html.escape(result.filename, quote=True)
    parts = [f'<document name="{name}" kind="{result.kind}" strategy="{result.strategy}">']
    if result.kind in ("spreadsheet", "csv"):
        parts.append("How to read this: tables are CSV whose first column is the sheet row number and "
                     "whose headers start with the column letter. Rows marked (total) are the file's own "
                     "totals; (group) rows are labels for the rows under them.")
    if result.warnings:
        parts.append("<conversion_notes>")
        parts += [f"- {esc(w)}" for w in result.warnings]
        parts.append("</conversion_notes>")
    parts += ["<content>", esc(result.text), "</content>"]
    cat = catalog_text(result.tables)
    if cat:
        parts.append(esc(cat))
    parts.append("</document>")
    return "\n".join(parts)
