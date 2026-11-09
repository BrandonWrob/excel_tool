"""Result types shared by every reader."""
from __future__ import annotations

from dataclasses import dataclass, field


class ConversionError(Exception):
    """A file couldn't be read. The message is written to be safe to show the user."""


@dataclass
class SqlTable:
    """One table's complete data, ready to load into the per-document SQLite database."""
    name: str                          # SQL table name
    columns: list[tuple[str, str]]     # (sql column name, sql type)
    labels: list[str]                  # original column label for each column
    rows: list[tuple]
    source: str = ""                   # e.g. "sheet 'Sales' A1:H9"
    note: str = ""                     # extra hint for the model (e.g. long-format crosstab)


@dataclass
class ConversionResult:
    filename: str
    kind: str                          # spreadsheet | csv | document | slides | pdf | text
    text: str                          # what to put in the prompt (large tables summarized)
    full_text: str = ""                # everything, unsummarized (for search indexing)
    warnings: list[str] = field(default_factory=list)
    tables: list[SqlTable] = field(default_factory=list)
    strategy: str = "inline"           # inline | sql | retrieval
    meta: dict = field(default_factory=dict)

    @property
    def est_tokens(self) -> int:
        return estimate_tokens(self.text)


def estimate_tokens(text: str) -> int:
    """Rough count (characters / 3.2). Use the token-counting API for exact numbers."""
    return int(len(text) / 3.2) + 1
