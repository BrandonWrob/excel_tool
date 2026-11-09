"""Every extracted table, every row, in a per-document SQLite file that Claude queries read-only.

Why SQLite per document: no schema management in your main database, no chance of one user's
upload being queryable from another user's chat, and the file can be stored next to the upload
(e.g. in Blob Storage) and deleted with it.

Read-only is enforced three ways: the file is opened with mode=ro, PRAGMA query_only is on,
and an authorizer rejects anything other than reading (no INSERT/UPDATE/DELETE/DROP/ATTACH/PRAGMA).
"""
from __future__ import annotations

import csv
import io
import sqlite3
import time
from pathlib import Path

from .models import ConversionResult, SqlTable

MAX_RESULT_ROWS = 200
QUERY_TIMEOUT_S = 5.0

# Generic on purpose: table and column names can contain business data, so they go in the
# message content (see catalog_text), never in the tool schema.
QUERY_TOOL = {
    "name": "query_spreadsheet",
    "description": (
        "Run one read-only SQL SELECT (SQLite dialect) against the tables extracted from files the "
        "user uploaded. The available tables and columns are listed in the <document> blocks under "
        "'SQL tables'. Use this for exact counts, sums, filters and lookups instead of reading rows "
        "by eye, and whenever a table was only summarized. Rows with _row_type = 'total' are the "
        "file's own total/subtotal rows: exclude them when you aggregate. _row is the original "
        "sheet row number; cite it so the user can find the row. Results are capped at "
        f"{MAX_RESULT_ROWS} rows."
    ),
    "input_schema": {
        "type": "object",
        "properties": {"sql": {"type": "string", "description": "A single SELECT or WITH ... SELECT statement."}},
        "required": ["sql"],
    },
}


def build_sqlite(results: list[ConversionResult] | ConversionResult, path: str | Path) -> Path:
    """Write all tables from one or more conversions to a new SQLite file."""
    if isinstance(results, ConversionResult):
        results = [results]
    path = Path(path)
    if path.exists():
        path.unlink()
    conn = sqlite3.connect(path)
    try:
        conn.execute("CREATE TABLE _tables (name TEXT PRIMARY KEY, file TEXT, source TEXT, row_count INTEGER, "
                     "columns TEXT, note TEXT)")
        for res in results:
            for t in res.tables:
                cols = ", ".join(f'"{n}" {typ}' for n, typ in t.columns)
                conn.execute(f'CREATE TABLE "{t.name}" ({cols})')
                marks = ", ".join("?" for _ in t.columns)
                conn.executemany(f'INSERT INTO "{t.name}" VALUES ({marks})', t.rows)
                desc = "; ".join(f"{n} = {lbl!r}" for (n, _), lbl in zip(t.columns, t.labels))
                conn.execute("INSERT INTO _tables VALUES (?, ?, ?, ?, ?, ?)",
                             (t.name, res.filename, t.source, len(t.rows), desc, t.note))
        conn.commit()
    finally:
        conn.close()
    return path


def catalog_text(tables: list[SqlTable]) -> str:
    """Short table list for the prompt so the model knows what it can query."""
    if not tables:
        return ""
    lines = ["SQL tables (query with the query_spreadsheet tool; every row is included):"]
    for t in tables:
        shown = []
        for (n, typ), lbl in zip(t.columns, t.labels):
            if n.startswith("_"):
                continue
            kind = {"REAL": " num", "INTEGER": " bool"}.get(typ, "")
            shown.append(f"{n}{kind}" if n == lbl.lower() else f"{n}{kind} ({lbl})")
        extra = ", ".join(n for n, _ in t.columns if n.startswith("_"))
        note = f" - {t.note}" if t.note else ""
        n_rows = f"{len(t.rows)} row{'s' if len(t.rows) != 1 else ''}"
        lines.append(f"- {t.name} [{t.source}, {n_rows}]{note}: {', '.join(shown)}; plus {extra}")
    return "\n".join(lines)


_ALLOWED = {getattr(sqlite3, n) for n in ("SQLITE_SELECT", "SQLITE_READ", "SQLITE_FUNCTION", "SQLITE_RECURSIVE")
            if hasattr(sqlite3, n)}


class ReadOnlyDB:
    def __init__(self, path: str | Path):
        uri = f"file:{Path(path).resolve()}?mode=ro"
        self.conn = sqlite3.connect(uri, uri=True, check_same_thread=False)
        self.conn.execute("PRAGMA query_only = ON")
        self.conn.set_authorizer(self._authorize)

    @staticmethod
    def _authorize(action, arg1, arg2, db_name, trigger):
        return sqlite3.SQLITE_OK if action in _ALLOWED else sqlite3.SQLITE_DENY

    def query(self, sql: str, max_rows: int = MAX_RESULT_ROWS, timeout_s: float = QUERY_TIMEOUT_S) -> str:
        """Run one statement; return CSV text (or an error message the model can act on)."""
        sql = sql.strip().rstrip(";").strip()
        if not sql.lower().startswith(("select", "with")):
            return "Error: only SELECT queries are allowed."
        deadline = time.monotonic() + timeout_s
        self.conn.set_progress_handler(lambda: 1 if time.monotonic() > deadline else 0, 10_000)
        try:
            cur = self.conn.execute(sql)
            rows = cur.fetchmany(max_rows + 1)
        except sqlite3.DatabaseError as e:
            msg = str(e)
            if "interrupted" in msg:
                msg = f"query took longer than {timeout_s:g}s; narrow it down"
            return f"Error: {msg}"
        finally:
            self.conn.set_progress_handler(None, 0)
        cols = [d[0] for d in cur.description or []]
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\n")
        w.writerow(cols)
        w.writerows(rows[:max_rows])
        out = buf.getvalue().rstrip("\n")
        if len(rows) > max_rows:
            out += f"\n(results truncated to {max_rows} rows; aggregate or filter to see the rest)"
        return out or "(no rows)"

    def close(self):
        self.conn.close()
