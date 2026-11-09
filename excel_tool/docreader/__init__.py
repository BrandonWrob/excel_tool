"""docreader - read uploaded files into text and SQL tables a model can answer questions from.

    from docreader import convert, to_prompt_block, build_sqlite, ReadOnlyDB, QUERY_TOOL

    result = convert("upload.xlsx")              # never modifies the file
    block = to_prompt_block(result)              # put in the user message
    build_sqlite(result, "upload.sqlite")        # every row, for the query tool
    db = ReadOnlyDB("upload.sqlite")
    db.query("SELECT region, SUM(revenue) FROM sales WHERE _row_type = 'data' GROUP BY region")
"""
from .core import convert, to_prompt_block
from .models import ConversionError, ConversionResult, SqlTable
from .sqlstore import QUERY_TOOL, ReadOnlyDB, build_sqlite, catalog_text

__all__ = ["convert", "to_prompt_block", "ConversionError", "ConversionResult", "SqlTable",
           "QUERY_TOOL", "ReadOnlyDB", "build_sqlite", "catalog_text"]
