"""Turn raw cell values into (display text, kind, raw value).

Kinds:  number currency percent numtext | date datetime time | bool | code | text | error | formula
"numtext" is a number stored as text in Excel; "code" is a digit string where leading zeros or
length matter (NDC, zip, lot, long IDs); "formula" is a formula with no saved result.
"""
from __future__ import annotations

import datetime as dt
import re

ERROR_VALUES = {"#N/A", "#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!",
                "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!", "#GETTING_DATA"}

_FAMILY = {"number": "number", "currency": "number", "percent": "number", "numtext": "number",
           "date": "date", "datetime": "date", "time": "date", "bool": "bool", "code": "code",
           "text": "text", "error": "error", "formula": "unknown", "empty": "empty"}


def family(kind: str) -> str:
    return _FAMILY.get(kind, "text")


# ------------------------------------------------------------------ text classification (CSV etc.)

_PLAIN_NUM = re.compile(r"^(\d+(\.\d*)?|\.\d+)$")
_SCI_NUM = re.compile(r"^\d+\.\d+[eE][-+]?\d+$")
_GROUPED = re.compile(r"^\d{1,3}(,\d{3})+(\.\d+)?$")
_EU_NUM = re.compile(r"^(\d{1,3}(\.\d{3})+|\d+)(,\d+)?$")
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2}[ T]\d{1,2}:\d{2}(:\d{2})?")
_SLASH_DATE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
_CURRENCY_CHARS = "$€£¥₹"


def parse_number(text: str, *, decimal_comma: bool = False) -> tuple[str, float] | None:
    """'$1,234.50' -> ('currency', 1234.5); '(12%)' -> ('percent', -0.12); 'abc' -> None."""
    s = text.replace(" ", " ").strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    if s[:1] in "+-":
        neg, s = (neg != (s[0] == "-")), s[1:].strip()
    pct = s.endswith("%")
    if pct:
        s = s[:-1].strip()
    cur = None
    if s[:1] and s[0] in _CURRENCY_CHARS:
        cur, s = s[0], s[1:].strip()
    elif s[-1:] and s[-1] in _CURRENCY_CHARS:
        cur, s = s[-1], s[:-1].strip()
    if s[:1] == "-":
        neg, s = not neg, s[1:]
    if decimal_comma and _EU_NUM.match(s) and ("," in s or "." in s):
        s = s.replace(".", "").replace(",", ".")
    elif _GROUPED.match(s):
        s = s.replace(",", "")
    if not (_PLAIN_NUM.match(s) or _SCI_NUM.match(s)):
        return None
    v = float(s)
    v = -v if neg else v
    if pct:
        return "percent", v / 100
    if cur:
        return "currency", v
    return "number", v


def classify_text(text: str, *, from_csv: bool, decimal_comma: bool = False) -> tuple[str, object]:
    """Kind and raw value for a value that arrived as text."""
    t = text.strip()
    if not t:
        return "empty", None
    up = t.upper()
    if up in ERROR_VALUES:
        return "error", None
    if up in ("TRUE", "FALSE"):
        return "bool", up == "TRUE"
    digits = t.replace("-", "")
    if (re.fullmatch(r"0\d+", t) or (digits.isdigit() and len(digits) > 15)
            or re.fullmatch(r"\d+(-\d+){1,3}", t) and not _ISO_DATE.match(t)):
        return "code", None  # leading zeros, long IDs, NDC-style 00093-7180-01
    parsed = parse_number(t, decimal_comma=decimal_comma)
    if parsed is not None:
        kind, val = parsed
        if kind == "number" and not from_csv:
            return "numtext", val
        return kind, val
    if _ISO_DATE.match(t):
        return "date", t
    if _ISO_DATETIME.match(t):
        return "datetime", t
    if _SLASH_DATE.match(t):
        return "date", None  # ambiguous M/D vs D/M - kept as text in SQL
    return "text", None


# ------------------------------------------------------------------ Excel values

_TIME_FMT = re.compile(r"[hs]", re.I)
_CURRENCY_FMT = re.compile(r'\[\$([^\-\]]+)[^\]]*\]|"([$€£¥₹])"|([$€£¥₹])')


def strip_format_literals(nf: str) -> str:
    """Remove quoted text, [colors]/[$-locale] and escapes from an Excel number format."""
    nf = re.sub(r'"[^"]*"', "", nf or "General")
    nf = re.sub(r"\[[^\]]*\]", "", nf)
    nf = re.sub(r"\\.", "", nf)
    return nf.split(";")[0]


def num_text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if v != v or v in (float("inf"), float("-inf")):
        return str(v)
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return format(v, ".15g")  # 0.1+0.2 -> 0.3


def _currency_text(value: float, nf: str, core: str) -> str | None:
    first = nf.split(";")[0]
    m = _CURRENCY_FMT.search(first)
    if not m:
        return None
    symbol = next(g for g in m.groups() if g)
    dec = re.search(r"\.(0+)", core)
    decimals = len(dec.group(1)) if dec else 0
    body = f"{abs(value):,.{decimals}f}" if "," in core else f"{abs(value):.{decimals}f}"
    digit = re.search(r"[0#?]", first)
    text = f"{body} {symbol}" if digit is not None and m.start() > digit.start() else f"{symbol}{body}"
    return f"-{text}" if value < 0 else text


def classify_excel(value, number_format: str | None) -> tuple[str, str, object]:
    """(display text, kind, raw) for an openpyxl cell value, shown the way Excel displays it."""
    nf = number_format or "General"
    if value is None:
        return "", "empty", None
    if isinstance(value, bool):
        return ("TRUE" if value else "FALSE"), "bool", value
    if isinstance(value, dt.datetime):
        if value.time() == dt.time(0) and not _TIME_FMT.search(strip_format_literals(nf)):
            return value.date().isoformat(), "date", value
        return value.isoformat(sep=" ", timespec="seconds"), "datetime", value
    if isinstance(value, dt.date):
        return value.isoformat(), "date", value
    if isinstance(value, dt.time):
        return value.isoformat(), "time", value
    if isinstance(value, dt.timedelta):
        return str(value), "time", value
    if isinstance(value, (int, float)):
        core = strip_format_literals(nf)
        if "%" in core:
            return num_text(value * 100) + "%", "percent", value
        if re.fullmatch(r"0{2,}", core) and float(value).is_integer():
            return str(int(value)).zfill(len(core)), "code", value  # "00000" keeps zip codes
        money = _currency_text(value, nf, core)
        if money is not None:
            return money, "currency", value
        return num_text(value), "number", value
    text = str(value).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    if text.startswith("="):
        return text, "formula", None
    kind, raw = classify_text(text, from_csv=False)
    if kind == "empty":
        return "", "empty", None
    return text, kind, raw
