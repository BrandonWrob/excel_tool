"""Small helpers: decoding, file guards, LibreOffice conversion, naming."""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from .models import ConversionError

MAX_UNCOMPRESSED_BYTES = 400 * 1024 * 1024  # zip-bomb guard for xlsx/docx/pptx
LIBREOFFICE_TIMEOUT_S = 120
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy Office binary / encrypted OOXML


def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def ranges_text(nums: list[int]) -> str:
    """[1,2,3,7,9,10] -> '1-3, 7, 9-10'"""
    if not nums:
        return ""
    nums = sorted(nums)
    out, start, prev = [], nums[0], nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == prev + 1:
            prev = n
            continue
        out.append(f"{start}" if start == prev else f"{start}-{prev}")
        if n is not None:
            start = prev = n
    return ", ".join(out)


def sql_ident(label: str, fallback: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", (label or "").strip().lower()).strip("_")[:60] or fallback
    return f"c_{s}" if s[0].isdigit() else s


def unique_names(names: list[str]) -> list[str]:
    seen: Counter[str] = Counter()
    out = []
    for n in names:
        seen[n] += 1
        out.append(n if seen[n] == 1 else f"{n}_{seen[n]}")
    return out


def decode_bytes(data: bytes) -> tuple[str, str]:
    """Decode predictably: BOM -> UTF-8 -> Windows-1252 -> best guess -> Latin-1."""
    for bom, enc in ((b"\xef\xbb\xbf", "utf-8-sig"), (b"\xff\xfe", "utf-16"), (b"\xfe\xff", "utf-16")):
        if data.startswith(bom):
            return data.decode(enc), enc
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    try:
        return data.decode("cp1252"), "cp1252"  # Excel's "CSV" on US Windows
    except UnicodeDecodeError:
        pass
    try:
        from charset_normalizer import from_bytes
        best = from_bytes(data).best()
        if best is not None:
            return str(best), best.encoding
    except ImportError:
        pass
    return data.decode("latin-1"), "latin-1"


def guard_zip(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as z:
            total = sum(i.file_size for i in z.infolist())
    except zipfile.BadZipFile as e:
        raise ConversionError(f"{path.name} is corrupt, encrypted, or not really a {path.suffix} file.") from e
    if total > MAX_UNCOMPRESSED_BYTES:
        raise ConversionError(f"{path.name} expands to more than {MAX_UNCOMPRESSED_BYTES // 2**20} MB.")


def is_ole(path: Path) -> bool:
    with open(path, "rb") as f:
        return f.read(8) == OLE_MAGIC


def via_libreoffice(path: Path, fmt: str, then, note: str | None = None):
    """Convert with headless LibreOffice into a temp dir, then call then(converted_path)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise ConversionError(f"{path.suffix} files need converting first; save as .{fmt} and re-upload.")
    with tempfile.TemporaryDirectory() as tmp:
        tmp_p = Path(tmp)
        src = tmp_p / ("src" + path.suffix.lower())
        shutil.copy(path, src)
        outdir = tmp_p / "out"
        # A private profile per call lets several conversions run at once.
        cmd = [soffice, f"-env:UserInstallation=file://{tmp_p / 'profile'}", "--headless",
               "--convert-to", fmt, "--outdir", str(outdir), str(src)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=LIBREOFFICE_TIMEOUT_S)
        except subprocess.TimeoutExpired as e:
            raise ConversionError(f"Converting {path.name} timed out.") from e
        except subprocess.CalledProcessError as e:
            raise ConversionError(f"LibreOffice could not convert {path.name}.") from e
        converted = outdir / f"src.{fmt}"
        if not converted.exists():
            raise ConversionError(f"LibreOffice could not convert {path.name} (encrypted or corrupt?).")
        res = then(converted)
    res.filename = path.name
    msg = note or (f"Converted from {path.suffix} to .{fmt} with LibreOffice before reading."
                   if path.suffix.lower() != f".{fmt}" else None)
    if msg:
        res.warnings.insert(0, msg)
    return res
