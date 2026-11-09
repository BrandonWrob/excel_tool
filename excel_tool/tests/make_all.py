"""Generate every sample file the checks use. Needs LibreOffice (soffice) on PATH.
Usage: python tests/make_all.py samples"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

here = Path(__file__).resolve().parent
out = Path(sys.argv[1]).resolve()
out.mkdir(parents=True, exist_ok=True)


def run(*args):
    subprocess.run([sys.executable, *map(str, args)], check=True)


def soffice(fmt, src):
    subprocess.run(["soffice", "-env:UserInstallation=file:///tmp/lo_make_all", "--headless", "--convert-to", fmt,
                    "--outdir", str(out), str(src)], check=True, capture_output=True, timeout=180)


run(here / "make_samples.py", out)
soffice("xls", out / "messy.xlsx")
soffice("pdf", out / "report.docx")
run(here / "make_rich_xlsx.py", out)
shutil.copy(out / "rich.xlsm", out / "macro.xlsm")
with zipfile.ZipFile(out / "macro.xlsm", "a") as z:
    z.writestr("xl/vbaProject.bin", b"\x00not-real-vba")  # only presence is checked
shutil.copy(out / "messy.xls", out / "renamed.xlsx")      # legacy/encrypted container with .xlsx name
run(here / "make_zoo.py", out / "zoo")
print("all samples in", out)
