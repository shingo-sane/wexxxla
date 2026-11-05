"""Office-Konvertierungen über Microsoft Office (falls installiert) oder LibreOffice."""
from __future__ import annotations

import os
import shutil
import tempfile
import threading
import uuid

from . import tools
from .engine import ConversionError, Ctx, route, _settings
from .formats import WRITER, CALC, IMPRESS

_lock = threading.Lock()  # Office-Programme nicht parallel starten

# ------------------------------------------------------------- LibreOffice
LO_FILTERS = {
    "pdf": "pdf",
    "docx": 'docx:MS Word 2007 XML',
    "doc": "doc:MS Word 97",
    "odt": "odt",
    "rtf": "rtf:Rich Text Format",
    "txt": "txt:Text (encoded):UTF8",
    "html": "html:XHTML Writer File:UTF8",
    "xlsx": "xlsx:Calc MS Excel 2007 XML",
    "xls": "xls:MS Excel 97",
    "ods": "ods",
    "csv": "csv:Text - txt - csv (StarCalc):44,34,76,1",
    "pptx": "pptx:Impress MS PowerPoint 2007 XML",
    "ppt": "ppt:MS PowerPoint 97",
    "odp": "odp",
    "png": "png",
}

LO_INFILTERS = {
    "html": "HTML (StarWriter)",
    "txt": "Text (encoded):UTF8",
    "csv": "CSV:44,34,76,1",
}


def _lo_profile() -> str:
    d = os.path.join(tempfile.gettempdir(), "ufc_libreoffice_profile")
    os.makedirs(d, exist_ok=True)
    return "file:///" + d.replace("\\", "/").lstrip("/")


def lo_convert(src: str, dst: str, src_ext: str, dst_ext: str, timeout: int = 300):
    soffice = tools.find_soffice()
    if not soffice:
        raise ConversionError("LibreOffice wurde nicht gefunden.")
    work = tempfile.mkdtemp(prefix="ufc_lo_")
    try:
        inp = os.path.join(work, f"input.{src_ext}")
        shutil.copyfile(src, inp)
        outdir = os.path.join(work, "out")
        os.makedirs(outdir)
        cmd = [soffice, f"-env:UserInstallation={_lo_profile()}", "--headless", "--norestore",
               "--nologo", "--nodefault", "--nolockcheck"]
        if src_ext in LO_INFILTERS:
            cmd.append(f"--infilter={LO_INFILTERS[src_ext]}")
        cmd += ["--convert-to", LO_FILTERS.get(dst_ext, dst_ext), "--outdir", outdir, inp]
        with _lock:
            try:
                res = tools.run(cmd, timeout=timeout)
            except Exception as e:
                raise ConversionError(f"LibreOffice reagiert nicht ({e}).")
        produced = [f for f in os.listdir(outdir)]
        if not produced:
            msg = (res.stderr or b"").decode(errors="ignore").strip()[-300:]
            raise ConversionError("LibreOffice konnte die Datei nicht umwandeln. " + msg)
        shutil.move(os.path.join(outdir, produced[0]), dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)


# ------------------------------------------------------- Microsoft Office
MS_WORD_FMT = {"docx": 16, "doc": 0, "odt": 23, "rtf": 6, "html": 10, "pdf": 17}
MS_EXCEL_FMT = {"xlsx": 51, "xls": 56, "ods": 60, "csv": 6, "pdf": -1}
MS_PPT_FMT = {"pptx": 24, "ppt": 1, "odp": 35, "pdf": 32}

PS_WORD = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0
try {
  $doc = $app.Documents.Open($env:UFC_SRC, $false, $true, $false)
  if ([int]$env:UFC_FMT -eq 17) { $doc.ExportAsFixedFormat($env:UFC_DST, 17) }
  else { $doc.SaveAs2($env:UFC_DST, [int]$env:UFC_FMT) }
  $doc.Close(0)
} finally {
  $app.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""

PS_EXCEL = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
try {
  $wb = $app.Workbooks.Open($env:UFC_SRC, 0, $true)
  if ([int]$env:UFC_FMT -eq -1) { $wb.ExportAsFixedFormat(0, $env:UFC_DST) }
  else { $wb.SaveAs($env:UFC_DST, [int]$env:UFC_FMT) }
  $wb.Close($false)
} finally {
  $app.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""

PS_PPT = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($env:UFC_SRC, -1, 0, 0)
  $pres.SaveAs($env:UFC_DST, [int]$env:UFC_FMT)
  $pres.Close()
} finally {
  $app.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""


def ms_convert(app: str, src: str, dst: str, dst_ext: str, timeout: int = 300):
    script, fmts = {"word": (PS_WORD, MS_WORD_FMT), "excel": (PS_EXCEL, MS_EXCEL_FMT),
                    "powerpoint": (PS_PPT, MS_PPT_FMT)}[app]
    if dst_ext not in fmts:
        raise ConversionError("Format wird von Microsoft Office nicht unterstützt.")
    work = tempfile.mkdtemp(prefix="ufc_ms_")
    try:
        ext = os.path.splitext(src)[1]
        inp = os.path.join(work, "input" + ext)
        shutil.copyfile(src, inp)
        out = os.path.join(work, f"output_{uuid.uuid4().hex[:6]}.{dst_ext}")
        ps1 = os.path.join(work, "convert.ps1")
        with open(ps1, "w", encoding="utf-8-sig") as f:
            f.write(script)
        env = dict(os.environ, UFC_SRC=os.path.abspath(inp), UFC_DST=os.path.abspath(out),
                   UFC_FMT=str(fmts[dst_ext]))
        with _lock:
            res = tools.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                             "-File", ps1], timeout=timeout, env=env)
        if not os.path.exists(out):
            raise ConversionError("Microsoft Office: " + (res.stderr or b"").decode(errors="ignore")[-300:])
        shutil.move(out, dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def office_convert(src: str, dst: str, ctx: Ctx, app: str):
    """app: word | excel | powerpoint | draw"""
    engine = _settings["office_engine"]
    ms_ok = app != "draw" and tools.find_msoffice().get(app, False)
    lo_ok = tools.find_soffice() is not None
    errors = []
    if ms_ok and engine in ("auto", "ms"):
        try:
            ctx.progress(0.1, "Microsoft Office arbeitet …")
            ms_convert(app, src, dst, ctx.dst_ext)
            return
        except Exception as e:
            errors.append(str(e))
    if lo_ok and engine in ("auto", "lo") or (lo_ok and not ms_ok):
        ctx.progress(0.1, "LibreOffice arbeitet …")
        try:
            lo_convert(src, dst, ctx.src_ext, ctx.dst_ext)
            return
        except Exception as e:
            errors.append(str(e))
    raise ConversionError(" | ".join(errors) or "Kein Office-Programm verfügbar (LibreOffice installieren).")


# ------------------------------------------------------------------ Routen
@route(WRITER, WRITER + ["pdf"], cost=3, backend="office-writer")
def writer_convert(src, dst, ctx):
    office_convert(src, dst, ctx, "word")


@route(CALC, CALC + ["pdf"], cost=3, backend="office-calc")
def calc_convert(src, dst, ctx):
    office_convert(src, dst, ctx, "excel")


@route(IMPRESS, IMPRESS + ["pdf"], cost=3, backend="office-impress")
def impress_convert(src, dst, ctx):
    office_convert(src, dst, ctx, "powerpoint")


@route("svg", ["pdf"], cost=3, backend="office-draw")
def draw_convert(src, dst, ctx):
    office_convert(src, dst, ctx, "draw")
