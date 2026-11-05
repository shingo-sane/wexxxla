"""Findet externe Programme: LibreOffice, Microsoft Office, ffmpeg, pandoc."""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
from functools import lru_cache

IS_WIN = os.name == "nt"
# Unter Windows keine Konsolenfenster für Unterprozesse aufpoppen lassen
NO_WINDOW = 0x08000000 if IS_WIN else 0


def app_dir() -> str:
    """Ordner, in dem die .exe bzw. app.py liegt."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run(cmd: list[str], timeout: int = 600, **kw) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        stdin=subprocess.DEVNULL,
        timeout=timeout,
        creationflags=NO_WINDOW,
        **kw,
    )


# ---------------------------------------------------------------- LibreOffice
@lru_cache(maxsize=1)
def find_soffice() -> str | None:
    candidates = []
    if IS_WIN:
        for base in (
            os.environ.get("PROGRAMFILES", r"C:\Program Files"),
            os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
            os.environ.get("LOCALAPPDATA", ""),
        ):
            if base:
                candidates += glob.glob(os.path.join(base, "LibreOffice*", "program", "soffice.exe"))
        # portable LibreOffice neben der App
        candidates += glob.glob(os.path.join(app_dir(), "LibreOffice*", "**", "soffice.exe"), recursive=True)
    else:
        candidates += [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice/program/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
        candidates += glob.glob("/opt/libreoffice*/program/soffice")
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            candidates.insert(0, p)
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


# ---------------------------------------------------------- Microsoft Office
@lru_cache(maxsize=1)
def find_msoffice() -> dict[str, bool]:
    """Welche MS-Office-Programme per COM ansprechbar sind (nur Windows)."""
    res = {"word": False, "excel": False, "powerpoint": False}
    if not IS_WIN:
        return res
    try:
        import winreg

        for key, progid in (
            ("word", "Word.Application"),
            ("excel", "Excel.Application"),
            ("powerpoint", "PowerPoint.Application"),
        ):
            try:
                winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"))
                res[key] = True
            except OSError:
                pass
    except Exception:
        pass
    return res


# --------------------------------------------------------------------- ffmpeg
@lru_cache(maxsize=1)
def find_ffmpeg() -> str | None:
    local = os.path.join(app_dir(), "ffmpeg.exe" if IS_WIN else "ffmpeg")
    if os.path.isfile(local):
        return local
    try:
        import imageio_ffmpeg  # bringt eine eigene ffmpeg-Binary mit

        p = imageio_ffmpeg.get_ffmpeg_exe()
        if p and os.path.isfile(p):
            return p
    except Exception:
        pass
    return shutil.which("ffmpeg")


# --------------------------------------------------------------------- pandoc
@lru_cache(maxsize=1)
def find_pandoc() -> str | None:
    local = os.path.join(app_dir(), "pandoc.exe" if IS_WIN else "pandoc")
    if os.path.isfile(local):
        return local
    try:
        import pypandoc  # pypandoc_binary bringt pandoc mit

        p = pypandoc.get_pandoc_path()
        if p and os.path.isfile(p):
            return p
    except Exception:
        pass
    return shutil.which("pandoc")


def status() -> dict:
    ms = find_msoffice()
    return {
        "libreoffice": find_soffice(),
        "msoffice": [k for k, v in ms.items() if v],
        "ffmpeg": find_ffmpeg(),
        "pandoc": find_pandoc(),
    }
