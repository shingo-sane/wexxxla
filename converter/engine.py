"""Kern: Registrierung der Konverter, Wegsuche (auch über Zwischenformate) und Ausführung."""
from __future__ import annotations

import heapq
import os
import shutil
import tempfile
import traceback
from dataclasses import dataclass, field
from typing import Callable

from . import tools
from .formats import FORMATS, normalize_ext, label


class ConversionError(Exception):
    """Fehler mit verständlicher Meldung für den Nutzer."""


@dataclass
class Ctx:
    src_ext: str
    dst_ext: str
    options: dict = field(default_factory=dict)
    tmpdir: str = ""
    progress: Callable[[float, str], None] = lambda p, m="": None

    def opt(self, key, default=None):
        v = self.options.get(key)
        return default if v in (None, "") else v


@dataclass
class Route:
    src: str
    dst: str
    func: Callable[[str, str, Ctx], None]
    cost: float
    backend: str  # builtin | office-writer | office-calc | office-impress | office-draw | ffmpeg | pandoc | heif
    multi: bool = False  # erzeugt evtl. mehrere Dateien -> nur als letzter Schritt erlaubt
    name: str = ""


ROUTES: list[Route] = []


def route(srcs, dsts, cost: float = 1.0, backend: str = "builtin", multi: bool = False):
    """Dekorator: registriert eine Konvertierfunktion für alle Paare srcs x dsts."""
    if isinstance(srcs, str):
        srcs = [srcs]
    if isinstance(dsts, str):
        dsts = [dsts]

    def deco(fn):
        for s in srcs:
            for d in dsts:
                if s != d:
                    ROUTES.append(Route(s, d, fn, cost, backend, multi, fn.__name__))
        return fn

    return deco


# ------------------------------------------------------------ Verfügbarkeit
_settings = {"office_engine": "auto"}  # auto | ms | lo


def set_office_engine(engine: str):
    _settings["office_engine"] = engine


def backend_available(backend: str) -> bool:
    if backend == "builtin":
        return True
    if backend.startswith("office-"):
        app = backend.split("-", 1)[1]
        lo = tools.find_soffice() is not None
        ms = tools.find_msoffice()
        ms_ok = {"writer": ms["word"], "calc": ms["excel"], "impress": ms["powerpoint"]}.get(app, False)
        eng = _settings["office_engine"]
        if eng == "lo":
            return lo
        if eng == "ms":
            return ms_ok or (app == "draw" and lo)
        return lo or ms_ok
    if backend == "ffmpeg":
        return tools.find_ffmpeg() is not None
    if backend == "pandoc":
        return tools.find_pandoc() is not None
    if backend == "heif":
        try:
            import pillow_heif  # noqa: F401
            return True
        except Exception:
            return False
    return False


def _available_routes() -> list[Route]:
    _ensure_loaded()
    return [r for r in ROUTES if backend_available(r.backend)]


def _ensure_loaded():
    # Module registrieren ihre Routen beim Import
    from . import office, pdf, builders, media, textfmt, data  # noqa: F401


# ------------------------------------------------------------------ Wegsuche
MAX_STEPS = 4


def allowed_formats(src: str) -> set[str]:
    """Sinnvolle Ziel- und Zwischenformate je nach Quellart (kein 'PNG → MP3' o. ä.)."""
    from .formats import category, RASTER, AUDIO, VIDEO

    doc_cats = {"PDF", "Dokument", "Präsentation", "Tabelle"}
    cat = category(src)
    if cat in doc_cats:
        return {f for f in FORMATS if category(f) in doc_cats} | (set(RASTER) - {"ico"})
    if cat == "Bild":
        s = set(RASTER) | {"pdf", "docx", "doc", "odt", "rtf", "pptx", "ppt", "odp"}
        if src == "gif":
            s |= set(VIDEO)
        return s
    if cat == "Video":
        return set(VIDEO) | set(AUDIO) | {"gif", "png", "jpg"}
    if cat == "Audio":
        return set(AUDIO)
    return set()


def _dijkstra(src: str, routes: list[Route]):
    """Kürzeste Wege von src zu allen Formaten. Gibt {fmt: (kosten, [routen])} zurück."""
    allowed = allowed_formats(src)
    by_src: dict[str, list[Route]] = {}
    for r in routes:
        if r.dst in allowed:
            by_src.setdefault(r.src, []).append(r)
    best: dict[str, tuple[float, list[Route]]] = {}
    heap = [(0.0, 0, src, [])]
    counter = 0
    settled = set()
    while heap:
        cost, _, fmt, path = heapq.heappop(heap)
        if fmt in settled:
            continue
        settled.add(fmt)
        if fmt != src:
            best[fmt] = (cost, path)
        if path and path[-1].multi:
            continue  # nach einem Mehrdatei-Schritt geht es nicht weiter
        if len(path) >= MAX_STEPS:
            continue
        for r in by_src.get(fmt, []):
            if r.dst in settled or r.dst == src:
                continue
            counter += 1
            heapq.heappush(heap, (cost + r.cost + 0.05, counter, r.dst, path + [r]))
    return best


def find_path(src: str, dst: str) -> list[Route] | None:
    src, dst = normalize_ext(src), normalize_ext(dst)
    res = _dijkstra(src, _available_routes())
    return res.get(dst, (None, None))[1]


def targets_for(src: str) -> list[str]:
    src = normalize_ext(src)
    res = _dijkstra(src, _available_routes())
    return [f for f in FORMATS if f in res]


def unavailable_targets(src: str) -> dict[str, str]:
    """Formate, die mit zusätzlicher Software erreichbar wären -> {fmt: grund}."""
    src = normalize_ext(src)
    _ensure_loaded()
    have = set(targets_for(src))
    everything = _dijkstra(src, ROUTES)
    out = {}
    for fmt, (_, path) in everything.items():
        if fmt in have or fmt not in FORMATS:
            continue
        missing = {r.backend for r in path if not backend_available(r.backend)}
        reasons = []
        if any(b.startswith("office-") for b in missing):
            reasons.append("LibreOffice oder Microsoft Office")
        if "ffmpeg" in missing:
            reasons.append("ffmpeg")
        if "pandoc" in missing:
            reasons.append("pandoc")
        if "heif" in missing:
            reasons.append("pillow-heif")
        out[fmt] = "benötigt " + " + ".join(reasons or ["Zusatzsoftware"])
    return out


def describe_path(path: list[Route]) -> str:
    if not path:
        return ""
    fmts = [path[0].src] + [r.dst for r in path]
    return " → ".join(f.upper() for f in fmts)


# --------------------------------------------------------------- Ausführung
def unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def convert(src_path: str, dst_fmt: str, out_dir: str, options: dict | None = None,
            progress: Callable[[float, str], None] | None = None,
            out_name: str | None = None) -> list[str]:
    """Konvertiert eine Datei. Gibt die Liste der erzeugten Dateien zurück."""
    options = options or {}
    progress = progress or (lambda p, m="": None)
    src_path, out_dir = os.path.abspath(src_path), os.path.abspath(out_dir)
    src_fmt = normalize_ext(src_path)
    dst_fmt = normalize_ext(dst_fmt)
    if src_fmt not in FORMATS:
        raise ConversionError(f"Das Format '.{src_fmt}' wird nicht unterstützt.")
    if src_fmt == dst_fmt:
        raise ConversionError("Quell- und Zielformat sind gleich.")
    path = find_path(src_fmt, dst_fmt)
    if not path:
        why = unavailable_targets(src_fmt).get(dst_fmt)
        if why:
            raise ConversionError(f"{label(src_fmt)} → {label(dst_fmt)} {why}.")
        raise ConversionError(f"{label(src_fmt)} lässt sich nicht in {label(dst_fmt)} umwandeln.")

    os.makedirs(out_dir, exist_ok=True)
    stem = out_name or os.path.splitext(os.path.basename(src_path))[0]
    options = dict(options, _title=options.get("_title") or os.path.splitext(os.path.basename(src_path))[0])
    tmp = tempfile.mkdtemp(prefix="ufc_")
    try:
        cur = src_path
        n = len(path)
        for i, r in enumerate(path):
            last = i == n - 1
            lo, hi = i / n, (i + 1) / n

            def step_progress(p, m="", lo=lo, hi=hi):
                progress(lo + (hi - lo) * max(0.0, min(1.0, p)), m)

            ctx = Ctx(r.src, r.dst, options, tmp, step_progress)
            if last:
                if r.multi:
                    target = unique_path(os.path.join(out_dir, f"{stem}_{r.dst}"))
                    os.makedirs(target, exist_ok=True)
                else:
                    target = unique_path(os.path.join(out_dir, f"{stem}.{r.dst}"))
            else:
                target = os.path.join(tmp, f"step{i}.{r.dst}")
            step_progress(0, f"{r.src.upper()} → {r.dst.upper()}")
            try:
                r.func(cur, target, ctx)
            except ConversionError:
                raise
            except Exception as e:  # verständlich machen
                traceback.print_exc()
                raise ConversionError(f"Fehler bei {r.src.upper()} → {r.dst.upper()}: {e}") from e
            if last and r.multi:
                files = sorted(os.path.join(target, f) for f in os.listdir(target))
                if not files:
                    shutil.rmtree(target, ignore_errors=True)
                    raise ConversionError("Es wurde keine Ausgabedatei erzeugt.")
                if len(files) == 1:  # nur eine Datei -> Ordner sparen
                    single = unique_path(os.path.join(out_dir, f"{stem}.{r.dst}"))
                    shutil.move(files[0], single)
                    shutil.rmtree(target, ignore_errors=True)
                    progress(1.0, "Fertig")
                    return [single]
                progress(1.0, "Fertig")
                return [target]
            if not os.path.exists(target) or (os.path.isfile(target) and os.path.getsize(target) == 0):
                raise ConversionError(f"{r.src.upper()} → {r.dst.upper()} hat keine Datei erzeugt.")
            cur = target
        progress(1.0, "Fertig")
        return [cur]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def merge_to_pdf(src_paths: list[str], out_dir: str, name: str = "Zusammengefuegt",
                 options: dict | None = None, progress=None) -> str:
    """Wandelt alle Dateien in PDF um und fügt sie zu einer PDF zusammen."""
    from pypdf import PdfWriter

    progress = progress or (lambda p, m="": None)
    tmp = tempfile.mkdtemp(prefix="ufc_merge_")
    try:
        pdfs = []
        n = len(src_paths)
        for i, p in enumerate(src_paths):
            if normalize_ext(p) == "pdf":
                pdfs.append(p)
            else:
                res = convert(p, "pdf", tmp, options,
                              lambda q, m="", i=i: progress((i + q) / (n + 1), m),
                              out_name=f"{i:03d}")
                pdfs.append(res[0])
            progress((i + 1) / (n + 1), os.path.basename(p))
        w = PdfWriter()
        for p in pdfs:
            w.append(p)
        os.makedirs(out_dir, exist_ok=True)
        out = unique_path(os.path.join(out_dir, f"{name}.pdf"))
        with open(out, "wb") as f:
            w.write(f)
        progress(1.0, "Fertig")
        return out
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
