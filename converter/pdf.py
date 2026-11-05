"""PDF als Quelle: → Word, PowerPoint, Excel, Text, Bilder."""
from __future__ import annotations

import io
import os
import re
from collections import Counter

from .engine import ConversionError, Ctx, route
from .formats import RASTER


# ------------------------------------------------------------ Hilfsfunktionen
def _open_pdfium(path):
    import pypdfium2 as pdfium

    try:
        return pdfium.PdfDocument(path)
    except Exception as e:
        if "password" in str(e).lower():
            raise ConversionError("Die PDF ist passwortgeschützt.")
        raise ConversionError(f"PDF kann nicht gelesen werden: {e}")


def render_pages(path: str, dpi: int = 150, progress=None):
    """Liefert für jede Seite ein PIL-Bild."""
    pdf = _open_pdfium(path)
    try:
        n = len(pdf)
        for i in range(n):
            page = pdf[i]
            img = page.render(scale=dpi / 72).to_pil()
            page.close()
            if progress:
                progress((i + 1) / n, f"Seite {i + 1}/{n}")
            yield img
    finally:
        pdf.close()


def page_sizes(path: str) -> list[tuple[float, float]]:
    pdf = _open_pdfium(path)
    try:
        return [pdf[i].get_size() for i in range(len(pdf))]
    finally:
        pdf.close()


def _is_bold(fontname: str) -> bool:
    f = (fontname or "").lower()
    return any(k in f for k in ("bold", "black", "heavy", "semibold", "demi"))


def _is_italic(fontname: str) -> bool:
    f = (fontname or "").lower()
    return "italic" in f or "oblique" in f


_PUA_BULLETS = {0xF0B7: "•", 0xF0A7: "▪", 0xF076: "•", 0xF0D8: "►", 0xF0FC: "✓", 0xF06C: "•", 0xF0A8: "□"}
BULLET_RE = r"^\s*([•▪●◦►✓□*–-]|\d{1,2}[.)])\s+"


def _clean(s: str) -> str:
    s = (s or "").translate(_PUA_BULLETS)
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)


def _inside(obj, bbox, pad=1.0):
    x0, top, x1, bottom = bbox
    cx = (obj["x0"] + obj["x1"]) / 2
    cy = (obj["top"] + obj["bottom"]) / 2
    return x0 - pad <= cx <= x1 + pad and top - pad <= cy <= bottom + pad


def extract_layout(pdf_path: str, progress=None):
    """Liest eine PDF strukturiert aus: pro Seite Zeilen, Tabellen und Bildbereiche."""
    import pdfplumber

    pages = []
    try:
        doc = pdfplumber.open(pdf_path)
    except Exception as e:
        raise ConversionError(f"PDF kann nicht gelesen werden: {e}")
    with doc:
        n = len(doc.pages)
        for pi, page in enumerate(doc.pages):
            tables = []
            try:
                for t in page.find_tables():
                    rows = t.extract()
                    if rows and len(rows) >= 1 and max(len(r) for r in rows) >= 2:
                        tables.append({"bbox": t.bbox, "rows": [[_clean(c or "") for c in r] for r in rows]})
            except Exception:
                pass
            tboxes = [t["bbox"] for t in tables]
            lines = []
            try:
                raw_lines = page.extract_text_lines(return_chars=True, strip=True)
            except Exception:
                raw_lines = []
            for ln in raw_lines:
                if any(_inside(ln, b) for b in tboxes):
                    continue
                chars = [c for c in ln.get("chars", []) if c.get("text", "").strip()]
                if not chars:
                    continue
                sizes = Counter(round(c.get("size", 10), 1) for c in chars)
                size = sizes.most_common(1)[0][0]
                bold = sum(_is_bold(c.get("fontname", "")) for c in chars) > len(chars) / 2
                italic = sum(_is_italic(c.get("fontname", "")) for c in chars) > len(chars) / 2
                color = None
                col = chars[0].get("non_stroking_color")
                if isinstance(col, (list, tuple)) and len(col) == 3 and all(isinstance(v, (int, float)) for v in col):
                    color = tuple(int(max(0, min(1, v)) * 255) for v in col)
                lines.append({
                    "text": _clean(ln["text"]), "x0": ln["x0"], "x1": ln["x1"],
                    "top": ln["top"], "bottom": ln["bottom"], "size": size,
                    "bold": bold, "italic": italic, "color": color,
                })
            images = []
            for im in page.images:
                x0, top = max(0, im["x0"]), max(0, im["top"])
                x1, bottom = min(page.width, im["x1"]), min(page.height, im["bottom"])
                if (x1 - x0) > 20 and (bottom - top) > 20:
                    images.append({"x0": x0, "top": top, "x1": x1, "bottom": bottom})
            pages.append({"width": page.width, "height": page.height, "lines": lines,
                          "tables": tables, "images": images})
            if progress:
                progress((pi + 1) / n, f"Analysiere Seite {pi + 1}/{n}")
    return pages


def _crop_images(pdf_path, pages, dpi=150):
    """Schneidet die Bildbereiche aus gerenderten Seiten aus -> {(seite, idx): PNG-bytes}."""
    crops = {}
    needed = [i for i, p in enumerate(pages) if p["images"]]
    if not needed:
        return crops
    pdf = _open_pdfium(pdf_path)
    try:
        scale = dpi / 72
        for i in needed:
            page = pdf[i]
            img = page.render(scale=scale).to_pil().convert("RGB")
            page.close()
            for k, b in enumerate(pages[i]["images"]):
                box = tuple(int(v * scale) for v in (b["x0"], b["top"], b["x1"], b["bottom"]))
                if box[2] - box[0] < 5 or box[3] - box[1] < 5:
                    continue
                bio = io.BytesIO()
                img.crop(box).save(bio, "PNG")
                crops[(i, k)] = bio.getvalue()
    finally:
        pdf.close()
    return crops


def _body_size(pages) -> float:
    sizes = Counter()
    for p in pages:
        for ln in p["lines"]:
            sizes[ln["size"]] += len(ln["text"])
    return sizes.most_common(1)[0][0] if sizes else 11.0


def _paragraphs(lines, body):
    """Fasst Zeilen zu Absätzen zusammen."""
    paras = []
    cur = None
    for ln in sorted(lines, key=lambda l: (round(l["top"]), l["x0"])):
        if cur is not None:
            gap = ln["top"] - cur["bottom"]
            same_style = abs(ln["size"] - cur["size"]) < 0.6 and ln["bold"] == cur["bold"]
            is_bullet = bool(re.match(BULLET_RE, ln["text"]))
            if same_style and 0 <= gap < cur["size"] * 0.9 and not is_bullet and cur["size"] < body * 1.25:
                joiner = "" if cur["text"].endswith("-") else " "
                if joiner == "":
                    cur["text"] = cur["text"][:-1]
                cur["text"] += joiner + ln["text"]
                cur["bottom"] = ln["bottom"]
                cur["x1"] = max(cur["x1"], ln["x1"])
                continue
            paras.append(cur)
        cur = dict(ln)
    if cur:
        paras.append(cur)
    return paras


# ---------------------------------------------------------------- PDF → Word
@route("pdf", "docx", cost=1)
def pdf_to_docx(src, dst, ctx: Ctx):
    mode = ctx.opt("pdf_word_mode", "auto")
    if mode == "pdf2docx":
        try:
            from pdf2docx import Converter

            ctx.progress(0.05, "pdf2docx analysiert Layout …")
            cv = Converter(src)
            try:
                cv.convert(dst)
            finally:
                cv.close()
            if os.path.exists(dst) and os.path.getsize(dst) > 0:
                return
        except Exception as e:
            print("pdf2docx nicht verfügbar, nutze eingebaute Engine:", e)
    from .pdflayout import analyze
    from .pdf_to_office import build_docx

    _check_pdf(src)
    models = analyze(src, lambda p, m: ctx.progress(p * 0.7, m))
    build_docx(models, dst, lambda p, m: ctx.progress(0.7 + p * 0.3, m))


def _check_pdf(src):
    pdf = _open_pdfium(src)
    n = len(pdf)
    pdf.close()
    if n == 0:
        raise ConversionError("Die PDF hat keine Seiten.")


# ---------------------------------------------------------- PDF → PowerPoint
@route("pdf", "pptx", cost=1)
def pdf_to_pptx(src, dst, ctx: Ctx):
    from pptx import Presentation
    from pptx.util import Emu

    _check_pdf(src)
    mode = ctx.opt("pdf_ppt_mode", "editable")
    if mode != "image":
        from .pdflayout import analyze
        from .pdf_to_office import build_pptx

        models = analyze(src, lambda p, m: ctx.progress(p * 0.7, m))
        build_pptx(models, dst, lambda p, m: ctx.progress(0.7 + p * 0.3, m))
        return
    # Modus "Originalgetreu": jede Seite als Bild, Text in den Notizen
    sizes = page_sizes(src)
    w, h = sizes[0]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(w * 12700)), Emu(int(h * 12700))
    blank = prs.slide_layouts[6]
    texts = []
    try:
        import pdfplumber
        with pdfplumber.open(src) as d:
            texts = [(p.extract_text() or "") for p in d.pages]
    except Exception:
        pass
    dpi = max(150, int(ctx.opt("pdf_dpi", 200)))
    for i, img in enumerate(render_pages(src, dpi, lambda p, m: ctx.progress(p, m))):
        slide = prs.slides.add_slide(blank)
        bio = io.BytesIO()
        img.convert("RGB").save(bio, "PNG", optimize=True)
        bio.seek(0)
        slide.shapes.add_picture(bio, 0, 0, prs.slide_width, prs.slide_height)
        if i < len(texts) and texts[i].strip():
            slide.notes_slide.notes_text_frame.text = texts[i]
    prs.save(dst)


# --------------------------------------------------------------- PDF → Excel
def _to_number(v: str):
    s = (v or "").strip().replace(" ", "").replace(" ", "")
    if not s:
        return v
    if re.fullmatch(r"-?\d+", s):
        try:
            return int(s)
        except ValueError:
            return v
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):  # 12.500 = deutsches Tausendertrennzeichen
        return int(s.replace(".", ""))
    if re.fullmatch(r"-?\d+[.,]\d+", s):
        return float(s.replace(",", "."))
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+,\d+", s):  # 1.234,56
        return float(s.replace(".", "").replace(",", "."))
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+\.\d+", s):  # 1,234.56
        return float(s.replace(",", ""))
    return v


@route("pdf", "xlsx", cost=1)
def pdf_to_xlsx(src, dst, ctx: Ctx):
    """Tabellen (mit und ohne Linien, inkl. verbundener Zellen und Farben) und Text je Seite auf ein Blatt."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from .pdflayout import analyze

    _check_pdf(src)
    models = analyze(src, lambda p, m: ctx.progress(p * 0.8, m))
    wb = Workbook()
    wb.remove(wb.active)
    any_table = any(m.tables for m in models)
    for pi, m in enumerate(models):
        ws = wb.create_sheet(f"Seite {pi + 1}")
        items = [(t.top, "t", t) for t in m.tables]
        if True:
            hidden = set(m.word_hidden_lines)
            for b in m.blocks:
                for para in b.paras:
                    if not all(id(l) in hidden for l in para.lines):
                        items.append((para.top, "p", para))
        items.sort(key=lambda x: x[0])
        row = 1
        widths = {}
        for _, kind, obj in items:
            if kind == "p":
                r0 = obj.runs()[0] if obj.runs() else None
                cell = ws.cell(row=row, column=1, value=_to_number(obj.text))
                if r0:
                    cell.font = Font(name=r0.font, size=r0.size, bold=r0.bold, italic=r0.italic,
                                     color="%02X%02X%02X" % r0.color if r0.color else None)
                row += 1
                continue
            t = obj
            for c in t.cells:
                r, k = row + c.r, 1 + c.c
                text = "\n".join(p.text for p in c.paras)
                cell = ws.cell(row=r, column=k, value=_to_number(text) if text else None)
                if c.paras and c.paras[0].runs():
                    r0 = c.paras[0].runs()[0]
                    cell.font = Font(name=r0.font, size=r0.size, bold=r0.bold, italic=r0.italic,
                                     color="%02X%02X%02X" % r0.color if r0.color else None)
                    al = c.paras[0].align
                    cell.alignment = Alignment(horizontal={"justify": "left"}.get(al, al), vertical="center",
                                               wrap_text=len(c.paras) > 1)
                if c.fill:
                    cell.fill = PatternFill("solid", fgColor="%02X%02X%02X" % c.fill)
                if t.ruled:
                    sides = {}
                    for key, name in (("l", "left"), ("r", "right"), ("t", "top"), ("b", "bottom")):
                        bd = c.borders.get(key)
                        sides[name] = Side(style="thin", color="%02X%02X%02X" % bd[0]) if bd else Side()
                    cell.border = Border(**sides)
                if c.rs > 1 or c.cs > 1:
                    ws.merge_cells(start_row=r, start_column=k, end_row=r + c.rs - 1, end_column=k + c.cs - 1)
                if c.cs == 1:
                    widths[k] = max(widths.get(k, 0), (t.xs[c.c + 1] - t.xs[c.c]) / 5.6)
            row += len(t.ys)
        for k, w in widths.items():
            ws.column_dimensions[get_column_letter(k)].width = min(max(w, 6), 80)
        if 1 not in widths:
            ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 10, 12)
        ctx.progress(0.8 + 0.2 * (pi + 1) / len(models), f"Seite {pi + 1}")
    # leere Blätter entfernen, solange eines übrig bleibt
    for name in list(wb.sheetnames):
        if wb[name].max_row <= 1 and wb[name]["A1"].value is None and len(wb.sheetnames) > 1:
            wb.remove(wb[name])
    wb.save(dst)


# ---------------------------------------------------------------- PDF → Text
@route("pdf", "txt", cost=1)
def pdf_to_txt(src, dst, ctx: Ctx):
    import pdfplumber

    out = []
    with pdfplumber.open(src) as doc:
        n = len(doc.pages)
        for i, page in enumerate(doc.pages):
            out.append(page.extract_text() or "")
            ctx.progress((i + 1) / n, f"Seite {i + 1}/{n}")
    text = "\n\n".join(out).strip()
    if not text:
        text = "(Diese PDF enthält keinen auslesbaren Text – vermutlich ein Scan.)"
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text + "\n")


# --------------------------------------------------------------- PDF → Bilder
def _save_image(img, path, fmt, ctx: Ctx):
    from .media import save_pil

    save_pil(img, path, fmt, ctx)


@route("pdf", [f for f in RASTER if f not in ("tiff",)], cost=1, multi=True)
def pdf_to_images(src, dst_dir, ctx: Ctx):
    dpi = int(ctx.opt("pdf_dpi", 150))
    stem = "Seite"
    for i, img in enumerate(render_pages(src, dpi, lambda p, m: ctx.progress(p, m))):
        _save_image(img, os.path.join(dst_dir, f"{stem}_{i + 1:03d}.{ctx.dst_ext}"), ctx.dst_ext, ctx)


@route("pdf", "tiff", cost=1)
def pdf_to_tiff(src, dst, ctx: Ctx):
    dpi = int(ctx.opt("pdf_dpi", 150))
    imgs = [im.convert("RGB") for im in render_pages(src, dpi, lambda p, m: ctx.progress(p, m))]
    imgs[0].save(dst, "TIFF", save_all=True, append_images=imgs[1:], compression="tiff_deflate")
