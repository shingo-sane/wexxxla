"""Eigene Konverter zwischen Word, PowerPoint und Excel (ohne Office nötig)."""
from __future__ import annotations

import datetime as dt
import io
import os
import tempfile

from .engine import Ctx, route, backend_available

# Farben für Tabellenköpfe
ACCENT = (0x2F, 0x55, 0x97)


# =============================================================== Hilfen: Word
def iter_docx_blocks(doc):
    """Absätze und Tabellen in Dokumentreihenfolge."""
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def para_images(par, doc) -> list[bytes]:
    from docx.oxml.ns import qn

    out = []
    for blip in par._element.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid and rid in doc.part.related_parts:
            try:
                out.append(doc.part.related_parts[rid].blob)
            except Exception:
                pass
    return out


def para_kind(par):
    """-> ('title'|'h1'|'h2'|'h3'|'list'|'p', level)"""
    from docx.oxml.ns import qn

    name = (par.style.name if par.style is not None else "") or ""
    low = name.lower()
    if low in ("title", "titel"):
        return "title", 0
    for lvl in (1, 2, 3, 4, 5, 6):
        if low in (f"heading {lvl}", f"überschrift {lvl}"):
            return ("h1", "h2", "h3")[min(lvl, 3) - 1], lvl
    ppr = par._element.pPr
    numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
    if numpr is not None or low.startswith("list") or low.startswith("liste"):
        ilvl = 0
        if numpr is not None:
            il = numpr.find(qn("w:ilvl"))
            if il is not None:
                ilvl = int(il.get(qn("w:val"), 0))
        elif low[-1:].isdigit():
            ilvl = int(low[-1]) - 1
        return "list", max(0, min(ilvl, 4))
    return "p", 0


def table_rows(table) -> list[list[str]]:
    rows = []
    for r in table.rows:
        row = []
        prev = None
        for c in r.cells:
            # verbundene Zellen nicht doppelt übernehmen
            row.append("" if c._tc is prev else c.text.strip())
            prev = c._tc
        rows.append(row)
    return rows


def fmt_value(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Ja" if v else "Nein"
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if isinstance(v, int):
        return str(v)
    if isinstance(v, dt.datetime):
        return v.strftime("%d.%m.%Y %H:%M") if (v.hour or v.minute) else v.strftime("%d.%m.%Y")
    if isinstance(v, dt.date):
        return v.strftime("%d.%m.%Y")
    return str(v)


def read_workbook(path) -> list[tuple[str, list[list]]]:
    """Alle Blätter als (Name, Zeilen) mit getrimmten leeren Rändern."""
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        while rows and all(v in (None, "") for v in rows[-1]):
            rows.pop()
        while rows and all(v in (None, "") for v in rows[0]):
            rows.pop(0)
        if not rows:
            continue
        ncols = max(len(r) for r in rows)
        rows = [r + [None] * (ncols - len(r)) for r in rows]
        # leere Spalten rechts entfernen
        while ncols > 1 and all(r[ncols - 1] in (None, "") for r in rows):
            ncols -= 1
        rows = [r[:ncols] for r in rows]
        sheets.append((ws.title, rows))
    wb.close()
    return sheets


# ========================================================== Hilfen: PowerPoint
def new_presentation():
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width, prs.slide_height = 12192000, 6858000  # 16:9
    return prs


def add_title_slide(prs, title, subtitle=""):
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = title
    if len(s.placeholders) > 1:
        s.placeholders[1].text = subtitle
    return s


def _font_size_for(total_chars: int):
    from pptx.util import Pt

    if total_chars < 250:
        return Pt(24)
    if total_chars < 450:
        return Pt(20)
    if total_chars < 700:
        return Pt(17)
    return Pt(14)


def add_bullet_slide(prs, title, bullets: list[tuple[str, int, bool]]):
    """bullets: (text, level, bold)"""
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = title
    body = s.placeholders[1]
    # Platzhalter auf 16:9 strecken
    body.left, body.top = int(prs.slide_width * 0.06), int(prs.slide_height * 0.22)
    body.width, body.height = int(prs.slide_width * 0.88), int(prs.slide_height * 0.72)
    s.shapes.title.left, s.shapes.title.width = int(prs.slide_width * 0.06), int(prs.slide_width * 0.88)
    tf = body.text_frame
    tf.word_wrap = True
    size = _font_size_for(sum(len(b[0]) for b in bullets))
    for i, (text, level, bold) in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = level
        run = p.add_run()
        run.text = text
        run.font.size = size
        run.font.bold = bold
    return s


def add_table_slide(prs, title, rows: list[list[str]], header=True):
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = title
    s.shapes.title.left, s.shapes.title.width = int(prs.slide_width * 0.04), int(prs.slide_width * 0.92)
    nrows, ncols = len(rows), max(len(r) for r in rows)
    left, top = int(prs.slide_width * 0.04), int(prs.slide_height * 0.2)
    width = int(prs.slide_width * 0.92)
    row_h = min(int(prs.slide_height * 0.75 / max(nrows, 1)), Emu(457200))
    shape = s.shapes.add_table(nrows, ncols, left, top, width, row_h * nrows)
    table = shape.table
    longest = max((len(str(v)) for r in rows for v in r), default=10)
    fsize = Pt(14 if nrows <= 8 and ncols <= 5 and longest < 30 else 11 if nrows <= 14 else 9)
    for r, row in enumerate(rows):
        for c in range(ncols):
            cell = table.cell(r, c)
            cell.text = str(row[c]) if c < len(row) and row[c] is not None else ""
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = fsize
                    if r == 0 and header:
                        run.font.bold = True
            if r == 0 and header:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(*ACCENT)
    return s


def add_picture_slide(prs, title, blob: bytes):
    from PIL import Image

    s = prs.slides.add_slide(prs.slide_layouts[5] if title else prs.slide_layouts[6])
    top = int(prs.slide_height * 0.04)
    if title:
        s.shapes.title.text = title
        top = int(prs.slide_height * 0.2)
    try:
        im = Image.open(io.BytesIO(blob))
        iw, ih = im.size
        if im.format not in ("PNG", "JPEG", "GIF", "BMP", "TIFF"):
            bio = io.BytesIO()
            im.convert("RGBA").save(bio, "PNG")
            blob = bio.getvalue()
    except Exception:
        return None
    maxw, maxh = prs.slide_width * 0.9, prs.slide_height - top - prs.slide_height * 0.05
    r = min(maxw / iw, maxh / ih)
    w, h = int(iw * r), int(ih * r)
    s.shapes.add_picture(io.BytesIO(blob), int((prs.slide_width - w) / 2), top, w, h)
    return s


def chunk(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


# ============================================================ Word → PowerPoint
@route("docx", "pptx", cost=1)
def docx_to_pptx(src, dst, ctx: Ctx):
    from docx import Document
    from docx.table import Table

    doc = Document(src)
    prs = new_presentation()
    title = (doc.core_properties.title or "").strip()
    blocks = list(iter_docx_blocks(doc))
    # Titel aus erstem Titel-Absatz
    if not title:
        for b in blocks[:5]:
            if not isinstance(b, Table) and para_kind(b)[0] == "title" and b.text.strip():
                title = b.text.strip()
                break
    title = title or ctx.opt("_title", "Präsentation")
    add_title_slide(prs, title, (doc.core_properties.author or "").strip())

    MAX_ITEMS, MAX_CHARS = 7, 650
    state = {"title": title, "bullets": [], "cont": False}

    def flush():
        if state["bullets"]:
            t = state["title"] + (" (Forts.)" if state["cont"] else "")
            add_bullet_slide(prs, t, state["bullets"])
            state["bullets"] = []
            state["cont"] = True

    n = len(blocks)
    for i, b in enumerate(blocks):
        if isinstance(b, Table):
            rows = table_rows(b)
            if rows and any(any(c for c in r) for r in rows):
                flush()
                head, body = rows[0], rows[1:] or [[""] * len(rows[0])]
                parts = list(chunk(body, 12))
                for k, part in enumerate(parts):
                    t = state["title"] + (f" ({k + 1}/{len(parts)})" if len(parts) > 1 else "")
                    add_table_slide(prs, t, [head] + part)
            continue
        kind, level = para_kind(b)
        text = b.text.strip()
        imgs = para_images(b, doc)
        if kind == "title" and text == title:
            continue
        if kind in ("title", "h1", "h2") and text:
            flush()
            state.update(title=text, cont=False)
        elif text:
            if kind == "h3":
                item = (text, 0, True)
            elif kind == "list":
                item = (text, min(level, 3), False)
            else:
                item = (text, 0, False)
            cur_chars = sum(len(x[0]) for x in state["bullets"])
            if state["bullets"] and (len(state["bullets"]) >= MAX_ITEMS or cur_chars + len(text) > MAX_CHARS):
                flush()
            state["bullets"].append(item)
        for blob in imgs:
            flush()
            add_picture_slide(prs, state["title"], blob)
            state["cont"] = True
        ctx.progress((i + 1) / n, "Folien werden erstellt")
    flush()
    if len(prs.slides) == 1:
        add_bullet_slide(prs, title, [("(Das Dokument enthält keinen Text.)", 0, False)])
    prs.save(dst)


# ============================================================ PowerPoint → Word
def _iter_shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(sh.shapes)
        else:
            yield sh


def _sorted_shapes(slide):
    return sorted(_iter_shapes(slide.shapes), key=lambda s: ((s.top or 0) // 200000, s.left or 0))


def _slide_title(slide):
    try:
        if slide.shapes.title is not None and slide.shapes.title.text.strip():
            return slide.shapes.title.text.strip()
    except Exception:
        pass
    return ""


def render_pptx_slides(src, ctx: Ctx, dpi=110) -> list:
    """Folien als Bilder (über Office → PDF). Leere Liste, wenn kein Office da ist."""
    if not backend_available("office-impress"):
        return []
    from .office import office_convert
    from .pdf import render_pages
    from .engine import Ctx as C

    tmp = tempfile.mkdtemp(prefix="ufc_slides_")
    try:
        pdf = os.path.join(tmp, "slides.pdf")
        office_convert(src, pdf, C("pptx", "pdf", {}, tmp, lambda p, m="": None), "powerpoint")
        return list(render_pages(pdf, dpi))
    except Exception as e:
        print("Folienbilder nicht möglich:", e)
        return []


@route("pptx", "docx", cost=1)
def pptx_to_docx(src, dst, ctx: Ctx):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(src)
    want_images = str(ctx.opt("ppt_word_images", "1")) not in ("0", "false", "False")
    ctx.progress(0.05, "Folien werden gelesen")
    images = render_pptx_slides(src, ctx) if want_images else []
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2)
    usable = sec.page_width - sec.left_margin - sec.right_margin
    doc.add_heading(ctx.opt("_title", "Präsentation"), level=0)
    n = len(prs.slides)
    for i, slide in enumerate(prs.slides):
        title = _slide_title(slide)
        doc.add_heading(f"Folie {i + 1}" + (f": {title}" if title else ""), level=1)
        if i < len(images):
            bio = io.BytesIO()
            images[i].convert("RGB").save(bio, "JPEG", quality=85)
            bio.seek(0)
            doc.add_picture(bio, width=usable)
        for sh in _sorted_shapes(slide):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                if title and sh.text_frame.text.strip() == title:
                    continue
                for p in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in p.runs).strip()
                    if not t:
                        continue
                    if p.level > 0 or len(sh.text_frame.paragraphs) > 1:
                        style = "List Bullet" if p.level == 0 else f"List Bullet {min(p.level + 1, 3)}"
                        doc.add_paragraph(t, style=style)
                    else:
                        doc.add_paragraph(t)
            elif getattr(sh, "has_table", False) and sh.has_table:
                tbl = sh.table
                rows = [[c.text.strip() for c in r.cells] for r in tbl.rows]
                if rows:
                    ncols = max(len(r) for r in rows)
                    t = doc.add_table(rows=len(rows), cols=ncols)
                    t.style = "Table Grid"
                    for r, row in enumerate(rows):
                        for c, v in enumerate(row):
                            t.cell(r, c).text = v
                            if r == 0:
                                for run in t.cell(r, c).paragraphs[0].runs:
                                    run.bold = True
                    doc.add_paragraph()
            elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE and not images:
                try:
                    blob = sh.image.blob
                    w = min(sh.width or usable, usable)
                    doc.add_picture(io.BytesIO(blob), width=w)
                except Exception:
                    pass
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                p = doc.add_paragraph()
                r = p.add_run("Notizen: ")
                r.bold = True
                r2 = p.add_run(notes)
                r2.italic = True
                r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        ctx.progress(0.3 + 0.7 * (i + 1) / n, f"Folie {i + 1}/{n}")
    doc.save(dst)


# =============================================================== Excel → PPT
@route("xlsx", "pptx", cost=1)
def xlsx_to_pptx(src, dst, ctx: Ctx):
    sheets = read_workbook(src)
    prs = new_presentation()
    add_title_slide(prs, ctx.opt("_title", "Tabelle"), f"{len(sheets)} Tabellenblatt/-blätter")
    ROWS, COLS = 14, 8
    for si, (name, rows) in enumerate(sheets):
        rows = [[fmt_value(v) for v in r] for r in rows]
        head, body = rows[0], rows[1:]
        col_groups = list(chunk(list(range(len(head))), COLS))
        row_parts = list(chunk(body, ROWS)) or [[]]
        total = len(col_groups) * len(row_parts)
        k = 0
        for cg in col_groups:
            for part in row_parts:
                k += 1
                sub = [[r[c] for c in cg] for r in [head] + part]
                t = name + (f" ({k}/{total})" if total > 1 else "")
                add_table_slide(prs, t, sub)
        ctx.progress((si + 1) / max(len(sheets), 1), name)
    if not sheets:
        add_bullet_slide(prs, "Leer", [("Die Tabelle enthält keine Daten.", 0, False)])
    prs.save(dst)


# ============================================================== Excel → Word
@route("xlsx", "docx", cost=1)
def xlsx_to_docx(src, dst, ctx: Ctx):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Pt, Cm, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    sheets = read_workbook(src)
    doc = Document()
    sec = doc.sections[0]
    if any(len(r[0]) > 6 for _, r in sheets):
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin = sec.right_margin = Cm(1.5)
    doc.add_heading(ctx.opt("_title", "Tabelle"), level=0)
    MAX_ROWS = 3000
    for si, (name, rows) in enumerate(sheets):
        doc.add_heading(name, level=1)
        rows = [[fmt_value(v) for v in r] for r in rows[:MAX_ROWS]]
        ncols = len(rows[0])
        table = doc.add_table(rows=len(rows), cols=ncols)
        table.style = "Table Grid"
        size = Pt(10 if ncols <= 6 else 8 if ncols <= 10 else 7)
        for r, row in enumerate(rows):
            cells = table.rows[r].cells
            for c, v in enumerate(row):
                cells[c].text = v
                for p in cells[c].paragraphs:
                    for run in p.runs:
                        run.font.size = size
                        if r == 0:
                            run.bold = True
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                if r == 0:
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), "%02X%02X%02X" % ACCENT)
                    cells[c]._tc.get_or_add_tcPr().append(shd)
        if len(rows) >= MAX_ROWS:
            doc.add_paragraph(f"(Gekürzt auf die ersten {MAX_ROWS} Zeilen.)")
        ctx.progress((si + 1) / max(len(sheets), 1), name)
    doc.save(dst)


# ===================================================== Word / PPT → Excel
def _autosize(ws):
    from openpyxl.utils import get_column_letter

    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 8), 60)


def _write_table(ws, rows, start_row=1):
    from openpyxl.styles import Font, PatternFill
    from .pdf import _to_number

    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            cell = ws.cell(row=start_row + r, column=c + 1, value=_to_number(v) if isinstance(v, str) else v)
            if r == 0:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="%02X%02X%02X" % ACCENT)
    return start_row + len(rows)


@route("docx", "xlsx", cost=1)
def docx_to_xlsx(src, dst, ctx: Ctx):
    from docx import Document
    from openpyxl import Workbook

    doc = Document(src)
    wb = Workbook()
    wb.remove(wb.active)
    for i, t in enumerate(doc.tables):
        ws = wb.create_sheet(f"Tabelle {i + 1}")
        _write_table(ws, table_rows(t))
        _autosize(ws)
    if not doc.tables:
        ws = wb.create_sheet("Text")
        ws.append(["Absatz", "Format", "Text"])
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip():
                ws.append([i + 1, p.style.name if p.style is not None else "", p.text.strip()])
        ws.column_dimensions["C"].width = 100
    wb.save(dst)


@route("pptx", "xlsx", cost=1)
def pptx_to_xlsx(src, dst, ctx: Ctx):
    from openpyxl import Workbook
    from pptx import Presentation

    prs = Presentation(src)
    wb = Workbook()
    wb.remove(wb.active)
    k = 0
    for si, slide in enumerate(prs.slides):
        for sh in _iter_shapes(slide.shapes):
            if getattr(sh, "has_table", False) and sh.has_table:
                k += 1
                ws = wb.create_sheet(f"Folie {si + 1} – T{k}"[:31])
                _write_table(ws, [[c.text.strip() for c in r.cells] for r in sh.table.rows])
                _autosize(ws)
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                k += 1
                ws = wb.create_sheet(f"Folie {si + 1} – Diagramm{k}"[:31])
                try:
                    ch = sh.chart
                    cats = list(ch.plots[0].categories)
                    rows = [["Kategorie"] + [s.name for s in ch.plots[0].series]]
                    vals = [list(s.values) for s in ch.plots[0].series]
                    for ci, cat in enumerate(cats):
                        rows.append([str(cat)] + [v[ci] if ci < len(v) else None for v in vals])
                    _write_table(ws, rows)
                except Exception:
                    ws.append(["Diagrammdaten nicht lesbar"])
    ws = wb.create_sheet("Folientexte", 0)
    ws.append(["Folie", "Titel", "Text", "Notizen"])
    for si, slide in enumerate(prs.slides):
        title = _slide_title(slide)
        texts = [sh.text_frame.text.strip() for sh in _sorted_shapes(slide)
                 if getattr(sh, "has_text_frame", False) and sh.has_text_frame and sh.text_frame.text.strip()
                 and sh.text_frame.text.strip() != title]
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        ws.append([si + 1, title, "\n".join(texts), notes])
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 90
    ws.column_dimensions["D"].width = 50
    wb.save(dst)


# ============================================================== Textausgaben
@route("docx", "txt", cost=1)
def docx_to_txt(src, dst, ctx: Ctx):
    from docx import Document
    from docx.table import Table

    doc = Document(src)
    out = []
    for b in iter_docx_blocks(doc):
        if isinstance(b, Table):
            for r in table_rows(b):
                out.append("\t".join(r))
            out.append("")
        else:
            kind, level = para_kind(b)
            t = b.text
            if kind == "list":
                t = "  " * level + "• " + t
            out.append(t)
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(out).strip() + "\n")


def _pptx_lines(src, markdown=False):
    from pptx import Presentation

    prs = Presentation(src)
    out = []
    for i, slide in enumerate(prs.slides):
        title = _slide_title(slide)
        head = f"Folie {i + 1}" + (f": {title}" if title else "")
        out.append(f"## {head}" if markdown else f"=== {head} ===")
        out.append("")
        for sh in _sorted_shapes(slide):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                if sh.text_frame.text.strip() == title:
                    continue
                for p in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in p.runs).strip()
                    if t:
                        out.append("  " * p.level + ("- " if markdown else "• ") + t)
            elif getattr(sh, "has_table", False) and sh.has_table:
                rows = [[c.text.strip().replace("|", "/") for c in r.cells] for r in sh.table.rows]
                if markdown and rows:
                    out.append("| " + " | ".join(rows[0]) + " |")
                    out.append("|" + "---|" * len(rows[0]))
                    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
                else:
                    out += ["\t".join(r) for r in rows]
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            out.append("")
            out.append(("> Notizen: " if markdown else "Notizen: ") + slide.notes_slide.notes_text_frame.text.strip())
        out.append("")
    return out


@route("pptx", "txt", cost=1)
def pptx_to_txt(src, dst, ctx: Ctx):
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(_pptx_lines(src)))


@route("pptx", "md", cost=1)
def pptx_to_md(src, dst, ctx: Ctx):
    with open(dst, "w", encoding="utf-8") as f:
        f.write(f"# {ctx.opt('_title', 'Präsentation')}\n\n" + "\n".join(_pptx_lines(src, markdown=True)))


@route("txt", "docx", cost=1)
def txt_to_docx(src, dst, ctx: Ctx):
    from docx import Document
    from docx.shared import Pt
    from .data import read_text

    doc = Document()
    text = read_text(src)
    for block in text.replace("\r\n", "\n").split("\n"):
        doc.add_paragraph(block)
    doc.save(dst)
