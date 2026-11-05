"""PDF → PowerPoint und PDF → Word auf Basis der Layout-Analyse (pdflayout)."""
from __future__ import annotations

import copy
import io
import statistics
from collections import Counter

from .pdflayout import PageModel, TextBlock, Para, Table, Run

TABLE_NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "Kein Format, kein Raster"


def _hex(c):
    return "%02X%02X%02X" % tuple(c)


# =========================================================================== PowerPoint
def build_pptx(models: list[PageModel], dst: str, progress=None):
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    W0, H0 = models[0].width, models[0].height
    prs.slide_width, prs.slide_height = Emu(int(W0 * 12700)), Emu(int(H0 * 12700))
    blank = prs.slide_layouts[6]
    for i, m in enumerate(models):
        slide = prs.slides.add_slide(blank)
        s = W0 / m.width
        _pptx_page(slide, m, s)
        if progress:
            progress((i + 1) / len(models), f"Folie {i + 1}/{len(models)}")
    prs.save(dst)


def _E(v, s):
    from pptx.util import Emu
    return Emu(int(round(v * s * 12700)))


def _pptx_page(slide, m: PageModel, s):
    from pptx.dml.color import RGBColor

    if m.bg:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*m.bg)
    filled = [sh for sh in m.shapes if sh.fill is not None]
    strokes = [sh for sh in m.shapes if sh.fill is None]
    filled.sort(key=lambda sh: -(sh.x1 - sh.x0) * (sh.bottom - sh.top))
    for sh in filled:
        _pptx_shape(slide, sh, s)
    for pic in m.pictures + m.graphics_img_notext:
        slide.shapes.add_picture(io.BytesIO(pic.png), _E(pic.x0, s), _E(pic.top, s),
                                 _E(pic.x1 - pic.x0, s), _E(pic.bottom - pic.top, s))
    for sh in strokes:
        _pptx_shape(slide, sh, s)
    for t in m.tables:
        _pptx_table(slide, t, s)
    for b in m.blocks:
        _pptx_textblock(slide, b, s)


def _style_line(line, color, lw, s):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    if color is None:
        line.fill.background()
    else:
        line.color.rgb = RGBColor(*color)
        line.width = Pt(max(lw, 0.25) * s)


def _pptx_shape(slide, sh, s):
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.dml.color import RGBColor

    try:
        if sh.kind == "rect":
            w, h = max(sh.x1 - sh.x0, 0.1), max(sh.bottom - sh.top, 0.1)
            sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _E(sh.x0, s), _E(sh.top, s), _E(w, s), _E(h, s))
        elif sh.kind == "line" and len(sh.pts) >= 2:
            (ax, ay), (bx, by) = sh.pts[0], sh.pts[-1]
            sp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, _E(ax, s), _E(ay, s), _E(bx, s), _E(by, s))
            _style_line(sp.line, sh.stroke, sh.lw, s)
            _drop_style(sp)
            return
        elif sh.kind == "path" and len(sh.pts) >= 2:
            pts = [(int(round(x * s * 12700)), int(round(y * s * 12700))) for x, y in sh.pts]
            fb = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1.0)
            fb.add_line_segments(pts[1:], close=sh.closed or sh.fill is not None)
            sp = fb.convert_to_shape()
        else:
            return
    except Exception as e:
        print("Form übersprungen:", e)
        return
    if sh.fill is not None:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor(*sh.fill)
    else:
        sp.fill.background()
    _style_line(sp.line, sh.stroke, sh.lw, s)
    _drop_style(sp)


def _drop_style(sp):
    """Theme-Stil entfernen (sonst zeigen manche Programme Schatten/Verläufe)."""
    from pptx.oxml.ns import qn

    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)


def _apply_runs(p, runs: list[Run], s, *, keep_size=True):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    for r in runs:
        if not r.text:
            continue
        run = p.add_run()
        run.text = r.text
        f = run.font
        f.name = r.font
        f.size = Pt(max(1, r.size * s))
        f.bold = r.bold
        f.italic = r.italic
        if r.color is not None:
            f.color.rgb = RGBColor(*r.color)


ALIGN_PPT = {"left": 1, "center": 2, "right": 3, "justify": 4}


def _set_align(p, align):
    from pptx.enum.text import PP_ALIGN

    p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
                   "justify": PP_ALIGN.JUSTIFY}.get(align, PP_ALIGN.LEFT)


def _split_bullet(runs):
    """'• Text' -> ('•', Rest-Läufe) oder (None, runs)."""
    from .pdflayout import BULLET_CHARS

    if runs and runs[0].text[:1] in BULLET_CHARS and len(runs[0].text) >= 1:
        first = Run(**runs[0].__dict__)
        bullet = first.text[0]
        first.text = first.text[1:].lstrip()
        rest = ([first] if first.text else []) + [Run(**r.__dict__) for r in runs[1:]]
        if rest:
            rest[0].text = rest[0].text.lstrip()
        return bullet, runs[0], rest
    return None, None, runs


def _para_text_x(para: Para):
    """x-Position, an der der Text nach dem Aufzählungszeichen beginnt."""
    ln = para.lines[0]
    if len(para.lines) > 1:
        return min(l.x0 for l in para.lines[1:])
    t = ln.text
    if len(t) > 2:
        return ln.x0 + (ln.x1 - ln.x0) * (len(t) - len(t[1:].lstrip())) / len(t)
    return ln.x0 + ln.size


def _pptx_textblock(slide, b: TextBlock, s):
    from pptx.util import Pt, Emu
    from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR

    pitch = b.pitch
    first = b.lines[0]
    # Grundlinie ≈ Unterkante - 0,21 × Schriftgröße; PowerPoint setzt sie bei ~0,8 × Zeilenabstand
    baseline = first.bottom - 0.21 * first.size
    top = baseline - 0.8 * max(pitch, first.size * 1.0)
    multi = len(b.lines) > 1
    # Bei mehrzeiligem Text exakt die Originalbreite (gleiche Zeilenumbrüche), sonst großzügig
    width = (b.x1 - b.x0) + (2.0 if multi else (b.x1 - b.x0) * 0.08 + 6)
    height = (b.bottom - top) + 2
    tb = slide.shapes.add_textbox(_E(b.x0 - 1, s), _E(top, s), _E(width, s), _E(height, s))
    tf = tb.text_frame
    tf.word_wrap = multi
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    prev = None
    for k, para in enumerate(b.paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        _set_align(p, para.align)
        ppitch = pitch if len(para.lines) == 1 or abs(para.size - b.lines[0].size) < 0.6 else _para_pitch(para)
        p.line_spacing = Pt(max(ppitch, para.size * 0.9) * s)
        if prev is not None:
            gap = para.lines[0].bottom - prev.lines[-1].bottom - ppitch
            if gap > 0.5:
                p.space_before = Pt(gap * s)
        runs = para.runs()
        bullet, brun, rest = _split_bullet(runs)
        pPr = p._p.get_or_add_pPr()
        if bullet:
            text_x = _para_text_x(para)
            mar = text_x - b.x0 + 1
            ind = para.x0 - text_x
            pPr.set("marL", str(int(mar * s * 12700)))
            pPr.set("indent", str(int(ind * s * 12700)))
            br = Run(**brun.__dict__)
            br.text = bullet + "\t"
            _apply_runs(p, [br] + rest, s)
        else:
            if para.x0 - b.x0 > 2 and para.align == "left":
                pPr.set("marL", str(int((para.x0 - b.x0) * s * 12700)))
            _apply_runs(p, runs, s)
        prev = para


def _para_pitch(para):
    ls = para.lines
    if len(ls) < 2:
        return para.size * 1.2
    d = [b.bottom - a.bottom for a, b in zip(ls, ls[1:])]
    return statistics.median(d)


def _pptx_table(slide, t: Table, s):
    from pptx.util import Pt
    from pptx.enum.text import MSO_ANCHOR
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree

    nr, nc = len(t.ys) - 1, len(t.xs) - 1
    gf = slide.shapes.add_table(nr, nc, _E(t.x0, s), _E(t.top, s), _E(t.x1 - t.x0, s), _E(t.bottom - t.top, s))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    tblPr = tbl._tbl.tblPr
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = TABLE_NO_STYLE
    for i in range(nc):
        tbl.columns[i].width = _E(t.xs[i + 1] - t.xs[i], s)
    for i in range(nr):
        tbl.rows[i].height = _E(t.ys[i + 1] - t.ys[i], s)
    for c in t.cells:
        if c.rs > 1 or c.cs > 1:
            try:
                tbl.cell(c.r, c.c).merge(tbl.cell(c.r + c.rs - 1, c.c + c.cs - 1))
            except Exception:
                pass
    for c in t.cells:
        for rr in range(c.r, c.r + c.rs):
            for cc in range(c.c, c.c + c.cs):
                _pptx_borders(tbl.cell(rr, cc), c.borders, s)
        tc = tbl.cell(c.r, c.c)
        tc.margin_left = tc.margin_right = _E(min(4, max(1, _pad(c))), s)
        tc.margin_top = tc.margin_bottom = _E(1, s)
        tc.vertical_anchor = MSO_ANCHOR.MIDDLE
        if c.fill:
            tc.fill.solid()
            tc.fill.fore_color.rgb = RGBColor(*c.fill)
        else:
            tc.fill.background()
        tf = tc.text_frame
        tf.word_wrap = True
        if not c.paras:
            p = tf.paragraphs[0]
            p.font.size = Pt(4 * s)
            continue
        for k, para in enumerate(c.paras):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            _set_align(p, para.align)
            _apply_runs(p, para.runs(), s)


def _pad(c):
    if not c.paras:
        return 3
    return min(p.x0 for p in c.paras) - c.x0 if c.paras[0].align == "left" else 3


def _pptx_borders(tc, borders, s):
    from pptx.oxml.ns import qn
    from lxml import etree

    tcPr = tc._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for side, tag in reversed((("l", "a:lnL"), ("r", "a:lnR"), ("t", "a:lnT"), ("b", "a:lnB"))):
        bd = (borders or {}).get(side)
        ln = etree.Element(qn(tag))
        if bd:
            color, w = bd
            ln.set("w", str(int(max(w, 0.25) * s * 12700)))
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", _hex(color))
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
        tcPr.insert(0, ln)


# =========================================================================== Word
def build_docx(models: list[PageModel], dst: str, progress=None):
    from docx import Document
    from docx.shared import Pt, Emu

    doc = Document()
    # Grundformat: häufigste Schrift
    fonts, sizes = Counter(), Counter()
    for m in models:
        for b in m.blocks:
            for p in b.paras:
                for r in p.runs():
                    fonts[r.font] += len(r.text)
                    sizes[r.size] += len(r.text)
    normal = doc.styles["Normal"]
    if fonts:
        normal.font.name = fonts.most_common(1)[0][0]
        _set_east_asia_font(normal.element, fonts.most_common(1)[0][0])
    if sizes:
        normal.font.size = Pt(sizes.most_common(1)[0][0])
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    body = sizes.most_common(1)[0][0] if sizes else 11

    # Seitenränder aus dem Inhalt ableiten
    W, H = models[0].width, models[0].height
    lefts, rights, tops = [], [], []
    for m in models:
        items = _page_items(m)
        flow = [it for it in items if it[0] in ("para", "table", "pic")]
        if flow:
            lefts.append(min(it[1] for it in flow))
            rights.append(max(it[3] for it in flow))
        if items:
            tops.append(min(it[2] for it in items))
    ml = max(18.0, min(90.0, min(lefts) - 2)) if lefts else 56
    mr = max(18.0, min(90.0, W - max(rights) - 2)) if rights else 56
    mt = max(0.0, min(90.0, min(tops) - 1)) if tops else 56
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Emu(int(W * 12700)), Emu(int(H * 12700))
    if W > H:
        from docx.enum.section import WD_ORIENT
        sec.orientation = WD_ORIENT.LANDSCAPE
    sec.left_margin, sec.right_margin = Pt(ml), Pt(mr)
    sec.top_margin, sec.bottom_margin = Pt(mt), Pt(18)
    sec.header_distance = sec.footer_distance = Pt(10)

    bgs = {m.bg for m in models}
    if len(bgs) == 1 and None not in bgs:
        _doc_background(doc, bgs.pop())
    ctx = {"doc": doc, "ml": ml, "mr": mr, "mt": mt, "W": W, "body": body,
           "head_sizes": sorted({round(p.size) for m in models for b in m.blocks for p in b.paras
                                 if p.size >= body * 1.25}, reverse=True)}
    for pi, m in enumerate(models):
        _docx_page(doc, m, ctx, pi)
        if progress:
            progress((pi + 1) / len(models), f"Seite {pi + 1}/{len(models)}")
    # Dokument darf nicht mit einer Tabelle enden
    body = doc.element.body
    kids = [k for k in body if not k.tag.endswith("}sectPr")]
    if kids and kids[-1].tag.endswith("}tbl"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
    # leeren Startabsatz entfernen
    first = doc.paragraphs[0] if doc.paragraphs else None
    if first is not None and not first.text and first._p.getprevious() is None and first._p.getnext() is not None \
            and not first._p.getnext().tag.endswith("}tbl") and not first._p.xpath(".//w:drawing"):
        first._p.getparent().remove(first._p)
    doc.save(dst)


def _doc_background(doc, color):
    """Seitenhintergrundfarbe (wird in Word angezeigt und gedruckt)."""
    from docx.oxml.ns import qn

    root = doc.element
    bg = root.makeelement(qn("w:background"), {qn("w:color"): _hex(color)})
    root.insert(0, bg)
    settings = doc.settings.element
    if settings.find(qn("w:displayBackgroundShape")) is None:
        settings.insert(0, settings.makeelement(qn("w:displayBackgroundShape"), {}))


def _set_east_asia_font(el, name):
    from docx.oxml.ns import qn

    rpr = el.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(k), name)


def _page_items(m: PageModel):
    """Alle Fluss-Elemente einer Seite: (art, x0, top, x1, bottom, objekt)."""
    hidden = set(m.word_hidden_lines)
    items = []
    for b in m.blocks:
        for p in b.paras:
            if all(id(l) in hidden for l in p.lines):
                continue
            items.append(("para", p.x0, p.top, p.x1, p.bottom, (p, b)))
    for t in m.tables:
        items.append(("table", t.x0, t.top, t.x1, t.bottom, t))
    for pic in m.pictures:
        if pic in m.graphics_img:
            continue
        items.append(("pic", pic.x0, pic.top, pic.x1, pic.bottom, pic))
    for g in m.graphics_img:
        items.append(("pic", g.x0, g.top, g.x1, g.bottom, g))
    for r in m.rules:
        items.append(("rule", r.x0, r.y - 0.5, r.x1, r.y + 0.5, r))
    # Absätze in Boxen (Infobox, Kopfbalken) zusammenfassen
    for bx in m.boxes:
        inside = [it for it in items if it[0] == "para" and it[1] >= bx.x0 - 1 and it[3] <= bx.x1 + 1
                  and it[2] >= bx.top - 1 and it[4] <= bx.bottom + 1]
        if not inside:
            continue
        items = [it for it in items if it not in inside]
        items.append(("box", bx.x0, bx.top, bx.x1, bx.bottom, (bx, [it[5] for it in inside])))
    # Bilder, die komplett in Grafiken liegen, nicht doppelt
    gboxes = [(g.x0, g.top, g.x1, g.bottom) for g in m.graphics_img]
    def _ov(it, g):
        w = max(0.0, min(it[3], g[2]) - max(it[1], g[0]))
        h = max(0.0, min(it[4], g[3]) - max(it[2], g[1]))
        return w * h / max((it[3] - it[1]) * (it[4] - it[2]), 1e-6)

    items = [it for it in items if not (it[0] == "pic" and it[5].kind == "image" and any(
        _ov(it, g) > 0.5 for g in gboxes))]
    # Text in Grafik-Bildern ausblenden
    items = [it for it in items if not (it[0] == "para" and any(
        it[1] >= g[0] - 1 and it[3] <= g[2] + 1 and it[2] >= g[1] - 1 and it[4] <= g[3] + 1 for g in gboxes))]
    items.sort(key=lambda it: (it[2], it[1]))
    return items


def _bands(items, W):
    """Teilt die Seite in waagrechte Bänder; nebeneinanderliegende Elemente -> Spalten."""
    bands = []
    for it in items:
        if bands and it[2] < bands[-1]["bottom"] - 1:
            b = bands[-1]
            b["items"].append(it)
            b["bottom"] = max(b["bottom"], it[4])
        else:
            bands.append({"items": [it], "bottom": it[4]})
    out = []
    for b in bands:
        its = b["items"]
        cols = _columns(its)
        out.append(cols)
    return out


def _columns(its):
    """Gruppiert Elemente eines Bandes nach überlappenden x-Bereichen."""
    if len(its) == 1:
        return [its]
    ivs = sorted(its, key=lambda it: it[1])
    groups = []
    for it in ivs:
        if groups and it[1] < groups[-1]["x1"] - 2:
            groups[-1]["items"].append(it)
            groups[-1]["x1"] = max(groups[-1]["x1"], it[3])
        else:
            groups.append({"items": [it], "x1": it[3]})
    return [sorted(g["items"], key=lambda it: (it[2], it[1])) for g in groups]


def _docx_page(doc, m: PageModel, ctx, pi):
    from docx.shared import Pt

    items = _page_items(m)
    state = {"y": ctx["mt"], "first": pi > 0, "page_top": True}
    for cols in _bands(items, m.width):
        if len(cols) == 1:
            for it in cols[0]:
                _docx_item(doc, it, m, ctx, state, ctx["ml"], ctx["W"] - ctx["mr"])
        else:
            _docx_side_by_side(doc, cols, m, ctx, state)
    if not items and pi > 0:
        p = doc.add_paragraph()
        p.paragraph_format.page_break_before = True


def _docx_side_by_side(doc, cols, m, ctx, state):
    """Nebeneinanderliegende Inhalte als unsichtbare Tabelle."""
    from docx.shared import Pt

    top = min(it[2] for c in cols for it in c)
    bottom = max(it[4] for c in cols for it in c)
    gap = _spacer(doc, ctx, state, top)
    if gap > 1:
        from docx.enum.text import WD_LINE_SPACING
        sp = _new_para(doc, state)
        sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        sp.paragraph_format.line_spacing = Pt(max(1, min(gap, 400) * 0.97))
    left, right = ctx["ml"], ctx["W"] - ctx["mr"]
    bounds = [left] + [(max(it[3] for it in a) + min(it[1] for it in b)) / 2 for a, b in zip(cols, cols[1:])] + [right]
    _before_table(doc, state)
    table = doc.add_table(rows=1, cols=len(cols))
    _table_layout(table, [bounds[i + 1] - bounds[i] for i in range(len(cols))], 0, borders=False)
    for k, col in enumerate(cols):
        cell = table.cell(0, k)
        _cell_margins(cell, 0, 0)
        cstate = {"y": top, "first": False, "page_top": False}
        p0 = cell.paragraphs[0]
        for it in col:
            _docx_item(cell, it, m, ctx, cstate, bounds[k], bounds[k + 1])
        if not p0.text and not p0._p.xpath(".//w:drawing") and len(cell.paragraphs) > 1:
            p0._p.getparent().remove(p0._p)
    state["y"] = bottom


def _spacer(container, ctx, state, top):
    """Abstand zum vorherigen Element (wird als 'Abstand vor' umgesetzt)."""
    y = state["y"]
    return 0 if y is None else max(0.0, top - y)


def _new_para(container, state):
    p = container.add_paragraph()
    if state.get("first") and state.get("page_top"):
        p.paragraph_format.page_break_before = True
    state["first"] = False
    state["page_top"] = False
    return p


def _docx_item(container, it, m, ctx, state, left, right):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    kind, x0, top, x1, bottom, obj = it
    gap = _spacer(container, ctx, state, top)
    if kind == "para":
        para, block = obj
        _docx_para(container, para, block, m, ctx, state, left, right, gap)
    elif kind == "pic":
        p = _new_para(container, state)
        pf = p.paragraph_format
        pf.space_before = Pt(gap * 0.95)
        pf.left_indent = Pt(max(0, x0 - left))
        avail = right - left
        w = min(x1 - x0, avail)
        p.add_run().add_picture(io.BytesIO(obj.png), width=Pt(w))
        h = (bottom - top) * (w / (x1 - x0))
        state["y"] = top + h + 1
    elif kind == "rule":
        p = _new_para(container, state)
        pf = p.paragraph_format
        pf.space_before = Pt(max(0, gap - 2) * 0.95)
        pf.left_indent = Pt(max(0, x0 - left))
        pf.right_indent = Pt(max(0, right - x1))
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(2)
        _para_border(p, "bottom", obj.color or (0, 0, 0), obj.lw)
        state["y"] = obj.y + 1
    elif kind == "box":
        _docx_box(container, obj, it, m, ctx, state, left, right, gap)
    elif kind == "table":
        if gap > 1:
            sp = _new_para(container, state)
            sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            sp.paragraph_format.line_spacing = Pt(max(1, min(gap, 400) * 0.95))
            if sp.paragraph_format.page_break_before:
                pass
        _docx_table(container, obj, ctx, state, left, right)
        state["y"] = bottom


def _docx_box(container, obj, it, m, ctx, state, left, right, gap):
    """Box mit Hintergrund/Rahmen -> einzellige Tabelle in Originalgröße."""
    from docx.shared import Pt
    from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_LINE_SPACING

    bx, paras = obj
    if gap > 1:
        sp = _new_para(container, state)
        sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        sp.paragraph_format.line_spacing = Pt(max(1, gap * 0.97))
    _before_table(container, state)
    table = container.add_table(rows=1, cols=1)
    _table_layout(table, [bx.x1 - bx.x0], bx.x0 - left, borders=False)
    row = table.rows[0]
    row.height = Pt(bx.bottom - bx.top)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    cell = table.cell(0, 0)
    tx0 = min(p.x0 for p, _ in paras)
    tx1 = max(p.x1 for p, _ in paras)
    padl = max(0.0, min(tx0 - bx.x0, 40))
    _cell_margins(cell, padl, 0)
    if bx.fill:
        _cell_shading(cell, bx.fill)
    side = (bx.stroke, max(bx.lw, 0.5)) if bx.stroke else None
    _cell_borders_docx(cell, {"l": side, "r": side, "t": side, "b": side})
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cstate = {"y": bx.top, "first": False, "page_top": False}
    p0 = cell.paragraphs[0]
    for para, block in sorted(paras, key=lambda pb: pb[0].top):
        _docx_para(cell, para, block, m, ctx, cstate, bx.x0 + padl, bx.x1 - 2, 0, in_box=bx)
    if not p0.text and len(cell.paragraphs) > 1:
        p0._p.getparent().remove(p0._p)
    state["y"] = bx.bottom


def _box_for(m, para):
    for bx in m.boxes:
        if para.x0 >= bx.x0 - 1 and para.x1 <= bx.x1 + 1 and para.top >= bx.top - 1 and para.bottom <= bx.bottom + 1:
            return bx
    return None


def _docx_para(container, para: Para, block: TextBlock, m, ctx, state, left, right, gap, in_box=None):
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    size = para.size
    head_level = None
    if size >= ctx["body"] * 1.25 and len(para.text) < 160:
        hs = ctx["head_sizes"]
        head_level = min(hs.index(round(size)) + 1 if round(size) in hs else 3, 3)
    elif para.lines[0].bold and len(para.lines) == 1 and size >= ctx["body"] * 1.05 and len(para.text) < 120:
        head_level = min(len(ctx["head_sizes"]) + 1, 3)
    p = _new_para(container, state)
    if head_level:
        try:
            p.style = f"Heading {head_level}"
        except Exception:
            pass
    pf = p.paragraph_format
    pitch = _para_pitch(para) if len(para.lines) > 1 else max(block.pitch if len(block.lines) > 1 else size * 1.18, size)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(max(pitch, size * 1.0))
    pf.space_after = Pt(0)
    # Abstand: Oberkante Zeile ≈ Grundlinie - 0,8·Zeilenabstand
    line_top = para.lines[0].bottom - 0.21 * para.lines[0].size - 0.8 * pitch
    gap_eff = 0 if state["y"] is None and gap == 0 else max(0.0, line_top - (state["y"] if state["y"] is not None else line_top))
    pf.space_before = Pt(round(gap_eff * 0.97, 1))
    pf.keep_with_next = bool(head_level)
    box = None
    # Ausrichtung & Einzüge
    align = para.align
    if len(para.lines) == 1 and align == "left":
        c = (para.x0 + para.x1) / 2
        if abs(c - (left + right) / 2) < 3 and para.x0 - left > 15:
            align = "center"
        elif abs(para.x1 - right) < 3 and para.x0 - left > 60:
            align = "right"
        elif in_box and abs(c - (in_box.x0 + in_box.x1) / 2) < 3 and para.x0 - in_box.x0 > 15:
            align = "center"
    pf.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    runs = para.runs()
    bullet, brun, rest = _split_bullet(runs)
    if box:
        ri = max(0.0, right - box.x1 + 5)
        li = max(0.0, para.x0 - left) if align in ("left", "justify") else max(0.0, box.x0 + 5 - left)
    elif align in ("left", "justify"):
        li = max(0.0, para.x0 - left)
        ri = max(0.0, right - block.x1 - 3) if len(para.lines) > 1 else 0.0
    elif align == "center":
        li = ri = 0.0
        if len(para.lines) > 1:
            li, ri = max(0.0, block.x0 - left), max(0.0, right - block.x1 - 3)
    else:  # right
        li, ri = 0.0, max(0.0, right - block.x1)
    if bullet:
        text_x = _para_text_x(para)
        pf.left_indent = Pt(max(0, text_x - left))
        pf.first_line_indent = Pt(para.x0 - text_x)
        pf.tab_stops.add_tab_stop(Pt(max(0, text_x - left)))
        br = Run(**brun.__dict__)
        br.text = bullet + "\t"
        runs = [br] + rest
    else:
        pf.left_indent = Pt(li)
        if len(para.lines) > 1:
            fl = para.lines[0].x0 - min(l.x0 for l in para.lines[1:])
            if abs(fl) > 2 and align in ("left", "justify"):
                pf.first_line_indent = Pt(fl)
    pf.right_indent = Pt(ri)
    if box:
        _para_box(p, box, para, m)
    for r in runs:
        if not r.text:
            continue
        run = p.add_run(r.text)
        f = run.font
        f.name = r.font
        _set_run_fonts(run, r.font)
        f.size = Pt(r.size)
        f.bold = r.bold
        f.italic = r.italic
        f.color.rgb = RGBColor(*(r.color or (0, 0, 0)))
    n_lines = len(para.lines)
    state["y"] = line_top + n_lines * pitch


def _set_run_fonts(run, name):
    from docx.oxml.ns import qn

    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rf)
    for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(k), name)
    for k in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(k)) is not None:
            del rf.attrib[qn(k)]


def _para_border(p, side, color, lw, space=1):
    from docx.oxml.ns import qn

    pPr = p._p.get_or_add_pPr()
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = pPr.makeelement(qn("w:pBdr"), {})
        # pBdr muss vor shd/tabs/spacing… stehen
        anchor = None
        for tag in ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:spacing", "w:ind", "w:jc"):
            anchor = pPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(bdr)
        else:
            pPr.append(bdr)
    el = bdr.makeelement(qn(f"w:{side}"), {qn("w:val"): "single", qn("w:sz"): str(max(2, int(lw * 8))),
                                            qn("w:space"): str(space), qn("w:color"): _hex(color)})
    order = ["top", "left", "bottom", "right"]
    bdr.append(el)
    kids = sorted(list(bdr), key=lambda e: order.index(e.tag.split("}")[1]) if e.tag.split("}")[1] in order else 9)
    for k in kids:
        bdr.remove(k)
        bdr.append(k)


def _para_box(p, box, para, m):
    from docx.oxml.ns import qn

    if box.stroke:
        for side in ("top", "left", "bottom", "right"):
            _para_border(p, side, box.stroke, box.lw or 0.75, space=4)
    if box.fill:
        pPr = p._p.get_or_add_pPr()
        shd = pPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): _hex(box.fill)})
        anchor = None
        for tag in ("w:tabs", "w:suppressAutoHyphens", "w:spacing", "w:ind", "w:jc"):
            anchor = pPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(shd)
        else:
            pPr.append(shd)


def _container_el(container):
    if hasattr(container, "_tc"):
        return container._tc
    return container.element.body


def _before_table(container, state):
    """Zwischen zwei Tabellen muss ein Absatz stehen (sonst verschmelzen sie); Seitenumbruch ebenfalls hier."""
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING

    el = _container_el(container)
    kids = [k for k in el if k.tag.split("}")[1] not in ("sectPr", "tcPr")]
    need = bool(kids) and kids[-1].tag.endswith("}tbl")
    brk = state.get("first") and state.get("page_top")
    if need or brk:
        p = container.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(1)
        pf.space_before = pf.space_after = Pt(0)
        if brk:
            pf.page_break_before = True
    state["first"] = False
    state["page_top"] = False


def _page_break_before_table(table):
    from docx.oxml.ns import qn

    # erste Zelle, erster Absatz: Seitenumbruch davor
    p = table.cell(0, 0).paragraphs[0]
    p.paragraph_format.page_break_before = True


def _table_layout(table, widths_pt, indent_pt, borders=True):
    from docx.shared import Pt
    from docx.oxml.ns import qn

    tbl = table._tbl
    tblPr = tbl.tblPr
    for tag in ("w:tblLayout", "w:tblInd", "w:tblW", "w:tblBorders"):
        for el in tblPr.findall(qn(tag)):
            tblPr.remove(el)
    lay = tblPr.makeelement(qn("w:tblLayout"), {qn("w:type"): "fixed"})
    tblw = tblPr.makeelement(qn("w:tblW"), {qn("w:w"): str(int(sum(widths_pt) * 20)), qn("w:type"): "dxa"})
    ind = tblPr.makeelement(qn("w:tblInd"), {qn("w:w"): str(int(indent_pt * 20)), qn("w:type"): "dxa"})
    style = tblPr.find(qn("w:tblStyle"))
    pos = list(tblPr).index(style) + 1 if style is not None else 0
    for el in (tblw, ind, lay):
        tblPr.insert(pos, el)
        pos += 1
    if not borders:
        bd = tblPr.makeelement(qn("w:tblBorders"), {})
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bd.append(bd.makeelement(qn(f"w:{side}"), {qn("w:val"): "nil"}))
        ind.addnext(bd)
    grid = tbl.tblGrid
    cols = grid.findall(qn("w:gridCol"))
    for gc, w in zip(cols, widths_pt):
        gc.set(qn("w:w"), str(int(w * 20)))
    for row in table.rows:
        for k, cell in enumerate(row.cells):
            if k < len(widths_pt):
                cell.width = Pt(widths_pt[k])


def _cell_margins(cell, lr, tb):
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    mar = tcPr.find(qn("w:tcMar"))
    if mar is None:
        mar = tcPr.makeelement(qn("w:tcMar"), {})
        tcPr.append(mar)
    for side, v in (("top", tb), ("left", lr), ("bottom", tb), ("right", lr)):
        el = mar.find(qn(f"w:{side}"))
        if el is None:
            el = mar.makeelement(qn(f"w:{side}"), {})
            mar.append(el)
        el.set(qn("w:w"), str(int(v * 20)))
        el.set(qn("w:type"), "dxa")


def _cell_shading(cell, color):
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): _hex(color)})
    tcPr.append(shd)


def _cell_borders_docx(cell, borders):
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    bd = tcPr.makeelement(qn("w:tcBorders"), {})
    for side, key in (("top", "t"), ("left", "l"), ("bottom", "b"), ("right", "r")):
        b = (borders or {}).get(key)
        if b:
            color, w = b
            bd.append(bd.makeelement(qn(f"w:{side}"), {qn("w:val"): "single", qn("w:sz"): str(max(2, int(w * 8))),
                                                      qn("w:space"): "0", qn("w:color"): _hex(color)}))
        else:
            bd.append(bd.makeelement(qn(f"w:{side}"), {qn("w:val"): "nil"}))
    # tcBorders steht nach tcW/gridSpan/vMerge
    anchor = None
    for tag in ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"):
        anchor = tcPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(bd)
    else:
        tcPr.append(bd)


def _docx_table(container, t: Table, ctx, state, left, right):
    from docx.shared import Pt, RGBColor
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    nr, nc = len(t.ys) - 1, len(t.xs) - 1
    _before_table(container, state)
    table = container.add_table(rows=nr, cols=nc)
    widths = [t.xs[i + 1] - t.xs[i] for i in range(nc)]
    _table_layout(table, widths, t.x0 - left, borders=False)
    for i, row in enumerate(table.rows):
        row.height = Pt(t.ys[i + 1] - t.ys[i])
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    for c in t.cells:
        if c.rs > 1 or c.cs > 1:
            try:
                table.cell(c.r, c.c).merge(table.cell(c.r + c.rs - 1, c.c + c.cs - 1))
            except Exception:
                pass
    for c in t.cells:
        cell = table.cell(c.r, c.c)
        _cell_margins(cell, min(5, max(1, _pad(c))), 0.5)
        if c.fill:
            _cell_shading(cell, c.fill)
        _cell_borders_docx(cell, c.borders)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p0 = cell.paragraphs[0]
        for k, para in enumerate(c.paras):
            p = p0 if k == 0 else cell.add_paragraph()
            p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                           "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}[para.align]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for r in para.runs():
                run = p.add_run(r.text)
                run.font.name = r.font
                _set_run_fonts(run, r.font)
                run.font.size = Pt(r.size)
                run.font.bold = r.bold
                run.font.italic = r.italic
                run.font.color.rgb = RGBColor(*(r.color or (0, 0, 0)))
        if not c.paras:
            p0.paragraph_format.space_after = Pt(0)
            # leere Zelle: kleine Schrift, damit sie nicht höher wird
            r = p0.add_run("")
            r.font.size = Pt(4)
