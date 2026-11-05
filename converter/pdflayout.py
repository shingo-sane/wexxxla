"""Layout-Analyse für PDF-Seiten.

Erkennt pro Seite:
  * Textblöcke mit Absätzen, Zeilen und Formatierung (Schriftart, Größe, fett, kursiv, Farbe, Ausrichtung)
  * Tabellen mit Linien (inkl. verbundener Zellen, Zellfarben, Rahmen) und Tabellen ohne Linien
  * Vektorgrafiken (Rechtecke, Linien, Kurven) – einfache werden als Formen, komplexe als Bild übernommen
  * Bilder
Daraus bauen pdf.py die PowerPoint- und Word-Dateien.
"""
from __future__ import annotations

import io
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from functools import lru_cache

# =================================================================== Modelle


@dataclass
class Run:
    text: str
    font: str
    size: float
    bold: bool
    italic: bool
    color: tuple | None


@dataclass
class Line:
    x0: float
    x1: float
    top: float
    bottom: float
    runs: list
    size: float

    @property
    def text(self):
        return "".join(r.text for r in self.runs)

    @property
    def bold(self):
        chars = [(len(r.text.strip()), r.bold) for r in self.runs if r.text.strip()]
        total = sum(n for n, _ in chars) or 1
        return sum(n for n, b in chars if b) / total > 0.6


@dataclass
class Para:
    lines: list
    align: str = "left"  # left | center | right | justify

    @property
    def x0(self):
        return min(l.x0 for l in self.lines)

    @property
    def x1(self):
        return max(l.x1 for l in self.lines)

    @property
    def top(self):
        return self.lines[0].top

    @property
    def bottom(self):
        return self.lines[-1].bottom

    @property
    def size(self):
        return Counter(round(l.size, 1) for l in self.lines).most_common(1)[0][0]

    @property
    def text(self):
        return " ".join(l.text for l in self.lines)

    def runs(self):
        """Alle Läufe, Zeilen mit Leerzeichen (bzw. Trennstrich-Auflösung) verbunden."""
        out: list[Run] = []
        for i, ln in enumerate(self.lines):
            runs = [Run(**r.__dict__) for r in ln.runs]
            if i > 0 and out:
                prev = out[-1]
                if prev.text.endswith("-") and len(prev.text) > 1 and prev.text[-2].isalpha() \
                        and runs and runs[0].text[:1].islower():
                    prev.text = prev.text[:-1]
                elif not prev.text.endswith(" "):
                    prev.text += " "
            for r in runs:
                if out and _same_fmt(out[-1], r):
                    out[-1].text += r.text
                else:
                    out.append(r)
        return out


@dataclass
class TextBlock:
    paras: list

    @property
    def lines(self):
        return [l for p in self.paras for l in p.lines]

    @property
    def x0(self):
        return min(p.x0 for p in self.paras)

    @property
    def x1(self):
        return max(p.x1 for p in self.paras)

    @property
    def top(self):
        return self.paras[0].top

    @property
    def bottom(self):
        return self.paras[-1].bottom

    @property
    def pitch(self):
        ls = self.lines
        if len(ls) < 2:
            return ls[0].size * 1.2
        diffs = [b.bottom - a.bottom for a, b in zip(ls, ls[1:]) if 0 < b.bottom - a.bottom < a.size * 2.2]
        return statistics.median(diffs) if diffs else ls[0].size * 1.2


@dataclass
class Cell:
    r: int
    c: int
    rs: int
    cs: int
    x0: float
    top: float
    x1: float
    bottom: float
    paras: list = field(default_factory=list)
    fill: tuple | None = None
    borders: dict = field(default_factory=dict)  # side -> (color, width) | None
    valign: str = "middle"


@dataclass
class Table:
    xs: list
    ys: list
    cells: list
    ruled: bool

    @property
    def x0(self):
        return self.xs[0]

    @property
    def x1(self):
        return self.xs[-1]

    @property
    def top(self):
        return self.ys[0]

    @property
    def bottom(self):
        return self.ys[-1]


@dataclass
class Picture:
    x0: float
    top: float
    x1: float
    bottom: float
    png: bytes
    kind: str = "image"  # image | graphic


@dataclass
class Shape:
    kind: str  # rect | line | path
    x0: float
    top: float
    x1: float
    bottom: float
    pts: list = field(default_factory=list)  # für line/path: [(x, top), ...]
    fill: tuple | None = None
    stroke: tuple | None = None
    lw: float = 0.0
    closed: bool = False


@dataclass
class Box:
    """Rechteck mit Füllung/Rahmen, das Text umschließt (Infobox, Kopfbalken)."""
    x0: float
    top: float
    x1: float
    bottom: float
    fill: tuple | None
    stroke: tuple | None
    lw: float


@dataclass
class Rule:
    """Waagrechte Trennlinie."""
    x0: float
    x1: float
    y: float
    color: tuple | None
    lw: float


@dataclass
class PageModel:
    width: float
    height: float
    bg: tuple | None = None
    blocks: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    pictures: list = field(default_factory=list)        # Bilder (ohne darüberliegenden Text)
    graphics_img: list = field(default_factory=list)    # komplexe Grafiken (für Word, mit Text)
    graphics_img_notext: list = field(default_factory=list)  # komplexe Grafiken (für PPT, ohne Text)
    shapes: list = field(default_factory=list)          # einfache Vektorgrafik (für PPT)
    boxes: list = field(default_factory=list)
    rules: list = field(default_factory=list)
    word_hidden_lines: list = field(default_factory=list)  # Zeilen, die in Word-Grafikbildern stecken


# =================================================================== Hilfen
def _same_fmt(a: Run, b: Run):
    return (a.font, round(a.size * 2), a.bold, a.italic, a.color) == (b.font, round(b.size * 2), b.bold, b.italic, b.color)


def to_rgb(c) -> tuple | None:
    if c is None or isinstance(c, str):
        return None
    try:
        vals = [float(v) for v in (c if isinstance(c, (list, tuple)) else [c])]
    except (TypeError, ValueError):
        return None
    if any(v > 1.0001 for v in vals):
        vals = [v / 255 for v in vals]
    if len(vals) == 1:
        g = int(round(vals[0] * 255))
        return (g, g, g)
    if len(vals) == 3:
        return tuple(int(round(max(0, min(1, v)) * 255)) for v in vals)
    if len(vals) == 4:
        cc, m, y, k = vals
        return tuple(int(round(255 * (1 - min(1, v)) * (1 - k))) for v in (cc, m, y))
    return None


FONT_MAP = {
    "helvetica": "Arial", "arial": "Arial", "arimo": "Arial", "liberationsans": "Arial",
    "helveticaneue": "Arial", "nimbussans": "Arial", "nimbussansl": "Arial",
    "times": "Times New Roman", "timesroman": "Times New Roman", "timesnewroman": "Times New Roman",
    "tinos": "Times New Roman", "liberationserif": "Times New Roman", "nimbusroman": "Times New Roman",
    "courier": "Courier New", "couriernew": "Courier New", "cousine": "Courier New",
    "liberationmono": "Courier New", "nimbusmono": "Courier New",
    "carlito": "Calibri", "caladea": "Cambria", "dejavusans": "Verdana", "dejavuserif": "Georgia",
    "dejavusansmono": "Consolas", "symbol": "Symbol", "zapfdingbats": "Wingdings",
    "calibri": "Calibri", "cambria": "Cambria", "candara": "Candara", "consolas": "Consolas",
    "constantia": "Constantia", "corbel": "Corbel", "segoeui": "Segoe UI", "tahoma": "Tahoma",
    "verdana": "Verdana", "georgia": "Georgia", "garamond": "Garamond", "trebuchetms": "Trebuchet MS",
    "bookantiqua": "Book Antiqua", "centurygothic": "Century Gothic", "palatinolinotype": "Palatino Linotype",
    "aptos": "Aptos", "gillsansmt": "Gill Sans MT", "franklingothic": "Franklin Gothic Book",
}
STYLE_WORDS = ("bolditalic", "boldoblique", "semibold", "demibold", "extrabold", "ultrabold", "bold",
               "italic", "oblique", "regular", "roman", "medium", "book", "normal", "black", "heavy",
               "light", "thin", "condensed", "it", "bd", "bi", "rg")


@lru_cache(maxsize=512)
def font_info(fontname: str) -> tuple[str, bool, bool]:
    """'ABCDEF+TimesNewRomanPS-BoldItalicMT' -> ('Times New Roman', True, True)"""
    n = (fontname or "Arial").split("+", 1)[-1]
    low = n.lower()
    bold = any(k in low for k in ("bold", "black", "heavy", "semibold", "demi", ",bd", "-bd"))
    italic = any(k in low for k in ("italic", "oblique", ",it", "-it")) or low.endswith("-bi")
    light = "light" in low
    base = re.split(r"[-,_]", n)[0]
    base = re.sub(r"(PSMT|PS|MT)$", "", base)
    # Stilwörter am Ende des Grundnamens entfernen (z. B. 'ArialBold')
    changed = True
    while changed:
        changed = False
        for w in STYLE_WORDS:
            if len(base) > len(w) + 2 and base.lower().endswith(w) and w not in ("it", "bd", "bi", "rg"):
                base = base[: -len(w)]
                changed = True
    key = re.sub(r"[^a-z]", "", base.lower())
    fam = FONT_MAP.get(key)
    if not fam:
        fam = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base).strip() or "Arial"
        if light and "light" not in fam.lower():
            fam += " Light"
    elif light and fam in ("Calibri", "Segoe UI", "Aptos"):
        fam += " Light"
    return fam, bold, italic


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _inside(x0, top, x1, bottom, bbox, pad=1.5):
    cx, cy = (x0 + x1) / 2, (top + bottom) / 2
    return bbox[0] - pad <= cx <= bbox[2] + pad and bbox[1] - pad <= cy <= bbox[3] + pad


def _cluster(values, tol):
    vals = sorted(values)
    groups = []
    for v in vals:
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _snap(v, edges):
    return min(range(len(edges)), key=lambda i: abs(edges[i] - v))


BULLET_CHARS = "•▪●◦►✓□■–-*‣⁃"
_PUA = {0xF0B7: "•", 0xF0A7: "▪", 0xF076: "•", 0xF0D8: "►", 0xF0FC: "✓", 0xF06C: "•", 0xF0A8: "□", 0xF06E: "■"}


_CID_BULLETS = {"(cid:127)": "•", "(cid:149)": "•", "(cid:183)": "•", "(cid:108)": "•", "(cid:110)": "■"}


def _upright(c):
    m = c.get("matrix")
    if m and len(m) >= 4:
        return abs(m[1]) < 0.02 and abs(m[2]) < 0.02 and m[0] > 0 and m[3] > 0
    return c.get("upright", True)


def _char_text(ch):
    t = ch.get("text", "")
    if t.startswith("(cid:"):
        return _CID_BULLETS.get(t, "")
    if len(t) == 1 and ord(t) in _PUA:
        return _PUA[ord(t)]
    if t and ord(t[0]) < 32:
        return ""
    return t


# =================================================================== Text → Zeilen
def chars_to_lines(chars, split_gaps=True) -> list[Line]:
    """Gruppiert Zeichen zu Zeilen; trennt an großen Lücken (Spalten) in eigene Segmente."""
    chars = [c for c in chars if _upright(c) and _char_text(c) != ""]
    if not chars:
        return []
    rows: list[dict] = []
    for c in sorted(chars, key=lambda c: (round(c["top"], 1), c["x0"])):
        h = max(c["bottom"] - c["top"], 0.5)
        best = None
        for r in reversed(rows[-6:]):
            ov = _overlap(c["top"], c["bottom"], r["top"], r["bottom"])
            if ov >= 0.5 * min(h, r["bottom"] - r["top"]):
                best = r
                break
        if best is None:
            rows.append({"top": c["top"], "bottom": c["bottom"], "chars": [c]})
        else:
            best["chars"].append(c)
            # Hoch-/Tiefstellungen sollen die Zeile nicht aufblähen
            if c["size"] >= 0.8 * statistics.median(x["size"] for x in best["chars"]):
                best["top"] = min(best["top"], c["top"])
                best["bottom"] = max(best["bottom"], c["bottom"])
    lines: list[Line] = []
    for r in rows:
        cs = sorted(r["chars"], key=lambda c: c["x0"])
        seg: list = []
        segs = []
        for c in cs:
            if seg:
                prev = seg[-1]
                gap = c["x0"] - prev["x1"]
                size = max(prev["size"], c["size"], 1)
                seg_text = "".join(_char_text(x) for x in seg).strip()
                if split_gaps and gap > max(size * 1.1, 7) and _char_text(prev).strip() and _char_text(c).strip() \
                        and not (len(seg_text) <= 2 and (seg_text[:1] in BULLET_CHARS or re.fullmatch(r"\d+[.)]|[a-z][.)]", seg_text))):
                    segs.append(seg)
                    seg = []
            seg.append(c)
        if seg:
            segs.append(seg)
        for s in segs:
            ln = _make_line(s)
            if ln and ln.text.strip():
                lines.append(ln)
    lines.sort(key=lambda l: (round(l.top), l.x0))
    return lines


def _make_line(chars) -> Line | None:
    runs: list[Run] = []
    prev = None
    for c in chars:
        t = _char_text(c)
        if prev is not None:
            gap = c["x0"] - prev["x1"]
            if gap > max(prev["size"], c["size"]) * 0.17 and t != " " and not runs[-1].text.endswith(" "):
                runs[-1].text += " "
        fam, b, i = font_info(c.get("fontname", ""))
        col = to_rgb(c.get("non_stroking_color"))
        r = Run(t, fam, round(float(c.get("size", 10)), 1), b, i, col)
        if t == " " and runs:
            if not runs[-1].text.endswith(" "):
                runs[-1].text += " "
        elif runs and _same_fmt(runs[-1], r):
            runs[-1].text += t
        else:
            runs.append(r)
        prev = c
    if not runs:
        return None
    runs[-1].text = runs[-1].text.rstrip()
    while runs and not runs[0].text.strip():
        runs.pop(0)
    if not runs:
        return None
    runs[0].text = runs[0].text.lstrip()
    sizes = Counter()
    for c in chars:
        sizes[round(float(c.get("size", 10)), 1)] += 1
    vis = [c for c in chars if _char_text(c).strip()] or chars
    return Line(min(c["x0"] for c in vis), max(c["x1"] for c in vis), min(c["top"] for c in vis),
                max(c["bottom"] for c in vis), runs, sizes.most_common(1)[0][0])


# =================================================================== Zeilen → Blöcke/Absätze
def _is_bullet(text: str) -> bool:
    t = text.lstrip()
    return bool(t) and (t[0] in BULLET_CHARS and (len(t) == 1 or t[1] in " \t")
                        or re.match(r"^(\d{1,2}|[a-zA-Z])[.)]\s", t) is not None)


def lines_to_blocks(lines: list[Line]) -> list[TextBlock]:
    blocks: list[list[Line]] = []
    for ln in sorted(lines, key=lambda l: (l.top, l.x0)):
        best, best_gap = None, None
        for b in blocks:
            last = b[-1]
            gap = ln.top - last.bottom
            size = max(last.size, ln.size)
            if gap < -size * 0.3 or gap > size * 0.95:
                continue
            bx0, bx1 = min(l.x0 for l in b), max(l.x1 for l in b)
            ov = _overlap(ln.x0, ln.x1, bx0, bx1)
            aligned = abs(ln.x0 - bx0) < 3 or abs(ln.x0 - last.x0) < 3
            bullet_cont = _is_bullet(last.text) and ln.x0 > last.x0
            if not (ov > 0.5 * min(ln.x1 - ln.x0, bx1 - bx0) or aligned or bullet_cont):
                continue
            if abs(ln.size - last.size) > 0.25 * size:
                continue
            if best is None or gap < best_gap:
                best, best_gap = b, gap
        if best is None:
            blocks.append([ln])
        else:
            best.append(ln)
    out = []
    for b in blocks:
        b.sort(key=lambda l: l.top)
        out.append(TextBlock(_split_paras(b)))
    out.sort(key=lambda b: (b.top, b.x0))
    return out


def _split_paras(lines: list[Line]) -> list[Para]:
    bx0, bx1 = min(l.x0 for l in lines), max(l.x1 for l in lines)
    width = bx1 - bx0
    paras: list[list[Line]] = []
    for ln in lines:
        if not paras:
            paras.append([ln])
            continue
        prev = paras[-1][-1]
        new = False
        if _is_bullet(ln.text):
            new = True
        elif ln.top - prev.bottom > max(prev.size, ln.size) * 0.55:
            new = True
        elif abs(ln.size - prev.size) > 0.6 or ln.bold != prev.bold and (len(prev.text) < 90 or len(ln.text) < 90):
            new = True
        elif bx1 - prev.x1 > _first_word_w(ln) + prev.size * 0.6 and not _centered(prev, bx0, bx1) \
                and width > 120 and len(lines) > 2:
            new = True  # vorherige Zeile endet früh, obwohl das nächste Wort noch Platz gehabt hätte
        elif ln.x0 - bx0 > ln.size * 1.2 and not _is_bullet(paras[-1][0].text) and abs(prev.x0 - bx0) < 3 \
                and prev.x1 > bx1 - prev.size * 3:
            new = True  # Einzug der ersten Zeile
        if new:
            paras.append([ln])
        else:
            paras[-1].append(ln)
    return [Para(p, _alignment(p, bx0, bx1)) for p in paras]


def _first_word_w(ln):
    word = ln.text.split(" ", 1)[0]
    if not word:
        return 0
    frac = len(word) / max(len(ln.text), 1)
    return (ln.x1 - ln.x0) * frac


def _centered(ln, bx0, bx1):
    return abs((ln.x0 + ln.x1) / 2 - (bx0 + bx1) / 2) < 2.5 and ln.x0 - bx0 > 6


def _alignment(lines, bx0, bx1):
    if len(lines) == 1:
        return "left"
    lefts = [l.x0 for l in lines]
    rights = [l.x1 for l in lines]
    centers = [(l.x0 + l.x1) / 2 for l in lines]
    rng = lambda v: max(v) - min(v)
    if rng(lefts) < 2.5:
        if len(lines) >= 3 and rng(rights[:-1]) < 2.5:
            return "justify"
        return "left"
    if rng(centers) < 3:
        return "center"
    if rng(rights) < 2.5:
        return "right"
    return "left"


# =================================================================== Tabellen
def _cell_paras(chars):
    lines = chars_to_lines(chars, split_gaps=False)
    if not lines:
        return []
    paras = []
    for b in lines_to_blocks(lines):
        paras.extend(b.paras)
    paras.sort(key=lambda p: p.top)
    return paras


def _cell_align(paras, x0, x1):
    if not paras:
        return "left"
    lgap = min(p.x0 for p in paras) - x0
    rgap = x1 - max(p.x1 for p in paras)
    if abs(lgap - rgap) < 4 and lgap > 8:
        return "center"
    if rgap < lgap - 6 and rgap < 12:
        return "right"
    return "left"


def find_ruled_tables(page, chars, edges, fills):
    try:
        found = page.find_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines",
                                  "snap_tolerance": 3, "join_tolerance": 3, "intersection_tolerance": 3})
    except Exception:
        return []
    tables = []
    for t in found:
        cells_bb = [c for c in t.cells if c]
        if len(cells_bb) < 2:
            continue
        xs = _cluster([c[0] for c in cells_bb] + [c[2] for c in cells_bb], 2.0)
        ys = _cluster([c[1] for c in cells_bb] + [c[3] for c in cells_bb], 2.0)
        if len(xs) < 3 and len(ys) < 3:
            continue
        cells = []
        texted = 0
        for bb in cells_bb:
            c0, c1 = _snap(bb[0], xs), _snap(bb[2], xs)
            r0, r1 = _snap(bb[1], ys), _snap(bb[3], ys)
            if c1 <= c0 or r1 <= r0:
                continue
            cc = [ch for ch in chars if _inside(ch["x0"], ch["top"], ch["x1"], ch["bottom"], bb, pad=0.5)]
            paras = _cell_paras(cc)
            if paras:
                texted += 1
            cell = Cell(r0, c0, r1 - r0, c1 - c0, xs[c0], ys[r0], xs[c1], ys[r1], paras)
            for p in paras:
                if p.align == "left" and len(p.lines) == 1:
                    p.align = _cell_align([p], cell.x0, cell.x1)
            cell.fill = _cell_fill(cell, fills)
            cell.borders = _cell_borders(cell, edges)
            cells.append(cell)
        # Nur echte Tabellen: genug Zellen mit Text (Balkendiagramme haben keinen Text in den "Zellen")
        if texted < 2 or texted < len(cells) * 0.3:
            continue
        tables.append(Table(xs, ys, _fill_grid(cells, xs, ys), True))
    return tables


def _fill_grid(cells, xs, ys):
    """Fehlende Gitterplätze mit leeren Zellen ergänzen."""
    used = set()
    for c in cells:
        for r in range(c.r, c.r + c.rs):
            for k in range(c.c, c.c + c.cs):
                used.add((r, k))
    for r in range(len(ys) - 1):
        for k in range(len(xs) - 1):
            if (r, k) not in used:
                cells.append(Cell(r, k, 1, 1, xs[k], ys[r], xs[k + 1], ys[r + 1], [], None,
                                  {"l": None, "r": None, "t": None, "b": None}))
    cells.sort(key=lambda c: (c.r, c.c))
    return cells


def _cell_fill(cell, fills):
    cx, cy = (cell.x0 + cell.x1) / 2, (cell.top + cell.bottom) / 2
    area = (cell.x1 - cell.x0) * (cell.bottom - cell.top)
    best = None
    for f in fills:
        if f["x0"] - 1 <= cx <= f["x1"] + 1 and f["top"] - 1 <= cy <= f["bottom"] + 1:
            fa = (f["x1"] - f["x0"]) * (f["bottom"] - f["top"])
            if fa >= area * 0.5 and (best is None or fa < best[0]):
                best = (fa, f["color"])
    return best[1] if best else None


def _cell_borders(cell, edges):
    out = {}
    for side in "ltrb":
        if side in "lr":
            x = cell.x0 if side == "l" else cell.x1
            need = cell.bottom - cell.top
            cov, col, w = 0.0, None, 0.0
            for e in edges:
                if e["orientation"] == "v" and abs(e["x0"] - x) < 2.5:
                    o = _overlap(e["top"], e["bottom"], cell.top, cell.bottom)
                    if o > 0:
                        cov += o
                        col, w = e["color"], e["lw"]
        else:
            y = cell.top if side == "t" else cell.bottom
            need = cell.x1 - cell.x0
            cov, col, w = 0.0, None, 0.0
            for e in edges:
                if e["orientation"] == "h" and abs(e["top"] - y) < 2.5:
                    o = _overlap(e["x0"], e["x1"], cell.x0, cell.x1)
                    if o > 0:
                        cov += o
                        col, w = e["color"], e["lw"]
        out[side] = (col or (0, 0, 0), max(w, 0.5)) if cov >= need * 0.6 and col is not None else None
    return out


def _char_rows(chars):
    """Zeichen zu Zeilen (Reihen) gruppieren – ohne Spaltentrennung."""
    rows: list[dict] = []
    for c in sorted(chars, key=lambda c: (round(c["top"], 1), c["x0"])):
        h = max(c["bottom"] - c["top"], 0.5)
        best = None
        for r in reversed(rows[-6:]):
            if _overlap(c["top"], c["bottom"], r["top"], r["bottom"]) >= 0.5 * min(h, r["bottom"] - r["top"]):
                best = r
                break
        if best is None:
            rows.append({"top": c["top"], "bottom": c["bottom"], "chars": [c]})
        else:
            best["chars"].append(c)
            best["top"], best["bottom"] = min(best["top"], c["top"]), max(best["bottom"], c["bottom"])
    for r in rows:
        r["chars"].sort(key=lambda c: c["x0"])
        r["words"] = _words(r["chars"])
        r["size"] = statistics.median(c["size"] for c in r["chars"])
    rows.sort(key=lambda r: r["top"])
    return [r for r in rows if r["words"]]


def _words(chars):
    words, cur = [], []
    for c in chars:
        t = _char_text(c)
        if not t.strip():
            if cur:
                words.append(cur)
                cur = []
            continue
        if cur and c["x0"] - cur[-1]["x1"] > max(cur[-1]["size"], c["size"]) * 0.2:
            words.append(cur)
            cur = []
        cur.append(c)
    if cur:
        words.append(cur)
    return [{"x0": w[0]["x0"], "x1": max(x["x1"] for x in w), "chars": w} for w in words]


def _channels(rows):
    """Spalten = x-Bereiche zwischen senkrechten Leerraum-Kanälen, die in allen Reihen frei sind."""
    words = [w for r in rows for w in r["words"]]
    lo, hi = min(w["x0"] for w in words), max(w["x1"] for w in words)
    size = statistics.median(r["size"] for r in rows)
    ivs = sorted((w["x0"], w["x1"]) for w in words)
    merged = []
    for a, b in ivs:
        if merged and a <= merged[-1][1] + max(size * 0.6, 4):
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return merged


def _row_has_gap(r):
    ws = r["words"]
    return any(b["x0"] - a["x1"] >= max(r["size"] * 0.9, 6) for a, b in zip(ws, ws[1:]))


def find_unruled_tables(chars, page_w, exclude=()):
    """Tabellen ohne Linien anhand senkrechter Leerraum-Kanäle. Gibt (Tabellen, benutzte Zeichen-IDs) zurück."""
    chars = [c for c in chars if _upright(c) and _char_text(c) != ""
             and not any(_inside(c["x0"], c["top"], c["x1"], c["bottom"], b, 0) for b in exclude)]
    rows = _char_rows(chars)
    tables, used = [], set()
    i = 0
    while i < len(rows):
        if not _row_has_gap(rows[i]):
            i += 1
            continue
        cur, j = [rows[i]], i
        cols = _channels(cur)
        while j + 1 < len(rows):
            nxt = rows[j + 1]
            h = cur[-1]["bottom"] - cur[-1]["top"]
            gap = nxt["top"] - cur[-1]["bottom"]
            if gap > h * 1.6 or gap < -1 or abs(nxt["size"] - cur[-1]["size"]) > cur[-1]["size"] * 0.35:
                break
            c2 = _channels(cur + [nxt])
            if len(c2) < 2 or len(c2) < len(cols) - 1:
                break
            cur, cols, j = cur + [nxt], c2, j + 1
        # auch nach oben erweitern (Kopfzeile hat oft kleinere Lücken)
        k = i
        while k - 1 >= 0 and len(cur) >= 2:
            prv = rows[k - 1]
            if any(id(c) in used for c in prv["chars"]):
                break
            h = prv["bottom"] - prv["top"]
            gap = cur[0]["top"] - prv["bottom"]
            if gap > h * 1.6 or gap < -1 or abs(prv["size"] - cur[0]["size"]) > cur[0]["size"] * 0.35:
                break
            if len(prv["words"]) < 2:
                break
            # Kopfzeile: Wörter den vorhandenen Spalten zuordnen (Spalten nicht neu berechnen)
            hit = set()
            for w in prv["words"]:
                cx = (w["x0"] + w["x1"]) / 2
                inside = [ci for ci, (a0, a1) in enumerate(cols) if a0 - 8 <= cx <= a1 + 8]
                if not inside:
                    hit = None
                    break
                hit.add(inside[0])
            ok = bool(hit) and len(hit) >= 2 and \
                (prv["words"][-1]["x1"] - prv["words"][0]["x0"]) <= (cols[-1][1] - cols[0][0]) * 1.3
            if not ok and k == i:
                ok = _header_columns([prv] + cur, cols) is not None
            if not ok:
                break
            cur, k = [prv] + cur, k - 1
        tbl = _make_unruled(cur, cols, page_w) if len(cur) >= 3 and len(cols) >= 2 else None
        if tbl:
            tables.append(tbl)
            for r in cur:
                for c in r["chars"]:
                    used.add(id(c))
            i = j + 1
        else:
            i += 1
    return tables, used


def _header_columns(rows, cols):
    """Spaltengrenzen aus der Kopfzeile (z. B. Excel-Export: Kopf linksbündig, Zahlen rechtsbündig)."""
    r0 = rows[0]
    hc = []
    for w in r0["words"]:
        if hc and w["x0"] - hc[-1][1] < max(r0["size"] * 0.9, 6):
            hc[-1][1] = w["x1"]
        else:
            hc.append([w["x0"], w["x1"]])
    if not (len(cols) <= len(hc) <= len(cols) + 2):
        return None
    allw = [w for r in rows for w in r["words"]]
    right = max(w["x1"] for w in allw) + 3
    xs = [min(hc[0][0], min(w["x0"] for w in allw)) - 3] + [h[0] - 3 for h in hc[1:]] + [right]
    if any(b <= a for a, b in zip(xs, xs[1:])):
        return None
    # jede Datenspalte muss in genau eine Kopfspalte fallen
    for a0, a1 in cols:
        cx = (a0 + a1) / 2
        if not any(xs[k] <= cx <= xs[k + 1] for k in range(len(hc))):
            return None
    return xs


def _make_unruled(rows, cols, page_w):
    body_cols = _channels(rows[1:]) if len(rows) > 3 else cols
    if len(body_cols) > len(cols):
        cols = body_cols
    xs_h = _header_columns(rows, cols)
    if xs_h:
        cols = [[xs_h[k] + 3, xs_h[k + 1] - 3] for k in range(len(xs_h) - 1)]

    def col_of(w):
        cx = (w["x0"] + w["x1"]) / 2
        for k, (a, b) in enumerate(cols):
            if a - 1 <= cx <= b + 1:
                return k
        return min(range(len(cols)), key=lambda k: abs((cols[k][0] + cols[k][1]) / 2 - cx))

    grid = [[[] for _ in cols] for _ in rows]
    for ri, r in enumerate(rows):
        for w in r["words"]:
            if xs_h:
                # Wörter, die über eine Spaltengrenze laufen (z. B. "###Produkt"), zeichenweise aufteilen
                parts = {}
                for c in w["chars"]:
                    cx = (c["x0"] + c["x1"]) / 2
                    k = next((k for k in range(len(xs_h) - 1) if cx < xs_h[k + 1]), len(xs_h) - 2)
                    parts.setdefault(k, []).append(c)
                for k, cs in parts.items():
                    grid[ri][k].append({"x0": cs[0]["x0"], "x1": max(c["x1"] for c in cs), "chars": cs})
            else:
                grid[ri][col_of(w)].append(w)
    filled = sum(1 for row in grid for cell in row if cell)
    multi_rows = sum(1 for row in grid if sum(1 for cell in row if cell) >= 2)
    if filled < len(rows) * len(cols) * 0.5 or multi_rows < max(2, len(rows) * 0.5):
        return None
    widths = [b - a for a, b in cols]
    words_per_cell = sum(len(cell) for row in grid for cell in row) / max(filled, 1)
    if len(cols) == 2 and min(widths) > page_w * 0.25 and words_per_cell > 3:
        return None  # zweispaltiger Fließtext
    first = ["".join(_char_text(c) for w in row[0] for c in w["chars"]).strip() for row in grid if row[0]]
    if first and all(len(t) <= 3 and (t[:1] in BULLET_CHARS or re.fullmatch(r"\d+[.)]?|[a-z][.)]", t)) for t in first):
        return None  # Aufzählung
    # Ausrichtung je Spalte
    col_align = []
    for k in range(len(cols)):
        ext = [(min(w["x0"] for w in row[k]), max(w["x1"] for w in row[k])) for row in grid[1:] if row[k]]
        x0s, x1s = [e[0] for e in ext], [e[1] for e in ext]
        if len(ext) > 1 and max(x1s) - min(x1s) < 2 and max(x0s) - min(x0s) > 2:
            col_align.append("right")
        elif len(ext) > 1 and max(x0s) - min(x0s) > 2 and max((a + b) / 2 for a, b in ext) - min((a + b) / 2 for a, b in ext) < 2:
            col_align.append("center")
        else:
            col_align.append("left")
    if xs_h:
        xs = list(xs_h)
    else:
        xs = [cols[0][0] - 3]
        for k in range(len(cols) - 1):
            if col_align[k] == "right":
                xs.append(cols[k][1] + 3)
            elif col_align[k + 1] == "left":
                xs.append(cols[k + 1][0] - 3)
            else:
                xs.append((cols[k][1] + cols[k + 1][0]) / 2)
        xs.append(cols[-1][1] + 3)
    # Spalten so weit verbreitern, dass der Text ohne Umbruch hineinpasst
    widest = [0.0] * len(cols)
    for row in grid:
        for k, ws in enumerate(row):
            if ws:
                widest[k] = max(widest[k], max(w["x1"] for w in ws) - min(w["x0"] for w in ws))
    for k in range(len(cols)):
        need = widest[k] + 9
        if xs[k + 1] - xs[k] < need:
            shift = need - (xs[k + 1] - xs[k])
            for j2 in range(k + 1, len(xs)):
                xs[j2] += shift
    ys = [rows[0]["top"] - 2]
    for a, b in zip(rows, rows[1:]):
        ys.append((a["bottom"] + b["top"]) / 2)
    ys.append(rows[-1]["bottom"] + 2)
    cells = []
    for ri in range(len(rows)):
        for ci in range(len(cols)):
            ws = grid[ri][ci]
            paras = []
            if ws:
                ln = _make_line([c for w in ws for c in w["chars"]])
                if ln:
                    al = col_align[ci]
                    if ri == 0 and abs(ln.x0 - xs[ci]) < 7:
                        al = "left"  # Kopfzeile steht oft linksbündig über rechtsbündigen Zahlen
                    paras = [Para([ln], al)]
            cells.append(Cell(ri, ci, 1, 1, xs[ci], ys[ri], xs[ci + 1], ys[ri + 1], paras, None,
                              {"l": None, "t": None, "r": None, "b": None}))
    return Table(xs, ys, cells, False)


# =================================================================== Grafik
def _bezier(p0, p1, p2, p3, n=10):
    pts = []
    for k in range(1, n + 1):
        t = k / n
        mt = 1 - t
        pts.append((mt ** 3 * p0[0] + 3 * mt * mt * t * p1[0] + 3 * mt * t * t * p2[0] + t ** 3 * p3[0],
                    mt ** 3 * p0[1] + 3 * mt * mt * t * p1[1] + 3 * mt * t * t * p2[1] + t ** 3 * p3[1]))
    return pts


def _path_points(obj):
    path = obj.get("path")
    if path:
        pts, cur, start = [], None, None
        closed = False
        for seg in path:
            op = seg[0]
            if op == "m":
                cur = start = seg[1]
                pts.append(cur)
            elif op == "l":
                cur = seg[1]
                pts.append(cur)
            elif op == "c" and cur is not None:
                pts.extend(_bezier(cur, seg[1], seg[2], seg[3]))
                cur = seg[3]
            elif op == "h":
                closed = True
                if start:
                    pts.append(start)
        return pts, closed or (len(pts) > 2 and pts[0] == pts[-1])
    pts = obj.get("pts") or []
    return list(pts), len(pts) > 2 and pts[0] == pts[-1]


def collect_vectors(page, ox, oy, table_boxes):
    """Liefert (Formen, Füllflächen, Kanten) einer Seite. Koordinaten relativ zur Seite."""
    shapes, fills = [], []
    W, H = float(page.width), float(page.height)

    def norm(o):
        return o["x0"] - ox, o["top"] - oy, o["x1"] - ox, o["bottom"] - oy

    for kind, objs in (("rect", page.rects), ("line", page.lines), ("path", page.curves)):
        for o in objs:
            x0, top, x1, bottom = norm(o)
            fill = to_rgb(o.get("non_stroking_color")) if o.get("fill") else None
            stroke = to_rgb(o.get("stroking_color")) if o.get("stroke") or kind == "line" else None
            lw = float(o.get("linewidth") or 0)
            if kind == "line" and stroke is None:
                stroke = (0, 0, 0)
            if stroke is not None and lw <= 0:
                lw = 0.5
            if fill is not None and fill == (255, 255, 255) and stroke is None:
                continue  # weiße Abdeckflächen
            if fill is None and stroke is None:
                continue
            if fill is not None:
                fills.append({"x0": x0, "top": top, "x1": x1, "bottom": bottom, "color": fill})
            pts, closed = [], False
            if kind in ("line", "path"):
                raw, closed = _path_points(o)
                pts = [(p[0] - ox, p[1] - oy) for p in raw]
            k = kind
            if kind == "path" and len(pts) == 2:
                k = "line"
            shapes.append(Shape(k, x0, top, x1, bottom, pts, fill, stroke, lw, closed))
    # Sichtbare Linien für Tabellenrahmen: Linien, Rahmen von Rechtecken und sehr dünne gefüllte Rechtecke
    edges = []

    def add_edge(o, x0, top, x1, bottom, col, lw):
        if col is None or col == (255, 255, 255):
            return
        edges.append({"orientation": o, "x0": x0, "x1": x1, "top": top, "bottom": bottom, "color": col, "lw": lw})

    for sh in shapes:
        if sh.kind == "line" and len(sh.pts) >= 2:
            (ax, ay), (bx, by) = sh.pts[0], sh.pts[-1]
            if abs(ay - by) < 1:
                add_edge("h", min(ax, bx), ay, max(ax, bx), ay, sh.stroke, sh.lw)
            elif abs(ax - bx) < 1:
                add_edge("v", ax, min(ay, by), ax, max(ay, by), sh.stroke, sh.lw)
        elif sh.kind == "rect":
            w, h = sh.x1 - sh.x0, sh.bottom - sh.top
            if sh.stroke is not None:
                for o, args in (("h", (sh.x0, sh.top, sh.x1, sh.top)), ("h", (sh.x0, sh.bottom, sh.x1, sh.bottom)),
                                ("v", (sh.x0, sh.top, sh.x0, sh.bottom)), ("v", (sh.x1, sh.top, sh.x1, sh.bottom))):
                    add_edge(o, *args, sh.stroke, sh.lw)
            elif sh.fill is not None and h <= 2.5 and w > h:
                y = (sh.top + sh.bottom) / 2
                add_edge("h", sh.x0, y, sh.x1, y, sh.fill, max(h, 0.25))
            elif sh.fill is not None and w <= 2.5 and h > w:
                x = (sh.x0 + sh.x1) / 2
                add_edge("v", x, sh.top, x, sh.bottom, sh.fill, max(w, 0.25))
    return shapes, fills, edges


def _union_clusters(items, gap=2.0):
    n = len(items)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    order = sorted(range(n), key=lambda i: items[i].x0)
    for a_i, a in enumerate(order):
        A = items[a]
        for b in order[a_i + 1:]:
            B = items[b]
            if B.x0 > A.x1 + gap:
                break
            if B.top <= A.bottom + gap and A.top <= B.bottom + gap:
                parent[find(a)] = find(b)
    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(items[i])
    return list(groups.values())


# =================================================================== Seite analysieren
def _normalize_rotation(pdf_path: str) -> str | None:
    """Gedrehte Seiten in eine ungedrehte Kopie umrechnen (Inhalt mitgedreht)."""
    try:
        import pypdf
        import tempfile

        r = pypdf.PdfReader(pdf_path)
        if r.is_encrypted:
            r.decrypt("")
        if not any((p.rotation or 0) % 360 for p in r.pages):
            return None
        w = pypdf.PdfWriter()
        for p in r.pages:
            if (p.rotation or 0) % 360:
                p.transfer_rotation_to_content()
            w.add_page(p)
        fd, tmp = tempfile.mkstemp(suffix=".pdf", prefix="ufc_rot_")
        with open(fd, "wb") as f:
            w.write(f)
        return tmp
    except Exception as e:
        print("Drehung konnte nicht ausgeglichen werden:", e)
        return None


def analyze(pdf_path: str, progress=None, dpi=200) -> list[PageModel]:
    tmp = _normalize_rotation(pdf_path)
    try:
        return _analyze(tmp or pdf_path, progress, dpi)
    finally:
        if tmp:
            import os
            try:
                os.remove(tmp)
            except OSError:
                pass


def _analyze(pdf_path: str, progress=None, dpi=200) -> list[PageModel]:
    import pdfplumber
    import pypdfium2 as pdfium

    models = []
    pdoc = pdfium.PdfDocument(pdf_path)
    try:
        with pdfplumber.open(pdf_path) as doc:
            n = len(doc.pages)
            for pi, page in enumerate(doc.pages):
                models.append(_analyze_page(page, pdoc, pi, dpi))
                if progress:
                    progress((pi + 1) / n, f"Analysiere Seite {pi + 1}/{n}")
    finally:
        pdoc.close()
    return models


def _render(pdoc, pi, scale, remove_text):
    page = pdoc[pi]
    try:
        if page.get_rotation():
            page.set_rotation(0)
        if remove_text:
            import pypdfium2.raw as raw

            objs = list(page.get_objects(filter=[raw.FPDF_PAGEOBJ_TEXT], max_depth=1))
            removed = 0
            for o in objs:
                try:
                    m = o.get_matrix()
                    if abs(m.b) < 1e-3 and abs(m.c) < 1e-3:  # gedrehten Text als Bild behalten
                        page.remove_obj(o)
                        o.close()
                        removed += 1
                except Exception:
                    pass
            del objs
            if removed:
                page.gen_content()
        return page.render(scale=scale).to_pil().convert("RGB")
    finally:
        page.close()


def _crop(img, scale, x0, top, x1, bottom, pad=0.0):
    W, H = img.size
    box = (max(0, int((x0 - pad) * scale)), max(0, int((top - pad) * scale)),
           min(W, int((x1 + pad) * scale + 0.999)), min(H, int((bottom + pad) * scale + 0.999)))
    if box[2] - box[0] < 2 or box[3] - box[1] < 2:
        return None
    bio = io.BytesIO()
    img.crop(box).save(bio, "PNG", optimize=True)
    return bio.getvalue()


def _analyze_page(page, pdoc, pi, dpi) -> PageModel:
    ox, oy = float(page.bbox[0]), float(page.bbox[1])
    W, H = float(page.width), float(page.height)
    m = PageModel(W, H)
    chars = []
    for c in page.chars:
        c = dict(c)
        c["x0"] -= ox
        c["x1"] -= ox
        c["top"] -= oy
        c["bottom"] -= oy
        if c["x1"] < 0 or c["x0"] > W or c["bottom"] < 0 or c["top"] > H:
            continue
        chars.append(c)

    # Seite mit überwiegend gedrehtem Text: originalgetreu als Bild übernehmen
    n_up = sum(1 for c in chars if _upright(c))
    if len(chars) - n_up > max(20, n_up):
        scale = dpi / 72
        full = _render(pdoc, pi, scale, remove_text=False)
        png = _crop(full, scale, 0, 0, W, H)
        if png:
            pic = Picture(0, 0, W, H, png, "graphic")
            m.pictures.append(pic)
            m.graphics_img.append(pic)
        return m

    shapes, fills, edges = collect_vectors(page, ox, oy, [])
    # Hintergrund
    for s in list(shapes):
        if s.kind == "rect" and s.fill and (s.x1 - s.x0) * (s.bottom - s.top) > 0.9 * W * H:
            if m.bg is None:
                m.bg = s.fill
            shapes.remove(s)
            fills = [f for f in fills if not (f["x0"] == s.x0 and f["top"] == s.top and f["x1"] == s.x1)]

    # ---- Tabellen mit Linien
    m.tables = _ruled(page, chars, edges, fills, ox, oy)
    tboxes = [(t.x0, t.top, t.x1, t.bottom) for t in m.tables]
    rest_chars = [c for c in chars if not any(_inside(c["x0"], c["top"], c["x1"], c["bottom"], b, 0.5) for b in tboxes)]
    shapes = [s for s in shapes if not any(s.x0 >= b[0] - 2.5 and s.x1 <= b[2] + 2.5 and s.top >= b[1] - 2.5
                                           and s.bottom <= b[3] + 2.5 for b in tboxes)]

    # ---- Text
    lines = chars_to_lines(rest_chars)

    # ---- Grafiken gruppieren
    need_full = False
    graphic_lines_word: set = set()
    others = []
    for grp in _union_clusters(shapes, 2.0):
        gx0, gtop = min(s.x0 for s in grp), min(s.top for s in grp)
        gx1, gbot = max(s.x1 for s in grp), max(s.bottom for s in grp)
        gw, gh = gx1 - gx0, gbot - gtop
        inner = [l for l in lines if _inside(l.x0, l.top, l.x1, l.bottom, (gx0, gtop, gx1, gbot), 0.5)]
        n_curve_pts = sum(len(s.pts) for s in grp if s.kind == "path")
        rects = [s for s in grp if s.kind == "rect"]
        # Trennlinie
        if len(grp) <= 2 and gh <= 3.5 and gw > 25 and not inner:
            st = grp[0]
            m.rules.append(Rule(gx0, gx1, (gtop + gbot) / 2, st.stroke or st.fill, max(st.lw, gh if st.fill else 0.75)))
            m.shapes.extend(grp)
            continue
        # Box um Text (Infobox, Kopfbalken)
        if len(grp) <= 3 and len(rects) == len(grp) and inner and \
                max((r.x1 - r.x0) * (r.bottom - r.top) for r in rects) >= 0.8 * gw * gh:
            big = max(rects, key=lambda r: (r.x1 - r.x0) * (r.bottom - r.top))
            m.boxes.append(Box(gx0, gtop, gx1, gbot, next((r.fill for r in rects if r.fill), None),
                               next((r.stroke for r in rects if r.stroke), None), big.lw))
            m.shapes.extend(grp)
            continue
        if gw < 4 and gh < 4:
            m.shapes.extend(grp)
            continue
        others.append(grp)

    # sonstige Grafiken, die nahe beieinander liegen (Diagramm + Legende), zusammenfassen
    merged_groups = []
    for grp in others:
        bb = [min(s.x0 for s in grp), min(s.top for s in grp), max(s.x1 for s in grp), max(s.bottom for s in grp)]
        merged_groups.append([bb, list(grp)])
    changed = True
    while changed:
        changed = False
        for a_i in range(len(merged_groups)):
            for b_i in range(a_i + 1, len(merged_groups)):
                A, B = merged_groups[a_i][0], merged_groups[b_i][0]
                if B[0] <= A[2] + 20 and A[0] <= B[2] + 20 and B[1] <= A[3] + 20 and A[1] <= B[3] + 20:
                    merged_groups[a_i][0] = [min(A[0], B[0]), min(A[1], B[1]), max(A[2], B[2]), max(A[3], B[3])]
                    merged_groups[a_i][1] += merged_groups[b_i][1]
                    del merged_groups[b_i]
                    changed = True
                    break
            if changed:
                break
    for (gx0, gtop, gx1, gbot), grp in merged_groups:
        inner = [l for l in lines if _inside(l.x0, l.top, l.x1, l.bottom, (gx0, gtop, gx1, gbot), 0.5)]
        n_curve_pts = sum(len(s.pts) for s in grp if s.kind == "path")
        # sonstige Grafik (Diagramm, Logo …)
        complexity = len(grp) + n_curve_pts / 6
        bbox = (gx0, gtop, gx1, gbot)
        need_full = True
        wb = list(bbox)
        for l in inner:
            graphic_lines_word.add(id(l))
        # auch Textzeilen knapp neben dem Diagramm (Achsenbeschriftung, Titel) gehören dazu
        for l in lines:
            if id(l) not in graphic_lines_word and _near(l, bbox, 14) and len(l.text) < 25:
                graphic_lines_word.add(id(l))
                wb = [min(wb[0], l.x0), min(wb[1], l.top), max(wb[2], l.x1), max(wb[3], l.bottom)]
        m.graphics_img.append(tuple(wb))  # später gerendert (mit Text) – für Word
        if complexity <= 220:
            m.shapes.extend(grp)  # für PowerPoint als echte Formen
        else:
            m.graphics_img_notext.append(bbox)

    # ---- Tabellen ohne Linien
    gboxes = [g for g in m.graphics_img]
    ut, used = find_unruled_tables(rest_chars, W, exclude=[(g[0], g[1], g[2], g[3]) for g in gboxes])
    if ut:
        m.tables.extend(ut)
        rest_chars = [c for c in rest_chars if id(c) not in used]
        old_hidden = [l for l in lines if id(l) in graphic_lines_word]
        lines = chars_to_lines(rest_chars)
        graphic_lines_word = {id(l) for l in lines if any(
            abs(l.top - o.top) < 0.5 and abs(l.x0 - o.x0) < 0.5 for o in old_hidden)}

    # ---- Textblöcke
    m.blocks = lines_to_blocks(lines)
    m.word_hidden_lines = list(graphic_lines_word)

    # ---- Bilder & Grafiken rendern
    imgs = []
    for im in page.images:
        x0, top = max(0.0, im["x0"] - ox), max(0.0, im["top"] - oy)
        x1, bottom = min(W, im["x1"] - ox), min(H, im["bottom"] - oy)
        if x1 - x0 > 3 and bottom - top > 3:
            imgs.append((x0, top, x1, bottom, im.get("srcsize")))
    if need_full:
        scale = dpi / 72
        full = _render(pdoc, pi, scale, remove_text=False)
        pics = []
        for b in m.graphics_img:
            b = (max(0, b[0] - 3), max(0, b[1] - 3), min(W, b[2] + 3), min(H, b[3] + 3))
            png = _crop(full, scale, *b)
            if png:
                pics.append(Picture(*b, png, "graphic"))
        m.graphics_img = pics
    if imgs or m.graphics_img_notext:
        dpi_needed = dpi
        for x0, top, x1, bottom, src in imgs:
            if src and x1 > x0:
                dpi_needed = max(dpi_needed, min(300, src[0] / ((x1 - x0) / 72)))
        scale = dpi_needed / 72
        notext = _render(pdoc, pi, scale, remove_text=True)
        for x0, top, x1, bottom, _ in imgs:
            png = _crop(notext, scale, x0, top, x1, bottom)
            if png:
                m.pictures.append(Picture(x0, top, x1, bottom, png, "image"))
        m.graphics_img_notext = [Picture(*b, _crop(notext, scale, *b, pad=1.5), "graphic")
                                 for b in m.graphics_img_notext]
        m.graphics_img_notext = [p for p in m.graphics_img_notext if p.png]
    return m


def _near(l, bbox, d):
    return (l.x1 > bbox[0] - d and l.x0 < bbox[2] + d and l.bottom > bbox[1] - d and l.top < bbox[3] + d)


def _ruled(page, chars, edges, fills, ox, oy):
    """find_tables arbeitet in Original-Koordinaten -> Ergebnisse auf Seitenkoordinaten verschieben."""
    tables = find_ruled_tables(page, [dict(c, x0=c["x0"] + ox, x1=c["x1"] + ox, top=c["top"] + oy,
                                           bottom=c["bottom"] + oy) for c in chars],
                               [dict(e, x0=e["x0"] + ox, x1=e["x1"] + ox, top=e["top"] + oy,
                                     bottom=e["bottom"] + oy) for e in edges],
                               [dict(f, x0=f["x0"] + ox, x1=f["x1"] + ox, top=f["top"] + oy,
                                     bottom=f["bottom"] + oy) for f in fills])
    if ox == 0 and oy == 0:
        return tables
    for t in tables:
        t.xs = [x - ox for x in t.xs]
        t.ys = [y - oy for y in t.ys]
        for c in t.cells:
            c.x0 -= ox
            c.x1 -= ox
            c.top -= oy
            c.bottom -= oy
            for p in c.paras:
                for ln in p.lines:
                    ln.x0 -= ox
                    ln.x1 -= ox
                    ln.top -= oy
                    ln.bottom -= oy
    return tables
