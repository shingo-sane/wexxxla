"""Textformate: Markdown, HTML, EPUB, ODT, RTF über pandoc – plus eingebaute Fallbacks."""
from __future__ import annotations

import html as htmlmod
import os
import re
import shutil

from . import tools
from .engine import ConversionError, Ctx, route
from .data import read_text

PANDOC_IN = {"md": "markdown", "html": "html", "docx": "docx", "odt": "odt", "epub": "epub"}
PANDOC_OUT = {"md": "gfm", "html": "html", "docx": "docx", "odt": "odt", "epub": "epub",
              "rtf": "rtf", "txt": "plain"}


@route(list(PANDOC_IN), list(PANDOC_OUT), cost=1, backend="pandoc")
def pandoc_convert(src, dst, ctx: Ctx):
    pandoc = tools.find_pandoc()
    if not pandoc:
        raise ConversionError("pandoc wurde nicht gefunden.")
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    cmd = [pandoc, src, "-f", PANDOC_IN[ctx.src_ext], "-t", PANDOC_OUT[ctx.dst_ext], "-o", dst, "--wrap=none"]
    title = ctx.opt("_title", "Dokument")
    if ctx.dst_ext in ("html", "epub", "rtf", "odt", "docx"):
        cmd += ["--standalone", "--metadata", f"pagetitle={title}"]
    if ctx.dst_ext == "epub":
        cmd += ["--metadata", f"title={title}"]
    if ctx.dst_ext == "html":
        cmd += ["--embed-resources"]
    res = tools.run(cmd, timeout=300, cwd=os.path.dirname(os.path.abspath(src)))
    if res.returncode != 0 and "--embed-resources" in cmd:
        cmd.remove("--embed-resources")
        res = tools.run(cmd, timeout=300, cwd=os.path.dirname(os.path.abspath(src)))
    if res.returncode != 0:
        raise ConversionError("pandoc: " + (res.stderr or b"").decode(errors="ignore").strip()[-300:])


# ------------------------------------------------------- eingebaute Fallbacks
def _html_page(body: str, title: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8"><title>{htmlmod.escape(title)}</title>
<style>body{{font-family:Segoe UI,Arial,sans-serif;max-width:46rem;margin:2rem auto;padding:0 1rem;
line-height:1.6;color:#222}}pre,code{{background:#f4f4f4;border-radius:4px}}pre{{padding:1rem;overflow:auto}}
table{{border-collapse:collapse}}td,th{{border:1px solid #ccc;padding:4px 8px}}</style>
</head><body>
{body}
</body></html>"""


@route("md", "html", cost=1.5)
def md_to_html(src, dst, ctx: Ctx):
    import markdown

    body = markdown.markdown(read_text(src), extensions=["tables", "fenced_code", "toc", "sane_lists"])
    with open(dst, "w", encoding="utf-8") as f:
        f.write(_html_page(body, ctx.opt("_title", "Dokument")))


@route("html", "md", cost=1.5)
def html_to_md(src, dst, ctx: Ctx):
    from markdownify import markdownify

    md = markdownify(read_text(src), heading_style="ATX")
    md = re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"
    with open(dst, "w", encoding="utf-8") as f:
        f.write(md)


@route("html", "txt", cost=1.5)
def html_to_txt(src, dst, ctx: Ctx):
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(read_text(src), "html.parser")
    for t in soup(["script", "style", "head"]):
        t.decompose()
    text = soup.get_text("\n")
    text = re.sub(r"\n\s*\n+", "\n\n", text).strip() + "\n"
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text)


@route("md", "txt", cost=1.5)
def md_to_txt(src, dst, ctx: Ctx):
    import markdown
    from bs4 import BeautifulSoup

    html = markdown.markdown(read_text(src), extensions=["tables", "fenced_code"])
    text = BeautifulSoup(html, "html.parser").get_text("\n")
    with open(dst, "w", encoding="utf-8") as f:
        f.write(re.sub(r"\n{3,}", "\n\n", text).strip() + "\n")


@route("txt", "md", cost=1)
def txt_to_md(src, dst, ctx: Ctx):
    with open(dst, "w", encoding="utf-8") as f:
        f.write(read_text(src))


@route("txt", "html", cost=1.5)
def txt_to_html(src, dst, ctx: Ctx):
    paras = [p.strip() for p in re.split(r"\n\s*\n", read_text(src)) if p.strip()]
    body = "\n".join("<p>" + htmlmod.escape(p).replace("\n", "<br>") + "</p>" for p in paras)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(_html_page(body, ctx.opt("_title", "Dokument")))
