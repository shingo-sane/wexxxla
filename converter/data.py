"""Tabellendaten: CSV, Excel, ODS, JSON, HTML-Tabellen (über pandas, ohne Office)."""
from __future__ import annotations

import csv
import io
import json
import os

from .engine import ConversionError, Ctx, route


def read_text(path) -> str:
    raw = open(path, "rb").read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def read_csv(path):
    import pandas as pd

    text = read_text(path)
    sample = text[:20000]
    sep = ","
    try:
        sep = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        counts = {d: sample.count(d) for d in (";", ",", "\t", "|")}
        sep = max(counts, key=counts.get)
    decimal = "," if sep == ";" else "."
    try:
        return pd.read_csv(io.StringIO(text), sep=sep, decimal=decimal, engine="python")
    except Exception as e:
        raise ConversionError(f"CSV kann nicht gelesen werden: {e}")


def read_sheets(path, ext) -> dict:
    """-> {Blattname: DataFrame}"""
    import pandas as pd

    if ext == "csv":
        return {"Tabelle1": read_csv(path)}
    if ext == "json":
        return read_json(path)
    if ext == "html":
        try:
            tables = pd.read_html(io.StringIO(read_text(path)))
        except ValueError:
            raise ConversionError("Die HTML-Datei enthält keine Tabelle.")
        return {f"Tabelle{i + 1}": t for i, t in enumerate(tables)}
    engine = {"ods": "odf", "xls": "xlrd"}.get(ext, "openpyxl")
    sheets = pd.read_excel(path, sheet_name=None, engine=engine)
    return {k: v for k, v in sheets.items() if not v.empty} or sheets


def read_json(path) -> dict:
    import pandas as pd

    try:
        data = json.loads(read_text(path))
    except json.JSONDecodeError as e:
        raise ConversionError(f"Ungültiges JSON: {e}")
    if isinstance(data, dict):
        # {"Blatt": [..], ...} oder ein einzelnes Objekt
        if data and all(isinstance(v, list) for v in data.values()):
            if all(all(isinstance(x, dict) for x in v) for v in data.values()):
                return {str(k)[:31]: pd.json_normalize(v) for k, v in data.items()}
            return {"Tabelle1": pd.DataFrame(dict((k, pd.Series(v)) for k, v in data.items()))}
        return {"Tabelle1": pd.json_normalize(data)}
    if isinstance(data, list):
        if all(isinstance(x, dict) for x in data):
            return {"Tabelle1": pd.json_normalize(data)}
        if all(isinstance(x, list) for x in data) and data:
            return {"Tabelle1": pd.DataFrame(data[1:], columns=[str(c) for c in data[0]])}
        return {"Tabelle1": pd.DataFrame({"Wert": data})}
    return {"Tabelle1": pd.DataFrame({"Wert": [data]})}


def write_xlsx(sheets: dict, dst):
    import pandas as pd
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(dst, engine="openpyxl") as xw:
        for name, df in sheets.items():
            safe = "".join(ch for ch in str(name) if ch not in '[]:*?/\\')[:31] or "Tabelle"
            df.to_excel(xw, sheet_name=safe, index=False)
            ws = xw.sheets[safe]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2F5597")
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, 1):
                vals = [str(col)] + [str(v) for v in df.iloc[:200, i - 1].tolist()]
                ws.column_dimensions[get_column_letter(i)].width = min(max(len(v) for v in vals) + 2, 60)


# ------------------------------------------------------------------ Routen
@route(["csv", "json", "ods"], "xlsx", cost=1)
def to_xlsx(src, dst, ctx: Ctx):
    write_xlsx(read_sheets(src, ctx.src_ext), dst)


@route("html", "xlsx", cost=2.5)
def html_to_xlsx(src, dst, ctx: Ctx):
    write_xlsx(read_sheets(src, "html"), dst)


try:
    import xlrd  # noqa: F401

    @route("xls", "xlsx", cost=1)
    def xls_to_xlsx(src, dst, ctx: Ctx):
        write_xlsx(read_sheets(src, "xls"), dst)
except ImportError:
    pass


@route(["xlsx", "ods", "json"], "csv", cost=1, multi=True)
def to_csv(src, dst_dir, ctx: Ctx):
    sep = ctx.opt("csv_sep", ";")
    sep = "\t" if sep in ("tab", "\\t") else sep
    sheets = read_sheets(src, ctx.src_ext)
    for name, df in sheets.items():
        safe = "".join(ch for ch in str(name) if ch.isalnum() or ch in " -_") or "Tabelle"
        df.to_csv(os.path.join(dst_dir, f"{safe}.csv"), sep=sep, index=False,
                  encoding="utf-8-sig", decimal="," if sep == ";" else ".")


@route(["xlsx", "csv", "ods"], "json", cost=1)
def to_json(src, dst, ctx: Ctx):
    sheets = read_sheets(src, ctx.src_ext)
    out = {}
    for name, df in sheets.items():
        out[name] = json.loads(df.to_json(orient="records", date_format="iso", force_ascii=False))
    data = next(iter(out.values())) if len(out) == 1 else out
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@route(["xlsx", "csv", "json", "ods"], "html", cost=1)
def table_to_html(src, dst, ctx: Ctx):
    import html as h

    sheets = read_sheets(src, ctx.src_ext)
    title = h.escape(ctx.opt("_title", "Tabelle"))
    parts = [f"""<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>{title}</title>
<style>body{{font-family:Segoe UI,Arial,sans-serif;margin:2rem;color:#222}}
table{{border-collapse:collapse;margin-bottom:2rem;font-size:14px}}
th{{background:#2F5597;color:#fff;text-align:left}}th,td{{border:1px solid #ccc;padding:6px 10px}}
tr:nth-child(even) td{{background:#f4f6fa}}</style></head><body><h1>{title}</h1>"""]
    for name, df in sheets.items():
        if len(sheets) > 1:
            parts.append(f"<h2>{h.escape(str(name))}</h2>")
        parts.append(df.fillna("").to_html(index=False, border=0))
    parts.append("</body></html>")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
