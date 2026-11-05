"""Bilder (Pillow) sowie Audio und Video (ffmpeg)."""
from __future__ import annotations

import io
import os
import re
import subprocess

from . import tools
from .engine import ConversionError, Ctx, route
from .formats import RASTER, AUDIO, VIDEO

PIL_FORMAT = {"png": "PNG", "jpg": "JPEG", "webp": "WEBP", "bmp": "BMP", "gif": "GIF",
              "tiff": "TIFF", "ico": "ICO", "pdf": "PDF"}


def _flatten(img, bg=(255, 255, 255)):
    from PIL import Image

    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        base = Image.new("RGB", img.size, bg)
        base.paste(img, mask=img.split()[-1])
        return base
    return img.convert("RGB")


def open_image(path):
    from PIL import Image, ImageOps

    if path.lower().endswith((".heic", ".heif")):
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except Exception:
            raise ConversionError("Für HEIC-Bilder wird 'pillow-heif' benötigt.")
    img = Image.open(path)
    try:
        img = ImageOps.exif_transpose(img)  # Handyfotos richtig drehen
    except Exception:
        pass
    return img


def save_pil(img, path, fmt, ctx: Ctx):
    quality = int(ctx.opt("image_quality", 90))
    pf = PIL_FORMAT[fmt]
    if fmt in ("jpg", "bmp", "pdf"):
        img = _flatten(img)
        kw = {"quality": quality, "optimize": True} if fmt == "jpg" else {}
        if fmt == "pdf":
            kw = {"resolution": 150}
        img.save(path, pf, **kw)
    elif fmt == "webp":
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")
        img.save(path, pf, quality=quality, method=4)
    elif fmt == "gif":
        img.convert("RGBA" if img.mode in ("RGBA", "LA", "P") else "RGB").save(path, pf)
    elif fmt == "ico":
        img = img.convert("RGBA")
        side = max(img.size)
        from PIL import Image
        sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        sq.paste(img, ((side - img.width) // 2, (side - img.height) // 2))
        sizes = [s for s in (16, 24, 32, 48, 64, 128, 256) if s <= side] or [16]
        sq.save(path, pf, sizes=[(s, s) for s in sizes])
    elif fmt == "png":
        if img.mode not in ("RGB", "RGBA", "L", "LA", "P", "I", "1"):
            img = img.convert("RGBA")
        img.save(path, pf, optimize=True)
    elif fmt == "tiff":
        if img.mode in ("P", "CMYK"):
            img = img.convert("RGB")
        img.save(path, pf, compression="tiff_deflate")
    else:
        img.save(path, pf)


# ------------------------------------------------------------ Bild ↔ Bild
@route(RASTER, RASTER, cost=1)
def image_to_image(src, dst, ctx: Ctx):
    img = open_image(src)
    if ctx.src_ext == "gif" and ctx.dst_ext == "webp" and getattr(img, "is_animated", False):
        img.save(dst, "WEBP", save_all=True, quality=int(ctx.opt("image_quality", 90)))
        return
    if getattr(img, "n_frames", 1) > 1 and ctx.dst_ext != "gif":
        img.seek(0)
    if ctx.src_ext == "gif" and ctx.dst_ext == "gif":
        img.save(dst, save_all=True)
        return
    save_pil(img, dst, ctx.dst_ext, ctx)


@route("heic", ["png", "jpg"], cost=1, backend="heif")
def heic_to_image(src, dst, ctx: Ctx):
    save_pil(open_image(src), dst, ctx.dst_ext, ctx)


@route(RASTER, "pdf", cost=1)
def image_to_pdf(src, dst, ctx: Ctx):
    img = open_image(src)
    frames = []
    try:
        n = getattr(img, "n_frames", 1)
        for i in range(n if ctx.src_ext == "tiff" else 1):  # mehrseitige TIFFs komplett
            img.seek(i)
            frames.append(_flatten(img.copy()))
    except EOFError:
        pass
    frames[0].save(dst, "PDF", resolution=float(img.info.get("dpi", (150, 150))[0] or 150),
                   save_all=True, append_images=frames[1:])


@route(RASTER, "docx", cost=1)
def image_to_docx(src, dst, ctx: Ctx):
    from docx import Document

    doc = Document()
    sec = doc.sections[0]
    width = sec.page_width - sec.left_margin - sec.right_margin
    img = _flatten(open_image(src))
    bio = io.BytesIO()
    img.save(bio, "PNG")
    bio.seek(0)
    height_ratio = img.height / img.width
    max_h = sec.page_height - sec.top_margin - sec.bottom_margin
    if width * height_ratio > max_h:
        doc.add_picture(bio, height=max_h)
    else:
        doc.add_picture(bio, width=width)
    doc.save(dst)


@route(RASTER, "pptx", cost=1)
def image_to_pptx(src, dst, ctx: Ctx):
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width, prs.slide_height = 12192000, 6858000
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    img = _flatten(open_image(src))
    bio = io.BytesIO()
    img.save(bio, "PNG")
    bio.seek(0)
    sw, sh = prs.slide_width, prs.slide_height
    r = min(sw / img.width, sh / img.height)
    w, h = int(img.width * r), int(img.height * r)
    slide.shapes.add_picture(bio, (sw - w) // 2, (sh - h) // 2, w, h)
    prs.save(dst)


# ------------------------------------------------------------ Audio / Video
AUDIO_ARGS = {
    "mp3": ["-vn", "-c:a", "libmp3lame", "-q:a", "2"],
    "wav": ["-vn", "-c:a", "pcm_s16le"],
    "flac": ["-vn", "-c:a", "flac"],
    "ogg": ["-vn", "-c:a", "libvorbis", "-q:a", "5"],
    "opus": ["-vn", "-c:a", "libopus", "-b:a", "128k"],
    "m4a": ["-vn", "-c:a", "aac", "-b:a", "192k"],
    "aac": ["-vn", "-c:a", "aac", "-b:a", "192k", "-f", "adts"],
    "wma": ["-vn", "-c:a", "wmav2", "-b:a", "192k"],
}

_EVEN = ["-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2"]
VIDEO_ARGS = {
    "mp4": ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart"],
    "mkv": ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "160k"],
    "mov": ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "160k"],
    "flv": ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "128k", "-ar", "44100"],
    "avi": ["-c:v", "mpeg4", "-q:v", "4", "-c:a", "libmp3lame", "-q:a", "3"],
    "webm": ["-c:v", "libvpx-vp9", "-crf", "33", "-b:v", "0", "-deadline", "realtime",
             "-cpu-used", "8", "-row-mt", "1", "-c:a", "libopus", "-b:a", "128k"],
    "wmv": ["-c:v", "wmv2", "-b:v", "3M", "-c:a", "wmav2", "-b:a", "160k"],
}

QUALITY_CRF = {"hoch": "19", "normal": "23", "klein": "28"}


def _duration(ffmpeg, src) -> float:
    try:
        r = tools.run([ffmpeg, "-hide_banner", "-i", src], timeout=60)
        m = re.search(rb"Duration: (\d+):(\d+):(\d+\.\d+)", r.stderr or b"")
        if m:
            h, mi, s = m.groups()
            return int(h) * 3600 + int(mi) * 60 + float(s)
    except Exception:
        pass
    return 0.0


def run_ffmpeg(src, dst, args, ctx: Ctx):
    ffmpeg = tools.find_ffmpeg()
    if not ffmpeg:
        raise ConversionError("ffmpeg wurde nicht gefunden.")
    total = _duration(ffmpeg, src)
    cmd = [ffmpeg, "-hide_banner", "-y", "-i", src] + args + ["-progress", "pipe:1", "-nostats", dst]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL,
                            creationflags=tools.NO_WINDOW)
    import threading
    err_buf = []
    t = threading.Thread(target=lambda: err_buf.append(proc.stderr.read()), daemon=True)
    t.start()
    for raw in proc.stdout:
        line = raw.decode(errors="ignore").strip()
        if line.startswith("out_time_ms=") and total > 0:
            try:
                sec = int(line.split("=")[1]) / 1_000_000
                ctx.progress(min(sec / total, 0.99), f"{int(sec)} / {int(total)} s")
            except ValueError:
                pass
    proc.wait()
    t.join(timeout=5)
    if proc.returncode != 0:
        msg = (err_buf[0] if err_buf else b"").decode(errors="ignore").strip().splitlines()
        raise ConversionError("ffmpeg: " + (msg[-1] if msg else "unbekannter Fehler"))


def _ffmpeg_with_fallback(src, dst, args, ctx):
    try:
        run_ffmpeg(src, dst, args, ctx)
    except ConversionError as e:
        # Fallback: ffmpeg wählt die Standard-Codecs für den Container selbst
        print("ffmpeg-Fallback:", e)
        if os.path.exists(dst):
            os.remove(dst)
        run_ffmpeg(src, dst, [a for a in args if a in ("-vn",)], ctx)


@route(AUDIO + VIDEO, AUDIO, cost=1, backend="ffmpeg")
def to_audio(src, dst, ctx: Ctx):
    args = list(AUDIO_ARGS[ctx.dst_ext])
    br = ctx.opt("audio_bitrate")
    if br and ctx.dst_ext in ("mp3",):
        args = ["-vn", "-c:a", "libmp3lame", "-b:a", f"{br}k"]
    _ffmpeg_with_fallback(src, dst, args, ctx)


@route(VIDEO + ["gif"], VIDEO, cost=1, backend="ffmpeg")
def to_video(src, dst, ctx: Ctx):
    args = list(VIDEO_ARGS[ctx.dst_ext])
    crf = QUALITY_CRF.get(ctx.opt("video_quality", "normal"))
    if crf and "-crf" in args and ctx.dst_ext != "webm":
        args[args.index("-crf") + 1] = crf
    if ctx.src_ext == "gif":
        # GIF hat keinen Ton: alle Audio-Optionen entfernen
        if "-c:a" in args:
            args = args[:args.index("-c:a")]
        args += ["-an"]
    if ctx.dst_ext in ("mp4", "mkv", "mov", "flv"):
        args = _EVEN + args  # H.264 braucht gerade Bildmaße
    _ffmpeg_with_fallback(src, dst, args, ctx)


@route(VIDEO, "gif", cost=1, backend="ffmpeg")
def video_to_gif(src, dst, ctx: Ctx):
    fps = int(ctx.opt("gif_fps", 12))
    width = int(ctx.opt("gif_width", 640))
    vf = (f"fps={fps},scale='min({width},iw)':-2:flags=lanczos,split[s0][s1];"
          f"[s0]palettegen=stats_mode=diff[p];[s1][p]paletteuse=dither=bayer:bayer_scale=4")
    run_ffmpeg(src, dst, ["-vf", vf, "-loop", "0", "-an"], ctx)


@route(VIDEO, ["png", "jpg"], cost=1.5, backend="ffmpeg")
def video_thumbnail(src, dst, ctx: Ctx):
    """Standbild aus der Mitte des Videos."""
    ffmpeg = tools.find_ffmpeg()
    t = _duration(ffmpeg, src) / 2
    r = tools.run([ffmpeg, "-hide_banner", "-y", "-ss", f"{t:.2f}", "-i", src, "-frames:v", "1", dst], timeout=120)
    if r.returncode != 0 or not os.path.exists(dst):
        raise ConversionError("Standbild konnte nicht erzeugt werden.")
