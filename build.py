"""Baut Wexxxla.exe mit PyInstaller (wird von build_exe.bat und von GitHub Actions benutzt)."""
import os
import sys

ICON = os.path.join("assets", "icon.ico")


def make_icon(path=ICON):
    """Zeichnet das Wexxxla-Icon (zwei Blätter mit Pfeil) als .ico."""
    from PIL import Image, ImageDraw

    size = 1024
    u = size / 26
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([1 * u, 3 * u, 14 * u, 20 * u], radius=2 * u, fill=(184, 49, 47, 255))
    d.rounded_rectangle([12 * u, 6 * u, 25 * u, 23 * u], radius=2 * u, fill=(40, 67, 184, 255))
    w = int(1.9 * u)
    d.line([(15 * u, 14.5 * u), (21 * u, 14.5 * u)], fill="white", width=w)
    d.line([(18.5 * u, 12 * u), (21 * u, 14.5 * u), (18.5 * u, 17 * u)], fill="white", width=w, joint="curve")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    im.save(path, sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])


if __name__ == "__main__":
    import PyInstaller.__main__

    if not os.path.exists(ICON):
        make_icon()
    sep = ";" if sys.platform.startswith("win") else ":"
    PyInstaller.__main__.run([
        "--noconfirm", "--clean", "--windowed",
        "--name", "Wexxxla",
        "--icon", ICON,
        "--add-data", f"web{sep}web",
        "--collect-submodules", "converter",
        "--collect-all", "pypdfium2", "--collect-all", "pypdfium2_raw",
        "--collect-data", "pdfminer", "--collect-data", "docx", "--collect-data", "pptx",
        "--collect-all", "imageio_ffmpeg", "--collect-all", "pypandoc",
        "--collect-all", "pillow_heif",
        "app.py",
    ])
