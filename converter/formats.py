"""Katalog aller unterstützten Dateiformate."""

# key: (Anzeigename, Kategorie)
FORMATS: dict[str, tuple[str, str]] = {
    # Dokumente
    "pdf": ("PDF", "PDF"),
    "docx": ("Word (.docx)", "Dokument"),
    "doc": ("Word 97-2003 (.doc)", "Dokument"),
    "odt": ("OpenDocument Text (.odt)", "Dokument"),
    "rtf": ("Rich Text (.rtf)", "Dokument"),
    "txt": ("Text (.txt)", "Dokument"),
    "md": ("Markdown (.md)", "Dokument"),
    "html": ("Webseite (.html)", "Dokument"),
    "epub": ("E-Book (.epub)", "Dokument"),
    # Präsentationen
    "pptx": ("PowerPoint (.pptx)", "Präsentation"),
    "ppt": ("PowerPoint 97-2003 (.ppt)", "Präsentation"),
    "odp": ("OpenDocument Präsentation (.odp)", "Präsentation"),
    # Tabellen
    "xlsx": ("Excel (.xlsx)", "Tabelle"),
    "xls": ("Excel 97-2003 (.xls)", "Tabelle"),
    "ods": ("OpenDocument Tabelle (.ods)", "Tabelle"),
    "csv": ("CSV (.csv)", "Tabelle"),
    "json": ("JSON (.json)", "Tabelle"),
    # Bilder
    "png": ("PNG", "Bild"),
    "jpg": ("JPG", "Bild"),
    "webp": ("WebP", "Bild"),
    "bmp": ("BMP", "Bild"),
    "gif": ("GIF", "Bild"),
    "tiff": ("TIFF", "Bild"),
    "ico": ("Icon (.ico)", "Bild"),
    "heic": ("HEIC (iPhone-Foto)", "Bild"),
    "svg": ("SVG (Vektor)", "Bild"),
    # Audio
    "mp3": ("MP3", "Audio"),
    "wav": ("WAV", "Audio"),
    "flac": ("FLAC", "Audio"),
    "ogg": ("OGG", "Audio"),
    "m4a": ("M4A", "Audio"),
    "aac": ("AAC", "Audio"),
    "opus": ("Opus", "Audio"),
    "wma": ("WMA", "Audio"),
    # Video
    "mp4": ("MP4", "Video"),
    "mkv": ("MKV", "Video"),
    "mov": ("MOV", "Video"),
    "avi": ("AVI", "Video"),
    "webm": ("WebM", "Video"),
    "wmv": ("WMV", "Video"),
    "flv": ("FLV", "Video"),
}

CATEGORY_ORDER = ["PDF", "Dokument", "Präsentation", "Tabelle", "Bild", "Audio", "Video"]

ALIASES = {
    "jpeg": "jpg", "jpe": "jpg", "jfif": "jpg",
    "tif": "tiff",
    "htm": "html", "xhtml": "html",
    "markdown": "md",
    "heif": "heic",
    "text": "txt", "log": "txt",
    "m4v": "mp4", "mpeg": "mp4", "mpg": "mp4", "3gp": "mp4", "ts": "mp4",
    "oga": "ogg",
    "tsv": "csv",
}


def normalize_ext(name_or_ext: str) -> str:
    ext = name_or_ext.rsplit(".", 1)[-1].lower().strip() if "." in name_or_ext else name_or_ext.lower()
    return ALIASES.get(ext, ext)


def category(fmt: str) -> str:
    return FORMATS.get(fmt, (fmt, "Sonstige"))[1]


def label(fmt: str) -> str:
    return FORMATS.get(fmt, (fmt.upper(), ""))[0]


# Gruppen, die öfter gebraucht werden
WRITER = ["docx", "doc", "odt", "rtf", "txt", "html"]
CALC = ["xlsx", "xls", "ods", "csv"]
IMPRESS = ["pptx", "ppt", "odp"]
RASTER = ["png", "jpg", "webp", "bmp", "gif", "tiff", "ico"]
AUDIO = ["mp3", "wav", "flac", "ogg", "m4a", "aac", "opus", "wma"]
VIDEO = ["mp4", "mkv", "mov", "avi", "webm", "wmv", "flv"]
