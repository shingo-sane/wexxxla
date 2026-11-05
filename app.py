"""Wexxxla – Startprogramm.

Startet einen kleinen lokalen Server und öffnet die Oberfläche als App-Fenster
(Microsoft Edge oder Chrome im App-Modus). Alles läuft lokal auf dem eigenen PC.

Aufrufe:
  Wexxxla.exe                       -> Oberfläche
  Wexxxla.exe --cli DATEI ZIEL [ORDNER]   -> ohne Oberfläche umwandeln
"""
from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import uuid
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from converter import __version__, VERSION_NAME, engine, tools
from converter.formats import FORMATS, CATEGORY_ORDER, normalize_ext

APP_NAME = "Wexxxla"
PORT = 47821
IS_WIN = os.name == "nt"


# ---------------------------------------------------------------- Pfade
def resource(*parts):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def config_dir():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(base, "Wexxxla")
    os.makedirs(d, exist_ok=True)
    return d


def default_out_dir():
    home = os.path.expanduser("~")
    dl = os.path.join(home, "Downloads")
    return os.path.join(dl if os.path.isdir(dl) else home, "Konvertiert")


UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "ufc_uploads")
SETTINGS_FILE = os.path.join(config_dir(), "settings.json")
DEFAULT_SETTINGS = {
    "settings_version": 2,
    "out_dir": default_out_dir(),
    "office_engine": "auto",
    "options": {
        "image_quality": 90,
        "pdf_ppt_mode": "editable",
        "pdf_word_mode": "auto",
        "pdf_dpi": 150,
        "csv_sep": ";",
        "video_quality": "normal",
        "ppt_word_images": "1",
        "gif_fps": 12,
        "gif_width": 640,
    },
}


def load_settings():
    s = json.loads(json.dumps(DEFAULT_SETTINGS))
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            saved = json.load(f)
        s.update({k: v for k, v in saved.items() if k != "options"})
        s["options"].update(saved.get("options", {}))
        if saved.get("settings_version", 1) < 2:  # neue Standardwerte für PDF → PowerPoint/Word
            s["options"]["pdf_ppt_mode"] = "editable"
            s["options"]["pdf_word_mode"] = "auto"
    except Exception:
        pass
    return s


def save_settings(s):
    s["settings_version"] = 2
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception:
        traceback.print_exc()


# ---------------------------------------------------------------- Zustand
STATE = {
    "settings": load_settings(),
    "uploads": {},      # id -> {path, name, ext, size}
    "jobs": {},         # id -> job dict
    "produced": set(),  # erlaubte Download-/Öffnen-Pfade
    "last_ping": time.time(),
    "bye_at": None,
}
LOCK = threading.Lock()
engine.set_office_engine(STATE["settings"]["office_engine"])


def file_info(uid):
    u = STATE["uploads"][uid]
    return {
        "id": uid, "name": u["name"], "ext": u["ext"], "size": u["size"],
        "supported": u["ext"] in FORMATS,
        "targets": engine.targets_for(u["ext"]) if u["ext"] in FORMATS else [],
        "unavailable": engine.unavailable_targets(u["ext"]) if u["ext"] in FORMATS else {},
    }


# ---------------------------------------------------------------- Jobs
def run_job(job):
    opts = dict(STATE["settings"]["options"])
    out_dir = job["out_dir"]
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception as e:
        for it in job["items"]:
            it.update(state="error", error=f"Zielordner nicht nutzbar: {e}")
        job["done"] = True
        return

    if job.get("merge"):
        paths = [STATE["uploads"][it["id"]]["path"] for it in job["items"]]
        names = [STATE["uploads"][it["id"]]["name"] for it in job["items"]]
        for it in job["items"]:
            it["state"] = "running"

        def prog(p, m=""):
            for it in job["items"]:
                it["progress"], it["msg"] = p, m

        try:
            # Originalnamen für die Zwischenschritte verwenden
            tmp = tempfile.mkdtemp(prefix="ufc_named_")
            named = []
            for i, (p, n) in enumerate(zip(paths, names)):
                dst = os.path.join(tmp, f"{i:03d}_{n}")
                shutil.copyfile(p, dst)
                named.append(dst)
            name = os.path.splitext(names[0])[0] + "_zusammengefuegt"
            out = engine.merge_to_pdf(named, out_dir, name, opts, prog)
            shutil.rmtree(tmp, ignore_errors=True)
            with LOCK:
                STATE["produced"].add(os.path.abspath(out))
            for it in job["items"]:
                it.update(state="done", progress=1.0, msg="Fertig",
                          outputs=[{"path": out, "name": os.path.basename(out), "isdir": False}])
        except Exception as e:
            traceback.print_exc()
            for it in job["items"]:
                it.update(state="error", error=str(e))
        job["done"] = True
        return

    for it in job["items"]:
        u = STATE["uploads"].get(it["id"])
        if not u:
            it.update(state="error", error="Datei nicht mehr vorhanden.")
            continue
        it["state"] = "running"

        def prog(p, m="", it=it):
            it["progress"], it["msg"] = round(p, 3), m

        try:
            t0 = time.time()
            stem = os.path.splitext(u["name"])[0]
            outs = engine.convert(u["path"], job["target"], out_dir, opts, prog, out_name=stem)
            with LOCK:
                for o in outs:
                    STATE["produced"].add(os.path.abspath(o))
            it.update(state="done", progress=1.0, msg=f"Fertig in {time.time() - t0:.1f} s",
                      outputs=[{"path": o, "name": os.path.basename(o), "isdir": os.path.isdir(o)} for o in outs])
        except Exception as e:
            traceback.print_exc()
            it.update(state="error", error=str(e))
    job["done"] = True


# ---------------------------------------------------------------- Helfer
def reveal(path):
    path = os.path.abspath(path)
    if IS_WIN:
        if os.path.isdir(path):
            os.startfile(path)
        else:
            subprocess.Popen(["explorer", "/select,", path])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", path if os.path.isdir(path) else os.path.dirname(path)])


def open_file(path):
    if IS_WIN:
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def pick_folder(initial):
    """Ordnerauswahl in einem eigenen Prozess (Tk braucht den Hauptthread)."""
    result = os.path.join(tempfile.gettempdir(), f"ufc_pick_{uuid.uuid4().hex[:8]}.txt")
    base = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__)]
    try:
        subprocess.run(base + ["--pick-folder", initial or "", result], stdin=subprocess.DEVNULL,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=900)
        with open(result, encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""
    finally:
        try:
            os.remove(result)
        except OSError:
            pass


def pick_folder_dialog(initial):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        p = filedialog.askdirectory(initialdir=initial if os.path.isdir(initial) else None,
                                    title="Zielordner wählen", mustexist=False)
        root.destroy()
        return p or ""
    except Exception:
        if IS_WIN:  # Fallback ohne Tk
            ps = ("Add-Type -AssemblyName System.Windows.Forms;"
                  "$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
                  "$d.Description='Zielordner wählen';"
                  "if($d.ShowDialog() -eq 'OK'){[Console]::OutputEncoding=[Text.Encoding]::UTF8;$d.SelectedPath}")
            r = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", ps], stdout=subprocess.PIPE,
                               stdin=subprocess.DEVNULL, creationflags=tools.NO_WINDOW)
            return r.stdout.decode("utf-8", errors="ignore").strip()
        return ""


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "Wexxxla/" + __version__

    def log_message(self, fmt, *args):  # ruhig bleiben
        pass

    # ---- Antworten
    def send_json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def read_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        return json.loads(self.rfile.read(n).decode("utf-8"))

    def _local_only(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        if host not in ("127.0.0.1", "localhost"):
            self.send_error(403)
            return False
        if self.command == "POST":
            origin = self.headers.get("Origin")
            if origin and not origin.startswith(("http://127.0.0.1:", "http://localhost:")):
                self.send_error(403)
                return False
        return True

    # ---- GET
    def do_GET(self):
        if not self._local_only():
            return
        url = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(url.query)
        if url.path in ("/", "/index.html"):
            with open(resource("web", "index.html"), "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)
        elif url.path == "/favicon.ico":
            svg = (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 26 26"><rect x="1" y="3" width="13" height="17" '
                   b'rx="2" fill="#B8312F"/><rect x="12" y="6" width="13" height="17" rx="2" fill="#2843B8"/></svg>')
            self.send_response(200)
            self.send_header("Content-Type", "image/svg+xml")
            self.send_header("Content-Length", str(len(svg)))
            self.end_headers()
            self.wfile.write(svg)
        elif url.path == "/api/info":
            st = tools.status()
            self.send_json({
                "app": APP_NAME, "version": __version__, "version_name": VERSION_NAME,
                "formats": {k: {"label": v[0], "cat": v[1]} for k, v in FORMATS.items()},
                "categories": CATEGORY_ORDER,
                "tools": {"libreoffice": bool(st["libreoffice"]), "msoffice": st["msoffice"],
                          "ffmpeg": bool(st["ffmpeg"]), "pandoc": bool(st["pandoc"])},
                "settings": STATE["settings"],
            })
        elif url.path == "/api/job":
            job = STATE["jobs"].get((q.get("id") or [""])[0])
            if not job:
                return self.send_json({"error": "Unbekannter Auftrag"}, 404)
            self.send_json({"done": job["done"], "items": job["items"]})
        elif url.path == "/api/download":
            path = os.path.abspath((q.get("path") or [""])[0])
            if path not in STATE["produced"] or not os.path.exists(path):
                return self.send_error(404)
            if os.path.isdir(path):
                bio = io.BytesIO()
                with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
                    for f in sorted(os.listdir(path)):
                        z.write(os.path.join(path, f), f)
                data, name = bio.getvalue(), os.path.basename(path) + ".zip"
            else:
                with open(path, "rb") as f:
                    data = f.read()
                name = os.path.basename(path)
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + urllib.parse.quote(name))
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_error(404)

    # ---- POST
    def do_POST(self):
        if not self._local_only():
            return
        url = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(url.query)
        try:
            if url.path == "/api/upload":
                name = os.path.basename((q.get("name") or ["datei"])[0]) or "datei"
                n = int(self.headers.get("Content-Length") or 0)
                uid = uuid.uuid4().hex[:12]
                d = os.path.join(UPLOAD_DIR, uid)
                os.makedirs(d, exist_ok=True)
                path = os.path.join(d, name)
                left = n
                with open(path, "wb") as f:
                    while left > 0:
                        chunk = self.rfile.read(min(1 << 20, left))
                        if not chunk:
                            break
                        f.write(chunk)
                        left -= len(chunk)
                STATE["uploads"][uid] = {"path": path, "name": name, "ext": normalize_ext(name), "size": n}
                return self.send_json(file_info(uid))
            if url.path == "/api/remove":
                uid = self.read_json().get("id")
                u = STATE["uploads"].pop(uid, None)
                if u:
                    shutil.rmtree(os.path.dirname(u["path"]), ignore_errors=True)
                return self.send_json({"ok": True})
            if url.path == "/api/convert":
                body = self.read_json()
                ids = [i for i in body.get("ids", []) if i in STATE["uploads"]]
                if not ids:
                    return self.send_json({"error": "Keine Dateien ausgewählt."}, 400)
                out_dir = body.get("out_dir") or STATE["settings"]["out_dir"]
                jid = uuid.uuid4().hex[:10]
                job = {"id": jid, "target": body.get("target"), "out_dir": out_dir,
                       "merge": bool(body.get("merge")), "done": False,
                       "items": [{"id": i, "name": STATE["uploads"][i]["name"], "state": "waiting",
                                  "progress": 0, "msg": "Wartet …", "outputs": [], "error": ""} for i in ids]}
                STATE["jobs"][jid] = job
                threading.Thread(target=run_job, args=(job,), daemon=True).start()
                return self.send_json({"job": jid})
            if url.path == "/api/route":
                body = self.read_json()
                p = engine.find_path(body.get("src", ""), body.get("dst", ""))
                return self.send_json({"route": engine.describe_path(p) if p else ""})
            if url.path == "/api/settings":
                body = self.read_json()
                s = STATE["settings"]
                if "out_dir" in body and body["out_dir"]:
                    s["out_dir"] = body["out_dir"]
                if "office_engine" in body:
                    s["office_engine"] = body["office_engine"]
                    engine.set_office_engine(body["office_engine"])
                s["options"].update(body.get("options", {}))
                save_settings(s)
                return self.send_json({"settings": s})
            if url.path == "/api/refresh":
                ids = self.read_json().get("ids", [])
                return self.send_json({"files": [file_info(i) for i in ids if i in STATE["uploads"]]})
            if url.path == "/api/pick-folder":
                p = pick_folder(STATE["settings"]["out_dir"])
                if p:
                    STATE["settings"]["out_dir"] = os.path.normpath(p)
                    save_settings(STATE["settings"])
                return self.send_json({"out_dir": STATE["settings"]["out_dir"], "changed": bool(p)})
            if url.path == "/api/open":
                body = self.read_json()
                path = os.path.abspath(body.get("path", ""))
                if body.get("what") == "outdir":
                    path = os.path.abspath(STATE["settings"]["out_dir"])
                    os.makedirs(path, exist_ok=True)
                    reveal(path)
                elif path in STATE["produced"] and os.path.exists(path):
                    (reveal if body.get("what") == "reveal" else open_file)(path)
                else:
                    return self.send_json({"error": "Datei nicht gefunden."}, 404)
                return self.send_json({"ok": True})
            if url.path == "/api/open-url":
                url_ = self.read_json().get("url", "")
                if url_.startswith("https://"):
                    webbrowser.open(url_)
                return self.send_json({"ok": True})
            if url.path == "/api/ping":
                STATE["last_ping"], STATE["bye_at"] = time.time(), None
                return self.send_json({"ok": True})
            if url.path == "/api/bye":
                STATE["bye_at"] = time.time()
                return self.send_json({"ok": True})
            self.send_error(404)
        except Exception as e:
            traceback.print_exc()
            self.send_json({"error": str(e)}, 500)


# ---------------------------------------------------------------- Fenster
def find_browser():
    if IS_WIN:
        cands = []
        for base in (os.environ.get("PROGRAMFILES(X86)"), os.environ.get("PROGRAMFILES"),
                     os.environ.get("LOCALAPPDATA")):
            if base:
                cands += [os.path.join(base, r"Microsoft\Edge\Application\msedge.exe"),
                          os.path.join(base, r"Google\Chrome\Application\chrome.exe")]
        return next((c for c in cands if os.path.isfile(c)), None)
    for n in ("google-chrome", "chromium", "chromium-browser", "microsoft-edge"):
        p = shutil.which(n)
        if p:
            return p
    return None


def open_window(url):
    """Öffnet die Oberfläche als App-Fenster. Gibt den Browser-Prozess zurück (oder None)."""
    br = find_browser()
    if br:
        profile = os.path.join(config_dir(), "window-profile")
        try:
            return subprocess.Popen([br, f"--app={url}", f"--user-data-dir={profile}",
                                     "--window-size=1200,860", "--no-first-run",
                                     "--no-default-browser-check", "--disable-sync"],
                                    stdin=subprocess.DEVNULL, creationflags=tools.NO_WINDOW)
        except Exception:
            traceback.print_exc()
    webbrowser.open(url)
    return None


def port_is_ours(port):
    try:
        import urllib.request
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/info", timeout=1.5) as r:
            return json.loads(r.read()).get("app") == APP_NAME
    except Exception:
        return False


def serve():
    # läuft schon eine Instanz? -> nur neues Fenster öffnen
    if port_is_ours(PORT):
        open_window(f"http://127.0.0.1:{PORT}/")
        return
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    except OSError:  # Port belegt -> beliebigen freien Port nehmen
        httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    port = httpd.server_address[1]
    httpd.daemon_threads = True
    url = f"http://127.0.0.1:{port}/"
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    print(f"{APP_NAME} läuft auf {url}")
    if os.environ.get("UFC_NO_WINDOW"):
        proc = None
    else:
        proc = open_window(url)
    started = time.time()
    STATE["last_ping"] = time.time()
    # Beenden, wenn das Fenster geschlossen wurde und nichts mehr läuft
    while True:
        time.sleep(1)
        busy = any(not j["done"] for j in STATE["jobs"].values())
        if busy:
            continue
        if os.environ.get("UFC_NO_WINDOW"):
            continue
        now = time.time()
        if STATE["bye_at"] and now - STATE["bye_at"] > 4:
            break
        if proc is not None and proc.poll() is not None and now - started > 10 and now - STATE["last_ping"] > 8:
            break
        if now - STATE["last_ping"] > 150:
            break
    httpd.shutdown()
    shutil.rmtree(UPLOAD_DIR, ignore_errors=True)


# ---------------------------------------------------------------- CLI
def cli(argv):
    if len(argv) < 2:
        print("Aufruf: --cli DATEI ZIELFORMAT [AUSGABEORDNER]")
        return 2
    src, target = argv[0], argv[1]
    out = argv[2] if len(argv) > 2 else os.path.dirname(os.path.abspath(src))
    p = engine.find_path(normalize_ext(src), target)
    print("Weg:", engine.describe_path(p) if p else "keiner")
    res = engine.convert(src, target, out, STATE["settings"]["options"],
                         lambda q, m="": print(f"\r{int(q * 100):3d}%  {m:<40}", end="", flush=True))
    print("\nErstellt:", *res, sep="\n  ")
    return 0


def main():
    if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
        log = open(os.path.join(config_dir(), "log.txt"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = log
    args = sys.argv[1:]
    if args[:1] == ["--pick-folder"]:
        p = pick_folder_dialog(args[1] if len(args) > 1 else "")
        if len(args) > 2:
            with open(args[2], "w", encoding="utf-8") as f:
                f.write(p)
        else:
            print(p)
        return 0
    if args[:1] == ["--cli"]:
        return cli(args[1:])
    serve()
    return 0


if __name__ == "__main__":
    sys.exit(main())
