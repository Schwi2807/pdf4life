"""PDF4life - local PDF toolbox. Run this file; the UI opens in your browser."""
import json
import os
import socket
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

import pdf_ops

if getattr(sys, "frozen", False):  # PyInstaller bundle
    BASE = Path(sys._MEIPASS)
else:
    BASE = Path(__file__).parent

WORK = Path(tempfile.gettempdir()) / "PDF4life"
WORK.mkdir(exist_ok=True)

app = Flask(__name__, template_folder=str(BASE / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512 MB

PDF_EXT = {".pdf"}
IMG_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp"}
DOC_EXT = {".docx", ".doc"}


# ---------------------------------------------------------------- store

def new_job():
    jid = uuid.uuid4().hex[:12]
    d = WORK / jid
    d.mkdir()
    (d / "out").mkdir()
    return jid, d


def job_dir(jid):
    d = WORK / secure_filename(jid)
    if not d.is_dir():
        raise FileNotFoundError("Unknown or expired file id - please re-upload.")
    return d


def stored_input(jid):
    d = job_dir(jid)
    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    return d / meta["stored"], meta


def save_upload(f, allowed):
    name = secure_filename(f.filename or "file")
    ext = Path(name).suffix.lower()
    if ext not in allowed:
        raise ValueError(f"'{name}': type {ext or '(none)'} not accepted here.")
    jid, d = new_job()
    stored = "input" + ext
    f.save(d / stored)
    meta = {"name": name, "stored": stored}
    if ext == ".pdf":
        import pymupdf
        with pymupdf.open(d / stored) as doc:
            meta["pages"] = doc.page_count
            meta["encrypted"] = doc.needs_pass
    meta["size"] = (d / stored).stat().st_size
    (d / "meta.json").write_text(json.dumps(meta), encoding="utf-8")
    meta["id"] = jid
    return meta


def result(jid, path, note=None):
    p = Path(path)
    out = {"download": f"/api/download/{jid}/{p.name}", "name": p.name,
           "size": p.stat().st_size}
    if note:
        out["note"] = note
    return jsonify(out)


def api(fn):
    """Wrap handlers: exceptions -> JSON {error} instead of HTML 500."""
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except Exception as e:  # noqa: BLE001 - surfaced to the UI on purpose
            app.logger.exception("operation failed")
            return jsonify({"error": str(e) or e.__class__.__name__}), 400
    wrapper.__name__ = fn.__name__
    return wrapper


# ---------------------------------------------------------------- generic routes

@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/health")
def health():
    return jsonify({"ok": True})


LAST_SEEN = time.time()


@app.before_request
def _touch():
    global LAST_SEEN
    LAST_SEEN = time.time()


@app.get("/api/ping")
def ping():
    return jsonify({"ok": True})


@app.post("/api/upload")
@api
def upload():
    kind = request.form.get("kind", "pdf")
    allowed = {"pdf": PDF_EXT, "image": IMG_EXT, "word": DOC_EXT,
               "any": PDF_EXT | IMG_EXT | DOC_EXT}[kind]
    files = request.files.getlist("files")
    if not files:
        raise ValueError("No files received.")
    out = [save_upload(f, allowed) for f in files]
    return jsonify(out)


@app.get("/api/thumb/<jid>/<int:page>")
@api
def thumb(jid, page):
    import pymupdf
    path, _ = stored_input(jid)
    w = min(int(request.args.get("w", 200)), 1600)
    with pymupdf.open(path) as doc:
        p = doc[page]
        zoom = w / max(p.rect.width, 1)
        pix = p.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        png = pix.tobytes("png")
    return app.response_class(png, mimetype="image/png")


@app.get("/api/download/<jid>/<name>")
@api
def download(jid, name):
    p = job_dir(jid) / "out" / secure_filename(name)
    if not p.is_file():
        raise FileNotFoundError("Result expired - please run the tool again.")
    return send_file(p, as_attachment=True, download_name=p.name)


# ---------------------------------------------------------------- tools

@app.post("/api/merge")
@api
def op_merge():
    ids = request.json["ids"]
    paths = [stored_input(i)[0] for i in ids]
    jid, d = new_job()
    out = d / "out" / "merged.pdf"
    pdf_ops.merge(paths, out)
    return result(jid, out)


@app.post("/api/split")
@api
def op_split():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    files = pdf_ops.split(str(path), j["mode"], j["value"], str(d / "out"))
    if len(files) == 1:
        return result(j["id"], files[0])
    z = pdf_ops.zip_files(files, str(d / "out" / f"{Path(meta['name']).stem}_split.zip"))
    return result(j["id"], z, note=f"{len(files)} PDFs inside")


@app.post("/api/compress")
@api
def op_compress():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_compressed.pdf"
    pdf_ops.compress(str(path), j.get("level", "medium"), str(out))
    before, after = meta["size"], out.stat().st_size
    hs = lambda n: f"{n/1e6:.2f} MB" if n >= 1e6 else f"{max(1, round(n/1024))} KB"
    pct = (1 - after / before) * 100 if before else 0
    tail = f"({pct:.0f}% smaller)" if pct >= 1 else "(already well optimized)"
    return result(j["id"], out, note=f"{hs(before)} -> {hs(after)} {tail}")


@app.post("/api/pdf2word")
@api
def op_pdf2word():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}.docx"
    pdf_ops.pdf_to_word(str(path), str(out))
    return result(j["id"], out)


@app.post("/api/word2pdf")
@api
def op_word2pdf():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}.pdf"
    pdf_ops.word_to_pdf(str(path), str(out))
    return result(j["id"], out)


@app.post("/api/pdf2img")
@api
def op_pdf2img():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    files = pdf_ops.pdf_to_images(str(path), j.get("fmt", "jpg"),
                                  j.get("dpi", 150), j.get("pages", ""),
                                  str(d / "out"))
    if len(files) == 1:
        return result(j["id"], files[0])
    z = pdf_ops.zip_files(files, str(d / "out" / f"{Path(meta['name']).stem}_images.zip"))
    return result(j["id"], z, note=f"{len(files)} images inside")


@app.post("/api/img2pdf")
@api
def op_img2pdf():
    ids = request.json["ids"]
    paths = [str(stored_input(i)[0]) for i in ids]
    jid, d = new_job()
    out = d / "out" / "images.pdf"
    pdf_ops.images_to_pdf(paths, str(out))
    return result(jid, out)


@app.post("/api/watermark")
@api
def op_watermark():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_watermarked.pdf"
    image_path = None
    if j.get("image_id"):
        raw = str(stored_input(j["image_id"])[0])
        image_path = pdf_ops.fade_image(raw, float(j.get("opacity", 0.25)),
                                        str(d / "wm_faded.png"))
    pdf_ops.watermark(str(path), str(out), text=j.get("text") or None,
                      image_path=image_path,
                      opacity=float(j.get("opacity", 0.25)),
                      rotation=float(j.get("rotation", 45)),
                      fontsize=float(j.get("fontsize", 48)),
                      color=j.get("color", "#d32f2f"),
                      position=j.get("position", "center"),
                      page_spec=j.get("pages", ""), tile=bool(j.get("tile")))
    return result(j["id"], out)


@app.post("/api/sign")
@api
def op_sign():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_signed.pdf"
    sig = d / "signature.png"
    if j.get("sig_data"):  # dataURL from the drawing canvas
        import base64
        sig.write_bytes(base64.b64decode(j["sig_data"].split(",", 1)[1]))
    elif j.get("sig_id"):
        import shutil
        shutil.copy(stored_input(j["sig_id"])[0], sig)
    else:
        raise ValueError("Draw or upload a signature first.")
    pdf_ops.sign(str(path), str(out), str(sig), j["placements"])
    return result(j["id"], out)


@app.post("/api/organize")
@api
def op_organize():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_organized.pdf"
    pdf_ops.organize(str(path), str(out), j["order"], j["rotations"])
    return result(j["id"], out)


@app.post("/api/pagenumbers")
@api
def op_pagenumbers():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_numbered.pdf"
    pdf_ops.add_page_numbers(str(path), str(out), fmt=j.get("fmt", "{n}"),
                             position=j.get("position", "bottom-center"),
                             start=int(j.get("start", 1)),
                             fontsize=float(j.get("fontsize", 11)),
                             page_spec=j.get("pages", ""),
                             color=j.get("color", "#333333"))
    return result(j["id"], out)


@app.post("/api/crop")
@api
def op_crop():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_cropped.pdf"
    pdf_ops.crop(str(path), str(out), j["rect"], j.get("pages", ""))
    return result(j["id"], out)


@app.post("/api/redact")
@api
def op_redact():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_redacted.pdf"
    n = pdf_ops.redact(str(path), str(out), areas=j.get("areas"),
                       search_text=j.get("search") or None,
                       fill=j.get("fill", "#000000"))
    return result(j["id"], out, note=f"{n} region(s) permanently removed")


@app.post("/api/compare")
@api
def op_compare():
    j = request.json
    pa, ma = stored_input(j["id_a"])
    pb, mb = stored_input(j["id_b"])
    jid, d = new_job()
    out = d / "out" / "comparison_report.pdf"
    diffs, total = pdf_ops.compare(str(pa), str(pb), str(out))
    note = ("No differences found." if diffs == 0
            else f"{diffs} of {total} page(s) differ")
    return result(jid, out, note=note)


@app.post("/api/forms/detect")
@api
def op_form_detect():
    path, _ = stored_input(request.json["id"])
    return jsonify({"fields": pdf_ops.form_fields(str(path))})


@app.post("/api/forms/fill")
@api
def op_form_fill():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_filled.pdf"
    pdf_ops.form_fill(str(path), str(out), j.get("values", {}),
                      flatten=bool(j.get("flatten")))
    return result(j["id"], out)


@app.post("/api/forms/create")
@api
def op_form_create():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_fillable.pdf"
    pdf_ops.form_create(str(path), str(out), j["fields"])
    return result(j["id"], out, note=f"{len(j['fields'])} field(s) added")


@app.get("/api/spans/<jid>/<int:page>")
@api
def op_spans(jid, page):
    path, _ = stored_input(jid)
    return jsonify(pdf_ops.text_spans(str(path), page))


@app.post("/api/edit")
@api
def op_edit():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}_edited.pdf"
    images = {op["image_id"]: str(stored_input(op["image_id"])[0])
              for op in j.get("ops", []) if op.get("image_id")}
    n = pdf_ops.edit_pdf(str(path), str(out), j.get("ops", []), images)
    return result(j["id"], out, note=f"{n} edit(s) applied")


@app.post("/api/pdf2md")
@api
def op_pdf2md():
    j = request.json
    path, meta = stored_input(j["id"])
    d = job_dir(j["id"])
    out = d / "out" / f"{Path(meta['name']).stem}.md"
    pdf_ops.pdf_to_markdown(str(path), str(out))
    return result(j["id"], out)


# ---------------------------------------------------------------- startup

def cleanup_old(hours=24):
    cutoff = time.time() - hours * 3600
    for d in WORK.iterdir():
        try:
            if d.is_dir() and d.stat().st_mtime < cutoff:
                import shutil
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return socket.gethostbyname(socket.gethostname())


def run_server(host, port):
    app.run(host=host, port=port, debug=False, use_reloader=False)


def watchdog(idle_seconds):
    """Exit when no browser tab has talked to us for a while (frozen browser
    mode has no console or window the user could close)."""
    while True:
        time.sleep(5)
        if time.time() - LAST_SEEN > idle_seconds:
            os._exit(0)


def wait_for_server(port, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            socket.create_connection(("127.0.0.1", port), 0.25).close()
            return True
        except OSError:
            time.sleep(0.1)
    return False


def main():
    cleanup_old()
    port = int(os.environ.get("PDF4LIFE_PORT", 0)) or free_port()
    if "--port" in sys.argv:
        port = int(sys.argv[sys.argv.index("--port") + 1])
    if "--no-browser" in sys.argv:
        os.environ["PDF4LIFE_NO_BROWSER"] = "1"

    host = "127.0.0.1"
    if "--lan" in sys.argv or os.environ.get("PDF4LIFE_LAN"):
        host = "0.0.0.0"
    if "--host" in sys.argv:
        host = sys.argv[sys.argv.index("--host") + 1]
    url = f"http://127.0.0.1:{port}"

    browser_mode = "--browser" in sys.argv or os.environ.get("PDF4LIFE_BROWSER")
    if not browser_mode:
        try:
            import webview  # noqa: F401
        except Exception:
            browser_mode = True  # pywebview/WebView2 unavailable - fall back

    frozen = getattr(sys, "frozen", False)
    lan_note = (f"   Shared on your network: http://{lan_ip()}:{port}\n"
                if host != "127.0.0.1" else "")

    if browser_mode:
        print()
        print("  ============================================")
        print("   PDF4life is running")
        print(f"   Open {url} in your browser")
        print(lan_note, end="")
        print("   Close this window to quit.")
        print("  ============================================")
        print()
        if frozen:
            # No console/window to close in a --noconsole exe: quit once no
            # browser tab has pinged for a while so the process never lingers.
            idle = int(os.environ.get("PDF4LIFE_IDLE", 120))
            threading.Thread(target=watchdog, args=(idle,), daemon=True).start()
        if not os.environ.get("PDF4LIFE_NO_BROWSER"):
            threading.Timer(1.0, lambda: webbrowser.open(url)).start()
        run_server(host, port)
        return

    import webview
    threading.Thread(target=run_server, args=(host, port), daemon=True).start()
    wait_for_server(port)
    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.create_window("PDF4life", url, width=1180, height=820,
                          min_size=(760, 560))
    webview.start()
    os._exit(0)  # window closed - stop the Flask thread immediately


if __name__ == "__main__":
    main()
