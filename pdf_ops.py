"""PDF4life - core PDF operations built on PyMuPDF (+ pdf2docx, pymupdf4llm)."""
import io
import os
import re
import shutil
import subprocess
import zipfile
import difflib
from pathlib import Path

import pymupdf as fitz


# ---------------------------------------------------------------- helpers

def parse_ranges(spec, page_count):
    """'1-3,5,7-9' -> list of lists of 0-based page indexes (one list per range)."""
    groups = []
    for part in re.split(r"[,;]+", spec.strip()):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
        elif part.isdigit():
            a = b = int(part)
        else:
            raise ValueError(f"Invalid page range: '{part}'")
        if a < 1 or b > page_count or a > b:
            raise ValueError(f"Range '{part}' outside 1-{page_count}")
        groups.append(list(range(a - 1, b)))
    if not groups:
        raise ValueError("No page ranges given")
    return groups


def pages_from_spec(spec, page_count):
    """Flatten a range spec into a sorted unique list of 0-based indexes.
    Empty spec means all pages."""
    if not spec or not spec.strip() or spec.strip().lower() == "all":
        return list(range(page_count))
    seen = []
    for grp in parse_ranges(spec, page_count):
        for p in grp:
            if p not in seen:
                seen.append(p)
    return sorted(seen)


def normalize_page(page):
    """Remove rotation so stamping coordinates match what the user sees."""
    if page.rotation != 0:
        page.remove_rotation()
    return page


def denorm_rect(page, nrect):
    """Normalized [x0,y0,x1,y1] (0-1, relative to the displayed page) -> fitz.Rect."""
    r = page.rect
    return fitz.Rect(
        r.x0 + nrect[0] * r.width, r.y0 + nrect[1] * r.height,
        r.x0 + nrect[2] * r.width, r.y0 + nrect[3] * r.height,
    )


def save_small(doc, out_path):
    doc.save(out_path, garbage=4, deflate=True, clean=False)
    doc.close()


# ---------------------------------------------------------------- basic ops

def merge(paths, out_path):
    out = fitz.open()
    for p in paths:
        with fitz.open(p) as src:
            out.insert_pdf(src)
    save_small(out, out_path)


def split(path, mode, value, out_dir):
    """mode: 'ranges' (value='1-3,4-6'), 'every' (value=N), 'extract' (value='1,3,5').
    Returns list of created file paths."""
    src = fitz.open(path)
    n = src.page_count
    stem = Path(path).stem
    outputs = []

    def emit(indexes, label):
        d = fitz.open()
        for i in indexes:
            d.insert_pdf(src, from_page=i, to_page=i)
        f = os.path.join(out_dir, f"{stem}_{label}.pdf")
        save_small(d, f)
        outputs.append(f)

    if mode == "every":
        step = max(1, int(value))
        part = 1
        for start in range(0, n, step):
            emit(list(range(start, min(start + step, n))), f"part{part}")
            part += 1
    elif mode == "extract":
        pages = pages_from_spec(value, n)
        emit(pages, "extracted")
    else:  # ranges
        groups = parse_ranges(value, n)
        for i, grp in enumerate(groups, 1):
            label = f"p{grp[0]+1}" if len(grp) == 1 else f"p{grp[0]+1}-{grp[-1]+1}"
            emit(grp, label)
    src.close()
    return outputs


def compress(path, level, out_path):
    """level: 'low' | 'medium' | 'high' (higher = smaller file)."""
    doc = fitz.open(path)
    if level in ("medium", "high"):
        dpi = 120 if level == "medium" else 90
        quality = 70 if level == "medium" else 50
        try:
            doc.rewrite_images(dpi_threshold=dpi + 30, dpi_target=dpi,
                               quality=quality, lossy=True, lossless=True)
        except Exception:
            pass  # older PyMuPDF or odd images - fall back to structural cleanup
        try:
            doc.subset_fonts()
        except Exception:
            pass
    doc.save(out_path, garbage=4, deflate=True, clean=True,
             deflate_images=True, deflate_fonts=True)
    doc.close()


# ---------------------------------------------------------------- conversions

def pdf_to_word(path, out_path):
    from pdf2docx import Converter
    cv = Converter(path)
    try:
        cv.convert(out_path)
    finally:
        cv.close()


def word_to_pdf(path, out_path):
    """Try MS Word COM first (best fidelity), then LibreOffice, else fail clearly."""
    # 1) Microsoft Word via COM
    try:
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        try:
            convert(path, out_path)
        finally:
            pythoncom.CoUninitialize()
        if os.path.exists(out_path):
            return
    except Exception:
        pass
    # 2) LibreOffice
    for soffice in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                    shutil.which("soffice")):
        if soffice and os.path.exists(soffice):
            outdir = os.path.dirname(out_path)
            subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            "--outdir", outdir, path], timeout=180,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            produced = os.path.join(outdir, Path(path).stem + ".pdf")
            if os.path.exists(produced):
                if os.path.abspath(produced) != os.path.abspath(out_path):
                    shutil.move(produced, out_path)
                return
    raise RuntimeError(
        "Word to PDF needs Microsoft Word or LibreOffice installed on this PC. "
        "Neither was found (or Word refused the conversion).")


def pdf_to_images(path, fmt, dpi, page_spec, out_dir):
    doc = fitz.open(path)
    pages = pages_from_spec(page_spec, doc.page_count)
    stem = Path(path).stem
    files = []
    for i in pages:
        pix = doc[i].get_pixmap(dpi=int(dpi))
        f = os.path.join(out_dir, f"{stem}_page{i+1}.{fmt}")
        if fmt in ("jpg", "jpeg"):
            pix.save(f, jpg_quality=90)
        else:
            pix.save(f)
        files.append(f)
    doc.close()
    return files


def images_to_pdf(paths, out_path):
    doc = fitz.open()
    for p in paths:
        with fitz.open(p) as img:
            rect = img[0].rect
            pdfbytes = img.convert_to_pdf()
        with fitz.open("pdf", pdfbytes) as imgpdf:
            page = doc.new_page(width=rect.width, height=rect.height)
            page.show_pdf_page(rect, imgpdf, 0)
    save_small(doc, out_path)


# ---------------------------------------------------------------- stamping

_POSITIONS = {  # (horizontal, vertical) anchors
    "top-left": (0, 0), "top-center": (1, 0), "top-right": (2, 0),
    "center": (1, 1),
    "bottom-left": (0, 2), "bottom-center": (1, 2), "bottom-right": (2, 2),
}


def _anchor_point(rect, pos, margin=36):
    hx, vy = _POSITIONS.get(pos, (1, 1))
    x = [rect.x0 + margin, (rect.x0 + rect.x1) / 2, rect.x1 - margin][hx]
    y = [rect.y0 + margin, (rect.y0 + rect.y1) / 2, rect.y1 - margin][vy]
    return fitz.Point(x, y)


def watermark(path, out_path, text=None, image_path=None, opacity=0.25,
              rotation=45, fontsize=48, color="#d32f2f", position="center",
              page_spec="", tile=False):
    doc = fitz.open(path)
    pages = pages_from_spec(page_spec, doc.page_count)
    rgb = tuple(int(color.lstrip("#")[i:i+2], 16) / 255 for i in (0, 2, 4))
    for i in pages:
        page = normalize_page(doc[i])
        r = page.rect
        if image_path:
            w = r.width * 0.5
            with fitz.open(image_path) as img:
                iw, ih = img[0].rect.width, img[0].rect.height
            h = w * ih / iw
            c = _anchor_point(r, position, margin=r.width * 0.05 + 20)
            target = fitz.Rect(c.x - w / 2, c.y - h / 2, c.x + w / 2, c.y + h / 2) & r
            page.insert_image(target, filename=image_path, overlay=True,
                              keep_proportion=True, alpha=1)
            # PyMuPDF has no direct image opacity; use optional-content-free fallback:
            # pre-bake transparency instead (handled by caller passing a faded PNG).
        elif text:
            positions = [_anchor_point(r, position)]
            if tile:
                positions = [fitz.Point(x, y)
                             for x in (r.width * .25, r.width * .5, r.width * .75)
                             for y in (r.height * .2, r.height * .5, r.height * .8)]
            for c in positions:
                box = fitz.Rect(c.x - 400, c.y - 100, c.x + 400, c.y + 100)
                page.insert_textbox(box, text, fontsize=fontsize, fontname="helv",
                                    color=rgb, align=fitz.TEXT_ALIGN_CENTER,
                                    fill_opacity=opacity, stroke_opacity=opacity,
                                    morph=(c, fitz.Matrix(-float(rotation))))
    save_small(doc, out_path)


def fade_image(src_path, opacity, out_path):
    """Bake opacity into a PNG (for image watermarks / signatures)."""
    from PIL import Image
    img = Image.open(src_path).convert("RGBA")
    alpha = img.getchannel("A").point(lambda a: int(a * opacity))
    img.putalpha(alpha)
    img.save(out_path, "PNG")
    return out_path


def sign(path, out_path, sig_png, placements):
    """placements: [{page:int(1-based), rect:[nx0,ny0,nx1,ny1]}] - signature image
    stamped at each spot, aspect preserved."""
    doc = fitz.open(path)
    for pl in placements:
        page = normalize_page(doc[int(pl["page"]) - 1])
        rect = denorm_rect(page, pl["rect"])
        page.insert_image(rect, filename=sig_png, keep_proportion=True, overlay=True)
    save_small(doc, out_path)


def add_page_numbers(path, out_path, fmt="{n}", position="bottom-center",
                     start=1, fontsize=11, page_spec="", color="#333333"):
    doc = fitz.open(path)
    pages = pages_from_spec(page_spec, doc.page_count)
    total = len(pages)
    rgb = tuple(int(color.lstrip("#")[i:i+2], 16) / 255 for i in (0, 2, 4))
    for k, i in enumerate(pages):
        page = normalize_page(doc[i])
        label = (fmt.replace("{n}", str(start + k)).replace("{N}", str(total))
                    .replace("{p}", str(i + 1)))
        r = page.rect
        c = _anchor_point(r, position, margin=28)
        width = fitz.get_text_length(label, fontname="helv", fontsize=fontsize)
        hx = _POSITIONS.get(position, (1, 2))[0]
        x = [c.x, c.x - width / 2, c.x - width][hx]
        vy = _POSITIONS.get(position, (1, 2))[1]
        y = c.y + (fontsize * 0.35 if vy != 0 else fontsize * 0.35)
        page.insert_text(fitz.Point(x, y), label, fontsize=fontsize,
                         fontname="helv", color=rgb)
    save_small(doc, out_path)


# ---------------------------------------------------------------- page geometry

def organize(path, out_path, order, rotations):
    """order: list of 0-based original page indexes to keep, in new order.
    rotations: parallel list of extra rotation (0/90/180/270) per kept page."""
    doc = fitz.open(path)
    doc.select(order)
    for pos, rot in enumerate(rotations):
        if rot:
            page = doc[pos]
            page.set_rotation((page.rotation + int(rot)) % 360)
    save_small(doc, out_path)


def crop(path, out_path, nrect, page_spec=""):
    doc = fitz.open(path)
    pages = pages_from_spec(page_spec, doc.page_count)
    for i in pages:
        page = doc[i]
        vis = denorm_rect(page, nrect)               # rect in displayed coords
        raw = vis * page.derotation_matrix           # -> unrotated coords
        raw.normalize()
        raw = raw & page.mediabox                    # keep inside mediabox
        if raw.is_empty or not raw.is_valid:
            continue
        page.set_cropbox(raw)
    save_small(doc, out_path)


def redact(path, out_path, areas=None, search_text=None, fill="#000000"):
    """areas: [{page:1-based, rect:[normalized]}]; search_text: string to find & remove.
    True redaction - content underneath is deleted."""
    doc = fitz.open(path)
    rgb = tuple(int(fill.lstrip("#")[i:i+2], 16) / 255 for i in (0, 2, 4))
    touched = set()
    count = 0
    if search_text:
        for i, page in enumerate(doc):
            for quad in page.search_for(search_text, quads=True):
                page.add_redact_annot(quad, fill=rgb)
                touched.add(i)
                count += 1
    for a in (areas or []):
        i = int(a["page"]) - 1
        page = doc[i]
        page.add_redact_annot(denorm_rect(page, a["rect"]), fill=rgb)
        touched.add(i)
        count += 1
    for i in touched:
        doc[i].apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)
    save_small(doc, out_path)
    return count


# ---------------------------------------------------------------- compare

def compare(path_a, path_b, out_path, dpi=100):
    """Build a PDF report: summary + side-by-side pages with differences boxed in red."""
    from PIL import Image, ImageChops, ImageDraw
    da, db = fitz.open(path_a), fitz.open(path_b)
    n = max(da.page_count, db.page_count)
    report = fitz.open()
    summary = []

    def render(doc, i):
        if i >= doc.page_count:
            return None
        pix = doc[i].get_pixmap(dpi=dpi)
        return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")

    def text_of(doc, i):
        return doc[i].get_text() if i < doc.page_count else ""

    for i in range(n):
        ia, ib = render(da, i), render(db, i)
        ta, tb = text_of(da, i), text_of(db, i)
        ratio = difflib.SequenceMatcher(None, ta, tb).ratio() if (ta or tb) else 1.0
        pixel_same = True
        boxes = []
        if ia and ib:
            if ia.size != ib.size:
                ib = ib.resize(ia.size)
            diff = ImageChops.difference(ia, ib).convert("L").point(
                lambda v: 255 if v > 24 else 0)
            bbox = diff.getbbox()
            pixel_same = bbox is None
            if bbox:
                # cluster diff regions per horizontal band for readable boxes
                w, h = diff.size
                band = max(20, h // 40)
                for y in range(0, h, band):
                    strip = diff.crop((0, y, w, min(y + band, h)))
                    bb = strip.getbbox()
                    if bb:
                        boxes.append((bb[0], y + bb[1], bb[2], y + bb[3]))
        identical = pixel_same and ta == tb
        summary.append((i + 1, identical, ratio,
                        ia is None and "missing in A" or ib is None and "missing in B" or ""))
        if identical:
            continue
        # side-by-side page in report
        for img, tag in ((ia, "A"), (ib, "B")):
            if img is None:
                continue
            if tag == "B":
                d = ImageDraw.Draw(img)
                for b in boxes:
                    d.rectangle(b, outline=(220, 30, 30), width=3)
        pw = 842, 595  # A4 landscape pts
        page = report.new_page(width=842, height=595)
        page.insert_text((30, 30), f"Page {i+1}   -   text similarity "
                         f"{ratio*100:.1f}%", fontsize=14, color=(0.8, 0.1, 0.1))
        half = fitz.Rect(20, 50, 421, 575), fitz.Rect(431, 50, 832, 575)
        for img, area, tag in ((ia, half[0], "File A"), (ib, half[1], "File B")):
            page.insert_text((area.x0 + 4, 48), tag, fontsize=9, color=(0.3, 0.3, 0.3))
            if img is None:
                page.insert_text(area.tl + (20, 40), "(page missing)", fontsize=12)
                continue
            buf = io.BytesIO()
            img.save(buf, "PNG")
            page.insert_image(area, stream=buf.getvalue(), keep_proportion=True)
    # summary front page
    front = report.new_page(pno=0, width=595, height=842)
    front.insert_text((50, 60), "PDF Comparison Report", fontsize=22)
    front.insert_text((50, 85), f"A: {Path(path_a).name}", fontsize=10)
    front.insert_text((50, 100), f"B: {Path(path_b).name}", fontsize=10)
    diff_pages = [s for s in summary if not s[1]]
    front.insert_text((50, 130),
                      f"{len(summary)} page(s) compared - "
                      f"{len(diff_pages)} with differences", fontsize=12,
                      color=(0.8, 0.1, 0.1) if diff_pages else (0, 0.5, 0))
    y = 160
    for pno, same, ratio, note in summary:
        line = f"Page {pno}:  {'identical' if same else f'DIFFERENT (text {ratio*100:.0f}% similar)'}"
        if note:
            line += f"  [{note}]"
        front.insert_text((60, y), line, fontsize=10,
                          color=(0, 0.5, 0) if same else (0.8, 0.1, 0.1))
        y += 15
        if y > 800:
            break
    save_small(report, out_path)
    da.close(); db.close()
    return len(diff_pages), len(summary)


# ---------------------------------------------------------------- forms

_WIDGET_TYPES = {
    "text": fitz.PDF_WIDGET_TYPE_TEXT,
    "checkbox": fitz.PDF_WIDGET_TYPE_CHECKBOX,
}


def form_fields(path):
    doc = fitz.open(path)
    fields = []
    for pno, page in enumerate(doc):
        for w in page.widgets() or []:
            fields.append({
                "page": pno + 1,
                "name": w.field_name or f"field_{len(fields)+1}",
                "type": w.field_type_string,
                "value": w.field_value,
                "options": getattr(w, "choice_values", None),
                "rect": list(w.rect),
            })
    doc.close()
    return fields


def form_fill(path, out_path, values, flatten=False):
    doc = fitz.open(path)
    for page in doc:
        for w in page.widgets() or []:
            if w.field_name in values:
                v = values[w.field_name]
                if w.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    w.field_value = v in (True, "true", "on", "yes", "1", "True")
                else:
                    w.field_value = str(v)
                w.update()
    if flatten:
        data = doc.convert_to_pdf()  # bakes appearances, drops interactivity
        doc.close()
        doc = fitz.open("pdf", data)
    save_small(doc, out_path)


def form_create(path, out_path, fields):
    """fields: [{page:1-based, rect:[normalized], name, type:'text'|'checkbox', label}]"""
    doc = fitz.open(path)
    for f in fields:
        page = normalize_page(doc[int(f["page"]) - 1])
        rect = denorm_rect(page, f["rect"])
        w = fitz.Widget()
        w.rect = rect
        w.field_name = f.get("name") or "field"
        w.field_type = _WIDGET_TYPES.get(f.get("type", "text"),
                                         fitz.PDF_WIDGET_TYPE_TEXT)
        w.border_color = (0.6, 0.6, 0.6)
        w.fill_color = (0.97, 0.97, 1.0)
        w.text_fontsize = 11
        if w.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            w.field_value = False
        page.add_widget(w)
    save_small(doc, out_path)


# ---------------------------------------------------------------- editor

_FONT_MAP = {"he": "helv", "ti": "tiro", "co": "cour"}


def _pick_font(bold=False, italic=False, serif=False, mono=False):
    base = "co" if mono else ("ti" if serif else "he")
    suffix = ("bi" if bold and italic else
              "bo" if bold else
              "it" if italic else "")
    return base + suffix if suffix else _FONT_MAP[base]


def _hexrgb(c):
    return tuple(int(c.lstrip("#")[i:i+2], 16) / 255 for i in (0, 2, 4))


def text_spans(path, pno):
    """All text spans of one page with normalized bboxes + style info,
    so the UI can offer click-to-edit."""
    doc = fitz.open(path)
    page = doc[pno]
    pr = page.rect
    spans = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for s in line["spans"]:
                if not s["text"].strip():
                    continue
                x0, y0, x1, y1 = s["bbox"]
                spans.append({
                    "rect": [(x0 - pr.x0) / pr.width, (y0 - pr.y0) / pr.height,
                             (x1 - pr.x0) / pr.width, (y1 - pr.y0) / pr.height],
                    "text": s["text"],
                    "size": round(s["size"], 1),
                    "color": "#%06x" % (s["color"] & 0xFFFFFF),
                    "bold": bool(s["flags"] & 16),
                    "italic": bool(s["flags"] & 2),
                    "serif": bool(s["flags"] & 4),
                    "mono": bool(s["flags"] & 8),
                })
    info = {"w": pr.width, "h": pr.height, "spans": spans}
    doc.close()
    return info


def _insert_fitted_text(page, rect, text, fontname, size, color):
    """insert_textbox that shrinks the font until the text fits the rect."""
    box = fitz.Rect(rect)
    box.x1 = min(box.x1 + box.width * 0.15 + 4, page.rect.x1)   # slack for metric drift
    box.y1 = min(box.y1 + size * 0.6, page.rect.y1)
    s = size
    while s > 4:
        leftover = page.insert_textbox(box, text, fontname=fontname, fontsize=s,
                                       color=color, align=fitz.TEXT_ALIGN_LEFT)
        if leftover >= 0:
            return
        s *= 0.9
    page.insert_textbox(box, text, fontname=fontname, fontsize=4, color=color)


def edit_pdf(path, out_path, ops, image_paths=None):
    """Replay editor operations. Each op: {type, page(1-based), rect|p1/p2|points
    (normalized), ...}. Types: replace_text, add_text, highlight, underline,
    strikeout, whiteout, rect, ellipse, line, ink, image."""
    doc = fitz.open(path)
    image_paths = image_paths or {}
    by_page = {}
    for op in ops:
        by_page.setdefault(int(op["page"]) - 1, []).append(op)

    for pno, plist in by_page.items():
        page = normalize_page(doc[pno])
        # phase 1: true removals (replaced text, whiteout) in one redaction pass
        removals = [o for o in plist if o["type"] in ("replace_text", "whiteout")]
        for o in removals:
            r = denorm_rect(page, o["rect"])
            fill = (1, 1, 1) if o["type"] == "whiteout" else False
            page.add_redact_annot(r, fill=fill)
        if removals:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)

        # phase 2: additions, in the order the user made them
        for o in plist:
            t = o["type"]
            if t == "whiteout":
                continue
            rect = denorm_rect(page, o["rect"]) if "rect" in o else None
            color = _hexrgb(o.get("color", "#111111"))
            width = float(o.get("width", 2))
            if t in ("replace_text", "add_text"):
                text = o.get("text", "")
                if not text.strip():
                    continue
                font = _pick_font(o.get("bold"), o.get("italic"),
                                  o.get("serif"), o.get("mono"))
                _insert_fitted_text(page, rect, text, font,
                                    float(o.get("size", 12)), color)
            elif t in ("highlight", "underline", "strikeout"):
                annot = {"highlight": page.add_highlight_annot,
                         "underline": page.add_underline_annot,
                         "strikeout": page.add_strikeout_annot}[t](rect)
                if o.get("color"):
                    annot.set_colors(stroke=color)
                annot.update()
            elif t == "rect":
                page.draw_rect(rect, color=color, width=width,
                               fill=_hexrgb(o["fill"]) if o.get("fill") else None)
            elif t == "ellipse":
                page.draw_oval(rect, color=color, width=width,
                               fill=_hexrgb(o["fill"]) if o.get("fill") else None)
            elif t == "line":
                r = page.rect
                p1 = fitz.Point(r.x0 + o["p1"][0] * r.width, r.y0 + o["p1"][1] * r.height)
                p2 = fitz.Point(r.x0 + o["p2"][0] * r.width, r.y0 + o["p2"][1] * r.height)
                page.draw_line(p1, p2, color=color, width=width)
            elif t == "ink":
                r = page.rect
                pts = [fitz.Point(r.x0 + x * r.width, r.y0 + y * r.height)
                       for x, y in o["points"]]
                if len(pts) > 1:
                    page.draw_polyline(pts, color=color, width=width,
                                       lineCap=1, lineJoin=1)
            elif t == "image":
                img = image_paths.get(o.get("image_id"))
                if not img:
                    raise ValueError("Image for an 'insert image' step is missing.")
                page.insert_image(rect, filename=img, keep_proportion=True,
                                  overlay=True)
    save_small(doc, out_path)
    return len(ops)


# ---------------------------------------------------------------- markdown

def pdf_to_markdown(path, out_path):
    import pymupdf4llm
    md = pymupdf4llm.to_markdown(path)
    Path(out_path).write_text(md, encoding="utf-8")


# ---------------------------------------------------------------- misc

def zip_files(paths, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    return zip_path


# ---------------------------------------------------------------- security

def protect(path, out_path, user_pw="", owner_pw="", allow_print=True,
            allow_copy=True, allow_edit=True, allow_annotate=True):
    """Encrypt with AES-256. user_pw opens the file; owner_pw lifts the
    permission restrictions (defaults to user_pw when not given)."""
    if not user_pw and not owner_pw:
        raise ValueError("Enter a password.")
    owner_pw = owner_pw or user_pw
    perm = fitz.PDF_PERM_ACCESSIBILITY
    if allow_print:
        perm |= fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
    if allow_copy:
        perm |= fitz.PDF_PERM_COPY
    if allow_edit:
        perm |= fitz.PDF_PERM_MODIFY | fitz.PDF_PERM_ASSEMBLE
    if allow_annotate:
        perm |= fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM
    doc = fitz.open(path)
    if doc.needs_pass:
        raise ValueError("This PDF is already password protected - unlock it first.")
    doc.save(out_path, garbage=3, deflate=True,
             encryption=fitz.PDF_ENCRYPT_AES_256,
             owner_pw=owner_pw, user_pw=user_pw, permissions=perm)
    doc.close()


def unlock(path, out_path, password=""):
    """Remove the password / permission restrictions (password required)."""
    doc = fitz.open(path)
    if not doc.is_encrypted:
        doc.close()
        raise ValueError("This PDF is not password protected.")
    if doc.needs_pass or password:
        # authenticate() returns 0 on failure, 2/4/6 for user/owner access
        if not doc.authenticate(password or ""):
            doc.close()
            raise ValueError("Wrong password.")
    doc.save(out_path, garbage=3, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
    doc.close()


# ---------------------------------------------------------------- metadata

META_KEYS = ("title", "author", "subject", "keywords", "creator", "producer")


def get_metadata(path):
    with fitz.open(path) as doc:
        m = doc.metadata or {}
        info = {k: m.get(k) or "" for k in META_KEYS}
        info["_pages"] = doc.page_count
        info["_format"] = m.get("format") or ""
        info["_created"] = m.get("creationDate") or ""
        info["_modified"] = m.get("modDate") or ""
        info["_encryption"] = m.get("encryption") or ""
        return info


def set_metadata(path, out_path, values, strip=False):
    """Write document info. strip=True removes all metadata (incl. XMP)."""
    doc = fitz.open(path)
    if strip:
        doc.set_metadata({})
        doc.del_xml_metadata()
    else:
        m = dict(doc.metadata or {})
        for k in META_KEYS:
            if k in values:
                m[k] = str(values[k] or "").strip()
        m["modDate"] = fitz.get_pdf_now()
        doc.set_metadata(m)
    save_small(doc, out_path)


def extract_images(path, out_dir, page_spec="", min_px=32):
    """Save every embedded image (original encoding where possible).
    Images reused on several pages are written once."""
    doc = fitz.open(path)
    stem = Path(path).stem
    seen, outputs = set(), []
    for pno in pages_from_spec(page_spec, doc.page_count):
        for n, info in enumerate(doc[pno].get_images(full=True), 1):
            xref = info[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                img = doc.extract_image(xref)
            except Exception:
                continue
            if not img or min(img.get("width", 0), img.get("height", 0)) < min_px:
                continue
            data, ext = img["image"], img.get("ext", "png")
            if info[1]:  # has an soft mask (transparency) - composite to PNG
                try:
                    pix = fitz.Pixmap(doc, xref)
                    mask = fitz.Pixmap(doc, info[1])
                    pix = fitz.Pixmap(pix, mask)
                    if pix.colorspace and pix.colorspace.n > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    data, ext = pix.tobytes("png"), "png"
                except Exception:
                    pass
            if ext not in ("png", "jpg", "jpeg", "jpx", "bmp", "tiff", "gif"):
                pix = fitz.Pixmap(doc, xref)
                if pix.colorspace and pix.colorspace.n > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                data, ext = pix.tobytes("png"), "png"
            f = os.path.join(out_dir, f"{stem}_p{pno+1}_img{n}.{ext}")
            Path(f).write_bytes(data)
            outputs.append(f)
    doc.close()
    if not outputs:
        raise ValueError("No embedded images found on the selected pages.")
    return outputs


# ---------------------------------------------------------------- OCR

def find_tessdata():
    """Locate Tesseract's tessdata folder, or None when Tesseract is missing."""
    cands = [os.environ.get("TESSDATA_PREFIX")]
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                 os.environ.get("LOCALAPPDATA", "")):
        if base:
            cands.append(os.path.join(base, "Tesseract-OCR", "tessdata"))
            cands.append(os.path.join(base, "Programs", "Tesseract-OCR", "tessdata"))
    try:
        cands.append(fitz.get_tessdata())
    except Exception:
        pass
    cands += ["/usr/share/tesseract-ocr/5/tessdata", "/usr/share/tesseract-ocr/4.00/tessdata",
              "/usr/share/tessdata", "/opt/homebrew/share/tessdata", "/usr/local/share/tessdata"]
    for c in cands:
        if c and os.path.isfile(os.path.join(c, "eng.traineddata")):
            return c
        if c and os.path.isdir(c) and any(f.endswith(".traineddata") for f in os.listdir(c)):
            return c
    return None


def ocr_languages():
    td = find_tessdata()
    if not td:
        return None
    return sorted(f[:-12] for f in os.listdir(td)
                  if f.endswith(".traineddata") and f[:-12] != "osd")


def ocr_pdf(path, out_path, language="eng", dpi=300, skip_text=True, page_spec=""):
    """Add an invisible, searchable text layer on top of scanned pages.
    The visible page content is left untouched (no re-rasterizing).
    Returns the number of pages OCR'd."""
    td = find_tessdata()
    if not td:
        raise RuntimeError(
            "Tesseract OCR is not installed. Install it (Windows installer: "
            "github.com/UB-Mannheim/tesseract/wiki) and restart PDF4life.")
    doc = fitz.open(path)
    done = 0
    for pno in pages_from_spec(page_spec, doc.page_count):
        page = doc[pno]
        if skip_text and page.get_text("text").strip():
            continue
        tp = page.get_textpage_ocr(language=language or "eng", dpi=int(dpi),
                                   full=True, tessdata=td)
        for x0, y0, x1, y1, word, *_ in page.get_text("words", textpage=tp):
            w, h = x1 - x0, y1 - y0
            if w <= 0 or h <= 0 or not word.strip():
                continue
            size = min(h * 0.95, w / max(fitz.get_text_length(word, "helv", 1), 1e-3))
            try:
                page.insert_text((x0, y1 - h * 0.2), word, fontname="helv",
                                 fontsize=max(size, 1), render_mode=3)
            except Exception:
                pass
        done += 1
    if done == 0:
        doc.close()
        raise ValueError("Every selected page already has text - nothing to OCR. "
                         "Untick 'Skip pages that already have text' to force it.")
    save_small(doc, out_path)
    return done
