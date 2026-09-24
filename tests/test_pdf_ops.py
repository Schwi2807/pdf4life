"""Smoke tests for pdf_ops and the Flask API (run: python -m pytest -q)."""
import io
import sys
from pathlib import Path

import pymupdf
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import app as appmod  # noqa: E402
import pdf_ops  # noqa: E402


@pytest.fixture
def sample(tmp_path):
    doc = pymupdf.open()
    for i in range(3):
        page = doc.new_page()
        page.insert_text((72, 72), f"Hello page {i + 1}", fontsize=24)
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 64, 64), False)
    pix.set_rect(pix.irect, (200, 30, 30))
    doc[0].insert_image(pymupdf.Rect(100, 100, 200, 200), pixmap=pix)
    p = tmp_path / "sample.pdf"
    doc.save(p)
    return p


@pytest.fixture
def scanned(tmp_path):
    """An image-only page containing rendered text (simulates a scan)."""
    src = pymupdf.open()
    src.new_page().insert_text((72, 120), "Searchable scan", fontsize=36)
    pix = src[0].get_pixmap(dpi=150)
    doc = pymupdf.open()
    page = doc.new_page(width=src[0].rect.width, height=src[0].rect.height)
    page.insert_image(page.rect, pixmap=pix)
    p = tmp_path / "scan.pdf"
    doc.save(p)
    return p


def test_merge_split(sample, tmp_path):
    out = tmp_path / "m.pdf"
    pdf_ops.merge([sample, sample], out)
    assert pymupdf.open(out).page_count == 6
    files = pdf_ops.split(str(out), "every", 2, str(tmp_path))
    assert len(files) == 3


def test_protect_unlock(sample, tmp_path):
    enc, dec = tmp_path / "enc.pdf", tmp_path / "dec.pdf"
    pdf_ops.protect(str(sample), str(enc), user_pw="secret", allow_copy=False)
    d = pymupdf.open(enc)
    assert d.needs_pass
    assert d.authenticate("secret")
    with pytest.raises(ValueError, match="Wrong password"):
        pdf_ops.unlock(str(enc), str(dec), "nope")
    pdf_ops.unlock(str(enc), str(dec), "secret")
    d = pymupdf.open(dec)
    assert not d.is_encrypted and "Hello page 1" in d[0].get_text()


def test_metadata(sample, tmp_path):
    out = tmp_path / "meta.pdf"
    pdf_ops.set_metadata(str(sample), str(out), {"title": "T", "author": "A"})
    m = pdf_ops.get_metadata(str(out))
    assert m["title"] == "T" and m["author"] == "A" and m["_pages"] == 3
    out2 = tmp_path / "strip.pdf"
    pdf_ops.set_metadata(str(out), str(out2), {}, strip=True)
    assert pdf_ops.get_metadata(str(out2))["title"] == ""


def test_extract_images(sample, tmp_path):
    files = pdf_ops.extract_images(str(sample), str(tmp_path))
    assert len(files) == 1 and Path(files[0]).stat().st_size > 0
    with pytest.raises(ValueError):
        pdf_ops.extract_images(str(sample), str(tmp_path), "2-3")


@pytest.mark.skipif(pdf_ops.find_tessdata() is None, reason="Tesseract not installed")
def test_ocr(scanned, tmp_path):
    assert pymupdf.open(scanned)[0].get_text().strip() == ""
    out = tmp_path / "ocr.pdf"
    assert pdf_ops.ocr_pdf(str(scanned), str(out), dpi=200) == 1
    text = pymupdf.open(out)[0].get_text().lower()
    assert "searchable" in text


def test_api_roundtrip(sample):
    c = appmod.app.test_client()
    r = c.post("/api/upload", data={"kind": "pdf",
               "files": (io.BytesIO(sample.read_bytes()), "doc.pdf")})
    meta = r.get_json()[0]
    assert meta["pages"] == 3
    r = c.post("/api/protect", json={"id": meta["id"], "user_pw": "pw"})
    assert r.status_code == 200, r.get_json()
    blob = c.get(r.get_json()["download"]).data
    # re-upload the encrypted result and unlock it through the API
    r = c.post("/api/upload", data={"kind": "pdf",
               "files": (io.BytesIO(blob), "doc_protected.pdf")})
    enc = r.get_json()[0]
    assert enc["encrypted"]
    assert c.post("/api/unlock", json={"id": enc["id"], "password": "x"}).status_code == 400
    assert c.post("/api/unlock", json={"id": enc["id"], "password": "pw"}).status_code == 200
    r = c.post("/api/metadata/get", json={"id": meta["id"]})
    assert r.get_json()["_pages"] == 3
    assert c.get("/api/ocr/status").get_json()["available"] in (True, False)
    assert c.get("/").status_code == 200
