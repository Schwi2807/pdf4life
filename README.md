# PDF4life

[![Build](https://github.com/schwi2807/pdf4life/actions/workflows/build.yml/badge.svg)](https://github.com/schwi2807/pdf4life/actions/workflows/build.yml)

A standalone, offline PDF toolbox for Windows. Everything runs locally on your
laptop — no files are ever uploaded to the internet.

**Download:** grab `PDF4life.exe` from the
[Releases](https://github.com/schwi2807/pdf4life/releases) page (or build it
yourself, see below), then double-click it — the app opens in its own window
(no browser, no console). Close the window to quit.

It uses Windows' built-in Edge WebView2 engine for the interface. If that is
ever unavailable, the app automatically falls back to opening in your default
browser instead. You can also force browser mode with
`PDF4life.exe --browser`.

**Process lifetime:** closing the app window terminates the process
immediately. In browser mode the exe has no window of its own, so it instead
quits itself automatically ~2 minutes after the last browser tab is closed
(the page sends a small keep-alive ping every 15 s while open; tune with the
`PDF4LIFE_IDLE` environment variable, in seconds).

## Sharing on your local network

By default the app binds to `127.0.0.1` and is unreachable from other devices.
To use it from your phone or another PC on the same Wi-Fi:

```powershell
PDF4life.exe --lan          # binds 0.0.0.0; also: --host <ip>, or set PDF4LIFE_LAN=1
```

In browser mode it prints the shareable `http://<your-ip>:<port>` address.
Windows Firewall will ask for permission the first time. Anyone on the network
can then use the app (and closing the app window still stops it for everyone),
so only use `--lan` on networks you trust.

## Tools

| Tool | What it does |
|---|---|
| Merge PDF | Combine several PDFs into one (drag to set order) |
| Split PDF | Split by custom ranges, every N pages, or extract pages |
| Compress PDF | Shrink file size (3 levels; image downsampling + font subsetting) |
| Edit PDF | Acrobat-style editing: click any text to retype it (matching size/style/color), add text boxes, highlight, whiteout (true erase), rectangles/ellipses/lines, freehand drawing, insert images — with live preview, undo, and click-to-remove |
| PDF → Word | Convert to an editable `.docx` (layout-aware, via pdf2docx) |
| Word → PDF | Convert `.docx` to PDF (uses your installed Microsoft Word, or LibreOffice) |
| PDF → JPG / PNG | Export pages as images at 100/150/300 DPI |
| JPG → PDF | Turn images into a PDF (drag to set page order) |
| Watermark | Text or image watermark; opacity, rotation, position, tiling |
| Sign PDF | Draw a signature with mouse/touch (or upload one) and place it on any pages |
| Organize | Thumbnail view — drag to reorder, rotate, delete pages |
| Page numbers | Several formats and positions, custom start number |
| Crop PDF | Draw the area to keep on a live preview |
| Redact | **True redaction** — draw boxes or search text; content underneath is deleted, not just covered |
| Compare PDFs | PDF report with pages side by side, differences boxed in red + text-similarity per page |
| PDF Forms | Detect existing fields and key in values (with optional flattening), or draw new text/checkbox fields to create a fillable PDF |
| PDF → Markdown | Structured text extraction to `.md` (headings, tables) |
| Protect PDF | Add a password with AES-256 encryption; choose whether printing, copying, editing and commenting are allowed |
| Unlock PDF | Remove the password / restrictions from a PDF you have the password for |
| OCR PDF | Make scanned PDFs searchable and copyable — adds an invisible text layer, the page looks unchanged (needs Tesseract, see below) |
| Metadata | View and edit title, author, subject, keywords…, or strip all metadata for privacy |
| Extract images | Save every embedded image in its original format (ZIP), with a size filter |

The home screen also has a **tool search** (press `/`), a **dark mode**
toggle (remembers your choice, follows Windows by default) and a **Recent
results** list so you can re-download anything produced in the last 24 hours.

## Notes & limits

- **Word → PDF** needs Microsoft Word (or LibreOffice) installed — the app
  drives it for a faithful conversion. Everything else is fully self-contained.
- **Sign PDF** stamps a signature image (like signing on paper). It is not a
  cryptographic digital signature.
- **Edit PDF** replaces text using the closest standard font (Helvetica /
  Times / Courier family, matching bold/italic) — documents with unusual
  embedded fonts will show a slight typeface change on edited lines, and
  unlike Acrobat, surrounding paragraphs do not reflow. If the replacement
  is much longer than the original, its font shrinks to fit the box.
- **OCR PDF** needs [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
  installed (the standard Windows installer is detected automatically; tick
  extra languages such as Malay in the installer if you need them, or set
  `TESSDATA_PREFIX`). Run a scanned PDF through OCR first to get real text
  out of PDF → Word / Markdown.
- Uploaded files and results are kept in `%TEMP%\PDF4life` and cleaned up
  automatically after 24 hours.
- The app binds to `127.0.0.1` only — it is not reachable from other devices
  unless you opt in with `--lan` (see above).

## Running from source

```powershell
.venv\Scripts\python.exe app.py        # or: run.bat
```

## Tests

```powershell
.venv\Scripts\python.exe -m pip install pytest
.venv\Scripts\python.exe -m pytest -q tests
```

## Rebuilding the exe

### Automatically (GitHub Actions)

Every push to `main` and every pull request runs the tests and builds
`PDF4life.exe` on a Windows runner (download it from the run's
**Artifacts**). Pushing a version tag also publishes a GitHub Release with the
exe attached:

```powershell
git tag v1.1.0
git push origin v1.1.0
```

### Locally

```powershell
python -m venv .venv
.venv\Scripts\pip.exe install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller.exe --noconfirm PDF4life.spec
```

or, equivalently, without the spec file:

```powershell
.venv\Scripts\pyinstaller.exe --noconfirm --onefile --noconsole --name PDF4life --icon icon.ico ^
  --add-data "templates;templates" --collect-data pymupdf ^
  --collect-submodules pdf2docx app.py
```

The result appears at `dist\PDF4life.exe`. You can copy that single file
anywhere (USB drive, another PC) — nothing else is needed.

## Project layout

- `app.py` — Flask server + startup (opens the browser, picks a free port)
- `pdf_ops.py` — all PDF operations (PyMuPDF, pdf2docx, pymupdf4llm, Pillow)
- `templates/index.html` — the whole UI (vanilla JS, no frameworks)
- `PDF4life.spec` — PyInstaller build recipe
- `tests/` — pytest smoke tests for the PDF operations and API
- `.github/workflows/build.yml` — CI: tests + Windows exe build + releases
