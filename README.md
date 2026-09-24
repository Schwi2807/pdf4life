# PDF4life

A standalone, offline PDF toolbox for Windows. Everything runs locally on your
laptop — no files are ever uploaded to the internet.

Double-click **`dist\PDF4life.exe`** — the app opens in its own window
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
- Scanned (image-only) PDFs convert to Word/Markdown as images/empty text —
  there is no OCR built in.
- Uploaded files and results are kept in `%TEMP%\PDF4life` and cleaned up
  automatically after 24 hours.
- The app binds to `127.0.0.1` only — it is not reachable from other devices
  unless you opt in with `--lan` (see above).

## Running from source

```powershell
.venv\Scripts\python.exe app.py        # or: run.bat
```

## Rebuilding the exe

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
