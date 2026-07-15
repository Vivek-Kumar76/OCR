# OCR to Structured Word Document

A FastAPI app that lets a user upload an image or PDF (including scanned/image-only PDFs), extracts the text and tables from it using **PaddleOCR** + **PP-StructureV3**, and lets them download the result as a properly structured **Word (.docx)** document — paragraphs stay as paragraphs, tables become real Word tables.

## Features

- Upload images (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tiff`) or PDFs (including scanned/image-only PDFs)
- Multi-page PDF support
- Preserves document structure:
  - Detected **tables** become real Word tables (rows/columns), not flattened text
  - Wrapped prose lines are merged back into single flowing paragraphs
  - Side-by-side text (e.g. label + value pairs like `Bank: Really Great Bank`) is kept on one line using a tab stop
  - Reading order follows each block's actual position on the page
- Simple drag-and-drop web UI
- Nothing is saved to disk — extraction and Word document generation both happen in memory
- Download button only becomes available after a successful extraction

## How it works

1. **Upload** — the user selects/drops a file and clicks **Extract Text**.
2. **`POST /ocr`** — the file is OCR'd (`PaddleOCR`) and analyzed for layout/tables (`PPStructureV3`). The result is combined into a structured, page-by-page representation and:
   - A text preview is returned and shown in the browser
   - The structured result is kept in memory (not saved to disk) so it can be reused for download
3. **Download** — once extraction succeeds, the **Download as Word** button becomes clickable.
4. **`GET /download`** — builds a `.docx` file **in memory** from the already-extracted structured result and streams it back to the browser as a file download.

## Project structure

```
your_project/
├── main.py              # FastAPI app
├── requirements.txt
└── static/
    └── index.html       # Upload UI (drag & drop, extract, download)
```

## Setup

Requires Python 3.9–3.11.

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

PaddleOCR/PP-StructureV3 will download their model weights automatically on first run (cached afterward).

### Extra dependency for PP-StructureV3

PP-StructureV3 needs additional PaddleX extras beyond the base install:

```bash
pip install "paddlex[ocr]>=3.7.0,<3.8.0"
```

(Match the version range to whatever your installed `paddleocr` requires — check with `pip show paddleocr`.)

## Running

```bash
python main.py
```

Then open **http://localhost:8000** in your browser.

## API reference

### `POST /ocr`
Upload a file, run OCR + structure extraction.

**Request:** `multipart/form-data` with a `file` field.

**Response:**
```json
{
  "filename": "invoice.png",
  "num_pages": 1,
  "text": "INVOICE\n#1024\n..."
}
```

### `GET /download`
Builds and returns a `.docx` from the most recently extracted result.

**Response:** streamed `.docx` file (`Content-Disposition: attachment`).

Returns `404` if called before any successful `/ocr` call.

## Known limitations

- **Single-user design:** the most recent extraction is held in a single in-memory variable. If two people use the same running instance at the same time, one's extraction can overwrite the other's before it's downloaded. Fine for local/personal use; would need per-session or per-request storage for multi-user deployments.
- **No persistence:** since nothing is saved to disk, restarting the server clears any pending (not-yet-downloaded) extraction — the user would need to re-upload and re-extract.
- **Table detection accuracy depends on table style:** tables with clear borders/gridlines are detected reliably. Borderless or thin-line tables (common in some invoice/letter templates) may be misclassified as plain text, or have rows/columns misaligned, since this depends on PP-StructureV3's table-structure model rather than something this app controls directly.
- **Row-grouping heuristic:** side-by-side text is detected using a vertical-position threshold (`y_threshold=12` in `group_into_rows`). Documents with unusually large or small line heights may need this value tuned.
- **Single language recognition per run:** the OCR engine is configured for one language (`lang="en"` by default in `main.py`). For documents in other languages, change the `lang` parameter — some language models (e.g. the Devanagari/Hindi model) also read Latin-script text reasonably well, so a single model can sometimes cover mixed-language documents; this varies by language pair and should be tested on real samples.

## Customizing

- **Language:** change `lang="en"` in both the `PaddleOCR(...)` and `PPStructureV3(...)` calls in `main.py`.
- **Label/value column alignment:** the tab stop position is set to `Inches(2.2)` in `build_docx()` — adjust to fit your documents' typical label-column width.
- **Allowed file types:** edit the `ALLOWED_EXTENSIONS` set in `main.py`.
