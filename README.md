# Electoral Roll OCR Extractor

Extracts voter data from image-based Tamil Nadu electoral roll PDFs and exports to Excel (`.xlsx`) with:

- One main sheet with all extracted records
- Additional sheets split by section
- OCR retry logic when extracted count significantly differs from summary count

## Project Files

| File | Purpose |
|---|---|
| `electoral_roll.py` | Main OCR extraction script |
| `server.py` | Flask backend — runs the script and serves results to the UI |
| `index.html` | Browser UI for batch-processing multiple PDFs |
| `requirements.txt` | Python dependencies |

## Prerequisites

### 1. Python 3.10+

### 2. Tesseract OCR

Install Tesseract and ensure the executable exists at:

```
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If your install path differs, update this line in `electoral_roll.py`:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### 3. Poppler (required by pdf2image)

Install Poppler binaries. `electoral_roll.py` tries a known WinGet install path; if that doesn't exist on your machine, either:

- add Poppler `bin` to system PATH, or
- update the `_poppler` path inside `electoral_roll.py`

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

### Option A — Browser UI (recommended for batch work)

1. Start the backend server:

```powershell
python server.py
```

2. Open `index.html` in your browser.

3. Drop one or more PDF files onto the page, adjust options (start page, DPI, output folder), and click **Run Queue**.

4. When each file completes, a **⬇ Download xlsx** link appears in the log panel. Output files are also saved to the `outputs/` folder next to `server.py`.

### Option B — CLI (single file)

```powershell
python electoral_roll.py "path\to\roll.pdf"
```

Useful options:

```powershell
# Specify page range and DPI
python electoral_roll.py "roll.pdf" --start-page 3 --end-page 250 --dpi 300

# Custom output path
python electoral_roll.py "roll.pdf" --output "results\roll.xlsx"

# Name the Excel sheet after the PDF filename (useful when processing many files)
python electoral_roll.py "roll.pdf" --batch

# Skip automatic OCR retry passes
python electoral_roll.py "roll.pdf" --no-retry
```

## Output

- **Excel** (`.xlsx`) — one sheet per section plus an "All Records" sheet
- Default save location: same folder as the input PDF (CLI) or `outputs/` (UI)

## Troubleshooting

| Error | Fix |
|---|---|
| `TesseractNotFoundError` | Install Tesseract; verify path in `electoral_roll.py` |
| `PDFInfoNotInstalledError` / Poppler errors | Install Poppler; add `bin` to PATH |
| Empty output | Check `--start-page` / `--end-page`; verify PDF is image-based |
| `Cannot reach backend` (UI) | Run `python server.py` before opening `index.html` |
