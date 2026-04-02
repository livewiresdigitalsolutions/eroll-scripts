# Electoral Roll OCR Extractor

Extracts voter data from image-based Tamil Nadu electoral roll PDFs and exports to Excel (`.xlsx`) with:

- One main sheet with all extracted records
- Additional sheets split by section
- OCR retry logic when extracted count significantly differs from summary count

## Project Files

- `electoral_roll.py` - main OCR extraction script
- `requirements.txt` - Python dependencies
- `index.html` - optional UI page in this repo

## Prerequisites

This project needs both Python packages and system tools.

### 1. Python

- Python 3.10+ recommended

### 2. Tesseract OCR (required)

Install Tesseract and ensure the executable exists at:

`C:\Program Files\Tesseract-OCR\tesseract.exe`

If your install path is different, update this line in `electoral_roll.py`:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### 3. Poppler (required by pdf2image)

Install Poppler binaries and make sure they are available in PATH.

Note: `electoral_roll.py` currently tries a user-specific Poppler folder path. If that path does not exist on another machine, install Poppler and either:

- add Poppler `bin` to system PATH, or
- update the `_poppler` path in `electoral_roll.py`

## Setup

From the project folder:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

Basic run:

```powershell
python electoral_roll.py "path\to\your\roll.pdf"
```

Useful options:

```powershell
python electoral_roll.py "path\to\your\roll.pdf" --start-page 3 --end-page 250 --dpi 300
python electoral_roll.py "path\to\your\roll.pdf" --output "out.xlsx"
python electoral_roll.py "path\to\your\roll.pdf" --batch
python electoral_roll.py "path\to\your\roll.pdf" --no-retry
```

## Output

By default, output Excel file is created next to the input PDF with `.xlsx` extension.

## Install Dependencies Only

If your friends already have a virtual environment active:

```powershell
pip install -r requirements.txt
```

## Troubleshooting

- `TesseractNotFoundError`: install Tesseract and verify path in script.
- `PDFInfoNotInstalledError` / Poppler errors: install Poppler and add `bin` to PATH.
- Empty output: verify page range (`--start-page`, `--end-page`) and PDF quality.
