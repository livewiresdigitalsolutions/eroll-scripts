"""
Flask backend for the Electoral Roll Extractor UI.

Run with:
    pip install flask flask-cors
    python server.py

The UI (index.html) sends PDFs here; this server runs electoral_roll.py
as a subprocess and returns JSON with the record count + xlsx filename.
The resulting xlsx is served for download via GET /download/<filename>.
"""

import os
import re
import subprocess
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # needed because index.html is opened as a local file (file://)

# Outputs land here when the user hasn't specified their own output folder.
OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

SCRIPT = Path(__file__).parent / "electoral_roll.py"


def _parse_record_count(log_text: str) -> int:
    """Pull the final extracted record count out of the script's log output."""
    m = re.search(r"Final totals:\s*(\d+)\s*records", log_text)
    if m:
        return int(m.group(1))
    # Fallback: last "→ N records extracted" line
    matches = re.findall(r"→\s*(\d+)\s*records\s*extracted", log_text)
    if matches:
        return int(matches[-1])
    return 0


@app.route("/health")
def health():
    return jsonify({"ok": True})


@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided"}), 400

    pdf_file = request.files["file"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "Only PDF files are accepted"}), 400

    start_page = request.form.get("start_page", "3")
    end_page   = request.form.get("end_page", "").strip()
    dpi        = request.form.get("dpi", "200")
    no_retry   = request.form.get("no_retry", "false").lower() == "true"
    output_dir = request.form.get("output_dir", "").strip()

    # Determine output xlsx path
    stem = re.sub(r"[^\w\-]", "_", Path(pdf_file.filename).stem)
    if output_dir:
        out_path = Path(output_dir) / f"{stem}.xlsx"
    else:
        out_path = OUTPUT_DIR / f"{stem}.xlsx"

    # Save uploaded PDF to a temp file
    tmp_fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf", prefix=f"{stem}_")
    os.close(tmp_fd)
    try:
        pdf_file.save(tmp_pdf)

        cmd = [
            "python", str(SCRIPT),
            tmp_pdf,
            "--start-page", start_page,
            "--dpi", dpi,
            "--output", str(out_path),
            "--batch",
        ]
        if end_page:
            cmd += ["--end-page", end_page]
        if no_retry:
            cmd += ["--no-retry"]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=str(Path(__file__).parent),
        )

        log_text = result.stdout + result.stderr
        records  = _parse_record_count(log_text)
        ok       = result.returncode == 0 and out_path.exists()

        return jsonify({
            "ok":      ok,
            "records": records,
            "xlsx":    out_path.name if ok else None,
            "log":     log_text,
            "error":   None if ok else f"Script exited with code {result.returncode}",
        })

    finally:
        if os.path.exists(tmp_pdf):
            os.unlink(tmp_pdf)


@app.route("/download/<path:filename>")
def download(filename):
    """Serve the generated xlsx for download."""
    # Strip any directory traversal attempts
    safe_name = Path(filename).name
    path = OUTPUT_DIR / safe_name
    if not path.exists():
        return jsonify({"error": "File not found"}), 404
    return send_file(str(path), as_attachment=True)


if __name__ == "__main__":
    print("Electoral Roll backend listening on http://localhost:5000")
    print(f"Outputs folder: {OUTPUT_DIR}")
    app.run(port=5000, debug=False)
