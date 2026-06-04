"""app.py — PDF to Excel API.
Local  : python app.py  →  http://localhost:5000
Vercel : git push → auto-deploy
"""
import os, traceback, tempfile
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

from flask import Flask, render_template, jsonify, request, Response
from werkzeug.utils import secure_filename

import extractor
import excel_writer
import supabase_client as supa

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# ── CORS — allow any origin (needed for public API) ───────────────────────────
@app.after_request
def add_cors(resp):
    resp.headers["Access-Control-Allow-Origin"]  = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type,Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    return resp

@app.route("/api/convert", methods=["OPTIONS"])
def convert_preflight():
    return "", 204

# ── Global error handler — always return JSON ─────────────────────────────────
@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/debug-env")
def debug_env():
    """Temporary: confirm env vars reach Vercel runtime."""
    import os
    groq  = os.environ.get("GROQ_API_KEY", "")
    supa  = os.environ.get("SUPABASE_URL", "")
    return jsonify({
        "GROQ_API_KEY":    (groq[:8] + "..." if groq else "NOT SET"),
        "GROQ_len":        len(groq),
        "SUPABASE_URL":    (supa[:30] + "..." if supa else "NOT SET"),
        "active_backend":  extractor.active_backend(),
    })


@app.route("/api/config")
def api_config():
    return jsonify({
        "ai_backend":     extractor.active_backend(),
        "api_key_set":    extractor.active_backend() != "none",
        "supabase_ready": supa.is_configured(),
    })


@app.route("/api/convert", methods=["POST"])
def api_convert():
    """
    Upload a PDF and convert to Excel — synchronous, returns result directly.

    Form fields:
      file : PDF file (multipart)
      type : "INV" or "PKL"

    Returns JSON:
      { filename, folder, rows, xlsx_name, download_url }
    """
    folder = request.form.get("type", "").upper()
    if folder not in ("INV", "PKL"):
        return jsonify({"error": "type must be INV or PKL"}), 400

    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file provided"}), 400
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are accepted"}), 400

    fname = secure_filename(f.filename)

    # Use temp dir — works on Vercel (no persistent local disk)
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path  = Path(tmp) / fname
        xlsx_name = Path(fname).stem + ".xlsx"
        xlsx_path = Path(tmp) / xlsx_name

        f.save(pdf_path)

        # Upload source PDF to Supabase
        conv_id = None
        if supa.is_configured():
            pdf_storage = f"{folder}/{fname}"
            supa.upload_file(supa.PDF_BUCKET, pdf_storage, pdf_path,
                             content_type="application/pdf")
            conv_id = supa.insert_conversion(fname, folder, pdf_storage)

        # Extract with AI
        try:
            items = extractor.extract(str(pdf_path), folder)
        except Exception as e:
            if conv_id:
                supa.update_conversion(conv_id, status="error", message=str(e))
            return jsonify({"error": f"Extraction failed: {e}"}), 500

        if not items:
            if conv_id:
                supa.update_conversion(conv_id, status="warn",
                                       message="No items extracted")
            return jsonify({"error": "No items extracted — check the PDF"}), 422

        # Write Excel
        excel_writer.write_excel(items, xlsx_path)

        # Upload Excel to Supabase
        download_url = f"/api/download/{folder}/{xlsx_name}"
        if supa.is_configured():
            xlsx_storage = f"{folder}/{xlsx_name}"
            supa.upload_file(
                supa.XLSX_BUCKET, xlsx_storage, xlsx_path,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            supa.update_conversion(conv_id,
                status="done",
                message=f"{len(items)} rows",
                rows_extracted=len(items),
                xlsx_storage_path=xlsx_storage,
            )

        return jsonify({
            "filename":     fname,
            "folder":       folder,
            "rows":         len(items),
            "xlsx_name":    xlsx_name,
            "download_url": download_url,
        })


@app.route("/api/download/<folder>/<filename>")
def api_download(folder, filename):
    if folder not in ("INV", "PKL"):
        return jsonify({"error": "Unknown folder"}), 400
    if not filename.endswith(".xlsx"):
        return jsonify({"error": "Invalid file"}), 400

    if not supa.is_configured():
        return jsonify({"error": "Supabase not configured"}), 503

    data = supa.download_file(supa.XLSX_BUCKET, f"{folder}/{filename}")
    return Response(
        data,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
    )


@app.route("/api/history")
def api_history():
    if not supa.is_configured():
        return jsonify([])
    return jsonify(supa.list_conversions(limit=20))


# ── Local server ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port    = int(os.environ.get("PORT", 5000))
    backend = extractor.active_backend()

    print("\n── DP World PDF → Excel ──────────────────────────────")
    print(f"  AI backend : {'✓ ' + backend.upper() if backend != 'none' else '✗ NOT SET — add GEMINI_API_KEY to .env'}")
    print(f"  Supabase   : {'✓ ready' if supa.is_configured() else '✗ not configured'}")
    print(f"  Open       : http://localhost:{port}")
    print("──────────────────────────────────────────────────────\n")

    from waitress import serve
    serve(app, host="0.0.0.0", port=port, threads=8)
