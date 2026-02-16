"""
Formater — drag & drop a DOCX, get it back formatted to MLA 9th edition.
Uses AI (OpenAI) to detect essay structure and check compliance.

Usage:
    python3 app.py
    -> opens http://localhost:8080
"""

import os, uuid, tempfile, threading, time, json
from flask import Flask, request, send_file, jsonify, render_template
from dotenv import load_dotenv

# Load .env from script directory
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# (#5) Single import — app.py just calls format_mla, no duplicate logic
from mla_formatter import format_mla

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "mla_formatter")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# (#6) Track files for cleanup: {file_id: {"paths": [...], "created": timestamp}}
file_registry = {}
CLEANUP_AGE = 3600  # 1 hour


def cleanup_old_files():
    """(#6) Remove temp files older than CLEANUP_AGE seconds."""
    now = time.time()
    expired = [fid for fid, info in file_registry.items() if now - info["created"] > CLEANUP_AGE]
    for fid in expired:
        for path in file_registry[fid]["paths"]:
            try:
                os.remove(path)
            except OSError:
                pass
        del file_registry[fid]
    # Schedule next cleanup
    threading.Timer(300, cleanup_old_files).start()


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    resp = app.make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/format", methods=["POST"])
def format_file():
    if "file" not in request.files:
        return jsonify(error="No file uploaded"), 400

    file = request.files["file"]
    if not file.filename.endswith(".docx"):
        return jsonify(error="Only .docx files are supported"), 400

    # (#11) Preserve original filename for download
    orig_name = os.path.splitext(file.filename)[0]

    # Save uploaded file
    file_id = uuid.uuid4().hex[:12]
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}_input.docx")
    output_path = os.path.join(UPLOAD_DIR, f"{file_id}_mla.docx")
    file.save(input_path)

    # Get optional fields
    name = request.form.get("name", "")
    instructor = request.form.get("instructor", "")
    course = request.form.get("course", "")
    date_str = request.form.get("date", "")
    # (#9) No heading checkbox
    no_heading = request.form.get("no_heading", "") == "true"

    # (#8) Custom heading field order from drag-to-reorder
    heading_order_raw = request.form.get("heading_order", "")
    heading_order = None
    if heading_order_raw:
        try:
            heading_order = json.loads(heading_order_raw)
        except Exception:
            pass

    api_key = os.environ.get("OPENAI_API_KEY", "")
    use_ai = bool(api_key)

    try:
        # (#5) Call format_mla directly — no duplicate logic
        result = format_mla(
            input_path, output_path,
            name=name, instructor=instructor, course=course, date=date_str,
            use_ai=use_ai, api_key=api_key, no_heading=no_heading,
            heading_order=heading_order,
        )
        result["file_id"] = file_id
        # (#11) Pass download name back to client
        result["download_name"] = f"{orig_name}_mla.docx"
    except Exception as e:
        return jsonify(error=f"Formatting failed: {str(e)}"), 500

    # (#6) Register files for cleanup
    file_registry[file_id] = {"paths": [input_path, output_path], "created": time.time()}

    return jsonify(result)


@app.route("/download/<file_id>")
def download(file_id):
    # Sanitize file_id to prevent path traversal
    if not file_id.isalnum():
        return "Invalid ID", 400
    path = os.path.join(UPLOAD_DIR, f"{file_id}_mla.docx")
    if not os.path.exists(path):
        return "File not found", 404
    # (#11) Use original filename from query param, fallback to generic
    dl_name = request.args.get("name", "essay_mla.docx")
    return send_file(path, as_attachment=True, download_name=dl_name)


# ── Run ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import webbrowser
    port = 8080
    print(f"\n  Formater running at: http://localhost:{port}\n")
    print(f"  Template auto-reload: {app.config['TEMPLATES_AUTO_RELOAD']}")
    print("  Cache mode: no-store for HTML templates\n")
    # (#6) Start cleanup timer
    threading.Timer(300, cleanup_old_files).start()
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(host="0.0.0.0", debug=True, use_reloader=True, port=port)
