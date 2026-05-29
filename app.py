#!/usr/bin/env python3
"""
BCAP Import Helper — web server
Run: python3 app.py
Then open: http://localhost:5050
"""

import os
import sys
import tempfile
import uuid

from flask import Flask, jsonify, render_template, request, send_file

sys.path.insert(0, os.path.dirname(__file__))
from bcap_import_helper import HeadingMismatchError, MissingTabsError, process

app = Flask(__name__)

# In-memory store: token -> (xlsx_path, csv_path, original_stem)
_outputs: dict = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process_file():
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file received."})

    f = request.files["file"]

    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"success": False, "error": "File must be an .xlsx file."})

    original_stem = os.path.splitext(f.filename)[0]

    # Save upload to a temp file
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    f.save(tmp.name)

    try:
        xlsx_path, csv_path = process(tmp.name)
        token = str(uuid.uuid4())
        _outputs[token] = (xlsx_path, csv_path, original_stem)
        return jsonify({"success": True, "token": token, "stem": original_stem})

    except MissingTabsError as e:
        return jsonify({"success": False, "error": str(e)})
    except HeadingMismatchError as e:
        return jsonify({"success": False, "error": str(e)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
    finally:
        # Clean up the upload temp file
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


@app.route("/download/<token>/<fmt>")
def download(token, fmt):
    if token not in _outputs:
        return "File not found or expired.", 404

    xlsx_path, csv_path, stem = _outputs[token]

    if fmt == "xlsx":
        path = xlsx_path
        mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = f"{stem}_import-ready.xlsx"
    elif fmt == "csv":
        path = csv_path
        mimetype = "text/csv"
        filename = f"{stem}_import-ready.csv"
    else:
        return "Invalid format.", 400

    if not os.path.exists(path):
        return "File not found.", 404

    return send_file(path, as_attachment=True, download_name=filename, mimetype=mimetype)


if __name__ == "__main__":
    print("Starting BCAP Import Helper at http://localhost:5050")
    app.run(port=5050, debug=False)
