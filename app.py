import os
import uuid
import traceback
from pathlib import Path
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

import openpyxl
from scripts.convert import find_sheet_and_header, parse, build

BASE = Path(__file__).parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "ไม่พบไฟล์ที่อัปโหลด"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "กรุณาเลือกไฟล์"}), 400
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "รองรับเฉพาะไฟล์ .xlsx"}), 400

    job_id = uuid.uuid4().hex[:12]
    safe_name = secure_filename(f.filename)
    in_path = UPLOAD_DIR / f"{job_id}_{safe_name}"
    out_name = safe_name.rsplit(".xlsx", 1)[0] + "_Media_Grid.xlsx"
    out_path = OUTPUT_DIR / f"{job_id}_{out_name}"

    f.save(in_path)

    meta = {
        "client": request.form.get("client", "").strip(),
        "product": request.form.get("product", "").strip(),
        "campaign": request.form.get("campaign", "").strip(),
        "docno": request.form.get("docno", "").strip(),
        "month": request.form.get("month", "").strip() or None,
        "year": request.form.get("year", "2569").strip() or "2569",
    }

    try:
        wb = openpyxl.load_workbook(in_path, data_only=True)
        ws, hdr, groups = find_sheet_and_header(wb)
        rows, date_wd, legend, dur_letter, month_txt = parse(ws, hdr, groups)
        if not rows:
            return jsonify({
                "error": "ไม่พบสปอตในไฟล์ — กรุณาตรวจสอบว่าไฟล์เป็น Pulzar-style (มีคอลัมน์ CH / weekday / Mat. / Cost / Time)"
            }), 400
        info = build(rows, date_wd, legend, dur_letter, month_txt, meta, str(out_path))
        return jsonify({
            "ok": True,
            "download_url": f"/download/{job_id}/{out_name}",
            "summary": {
                "rows": info["rows"],
                "channels": info["channels"],
                "spots": info["spots"],
                "money": info["money"],
                "ndays": info["ndays"],
                "month": info["month"],
            },
        })
    except SystemExit as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"เกิดข้อผิดพลาด: {e}"}), 500


@app.route("/download/<job_id>/<name>")
def download(job_id, name):
    safe_name = secure_filename(name)
    path = OUTPUT_DIR / f"{job_id}_{safe_name}"
    if not path.exists():
        return "ไม่พบไฟล์", 404
    return send_file(
        path,
        as_attachment=True,
        download_name=safe_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port)
