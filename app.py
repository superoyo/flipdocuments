import os
import uuid
import json
import sqlite3
import traceback
from datetime import datetime, timezone
from pathlib import Path
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

import openpyxl
from scripts.convert import find_sheet_and_header, parse, build

BASE = Path(__file__).parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "outputs"
DATA_DIR = BASE / "data"
for d in (UPLOAD_DIR, OUTPUT_DIR, DATA_DIR):
    d.mkdir(exist_ok=True)

DB_PATH = DATA_DIR / "history.db"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT UNIQUE NOT NULL,
                created_at TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                input_size INTEGER,
                output_filename TEXT,
                status TEXT NOT NULL,
                error_msg TEXT,
                meta TEXT,
                summary TEXT
            )
        """)
        conn.commit()


init_db()


def log_history(job_id, original_filename, input_size, status,
                output_filename=None, error_msg=None, meta=None, summary=None):
    with get_db() as conn:
        conn.execute("""
            INSERT INTO history
            (job_id, created_at, original_filename, input_size, output_filename,
             status, error_msg, meta, summary)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            job_id,
            datetime.now(timezone.utc).isoformat(),
            original_filename,
            input_size,
            output_filename,
            status,
            error_msg,
            json.dumps(meta, ensure_ascii=False) if meta else None,
            json.dumps(summary, ensure_ascii=False) if summary else None,
        ))
        conn.commit()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/history-page")
def history_page():
    return render_template("history.html")


@app.route("/history")
def history_list():
    with get_db() as conn:
        rows = conn.execute("""
            SELECT job_id, created_at, original_filename, input_size,
                   output_filename, status, error_msg, summary
            FROM history
            ORDER BY id DESC
            LIMIT 500
        """).fetchall()
    items = []
    for r in rows:
        in_exists = (UPLOAD_DIR / f"{r['job_id']}.xlsx").exists()
        out_exists = r["output_filename"] and (OUTPUT_DIR / f"{r['job_id']}_out.xlsx").exists()
        items.append({
            "job_id": r["job_id"],
            "created_at": r["created_at"],
            "original_filename": r["original_filename"],
            "input_size": r["input_size"],
            "output_filename": r["output_filename"],
            "status": r["status"],
            "error_msg": r["error_msg"],
            "summary": json.loads(r["summary"]) if r["summary"] else None,
            "input_url": f"/download-input/{r['job_id']}" if in_exists else None,
            "output_url": f"/download/{r['job_id']}" if out_exists else None,
        })
    return jsonify({"items": items})


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
    original = f.filename
    in_path = UPLOAD_DIR / f"{job_id}.xlsx"
    out_name = original.rsplit(".xlsx", 1)[0] + "_Media_Grid.xlsx"
    out_path = OUTPUT_DIR / f"{job_id}_out.xlsx"

    f.save(in_path)
    in_size = in_path.stat().st_size

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
            err = "ไม่พบสปอตในไฟล์ — กรุณาตรวจสอบว่าไฟล์เป็น Pulzar-style (มีคอลัมน์ CH / weekday / Mat. / Cost / Time)"
            log_history(job_id, original, in_size, "error", error_msg=err, meta=meta)
            return jsonify({"error": err}), 400
        info = build(rows, date_wd, legend, dur_letter, month_txt, meta, str(out_path))
        summary = {
            "rows": info["rows"],
            "channels": info["channels"],
            "spots": info["spots"],
            "money": info["money"],
            "ndays": info["ndays"],
            "month": info["month"],
        }
        log_history(job_id, original, in_size, "success",
                    output_filename=out_name, meta=meta, summary=summary)
        return jsonify({
            "ok": True,
            "download_url": f"/download/{job_id}",
            "summary": summary,
        })
    except SystemExit as e:
        err = str(e)
        log_history(job_id, original, in_size, "error", error_msg=err, meta=meta)
        return jsonify({"error": err}), 400
    except Exception as e:
        traceback.print_exc()
        err = f"เกิดข้อผิดพลาด: {e}"
        log_history(job_id, original, in_size, "error", error_msg=err, meta=meta)
        return jsonify({"error": err}), 500


@app.route("/download/<job_id>")
@app.route("/download/<job_id>/<name>")  # backward compat
def download(job_id, name=None):
    with get_db() as conn:
        row = conn.execute(
            "SELECT output_filename FROM history WHERE job_id=? AND status='success'",
            (job_id,)
        ).fetchone()
    path = OUTPUT_DIR / f"{job_id}_out.xlsx"
    if not row or not path.exists():
        return "ไม่พบไฟล์", 404
    return send_file(
        path,
        as_attachment=True,
        download_name=row["output_filename"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download-input/<job_id>")
def download_input(job_id):
    with get_db() as conn:
        row = conn.execute(
            "SELECT original_filename FROM history WHERE job_id=?",
            (job_id,)
        ).fetchone()
    path = UPLOAD_DIR / f"{job_id}.xlsx"
    if not row or not path.exists():
        return "ไม่พบไฟล์", 404
    return send_file(
        path,
        as_attachment=True,
        download_name=row["original_filename"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port)
