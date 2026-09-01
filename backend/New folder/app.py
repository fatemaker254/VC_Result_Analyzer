"""
SGPA Report Generator - Flask Backend
--------------------------------------
Endpoints:
  GET  /                       -> serves the frontend UI
  POST /api/upload             -> upload xlsx, analyze, return JSON summary + chart URLs
  GET  /api/chart/<rid>/<name> -> serve a generated chart PNG
  GET  /api/report/<rid>/download -> generate & download a .docx report with charts

Run:
  export OPENAI_API_KEY=sk-...      (or on Windows PowerShell: $env:OPENAI_API_KEY="sk-...")
  pip install -r requirements.txt
  python app.py
"""

import os
import json
import uuid

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, request, jsonify, send_file, send_from_directory, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename

from openai_helper import detect_paper_columns
from analysis import analyze_dataframe
from docx_report import build_docx_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)
CORS(app)

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}


def allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part in request. Field name must be 'file'."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only .xlsx / .xls files are supported."}), 400

    sgpa_col_hint = request.form.get("sgpa_column", "").strip() or None

    report_id = uuid.uuid4().hex[:12]
    safe_name = secure_filename(file.filename)
    save_path = os.path.join(UPLOAD_DIR, f"{report_id}_{safe_name}")
    file.save(save_path)

    try:
        df = pd.read_excel(save_path)
    except Exception as e:
        return jsonify({"error": f"Could not read Excel file: {e}"}), 400

    if df.empty:
        return jsonify({"error": "The uploaded sheet appears to be empty."}), 400

    columns = [str(c) for c in df.columns]

    try:
        detection = detect_paper_columns(columns, sample_rows=df.head(5).to_dict(orient="records"))
    except Exception as e:
        detection = heuristic_fallback(columns)
        detection["openai_error"] = str(e)

    if sgpa_col_hint and sgpa_col_hint in df.columns:
        detection["sgpa_column"] = sgpa_col_hint

    try:
        result = analyze_dataframe(df, detection)
    except Exception as e:
        return jsonify({"error": f"Analysis failed: {e}", "detection": detection}), 400

    chart_dir = os.path.join(REPORTS_DIR, report_id)
    os.makedirs(chart_dir, exist_ok=True)
    chart_files = generate_charts(result, chart_dir)

    with open(os.path.join(chart_dir, "result.json"), "w") as f:
        json.dump({"detection": detection, "result": result}, f)

    response = {
        "report_id": report_id,
        "detected": detection,
        "total_students": result["total_students"],
        "metrics": result["metrics"],
        "charts": {
            name: f"/api/chart/{report_id}/{name}.png" for name in chart_files
        },
    }
    return jsonify(response)


def heuristic_fallback(columns):
    """Fallback if OpenAI call fails: find one status/result column per paper
    (columns ending in 'Status', 'Result', or 'Grade' following a course-code
    prefix), plus the sgpa and overall-remarks columns by name."""
    status_cols, sgpa_col, remarks_col, id_cols = [], None, None, []
    for c in columns:
        lc = c.lower().strip()
        if lc == "sgpa" or lc.endswith(" sgpa") or "cgpa" in lc:
            sgpa_col = c
        elif lc in ("remarks", "result", "overall result", "overall remarks"):
            remarks_col = c
        elif any(k in lc for k in ["roll", "name", "regn", "registration", "id", "semester"]):
            id_cols.append(c)
        elif lc.endswith(" status"):
            status_cols.append(c)

    return {
        "status_columns": status_cols,
        "sgpa_column": sgpa_col,
        "remarks_column": remarks_col,
        "id_columns": id_cols,
        "total_papers": len(status_cols),
        "notes": "heuristic fallback used (OpenAI detection unavailable)",
    }


def _safe_pie(ax, values, labels, colors, title):
    """Draw a pie chart, handling the all-zero case gracefully instead of
    crashing (matplotlib's pie() divides by the total to compute wedge
    angles, so an all-zero slice set produces a 0/0 -> NaN -> ValueError).
    This happens e.g. when a sheet has very few students and all of them
    fall outside every SGPA band (e.g. everyone failed, so SGPA is blank)."""
    total = sum(values)
    if total == 0:
        ax.text(0.5, 0.5, "No data\nfor this chart", ha="center", va="center",
                 fontsize=12, color="#666666", transform=ax.transAxes)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.set_title(title)
        return
    ax.pie(values, labels=labels, autopct="%1.1f%%", colors=colors, startangle=90)
    ax.set_title(title)


def generate_charts(result, out_dir):
    """Create PNG charts for each metric + one summary bar chart.
    Returns list of chart base names (without extension)."""
    plt.rcParams.update({"font.size": 11})
    names = []

    metrics = result["metrics"]
    labels = [
        "All Papers\nCleared",
        "Fail\n(>=1 paper)",
        "SGPA < 5",
        "SGPA 5.1-7",
        "SGPA > 7.1",
    ]
    values = [
        metrics["all_cleared"]["count"],
        metrics["fail"]["count"],
        metrics["sgpa_below_5"]["count"],
        metrics["sgpa_5_1_to_7"]["count"],
        metrics["sgpa_above_7_1"]["count"],
    ]
    colors = ["#2e7d32", "#c62828", "#ef6c00", "#f9a825", "#1565c0"]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(labels, values, color=colors)
    ax.set_ylabel("Number of Students")
    ax.set_title("Overall Performance Summary")
    ax.set_ylim(0, max(values + [1]) * 1.15)  # avoid a degenerate 0-height axis too
    for b in bars:
        h = b.get_height()
        ax.annotate(str(int(h)), xy=(b.get_x() + b.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points", ha="center")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "summary.png"), dpi=150)
    plt.close(fig)
    names.append("summary")

    fig, ax = plt.subplots(figsize=(5, 5))
    pass_fail_vals = [metrics["all_cleared"]["count"], metrics["fail"]["count"]]
    _safe_pie(ax, pass_fail_vals, ["All Papers Cleared", "Fail"],
              ["#2e7d32", "#c62828"], "Pass vs Fail")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "pass_fail.png"), dpi=150)
    plt.close(fig)
    names.append("pass_fail")

    fig, ax = plt.subplots(figsize=(5, 5))
    sgpa_vals = [
        metrics["sgpa_below_5"]["count"],
        metrics["sgpa_5_1_to_7"]["count"],
        metrics["sgpa_above_7_1"]["count"],
    ]
    _safe_pie(ax, sgpa_vals, ["< 5", "5.1 - 7", "> 7.1"],
              ["#ef6c00", "#f9a825", "#1565c0"], "SGPA Distribution")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "sgpa_distribution.png"), dpi=150)
    plt.close(fig)
    names.append("sgpa_distribution")

    return names


@app.route("/api/chart/<report_id>/<filename>")
def get_chart(report_id, filename):
    chart_dir = os.path.join(REPORTS_DIR, secure_filename(report_id))
    return send_from_directory(chart_dir, secure_filename(filename))


@app.route("/api/report/<report_id>/download")
def download_report(report_id):
    chart_dir = os.path.join(REPORTS_DIR, secure_filename(report_id))
    result_path = os.path.join(chart_dir, "result.json")
    if not os.path.exists(result_path):
        return jsonify({"error": "Report not found. Please upload again."}), 404

    with open(result_path) as f:
        saved = json.load(f)

    docx_path = os.path.join(chart_dir, "SGPA_Report.docx")
    build_docx_report(saved["result"], saved["detection"], chart_dir, docx_path)

    return send_file(
        docx_path,
        as_attachment=True,
        download_name="SGPA_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
