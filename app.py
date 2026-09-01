"""
SGPA Report Generator - Flask Backend
--------------------------------------
Endpoints:
  GET  /                                            -> serves the frontend UI

  --- Original single-file flow (kept exactly as before) ---
  POST /api/upload                                  -> upload ONE xlsx, analyze, return JSON
  GET  /api/chart/<rid>/<name>                       -> serve a generated chart PNG
  GET  /api/report/<rid>/download                    -> download that report as .docx

  --- New multi-department flow ---
  POST /api/upload-batch                             -> upload 1+ xlsx files (+ optional
                                                          "semester" field), analyze each as
                                                          its own department, plus a
                                                          cross-department comparison
  GET  /api/batch-chart/<batch_id>/<dept_code>/<name> -> serve a per-department chart PNG
  GET  /api/batch-chart/<batch_id>/comparison/<name> -> serve a comparison chart PNG
  GET  /api/batch/<batch_id>/department/<dept_code>/download  -> download one dept's .docx
  GET  /api/batch/<batch_id>/comparison/download               -> download the comparison .docx

If you upload just one file through /api/upload-batch, the "comparison"
section in the response is still computed (trivially, over 1 department) —
the frontend simply doesn't render it when there's only one department, so
the experience matches the original single-file flow.

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
from analysis import analyze_dataframe, build_comparison
from docx_report import build_docx_report, build_comparison_docx
from department_map import get_department_name, guess_department_code

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
    """Draw a pie chart, handling two edge cases that matplotlib doesn't:

    1) All-zero values: ax.pie() divides each value by the total to compute
       wedge angles, so an all-zero slice set produces a 0/0 -> NaN ->
       ValueError. Happens e.g. when very few students exist and none of
       them fall into any SGPA band (everyone failed, so SGPA is blank).

    2) Some (but not all) values are zero: matplotlib still draws a
       zero-width wedge for a 0% slice and places its label at the same
       angle as the neighboring slice's label, so the two labels/percentages
       overlap and become unreadable. Fix: drop zero-value slices entirely
       before plotting -- a slice with no students shouldn't be drawn.
    """
    total = sum(values)
    if total == 0:
        ax.text(0.5, 0.5, "No data\nfor this chart", ha="center", va="center",
                 fontsize=12, color="#666666", transform=ax.transAxes)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.set_title(title)
        return

    nonzero = [(v, l, c) for v, l, c in zip(values, labels, colors) if v > 0]
    values, labels, colors = zip(*nonzero)

    ax.pie(values, labels=labels, autopct="%1.1f%%", colors=colors, startangle=90)
    ax.set_title(title)


def generate_charts(result, out_dir):
    """Create PNG charts for one department: a summary bar chart, a pass/fail
    pie, and an SGPA-distribution pie. Returns list of chart base names
    (without extension). Used by both the single-file flow and each
    department inside a batch upload."""
    plt.rcParams.update({"font.size": 11})
    names = []

    metrics = result["metrics"]
    labels = [
        "All Papers\nCleared",
        "Fail\n(>=1 paper)",
        "SGPA < 5",
        "SGPA 5-7.5",
        "SGPA > 7.5",
    ]
    values = [
        metrics["all_cleared"]["count"],
        metrics["fail"]["count"],
        metrics["sgpa_below_5"]["count"],
        metrics["sgpa_5_to_7_5"]["count"],
        metrics["sgpa_above_7_5"]["count"],
    ]
    colors = ["#2e7d32", "#c62828", "#ef6c00", "#f9a825", "#1565c0"]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(labels, values, color=colors)
    ax.set_ylabel("Number of Students")
    ax.set_title("Overall Performance Summary")
    ax.set_ylim(0, max(values + [1]) * 1.15)
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
        metrics["sgpa_5_to_7_5"]["count"],
        metrics["sgpa_above_7_5"]["count"],
    ]
    _safe_pie(ax, sgpa_vals, ["< 5", "5 - 7.5", "> 7.5"],
              ["#ef6c00", "#f9a825", "#1565c0"], "SGPA Distribution")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "sgpa_distribution.png"), dpi=150)
    plt.close(fig)
    names.append("sgpa_distribution")

    return names


def generate_comparison_charts(comparison, out_dir):
    """Create the cross-department comparison charts: a grouped bar chart of
    cleared% vs not-cleared% per department (matches the client's Table 2),
    and a stacked bar chart of SGPA-band % per department. Returns list of
    chart base names (without extension)."""
    plt.rcParams.update({"font.size": 10})
    names = []

    rows = comparison["rows"] + [comparison["aggregate"]]
    labels = [r["department_name"] for r in rows]
    width_fig = max(8, len(labels) * 0.9)

    # --- Table 2 style: cleared% vs not-cleared% grouped bars ---
    cleared_pct = [r["cleared_pct"] for r in rows]
    not_cleared_pct = [r["not_cleared_pct"] for r in rows]
    x = list(range(len(labels)))
    bar_width = 0.38

    fig, ax = plt.subplots(figsize=(width_fig, 5.5))
    b1 = ax.bar([i - bar_width / 2 for i in x], cleared_pct, bar_width,
                label="Cleared %", color="#4c72b0")
    b2 = ax.bar([i + bar_width / 2 for i in x], not_cleared_pct, bar_width,
                label="Not Cleared %", color="#c44e52")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Percent")
    ax.set_title("Department-wise Students' Performance")
    ax.legend()
    for bars in (b1, b2):
        for b in bars:
            h = b.get_height()
            if h > 0:
                ax.annotate(f"{h:.1f}", xy=(b.get_x() + b.get_width() / 2, h),
                            xytext=(0, 2), textcoords="offset points", ha="center", fontsize=7)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "dept_comparison.png"), dpi=150)
    plt.close(fig)
    names.append("dept_comparison")

    # --- Table 3 style: SGPA band % stacked bars ---
    below = [r["sgpa_below_5"]["percent"] for r in rows]
    mid = [r["sgpa_5_to_7_5"]["percent"] for r in rows]
    above = [r["sgpa_above_7_5"]["percent"] for r in rows]
    bottom_mid = below
    bottom_above = [b + m for b, m in zip(below, mid)]

    fig, ax = plt.subplots(figsize=(width_fig, 5.5))
    ax.bar(labels, below, label="SGPA < 5", color="#ef6c00")
    ax.bar(labels, mid, bottom=bottom_mid, label="SGPA 5 - 7.5", color="#f9a825")
    ax.bar(labels, above, bottom=bottom_above, label="SGPA > 7.5", color="#1565c0")
    ax.set_ylabel("Percent")
    ax.set_title("SGPA-wise Department Performance (%)")
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.legend()
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "sgpa_comparison.png"), dpi=150)
    plt.close(fig)
    names.append("sgpa_comparison")

    return names


def _analyze_one_file(file_storage, tag, sgpa_override=None):
    """Shared logic: save an uploaded file, read it, run OpenAI/heuristic
    column detection, and analyze it. Returns (detection, result) or
    raises with a descriptive message."""
    safe_name = secure_filename(file_storage.filename)
    save_path = os.path.join(UPLOAD_DIR, f"{tag}_{safe_name}")
    file_storage.save(save_path)

    df = pd.read_excel(save_path)
    if df.empty:
        raise ValueError("The uploaded sheet appears to be empty.")

    columns = [str(c) for c in df.columns]

    try:
        detection = detect_paper_columns(columns, sample_rows=df.head(5).to_dict(orient="records"))
    except Exception as e:
        detection = heuristic_fallback(columns)
        detection["openai_error"] = str(e)

    if sgpa_override and sgpa_override in df.columns:
        detection["sgpa_column"] = sgpa_override

    result = analyze_dataframe(df, detection)
    return detection, result


# ---------------------------------------------------------------------------
# Original single-file flow — UNCHANGED behavior (still works exactly as
# before; use this if you only ever have one department per report)
# ---------------------------------------------------------------------------

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

    try:
        detection, result = _analyze_one_file(file, report_id, sgpa_override=sgpa_col_hint)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

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
        "charts": {name: f"/api/chart/{report_id}/{name}.png" for name in chart_files},
    }
    return jsonify(response)


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


# ---------------------------------------------------------------------------
# New multi-department batch flow
# ---------------------------------------------------------------------------

@app.route("/api/upload-batch", methods=["POST"])
def upload_batch():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded. Field name must be 'files'."}), 400

    semester = request.form.get("semester", "").strip()

    batch_id = uuid.uuid4().hex[:12]
    batch_dir = os.path.join(REPORTS_DIR, "batch_" + batch_id)
    os.makedirs(batch_dir, exist_ok=True)

    departments = []
    file_errors = []

    for f in files:
        if f.filename == "" or not allowed_file(f.filename):
            file_errors.append({"file": f.filename, "error": "Unsupported or empty file."})
            continue

        try:
            detection, result = _analyze_one_file(f, f"{batch_id}_{len(departments)}")
        except Exception as e:
            file_errors.append({"file": f.filename, "error": str(e)})
            continue

        dept_code = guess_department_code(result["status_columns"], filename=f.filename)
        dept_name = get_department_name(dept_code)

        dept_dir = os.path.join(batch_dir, secure_filename(dept_code))
        os.makedirs(dept_dir, exist_ok=True)
        chart_files = generate_charts(result, dept_dir)

        with open(os.path.join(dept_dir, "result.json"), "w") as fh:
            json.dump({
                "detection": detection,
                "result": result,
                "department_code": dept_code,
                "department_name": dept_name,
                "semester": semester,
            }, fh)

        departments.append({
            "department_code": dept_code,
            "department_name": dept_name,
            "total_students": result["total_students"],
            "total_papers": result["total_papers"],
            "metrics": result["metrics"],
            "charts": {
                name: f"/api/batch-chart/{batch_id}/{dept_code}/{name}.png" for name in chart_files
            },
        })

    if not departments:
        return jsonify({"error": "No files could be analyzed.", "file_errors": file_errors}), 400

    comparison = build_comparison(departments)

    comp_dir = os.path.join(batch_dir, "comparison")
    os.makedirs(comp_dir, exist_ok=True)
    comparison_chart_files = generate_comparison_charts(comparison, comp_dir)
    comparison["charts"] = {
        name: f"/api/batch-chart/{batch_id}/comparison/{name}.png" for name in comparison_chart_files
    }

    batch_meta = {
        "batch_id": batch_id,
        "semester": semester,
        "departments": departments,
        "comparison": comparison,
        "file_errors": file_errors,
    }
    with open(os.path.join(batch_dir, "batch.json"), "w") as fh:
        json.dump(batch_meta, fh)

    return jsonify(batch_meta)


@app.route("/api/batch-chart/<batch_id>/<dept_code>/<filename>")
def get_batch_chart(batch_id, dept_code, filename):
    chart_dir = os.path.join(REPORTS_DIR, "batch_" + secure_filename(batch_id), secure_filename(dept_code))
    return send_from_directory(chart_dir, secure_filename(filename))


@app.route("/api/batch/<batch_id>/department/<dept_code>/download")
def download_department_report(batch_id, dept_code):
    dept_dir = os.path.join(REPORTS_DIR, "batch_" + secure_filename(batch_id), secure_filename(dept_code))
    result_path = os.path.join(dept_dir, "result.json")
    if not os.path.exists(result_path):
        return jsonify({"error": "Department report not found."}), 404

    with open(result_path) as f:
        saved = json.load(f)

    docx_path = os.path.join(dept_dir, "Department_Report.docx")
    build_docx_report(
        saved["result"], saved["detection"], dept_dir, docx_path,
        department_name=saved.get("department_name"),
        semester=saved.get("semester"),
    )

    download_name = f"{saved.get('department_name', dept_code)}_Report.docx".replace(" ", "_")
    return send_file(
        docx_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/batch/<batch_id>/comparison/download")
def download_comparison_report(batch_id):
    batch_dir = os.path.join(REPORTS_DIR, "batch_" + secure_filename(batch_id))
    meta_path = os.path.join(batch_dir, "batch.json")
    if not os.path.exists(meta_path):
        return jsonify({"error": "Batch not found. Please upload again."}), 404

    with open(meta_path) as f:
        batch_meta = json.load(f)

    comp_dir = os.path.join(batch_dir, "comparison")
    docx_path = os.path.join(comp_dir, "Comparison_Report.docx")
    build_comparison_docx(batch_meta, comp_dir, docx_path)

    return send_file(
        docx_path,
        as_attachment=True,
        download_name="Department_Comparison_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)