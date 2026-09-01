# SGPA Report Generator

Upload a mark-sheet Excel file → auto-detects paper/subject columns via the
OpenAI API → analyzes pass/fail and SGPA bands with pandas → shows charts in
the browser → download a .docx report with the charts embedded.

## Structure
```
backend/
  app.py            Flask API (upload, chart serving, docx download)
  openai_helper.py  Calls OpenAI to classify columns (paper/sgpa/id)
  analysis.py       Pass/fail + SGPA-band computation (pandas)
  docx_report.py    Builds the downloadable .docx with python-docx
  requirements.txt
frontend/
  index.html        Single-page UI (vanilla JS, no build step)
```

## Setup

```bash
cd backend
pip install -r requirements.txt
export OPENAI_API_KEY=sk-...        # required for column auto-detection
python app.py                        # runs on http://localhost:5000
```

Then just open `frontend/index.html` in a browser (or serve it with any
static server). It talks to `http://localhost:5000` by default — change
`API_BASE` at the top of the `<script>` in `index.html` if your backend runs
elsewhere.

## How column detection works

`openai_helper.py` sends only the **column headers** + a few sample rows
(not the whole dataset) to the OpenAI API and asks it to classify each
column as `paper`, `sgpa`, `id`, or `other`. If the API key is missing or the
call fails, `app.py` falls back to a simple heuristic (column names
containing "paper"/"sub"/"subject", or "sgpa") so the app still works without
a key — just less robustly.

You can also override the SGPA column manually by sending a `sgpa_column`
form field with the upload request.

## Adjusting pass/fail rules

Real mark sheets vary a lot. Open `backend/analysis.py` and adjust:
- `FAIL_MARKERS` — text values that count as a fail (F, AB, RA, etc.)
- `PASS_MARK_THRESHOLD` — numeric cutoff for marks-out-of-100 columns
  (columns that look like 0–10 grade points instead use "0 = fail")

This was built against a synthetic sample sheet since the actual
`Bengali.xlsx` file didn't come through in the upload. **Once you share the
real file (or its column headers), send it over and I'll tune
`analysis.py`'s fail-detection and SGPA parsing to match your exact format**
— e.g. if fails are marked with a specific grade letter, or SGPA is stored
as text, etc.

## Endpoints

- `POST /api/upload` — multipart form, field `file` (the .xlsx). Optional
  field `sgpa_column` to override auto-detection.
- `GET /api/chart/<report_id>/<name>.png` — chart images
  (`summary`, `pass_fail`, `sgpa_distribution`)
- `GET /api/report/<report_id>/download` — the .docx report
