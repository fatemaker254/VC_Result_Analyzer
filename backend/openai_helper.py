"""
Uses the OpenAI API to inspect the Excel column headers (and a few sample
rows) and classify them for a university mark-sheet where each paper/course
spans a GROUP of columns (Theory, Tutorial/Practical, Internal Assessment,
Total, Grade, Credit, Credit Points, Status).

We only send column headers + a handful of sample rows (not the full
dataset) to keep this cheap and fast, and to avoid sending unnecessary
student data.
"""

import os
import json
from openai import OpenAI

MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

_client = None


def _get_client():
    global _client
    if _client is None:
        _client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    return _client


SYSTEM_PROMPT = """You are a data-structure classifier for university mark sheets / grade sheets.
You will be given a list of column headers from an Excel sheet, plus a few sample rows.

Many university mark sheets group several columns per paper/subject/course, e.g. for a course
named "BNGA-CC5" you might see: "BNGA-CC5 Theory", "BNGA-CC5 Theory Full Marks",
"BNGA-CC5 Tutorial", "BNGA-CC5 Internal Assessment", "BNGA-CC5 Total", "BNGA-CC5 Grade",
"BNGA-CC5 Credit", "BNGA-CC5 Credit Points", "BNGA-CC5 Status". In cases like this, each
DISTINCT course/paper corresponds to ONE "Status" column (or if there is no "Status" column,
one "Grade" or "Result" column) -- count papers by counting these, not by counting every
column in the group.

Identify:
  - "status_columns": the list of column headers that hold a PASS/FAIL style result for an
     individual paper/course (commonly ending in "Status", "Result", or "Grade" and repeated
     once per paper/subject). One entry per distinct paper/course.
  - "sgpa_column": the column holding the overall precomputed SGPA / CGPA for the student
     (may contain non-numeric values like "N.A." for failing students -- still pick it).
  - "remarks_column": a column giving the OVERALL semester/exam result for the student as a
     whole (e.g. values like "Semester Cleared" / "Semester Not Cleared", "Pass"/"Fail",
     "Promoted"/"Detained"). This is different from per-paper status. Null if none exists.
  - "id_columns": identifying info such as roll number, registration number, name, semester.

Respond ONLY with strict JSON, no markdown, no commentary, in this exact shape:
{
  "status_columns": ["<header>", ...],
  "sgpa_column": "<header or null>",
  "remarks_column": "<header or null>",
  "id_columns": ["<header>", ...],
  "total_papers": <integer, count of status_columns>
}
"""


def detect_paper_columns(columns, sample_rows=None):
    """
    columns: list of column header strings
    sample_rows: list of dicts (a few sample rows) to give the model context
    Returns a dict with status_columns, sgpa_column, remarks_column, id_columns, total_papers
    """
    if not os.environ.get("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY is not set")

    user_content = {
        "columns": columns,
        "sample_rows": sample_rows or [],
    }

    resp = _get_client().chat.completions.create(
        model=MODEL,
        temperature=0,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(user_content, default=str)},
        ],
        response_format={"type": "json_object"},
    )

    raw = resp.choices[0].message.content
    data = json.loads(raw)

    data.setdefault("status_columns", [])
    data.setdefault("sgpa_column", None)
    data.setdefault("remarks_column", None)
    data.setdefault("id_columns", [])
    data["total_papers"] = len(data["status_columns"])

    return data
