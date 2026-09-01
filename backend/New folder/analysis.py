"""
Core analysis logic — tuned to real university mark-sheet exports where each
paper/course occupies a GROUP of columns (Theory, Tutorial/Practical,
Internal Assessment, Total, Grade, Credit, Credit Points, Status), and there
is one "<Course> Status" column per paper holding a per-paper pass/fail
result (e.g. "P", "F(TH)").

Two ways a student's overall pass/fail is determined, in priority order:

  1) A "remarks_column" (e.g. "Remarks") holding an overall semester result
     like "Semester Cleared" / "Semester Not Cleared". This is the most
     authoritative source since it already accounts for the institution's
     own credit/aggregate rules, so we use it whenever present.

  2) Fallback: if no remarks column exists, a student "fails" if ANY of
     their per-paper status columns is not a passing status. Passing status
     values are configurable in PASS_STATUS_MARKERS_PREFIX below (default:
     any status starting with "P", e.g. "P", "Pass").

SGPA is read from the detected sgpa_column and coerced to numeric — sheets
that put "N.A." (or leave it blank) for failing students naturally fall out
of all three SGPA bands, which is usually what's wanted (SGPA bands describe
students who actually have an SGPA).
"""

import math
import pandas as pd

PASS_STATUS_MARKERS_PREFIX = "p"  # per-paper status values starting with this (case-insensitive) count as pass


def _paper_passed(status_value):
    if status_value is None:
        return False
    if isinstance(status_value, float) and math.isnan(status_value):
        return False
    s = str(status_value).strip().lower()
    return s.startswith(PASS_STATUS_MARKERS_PREFIX)


def _remarks_says_fail(remarks_value):
    if remarks_value is None:
        return None
    if isinstance(remarks_value, float) and math.isnan(remarks_value):
        return None
    s = str(remarks_value).strip().lower()
    if "not" in s and "clear" in s:
        return True
    if "clear" in s:
        return False
    if "fail" in s:
        return True
    if "pass" in s:
        return False
    return None


def analyze_dataframe(df: pd.DataFrame, detection: dict) -> dict:
    status_columns = [c for c in detection.get("status_columns", []) if c in df.columns]
    sgpa_column = detection.get("sgpa_column")
    remarks_column = detection.get("remarks_column")

    if not status_columns:
        raise ValueError(
            "No per-paper status columns were detected (looked for columns like "
            "'<Course> Status'). Check that the sheet has one such column per paper, "
            "or pass an override."
        )
    if not sgpa_column or sgpa_column not in df.columns:
        raise ValueError(
            f"SGPA column '{sgpa_column}' was not found in the sheet columns: {list(df.columns)}"
        )

    total_students = len(df)

    if remarks_column and remarks_column in df.columns:
        classified = df[remarks_column].apply(_remarks_says_fail)
        unknown_mask = classified.isna()
        if unknown_mask.any():
            fallback_fail = df.loc[unknown_mask, status_columns].apply(
                lambda row: any(not _paper_passed(v) for v in row), axis=1
            )
            classified.loc[unknown_mask] = fallback_fail
        fail_mask = classified.astype(bool)
    else:
        fail_mask = df[status_columns].apply(
            lambda row: any(not _paper_passed(v) for v in row), axis=1
        )

    all_cleared_mask = ~fail_mask

    sgpa_numeric = pd.to_numeric(df[sgpa_column], errors="coerce")

    below_5_mask = sgpa_numeric < 5.0
    mid_mask = (sgpa_numeric >= 5.1) & (sgpa_numeric <= 7.0)
    above_mask = sgpa_numeric >= 7.1

    def pack(mask):
        return {
            "count": int(mask.sum()),
            "percent": round(float(mask.sum()) / total_students * 100, 2) if total_students else 0.0,
        }

    metrics = {
        "all_cleared": pack(all_cleared_mask),
        "fail": pack(fail_mask),
        "sgpa_below_5": pack(below_5_mask),
        "sgpa_5_1_to_7": pack(mid_mask),
        "sgpa_above_7_1": pack(above_mask),
    }

    paper_names = [c.rsplit(" ", 1)[0] for c in status_columns]

    return {
        "total_students": total_students,
        "total_papers": len(status_columns),
        "paper_columns": paper_names,
        "status_columns": status_columns,
        "sgpa_column": sgpa_column,
        "remarks_column": remarks_column,
        "metrics": metrics,
    }
