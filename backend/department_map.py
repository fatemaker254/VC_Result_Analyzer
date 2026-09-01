"""
Hashmap of department/subject short-form codes (as used in course-code
prefixes like "BNGA-CC5", or in filenames like "MBGN.xlsx") to their full
department/subject name, used for display in reports.

EDIT DEPARTMENT_MAP below to match your institution's actual course-code
conventions. This is intentionally a plain, editable Python dict rather than
something buried in the analysis logic, so you can extend it as new
departments/codes show up.

Unknown codes don't break anything -- get_department_name() falls back to a
title-cased version of the code itself.
"""

import os
import re
from collections import Counter

DEPARTMENT_MAP = {
    "BNGA": "Bengali",
    "MBGN": "Bengali",
    "BOT": "Botany",
    "BOTA": "Botany",
    "MTH": "Mathematics",
    "MATH": "Mathematics",
    "BCH": "Bio-Chemistry",
    "BIOC": "Bio-Chemistry",
    "ECO": "Economics",
    "ECOA": "Economics",
    "CSC": "Computer Science",
    "COMS": "Computer Science",
    "ENG": "English",
    "ENGA": "English",
    "GEO": "Geography",
    "GEOG": "Geography",
    "PSC": "Political Science",
    "POLS": "Political Science",
    "HIS": "History",
    "HISA": "History",
    "STA": "Statistics",
    "STAT": "Statistics",
    "ZOO": "Zoology",
    "ZOOA": "Zoology",
    "MPHI": "Philosophy",
    "PHI": "Philosophy",
    "MWOS": "Women's Studies",
    "WOS": "Women's Studies",
    "JORG": "Journalism & Mass Communication",
}


def get_department_name(short_code: str) -> str:
    """Look up the full department name for a short code. Falls back to a
    title-cased version of the code itself if it's not in the map, so
    unknown codes still display reasonably instead of breaking."""
    if not short_code:
        return "Unknown Department"
    key = short_code.strip().upper()
    return DEPARTMENT_MAP.get(key, short_code.strip().title())


def guess_department_code(status_columns, filename=None):
    """
    Guess a department's short code, either from:
      1) the common alphabetic prefix shared by its "<Course> Status"
         columns (e.g. "BNGA-CC5 Status", "BNGA-CC6 Status" -> "BNGA"), or
      2) the uploaded filename (e.g. "Bengali.xlsx" -> "BENGALI") as a
         fallback if no clean common prefix is found.
    """
    prefixes = []
    for col in status_columns:
        base = col.rsplit(" ", 1)[0] if col.lower().endswith("status") else col
        m = re.match(r"^([A-Za-z]+)", base)
        if m:
            prefixes.append(m.group(1).upper())

    if prefixes:
        most_common, _ = Counter(prefixes).most_common(1)[0]
        return most_common

    if filename:
        name = os.path.splitext(os.path.basename(filename))[0]
        return name.strip().upper()

    return "UNKNOWN"
