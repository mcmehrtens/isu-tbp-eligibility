#!/usr/bin/env python3
"""Prepare TBP eligibility data from ISU enrollment spreadsheets.

Reads an Excel spreadsheet (optionally password-protected) containing student
enrollment data, cross-references it against the preloaded members on the TBP
template spreadsheet, and produces a CSV file formatted for the Tau Beta Pi
eligibility template.

Usage:
    python prep_eligibility_data.py <semester> <input.xlsx> <template.xls> [-o output.csv]

The <semester> argument identifies the current semester and is used to detect
students whose computed graduation date has already passed (e.g. S2026 for
Spring 2026, F2026 for Fall 2026).
"""

import argparse
import csv
import getpass
import io
import re
import sys
from collections import Counter
from pathlib import Path

import msoffcrypto
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import xlrd

# =============================================================================
# CONFIGURATION — Edit these values as needed for future semesters.
# =============================================================================

# Sheet names in the input workbook to read as eligibility data.
JUNIOR_SHEET = "Juniors"
SENIOR_SHEET = "Seniors"

# Sheet name in the TBP template workbook that contains preloaded members.
TEMPLATE_MEMBERS_SHEET = "CurrentMembers"

# Number of semesters (fall/spring only) from admission to graduation.
# ISU standard undergraduate timeline is 8 semesters (4 years).
SEMESTERS_TO_GRADUATION = 8

# ISU graduation months by semester.
FALL_GRADUATION_MONTH = "Dec"
SPRING_GRADUATION_MONTH = "May"

# Column headers in the INPUT spreadsheet.
COL_FIRST = "First Name"
COL_MIDDLE = "Middle Name"
COL_LAST = "Last Name"
COL_ADMITTED = "Admitted To Academic Period"
COL_PROGRAM = "Program of Study"
COL_EMAIL = "PRIMARY_INSTITUTIONAL_EMAIL_ADDRESS"
COL_TBP_UNDERGRAD = "Tau_Beta_Pi_Undergrad"

# Column headers in the TEMPLATE spreadsheet's CurrentMembers sheet.
TMPL_FIRST = "FirstName"
TMPL_MIDDLE = "MiddleName"
TMPL_LAST = "LastName"
TMPL_JRSR = "JRSR"
TMPL_GRAD_MONTH = "GradMonth"
TMPL_GRAD_YEAR = "GradYear"
TMPL_CURRICULUM = "Curriculum"
TMPL_MEMBER = "CurrentMember"
TMPL_EMAIL = "Email"

# Value in the Tau_Beta_Pi_Undergrad column that indicates current membership.
TBP_MEMBER_VALUE = "Tau Beta Pi - Undergraduate"

# Mapping from "Program of Study" values in the input spreadsheet to the
# curriculum names used in the TBP eligibility template. If a new program is
# added at ISU, add a new entry here.
CURRICULUM_MAP: dict[str, str] = {
    "Aerospace Engineering, B.S.": "Aerospace engg",
    "Agricultural Engineering, B.S.": "Agricultural engg",
    "Biological Systems Engineering, B.S.": "Biological systems engg",
    "Biomedical Engineering, B.S.": "Biomedical engg",
    "Chemical Engineering, B.S.": "Chemical engg",
    "Civil Engineering, B.S.": "Civil engg",
    "Computer Engineering, B.S.": "Computer engg",
    "Construction Engineering, B.S.": "Construction engg",
    "Cyber Security Engineering, B.S.": "Cyber security engg",
    "Electrical Engineering, B.S.": "Electrical engg",
    "Environmental Engineering, B.S.": "Environmental engg",
    "Industrial Engineering, B.S.": "Industrial engg",
    "Materials Engineering, B.S.": "Materials engg",
    "Mechanical Engineering, B.S.": "Mechanical engg",
    "Software Engineering, B.S.": "Software engg",
}

# Output CSV column headers (must match the TBP eligibility template).
OUTPUT_HEADERS = [
    "First",
    "Middle",
    "Last",
    "Junior or Senior Class",
    "Month of Graduation",
    "Year of Graduation",
    "Curriculum",
    "Present Member",
    "Email Address",
]

# When a student's computed graduation date is in the past, assume they will
# graduate this many semesters from the current semester (including it).
FALLBACK_SEMESTERS_JUNIOR = 4
FALLBACK_SEMESTERS_SENIOR = 2

# =============================================================================
# END CONFIGURATION
# =============================================================================


# -- Semester arithmetic -------------------------------------------------------

# Semesters are represented as (year, order) tuples for easy comparison.
# Spring comes before Fall within a calendar year.
_SEMESTER_ORDER = {"Spring": 0, "Fall": 1}


def parse_current_semester(value: str) -> tuple[int, str]:
    """Parse a CLI semester argument like ``S2026`` or ``F2026``.

    Parameters
    ----------
    value : str
        ``"S"`` or ``"F"`` followed by a four-digit year.

    Returns
    -------
    tuple[int, str]
        ``(year, semester)`` where *semester* is ``"Spring"`` or ``"Fall"``.

    Raises
    ------
    argparse.ArgumentTypeError
        If the value does not match the expected format.
    """
    import argparse as _ap

    m = re.fullmatch(r"([SF])(\d{4})", value, re.IGNORECASE)
    if not m:
        raise _ap.ArgumentTypeError(
            f"Invalid semester {value!r}. Expected format: S2026 or F2026."
        )
    sem = "Spring" if m.group(1).upper() == "S" else "Fall"
    return int(m.group(2)), sem


def _semester_key(year: int, semester: str) -> tuple[int, int]:
    """Return a sortable key for a (year, semester) pair.

    Parameters
    ----------
    year : int
        Calendar year.
    semester : str
        ``"Fall"`` or ``"Spring"``.

    Returns
    -------
    tuple[int, int]
        A tuple that sorts semesters in chronological order.
    """
    return (year, _SEMESTER_ORDER[semester])


def advance_semesters(year: int, semester: str, count: int) -> tuple[int, str]:
    """Advance a semester by *count* positions (including the starting one).

    For example, advancing Spring 2026 by 2 (including itself) yields
    Fall 2026: Spring(1) -> Fall(2).

    Parameters
    ----------
    year : int
        Starting calendar year.
    semester : str
        ``"Fall"`` or ``"Spring"``.
    count : int
        Number of semesters to span, **including** the starting semester.

    Returns
    -------
    tuple[int, str]
        The resulting ``(year, semester)`` pair.
    """
    # Convert to a linear index, advance, and convert back.
    # Spring Y = 2*Y + 0, Fall Y = 2*Y + 1
    idx = 2 * year + _SEMESTER_ORDER[semester]
    idx += count - 1  # -1 because count includes the starting semester
    result_year, result_ord = divmod(idx, 2)
    result_semester = "Spring" if result_ord == 0 else "Fall"
    return result_year, result_semester


# -- Admitted period parsing ---------------------------------------------------

# Pattern 1: ACADEMIC_PERIOD-YYYYSemester  (e.g. ACADEMIC_PERIOD-2023Fall)
_PATTERN_CODED = re.compile(
    r"ACADEMIC_PERIOD-(\d{4})(Fall|Spring|Summer)", re.IGNORECASE
)

# Pattern 2: YYYY Semester Semester (MM/DD/YYYY-MM/DD/YYYY)
_PATTERN_DESCRIPTIVE = re.compile(
    r"(\d{4})\s+(Fall|Spring|Summer)\s+Semester", re.IGNORECASE
)


def parse_admitted_period(value: str) -> tuple[int, str]:
    """Extract the admission year and semester from an academic period string.

    Parameters
    ----------
    value : str
        The raw "Admitted To Academic Period" value from the spreadsheet.

    Returns
    -------
    tuple[int, str]
        A ``(year, semester)`` pair where *semester* is one of
        ``"Fall"``, ``"Spring"``, or ``"Summer"``.

    Raises
    ------
    ValueError
        If the value does not match any known format.
    """
    if m := _PATTERN_CODED.search(value):
        return int(m.group(1)), m.group(2).capitalize()
    if m := _PATTERN_DESCRIPTIVE.search(value):
        return int(m.group(1)), m.group(2).capitalize()
    raise ValueError(f"Unrecognized academic period format: {value!r}")


def compute_graduation(admit_year: int, admit_semester: str) -> tuple[str, int]:
    """Compute expected graduation month and year from admission data.

    Students who start in summer are treated as starting in the following fall.
    Graduation is calculated by advancing ``SEMESTERS_TO_GRADUATION`` semesters
    (counting only fall and spring) from the admission semester.

    Parameters
    ----------
    admit_year : int
        The year the student was admitted.
    admit_semester : str
        One of ``"Fall"``, ``"Spring"``, or ``"Summer"``.

    Returns
    -------
    tuple[str, int]
        A ``(month, year)`` pair — *month* is a three-letter abbreviation
        (e.g. ``"Dec"``), *year* is a four-digit integer.
    """
    # Treat summer admits as fall admits.
    if admit_semester == "Summer":
        admit_semester = "Fall"

    # Each academic year has two semesters: Fall then Spring.
    # Fall  of year Y -> Spring of year Y+1 -> Fall of year Y+1 -> ...
    if admit_semester == "Fall":
        # After N semesters: Fall(1), Spring(2), Fall(3), Spring(4), ...
        # N even  -> Spring of (admit_year + N//2)
        # N odd   -> Fall   of (admit_year + (N-1)//2)  [incomplete year]
        grad_semesters_ahead = SEMESTERS_TO_GRADUATION
        if grad_semesters_ahead % 2 == 0:
            grad_semester = "Spring"
            grad_year = admit_year + grad_semesters_ahead // 2
        else:
            grad_semester = "Fall"
            grad_year = admit_year + (grad_semesters_ahead + 1) // 2
    else:  # Spring
        # Spring(1), Fall(2), Spring(3), Fall(4), ...
        grad_semesters_ahead = SEMESTERS_TO_GRADUATION
        if grad_semesters_ahead % 2 == 0:
            grad_semester = "Fall"
            grad_year = admit_year + grad_semesters_ahead // 2 - 1
        else:
            grad_semester = "Spring"
            grad_year = admit_year + (grad_semesters_ahead + 1) // 2

    month = (
        FALL_GRADUATION_MONTH if grad_semester == "Fall" else SPRING_GRADUATION_MONTH
    )
    return month, grad_year


# -- Spreadsheet reading -------------------------------------------------------


def decrypt_workbook(path: Path, password: str) -> openpyxl.Workbook:
    """Open a password-protected .xlsx file and return an openpyxl Workbook.

    Parameters
    ----------
    path : Path
        Path to the encrypted ``.xlsx`` file.
    password : str
        Password used to decrypt the file.

    Returns
    -------
    openpyxl.Workbook
        The decrypted workbook loaded in read-only mode.

    Raises
    ------
    msoffcrypto.exceptions.DecryptionError
        If the password is incorrect.
    """
    with open(path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return openpyxl.load_workbook(decrypted, read_only=True, data_only=True)


def read_sheet(
    wb: openpyxl.Workbook, sheet_name: str
) -> list[dict[str, str]]:
    """Read a worksheet into a list of row dictionaries.

    The first row is treated as headers. Cell values that are ``None`` or
    consist only of whitespace are normalised to empty strings.

    Parameters
    ----------
    wb : openpyxl.Workbook
        An open openpyxl workbook.
    sheet_name : str
        Name of the sheet to read.

    Returns
    -------
    list[dict[str, str]]
        Each element maps column header -> cell value (as a stripped string).
    """
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h else "" for h in rows[0]]
    records: list[dict[str, str]] = []
    for row in rows[1:]:
        record = {}
        for header, cell in zip(headers, row):
            if not header:
                continue
            text = str(cell).strip() if cell is not None else ""
            record[header] = text
        records.append(record)
    return records


def read_template_members(path: Path) -> list[dict[str, str]]:
    """Read preloaded members from the TBP template spreadsheet (.xls).

    Parameters
    ----------
    path : Path
        Path to the TBP template ``.xls`` file.

    Returns
    -------
    list[dict[str, str]]
        Each element maps column header -> cell value (as a stripped string).
        Only non-empty rows are included.
    """
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_name(TEMPLATE_MEMBERS_SHEET)
    headers = [ws.cell_value(0, c).strip() for c in range(ws.ncols)]
    records: list[dict[str, str]] = []
    for r in range(1, ws.nrows):
        record = {}
        all_empty = True
        for c, header in enumerate(headers):
            if not header:
                continue
            val = ws.cell_value(r, c)
            # xlrd returns floats for numeric cells (e.g. year 2026.0).
            if isinstance(val, float) and val == int(val):
                text = str(int(val))
            else:
                text = str(val).strip() if val else ""
            record[header] = text
            if text:
                all_empty = False
        if not all_empty:
            records.append(record)
    return records


def read_template_curriculums(path: Path) -> set[str]:
    """Read valid curriculum names from the template's CurriculumData sheet."""
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_name("CurriculumData")
    names: set[str] = set()
    for r in range(1, ws.nrows):  # skip header row
        val = str(ws.cell_value(r, 0)).strip()
        if val:
            names.add(val)
    return names


# -- Template member matching --------------------------------------------------


def _normalize_name(name: str) -> str:
    """Remove all non-letter/digit characters and lowercase for comparison.

    This allows names with different punctuation or whitespace to compare
    equal: ``"John A. B."`` and ``"John AB"`` both become ``"johnab"``.
    """
    return re.sub(r"[^a-zA-Z0-9]", "", name).lower()


def _clean_name(name: str) -> str:
    """Strip punctuation from a name for template output.

    Interior punctuation is replaced with a space.  Leading/trailing
    punctuation is stripped entirely.  Whitespace is normalized.
    """
    cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", name)
    return " ".join(cleaned.split())


def _match_key(first: str, middle: str, last: str, email: str) -> tuple[str, ...]:
    """Build a matching key from name and email fields.

    Name fields are normalized (punctuation/whitespace removed, lowercased)
    so that minor formatting differences don't prevent matches.  The email
    is only lowercased/stripped.
    """
    return (
        _normalize_name(first),
        _normalize_name(middle),
        _normalize_name(last),
        email.strip().lower(),
    )


def _tmpl_key(tmpl: dict[str, str]) -> tuple[str, ...]:
    """Return the exact match key for a template record."""
    return _match_key(tmpl[TMPL_FIRST], tmpl[TMPL_MIDDLE], tmpl[TMPL_LAST], tmpl[TMPL_EMAIL])


def _fullname_key(first: str, middle: str, last: str) -> str:
    """Return a normalized full-name string for name-splitting-agnostic comparison.

    All punctuation and whitespace are removed so that different name
    splitting between systems still produces the same key.
    """
    return _normalize_name(f"{first} {middle} {last}")


def _find_template_match(
    first: str,
    middle: str,
    last: str,
    email: str,
    template_lookup: dict[tuple[str, ...], dict[str, str]],
    template_by_name: dict[tuple[str, ...], list[dict[str, str]]],
    template_by_first_last_email: dict[tuple[str, ...], list[dict[str, str]]],
    template_by_fullname_email: dict[tuple[str, ...], list[dict[str, str]]],
    matched_keys: set[tuple[str, ...]],
) -> tuple[dict[str, str] | None, str, list[dict[str, str]]]:
    """Find a template member using tiered matching.

    Tries four tiers in order, returning the first that produces exactly
    one unclaimed candidate:

    1. **Exact** — all four fields match (case-insensitive).
    2. **Email-relaxed** — first, middle, and last match; the template's
       email is not ``@iastate.edu`` (personal email registered instead).
    3. **Middle-relaxed** — first, last, and email match; middle names
       differ (abbreviation, initial, or missing).
    4. **Full-name-relaxed** — the space-joined concatenation of
       first+middle+last matches (handling different name splitting
       between systems); email matches.

    No two tiers combine.

    Parameters
    ----------
    first, middle, last, email : str
        Input student fields.
    template_lookup : dict
        Exact ``(first, middle, last, email)`` -> template record.
    template_by_name : dict
        ``(first, middle, last)`` -> list of template records.
    template_by_first_last_email : dict
        ``(first, last, email)`` -> list of template records.
    template_by_fullname_email : dict
        ``(fullname, email)`` -> list of template records.
    matched_keys : set
        Exact keys of template records already claimed.

    Returns
    -------
    tuple[dict | None, str, list[dict]]
        ``(template_record, match_type, ambiguous_candidates)``.
        *match_type* is ``"exact"``, ``"email_relaxed"``,
        ``"middle_relaxed"``, ``"fullname_relaxed"``,
        ``"ambiguous"``, or ``"none"``.
        *ambiguous_candidates* is populated only when *match_type* is
        ``"ambiguous"``.
    """
    fl = _normalize_name(first)
    ml = _normalize_name(middle)
    ll = _normalize_name(last)
    el = email.strip().lower()

    # Tier 1: exact match.
    key = (fl, ml, ll, el)
    tmpl = template_lookup.get(key)
    if tmpl is not None and key not in matched_keys:
        return tmpl, "exact", []

    ambiguous: list[dict[str, str]] = []

    # Tier 2: email-relaxed (names match exactly, template email is not @iastate.edu).
    name_key = (fl, ml, ll)
    candidates = [
        t for t in template_by_name.get(name_key, [])
        if not t[TMPL_EMAIL].strip().lower().endswith("@iastate.edu")
        and _tmpl_key(t) not in matched_keys
    ]
    if len(candidates) == 1:
        return candidates[0], "email_relaxed", []
    if len(candidates) > 1:
        ambiguous.extend(candidates)

    # Tier 3: middle-name-relaxed (first + last + email match, middle differs).
    fle_key = (fl, ll, el)
    candidates = [
        t for t in template_by_first_last_email.get(fle_key, [])
        if _tmpl_key(t) not in matched_keys
    ]
    if len(candidates) == 1:
        return candidates[0], "middle_relaxed", []
    if len(candidates) > 1:
        ambiguous.extend(candidates)

    # Tier 4: full-name-relaxed (concatenated name matches, email matches).
    fn_key = (_fullname_key(first, middle, last), el)
    candidates = [
        t for t in template_by_fullname_email.get(fn_key, [])
        if _tmpl_key(t) not in matched_keys
    ]
    if len(candidates) == 1:
        return candidates[0], "fullname_relaxed", []
    if len(candidates) > 1:
        ambiguous.extend(candidates)

    if ambiguous:
        # Deduplicate by template key.
        seen: set[tuple[str, ...]] = set()
        unique: list[dict[str, str]] = []
        for t in ambiguous:
            tk = _tmpl_key(t)
            if tk not in seen:
                seen.add(tk)
                unique.append(t)
        return None, "ambiguous", unique

    return None, "none", []


def template_row_to_output(
    tmpl: dict[str, str], class_standing: str
) -> dict[str, str]:
    """Convert a template member record to an output-CSV row.

    All fields are taken from the template exactly as-is, except the
    Junior/Senior class standing which comes from the input spreadsheet.

    Parameters
    ----------
    tmpl : dict[str, str]
        A record from the template's ``CurrentMembers`` sheet.
    class_standing : str
        ``"Junior"`` or ``"Senior"`` — from the input spreadsheet tab.

    Returns
    -------
    dict[str, str]
        A row ready for the output CSV.
    """
    return {
        "First": tmpl[TMPL_FIRST],
        "Middle": tmpl[TMPL_MIDDLE],
        "Last": tmpl[TMPL_LAST],
        "Junior or Senior Class": class_standing,
        "Month of Graduation": tmpl[TMPL_GRAD_MONTH],
        "Year of Graduation": tmpl[TMPL_GRAD_YEAR],
        "Curriculum": tmpl[TMPL_CURRICULUM],
        "Present Member": tmpl[TMPL_MEMBER],
        "Email Address": tmpl[TMPL_EMAIL],
    }


# -- Row transformation --------------------------------------------------------


def transform_row(
    row: dict[str, str],
    class_standing: str,
    current_semester: tuple[int, str],
) -> dict[str, str] | None:
    """Convert one input row into an output-CSV row dictionary.

    Only used for non-member students. Member students are handled separately
    via template matching.

    Parameters
    ----------
    row : dict[str, str]
        A single record from the input spreadsheet.
    class_standing : str
        ``"Junior"`` or ``"Senior"`` — determined by which sheet the row
        came from.
    current_semester : tuple[int, str]
        The current ``(year, semester)`` used to detect past graduation dates.

    Returns
    -------
    dict[str, str] | None
        The transformed row ready for CSV output, or ``None`` if the row
        could not be processed (a warning is printed to stderr).
    """
    admitted_raw = row.get(COL_ADMITTED, "")
    program = row.get(COL_PROGRAM, "")
    email = row.get(COL_EMAIL, "")
    first = row.get(COL_FIRST, "")
    middle = row.get(COL_MIDDLE, "")
    last = row.get(COL_LAST, "")

    # Look up curriculum.
    curriculum = CURRICULUM_MAP.get(program)
    if curriculum is None:
        print(
            f"  WARNING: Unknown program {program!r} for {first} {last} — skipping.",
            file=sys.stderr,
        )
        return None

    # Parse admission period and compute graduation.
    try:
        admit_year, admit_semester = parse_admitted_period(admitted_raw)
    except ValueError:
        print(
            f"  WARNING: Could not parse admitted period {admitted_raw!r} "
            f"for {first} {last} — skipping.",
            file=sys.stderr,
        )
        return None
    grad_month, grad_year = compute_graduation(admit_year, admit_semester)

    # If the computed graduation is in the past, assume the student will
    # graduate a fixed number of semesters from now based on class standing.
    grad_sem = "Fall" if grad_month == FALL_GRADUATION_MONTH else "Spring"
    if _semester_key(grad_year, grad_sem) < _semester_key(*current_semester):
        fallback = (
            FALLBACK_SEMESTERS_JUNIOR
            if class_standing == "Junior"
            else FALLBACK_SEMESTERS_SENIOR
        )
        new_year, new_sem = advance_semesters(*current_semester, fallback)
        new_month = (
            FALL_GRADUATION_MONTH
            if new_sem == "Fall"
            else SPRING_GRADUATION_MONTH
        )
        grad_month, grad_year = new_month, new_year

    return {
        "First": _clean_name(first),
        "Middle": _clean_name(middle),
        "Last": _clean_name(last),
        "Junior or Senior Class": class_standing,
        "Month of Graduation": grad_month,
        "Year of Graduation": str(grad_year),
        "Curriculum": curriculum,
        "Present Member": "",
        "Email Address": email,
    }


# -- Manual review workbook ----------------------------------------------------

# Styling constants for the manual review spreadsheet.
_SECTION_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
_SECTION_FONT = Font(bold=True, color="FFFFFF", size=12)
_HEADER_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
_HEADER_FONT = Font(bold=True)
_INSTRUCTION_FONT = Font(italic=True, size=10)
_WARNING_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
_WARNING_FONT = Font(bold=True, color="BF8F00", size=11)


def _append_styled_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    values: list[str],
    fill: PatternFill | None = None,
    font: Font | None = None,
) -> None:
    """Append a row and apply optional fill and font to each cell.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The worksheet to append to.
    values : list[str]
        Cell values for the row.
    fill : PatternFill | None
        Optional fill to apply to every cell in the row.
    font : Font | None
        Optional font to apply to every cell in the row.
    """
    ws.append(values)
    row_idx = ws.max_row
    for col_idx in range(1, len(values) + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        if fill:
            cell.fill = fill
        if font:
            cell.font = font


def write_manual_review(
    path: Path,
    unmatched_input: list[tuple[dict[str, str], str]],
    unmatched_template: list[dict[str, str]],
    miscoded_nonmembers: list[tuple[dict[str, str], str]],
) -> None:
    """Write an Excel workbook listing members that need manual review.

    All sections are on a single sheet, separated by styled section headers.

    Parameters
    ----------
    path : Path
        Output path for the ``.xlsx`` file.
    unmatched_input : list[tuple[dict[str, str], str]]
        Members from the input spreadsheet that had no match on the template.
        Each element is ``(output_row_dict, note)`` where *note* is an
        optional string describing ambiguous fuzzy-match candidates.
    unmatched_template : list[dict[str, str]]
        Members from the template spreadsheet that had no match in the input
        spreadsheet. Each dict has the template column keys.
    miscoded_nonmembers : list[tuple[dict[str, str], str]]
        Non-member students whose data matched a template entry. Each
        element is ``(output_row_dict, note)``.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Manual Review"

    # -- Section 1: Input members without a template match ---------------------
    _append_styled_row(
        ws,
        ["SECTION 1: Input members with no template match",
         "", "", "", "", "", "", ""],
        fill=_SECTION_FILL, font=_SECTION_FONT,
    )
    ws.append([
        "These students are marked as TBP members in the input spreadsheet "
        "but could NOT be matched (by first name, middle name, last name, and "
        "email) to any preloaded member on the TBP template."
    ])
    ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
    ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)
    ws.append([
        "ACTION: For each student below, find their correct entry in "
        "Section 3 and manually add them to the eligibility report using the "
        "template's data (updating Junior/Senior status as shown here)."
    ])
    ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
    ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)
    ws.append([
        "If no match exists in Section 3, look up the student in the TBP "
        "member lookup system and use that data EXACTLY as provided (except "
        "Junior/Senior status). If they are not found there either, they may "
        "be miscoded as a TBP member in the input spreadsheet."
    ])
    ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
    ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)

    headers_1 = [
        "Reviewed (X)", "First", "Middle", "Last",
        "Junior or Senior Class", "Curriculum", "Email", "Notes",
    ]
    _append_styled_row(ws, headers_1, fill=_HEADER_FILL, font=_HEADER_FONT)
    for row, note in unmatched_input:
        ws.append([
            "", row["First"], row["Middle"], row["Last"],
            row["Junior or Senior Class"], row["Curriculum"], row["Email Address"],
            note,
        ])

    # Spacer rows.
    ws.append([])
    ws.append([])

    # -- Section 2: Possibly miscoded non-members --------------------------------
    if miscoded_nonmembers:
        _append_styled_row(
            ws,
            ["SECTION 2: Possibly miscoded non-members",
             "", "", "", "", "", "", ""],
            fill=_WARNING_FILL, font=_WARNING_FONT,
        )
        ws.append([
            "These students are NOT marked as TBP members in the input "
            "spreadsheet, but their data matches (exactly or via fuzzy "
            "match) a preloaded member on the TBP template. They may be "
            "miscoded in the input data. Their corresponding template "
            "entries appear in Section 3."
        ])
        ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)
        ws.append([
            "ACTION: Verify whether each student below is actually a TBP "
            "member. If so, add them to the eligibility report using the "
            "template's data from Section 3 (with 'M' in Present Member, "
            "updating Junior/Senior status as shown here). If they are "
            "truly non-members, add them manually as non-members."
        ])
        ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)

        headers_2 = [
            "Reviewed (X)", "First", "Middle", "Last",
            "Junior or Senior Class", "Curriculum", "Email", "Notes",
        ]
        _append_styled_row(ws, headers_2, fill=_HEADER_FILL, font=_HEADER_FONT)
        for input_row, note in miscoded_nonmembers:
            ws.append([
                "", input_row["First"], input_row["Middle"],
                input_row["Last"], input_row["Junior or Senior Class"],
                input_row["Curriculum"], input_row["Email Address"], note,
            ])

        ws.append([])
        ws.append([])

    # -- Section 3: Template members without an input match --------------------
    _append_styled_row(
        ws,
        ["SECTION 3: Template members with no input member match",
         "", "", "", "", "", "", "", "", ""],
        fill=_SECTION_FILL, font=_SECTION_FONT,
    )
    ws.append([
        "These students are preloaded on the TBP template but were NOT matched "
        "to any current member in the input spreadsheet. They may have "
        "graduated, lost eligibility, or have a name/email mismatch. Also "
        "contains entries for Section 2 students."
    ])
    ws.cell(row=ws.max_row, column=1).font = _INSTRUCTION_FONT
    ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)

    headers_3 = [
        "Added (X)", "First", "Middle", "Last",
        "Junior/Senior", "Grad Month", "Grad Year",
        "Curriculum", "Present Member", "Email",
    ]
    _append_styled_row(ws, headers_3, fill=_HEADER_FILL, font=_HEADER_FONT)
    for row in unmatched_template:
        # Parse graduation year as an integer for proper numeric storage.
        try:
            grad_year = int(row[TMPL_GRAD_YEAR])
        except (ValueError, TypeError):
            grad_year = row[TMPL_GRAD_YEAR]
        ws.append([
            "", row[TMPL_FIRST], row[TMPL_MIDDLE], row[TMPL_LAST],
            row[TMPL_JRSR], row[TMPL_GRAD_MONTH], grad_year,
            row[TMPL_CURRICULUM], row[TMPL_MEMBER], row[TMPL_EMAIL],
        ])

    # Auto-size columns for readability.
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    wb.save(path)


# -- Statistics ----------------------------------------------------------------


def print_stats(
    output_rows: list[dict[str, str]],
    unmatched_input: list[tuple[dict[str, str], str]],
    miscoded_nonmembers: list[tuple[dict[str, str], str]],
    matched_members: int,
    fuzzy_matched: int,
    unmatched_template_count: int,
) -> None:
    """Print summary statistics about the processed eligibility data.

    Parameters
    ----------
    output_rows : list[dict[str, str]]
        Rows written to the output CSV.
    unmatched_input : list[tuple[dict[str, str], str]]
        Members sent to manual review (not in the CSV yet, but still
        eligible). Each element is ``(row_dict, note)``.
    miscoded_nonmembers : list[tuple[dict[str, str], str]]
        Non-members excluded from the CSV pending review. Each element
        is ``(row_dict, note)``.
    matched_members : int
        Number of members matched between input and template (exact + fuzzy).
    fuzzy_matched : int
        Number of members matched via fuzzy matching (subset of *matched_members*).
    unmatched_template_count : int
        Number of template members with no input match (sent to manual review).
    """
    # Combine all eligible students for the true eligibility count.
    unmatched_rows = [r for r, _ in unmatched_input]
    miscoded_rows = [r for r, _ in miscoded_nonmembers]
    all_eligible = output_rows + unmatched_rows + miscoded_rows

    juniors = [r for r in all_eligible if r["Junior or Senior Class"] == "Junior"]
    seniors = [r for r in all_eligible if r["Junior or Senior Class"] == "Senior"]

    print("\n=== Eligibility Totals ===\n")
    print(f"  Eligible juniors:  {len(juniors)}")
    print(f"  Eligible seniors:  {len(seniors)}")
    print(f"  ───────────────────────")
    print(f"  Total eligible:    {len(all_eligible)}")

    total_members = matched_members + len(unmatched_rows)
    print(f"\n=== Member Matching ===\n")
    print(f"  Total input members:        {total_members}")
    print(f"  Members matched:            {matched_members}")
    if fuzzy_matched:
        print(f"    (of which fuzzy-matched:  {fuzzy_matched})")
    print(f"  Members unmatched:          {len(unmatched_rows)}  (manual review needed)")
    print(f"  Template members unmatched: {unmatched_template_count}  (likely graduated/ineligible)")
    if miscoded_rows:
        print(f"  Possibly miscoded:          {len(miscoded_rows)}  (non-members matching template)")

    # Duplicates by email.
    email_counts = Counter(r["Email Address"].lower() for r in all_eligible)
    duplicates = {email: count for email, count in email_counts.items() if count > 1}
    print(f"\n  Duplicate entries (by email): {len(duplicates)}")
    if duplicates:
        for email, count in sorted(duplicates.items()):
            names = [
                f"{r['First']} {r['Last']}"
                for r in all_eligible
                if r["Email Address"].lower() == email
            ]
            print(f"    {email} ({count}x): {', '.join(names)}")

    # By curriculum.
    curriculum_counts = Counter(r["Curriculum"] for r in all_eligible)
    print("\n  By curriculum:")
    for curriculum, count in curriculum_counts.most_common():
        members = sum(
            1
            for r in all_eligible
            if r["Curriculum"] == curriculum and r["Present Member"] == "M"
        )
        member_str = f"  ({members} members)" if members else ""
        print(f"    {curriculum:<30s} {count:>4}{member_str}")
    print()


# -- Main entry point ----------------------------------------------------------


def main() -> None:
    """Parse arguments, read input, transform data, and write CSV output."""
    parser = argparse.ArgumentParser(
        description="Prepare TBP eligibility data from an ISU enrollment spreadsheet."
    )
    parser.add_argument(
        "semester",
        type=parse_current_semester,
        help="Current semester (e.g. S2026 for Spring 2026, F2026 for Fall 2026).",
    )
    parser.add_argument("input", type=Path, help="Path to the input .xlsx file.")
    parser.add_argument(
        "template", type=Path, help="Path to the TBP template .xls file."
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Path for the output CSV file. Defaults to <input_stem>_eligibility.csv.",
    )
    args = parser.parse_args()

    current_semester: tuple[int, str] = args.semester
    input_path: Path = args.input
    template_path: Path = args.template
    if not input_path.exists():
        print(f"Error: {input_path} not found.", file=sys.stderr)
        sys.exit(1)
    if not template_path.exists():
        print(f"Error: {template_path} not found.", file=sys.stderr)
        sys.exit(1)

    output_path: Path = args.output or input_path.with_name(
        f"{input_path.stem}_eligibility.csv"
    )
    review_path = output_path.with_name(
        f"{output_path.stem}_manual_review.xlsx"
    )

    # -- Read the TBP template -------------------------------------------------
    template_members = read_template_members(template_path)
    print(f"Read {len(template_members)} preloaded members from template.")

    # Validate that every CURRICULUM_MAP value exists in the template.
    template_curriculums = read_template_curriculums(template_path)
    bad_curriculums = sorted(
        v for v in CURRICULUM_MAP.values() if v not in template_curriculums
    )
    if bad_curriculums:
        print(
            "\n"
            "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
            "  WARNING: The following CURRICULUM_MAP values do NOT exist\n"
            "  in the template's CurriculumData sheet:\n",
            file=sys.stderr,
        )
        for c in bad_curriculums:
            print(f"    - {c!r}", file=sys.stderr)
        print(
            "\n"
            "  Students with these curriculums will have invalid data in\n"
            "  the eligibility report. Update CURRICULUM_MAP or the template.\n"
            "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n",
            file=sys.stderr,
        )

    # Build lookup structures for tiered template matching.
    # Exact: (first, middle, last, email) -> template record.
    template_lookup: dict[tuple[str, ...], dict[str, str]] = {}
    # By full name: (first, middle, last) -> list of template records.
    template_by_name: dict[tuple[str, ...], list[dict[str, str]]] = {}
    # By first+last+email: (first, last, email) -> list of template records.
    template_by_fle: dict[tuple[str, ...], list[dict[str, str]]] = {}
    # By concatenated fullname+email: (fullname, email) -> list of template records.
    template_by_fne: dict[tuple[str, ...], list[dict[str, str]]] = {}
    for tmpl in template_members:
        key = _tmpl_key(tmpl)
        template_lookup[key] = tmpl
        nk = (key[0], key[1], key[2])  # first, middle, last (already lowered)
        template_by_name.setdefault(nk, []).append(tmpl)
        flek = (key[0], key[2], key[3])  # first, last, email (already lowered)
        template_by_fle.setdefault(flek, []).append(tmpl)
        fnk = (_fullname_key(tmpl[TMPL_FIRST], tmpl[TMPL_MIDDLE], tmpl[TMPL_LAST]), key[3])
        template_by_fne.setdefault(fnk, []).append(tmpl)

    # Track which template members get matched (by members or non-members).
    matched_template_keys: set[tuple[str, ...]] = set()

    # -- Read the input spreadsheet --------------------------------------------
    try:
        wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    except Exception:
        password = getpass.getpass("Enter spreadsheet password: ")
        try:
            wb = decrypt_workbook(input_path, password)
        except Exception as exc:
            print(f"Error: Could not decrypt file — {exc}", file=sys.stderr)
            sys.exit(1)

    junior_rows = read_sheet(wb, JUNIOR_SHEET)
    senior_rows = read_sheet(wb, SENIOR_SHEET)
    wb.close()

    print(f"Read {len(junior_rows)} juniors and {len(senior_rows)} seniors.")

    # -- Process rows ----------------------------------------------------------
    output_rows: list[dict[str, str]] = []
    unmatched_input_members: list[tuple[dict[str, str], str]] = []
    miscoded_nonmembers: list[tuple[dict[str, str], str]] = []
    skipped = 0
    fuzzy_matched = 0

    match_args = (
        template_lookup, template_by_name, template_by_fle, template_by_fne,
        matched_template_keys,
    )

    for class_standing, sheet_rows in [("Junior", junior_rows), ("Senior", senior_rows)]:
        for row in sheet_rows:
            first = row.get(COL_FIRST, "")
            middle = row.get(COL_MIDDLE, "")
            last = row.get(COL_LAST, "")
            email = row.get(COL_EMAIL, "")
            is_member = row.get(COL_TBP_UNDERGRAD, "") == TBP_MEMBER_VALUE

            if is_member:
                # Try to match against template (exact, then fuzzy).
                tmpl, match_type, ambiguous = _find_template_match(
                    first, middle, last, email, *match_args,
                )

                if tmpl is not None:
                    output_rows.append(template_row_to_output(tmpl, class_standing))
                    matched_template_keys.add(_tmpl_key(tmpl))
                    if match_type == "email_relaxed":
                        fuzzy_matched += 1
                        print(
                            f"  FUZZY MATCH: {first} {last} — email mismatch "
                            f"(input: {email}, template: {tmpl[TMPL_EMAIL]})"
                        )
                    elif match_type == "middle_relaxed":
                        fuzzy_matched += 1
                        print(
                            f"  FUZZY MATCH: {first} {last} — middle name "
                            f"mismatch (input: '{middle}', "
                            f"template: '{tmpl[TMPL_MIDDLE]}')"
                        )
                    elif match_type == "fullname_relaxed":
                        fuzzy_matched += 1
                        print(
                            f"  FUZZY MATCH: {first} {last} — name splitting "
                            f"differs (input: '{first}'/'{middle}'/'{last}', "
                            f"template: '{tmpl[TMPL_FIRST]}'/'"
                            f"{tmpl[TMPL_MIDDLE]}'/'{tmpl[TMPL_LAST]}')"
                        )
                elif match_type == "ambiguous":
                    # Multiple fuzzy matches — cannot auto-resolve.
                    result = transform_row(row, class_standing, current_semester)
                    if result:
                        result["Present Member"] = "M"
                        cands = "; ".join(
                            f"{t[TMPL_FIRST]} {t[TMPL_MIDDLE]} {t[TMPL_LAST]} "
                            f"({t[TMPL_EMAIL]})"
                            for t in ambiguous
                        )
                        print(
                            f"  WARNING: {first} {last} — multiple fuzzy "
                            f"matches; treating as unmatched. Candidates: {cands}"
                        )
                        unmatched_input_members.append(
                            (result, f"Fuzzy matched: {cands}")
                        )
                    else:
                        skipped += 1
                else:
                    # No match at all — flag for manual review.
                    result = transform_row(row, class_standing, current_semester)
                    if result:
                        result["Present Member"] = "M"
                        unmatched_input_members.append((result, ""))
                    else:
                        skipped += 1
            else:
                # Non-member — compute everything from input data.
                result = transform_row(row, class_standing, current_semester)
                if result is None:
                    skipped += 1
                    continue

                # Check if this non-member unexpectedly matches a template
                # entry (exact or fuzzy), flagging a possible data coding
                # error.  These students are excluded from the output CSV
                # and sent to manual review instead.
                tmpl, match_type, ambiguous = _find_template_match(
                    first, middle, last, email, *match_args,
                )
                if tmpl is not None:
                    # Don't claim the template entry — it should still
                    # appear in Section 3 for the reviewer's reference.
                    if match_type == "exact":
                        print(
                            f"  WARNING: {first} {last} is NOT marked as a "
                            f"member but matches a template entry — "
                            f"possible data error.",
                        )
                        miscoded_nonmembers.append((result, ""))
                    elif match_type == "email_relaxed":
                        print(
                            f"  WARNING: {first} {last} is NOT marked as a "
                            f"member but fuzzy-matches a template entry "
                            f"(email mismatch: {email} vs "
                            f"{tmpl[TMPL_EMAIL]}) — possible data error.",
                        )
                        miscoded_nonmembers.append(
                            (result, "Email mismatch with template")
                        )
                    elif match_type == "middle_relaxed":
                        print(
                            f"  WARNING: {first} {last} is NOT marked as a "
                            f"member but fuzzy-matches a template entry "
                            f"(middle name mismatch: '{middle}' vs "
                            f"'{tmpl[TMPL_MIDDLE]}') — possible data error.",
                        )
                        miscoded_nonmembers.append(
                            (result, "Middle name mismatch with template")
                        )
                    elif match_type == "fullname_relaxed":
                        print(
                            f"  WARNING: {first} {last} is NOT marked as a "
                            f"member but fuzzy-matches a template entry "
                            f"(name splitting differs) — possible data error.",
                        )
                        miscoded_nonmembers.append(
                            (result, "Name splitting differs from template")
                        )
                elif match_type == "ambiguous":
                    cands = "; ".join(
                        f"{t[TMPL_FIRST]} {t[TMPL_MIDDLE]} {t[TMPL_LAST]} "
                        f"({t[TMPL_EMAIL]})"
                        for t in ambiguous
                    )
                    print(
                        f"  WARNING: {first} {last} is NOT marked as a "
                        f"member but has multiple fuzzy template matches — "
                        f"possible data error. Candidates: {cands}"
                    )
                    miscoded_nonmembers.append(
                        (result, f"Fuzzy matched: {cands}")
                    )
                else:
                    # No template match — normal non-member.
                    output_rows.append(result)

    if skipped:
        print(f"Skipped {skipped} rows due to warnings (see above).")

    # -- Identify unmatched template members -----------------------------------
    unmatched_template = [
        tmpl
        for tmpl in template_members
        if _match_key(
            tmpl[TMPL_FIRST], tmpl[TMPL_MIDDLE], tmpl[TMPL_LAST], tmpl[TMPL_EMAIL]
        )
        not in matched_template_keys
    ]

    # -- Stats -----------------------------------------------------------------
    print_stats(
        output_rows,
        unmatched_input=unmatched_input_members,
        miscoded_nonmembers=miscoded_nonmembers,
        matched_members=len(matched_template_keys),
        fuzzy_matched=fuzzy_matched,
        unmatched_template_count=len(unmatched_template),
    )

    # -- Write outputs ---------------------------------------------------------
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(output_rows)
    print(f"Wrote {len(output_rows)} rows to {output_path}")

    if unmatched_input_members or unmatched_template or miscoded_nonmembers:
        write_manual_review(
            review_path, unmatched_input_members, unmatched_template,
            miscoded_nonmembers,
        )
        print(f"Wrote manual review workbook to {review_path}")
    else:
        print("All members matched — no manual review needed.")


if __name__ == "__main__":
    main()
