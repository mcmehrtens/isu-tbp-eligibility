# TBP Eligibility

Scripts for preparing the Tau Beta Pi (TBP) **Report of Eligibility** for the
Iowa Alpha chapter (Iowa State University) from the university-provided
eligibility spreadsheets.

- `prep_eligibility_data.py` — turns the ISU enrollment workbook into a CSV
  formatted for the TBP HQ eligibility template, cross-referencing current
  members against the template's preloaded member data. Also produces a
  manual-review workbook for anything it can't resolve automatically.
- `compare_reports.py` — diffs the member rows of two filled-in TBP report
  workbooks (e.g. one generated from this CSV vs. one done by hand).
- `split_by_department.py` — splits the eligible non-members into
  per-department lists (an Excel tab and a Net-ID CSV per department) for
  sending invitations. Also keeps `Major-Department Reference.xlsx` (in this
  repo) up to date with every major seen and its department.

> **Student data is never committed.** `.gitignore` excludes all spreadsheets,
> CSVs, and documents, plus a `data/` directory. Keep inputs and outputs there.
> The one exception is `Major-Department Reference.xlsx`, which holds no
> student data.

## Semester workflow

1. **Get the ISU eligibility workbook.** The College of Engineering's
   enrollment-research group provides a Power BI report and an accompanying
   Excel workbook with instructions. Follow those instructions to export each
   tab and paste into the workbook; the script reads the `Juniors` and
   `Seniors` tabs. The file is usually password-protected.
2. **Download the TBP template** (`.xls`) from the TBP HQ chapter site. Its
   hidden `CurrentMembers` sheet holds HQ's authoritative record of current
   members and must be used verbatim for those students.
3. **Run the script** (below). Put the inputs in `data/`.
4. **Paste the CSV** into the template's `TBP-Report of Eligibility` sheet.
5. **Work through the manual-review workbook**, if one was generated, adding
   the remaining members by hand.
6. Submit the template to HQ.

## Setup

Requires [uv](https://docs.astral.sh/uv/). Python 3.12+ is installed
automatically.

```
uv sync
```

## Usage

```
uv run prep_eligibility_data.py <semester> <input.xlsx> <template.xls> [-o output.csv]
```

| Argument | Description |
|---|---|
| `semester` | Current semester: `S2026` (Spring 2026) or `F2026` (Fall 2026) |
| `input.xlsx` | ISU eligibility workbook (password prompt if encrypted) |
| `template.xls` | TBP eligibility template from HQ |
| `-o output.csv` | Optional output path (default: `<input_stem>_eligibility.csv`) |

Example:

```
uv run prep_eligibility_data.py S2026 data/Spring2026TBPLists.xlsx data/Eligibility.xls
```

To diff two filled-in reports:

```
uv run compare_reports.py data/report-a.xls data/report-b.xls [-o discrepancies.xlsx]
```

### Outputs

1. **`*_eligibility.csv`** — rows ready to paste into the template. Columns:
   First, Middle, Last, Junior or Senior Class, Month of Graduation,
   Year of Graduation, Curriculum, Present Member, Email Address.
2. **`*_eligibility_manual_review.xlsx`** — generated only when member
   matching is imperfect. One sheet with up to three sections, each with
   instructions and a checkbox column:
   - **Section 1** — input members with no template match. Find them in
     Section 3 (or the TBP member lookup) and add them using HQ's data,
     updating only Junior/Senior. Includes the graduation date computed from
     the ISU data, for students with no Section 3 match.
   - **Section 2** — non-members whose data matches a template entry
     (possible miscoding in the ISU data). Excluded from the CSV pending
     review. Includes the computed graduation date, for adding any that turn
     out to be true non-members.
   - **Section 3** — template members with no input match. Likely graduated,
     or a name/email mismatch with Section 1.

Summary statistics (eligible counts, match results, duplicate emails,
per-curriculum breakdown) and all fuzzy matches are printed to stdout.

## Department lists

`split_by_department.py` takes the same ISU eligibility workbook and produces
the lists departments use to reach invitees. It reads all four eligible tabs
(`Juniors`, `Seniors`, `Masters`, `PhD`), so graduate students are included.

```
TBP_PASSWORD=... uv run split_by_department.py <semester> <input.xlsx> \
    [--prior SEMESTER=PATH ...] [-o OUTDIR] [--reference PATH] [--no-encrypt]
```

| Argument | Description |
|---|---|
| `semester` | Current semester (`F2026`); used in output file names |
| `input.xlsx` | ISU eligibility workbook |
| `--prior SEMESTER=PATH` | An earlier semester's workbook, used only to fill in the majors reference (repeatable) |
| `-o OUTDIR` | Output directory for the department workbook and CSVs (default: `data/`) |
| `--reference PATH` | Majors reference workbook to update (default: `Major-Department Reference.xlsx` in this repo) |
| `--no-encrypt` | Don't password-protect the department workbook |

The password comes from `TBP_PASSWORD` if set; otherwise you are prompted
when an input is encrypted.

Example:

```
TBP_PASSWORD=... uv run split_by_department.py F2026 "data/Fall 2026 Lists.xlsx" \
    --prior "S2026=data/Spring2026TBPLists.xlsx" -o data/dept-F2026
```

### Outputs

The first two go in `OUTDIR`; the third is updated in place in the repo.

1. **`<Semester> Non-Member Lists - Separated by Department.xlsx`**: one tab
   per department with Last Name, First Name, Net-ID, and Major, sorted by
   last name. Encrypted with the input workbook's password.
2. **`Department Net-IDs/<Semester> Net-IDs - <DEPT>.csv`**: Net-IDs only, one
   per line, no header, in the same order as the matching tab.
3. **`Major-Department Reference.xlsx`** (repo root, tracked in git): every
   Program of Study seen so far, with its department, the semesters it
   appeared in, and notes explaining non-obvious placements. Each run merges
   the programs from the current and `--prior` workbooks into the existing
   file, so earlier semesters are kept and `--prior` is only needed to
   backfill. Commit the updated file after each semester's run. It is not
   encrypted, since it holds no student data.

### Decision logic

- **Non-member**: both `Tau_Beta_Pi_Grad` and `Tau_Beta_Pi_Undergrad` are
  blank. As a safety net, anyone whose email is on a member tab
  (`Current TBP Members` / `Members`) is also excluded.
- **Net-ID**: the part of the institutional email before `@`.
- **Department**: the program name (the Program of Study text before the
  first comma) is looked up in `PROGRAM_DEPARTMENTS`. Cross-department
  programs are placed as follows. See the notes in the reference workbook
  for sources.
  - Engineering Mechanics → AERE
  - Industrial and Agricultural Technology → ABE
  - Biomedical Engineering → CBE
  - Software Engineering → ECpE
  - Engineering Management and Systems Engineering → IMSE
  - Energy Systems Engineering → ME
- **Unknown programs**: if any program on any tab is not mapped, the script
  lists them and exits without writing anything. Add them to
  `PROGRAM_DEPARTMENTS` and re-run.
- **Duplicates**: a student on more than one tab (e.g. both Masters and PhD)
  appears once, with the majors joined by `; `.

## Decision logic

**Graduation date.** Derived from `Admitted To Academic Period` by advancing 8
fall/spring semesters. Summer admits count as fall. Fall → `Dec`, spring →
`May`. Two input formats are recognized: `ACADEMIC_PERIOD-2023Fall` and
`2025 Spring Semester (01/21/2025-05/16/2025)`. If the computed date is before
the current semester, juniors are assumed to graduate 4 semesters out and
seniors 2 (counting the current one).

**Curriculum.** `Program of Study` (e.g. `Mechanical Engineering, B.S.`) is
mapped to TBP's short names (`Mechanical engg`) via `CURRICULUM_MAP`. Unknown
programs are skipped with a warning. At startup every mapped value is checked
against the template's `CurriculumData` sheet.

**Member matching.** HQ's instructions say preloaded member names must not be
changed, so each input member is matched to the template and the template's
data is used verbatim (only Junior/Senior is updated). Names are compared with
punctuation, whitespace, and case removed; emails are case-insensitive. Tiers,
tried in order, and never combined:

1. **Exact** — first, middle, last, and email all match.
2. **Email-relaxed** — names match; template email is not `@iastate.edu`.
3. **Middle-relaxed** — first, last, and email match; middle differs
   (initial vs. full, present vs. missing).
4. **Full-name-relaxed** — first+middle+last concatenated matches and email
   matches, but the split differs (e.g. first=`Ana Maria`, last=`Silva`
   vs. first=`Ana`, middle=`Maria`, last=`Silva`).

Multiple candidates in a tier → treated as unmatched and sent to review with
the candidates noted.

**Non-member safeguard.** Non-members are also run through the matcher. Any
that match a template entry are flagged as possibly miscoded and excluded from
the CSV (Section 2 of the review workbook).

**Name punctuation.** HQ's system can't handle punctuation in names. For rows
built from ISU data, interior punctuation becomes a space and edge punctuation
is dropped (`O'Brien` → `O Brien`). Template data is never modified.

## What to update each semester

Everything likely to change lives in the `CONFIGURATION` block at the top of
`prep_eligibility_data.py`:

| Constant | Update when |
|---|---|
| `CURRICULUM_MAP` | ISU adds, removes, or renames an engineering program |
| `JUNIOR_SHEET` / `SENIOR_SHEET` | ISU workbook tab names change |
| `COL_*` / `TMPL_*` | ISU or template column headers change |
| `TEMPLATE_MEMBERS_SHEET` | HQ renames the hidden members sheet |
| `TBP_MEMBER_VALUE` | ISU's membership indicator string changes |
| `SEMESTERS_TO_GRADUATION`, `*_GRADUATION_MONTH` | Degree timeline or commencement schedule changes |
| `FALLBACK_SEMESTERS_*` | Past-graduation fallback policy changes |

For `split_by_department.py`, the `CONFIGURATION` block holds:

| Constant | Update when |
|---|---|
| `PROGRAM_DEPARTMENTS` | ISU adds or renames a program, or it moves department (the script stops on unmapped programs) |
| `DEPARTMENTS` | A department is added, renamed, or merged |
| `DEPARTMENT_NOTES` | Rationale/source for a cross-department placement changes |
| `ELIGIBLE_SHEETS` / `MEMBER_SHEETS` | ISU workbook tab names change |
| `COL_*` | ISU column headers change |

`compare_reports.py` reads columns 11–19 of the template's `TBPEligMembers`
sheet; adjust `MEMBER_COLS` if HQ changes the template layout.
