"""
Monthly supervisor reports: Excel -> Word

Same idea as make_reports.py, for the monthly workbook layout:
    Date | Name | CSID | Ford TL | Ford Supervisor | Feature name |
    JIRA ID/Test Rail Test Run | Activity Description |
    TC executed, TC created/Updated, ... | Program | (anything else is ignored)

Creates ONE Word document per supervisor with:
    - the supervisor name in a shaded box at the top
    - for every week: "Week 36 (September 1 – September 6):  - Supervisor: <name>"
      and a table with one row each for Basic / Intermediate / Senior

Excel column               -> Word column
    Activity Description       -> Activity Title
    JIRA ID/Test Rail Test Run -> JIRA ID/TMT Request ID/Test Rail Test Run
    TC executed, ...           -> Activity Description
    TC executed, ...           -> TC executed, ...
    Program                    -> Vehicle Program
    (none)                     -> Total Amount of hours (left blank for now)

Weeks come from the Date column (ISO week number, Monday to Saturday, or to
Sunday if someone logged a Sunday). A sheet without a Date column is treated as
one week named after the sheet, like in make_reports.py.

Engineer levels come from the same engineer_levels.xlsx as make_reports.py
(make_levels.py also works on the monthly workbook).

Usage:
    pip install pandas openpyxl python-docx
    python make_monthly_reports.py monthly.xlsx
    python make_monthly_reports.py monthly.xlsx --levels engineer_levels.xlsx --out monthly_reports
"""

import argparse
import re
import sys
from datetime import timedelta
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

from make_reports import (LEVELS, COL_NAME, COL_SUPERVISOR, COL_JIRA,
                          norm, clean, load_levels, read_week_sheet,
                          set_cell_shading, write_cell, write_bullets)

HEADERS = ["Service Level", "Activity Title",
           "JIRA ID/TMT Request ID/Test Rail Test Run", "Activity Description",
           "TC executed, TC created/Updated, System Issues Retested/Reported, "
           "Test Procedure issues resolved, Deliverables completed",
           "Vehicle Program", "Total Amount of hours"]
COL_WIDTHS = [Cm(2.6), Cm(3.1), Cm(2.5), Cm(2.6), Cm(3.1), Cm(2.6), Cm(2.0)]
TITLE_FILL = "F6C5AC"

# Columns we need from the monthly sheet (plus Name / Supervisor / JIRA ID)
COL_DATE = "Date"
COL_DESCRIPTION = "Activity Description"
COL_TC = "TC executed"
COL_PROGRAM = "Program"
# Excel column feeding each Word column after "Service Level", in order
# (None = left blank). Total Amount of hours is always blank.
ENTRY_COLS = [COL_DESCRIPTION,   # -> Activity Title
              COL_JIRA,          # -> JIRA ID/TMT Request ID/Test Rail Test Run
              COL_TC,            # -> Activity Description
              COL_TC,            # -> TC executed, ...
              COL_PROGRAM]       # -> Vehicle Program
# dict.fromkeys: each Excel column once, even if it feeds two Word columns
NEEDED = list(dict.fromkeys([COL_NAME, COL_SUPERVISOR] + [c for c in ENTRY_COLS if c]))


# ----------------------------------------------------------------- helpers
def find_columns(header_cells):
    """
    Work out which column is which in the monthly layout, tolerating small
    wording differences. Returns {standard name: column index}.
    """
    cells = [norm(c) for c in header_cells]

    def first(*matches):
        for match in matches:                  # try the strictest match first
            for idx, c in enumerate(cells):
                if c and match(c):
                    return idx
        return None

    others = ("supervisor", "program", "feature", "activity", "model")
    found = {
        # \b so "TC created/Updated" doesn't count as a date column
        COL_DATE: first(lambda c: c == "date", lambda c: re.search(r"\bdate\b", c)),
        COL_NAME: first(lambda c: c == "name",
                        lambda c: "name" in c and not any(o in c for o in others)),
        COL_SUPERVISOR: first(lambda c: "supervisor" in c),
        COL_JIRA: first(lambda c: "jira" in c),
        COL_DESCRIPTION: first(lambda c: "description" in c),
        COL_TC: first(lambda c: c.startswith("tc executed"),
                      lambda c: "tc executed" in c),
        COL_PROGRAM: first(lambda c: c == "program", lambda c: "program" in c),
    }
    return {k: v for k, v in found.items() if v is not None}


def to_date(value):
    """Excel date cell (or text / serial number) -> date, or None."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if pd.isna(value):
            return None
        value = pd.Timestamp("1899-12-30") + pd.Timedelta(days=value)
    d = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(d) else d.date()


def day(d):
    return f"{d:%B} {d.day}"


def week_label(monday, has_sunday):
    end = monday + timedelta(days=6 if has_sunday else 5)
    return f"Week {monday.isocalendar()[1]} ({day(monday)} – {day(end)})"


# -------------------------------------------------------------- the page
def add_title_box(doc, supervisor):
    box = doc.add_table(rows=1, cols=1)
    box.style = "Table Grid"
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = box.rows[0].cells[0]
    cell.width = Cm(13.75)
    set_cell_shading(cell, TITLE_FILL)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(supervisor)
    r.font.name = "Calibri Light"
    r.font.size = Pt(16)
    doc.add_paragraph()


def repeat_as_header(row):
    """Repeat this row at the top of every page the table runs onto."""
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))


def add_week(doc, supervisor, label, activities_by_level):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(f"{label}:  - ")
    p.add_run(f"Supervisor: {supervisor}").bold = True

    table = doc.add_table(rows=1, cols=len(HEADERS))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h in enumerate(HEADERS):
        write_cell(table.rows[0].cells[i], h)
    repeat_as_header(table.rows[0])

    # One row per level; each column a bullet list, one bullet per activity
    for level in LEVELS:
        items = activities_by_level.get(level, [])
        cells = table.add_row().cells
        write_cell(cells[0], level, bold=True, center=True)
        for col, source in enumerate(ENTRY_COLS):
            if source:
                write_bullets(cells[col + 1], [item[col] for item in items])
            else:
                write_cell(cells[col + 1], "")
        write_cell(cells[-1], "")             # Total hours: left blank for now

    table.autofit = False
    for i, w in enumerate(COL_WIDTHS):
        table.columns[i].width = w
    for row in table.rows:
        for i, w in enumerate(COL_WIDTHS):
            row.cells[i].width = w

    doc.add_paragraph()


# ------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Build monthly supervisor reports.")
    ap.add_argument("hours", help="Monthly Excel workbook")
    ap.add_argument("--levels", default="engineer_levels.xlsx",
                    help="Excel file with columns Name | Level")
    ap.add_argument("--out", default="monthly_reports", help="Output folder")
    args = ap.parse_args()

    levels = load_levels(args.levels)
    xls = pd.ExcelFile(args.hours)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    # data[supervisor][week key][level] = [entry following ENTRY_COLS, ...]
    # week key: (0, monday ordinal) for dated rows, (1, sheet index) for undated sheets
    data, labels = {}, {}
    sunday_weeks = set()
    unassigned = set()
    bad_dates = 0

    for sheet_idx, sheet in enumerate(xls.sheet_names):
        df, problem = read_week_sheet(xls, sheet, needed=NEEDED,
                                      optional=[COL_DATE], finder=find_columns)
        if df is None:
            print(f"  - Skipping sheet '{sheet}':\n      {problem}")
            continue
        dated = COL_DATE in df.columns

        for _, row in df.iterrows():
            supervisor = clean(row[COL_SUPERVISOR])
            name = clean(row[COL_NAME])
            if not supervisor or not name:
                continue
            level = levels.get(norm(name))
            if level is None:
                unassigned.add(name)
                continue
            entry = tuple(clean(row[c]) if c else "" for c in ENTRY_COLS)
            if not any(entry):
                continue

            if dated:
                d = to_date(row[COL_DATE])
                if d is None:
                    bad_dates += 1
                    continue
                monday = d - timedelta(days=d.weekday())
                week = (0, monday.toordinal())
                labels[week] = monday
                if d.weekday() == 6:
                    sunday_weeks.add(week)
            else:
                week = (1, sheet_idx)
                labels[week] = sheet

            bucket = (data.setdefault(supervisor, {})
                          .setdefault(week, {})
                          .setdefault(level, []))
            if entry not in bucket:          # same activity on several days -> once
                bucket.append(entry)

    if not data:
        sys.exit("No rows found. Check the sheet layout and the levels file.")

    for week, value in labels.items():
        if week[0] == 0:
            labels[week] = week_label(value, week in sunday_weeks)

    for supervisor, weeks in data.items():
        doc = Document()
        for section in doc.sections:
            section.top_margin = section.bottom_margin = Cm(2.5)
            section.left_margin = section.right_margin = Cm(1.5)
        style = doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(11)

        add_title_box(doc, supervisor)
        for week in sorted(weeks):
            add_week(doc, supervisor, labels[week], weeks[week])

        safe = re.sub(r'[\\/:*?"<>|]+', "_", supervisor)
        path = out_dir / f"Monthly Report - {safe}.docx"
        try:
            doc.save(path)
        except PermissionError:
            sys.exit(f"Can't write '{path}' - close it in Word and run again.")
        print(f"  ✓ {path}  ({len(weeks)} week(s))")

    if bad_dates:
        print(f"\n  ! {bad_dates} row(s) had an empty or unreadable Date and were left out.")
    if unassigned:
        print("\n  ! These engineers have no level in the levels file and were left out:")
        for n in sorted(unassigned):
            print(f"      - {n}")
        print("    Add them to the levels file (python make_levels.py <workbook>) and run again.")


if __name__ == "__main__":
    main()
