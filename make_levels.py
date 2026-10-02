"""
Build / update engineer_levels.xlsx from the hours workbook.

Pulls every engineer name out of all the weekly sheets, so the names in the
levels file are spelled exactly as in the hours file. You then only have to
pick a level for each one (the Level column has a Basic / Intermediate /
Senior dropdown).

If the levels file already exists, levels you already filled in are kept and
only new names are added, so it's safe to run again every week.

Usage:
    pip install pandas openpyxl
    python make_levels.py hours.xlsx
    python make_levels.py hours.xlsx --levels engineer_levels.xlsx
"""

import argparse
import difflib
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from make_reports import (LEVELS, COL_NAME, COL_SUPERVISOR,
                          norm, clean, read_week_sheet)


def collect_names(hours_path):
    """Return {normalised name: (spelling, [supervisors])} from every sheet."""
    xls = pd.ExcelFile(hours_path)
    people = {}
    for sheet in xls.sheet_names:
        # Only names and supervisors are needed, so this also works on the
        # monthly workbook (which has no "Activity title" column)
        df, problem = read_week_sheet(xls, sheet, needed=(COL_NAME, COL_SUPERVISOR))
        if df is None:
            print(f"  - Skipping sheet '{sheet}':\n      {problem}")
            continue
        for _, row in df.iterrows():
            name = clean(row[COL_NAME])
            if not name:
                continue
            spelling, supervisors = people.setdefault(norm(name), (" ".join(name.split()), []))
            supervisor = clean(row[COL_SUPERVISOR])
            if supervisor and supervisor not in supervisors:
                supervisors.append(supervisor)
    return people


def read_existing(path):
    """Return {normalised name: (spelling, level)} from an existing levels file."""
    if not path.exists():
        return {}
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    if "Name" not in df.columns or "Level" not in df.columns:
        sys.exit(f"'{path}' must have the columns 'Name' and 'Level'.")
    existing = {}
    for _, row in df.dropna(subset=["Name"]).iterrows():
        existing[norm(row["Name"])] = (clean(row["Name"]), clean(row["Level"]))
    return existing


def warn_similar(names):
    """Point out names that look like typos of each other (e.g. Jon / John)."""
    names = sorted(names, key=str.lower)
    lowered = [n.lower() for n in names]
    pairs = []
    for i, a in enumerate(lowered):
        for j in range(i + 1, len(lowered)):
            if difflib.SequenceMatcher(None, a, lowered[j]).ratio() >= 0.85:
                pairs.append((names[i], names[j]))
    if pairs:
        print("\n  ! These names look very similar - check the hours file for typos:")
        for a, b in pairs:
            print(f"      - '{a}'  vs  '{b}'")


def write_levels(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Levels"
    ws.append(["Name", "Level", "Supervisor(s)"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E2F3")
    for row in rows:
        ws.append(row)

    # Dropdown so the level can't be misspelled either
    dv = DataValidation(type="list", formula1=f'"{",".join(LEVELS)}"',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Invalid level",
                        error="Pick Basic, Intermediate or Senior.")
    ws.add_data_validation(dv)
    dv.add(f"B2:B{max(len(rows) + 1, 2) + 200}")   # room for rows added by hand

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 45
    ws.freeze_panes = "A2"

    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"Can't write '{path}' - close it in Excel and run again.")


def main():
    ap = argparse.ArgumentParser(description="Create/update the engineer levels file.")
    ap.add_argument("hours", help="Excel workbook with one sheet per week")
    ap.add_argument("--levels", default="engineer_levels.xlsx",
                    help="Levels file to create or update (Name | Level)")
    args = ap.parse_args()

    levels_path = Path(args.levels)
    people = collect_names(args.hours)
    if not people:
        sys.exit("No engineer names found. Check the sheet layout.")
    existing = read_existing(levels_path)

    rows, new, missing_level = [], [], []
    for key, (spelling, supervisors) in people.items():
        old_spelling, level = existing.get(key, (None, ""))
        if old_spelling is None:
            new.append(spelling)
        if level.capitalize() in LEVELS:
            level = level.capitalize()
        else:
            missing_level.append(spelling)
        rows.append([spelling, level, ", ".join(supervisors)])

    # Keep people already in the levels file even if they're not in this hours file
    for key, (spelling, level) in existing.items():
        if key not in people:
            rows.append([spelling, level, ""])

    rows.sort(key=lambda r: r[0].lower())
    write_levels(levels_path, rows)

    print(f"  ✓ {levels_path}  ({len(rows)} engineer(s), {len(new)} new)")
    if new and existing:
        print("\n  New names added:")
        for n in sorted(new, key=str.lower):
            print(f"      - {n}")
    if missing_level:
        print(f"\n  ! {len(missing_level)} engineer(s) still need a level - "
              f"open {levels_path.name} and fill in the Level column.")
    warn_similar([s for s, _ in people.values()])


if __name__ == "__main__":
    main()
