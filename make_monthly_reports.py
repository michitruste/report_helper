"""
Monthly supervisor reports: Excel -> Word

Same idea as make_reports.py, for the monthly workbook layout:
    Date | Name | CSID | Ford TL | Ford Supervisor | Feature name |
    JIRA ID/Test Rail Test Run | Activity Description |
    TC executed, TC created/Updated, ... | Program | (anything else is ignored)

Creates ONE Word document per supervisor with:
    - the supervisor name in a shaded box at the top
    - for every week: "Week 39: 01/09/2026 - 06/09/2026  - Supervisor: <name>"
      and a table with one row each for Basic / Intermediate / Senior

Excel column               -> Word column
    Activity Description       -> Activity Title
    JIRA ID/Test Rail Test Run -> JIRA ID/TMT Request ID/Test Rail Test Run
    TC executed, ...           -> Activity Description
    TC executed, ...           -> TC executed, ...
    Program                    -> Vehicle Program
    (none)                     -> Total Amount of hours (left blank for now)

Weeks come from the Date column, not from the sheet names: Monday to Sunday,
cut at the start and end of the month (so September 2026 gives 01/09 - 06/09,
07/09 - 13/09, ..., 28/09 - 30/09). Week numbers are ISO week numbers unless
--first-week is given, which numbers the first week N and counts up from there.
A sheet without a Date column is treated as one week named after the sheet,
like in make_reports.py.

Engineer levels come from the same engineer_levels.xlsx as make_reports.py
(make_levels.py also works on the monthly workbook).

The columns, the words used to find them and the Word table can be changed
in the window (report_gui.pyw); this script uses the default "Monthly" layout.
The logic lives in report_engine.py.

Usage:
    pip install pandas openpyxl python-docx
    python make_monthly_reports.py monthly.xlsx
    python make_monthly_reports.py monthly.xlsx --levels engineer_levels.xlsx --out monthly_reports
    python make_monthly_reports.py monthly.xlsx --first-week 39
"""

import argparse
import sys

from report_engine import ReportError, build_reports, default_profile


def main():
    ap = argparse.ArgumentParser(description="Build monthly supervisor reports.")
    ap.add_argument("hours", help="Monthly Excel workbook")
    ap.add_argument("--levels", default="engineer_levels.xlsx",
                    help="Excel file with columns Name | Level")
    ap.add_argument("--out", default="monthly_reports", help="Output folder")
    ap.add_argument("--first-week", type=int,
                    help="Number of the first week (e.g. 39); the rest count up "
                         "from it. Default: ISO week numbers")
    args = ap.parse_args()

    profile = default_profile("Monthly")
    profile.update(hours_file=args.hours, levels_file=args.levels,
                   out_dir=args.out, first_week=args.first_week)
    try:
        build_reports(profile)
    except ReportError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
