"""
Weekly supervisor reports: Excel -> Word

Creates ONE Word document per supervisor. Each week (each sheet in the
workbook, e.g. "Aug 10 to 16") becomes ONE page with:
    - supervisor name, centered
    - the week
    - a table: Service level | Activity title | JIRA ID | Total hours
      grouped into Basic / Intermediate / Senior

Engineer levels come from a small Excel file you maintain by hand
(engineer_levels.xlsx with columns: Name | Level).

The columns, the words used to find them and the Word table can be changed
in the window (report_gui.pyw); this script uses the default "Weekly" layout.
The logic lives in report_engine.py.

Usage:
    pip install pandas openpyxl python-docx
    python make_reports.py hours.xlsx
    python make_reports.py hours.xlsx --levels engineer_levels.xlsx --out reports
"""

import argparse
import sys

from report_engine import ReportError, build_reports, default_profile


def main():
    ap = argparse.ArgumentParser(description="Build weekly supervisor reports.")
    ap.add_argument("hours", help="Excel workbook with one sheet per week")
    ap.add_argument("--levels", default="engineer_levels.xlsx",
                    help="Excel file with columns Name | Level")
    ap.add_argument("--out", default="reports", help="Output folder")
    args = ap.parse_args()

    profile = default_profile("Weekly")
    profile.update(hours_file=args.hours, levels_file=args.levels, out_dir=args.out)
    try:
        build_reports(profile)
    except ReportError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
