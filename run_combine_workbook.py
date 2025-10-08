"""run_combine_workbook.py

Convenience wrapper so you can call a single Python script with a *source* Excel
workbook that contains BOTH the Timesheet sheet and the Job Sheet ("New Formula
Job Sheet") and get the combined output in the `outputs` folder.

Usage (CLI):
    python run_combine_workbook.py --workbook path\to\Workbook.xlsx \
        --timesheet-sheet Timesheet \
        --jobsheet-sheet "New Formula Job Sheet" \
        --outdir outputs --per-sheet

All arguments except --workbook are optional (defaults match existing tools).

This is just a thin wrapper around:
  - parse_timesheet.process_timesheet
  - parse_job_sheet.process_job_sheet
  - combine_parsers.combine_daily_reports

It produces (inside --outdir):
  - timesheet_daily_summary.xlsx  (intermediate)
  - ex_job_sheet_normalized.xlsx  (intermediate)
  - combined_daily_report_<timestamp>.xlsx  (final combined)

Return code 0 on success; prints the path to the combined report.
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

from parse_timesheet import process_timesheet
from parse_job_sheet import process_job_sheet
from combine_parsers import combine_daily_reports


def combine_workbook(
    workbook_path: str,
    timesheet_sheet: str = "Timesheet",
    jobsheet_sheet: str = "New Formula Job Sheet",
    outdir: str = "outputs",
    per_sheet: bool = True,
    name: str | None = None,
    fuzzy: bool = True,
) -> str:
    if not os.path.exists(workbook_path):
        raise FileNotFoundError(workbook_path)
    os.makedirs(outdir, exist_ok=True)

    # Intermediate output paths (matching existing conventions)
    timesheet_summary_path = os.path.join(outdir, "timesheet_daily_summary.xlsx")
    job_sheet_norm_path = os.path.join(outdir, "ex_job_sheet_normalized.xlsx")

    # Run the individual parsers
    process_timesheet(workbook_path, output_dir=outdir, sheet_name=timesheet_sheet)
    process_job_sheet(workbook_path, sheet_name=jobsheet_sheet, output_path=job_sheet_norm_path)

    if name is None:
        name = f"combined_daily_report_{datetime.now():%Y%m%d_%H%M%S}"

    combined_path = combine_daily_reports(
        timesheet_summary_path=timesheet_summary_path,
        job_sheet_table_path=job_sheet_norm_path,
        output_dir=outdir,
        output_basename=name,
        write_csv=True,
        fuzzy=fuzzy,
        per_sheet=per_sheet,
    )
    return combined_path


def _parse_args(argv=None):  # pragma: no cover (thin wrapper)
    p = argparse.ArgumentParser(description="Combine a single workbook (Timesheet + Job Sheet) into a consolidated report")
    p.add_argument("--workbook", required=True, help="Path to source Excel workbook containing both sheets")
    p.add_argument("--timesheet-sheet", default="Timesheet", help="Sheet name for timesheet (default: Timesheet)")
    p.add_argument("--jobsheet-sheet", default="New Formula Job Sheet", help="Sheet name for job sheet (default: New Formula Job Sheet)")
    p.add_argument("--outdir", default="outputs", help="Output directory (default: outputs)")
    p.add_argument("--name", help="Optional base name for combined output (timestamp used if omitted)")
    p.add_argument("--no-fuzzy", action="store_true", help="Disable fuzzy job name reconciliation")
    p.add_argument("--single-sheet", action="store_true", help="Write only a single sheet (no per-job sheets)")
    return p.parse_args(argv)


def main():  # pragma: no cover
    args = _parse_args()
    try:
        out = combine_workbook(
            workbook_path=args.workbook,
            timesheet_sheet=args.timesheet_sheet,
            jobsheet_sheet=args.jobsheet_sheet,
            outdir=args.outdir,
            per_sheet=not args.single_sheet,
            name=args.name,
            fuzzy=not args.no_fuzzy,
        )
        print(f"Combined report written: {out}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":  # pragma: no cover
    main()
