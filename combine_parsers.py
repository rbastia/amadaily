"""combine_parsers.py

Call both parsers (timesheet and job sheet), merge their outputs by Date+Job,
and write a combined table to outputs.

Usage:
  from combine_parsers import combine_parsers
  out = combine_parsers(r"path\to\Timesheet.xlsx", r"path\to\job_sheet.xlsx", output_dir="outputs")
"""
from __future__ import annotations

import os
from typing import Optional

import pandas as pd
from datetime import datetime
import shutil

# import the existing parser functions
from parse_timesheet import process_timesheet
from parse_job_sheet import process_job_sheet, process_job_sheet_file


def _normalize_job_name(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    s2 = str(s).strip()
    if not s2:
        return None
    # normalize to lower-case, remove punctuation (keep letters/numbers/spaces),
    # collapse multiple internal spaces
    import re
    s2 = s2.lower()
    # Replace common separators and punctuation with space
    s2 = re.sub(r"[^a-z0-9 ]+", " ", s2)
    s2 = " ".join(s2.split())
    return s2


def _is_noise_job(s: Optional[str]) -> bool:
    """Return True for job names that are likely noise/artifacts (e.g. Column1).

    This helps filter out header-like or placeholder job names that appear in
    the timesheet parsing but aren't meaningful jobs to merge on.
    """
    if s is None:
        return True
    t = str(s).strip().lower()
    if not t:
        return True
    import re
    # column followed by digits (Column8, Column1) -> noise
    if re.match(r"^column\s*\d+$", t):
        return True
    # generic single-word placeholders
    if t in ("column", "col", "trk#", "trk"):
        return True
    # purely numeric job names are probably noise
    if re.fullmatch(r"\d+", t):
        return True
    return False


def _ensure_date_col(df: pd.DataFrame, col: str) -> pd.Series:
    """Return a Series of python.date objects for the given column."""
    if col not in df.columns:
        return pd.Series([pd.NaT] * len(df))
    ser = pd.to_datetime(df[col], errors="coerce")
    # keep only date portion
    return ser.dt.date


def combine_parsers(timesheet_path: str, job_sheet_path: str, output_dir: str = "outputs", sheet_name: str = "New Formula Job Sheet") -> str:
    """Run both parsers and merge their outputs on Date+Job.

    Writes two files into output_dir:
      - combined_daily_report.xlsx
      - combined_daily_report.csv

    Returns the Excel output path.
    """
    os.makedirs(output_dir, exist_ok=True)

    # 1) Run timesheet parser which writes outputs into output_dir
    # process_timesheet writes timesheet_long_parsed.csv and timesheet_daily_summary.xlsx
    timesheet_summary_path = os.path.join(output_dir, "timesheet_daily_summary.xlsx")

    try:
        process_timesheet(timesheet_path, output_dir=output_dir)
        # if successful, the summary should now exist at timesheet_summary_path
        if not os.path.exists(timesheet_summary_path):
            raise FileNotFoundError(f"Expected timesheet summary at {timesheet_summary_path}")
    except PermissionError as e:
        # Common on Windows when the target Excel file is open. Fallback: write into
        # a timestamped temp subdirectory and continue from there.
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        alt_dir = os.path.join(output_dir, f'_tmp_timesheet_{ts}')
        print(f"PermissionError writing to {timesheet_summary_path}: {e}. Retrying in {alt_dir}.")
        os.makedirs(alt_dir, exist_ok=True)
        process_timesheet(timesheet_path, output_dir=alt_dir)
        timesheet_summary_path = os.path.join(alt_dir, "timesheet_daily_summary.xlsx")
        if not os.path.exists(timesheet_summary_path):
            raise FileNotFoundError(f"Expected timesheet summary at {timesheet_summary_path} after fallback")

    timesheet_df = pd.read_excel(timesheet_summary_path)

    # 2) Run job sheet parser - we can call lower-level function to get DataFrame
    # process_job_sheet returns a DataFrame; we won't rely on the file written by process_job_sheet_file
    job_df = process_job_sheet(job_sheet_path, sheet_name=sheet_name)

    # 3) Normalize Date and Job columns on both frames so they match
    timesheet_df = timesheet_df.copy()
    job_df = job_df.copy()

    # Ensure Date columns are date objects
    timesheet_df['Date'] = _ensure_date_col(timesheet_df, 'Date')
    job_df['Date'] = _ensure_date_col(job_df, 'Date')

    # Normalize Job names (strip and collapse spaces)
    timesheet_df['Job_norm'] = timesheet_df['Job'].apply(_normalize_job_name)
    job_df['Job_norm'] = job_df['Job'].apply(_normalize_job_name)

    # Merge on Date and normalized Job name. Use outer to keep rows from both.
    merged = pd.merge(
        timesheet_df,
        job_df,
        left_on=['Date', 'Job_norm'],
        right_on=['Date', 'Job_norm'],
        how='outer',
        suffixes=('_timesheet', '_job')
    )

    # Drop the helper normalized job column and optionally keep a single Job column
    # Prefer the explicit Job column from job_df when available, otherwise timesheet one
    def pick_job(row):
        if pd.notna(row.get('Job_job')) and str(row.get('Job_job')).strip():
            return row.get('Job_job')
        return row.get('Job_timesheet')

    merged['Job'] = merged.apply(pick_job, axis=1)
    # --- Simple fuzzy matching pass ---
    # For timesheet rows that didn't find a job-sheet match (no Job_job), try a
    # permissive substring match against job_df on the same Date. If there's a
    # single candidate, copy the job-specific columns into the merged row.
    job_cols = [c for c in job_df.columns if c not in ("Date", "Job", "Job_norm")]
    # build lookup by (Date) -> list of (job_norm, row_index)
    job_lookup = {}
    for i, row in job_df.iterrows():
        d = row['Date']
        jn = row.get('Job_norm')
        job_lookup.setdefault(d, []).append((jn, i))

    def try_fuzzy_fill(row):
        # only operate on rows that have timesheet data and no job info
        if pd.notna(row.get('Job_job')):
            return row
        date = row['Date']
        if pd.isna(date):
            return row
        tjn = row.get('Job_norm')
        if not tjn or _is_noise_job(tjn):
            return row
        candidates = job_lookup.get(date, [])
        matches = []
        for jn, idx in candidates:
            if not jn:
                continue
            # prefer exact containment both ways
            if jn in tjn or tjn in jn:
                matches.append(idx)
        if len(matches) == 1:
            # copy job-specific columns from job_df
            src = job_df.loc[matches[0]]
            for c in ['Truck(s)', 'Description', 'Concrete', 'Concrete Yds', 'Stone', 'Stone Lds']:
                if c in src.index and (pd.isna(row.get(c)) or not str(row.get(c)).strip()):
                    row[c] = src.get(c)
            # prefer the job text from job sheet
            if pd.notna(src.get('Job')) and str(src.get('Job')).strip():
                row['Job'] = src.get('Job')
        return row

    merged = merged.apply(try_fuzzy_fill, axis=1)
    # Reorder columns: Date, Job, then helpful columns from both
    # Keep timesheet summary columns if present
    cols = ['Date', 'Job']
    # Add timesheet summary columns if they exist
    for c in ['EmployeeCount', 'TotalHours', 'Employees']:
        if c in merged.columns:
            cols.append(c)
    # Add job-specific columns
    for c in ['Truck(s)', 'Description', 'Concrete', 'Concrete Yds', 'Stone', 'Stone Lds']:
        if c in merged.columns:
            cols.append(c)

    # Remove helper/internal columns created by the merge
    drop_cols = [c for c in merged.columns if c.endswith('_timesheet') or c.endswith('_job') or c == 'Job_norm']
    # also explicitly drop Job_timesheet/Job_job if present (we already have unified 'Job')
    for x in ('Job_timesheet', 'Job_job'):
        if x in drop_cols:
            drop_cols.remove(x)
    # But we want to drop them - ensure they are in the list
    drop_cols = list(set(drop_cols) | set([c for c in ('Job_timesheet', 'Job_job', 'Job_norm') if c in merged.columns]))
    out_df = merged.drop(columns=drop_cols, errors='ignore')
    # Reorder columns: Date, Job, then timesheet summary columns if present, then job-specific cols
    ordered = ['Date', 'Job']
    for c in ['EmployeeCount', 'TotalHours', 'Employees']:
        if c in out_df.columns:
            ordered.append(c)
    for c in ['Truck(s)', 'Description', 'Concrete', 'Concrete Yds', 'Stone', 'Stone Lds']:
        if c in out_df.columns:
            ordered.append(c)
    # append any remaining columns
    remaining = [c for c in out_df.columns if c not in ordered]
    out_df = out_df[ordered + remaining]

    out_xlsx = os.path.join(output_dir, 'combined_daily_report.xlsx')
    out_csv = os.path.join(output_dir, 'combined_daily_report.csv')

    try:
        out_df.to_excel(out_xlsx, index=False)
        out_df.to_csv(out_csv, index=False)
        return out_xlsx
    except PermissionError as e:
        # If the target file is locked (common on Windows when open in Excel),
        # fall back to a timestamped filename in the same output directory.
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        alt_xlsx = os.path.join(output_dir, f'combined_daily_report_{ts}.xlsx')
        alt_csv = os.path.join(output_dir, f'combined_daily_report_{ts}.csv')
        print(f"PermissionError writing combined report to {out_xlsx}: {e}."
              f" Writing to {alt_xlsx} instead.")
        out_df.to_excel(alt_xlsx, index=False)
        out_df.to_csv(alt_csv, index=False)
        return alt_xlsx


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Combine timesheet and job sheet outputs into one merged report')
    parser.add_argument('timesheet', help='Path to Timesheet.xlsx')
    parser.add_argument('job_sheet', help='Path to job sheet .xlsx or csv')
    parser.add_argument('-o', '--output_dir', default='outputs')
    args = parser.parse_args()
    print(combine_parsers(args.timesheet, args.job_sheet, output_dir=args.output_dir))
