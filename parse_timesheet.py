"""
parse_timesheet.py (very simple, well-commented)
===============================================

This is a minimal parser for the company's Timesheet file. It is written to be
easy to read for someone new to Python and pandas. The parser assumes the
Timesheet layout is consistent (as in the CSV you provided) and tolerates empty
cells.

Outputs:
 - timesheet_long_parsed.csv  (one row per Employee/Date/Job)
 - timesheet_daily_summary.xlsx (grouped by Date+Job with total hours and employee list)

Usage:
    from parse_timesheet import process_timesheet
    process_timesheet(r"path\to\Timesheet.xlsx", output_dir="outputs")

Dependencies:
    pip install pandas openpyxl python-dateutil
"""

import os
import re
from datetime import datetime
import pandas as pd
from dateutil import parser as dateparser

# ---------------------------------------------------------------------------
# Placeholder / noise job name detection
# ---------------------------------------------------------------------------
# Timesheet exports sometimes contain artifact strings like 'Column8', 'Column13',
# or bare header tokens like 'Job'. We do NOT want to treat those as real jobs.
# We'll filter them out before building the long records list.

_PLACEHOLDER_JOB_REGEX = re.compile(r"^column\s*\d+$", re.IGNORECASE)
_GENERIC_JOB_TOKENS = {"job", "job#", "job number", "column", "col"}

def _is_placeholder_job(value: str | None) -> bool:
    if value is None:
        return True  # treat missing like placeholder so we skip it
    s = str(value).strip()
    if not s:
        return True
    if _PLACEHOLDER_JOB_REGEX.fullmatch(s):
        return True
    if s.lower() in _GENERIC_JOB_TOKENS:
        return True
    return False

def process_timesheet(timesheet_path, output_dir="output", sheet_name: str = "Timesheet", include_driving: bool = False):
    """Parse the Timesheet and write outputs to `output_dir`.

    This function is intentionally simple and documented so a beginner can follow
    the steps.

    Parameters:
        timesheet_path (str): Path to the Timesheet Excel file.
        output_dir (str): Folder for output CSVs.
        sheet_name (str): Sheet name inside the workbook.
        include_driving (bool): If True, add driving (D column) hours into Hours/Total.
            If False (default), driving hours are ignored (preserves original behavior
            and schema). No additional columns are added to outputs either way.
    """

    # 1) Ensure the output folder exists
    os.makedirs(output_dir, exist_ok=True)

    # 2) Read the timesheet into a DataFrame. Support either Excel or CSV input.
    #    We use header=None because the layout does not have a clean single
    #    header row (it has merged cells and repeated blocks across columns).
    lower_path = str(timesheet_path).lower()
    if lower_path.endswith('.csv'):
        # CSV doesn't have sheets; read as plain CSV
        df = pd.read_csv(timesheet_path, header=None, dtype=object)
    else:
        df = pd.read_excel(timesheet_path, sheet_name=sheet_name, header=None)

    # 3) Find the header row index that contains the word 'Employee'. We look at
    #    the first 15 rows for safety.
    header_row_idx = None
    for i in range(min(15, df.shape[0])):
        # Convert every cell in the row to a string and check for 'employee'
        row_text = df.iloc[i].astype(str).fillna("").str.strip().str.lower()
        if any(cell.startswith('employee') for cell in row_text if cell):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise RuntimeError("Couldn't find header row containing 'Employee'.")

    # 4) Read the header labels (as strings) for each column
    header = df.iloc[header_row_idx].astype(str).fillna("").str.strip().tolist()

    # 5) Detect the starting column for each day's block by finding columns that
    #    include 'Trk' in the header (this matches the example layout).
    day_start_cols = [i for i, h in enumerate(header) if isinstance(h, str) and 'trk' in h.lower()]
    if not day_start_cols:
        raise RuntimeError("Couldn't detect day blocks (no 'Trk' column in header).")

    # 6) Helper to parse a date label (the date is usually in the row above
    #    the header). This helper tries a few strategies to be robust.
    def parse_date_label(cell_value):
        if pd.isna(cell_value):
            return None
        # If Excel stored a real date, pandas will already give us a datetime
        if isinstance(cell_value, (datetime, pd.Timestamp)):
            return pd.to_datetime(cell_value).date()
        text = str(cell_value).strip()
        if not text:
            return None
        # Try pandas first
        dt = pd.to_datetime(text, errors='coerce')
        if pd.notna(dt):
            return dt.date()
        # Try a simple regex like 9/8 or 9-8 or 9-8-25
        m = re.search(r"(\d{1,2})[-/](\d{1,2})(?:[-/](\d{2,4}))?", text)
        if m:
            mon = int(m.group(1))
            day = int(m.group(2))
            yr = m.group(3)
            if yr:
                yr = int(yr)
                if yr < 100:
                    yr += 2000
            else:
                yr = datetime.now().year
            try:
                return datetime(yr, mon, day).date()
            except Exception:
                pass
        # As a last resort use dateutil
        try:
            dt = dateparser.parse(text, default=datetime(datetime.now().year, 1, 1))
            return dt.date() if dt else None
        except Exception:
            return None

    # The row just above header often contains the date labels
    date_label_row = header_row_idx - 1 if header_row_idx > 0 else None

    # 7) Find the first employee row (first non-empty cell in column 0 after header)
    first_emp_row = None
    for r in range(header_row_idx + 1, df.shape[0]):
        v = df.iat[r, 0]
        if pd.notna(v) and str(v).strip() and not str(v).strip().lower().startswith('total'):
            first_emp_row = r
            break
    if first_emp_row is None:
        first_emp_row = header_row_idx + 1

    # 8) Iterate day blocks and employee rows to build records
    records = []
    for day_col in day_start_cols:
        # Within each day block, we only want to look at THAT day's columns.
        # A block in the provided layout is exactly 5 columns wide:
        #   Trk # | Job | H | W/S | D
        # Previously we scanned up to 7 columns which allowed us to "see"
        # the next day's 'Job' header and overwrite job_col with the following
        # day's Job column. That caused most Friday (for example) jobs to be
        # skipped unless the employee also had a Saturday entry (because we
        # were pairing Friday hours with Saturday job cells that were blank).
        #
        # Fix: restrict the scan strictly to the current block width (5 cols),
        # or stop early if we encounter another 'trk' header which signals the
        # next day has started.
        BLOCK_WIDTH = 5  # conservative; adjust if layout changes in future
        job_col = None
        hours_col = None  # 'H' column (work hours)
        driving_col = None  # 'D' column (driving hours)
        for offset in range(BLOCK_WIDTH):
            c = day_col + offset
            if c >= len(header):
                break
            label = str(header[c]).strip().lower()
            # If (unlikely) we hit another 'trk' within the expected width (layout change), stop.
            if offset > 0 and label.startswith('trk'):
                break
            if label.startswith('job') and job_col is None:  # first match within block
                job_col = c
            # accept headers like 'H', 'hours', or combined tokens like 'H,W/S,D'
            if label.startswith('h') and hours_col is None:
                hours_col = c
            # driving column often just 'd'
            if (label == 'd' or label.startswith('drive')) and driving_col is None:
                driving_col = c

        # If we somehow didn't find required base columns, skip this day block gracefully
        if job_col is None or hours_col is None:
            continue

        # Parse the date for this day from the row above the header
        date_val = df.iat[date_label_row, day_col] if date_label_row is not None else None
        the_date = parse_date_label(date_val)
        if the_date is None:
            # Try neighboring columns if date not found in the expected spot
            for adj in (-1, 1, 2, -2):
                c = day_col + adj
                if 0 <= c < df.shape[1]:
                    the_date = parse_date_label(df.iat[date_label_row, c])
                    if the_date:
                        break

        # Now iterate employees
        for r in range(first_emp_row, df.shape[0]):
            emp_cell = df.iat[r, 0]
            if pd.isna(emp_cell):
                continue
            emp = str(emp_cell).strip()
            if not emp or emp.lower().startswith('total'):
                continue

            # Read job and hours safely (columns may be missing or empty)
            job_cell = df.iat[r, job_col] if job_col is not None and job_col < df.shape[1] else None
            hours = df.iat[r, hours_col] if hours_col is not None and hours_col < df.shape[1] else None
            driving = df.iat[r, driving_col] if driving_col is not None and driving_col < df.shape[1] else None

            # Normalize job text early; skip placeholder artifacts
            job = None
            if job_cell is not None and not pd.isna(job_cell):
                job_text = str(job_cell).strip()
                if job_text and not _is_placeholder_job(job_text):
                    job = job_text

            # Convert hours to float when possible
            work_hrs = 0.0  # base work (H) hours
            if pd.notna(hours) and str(hours).strip() != '':
                m = re.search(r"(\d+(?:\.\d+)?)", str(hours))
                if m:
                    try:
                        work_hrs = float(m.group(1))
                    except Exception:
                        work_hrs = 0.0

            # Always parse driving hours when present so we can expose them in the
            # output DrivingHours column. Whether driving is added into the
            # 'Hours' field is controlled by include_driving (below).
            drive_hrs = 0.0
            if driving_col is not None and pd.notna(driving) and str(driving).strip() != '':
                m2 = re.search(r"(\d+(?:\.\d+)?)", str(driving))
                if m2:
                    try:
                        drive_hrs = float(m2.group(1))
                    except Exception:
                        drive_hrs = 0.0

            effective_hours = work_hrs + (drive_hrs if include_driving else 0.0)

            # Include rows that have either work hours or driving hours. The
            # 'Hours' field will include driving only when include_driving=True.
            if job is None or (work_hrs == 0.0 and drive_hrs == 0.0):
                continue

            records.append({
                'Employee': emp,
                'Date': the_date,
                'Job': str(job).strip(),
                'Hours': effective_hours,  # remains singular; optionally includes driving
                'DrivingHours': drive_hrs,
                'Trk#': df.iat[r, day_col] if day_col < df.shape[1] else None,
            })

    # 9) Convert to DataFrame and save outputs
    if not records:
        raise RuntimeError('No records parsed. Check the sheet layout and header detection.')

    long_df = pd.DataFrame.from_records(records)
    long_df['Date'] = pd.to_datetime(long_df['Date']).dt.date

    summary = (
        long_df
        .groupby(['Date', 'Job'], dropna=False)
        .agg(
            EmployeeCount=('Employee', lambda x: x.nunique()),
            TotalHours=('Hours', 'sum'),  # Hours may or may not include driving based on flag
            TotalDrivingHours=('DrivingHours', 'sum'),
            Employees=('Employee', lambda x: ', '.join(sorted(set(x))))
        )
        .reset_index()
    )

    out_csv = os.path.join(output_dir, 'timesheet_long_parsed.csv')
    out_xlsx = os.path.join(output_dir, 'timesheet_daily_summary.xlsx')

    long_df.to_csv(out_csv, index=False)
    # Safe write for summary: handle case where file is open (Windows Excel lock)
    try:
        summary.to_excel(out_xlsx, index=False)
    except PermissionError:
        # If XLSX is locked, write CSV instead and notify via returned path
        base, ext = os.path.splitext(out_xlsx)
        alt_xlsx = base + '_new' + ext
        try:
            summary.to_excel(alt_xlsx, index=False)
            out_xlsx = alt_xlsx
        except PermissionError:
            out_csv_summary = base + '_daily_summary.csv'
            summary.to_csv(out_csv_summary, index=False)
            out_xlsx = out_csv_summary

    # out_csv = os.path.join(output_dir, 'timesheet_long_parsed.csv')   
    # out_xlsx = os.path.join(output_dir, 'timesheet_daily_summary.csv')

    # long_df.to_csv(out_csv, index=False)

    # Safe write for summary: handle case where file is open (Windows Excel lock)
    # try:
    #     summary.to_csv(out_xlsx, index=False)
    # except PermissionError:
    #     base, ext = os.path.splitext(out_xlsx)
    #     alt_path = base + '_new' + ext
    #     summary.to_csv(alt_path, index=False)
    #     out_xlsx = alt_path

    return out_xlsx
