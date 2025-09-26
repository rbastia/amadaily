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

def process_timesheet(timesheet_path, output_dir="output"):
    """Parse the Timesheet and write outputs to `output_dir`.

    This function is intentionally simple and documented so a beginner can follow
    the steps.
    """

    # 1) Ensure the output folder exists
    os.makedirs(output_dir, exist_ok=True)

    # 2) Read the 'Timesheet' sheet into a DataFrame. We use header=None because
    #    the Excel layout does not have a clean single header row (it has merged
    #    cells and repeated blocks across columns).
    df = pd.read_excel(timesheet_path, sheet_name="Timesheet", header=None)

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
        # Within each block, locate 'Job' and 'H' (hours) columns by looking a few
        # columns to the right of the Trk column (this matches the example layout).
        job_col = None
        hours_col = None
        for offset in range(0, 7):
            c = day_col + offset
            if c >= len(header):
                break
            label = str(header[c]).strip().lower()
            if label.startswith('job'):
                job_col = c
            if label == 'h' or label == 'hours':
                hours_col = c

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
            job = df.iat[r, job_col] if job_col is not None and job_col < df.shape[1] else None
            hours = df.iat[r, hours_col] if hours_col is not None and hours_col < df.shape[1] else None

            # Convert hours to float when possible
            hrs = 0.0
            if pd.notna(hours) and str(hours).strip() != '':
                m = re.search(r"(\d+(?:\.\d+)?)", str(hours))
                if m:
                    try:
                        hrs = float(m.group(1))
                    except Exception:
                        hrs = 0.0

            # Skip rows with no job or zero hours
            if pd.isna(job) or not str(job).strip() or hrs == 0.0:
                continue

            records.append({
                'Employee': emp,
                'Date': the_date,
                'Job': str(job).strip(),
                'Hours': hrs,
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
            TotalHours=('Hours', 'sum'),
            Employees=('Employee', lambda x: ', '.join(sorted(set(x))))
        )
        .reset_index()
    )

    out_csv = os.path.join(output_dir, 'timesheet_long_parsed.csv')
    out_xlsx = os.path.join(output_dir, 'timesheet_daily_summary.xlsx')

    long_df.to_csv(out_csv, index=False)
    summary.to_excel(out_xlsx, index=False)

    return out_xlsx
