"""
parse_timesheet.py
==================
This script reads a construction company timesheet Excel file, extracts daily
work information (who worked on which job, how many hours, etc.), and produces:
  - A long-format CSV with one row per (Employee, Date, Job, Hours)
  - A summary Excel file that groups by (Date, Job)

Libraries required:
    pip install pandas openpyxl python-dateutil
"""

# --- Imports ---
import pandas as pd                # main data handling library
import os, re                      # os for filenames, re for regex text matching
from datetime import datetime       # for date handling
from dateutil import parser as dateparser   # powerful date parser

# --- CONFIGURATION ---
timesheet_path = "Timesheet 9-7-25 thru 9-13-25.xlsx"  # your Excel file
sheet_name = "Timesheet"                               # name of the worksheet/tab
# ======================


# --- STEP 1: Load the raw sheet into pandas ---
# We use header=None because the Excel sheet does NOT have a simple header row.
# It's got weird merged cells and repeating blocks instead.
df_raw = pd.read_excel(timesheet_path, sheet_name=sheet_name, header=None)


# --- STEP 2: Try to guess the year from the filename ---
# Example filename: "Timesheet 9-7-25 thru 9-13-25.xlsx"
# This helps if the sheet only has dates like "9-7" without a year.
def infer_year_from_filename(path):
    fn = os.path.basename(path)  # just the filename, no folder path
    # Look for patterns like "9-7-25" or "09/07/2025"
    m = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})', fn)
    if m:
        mon, day, yr = m.groups()
        yr = int(yr)
        # If year is two digits (like 25), assume 2000+ (→ 2025)
        if yr < 100:
            yr += 2000
        return yr
    # Fallback: if nothing found, use current year
    return datetime.now().year

default_year = infer_year_from_filename(timesheet_path)


# --- STEP 3: Find the header row ---
# The header row is the one containing "Employee" (usually near the top).
header_row_idx = None
for i in range(0, min(10, df_raw.shape[0])):  # only scan first 10 rows
    row_vals = df_raw.iloc[i].astype(str).str.strip().fillna('').tolist()
    # If any cell in this row starts with "employee" (case-insensitive), we found it
    if any(v.lower().startswith('employee') for v in row_vals if v):
        header_row_idx = i
        break

if header_row_idx is None:
    raise RuntimeError("Couldn't find header row containing 'Employee' in first 10 rows.")

# Save the header row as a list of labels
header_row = df_raw.iloc[header_row_idx].astype(str).str.strip().fillna('').tolist()


# --- STEP 4: Detect day-block starts ---
# The timesheet is structured in repeating column groups:
#   Employee | Trk# | Job | H | ...
# We detect where each block starts by finding "Trk #".
day_start_cols = [i for i, v in enumerate(header_row) if v and v.lower().startswith('trk')]
if not day_start_cols:
    # Backup: catch cases where it's written differently like "trk number"
    day_start_cols = [i for i, v in enumerate(header_row) if 'trk' in v.lower()]


# --- STEP 5: Parse a date from the label above the header ---
# For each day-block, the row ABOVE the header has a date like "Monday 9-8".
def parse_date_label(cell_value, default_year=default_year):
    if pd.isna(cell_value):
        return None

    # If Excel stored this as an actual date object, just return it
    if isinstance(cell_value, (datetime, pd.Timestamp)):
        return pd.to_datetime(cell_value).date()

    # Otherwise treat it as text
    s = str(cell_value).strip()

    # Try pandas' date parser first
    try:
        dt = pd.to_datetime(s, errors='coerce')
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass

    # Try a regex like "9-8" or "9-8-25"
    m = re.search(r'(\d{1,2})[-/](\d{1,2})(?:[-/](\d{2,4}))?', s)
    if m:
        mon = int(m.group(1))
        day = int(m.group(2))
        yr = m.group(3)
        if yr:
            yr = int(yr)
            if yr < 100:
                yr += 2000
        else:
            yr = default_year
        try:
            return datetime(yr, mon, day).date()
        except Exception:
            return None

    # Final attempt: use python-dateutil (super flexible)
    try:
        dt = dateparser.parse(s, default=datetime(default_year, 1, 1))
        if dt:
            return dt.date()
    except Exception:
        pass

    return None


# --- STEP 6: Find the first employee row ---
# After the header row, the first real person’s name marks where employees start.
first_employee_row = None
for r in range(header_row_idx+1, df_raw.shape[0]):
    v = df_raw.iloc[r, 0]   # employee name is in first column
    if pd.notna(v):
        s = str(v).strip()
        # Skip blank/total rows; take first row that looks like a real name
        if s and not s.lower().startswith('column') and not s.lower().startswith('total') and len(s) > 1:
            first_employee_row = r
            break

if first_employee_row is None:
    first_employee_row = header_row_idx + 3   # fallback guess


# --- STEP 7: Iterate through employees & day blocks ---
# Build a list of "records" with Employee, Date, Job, Hours.
records = []
num_cols = df_raw.shape[1]

for day_col in day_start_cols:

    # Look for which subcolumns in this block are Job and Hours
    job_col = None
    hours_col = None
    for offset in range(0, 7):  # search up to 7 columns from start of block
        idx = day_col + offset
        if idx >= len(header_row):
            break
        label = header_row[idx].strip().lower()
        if label.startswith('job'):
            job_col = idx
        if label.upper() == 'H' or label.strip().lower() == 'h':
            hours_col = idx

    # Grab the date label above this block
    date_cell = df_raw.iloc[header_row_idx-1, day_col] if header_row_idx-1 >= 0 else None
    parsed_date = parse_date_label(date_cell)

    # If parsing failed, look around neighboring columns
    if parsed_date is None:
        for adj in [-1, 1, 2, -2]:
            c = day_col + adj
            if 0 <= c < num_cols:
                parsed_date = parse_date_label(df_raw.iloc[header_row_idx-1, c])
                if parsed_date is not None:
                    break

    # Loop through every employee row
    for r in range(first_employee_row, df_raw.shape[0]):
        emp = df_raw.iloc[r, 0]  # employee name
        if pd.isna(emp):
            continue
        emp_s = str(emp).strip()
        # Skip blank, totals, or non-names
        if not emp_s or emp_s.lower().startswith('total') or emp_s.lower().startswith('ama'):
            continue

        # Get job code and hours worked from this row/block
        job = df_raw.iloc[r, job_col] if job_col is not None and job_col < num_cols else None
        hours = df_raw.iloc[r, hours_col] if hours_col is not None and hours_col < num_cols else None

        # Convert hours safely to float
        try:
            hrs = float(hours) if pd.notna(hours) and str(hours).strip() != '' else 0.0
        except Exception:
            # If it's something weird like "8 hrs", grab just the number
            m = re.search(r'(\d+(?:\.\d+)?)', str(hours))
            hrs = float(m.group(1)) if m else 0.0

        # Skip rows with no job or 0 hours
        if pd.isna(job) or (not str(job).strip()) or hrs == 0:
            continue

        # Save this record
        records.append({
            "Employee": emp_s,
            "Date": pd.to_datetime(parsed_date) if parsed_date is not None else None,
            "Job": str(job).strip(),
            "Hours": hrs,
            "Trk#": df_raw.iloc[r, day_col] if day_col < num_cols else None,
            "SourceRow": r,          # (debug info)
            "DayStartCol": day_col   # (debug info)
        })


# --- STEP 8: Turn into DataFrames ---
if not records:
    raise RuntimeError("No records parsed -- header detection may need tweaking.")

# "long_df" = one row per employee/day/job/hours
long_df = pd.DataFrame.from_records(records)
# Ensure Date column is just a date (no time)
long_df['Date'] = pd.to_datetime(long_df['Date']).dt.date

# "summary" = group by Date + Job
summary = (
    long_df
    .groupby(['Date', 'Job'], dropna=False)
    .agg(
        EmployeeCount=('Employee', lambda x: x.nunique()),   # how many distinct employees
        TotalHours=('Hours', 'sum'),                         # total hours worked
        Employees=('Employee', lambda x: ', '.join(sorted(set(x))))  # list employees
    )
    .reset_index()
)


# --- STEP 9: Save results ---
out_csv = "timesheet_long_parsed.csv"
out_summary = "timesheet_daily_summary.xlsx"

# Save detailed records to CSV
long_df.to_csv(out_csv, index=False)
# Save summary table to Excel
summary.to_excel(out_summary, index=False)

print("Saved:", out_csv, out_summary)
