"""parse_job_sheet.py

Lightweight parser for the "New Formula Job Sheet" layout.

Produces a flat table with one row per (Date, Job) and these columns:
    Date, Job, Truck(s), Description, Concrete, Concrete Yds, Stone, Stone Lds

Supports input as a CSV export or an Excel workbook. Minimal, clear
comments are provided to make the parsing steps easy to follow.
"""
from __future__ import annotations

import os
import re
from typing import Optional

import pandas as pd
from dateutil.parser import parse as parse_date


def _read_sheet(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Read file as a raw table (header=None) so we can inspect layout.

    We return a DataFrame where row 0 is the original sheet header and
    following rows match the sheet rows. CSV inputs are read as strings.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
        # read the requested sheet (fallback to first sheet)
        try:
            df = pd.read_excel(path, sheet_name=sheet_name or 0, header=None, engine="openpyxl")
        except Exception:
            # try without engine (older pandas)
            df = pd.read_excel(path, sheet_name=sheet_name or 0, header=None)
    else:
        # assume CSV representing a single sheet export
        df = pd.read_csv(path, header=None, dtype=str)
    # Normalize: ensure all cells are strings (keep NaN for empties)
    df = df.astype(object)
    return df


def process_job_sheet(path: str, sheet_name: str = "New Formula Job Sheet", output_path: Optional[str] = None) -> pd.DataFrame:
    """Parse the job sheet and return a normalized DataFrame.

    The function expects the top row to alternate DayName / Date columns (e.g. Monday, 9/8/2025, Tuesday, 9/9/2025, ...).
    It detects 'Job & Truck' blocks and extracts the related Description, Concrete & Yds, and Stone & Lds rows.

    Returns a DataFrame with columns: Date, Job, Truck(s), Description, Concrete, Concrete Yds, Stone, Stone Lds
    """
    raw = _read_sheet(path, sheet_name=sheet_name)

    if raw.shape[0] == 0:
        return pd.DataFrame(columns=["Date", "Job", "Truck(s)", "Description", "Concrete", "Concrete Yds", "Stone", "Stone Lds"])

    header_row = raw.iloc[0].tolist()
    # The sheet's first row alternates day-name and date. Find columns that
    # contain actual dates (we only try to parse cells that include '/' or '-'). This is important because 
    date_cols = []
    date_values = {}
    for col_idx, cell in enumerate(header_row):
        if pd.isna(cell):
            continue
        s = str(cell).strip()
        if '/' not in s and '-' not in s:
            # not in an explicit date-like format we expect from the sheet
            continue
        try:
            dt = parse_date(s, fuzzy=False)
            date_cols.append(col_idx)
            date_values[col_idx] = dt.date().isoformat()
        except Exception:
            # ignore cells that aren't valid dates even if they contain separators
            continue

    # For each date column remember the previous column (where the job name sits)
    pairs = []
    for date_col in date_cols:
        day_col = date_col - 1 if date_col > 0 else date_col
        pairs.append((day_col, date_col, date_values[date_col]))

    # Find rows that start job blocks (left-most cell contains 'Job & Truck')
    first_col = raw.iloc[:, 0].astype(str).fillna("")
    group_rows = [i for i, v in enumerate(first_col) if "Job & Truck" in v]

    # Build a mapping from job-block row index -> concrete row and stone row
    # by scanning the sheet top-to-bottom. When we encounter a 'Job & Truck'
    # row we set the current_group; any subsequent 'Concrete' or 'Stone' rows
    # are attached to that current_group until the next job appears.
    concrete_map = {}
    stone_map = {}
    current_group = None
    for i in range(len(raw)):
        cell0 = str(raw.iat[i, 0]) if i < len(raw) else ""
        if "Job & Truck" in cell0:
            current_group = i
            # ensure maps have a default entry (not strictly necessary)
            concrete_map.setdefault(current_group, [None] * raw.shape[1])
            stone_map.setdefault(current_group, [None] * raw.shape[1])
        elif "Concrete" in cell0 and current_group is not None:
            concrete_map[current_group] = raw.iloc[i].tolist()
        elif "Stone" in cell0 and current_group is not None:
            stone_map[current_group] = raw.iloc[i].tolist()

    records = []
    for gr in group_rows:
        job_row = raw.iloc[gr].tolist()
        desc_row = raw.iloc[gr + 1].tolist() if gr + 1 < len(raw) else [None] * raw.shape[1]

        # Use the concrete/stone rows attached to this job block (if any)
        concrete_row = concrete_map.get(gr, [None] * raw.shape[1])
        stone_row = stone_map.get(gr, [None] * raw.shape[1])

        for day_col, date_col, iso_date in pairs:
            # Read fields for this (date, job) from the appropriate columns
            job = _safe_get(job_row, day_col)
            trucks = _safe_get(job_row, date_col)
            desc = _safe_get(desc_row, day_col)
            concrete = _safe_get(concrete_row, day_col)
            concrete_yds = _safe_get(concrete_row, date_col)
            stone = _safe_get(stone_row, day_col)
            stone_lds = _safe_get(stone_row, date_col)

            # Only emit a row when there's something meaningful
            if any([_has_value(x) for x in (job, trucks, desc, concrete, concrete_yds, stone, stone_lds)]):
                records.append({
                    "Date": iso_date,
                    "Job": _clean_str(job),
                    "Truck(s)": _normalize_trucks(_clean_str(trucks)),
                    "Description": _clean_str(desc),
                    "Concrete": _clean_str(concrete),
                    "Concrete Yds": _clean_str(concrete_yds),
                    "Stone": _clean_str(stone),
                    "Stone Lds": _clean_str(stone_lds),
                })

    out = pd.DataFrame.from_records(records, columns=["Date", "Job", "Truck(s)", "Description", "Concrete", "Concrete Yds", "Stone", "Stone Lds"]) 

    # If output_path provided, write to Excel for .xlsx/.xls or to CSV otherwise
    if output_path:
        _, ext = os.path.splitext(output_path)
        ext = ext.lower()
        if ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
            # write Excel (requires openpyxl for xlsx)
            try:
                out.to_excel(output_path, index=False, engine="openpyxl")
            except Exception:
                # fallback without engine
                out.to_excel(output_path, index=False)
        else:
            out.to_csv(output_path, index=False)

    return out


def _safe_get(row: list, idx: int):
    try:
        return row[idx]
    except Exception:
        return None


def _has_value(v) -> bool:
    if v is None:
        return False
    s = str(v).strip()
    return s != "" and s.lower() not in ("nan",)


def _clean_str(v) -> Optional[str]:
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    # normalize quotes and multiple commas
    s = s.replace('"', '')
    # trim whitespace around commas
    parts = [p.strip() for p in s.split(',') if p.strip()]
    if len(parts) > 1 and any(p.isdigit() for p in parts):
        # join numeric lists with comma+space
        return ", ".join(parts)
    return s


def _normalize_trucks(s: str) -> str:
    """Split concatenated truck identifiers into a comma+space list.

    Example: "125126" -> "125, 126" if it cleanly tokenizes into multiple
    alphanumeric groups whose concatenation equals the original string.

    We only modify when:
      - There are no existing delimiters (comma, space, slash, dash)
      - Regex finds 2+ tokens of the pattern [A-Za-z]*\d+[A-Za-z]*
      - Joining the tokens exactly reconstructs the original string
    """
    if not s:
        return s
    if any(d in s for d in [',', ' ', '/', '-']):  # already delimited
        return s
    # 1. Heuristic for pure digits: attempt fixed-width segmentation (common truck ID lengths)
    if s.isdigit() and len(s) >= 6:  # at least two IDs of length >=3
        # Prefer 3-digit IDs (e.g., 125126 -> 125, 126). Fall back to 4 or 2 if evenly divisible.
        for width in (3, 4, 2):
            if len(s) % width == 0 and len(s) // width >= 2:
                parts = [s[i:i+width] for i in range(0, len(s), width)]
                # Sanity check: avoid splitting something like a single repeated digit sequence oddly
                if all(part.lstrip('0') for part in parts):
                    return ', '.join(parts)
    # 2. General alphanumeric token attempt (original logic, but only if it would create >1 token)
    tokens = re.findall(r'[A-Za-z]*\d+[A-Za-z]*', s)
    if len(tokens) > 1 and ''.join(tokens) == s:
        return ', '.join(tokens)
    return s


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process New Formula Job Sheet into normalized CSV/Excel")
    parser.add_argument("input", help="Path to input .xlsx or .csv file")
    parser.add_argument("-o", "--output", help="Optional output file path (.csv or .xlsx). If omitted, no file is written and results print to stdout")
    parser.add_argument("-s", "--sheet", default="New Formula Job Sheet", help="Sheet name for Excel files (default: New Formula Job Sheet)")
    args = parser.parse_args()
    df = process_job_sheet(args.input, sheet_name=args.sheet, output_path=args.output)
    # print CSV to stdout for quick preview
    print(df.to_csv(index=False))

# This is designed to make all of this 
def process_job_sheet_file(input_path: str, output_dir: str = "outputs", sheet_name: str = "New Formula Job Sheet") -> str:
        """High-level helper to match the parse_timesheet API style.

        Usage:
            from parse_job_sheet import process_job_sheet_file
            out = process_job_sheet_file(r"path\to\file.xlsx", output_dir="outputs")

        The function ensures output_dir exists, writes an Excel file
        called 'ex_job_sheet_normalized.xlsx' into that folder, and returns
        the output path.
        """
        os.makedirs(output_dir, exist_ok=True)
        out_path = os.path.join(output_dir, "ex_job_sheet_normalized.xlsx")
        # Call the lower-level parser which returns a DataFrame and writes the file
        process_job_sheet(input_path, sheet_name=sheet_name, output_path=out_path)
        return out_path
