"""combine_parsers.py

Utilities to combine the outputs of `parse_timesheet.process_timesheet` and
`parse_job_sheet.process_job_sheet` into a single daily report.

Expected Inputs
---------------
1. Timesheet summary Excel produced by `process_timesheet`, typically
   `timesheet_daily_summary.xlsx` with columns:
	  Date, Job, EmployeeCount, TotalHours, TotalDrivingHours, Employees
2. Normalized Job Sheet table (either CSV or Excel) produced by
   `process_job_sheet` (or a saved version) with columns:
	  Date, Job, Truck(s), Description, Concrete, Concrete Yds, Stone, Stone Lds

Output
------
An Excel (and optional CSV) file combining both sets of columns on (Date, Job).
If a job appears in one source but not the other it is still included (outer
join). Basic text normalization is applied to improve matching.

Usage (Programmatic)
--------------------
from combine_parsers import combine_daily_reports
path = combine_daily_reports(
	timesheet_summary_path="outputs/timesheet_daily_summary.xlsx",
	job_sheet_table_path="outputs/ex_job_sheet_normalized.xlsx"
)
print("Wrote", path)

CLI
---
python combine_parsers.py \
	--timesheet outputs/timesheet_daily_summary.xlsx \
	--jobsheet outputs/ex_job_sheet_normalized.xlsx \
	--outdir outputs

Notes
-----
* We purposefully avoid heavy fuzzy matching libraries; a light difflib pass
  is used only for unmatched rows.
* A `_CanonicalJob` column is added internally for merge logic and removed in
  the final output.
"""


from __future__ import annotations

import os
import argparse
import difflib
from datetime import date, datetime
from typing import Tuple, Dict

import pandas as pd


# ------------------------- Normalization Helpers ------------------------- #
def _norm_job(job: str | None) -> str:
	"""Return a normalized key for a job name.

	Strategy: strip, collapse internal whitespace, remove trailing commas,
	lowercase, and remove duplicate spaces. This keeps alphanumerics & symbols
	(so ARA3A vs ARA3A Moorefield stay distinct) while smoothing accidental
	spacing differences.
	"""
	if job is None:
		return ""
	s = str(job).strip().strip(',')
	if not s:
		return ""
	# Collapse all internal whitespace to single spaces
	parts = s.split()
	s = " ".join(parts)
	return s.lower()


def _coerce_date(v) -> date | None:
	if pd.isna(v):
		return None
	if isinstance(v, date) and not isinstance(v, datetime):  # already date
		return v
	try:
		return pd.to_datetime(v).date()
	except Exception:
		return None


def _prepare_timesheet_df(df: pd.DataFrame) -> pd.DataFrame:
	required = {"Date", "Job"}
	missing = required - set(df.columns)
	if missing:
		raise ValueError(f"Timesheet summary missing columns: {missing}")
	df = df.copy()
	df["Date"] = df["Date"].apply(_coerce_date)
	df["_CanonicalJob"] = df["Job"].apply(_norm_job)
	return df


def _prepare_job_sheet_df(df: pd.DataFrame) -> pd.DataFrame:
	required = {"Date", "Job"}
	missing = required - set(df.columns)
	if missing:
		raise ValueError(f"Job sheet table missing columns: {missing}")
	df = df.copy()

	# Heuristic renaming for placeholder-style columns that appeared in some
	# normalized job sheet outputs (user reported '_5', '_8', '_10'). These
	# likely represent Truck(s), Concrete Yds, and Stone Lds respectively when
	# the original header extraction produced generic positional names.
	# Expand mapping to handle variants that pandas may create: plain numbers,
	# 'Unnamed: X', or with leading underscore. All keys compared case-insensitively.
	_alias_targets = {
		"5": "Truck(s)",
		"_5": "Truck(s)",
		"unnamed: 5": "Truck(s)",
		"8": "Concrete Yds",
		"_8": "Concrete Yds",
		"unnamed: 8": "Concrete Yds",
		"10": "Stone Lds",
		"_10": "Stone Lds",
		"unnamed: 10": "Stone Lds",
	}
	# Build rename dict by scanning existing columns (cast to string for safety)
	rename_dict = {}
	for col in list(df.columns):
		col_key = str(col).strip().lower()
		if col_key in _alias_targets:
			desired = _alias_targets[col_key]
			if desired not in df.columns:  # avoid overwriting if already present
				rename_dict[col] = desired
	if rename_dict:
		df = df.rename(columns=rename_dict)

	df["Date"] = df["Date"].apply(_coerce_date)
	df["_CanonicalJob"] = df["Job"].apply(_norm_job)
	return df


# --------------------------- Fuzzy Reconciliation --------------------------- #
def _fuzzy_reconcile(left: pd.DataFrame, right: pd.DataFrame) -> Dict[Tuple[date, str], str]:
	"""Attempt fuzzy reconciliation for left rows whose canonical job didn't
	find an exact match in right for the same date.

	Returns a mapping of (date, left_canonical_job) -> right_canonical_job.
	Only applied when there's exactly one reasonably close candidate.
	"""
	mapping: Dict[Tuple[date, str], str] = {}
	# Build per-date index of right canonical jobs
	by_date: Dict[date, list] = {}
	for d, sub in right.groupby("Date"):
		by_date[d] = list(sub["_CanonicalJob"].unique())

	unmatched = left[left["_CanonicalJob"].notna()][["Date", "_CanonicalJob"]]
	merged_keys = set(zip(right["Date"], right["_CanonicalJob"]))
	for d, cjob in unmatched.itertuples(index=False):
		if (d, cjob) in merged_keys or d not in by_date:
			continue
		candidates = by_date.get(d, [])
		if not candidates:
			continue
		# Use difflib; cutoff 0.82 for reasonably similar names
		close = difflib.get_close_matches(cjob, candidates, n=2, cutoff=0.82)
		if len(close) == 1:
			mapping[(d, cjob)] = close[0]
	return mapping


# ------------------------------- Main Combine ------------------------------- #
def combine_daily_reports(
	timesheet_summary_path: str,
	job_sheet_table_path: str,
	output_dir: str = "outputs",
	output_basename: str | None = None,
	write_csv: bool = True,
	fuzzy: bool = True,
	per_sheet: bool = False,
) -> str:
	"""Combine the two parser outputs and write an Excel file.

	Parameters
	----------
	timesheet_summary_path : str
		Path to the timesheet daily summary Excel (columns: Date, Job, ...).
	job_sheet_table_path : str
		Path to the normalized job sheet table (CSV or Excel) with Date & Job.
	output_dir : str, default 'outputs'
		Folder to place the combined output.
	output_basename : str | None
		Base name (without extension). Default builds from current timestamp.
	write_csv : bool, default True
		Also write a CSV next to the Excel.
	fuzzy : bool, default True
		Attempt a light fuzzy reconciliation for unmatched jobs on same date.

	Returns
	-------
	str : path to the written Excel file.
	"""
	if not os.path.exists(timesheet_summary_path):
		raise FileNotFoundError(timesheet_summary_path)
	if not os.path.exists(job_sheet_table_path):
		raise FileNotFoundError(job_sheet_table_path)

	# Read inputs
	ts = pd.read_excel(timesheet_summary_path) if timesheet_summary_path.lower().endswith(".xlsx") else pd.read_csv(timesheet_summary_path)
	js = pd.read_excel(job_sheet_table_path) if job_sheet_table_path.lower().endswith(".xlsx") else pd.read_csv(job_sheet_table_path)

	ts_prep = _prepare_timesheet_df(ts)
	js_prep = _prepare_job_sheet_df(js)

	# Optional fuzzy reconciliation mapping
	if fuzzy:
		mapping = _fuzzy_reconcile(ts_prep, js_prep)
		if mapping:
			# Replace left canonical job with the matched right canonical job key to enable merge
			def _map_row(row):
				key = (row["Date"], row["_CanonicalJob"])
				return mapping.get(key, row["_CanonicalJob"])
			ts_prep["_CanonicalJob"] = ts_prep.apply(_map_row, axis=1)

	# Merge on Date + _CanonicalJob (outer)
	left_cols = [c for c in ts_prep.columns if c not in ("_CanonicalJob",)]
	right_cols = [c for c in js_prep.columns if c not in ("_CanonicalJob",)]

	merged = pd.merge(
		ts_prep, js_prep,
		how="outer",
		on=["Date", "_CanonicalJob"],
		suffixes=("_TS", "_JS")
	)

	# Resolve Job column preference: prefer job sheet's original Job when present
	def _choose_job(row):
		job_js = row.get("Job_JS")
		job_ts = row.get("Job_TS")
		if isinstance(job_js, str) and job_js.strip():
			return job_js
		return job_ts
	# Add a human-facing Job column (prefer job sheet Job when available)
	merged["Job"] = merged.apply(_choose_job, axis=1)

	# Reorder & select columns: Job first, Date second
	# Helper to pick the first available column among base and common merge suffixes
	def _pick_first(df_cols, base: str):
		for cand in (base, f"{base}_TS", f"{base}_JS"):
			if cand in df_cols:
				return cand
		return None

	_timesheet_bases = ["EmployeeCount", "TotalHours", "TotalDrivingHours", "Employees"]
	_timesheet_metrics = [c for c in (_pick_first(merged.columns, b) for b in _timesheet_bases) if c]

	ordered = [
		"Job", "Date",
		*(_timesheet_metrics),
		*[c for c in ("Truck(s)", "Description", "Concrete", "Concrete Yds", "Stone", "Stone Lds") if c in merged.columns],
	]

	# Bring over any remaining columns (excluding internal / duplicates)
	internal = set(ordered + ["_CanonicalJob", "Job_TS", "Job_JS"])
	remaining = [c for c in merged.columns if c not in internal]
	final_df = merged[ordered + remaining].copy()

	# Shorter rename logic: map first found TotalHours -> Working Hours; first found
	# TotalDrivingHours -> Driving Hours
	_rename_map = {}
	wh = _pick_first(final_df.columns, "TotalHours")
	if wh and "Working Hours" not in final_df.columns:
		_rename_map[wh] = "Working Hours"
	dh = _pick_first(final_df.columns, "TotalDrivingHours")
	if dh and "Driving Hours" not in final_df.columns:
		_rename_map[dh] = "Driving Hours"
	if _rename_map:
		final_df = final_df.rename(columns=_rename_map)

	# Move Driving Hours next to Working Hours if both present
	if "Working Hours" in final_df.columns and "Driving Hours" in final_df.columns:
		cols = list(final_df.columns)
		cols.remove("Driving Hours")
		cols.insert(cols.index("Working Hours") + 1, "Driving Hours")
		final_df = final_df[cols]

	# Sort by Job (alphabetically, case-insensitive) then Date
	# Use a temporary lower-cased key so sorting is deterministic regardless of case
	final_df["_job_sort_key"] = final_df["Job"].astype(str).str.lower()
	final_df = final_df.sort_values(["_job_sort_key", "Date"], kind="stable").drop(columns=["_job_sort_key"]).reset_index(drop=True)
	# Output paths
	os.makedirs(output_dir, exist_ok=True)
	if output_basename is None:
		# Default output name: prefix the input timesheet filename with "COMBINED "
		# Example: "Timesheet 9-7-25 thru 9-13-25.xlsx" ->
		# "COMBINED Timesheet 9-7-25 thru 9-13-25.xlsx"
		input_basename = os.path.basename(timesheet_summary_path)
		# Avoid double-prefixing if the input already starts with COMBINED
		if input_basename.lower().startswith("combined "):
			xlsx_name = input_basename if input_basename.lower().endswith(".xlsx") else input_basename
		else:
			if input_basename.lower().endswith(".xlsx"):
				xlsx_name = f"COMBINED {input_basename}"
			else:
				name_no_ext, _ = os.path.splitext(input_basename)
				xlsx_name = f"COMBINED {name_no_ext}.xlsx"
		xlsx_path = os.path.join(output_dir, xlsx_name)
		# CSV sidecar uses same base name with .csv extension
		csv_path = os.path.join(output_dir, os.path.splitext(xlsx_name)[0] + ".csv")
	else:
		xlsx_path = os.path.join(output_dir, f"{output_basename}.xlsx")
		csv_path = os.path.join(output_dir, f"{output_basename}.csv")

	if per_sheet:
		# Write one worksheet per (Date, Job) plus an 'All' summary sheet.
		def _sanitize_sheet_name(name: str) -> str:
			# Remove characters Excel forbids and trim to 31 chars.
			invalid = set('[]:*?/\\')
			clean = ''.join(ch for ch in name if ch not in invalid)
			return clean[:31] if len(clean) > 31 else clean

		with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
			# All sheet first
			final_df.to_excel(writer, sheet_name="All", index=False)
			used_names = {"All"}
			# One sheet per unique Job (all dates included). Sort jobs alphabetically for predictability.
			# Treat missing jobs as 'UnknownJob'.
			jobs = list(final_df["Job"].fillna("UnknownJob").unique())
			# Sort case-insensitively
			jobs = sorted(jobs, key=lambda v: str(v).lower())
			for job_val in jobs:
				sub = final_df[final_df["Job"].fillna("UnknownJob") == job_val]
				job_str = str(job_val) if job_val is not None else 'UnknownJob'
				base = _sanitize_sheet_name(job_str)
				name = base if base else 'Sheet'
				# Ensure uniqueness and respect Excel 31-char limit
				counter = 2
				while name in used_names:
					# Reserve room for suffix like _2
					suffix = f"_{counter}"
					trunc = base[:31 - len(suffix)]
					name = _sanitize_sheet_name(f"{trunc}{suffix}")
					counter += 1
				used_names.add(name)
				sub.to_excel(writer, sheet_name=name, index=False)
	else:
		# Single sheet mode (original behavior)
		final_df.to_excel(xlsx_path, index=False)
		if write_csv:
			final_df.to_csv(csv_path, index=False)

	return xlsx_path


# ----------------------------------- CLI ----------------------------------- #
def _build_arg_parser() -> argparse.ArgumentParser:
	p = argparse.ArgumentParser(description="Combine timesheet & job sheet outputs into a single daily report")
	p.add_argument("--timesheet", required=True, help="Path to timesheet_daily_summary.xlsx (or CSV)")
	p.add_argument("--jobsheet", required=True, help="Path to normalized job sheet table (.xlsx or .csv)")
	# New unified workbook mode: user can supply a single Excel file containing both
	# sheets. If provided, we will parse both from that workbook and ignore the
	# explicit --timesheet/--jobsheet file paths (unless user overrides).
	p.add_argument("--single-workbook", help="Path to one Excel file containing both Timesheet and Job Sheet source sheets")
	p.add_argument("--timesheet-sheet", default="Timesheet", help="Sheet name for the timesheet inside the single workbook (default: Timesheet)")
	p.add_argument("--jobsheet-sheet", default="New Formula Job Sheet", help="Sheet name for the job sheet inside the single workbook (default: New Formula Job Sheet)")
	p.add_argument("--outdir", default="outputs", help="Output directory (default: outputs)")
	p.add_argument("--name", help="Optional base filename without extension")
	p.add_argument("--no-csv", action="store_true", help="Do not write CSV alongside Excel")
	p.add_argument("--no-fuzzy", action="store_true", help="Disable fuzzy job name reconciliation")
	p.add_argument("--per-sheet", action="store_true", help="Write each Date+Job combo to its own worksheet (also keeps an 'All' sheet)")
	return p


def main():  # pragma: no cover - simple CLI wrapper
	args = _build_arg_parser().parse_args()

	# If unified workbook is supplied, generate intermediate normalized outputs first
	if args.single_workbook:
		from parse_timesheet import process_timesheet
		from parse_job_sheet import process_job_sheet
		os.makedirs(args.outdir, exist_ok=True)
		# Derive temp paths we will feed to the combiner
		timesheet_summary_path = os.path.join(args.outdir, "timesheet_daily_summary.xlsx")
		job_sheet_norm_path = os.path.join(args.outdir, "ex_job_sheet_normalized.xlsx")
		# Run parsers
		process_timesheet(args.single_workbook, output_dir=args.outdir, sheet_name=args.timesheet_sheet)
		process_job_sheet(args.single_workbook, sheet_name=args.jobsheet_sheet, output_path=job_sheet_norm_path)
		# Override args for downstream combine
		args.timesheet = timesheet_summary_path
		args.jobsheet = job_sheet_norm_path
		# If the user didn't provide an explicit --name, prefer the original
		# uploaded workbook's basename so the combined output mirrors it.
		if not args.name:
			uploaded_base = os.path.basename(args.single_workbook)
			name_no_ext, _ = os.path.splitext(uploaded_base)
			args.name = f"COMBINED {name_no_ext}"
	path = combine_daily_reports(
		timesheet_summary_path=args.timesheet,
		job_sheet_table_path=args.jobsheet,
		output_dir=args.outdir,
		output_basename=args.name,
		write_csv=not args.no_csv,
		fuzzy=not args.no_fuzzy,
		per_sheet=args.per_sheet,
	)
	print(f"Combined report written: {path}")


if __name__ == "__main__":  # pragma: no cover
	main()

