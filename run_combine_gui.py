"""run_combine_gui.py

One-click (well, a few clicks) GUI to combine ANY Excel workbook that contains
both the Timesheet sheet and Job Sheet into the consolidated report.

No command line needed. Launch it, pick the workbook, optionally adjust sheet
names, and press the big Combine button.

Default expected sheet names:
  Timesheet
  New Formula Job Sheet

Outputs go to the selected output folder (default: ./outputs).

Requires: pandas, openpyxl, python-dateutil (already in requirements). Tkinter
is bundled with standard Python on Windows.
"""
from __future__ import annotations

import os
import threading
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd  # noqa: F401 (just to ensure dependency error surfaces early)

from parse_timesheet import process_timesheet
from parse_job_sheet import process_job_sheet
from combine_parsers import combine_daily_reports

APP_TITLE = "Daily Report Combiner"


def _safe_run(fn, on_done):
    """Run fn() in a thread then call on_done(success, message, path)."""
    def runner():
        try:
            result = fn()
            on_done(True, "Success", result)
        except Exception as e:  # noqa: BLE001 - show full trace
            tb = traceback.format_exc()
            on_done(False, f"{e}\n\n{tb}", None)
    threading.Thread(target=runner, daemon=True).start()


def launch():  # pragma: no cover - GUI launcher
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("560x360")

    # State variables
    workbook_var = tk.StringVar()
    outdir_var = tk.StringVar(value=os.path.abspath("outputs"))
    timesheet_sheet_var = tk.StringVar(value="Timesheet")
    jobsheet_sheet_var = tk.StringVar(value="New Formula Job Sheet")
    per_sheet_var = tk.BooleanVar(value=True)
    fuzzy_var = tk.BooleanVar(value=True)
    name_var = tk.StringVar(value="")

    status_var = tk.StringVar(value="Idle")

    def pick_workbook():
        path = filedialog.askopenfilename(
            title="Select Source Workbook",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb *.xls"), ("All", "*.*")],
        )
        if path:
            workbook_var.set(path)

    def pick_outdir():
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            outdir_var.set(path)

    def do_combine():
        wb = workbook_var.get().strip()
        if not wb:
            messagebox.showerror("Missing", "Please select a workbook first.")
            return
        if not os.path.exists(wb):
            messagebox.showerror("Missing", f"Workbook not found: {wb}")
            return
        outdir = outdir_var.get().strip() or "outputs"
        os.makedirs(outdir, exist_ok=True)

        t_sheet = timesheet_sheet_var.get().strip() or "Timesheet"
        j_sheet = jobsheet_sheet_var.get().strip() or "New Formula Job Sheet"
        per_sheet = per_sheet_var.get()
        fuzzy = fuzzy_var.get()
        custom_name = name_var.get().strip() or None

        status_var.set("Working...")
        combine_btn.config(state=tk.DISABLED)

        def work():
            # Produce intermediate files then combine
            timesheet_summary_path = os.path.join(outdir, "timesheet_daily_summary.xlsx")
            job_sheet_norm_path = os.path.join(outdir, "ex_job_sheet_normalized.xlsx")
            process_timesheet(wb, output_dir=outdir, sheet_name=t_sheet)
            process_job_sheet(wb, sheet_name=j_sheet, output_path=job_sheet_norm_path)
            if custom_name is None:
                base_name = f"combined_daily_report_{datetime.now():%Y%m%d_%H%M%S}"
            else:
                base_name = custom_name
            combined_path = combine_daily_reports(
                timesheet_summary_path=timesheet_summary_path,
                job_sheet_table_path=job_sheet_norm_path,
                output_dir=outdir,
                output_basename=base_name,
                write_csv=True,
                fuzzy=fuzzy,
                per_sheet=per_sheet,
            )
            return combined_path

        def done(success: bool, msg: str, path: str | None):
            combine_btn.config(state=tk.NORMAL)
            if success:
                status_var.set(f"Done: {os.path.basename(path)}")
                open_btn.config(state=tk.NORMAL)
                messagebox.showinfo("Success", f"Combined report written:\n{path}")
            else:
                status_var.set("Error")
                messagebox.showerror("Error", msg)

        _safe_run(work, done)

    def open_output():  # open folder in explorer
        folder = outdir_var.get().strip() or os.path.abspath("outputs")
        if not os.path.exists(folder):
            messagebox.showwarning("Missing", f"Folder not found: {folder}")
            return
        os.startfile(folder)  # Windows only

    # Layout
    pad = {"padx": 6, "pady": 4, "sticky": "w"}

    frm = tk.Frame(root)
    frm.pack(fill="both", expand=True, padx=10, pady=10)

    tk.Label(frm, text="Source Workbook:").grid(row=0, column=0, **pad)
    tk.Entry(frm, textvariable=workbook_var, width=50).grid(row=0, column=1, **pad)
    tk.Button(frm, text="Browse", command=pick_workbook).grid(row=0, column=2, **pad)

    tk.Label(frm, text="Output Folder:").grid(row=1, column=0, **pad)
    tk.Entry(frm, textvariable=outdir_var, width=50).grid(row=1, column=1, **pad)
    tk.Button(frm, text="Browse", command=pick_outdir).grid(row=1, column=2, **pad)

    tk.Label(frm, text="Timesheet Sheet Name:").grid(row=2, column=0, **pad)
    tk.Entry(frm, textvariable=timesheet_sheet_var, width=25).grid(row=2, column=1, **pad)

    tk.Label(frm, text="Job Sheet Name:").grid(row=3, column=0, **pad)
    tk.Entry(frm, textvariable=jobsheet_sheet_var, width=25).grid(row=3, column=1, **pad)

    tk.Label(frm, text="Custom Base Name (optional):").grid(row=4, column=0, **pad)
    tk.Entry(frm, textvariable=name_var, width=30).grid(row=4, column=1, **pad)

    tk.Checkbutton(frm, text="Per-Job Sheets", variable=per_sheet_var).grid(row=5, column=0, **pad)
    tk.Checkbutton(frm, text="Fuzzy Match Jobs", variable=fuzzy_var).grid(row=5, column=1, **pad)

    combine_btn = tk.Button(frm, text="Combine Now", font=("Segoe UI", 12, "bold"), bg="#4caf50", fg="white", command=do_combine)
    combine_btn.grid(row=6, column=0, columnspan=2, padx=6, pady=12, sticky="we")

    open_btn = tk.Button(frm, text="Open Output Folder", command=open_output, state=tk.DISABLED)
    open_btn.grid(row=6, column=2, padx=6, pady=12, sticky="we")

    tk.Label(frm, textvariable=status_var, fg="#555").grid(row=7, column=0, columnspan=3, padx=6, pady=8, sticky="w")

    tk.Label(frm, text="Tip: Select a workbook, adjust names only if your sheet titles differ, then click Combine.", fg="#777").grid(row=8, column=0, columnspan=3, padx=6, pady=4, sticky="w")

    root.mainloop()


if __name__ == "__main__":  # pragma: no cover
    launch()
