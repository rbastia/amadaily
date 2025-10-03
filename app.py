from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
import pandas as pd
from parse_timesheet import process_timesheet
from parse_job_sheet import process_job_sheet
from combine_parsers import combine_daily_reports

# Create a Flask app instance
app = Flask(__name__)

# Upload safety config
# limit request size to 50 MB by default (can override via env var MAX_CONTENT_LENGTH)
default_max = int(os.environ.get("MAX_CONTENT_LENGTH", 50 * 1024 * 1024))
app.config["MAX_CONTENT_LENGTH"] = default_max

# Only allow Excel workbooks (must contain both required sheets)
ALLOWED_EXTENSIONS = {"xlsx", "xls", "xlsm", "xlsb"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# Folders where files will live
UPLOAD_FOLDER = "uploads"   # where uploaded Excel files go
OUTPUT_FOLDER = "outputs"   # where parsed summary files go

# Make sure those folders exist so saving files doesn’t crash
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")  # when you go to http://127.0.0.1:5000/
def index():
    # Render the HTML form page
    return render_template("index.html")


@app.route("/upload", methods=["POST"])  # single workbook upload
def upload_file():
    file = request.files.get("workbook")
    if not file or file.filename == "":
        return "No file uploaded", 400
    if not allowed_file(file.filename):
        return "Unsupported file type (must be Excel)", 400

    fname = secure_filename(file.filename)
    saved_path = os.path.join(UPLOAD_FOLDER, fname)
    try:
        file.save(saved_path)
    except Exception:
        return "Failed to save file", 500

    # Validate required sheets exist first to give a quick failure if not
    required_sheets = {"Timesheet", "New Formula Job Sheet"}
    try:
        xl = pd.ExcelFile(saved_path)
        sheet_set = set(xl.sheet_names)
        missing = required_sheets - sheet_set
        if missing:
            return f"Workbook missing required sheet(s): {', '.join(missing)}", 400
    except Exception as e:
        return f"Could not open workbook: {e}", 400

    try:
        # Run both parsers against their respective sheets inside the single workbook
        timesheet_summary_path = process_timesheet(saved_path, OUTPUT_FOLDER, sheet_name="Timesheet")
        job_sheet_normalized_path = os.path.join(OUTPUT_FOLDER, "ex_job_sheet_normalized.xlsx")
        process_job_sheet(saved_path, sheet_name="New Formula Job Sheet", output_path=job_sheet_normalized_path)

        # Determine per-sheet flag from form checkbox (present when checked)
        per_sheet_flag = bool(request.form.get('per_sheet'))

        combined_path = combine_daily_reports(
            timesheet_summary_path=timesheet_summary_path,
            job_sheet_table_path=job_sheet_normalized_path,
            output_dir=OUTPUT_FOLDER,
            output_basename="combined_daily_report",
            write_csv=False,
            fuzzy=True,
            per_sheet=per_sheet_flag,
        )

        # Optionally remove intermediate artifacts so only combined output remains.
        # Set KEEP_INTERMEDIATE=1 to skip deletion (for debugging).
        if os.environ.get("KEEP_INTERMEDIATE") != "1":
            intermediate = [
                timesheet_summary_path,
                job_sheet_normalized_path,
                os.path.join(OUTPUT_FOLDER, "timesheet_long_parsed.csv"),  # created by process_timesheet
            ]
            # Don't delete the combined Excel output
            protected = {combined_path}
            for p in intermediate:
                if p and os.path.exists(p) and p not in protected:
                    try:
                        os.remove(p)
                    except Exception:
                        pass
    except Exception as e:
        return f"Processing failed: {e}", 500

    return send_file(combined_path, as_attachment=True)


if __name__ == "__main__":
    # Run the Flask development server
    app.run(debug=True)
