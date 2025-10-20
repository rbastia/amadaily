from flask import Flask, render_template, request, send_file
import shutil
import sys
import subprocess
import os
from werkzeug.utils import secure_filename
import pandas as pd
from parse_timesheet import process_timesheet
from parse_job_sheet import process_job_sheet
from combine_parsers import combine_daily_reports
from flask import jsonify, send_from_directory

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

        # Build output basename from the original uploaded filename so the
        # combined workbook mirrors the input name (prefixed with "COMBINED ").
        # secure_filename was used for saving so reuse `fname` here.
        uploaded_base, _ = os.path.splitext(fname)
        output_basename = f"COMBINED {uploaded_base}"
        combined_path = combine_daily_reports(
            timesheet_summary_path=timesheet_summary_path,
            job_sheet_table_path=job_sheet_normalized_path,
            output_dir=OUTPUT_FOLDER,
            output_basename=output_basename,
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
        # Log full traceback to outputs/error.log for easier debugging when
        # running from an EXE or headless environment.
        import traceback
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        tb = traceback.format_exc()
        try:
            with open(os.path.join(OUTPUT_FOLDER, "error.log"), "a", encoding="utf-8") as fh:
                fh.write("\n--- ERROR: upload_file failed ---\n")
                fh.write(tb)
        except Exception:
            # If logging fails, ignore to avoid masking original error
            pass
        # Return a short message to the client and point them to the log file
        return ("Processing failed; full error written to outputs/error.log. "
                f"Summary: {e}"), 500

    # Ensure we use an absolute path (helps when running from an EXE with
    # a different current working directory).
    combined_path = os.path.abspath(combined_path)

    # If we're running as a frozen EXE, some browsers/packagers have trouble
    # returning files directly. Instead, copy the file to the user's
    # Downloads folder and open Explorer to show it.
    if getattr(sys, "frozen", False):
        try:
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(downloads, exist_ok=True)
            dest_base = os.path.join(downloads, os.path.basename(combined_path))
            dest = dest_base
            name, ext = os.path.splitext(dest_base)
            idx = 1
            while os.path.exists(dest):
                dest = f"{name} ({idx}){ext}"
                idx += 1
            shutil.copy2(combined_path, dest)
            # Try to open Explorer and select the file so the user sees it.
            try:
                subprocess.Popen(["explorer", "/select,", dest])
            except Exception:
                try:
                    # Fallback: just open the Downloads folder
                    subprocess.Popen(["explorer", downloads])
                except Exception:
                    pass
            return (f"Saved combined file to your Downloads folder: {dest}"), 200
        except Exception as e:
            import traceback, tempfile
            tb = traceback.format_exc()
            try:
                os.makedirs(OUTPUT_FOLDER, exist_ok=True)
                with open(os.path.join(OUTPUT_FOLDER, "error.log"), "a", encoding="utf-8") as fh:
                    fh.write("\n--- ERROR: copy to Downloads failed ---\n")
                    fh.write(tb)
            except Exception:
                pass
            try:
                tempfn = os.path.join(tempfile.gettempdir(), "ama_daily_error.log")
                with open(tempfn, "a", encoding="utf-8") as fh:
                    fh.write("\n--- ERROR: copy to Downloads failed ---\n")
                    fh.write(tb)
            except Exception:
                pass
            return ("Processing failed when saving the combined file to Downloads. "
                    "Full error written to outputs/error.log and to your system temp folder."), 500

    # Not frozen: return file as HTTP attachment (original behavior)
    is_xhr = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    if is_xhr:
        # Provide a safe URL that serves files from the outputs folder
        download_name = os.path.basename(combined_path)
        return jsonify({"download_url": f"/download/{download_name}"})

    try:
        return send_file(combined_path, as_attachment=True)
    except Exception as e:
        # Log full traceback to both outputs/error.log (relative) and a
        # known temp location for easier discovery when running the EXE.
        import traceback, tempfile
        tb = traceback.format_exc()
        # relative outputs in the current working dir
        try:
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)
            with open(os.path.join(OUTPUT_FOLDER, "error.log"), "a", encoding="utf-8") as fh:
                fh.write("\n--- ERROR: send_file failed ---\n")
                fh.write(tb)
        except Exception:
            pass
        # also write to system temp directory so it's easy to find
        try:
            tempfn = os.path.join(tempfile.gettempdir(), "ama_daily_error.log")
            with open(tempfn, "a", encoding="utf-8") as fh:
                fh.write("\n--- ERROR: send_file failed ---\n")
                fh.write(tb)
        except Exception:
            pass
        return ("Processing failed when sending the combined file. "
                "Full error written to outputs/error.log and to your system temp folder."), 500




if __name__ == "__main__":
    # Run the Flask development server
    app.run(debug=True)


@app.route('/download/<path:filename>')
def download_file(filename):
    # Prevent path traversal and only serve files from OUTPUT_FOLDER
    safe_name = os.path.basename(filename)
    out_dir = os.path.abspath(OUTPUT_FOLDER)
    file_path = os.path.join(out_dir, safe_name)
    if not os.path.exists(file_path):
        return ("File not found"), 404
    # Use send_from_directory to let Flask set headers correctly
    return send_from_directory(out_dir, safe_name, as_attachment=True)
