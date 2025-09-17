from flask import Flask, render_template, request, send_file, abort
import os
from werkzeug.utils import secure_filename
from parse_timesheet import process_timesheet  # import your function

# Create a Flask app instance
app = Flask(__name__)

# Upload safety config
# limit request size to 50 MB by default (can override via env var MAX_CONTENT_LENGTH)
default_max = int(os.environ.get("MAX_CONTENT_LENGTH", 50 * 1024 * 1024))
app.config["MAX_CONTENT_LENGTH"] = default_max

# allowed file extensions
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

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


@app.route("/upload", methods=["POST"])  # handle form submissions
def upload_file():
    # Check the form actually had a file in it
    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]

    if file.filename == "":
        return "No file selected", 400

    if not allowed_file(file.filename):
        return "Unsupported file type", 400

    # Use a secure filename and save to uploads/
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    try:
        file.save(filepath)
    except Exception:
        return "Failed to save file", 500

    # Run your parser on the uploaded file
    try:
        summary_file = process_timesheet(filepath, OUTPUT_FOLDER)
    except Exception as e:
        # If processing fails, return an error
        return f"Processing failed: {e}", 500

    # Send the summary Excel back to the user as a download
    return send_file(summary_file, as_attachment=True)


if __name__ == "__main__":
    # Run the Flask development server
    app.run(debug=True)
