# AMA DailyReport

Small Flask app that parses a construction timesheet Excel file and returns a summarized Excel file.

Contents added to help deploy or package the app locally and for others to run.

Quick start (development)

1. Create a virtual environment and activate it:

   On Windows (PowerShell):

   ```powershell
   python -m venv .venv; .\.venv\\Scripts\\Activate.ps1
   pip install -r requirements.txt
   ```

2. Run the app:

   ```powershell
   python app.py
   ```

3. Open http://127.0.0.1:5000 and upload an .xlsx timesheet.

Deploy as a Docker container

1. Build the image:

   ```powershell
   docker build -t amadaily:latest .
   ```

2. Run the container:

   ```powershell
   docker run -p 5000:5000 amadaily:latest
   ```

Deploy to Render / Heroku / similar

- Ensure `requirements.txt` and `Procfile` are in the repo. The Procfile tells the platform to run gunicorn. Push to the service and follow their steps.

Make a Windows executable (standalone) using PyInstaller

This produces a single-folder or single-file executable others can run without Python installed.

1. Install dev deps (in your activated venv):

   ```powershell
   pip install pyinstaller
   ```

2. Create a spec or run a simple build (single-folder is safer for pandas):

   ```powershell
   pyinstaller --add-data "templates;templates" --add-data "uploads;uploads" --add-data "outputs;outputs" --onedir app.py
   ```

3. The result will be in the `dist/app` folder. Share that folder with users.

Notes & next steps

- Consider adding authentication or limiting uploads for a public deployment.
- Add simple unit tests for `parse_timesheet.py` for safer refactors.
- For a desktop UI experience, consider packaging the parser into a small GUI using PySimpleGUI or Electron (web wrapper).

Security and runtime notes

- The app currently accepts and processes uploaded Excel files. If you deploy publicly, add these protections:
   - File size limits and content-type checks.
   - A temporary upload area and periodic cleanup of `uploads/` to avoid disk growth.
   - Run behind HTTPS and enable basic auth or a login for restricted use.

Choosing between deploy vs. distribution

- Use Docker + a VPS or Render for a hosted website that multiple people can access.
- Use PyInstaller to create a portable folder to share with teammates who want an offline, double-clickable tool.

If you'd like, I can:

- Create a small PowerShell script that automates the PyInstaller build for Windows.
- Add a basic Gunicorn systemd service file example for a Linux server.
- Add simple unit tests for `parse_timesheet.py`.

