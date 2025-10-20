"""Launcher for a desktop build of the Flask app using pywebview.

Starts the Flask app on a free localhost port and opens a native window
pointing at the app. Intended for packaging with PyInstaller.
"""
from __future__ import annotations

import socket
from threading import Thread
import os

try:
    import webview
    _HAVE_WEBVIEW = True
except Exception:
    _HAVE_WEBVIEW = False
    import webbrowser

# Import the Flask app instance from your app module
from app import app


def _find_free_port() -> int:
    s = socket.socket()
    s.bind(("", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _run_server(port: int) -> None:
    # For local desktop usage the Flask dev server is acceptable. Disable
    # reloader to avoid spawning extra processes when packaged.
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


def main() -> None:
    port = _find_free_port()
    # Start the server in a non-daemon thread so the process stays alive
    # while the web UI is open (fallback path joins this thread).
    t = Thread(target=_run_server, args=(port,), daemon=False)
    t.start()

    url = f"http://127.0.0.1:{port}/"
    # Create a window and start the GUI loop. Size can be adjusted.
    if _HAVE_WEBVIEW:
        webview.create_window("AMA Combined Daily Report", url, width=1100, height=700)
        webview.start()
    else:
        # Fallback: open default browser and keep the server running
        print("pywebview not available, opening default browser instead.")
        webbrowser.open(url)
        # Block here to keep the Flask server thread alive
        t.join()


if __name__ == "__main__":
    main()
