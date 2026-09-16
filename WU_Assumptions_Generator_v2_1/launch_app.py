"""Windows-friendly launcher that opens the local Flask UI in a browser."""
from __future__ import annotations

import threading
import time
import webbrowser

from app import app

URL = "http://127.0.0.1:5051"


def _open_browser() -> None:
    time.sleep(1.2)
    webbrowser.open(URL)


if __name__ == "__main__":
    threading.Thread(target=_open_browser, daemon=True).start()
    print(f"Lower Basin Water Use Report Generator: {URL}")
    print("Keep this window open while using the application. Press Ctrl+C to stop it.")
    app.run(host="127.0.0.1", port=5051, debug=False, use_reloader=False)
