#!/usr/bin/env python3
"""
V4.3.3.5 – Crawler-Server
Startet einen lokalen HTTP-Server auf Port 8765, damit der Dashboard-Button
den Crawler direkt auslösen kann.

Verwendung:
  python3 crawler_server.py
  → Dann im Dashboard den Button "🌍 Länder aktualisieren" klicken.
"""

import json, threading, sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

# Import des Crawlers
sys.path.insert(0, __file__.rsplit("/", 1)[0])
from update_country_data import run_all

PORT = 8765
_job_state = {"running": False, "progress": [], "done": False, "error": None}
_lock = threading.Lock()


def progress_cb(info: dict):
    with _lock:
        _job_state["progress"].append(info)


def run_crawler_thread():
    with _lock:
        _job_state["running"] = True
        _job_state["progress"] = []
        _job_state["done"] = False
        _job_state["error"] = None
    try:
        run_all(dry_run=False, progress_cb=progress_cb)
    except Exception as e:
        with _lock:
            _job_state["error"] = str(e)
    finally:
        with _lock:
            _job_state["running"] = False
            _job_state["done"] = True


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # Konsole nicht zumüllen

    def send_json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/status":
            with _lock:
                self.send_json(200, dict(_job_state))
        elif path == "/ping":
            self.send_json(200, {"ok": True, "version": "4.3.3.5"})
        else:
            self.send_json(404, {"error": "Not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/run":
            with _lock:
                if _job_state["running"]:
                    self.send_json(409, {"error": "Crawler läuft bereits"})
                    return
            t = threading.Thread(target=run_crawler_thread, daemon=True)
            t.start()
            self.send_json(200, {"ok": True, "message": "Crawler gestartet"})
        else:
            self.send_json(404, {"error": "Not found"})


if __name__ == "__main__":
    server = HTTPServer(("localhost", PORT), Handler)
    print(f"""
╔══════════════════════════════════════════════════╗
║   V4.3.3.5  Crawler-Server läuft                ║
║   http://localhost:{PORT}                          ║
║   Im Dashboard auf "🌍 Länder aktualisieren"     ║
║   klicken, um den Crawler zu starten.            ║
║   Beenden mit: Ctrl+C                            ║
╚══════════════════════════════════════════════════╝
""")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer beendet.")
