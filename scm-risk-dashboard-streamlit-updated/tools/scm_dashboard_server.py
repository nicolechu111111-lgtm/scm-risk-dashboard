from __future__ import annotations

import cgi
import json
import os
import shutil
import subprocess
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path("/Users/blue/Documents/跨境供应链跟单")
PYTHON = Path("/Users/blue/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3")
BUILDER = ROOT / "tools" / "build_scm_html_mvp.py"
OUT_DIR = ROOT / "outputs" / "scm_html_mvp_live"
UPLOAD_DIR = ROOT / "outputs" / "scm_html_mvp_live_uploads"
HTML_NAME = "SCM Risk Dashboard MVP 2026-06-20.html"
DEFAULT_HTML = ROOT / "outputs" / "scm_html_mvp_20260620" / HTML_NAME


def dashboard_html() -> Path:
    live = OUT_DIR / HTML_NAME
    return live if live.exists() else DEFAULT_HTML


def send_json(handler: BaseHTTPRequestHandler, status: int, payload: dict):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


class Handler(BaseHTTPRequestHandler):
    server_version = "SCMDashboard/0.1"

    def log_message(self, fmt, *args):
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def do_GET(self):
        path = urlparse(self.path).path
        if path in {"/", "/dashboard", f"/{HTML_NAME}"}:
            html = dashboard_html()
            body = html.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/data.json":
            p = OUT_DIR / "scm_risk_dashboard_data_2026-06-20.json"
            if not p.exists():
                p = ROOT / "outputs" / "scm_html_mvp_20260620" / "scm_risk_dashboard_data_2026-06-20.json"
            body = p.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path != "/api/upload_followup":
            self.send_error(404)
            return

        ctype, pdict = cgi.parse_header(self.headers.get("Content-Type", ""))
        if ctype != "multipart/form-data":
            send_json(self, 400, {"ok": False, "error": "Expected multipart upload."})
            return
        pdict["boundary"] = bytes(pdict["boundary"], "utf-8")
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")})
        item = form["followup"] if "followup" in form else None
        if item is None or not item.filename:
            send_json(self, 400, {"ok": False, "error": "No Follow Up workbook uploaded."})
            return
        if not item.filename.lower().endswith((".xlsx", ".xlsm")):
            send_json(self, 400, {"ok": False, "error": "Please upload an .xlsx or .xlsm file."})
            return

        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = Path(item.filename).name.replace("/", "_")
        uploaded = UPLOAD_DIR / f"{stamp}_{safe_name}"
        with uploaded.open("wb") as f:
            shutil.copyfileobj(item.file, f)

        env = os.environ.copy()
        env["SCM_WORKBOOK"] = str(uploaded)
        env["SCM_OUT_DIR"] = str(OUT_DIR)
        env["SCM_TODAY"] = datetime.now().strftime("%Y-%m-%d")
        result = subprocess.run([str(PYTHON), str(BUILDER)], cwd=str(ROOT), env=env, text=True, capture_output=True)
        if result.returncode != 0:
            send_json(self, 500, {"ok": False, "error": result.stderr[-2000:] or result.stdout[-2000:] or "Recalculation failed."})
            return
        send_json(self, 200, {"ok": True, "url": "/dashboard", "uploaded": str(uploaded)})


def main():
    host = "127.0.0.1"
    port = int(os.environ.get("SCM_DASHBOARD_PORT", "8765"))
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"SCM dashboard server running at http://{host}:{port}/dashboard")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
