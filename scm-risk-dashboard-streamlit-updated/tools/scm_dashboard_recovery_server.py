from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path


import os


HOST = "127.0.0.1"
PORT = int(os.environ.get("SCM_PORT", "8765"))
HTML_PATH = Path(
    "/Users/blue/Documents/跨境供应链跟单/outputs/scm_html_mvp_live2/SCM Risk Dashboard MVP 2026-06-20.html"
)


class Handler(BaseHTTPRequestHandler):
    def send_dashboard(self, include_body=True):
        body = HTML_PATH.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if include_body:
            self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/dashboard"):
            self.send_dashboard(include_body=True)
            return
        self.send_response(404)
        self.end_headers()

    def do_HEAD(self):
        if self.path in ("/", "/dashboard"):
            self.send_dashboard(include_body=False)
            return
        self.send_response(404)
        self.end_headers()

    def log_message(self, fmt, *args):
        return


if __name__ == "__main__":
    print(f"http://{HOST}:{PORT}/dashboard")
    HTTPServer((HOST, PORT), Handler).serve_forever()
