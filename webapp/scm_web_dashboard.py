from __future__ import annotations

import cgi
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


ROOT = Path(os.environ.get("SCM_ROOT", Path(__file__).resolve().parents[1]))
BUILDER = ROOT / "tools" / "build_scm_html_mvp.py"
DATA_DIR = Path(os.environ.get("SCM_DATA_DIR", ROOT / "webapp_data"))
OUT_DIR = DATA_DIR / "live"
UPLOAD_DIR = DATA_DIR / "uploads"
HTML_NAME = "SCM Risk Dashboard MVP 2026-06-20.html"
JSON_NAME = "scm_risk_dashboard_data_2026-06-20.json"
PASSWORD = os.environ.get("SCM_DASHBOARD_PASSWORD", "").strip()


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def html_response(handler: BaseHTTPRequestHandler, status: int, text: str):
    body = text.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def live_html() -> Path:
    return OUT_DIR / HTML_NAME


def live_json() -> Path:
    return OUT_DIR / JSON_NAME


def latest_upload() -> Path | None:
    if not UPLOAD_DIR.exists():
        return None
    files = sorted(UPLOAD_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def recalculate(workbook: Path) -> subprocess.CompletedProcess:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["SCM_WORKBOOK"] = str(workbook)
    env["SCM_OUT_DIR"] = str(OUT_DIR)
    env["SCM_TODAY"] = datetime.now().strftime("%Y-%m-%d")
    return subprocess.run(
        [sys.executable, str(BUILDER)],
        cwd=str(ROOT),
        env=env,
        text=True,
        capture_output=True,
    )


def ensure_initial_dashboard():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    upload = latest_upload()
    if upload and not live_html().exists():
        recalculate(upload)


def is_authenticated(handler: BaseHTTPRequestHandler) -> bool:
    if not PASSWORD:
        return True
    raw = handler.headers.get("Cookie", "")
    jar = cookies.SimpleCookie(raw)
    return jar.get("scm_auth") and jar["scm_auth"].value == PASSWORD


def login_page(message: str = "") -> str:
    note = f"<p class='msg'>{message}</p>" if message else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SCM 看板登录</title>
  <style>
    body {{ margin:0; font-family: Arial, sans-serif; background:#f4f7fb; color:#172231; }}
    main {{ max-width:420px; margin:14vh auto; background:#fff; border:1px solid #d9e4f0; border-radius:8px; padding:28px; }}
    h1 {{ margin:0 0 10px; font-size:26px; }}
    p {{ color:#607082; }}
    .msg {{ color:#b42318; }}
    input, button {{ box-sizing:border-box; width:100%; font-size:16px; padding:12px; border-radius:6px; }}
    input {{ border:1px solid #cfdbea; margin:12px 0; }}
    button {{ border:0; background:#2d6693; color:#fff; font-weight:700; cursor:pointer; }}
  </style>
</head>
<body>
  <main>
    <h1>SCM 看板</h1>
    <p>请输入共享密码后查看或更新看板。</p>
    {note}
    <form method="post" action="/login">
      <input name="password" type="password" placeholder="共享密码" autofocus>
      <button type="submit">进入看板</button>
    </form>
  </main>
</body>
</html>"""


def setup_page() -> str:
    return """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>上传 Follow Up</title>
  <style>
    body { margin:0; font-family: Arial, sans-serif; background:#f4f7fb; color:#172231; }
    main { max-width:760px; margin:10vh auto; background:#fff; border:1px solid #d9e4f0; border-radius:8px; padding:28px; }
    h1 { margin:0 0 10px; font-size:28px; }
    p { color:#607082; line-height:1.6; }
    input, button { font-size:16px; }
    button { padding:10px 16px; border:1px solid #cfdbea; border-radius:6px; background:#2d6693; color:#fff; cursor:pointer; }
  </style>
</head>
<body>
  <main>
    <h1>请先上传 Follow Up</h1>
    <p>在线版看板需要先上传一份最新的 Follow Up 表格。上传后系统会自动重新计算，所有人刷新页面都能看到同一版数据。</p>
    <form method="post" action="/api/upload_followup" enctype="multipart/form-data">
      <input type="file" name="followup" accept=".xlsx,.xlsm" required>
      <button type="submit">上传并生成看板</button>
    </form>
  </main>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    server_version = "SCMWebDashboard/1.0"

    def log_message(self, fmt, *args):
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def require_auth(self) -> bool:
        if is_authenticated(self):
            return True
        html_response(self, 401, login_page())
        return False

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/health":
            json_response(self, 200, {"ok": True})
            return
        if path == "/login":
            html_response(self, 200, login_page())
            return
        if not self.require_auth():
            return
        if path in {"/", "/dashboard"}:
            ensure_initial_dashboard()
            if not live_html().exists():
                html_response(self, 200, setup_page())
                return
            body = live_html().read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/data.json":
            if not live_json().exists():
                json_response(self, 404, {"ok": False, "error": "还没有生成看板数据。"})
                return
            body = live_json().read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/login":
            length = int(self.headers.get("Content-Length", "0") or "0")
            fields = parse_qs(self.rfile.read(length).decode("utf-8"))
            password = fields.get("password", [""])[0]
            if not PASSWORD or password == PASSWORD:
                self.send_response(302)
                self.send_header("Location", "/dashboard")
                self.send_header("Set-Cookie", f"scm_auth={password}; Path=/; SameSite=Lax")
                self.end_headers()
                return
            html_response(self, 401, login_page("密码不正确。"))
            return
        if not self.require_auth():
            return
        if path != "/api/upload_followup":
            self.send_error(404)
            return
        ctype, _ = cgi.parse_header(self.headers.get("Content-Type", ""))
        if ctype != "multipart/form-data":
            json_response(self, 400, {"ok": False, "error": "请上传 Excel 文件。"})
            return
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
        )
        item = form["followup"] if "followup" in form else None
        if item is None or not item.filename:
            json_response(self, 400, {"ok": False, "error": "没有收到 Follow Up 文件。"})
            return
        if not item.filename.lower().endswith((".xlsx", ".xlsm")):
            json_response(self, 400, {"ok": False, "error": "请上传 .xlsx 或 .xlsm 文件。"})
            return
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = Path(item.filename).name.replace("/", "_")
        uploaded = UPLOAD_DIR / f"{stamp}_{safe_name}"
        with uploaded.open("wb") as f:
            shutil.copyfileobj(item.file, f)
        result = recalculate(uploaded)
        if result.returncode != 0:
            json_response(
                self,
                500,
                {
                    "ok": False,
                    "error": result.stderr[-3000:] or result.stdout[-3000:] or "重新计算失败。",
                },
            )
            return
        json_response(self, 200, {"ok": True, "url": "/dashboard", "uploaded": str(uploaded)})


def main():
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("SCM_PORT", "8080")))
    ensure_initial_dashboard()
    print(f"SCM Web Dashboard: http://{host}:{port}/dashboard")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    main()
