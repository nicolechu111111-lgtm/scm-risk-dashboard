from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components


ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / "tools" / "build_scm_html_mvp.py"
DATA_DIR = ROOT / "streamlit_data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUT_DIR = DATA_DIR / "live"
HTML_NAME = "SCM Risk Dashboard MVP 2026-06-20.html"


st.set_page_config(page_title="SCM 看板", layout="wide")


def check_password() -> bool:
    password = st.secrets.get("SCM_DASHBOARD_PASSWORD", "")
    if not password:
        return True
    if st.session_state.get("authenticated"):
        return True
    st.title("SCM 看板")
    entered = st.text_input("共享密码", type="password")
    if st.button("进入看板"):
        if entered == password:
            st.session_state.authenticated = True
            st.rerun()
        st.error("密码不正确。")
    return False


def latest_upload() -> Path | None:
    if not UPLOAD_DIR.exists():
        return None
    files = sorted(UPLOAD_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def live_html() -> Path:
    return OUT_DIR / HTML_NAME


def recalculate(workbook: Path) -> tuple[bool, str]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["SCM_WORKBOOK"] = str(workbook)
    env["SCM_OUT_DIR"] = str(OUT_DIR)
    env["SCM_TODAY"] = datetime.now().strftime("%Y-%m-%d")
    result = subprocess.run(
        [sys.executable, str(BUILDER)],
        cwd=str(ROOT),
        env=env,
        text=True,
        capture_output=True,
    )
    if result.returncode != 0:
        return False, result.stderr[-3000:] or result.stdout[-3000:] or "重新计算失败。"
    return True, "重新计算完成。"


if not check_password():
    st.stop()

st.title("SCM Risk Dashboard / 供应链风险看板")

with st.sidebar:
    st.header("上传 Follow Up")
    st.caption("上传最新版 Follow Up 后，系统会重新计算看板。")
    uploaded = st.file_uploader("选择 Follow Up 表格", type=["xlsx", "xlsm"])
    if uploaded and st.button("上传并重新计算", type="primary"):
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = Path(uploaded.name).name.replace("/", "_")
        target = UPLOAD_DIR / f"{stamp}_{safe_name}"
        target.write_bytes(uploaded.getbuffer())
        ok, message = recalculate(target)
        if ok:
            st.success(message)
            st.rerun()
        else:
            st.error(message)

    current = latest_upload()
    if current:
        st.caption(f"当前文件：{current.name}")
    st.caption("免费部署版本的文件保存在应用运行环境内；如果平台休眠或重建，可能需要重新上传。")

if not live_html().exists():
    current = latest_upload()
    if current:
        ok, message = recalculate(current)
        if not ok:
            st.error(message)
            st.stop()
    else:
        st.info("请先在左侧上传最新版 Follow Up。")
        st.stop()

html = live_html().read_text(encoding="utf-8")
components.html(html, height=1200, scrolling=True)
