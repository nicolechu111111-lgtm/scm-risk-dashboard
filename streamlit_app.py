from __future__ import annotations

import os
import json
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components


ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / "tools" / "build_scm_html_mvp.py"
DATA_DIR = ROOT / "streamlit_data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUT_DIR = DATA_DIR / "live"
HTML_NAME = "SCM Risk Dashboard MVP 2026-06-20.html"
JSON_NAME = "scm_risk_dashboard_data_2026-06-20.json"
EMAIL_SENT_PATH = DATA_DIR / "warehouse_email_sent.json"


st.set_page_config(page_title="SCM 看板", layout="wide")


def check_password() -> bool:
    configured_password = str(
        st.secrets.get("SCM_DASHBOARD_PASSWORD", "")
        or os.environ.get("SCM_DASHBOARD_PASSWORD", "")
        or ""
    ).strip()
    fallback_password = str(os.environ.get("SCM_FALLBACK_PASSWORD", "scm2026")).strip()
    allowed_passwords = {p for p in (configured_password, fallback_password) if p}
    if not allowed_passwords:
        return True
    if st.session_state.get("authenticated"):
        return True
    st.title("SCM 看板")
    entered = st.text_input("共享密码", type="password")
    if st.button("进入看板"):
        if entered.strip() in allowed_passwords:
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


def live_json() -> Path:
    return OUT_DIR / JSON_NAME


def load_json_file(path: Path, fallback):
    try:
        if not path.exists():
            return fallback
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback


def save_json_file(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_email_sent() -> dict:
    data = load_json_file(EMAIL_SENT_PATH, {})
    return data if isinstance(data, dict) else {}


def save_email_sent(data: dict) -> None:
    save_json_file(EMAIL_SENT_PATH, data)


def recalculate(workbook: Path) -> tuple[bool, str]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["SCM_WORKBOOK"] = str(workbook)
    env["SCM_OUT_DIR"] = str(OUT_DIR)
    env["SCM_TODAY"] = datetime.now().strftime("%Y-%m-%d")
    env["SCM_SHARED_EMAIL_MODE"] = "1"
    env["SCM_EMAIL_SENT_JSON"] = str(EMAIL_SENT_PATH)
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


def date_text(value: str, delta_days: int = 0) -> str:
    try:
        d = datetime.strptime(str(value), "%Y-%m-%d").date()
    except Exception:
        return ""
    return (d + timedelta(days=delta_days)).isoformat()


def day_diff(left: str, right: str) -> int | None:
    try:
        l_date = datetime.strptime(str(left), "%Y-%m-%d").date()
        r_date = datetime.strptime(str(right), "%Y-%m-%d").date()
    except Exception:
        return None
    return (l_date - r_date).days


def warehouse_email_rows(data: dict) -> list[dict]:
    generated_at = data.get("generated_at", datetime.now().strftime("%Y-%m-%d"))
    sent = load_email_sent()
    rows = []
    for so in data.get("so_board", []):
        so_no = str(so.get("so", ""))
        email_due = date_text(so.get("required_arrival", ""), -21)
        days = day_diff(email_due, generated_at) if email_due else None
        is_sent = so_no in sent
        if is_sent:
            status = "已发送"
        elif days is not None and days < 0:
            status = "已逾期"
        elif days == 0:
            status = "今日到期"
        elif days is not None and days <= 7:
            status = "即将到期"
        else:
            status = "未到期"
        rows.append(
            {
                "状态": status,
                "SO": so_no,
                "客户": so.get("customer", ""),
                "客户仓": so.get("delivery_center", ""),
                "客户要求到仓日": so.get("required_arrival", ""),
                "邮件提醒日": email_due,
                "剩余天数": "" if days is None else days,
                "风险": so.get("status", ""),
                "问题SKU数": so.get("issue_skus", 0),
                "sent_at": sent.get(so_no, {}).get("sent_at", ""),
            }
        )
    rank = {"已逾期": 0, "今日到期": 1, "即将到期": 2, "未到期": 3, "已发送": 4}
    rows.sort(key=lambda r: (rank.get(r["状态"], 9), str(r["邮件提醒日"] or "9999-12-31"), str(r["SO"])))
    return rows


def render_shared_email_controls(data: dict) -> None:
    rows = warehouse_email_rows(data)
    active = [r for r in rows if r["状态"] in {"已逾期", "今日到期", "即将到期"}]
    sent = [r for r in rows if r["状态"] == "已发送"]

    st.sidebar.divider()
    st.sidebar.header("仓库邮件共享确认")
    st.sidebar.caption("这里的标记会同步给所有打开线上看板的人。")
    if not active:
        st.sidebar.success("暂无需要确认发送的仓库邮件。")
    for row in active[:20]:
        label = f"{row['SO']} · {row['客户仓']} · {row['邮件提醒日']}"
        with st.sidebar.expander(label):
            st.write(f"客户要求到仓日：{row['客户要求到仓日']}")
            st.write(f"风险：{row['风险']} / 问题SKU数：{row['问题SKU数']}")
            if st.button("标记已发送", key=f"email_sent_{row['SO']}"):
                sent_map = load_email_sent()
                sent_map[row["SO"]] = {"so": row["SO"], "sent_at": datetime.now().isoformat(timespec="seconds")}
                save_email_sent(sent_map)
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()

    with st.sidebar.expander(f"已发送记录 ({len(sent)})"):
        if not sent:
            st.caption("暂无已发送记录。")
        for row in sent[:30]:
            cols = st.columns([3, 1])
            cols[0].caption(f"{row['SO']} · {row['sent_at']}")
            if cols[1].button("撤销", key=f"email_unsent_{row['SO']}"):
                sent_map = load_email_sent()
                sent_map.pop(row["SO"], None)
                save_email_sent(sent_map)
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()


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

live_data = load_json_file(live_json(), {})
if live_data:
    render_shared_email_controls(live_data)

html = live_html().read_text(encoding="utf-8")
components.html(html, height=1200, scrolling=True)
