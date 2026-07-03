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
SHARED_STATE_PATH = DATA_DIR / "shared_state.json"
UPLOAD_HISTORY_PATH = DATA_DIR / "upload_history.json"
OPERATION_LOG_PATH = DATA_DIR / "operation_log.json"


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


def default_shared_state() -> dict:
    return {"manual_allocations": {}, "transit_settings": {"stateDays": {}, "dcDays": {}}, "confirmed_sps_imports": []}


def load_shared_state() -> dict:
    data = load_json_file(SHARED_STATE_PATH, default_shared_state())
    if not isinstance(data, dict):
        data = default_shared_state()
    data.setdefault("manual_allocations", {})
    data.setdefault("transit_settings", {"stateDays": {}, "dcDays": {}})
    data.setdefault("confirmed_sps_imports", [])
    if not isinstance(data["manual_allocations"], dict):
        data["manual_allocations"] = {}
    if not isinstance(data["transit_settings"], dict):
        data["transit_settings"] = {"stateDays": {}, "dcDays": {}}
    data["transit_settings"].setdefault("stateDays", {})
    data["transit_settings"].setdefault("dcDays", {})
    if not isinstance(data["confirmed_sps_imports"], list):
        data["confirmed_sps_imports"] = []
    return data


def save_shared_state(data: dict) -> None:
    save_json_file(SHARED_STATE_PATH, data)


def append_log(action: str, detail: str = "") -> None:
    rows = load_json_file(OPERATION_LOG_PATH, [])
    if not isinstance(rows, list):
        rows = []
    rows.insert(0, {"time": datetime.now().isoformat(timespec="seconds"), "action": action, "detail": detail})
    save_json_file(OPERATION_LOG_PATH, rows[:300])


def append_upload_history(filename: str) -> None:
    rows = load_json_file(UPLOAD_HISTORY_PATH, [])
    if not isinstance(rows, list):
        rows = []
    rows.insert(0, {"time": datetime.now().isoformat(timespec="seconds"), "file": filename})
    save_json_file(UPLOAD_HISTORY_PATH, rows[:80])


def recalculate(workbook: Path) -> tuple[bool, str]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["SCM_WORKBOOK"] = str(workbook)
    env["SCM_OUT_DIR"] = str(OUT_DIR)
    env["SCM_TODAY"] = datetime.now().strftime("%Y-%m-%d")
    env["SCM_SHARED_EMAIL_MODE"] = "1"
    env["SCM_EMAIL_SENT_JSON"] = str(EMAIL_SENT_PATH)
    env["SCM_SHARED_STATE_MODE"] = "1"
    env["SCM_SHARED_STATE_JSON"] = str(SHARED_STATE_PATH)
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
                append_log("仓库邮件标记已发送", str(row["SO"]))
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
                append_log("撤销仓库邮件已发送", str(row["SO"]))
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()


def allocation_key(sku: str, so: str) -> str:
    return f"{sku}__{so}"


def render_shared_allocation_controls(data: dict) -> None:
    lines = data.get("so_lines", [])
    if not lines:
        return
    state = load_shared_state()
    allocations = state["manual_allocations"]
    sku_options = sorted({str(x.get("sku", "")) for x in lines if x.get("sku")})

    st.sidebar.divider()
    st.sidebar.header("人工分配共享管理")
    st.sidebar.caption("这里保存后，所有人刷新都会看到同一套人工分配。")
    selected_sku = st.sidebar.selectbox("选择 SKU", sku_options, key="shared_alloc_sku")
    sku_lines = [x for x in lines if str(x.get("sku", "")) == selected_sku]
    sku_info = next((x for x in data.get("sku_summary", []) if str(x.get("sku", "")) == selected_sku), {})
    if not sku_info:
        sku_info = data.get("inventory_by_sku", {}).get(selected_sku, {})
    stock_limit = int(float(sku_info.get("current_onhand", 0) or 0))
    assigned = 0
    for line in sku_lines:
        saved = allocations.get(allocation_key(selected_sku, str(line.get("so", ""))), {})
        raw = str(saved.get("assign_qty", "")).strip()
        if raw:
            try:
                assigned += int(float(raw))
            except Exception:
                pass
    st.sidebar.caption(f"当前库存：{stock_limit} / 人工已分配：{assigned} / 剩余：{max(stock_limit - assigned, 0)}")

    with st.sidebar.expander("编辑该 SKU 的 SO 分配", expanded=False):
        for line in sku_lines[:80]:
            so = str(line.get("so", ""))
            key = allocation_key(selected_sku, so)
            saved = allocations.get(key, {})
            label = f"{so} · {line.get('delivery_center','')} · 需求 {line.get('qty', '')}"
            st.caption(label)
            c1, c2 = st.columns([1, 1])
            default_qty = saved.get("assign_qty", "")
            qty_value = c1.text_input("人工分配库存", value=str(default_qty), key=f"alloc_qty_{key}", placeholder="空=按系统")
            note_value = c2.text_input("备注", value=str(saved.get("note", "")), key=f"alloc_note_{key}")
            if st.button("保存这一行", key=f"save_alloc_{key}"):
                qty_text = str(qty_value).strip()
                if qty_text:
                    try:
                        qty_num = max(int(float(qty_text)), 0)
                    except Exception:
                        st.error("人工分配库存必须是数字。")
                        st.stop()
                    if qty_num > int(float(line.get("qty", 0) or 0)):
                        st.error("人工分配不能超过该 SO 的需求数量。")
                        st.stop()
                    allocations[key] = {
                        "sku": selected_sku,
                        "so": so,
                        "assign": "manual",
                        "assign_qty": str(qty_num),
                        "note": note_value,
                        "updated_at": datetime.now().isoformat(timespec="seconds"),
                    }
                elif note_value.strip():
                    allocations[key] = {
                        "sku": selected_sku,
                        "so": so,
                        "assign": "system",
                        "assign_qty": "",
                        "note": note_value,
                        "updated_at": datetime.now().isoformat(timespec="seconds"),
                    }
                else:
                    allocations.pop(key, None)
                state["manual_allocations"] = allocations
                save_shared_state(state)
                append_log("更新人工分配", f"{selected_sku} / {so} / {qty_text or '按系统'}")
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()
        if st.button("清空该 SKU 的人工分配", key=f"clear_alloc_{selected_sku}"):
            for key in list(allocations):
                if allocations[key].get("sku") == selected_sku:
                    allocations.pop(key, None)
            state["manual_allocations"] = allocations
            save_shared_state(state)
            append_log("清空 SKU 人工分配", selected_sku)
            current = latest_upload()
            if current:
                recalculate(current)
            st.rerun()


DEFAULT_TRANSIT_DAYS = {"CA": 2, "NV": 3, "AZ": 3, "OR": 4, "WA": 4, "UT": 5, "CO": 5, "TX": 5, "IL": 7, "GA": 8, "SC": 8, "NC": 8, "NJ": 9, "PA": 9, "NY": 9, "MD": 9, "VA": 9, "MA": 10, "CT": 10, "FL": 10}


def dc_key(customer: str, dc: str) -> str:
    return f"{customer}__{dc}"


def render_shared_transit_controls(data: dict) -> None:
    state = load_shared_state()
    settings = state["transit_settings"]
    settings.setdefault("stateDays", {})
    settings.setdefault("dcDays", {})
    lines = data.get("so_lines", [])
    dc_rows = []
    seen = set()
    for line in lines:
        customer = str(line.get("customer", ""))
        dc = str(line.get("delivery_center", ""))
        if not dc:
            continue
        key = dc_key(customer, dc)
        if key in seen:
            continue
        seen.add(key)
        dc_rows.append({"key": key, "customer": customer, "dc": dc, "days": settings["dcDays"].get(key, "")})

    st.sidebar.divider()
    st.sidebar.header("运输设置共享管理")
    st.sidebar.caption("客户仓运输天数会同步给所有人。")
    with st.sidebar.expander("客户仓天数", expanded=False):
        for row in dc_rows[:80]:
            value = st.text_input(f"{row['customer']} · {row['dc']}", value=str(row["days"]), key=f"dc_days_{row['key']}", placeholder="空=按默认")
            if st.button("保存", key=f"save_dc_{row['key']}"):
                raw = str(value).strip()
                if raw:
                    try:
                        settings["dcDays"][row["key"]] = max(int(float(raw)), 0)
                    except Exception:
                        st.error("运输天数必须是数字。")
                        st.stop()
                else:
                    settings["dcDays"].pop(row["key"], None)
                state["transit_settings"] = settings
                save_shared_state(state)
                append_log("更新客户仓运输天数", f"{row['dc']} = {raw or '默认'}")
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()
    with st.sidebar.expander("州默认天数", expanded=False):
        for st_code, default_days in DEFAULT_TRANSIT_DAYS.items():
            value = st.text_input(st_code, value=str(settings["stateDays"].get(st_code, default_days)), key=f"state_days_{st_code}")
            if st.button("保存州设置", key=f"save_state_{st_code}"):
                try:
                    settings["stateDays"][st_code] = max(int(float(value)), 0)
                except Exception:
                    st.error("运输天数必须是数字。")
                    st.stop()
                state["transit_settings"] = settings
                save_shared_state(state)
                append_log("更新州运输天数", f"{st_code} = {settings['stateDays'][st_code]}")
                current = latest_upload()
                if current:
                    recalculate(current)
                st.rerun()


def render_history_and_logs() -> None:
    st.sidebar.divider()
    st.sidebar.header("版本与操作记录")
    uploads = load_json_file(UPLOAD_HISTORY_PATH, [])
    logs = load_json_file(OPERATION_LOG_PATH, [])
    with st.sidebar.expander(f"上传历史 ({len(uploads) if isinstance(uploads, list) else 0})"):
        if not uploads:
            st.caption("暂无上传记录。")
        for row in (uploads if isinstance(uploads, list) else [])[:20]:
            st.caption(f"{row.get('time','')} · {row.get('file','')}")
    with st.sidebar.expander(f"操作日志 ({len(logs) if isinstance(logs, list) else 0})"):
        if not logs:
            st.caption("暂无操作记录。")
        for row in (logs if isinstance(logs, list) else [])[:30]:
            st.caption(f"{row.get('time','')} · {row.get('action','')} · {row.get('detail','')}")


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
            append_upload_history(target.name)
            append_log("上传 Follow Up", target.name)
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
    render_shared_allocation_controls(live_data)
    render_shared_transit_controls(live_data)
    render_history_and_logs()

html = live_html().read_text(encoding="utf-8")
components.html(html, height=1200, scrolling=True)
