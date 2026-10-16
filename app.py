"""
Export invoice automation - Streamlit dashboard
Starts rpa_customs.py (Update Customs No) and rpa_erp.py (Invoice Header) as background runs
(runner.py) and shows live progress, per-invoice results, screenshots and run history.
Runs keep going - and are saved - even if you switch tabs, refresh the page or restart Streamlit.

Start:  streamlit run app.py
"""
import ast
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import psutil
import streamlit as st
from dotenv import load_dotenv

from excel_io import build_template, convert_raw, read_invoice_file

BASE = Path(__file__).parent
load_dotenv(BASE / ".env")

RUNS = BASE / "runs"
RUNS.mkdir(exist_ok=True)
SCRIPT_CUSTOMS = BASE / "rpa_customs.py"
SCRIPT_HEADER = BASE / "rpa_erp.py"

st.set_page_config(page_title="Export invoice automation", page_icon=":material/receipt_long:", layout="wide")
st.markdown(
    """
    <style>
    .block-container {padding-top: 2.2rem; max-width: 1180px;}
    h1 {font-weight: 650; letter-spacing: -0.01em; margin-bottom: 0.2rem;}
    [data-testid="stMetricValue"] {font-variant-numeric: tabular-nums;}
    .subtle {color: #5B6B7A; margin-bottom: 1.4rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Hosted on Linux (Streamlit Cloud / Docker): keep the Playwright browser in a folder the app controls,
# so the dashboard and the runs it starts always look in the same place. A value that is already set
# (for example by the Windows installer's start_dashboard.bat) is kept.
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(BASE / ".pw-browsers"))


@st.cache_resource(show_spinner="Installing the browser for the automation (first start only, about a minute)...")
def ensure_browser():
    """Download Chromium once per server start. Only on Linux hosting; on Windows it is already installed."""
    if not sys.platform.startswith("linux"):
        return "skipped"
    try:
        r = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
                           capture_output=True, text=True, timeout=600)
    except Exception as e:                                   # noqa: BLE001
        return f"failed: {e}"
    return "ok" if r.returncode == 0 else "failed: " + (r.stderr or r.stdout)[-500:]


_browser_status = ensure_browser()
if _browser_status.startswith("failed"):
    st.warning("The automation browser could not be installed on this server. " + _browser_status)


# ----------------------------------------------------------------- helpers
def get_secret(name, default=""):
    """st.secrets first (Streamlit hosting), then environment / .env."""
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name, default)


def _now():
    return datetime.now().isoformat(timespec="seconds")


def read_meta(run_dir):
    try:
        return json.loads((run_dir / "meta.json").read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_meta(run_dir, meta):
    tmp = run_dir / "meta.tmp"
    tmp.write_text(json.dumps(meta), encoding="utf-8")
    os.replace(tmp, run_dir / "meta.json")


def run_alive(meta):
    """True while the background runner process exists."""
    pid = meta.get("pid")
    if pid:
        return psutil.pid_exists(pid)
    # the runner has not written its pid yet (it has just been started): allow 30 s
    try:
        return (datetime.now() - datetime.fromisoformat(meta["started"])).total_seconds() < 30
    except Exception:
        return False


def get_active_run():
    """Name of the run in progress (or None). One run at a time: two logins with the same ERP user would collide."""
    dirs = sorted((d for d in RUNS.iterdir() if d.is_dir()), reverse=True)[:20]
    for d in dirs:
        meta = read_meta(d)
        if meta.get("status") != "running":
            continue
        if run_alive(meta):
            return d.name
        meta.update(status="interrupted", finished=_now())     # runner died (e.g. server restart)
        write_meta(d, meta)
    return None


def stop_run(name):
    """Stop a run: mark it 'stopped' first, then kill the runner, the script and any browser it opened."""
    d = RUNS / name
    meta = read_meta(d)
    started = datetime.fromisoformat(meta["started"])
    meta.update(status="stopped", finished=_now(), seconds=round((datetime.now() - started).total_seconds()))
    write_meta(d, meta)
    pid = meta.get("pid")
    if pid:
        try:
            parent = psutil.Process(pid)
            for child in parent.children(recursive=True):
                try:
                    child.kill()
                except psutil.Error:
                    pass
            parent.kill()
        except psutil.Error:
            pass


def parse_rows(text):
    """Lines 'CustomsNo  InvoiceNo' (tab/space separated; a date between them is ignored) -> {customs_no: [invoice, ...]}."""
    groups = {}
    for line in text.splitlines():
        parts = line.split()
        if len(parts) < 2:
            continue
        cust, inv = parts[0], parts[-1]
        if inv not in groups.setdefault(cust, []):
            groups[cust].append(inv)
    return groups


def default_header_fields():
    """Read HEADER_FIELDS from rpa_erp.py without importing it."""
    tree = ast.parse(SCRIPT_HEADER.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") == "HEADER_FIELDS":
            return ast.literal_eval(node.value)
    return []


def launch(kind, script, args, files, env_extra, total_steps):
    """Create the run folder and start runner.py detached from Streamlit. Returns the run folder."""
    if get_active_run():
        st.error("Another automation is already running. Wait for it to finish or stop it first.")
        return None
    run_dir = RUNS / f"{datetime.now():%Y%m%d_%H%M%S}_{kind}"
    run_dir.mkdir()
    for name, content in files.items():
        if isinstance(content, bytes):
            (run_dir / name).write_bytes(content)
        else:
            (run_dir / name).write_text(content, encoding="utf-8")
    write_meta(run_dir, {"kind": kind, "status": "running", "started": _now(),
                         "total_steps": total_steps, "pid": None})

    env = os.environ.copy()
    env.update(env_extra)                      # credentials travel in the environment only, never on disk
    env["RPA_RESULT_JSON"] = str(run_dir / "result.json")
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    detach = {}
    if os.name == "nt":
        detach["creationflags"] = 0x00000008 | 0x00000200      # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    else:
        detach["start_new_session"] = True
    subprocess.Popen(
        [sys.executable, str(BASE / "runner.py"), str(run_dir), str(script), *args],
        cwd=run_dir, env=env, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        **detach,
    )
    return run_dir


KIND_LABEL = {"customs": "Update Customs No", "header": "Invoice header"}
STATUS_LABEL = {"running": "Running", "finished": "Finished", "error": "Error",
                "stopped": "Stopped", "interrupted": "Interrupted"}


def run_status(meta):
    """Status of a run; runs saved by an older version only have an exit code."""
    if meta.get("status"):
        return meta["status"]
    return "finished" if meta.get("exit_code", 0) == 0 else "error"


def render_active_run(name):
    """Live panel for the run in progress. Runs as a fragment: refreshes itself every 2 s without
    re-running the page, and the run itself is a separate process, so nothing here can interrupt it."""
    d = RUNS / name
    meta = read_meta(d)
    if meta.get("status") != "running" or not run_alive(meta):
        st.rerun()                                              # finished: redraw the whole page
    log = d / "log.txt"
    lines = log.read_text(encoding="utf-8", errors="replace").splitlines() if log.exists() else []
    steps = sum(1 for ln in lines if re.match(r"^Step 3\.\d+:", ln))
    total = max(int(meta.get("total_steps") or 1), 1)
    started = datetime.fromisoformat(meta["started"])
    elapsed = int((datetime.now() - started).total_seconds())
    with st.container(border=True):
        st.markdown(f"**Run in progress: {KIND_LABEL.get(meta.get('kind'), '')}**  \n"
                    f"Started {started:%H:%M:%S}, running for {elapsed // 60} min {elapsed % 60:02d} s")
        st.progress(min(steps / total, 1.0), text=f"{min(steps, total)} of {total} steps")
        st.code("\n".join(lines[-14:]) or "Starting...", language="text")
        c1, c2 = st.columns([1, 5], vertical_alignment="center")
        if c1.button("Stop run", key=f"stop_{name}"):
            stop_run(name)
            st.rerun()
        c2.caption("You can switch tabs or reload the page. The run keeps going and is saved in Run history.")


def render_finished_run(name):
    d = RUNS / name
    meta = read_meta(d)
    with st.container(border=True):
        head, btn = st.columns([6, 1], vertical_alignment="center")
        head.markdown(f"**Last run: {KIND_LABEL.get(meta.get('kind'), '')}**  \n"
                      f"Started {meta.get('started', '').replace('T', ' ')}")
        if btn.button("Dismiss", key=f"dismiss_{name}"):
            st.session_state.pop("finished_run", None)
            st.rerun()
        (show_customs_result if meta.get("kind") == "customs" else show_header_result)(d, "top")


def show_screenshots(run_dir, first=()):
    pngs = sorted(run_dir.glob("*.png"))
    pngs.sort(key=lambda p: (p.name not in first, p.name))
    if not pngs:
        return
    with st.expander(f"Screenshots ({len(pngs)})", expanded=any(p.name in first for p in pngs)):
        cols = st.columns(2)
        for i, p in enumerate(pngs):
            cols[i % 2].image(str(p), caption=p.name, use_container_width=True)


def show_log(run_dir, tag="x"):
    log = run_dir / "log.txt"
    if log.exists():
        with st.expander("Full log"):
            st.code(log.read_text(encoding="utf-8"), language="text")
        st.download_button("Download log", log.read_bytes(), file_name=f"{run_dir.name}.log", key=f"dl_{tag}_{run_dir.name}")


def show_customs_result(run_dir, tag="x"):
    meta = read_meta(run_dir)
    result = {}
    if (run_dir / "result.json").exists():
        result = json.loads((run_dir / "result.json").read_text())
    planned = parse_rows((run_dir / "data.txt").read_text(encoding="utf-8")) if (run_dir / "data.txt").exists() else {}
    res = result.get("results", {})

    rows, added, skipped, failed = [], 0, 0, 0
    updated_msgs = []
    for cust, invoices in planned.items():
        r = res.get(cust, {})
        for inv in invoices:
            if inv in r.get("added", []):
                status, note = "Added", ""
                added += 1
            elif inv in r.get("skipped", []):
                status, note = "Already in list", ""
                skipped += 1
            elif inv in r.get("failed", {}):
                status, note = "Failed", r["failed"][inv]
                failed += 1
            else:
                status, note = "Not run", ""
            rows.append({"Customs No": cust, "Invoice No": inv, "Status": status, "Note": note})
        if r.get("update_message"):
            updated_msgs.append(r["update_message"])

    status = run_status(meta)
    if status == "error" or result.get("error"):
        st.error(f"The run stopped with an error: {result.get('error') or 'see the log below'}")
    elif status == "stopped":
        st.warning("The run was stopped before it finished. Invoices marked 'Not run' were not processed.")
    elif status == "interrupted":
        st.warning("The run was interrupted (the computer or server stopped). Check the log below.")
    elif status == "running":
        st.info("This run is still in progress.")
    elif planned:
        st.success("The run finished.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Added", added)
    c2.metric("Already in list", skipped)
    c3.metric("Failed", failed)
    c4.metric("Customs No updated", "Yes" if updated_msgs else "No")

    for m in updated_msgs:
        st.info(m)
    if rows:
        df = pd.DataFrame(rows)
        palette = {"Added": "#E3F2EA", "Already in list": "#EEF1F4", "Failed": "#F9E3DF", "Not run": "#FFF4D6"}
        styler = df.style
        apply = getattr(styler, "map", None) or styler.applymap
        styled = apply(lambda v: f"background-color: {palette.get(v, '')}", subset=["Status"])
        st.dataframe(styled, use_container_width=True, hide_index=True)
    show_screenshots(run_dir, first=("error.png", "customs_updated.png"))
    show_log(run_dir, tag)


def show_header_result(run_dir, tag="x"):
    meta = read_meta(run_dir)
    log = (run_dir / "log.txt").read_text(encoding="utf-8", errors="replace") if (run_dir / "log.txt").exists() else ""
    status = run_status(meta)
    if status == "finished" and "Step 3 done" in log:
        st.success("Invoice header filled. Check the screenshot before saving in D365.")
    elif status == "running":
        st.info("This run is still in progress.")
    elif status in ("stopped", "interrupted"):
        st.warning(f"The run was {status} before it finished.")
    else:
        st.error("The run did not finish. See the last lines of the log and the screenshots.")
    show_screenshots(run_dir, first=("step3_filled.png", "error.png"))
    show_log(run_dir, tag)


def list_runs():
    runs = []
    for d in sorted(RUNS.iterdir(), reverse=True):
        if d.is_dir() and (d / "meta.json").exists():
            runs.append((d, read_meta(d)))
    return runs


# ------------------------------------------------------------- access gate
app_password = get_secret("APP_PASSWORD")
if app_password and not st.session_state.get("auth"):
    st.title("Export invoice automation")
    with st.form("gate"):
        pw = st.text_input("Dashboard password", type="password")
        if st.form_submit_button("Open dashboard", type="primary"):
            if pw == app_password:
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("That password is not correct.")
    st.stop()

# ----------------------------------------------------------------- sidebar
with st.sidebar:
    st.subheader("D365 connection")
    saved_user = get_secret("ERP_USER")
    saved_pw = get_secret("ERP_PASSWORD")
    erp_user = st.text_input("User", value=saved_user)
    erp_pw = st.text_input("Password", type="password",
                           placeholder="Using the saved password" if saved_pw else "Enter password")
    erp_pw = erp_pw or saved_pw
    st.caption("Passwords are passed to the run only and are never written to the run folder.")

    st.subheader("Run options")
    show_browser = st.toggle("Show the browser window", value=False,
                             help="Turn on only when the dashboard runs on your own PC and you want to watch the steps.")
    do_update = st.toggle("Click Update Customs No at the end", value=True,
                          help="Writes to production data. Turn off to fill and add invoices only.")

creds_ok = bool(erp_user and erp_pw)
run_env = {
    "ERP_USER": erp_user, "ERP_PASSWORD": erp_pw,
    "HEADLESS": "0" if show_browser else "1",
}
form_url = get_secret("ERP_FORM_URL")
if form_url:
    run_env["ERP_FORM_URL"] = form_url

# -------------------------------------------------------------------- page
st.title("Export invoice automation")
st.markdown('<div class="subtle">Runs the D365 Production export invoice steps for you and keeps a record of every run.</div>',
            unsafe_allow_html=True)
if not creds_ok:
    st.warning("Enter the D365 user and password in the sidebar before running.")

# Live panel (while a run is in progress) or the result of the last run
active = get_active_run()
busy = active is not None
if active:
    st.session_state["watching"] = active
    st.fragment(run_every=2)(render_active_run)(active)
else:
    just_done = st.session_state.pop("watching", None)
    if just_done:
        st.session_state["finished_run"] = just_done
    finished = st.session_state.get("finished_run")
    if finished and (RUNS / finished).exists():
        render_finished_run(finished)

tab_customs, tab_header, tab_history = st.tabs(["Update Customs No", "Invoice header", "Run history"])

# ---------------------------------------------------------- Update Customs No
with tab_customs:
    head_l, head_r = st.columns([4, 1], vertical_alignment="bottom")
    head_l.markdown("Upload the filled Excel template. Each row holds a Customs No and an Invoice No.")
    head_r.download_button("Download template", build_template(), file_name="invoice_template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    source = st.radio("Data source", ["Upload filled template", "Convert raw data", "Paste rows"],
                      horizontal=True, label_visibility="collapsed")

    groups, rows_df, errors, warnings, infos = {}, None, [], [], []
    upload_name, upload_bytes, raw_copy = None, None, {}
    if source == "Upload filled template":
        up = st.file_uploader("Filled template", type=["xlsx", "xlsm", "csv"], label_visibility="collapsed",
                              key="tpl_up")
        if up is not None:
            upload_name, upload_bytes = up.name, up.getvalue()
            parsed = read_invoice_file(upload_bytes, up.name)
            groups, rows_df, errors, warnings = parsed["groups"], parsed["rows"], parsed["errors"], parsed["warnings"]
        else:
            st.info("No file yet. Download the template, fill the 'Invoices' sheet, then upload it here.")

    elif source == "Convert raw data":
        st.caption("Upload your raw customs list. CUSTOM becomes Customs No and INV# becomes Invoice No; "
                   "other columns, title rows and blank separator rows are ignored.")
        sc_only = st.checkbox("For invoices starting with SC, keep only the base invoice (skip -1, -2, -3 ...)",
                              value=True, help="SC260389-1 to SC260389-22 become one row: SC260389. "
                                              "Invoices that do not start with SC, such as C265386-1, are kept as they are.")
        up = st.file_uploader("Raw customs list", type=["xlsx", "xlsm", "csv"], label_visibility="collapsed",
                              key="raw_up")
        if up is not None:
            conv = convert_raw(up.getvalue(), up.name, sc_base_only=sc_only)
            infos = conv["info"]
            errors += conv["errors"]
            warnings += conv["warnings"]
            if conv["template"]:
                stem = Path(up.name).stem
                upload_name, upload_bytes = f"converted_{stem}.xlsx", conv["template"]
                raw_copy = {f"raw_{up.name}": up.getvalue()}
                parsed = read_invoice_file(conv["template"], upload_name)
                groups, rows_df = parsed["groups"], parsed["rows"]
                errors += parsed["errors"]
                warnings += parsed["warnings"]
                st.download_button("Download converted template", conv["template"],
                                   file_name=f"invoice_template_{stem}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No file yet. Upload the raw list and it is converted to the template format here.")

    else:
        raw = st.text_area("Paste rows from Excel", height=160, placeholder="E274120\tC265825",
                           help="One row per invoice: Customs No, then Invoice No (tab or space between).")
        groups = parse_rows(raw)
        rows_df = pd.DataFrame([{"Customs No": c, "Invoice No": i} for c, v in groups.items() for i in v])
        if raw.strip() and not groups:
            errors.append("No rows recognised. Use: Customs No, Invoice No.")

    total_invoices = sum(len(v) for v in groups.values())
    for e in errors:
        st.error(e)
    for w in warnings:
        st.warning(w)
    for note in (infos if source == "Convert raw data" else []):
        st.info(note)

    if groups and not errors:
        st.success(f"Read {total_invoices} invoice(s) for {len(groups)} Customs No.")
        left, right = st.columns([2, 3], gap="large")
        with left:
            st.markdown("**Customs No groups**")
            st.dataframe(pd.DataFrame([{"Customs No": c, "Invoices": len(v)} for c, v in groups.items()]),
                         hide_index=True, use_container_width=True)
            st.caption("Customs No is filled once per group, then each invoice is added in turn."
                       + (" Update Customs No is clicked after each group." if do_update
                          else " Update Customs No is not clicked."))
        with right:
            st.markdown("**Rows to process**")
            st.dataframe(rows_df, hide_index=True, use_container_width=True, height=min(35 * (len(rows_df) + 1) + 3, 320))

    can_run = bool(groups) and not errors and creds_ok and not busy
    if busy:
        st.caption("A run is in progress. Follow it in the panel at the top, or stop it there.")
    if st.button("Run Customs No update", type="primary", disabled=not can_run, key="run_customs"):
        data_txt = "\n".join(f"{c}\t-\t{inv}" for c, invs in groups.items() for inv in invs)
        files = {"data.txt": data_txt}
        if upload_bytes:
            files["upload_" + upload_name] = upload_bytes        # keep the uploaded / converted file with the run
        files.update(raw_copy)                                   # and the original raw file, if one was converted
        args = ["--file", "data.txt"] + ([] if do_update else ["--no-update"])
        if launch("customs", SCRIPT_CUSTOMS, args, files, run_env, total_invoices):
            st.rerun()

# ------------------------------------------------------------ Invoice header
with tab_header:
    st.markdown("Fills the Invoice Header tab for one export invoice. Edit the values below for this run.")
    invoice_no = st.text_input("Export invoice no.", placeholder="SC260309")

    fields = default_header_fields()
    df_fields = pd.DataFrame(
        [{"Field": f[0], "Value": f[1], "Type": f[2], "Dropdown row": ", ".join(f[3]) if len(f) > 3 else ""} for f in fields]
    )
    edited = st.data_editor(
        df_fields, hide_index=True, use_container_width=True, num_rows="dynamic",
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["text", "lookup"], required=True,
                                                     help="text: type and press Tab. lookup: type, then click the matching dropdown row."),
            "Dropdown row": st.column_config.TextColumn("Dropdown row", help="Optional. Texts to click in the dropdown, separated by commas."),
        },
    )
    st.caption("Dates are typed as d/m/yyyy, the way D365 shows them. The run does not click Save.")

    if st.button("Fill invoice header", type="primary", disabled=not (invoice_no.strip() and creds_ok) or busy, key="run_header"):
        payload = []
        for _, r in edited.dropna(subset=["Field"]).iterrows():
            picks = [p.strip() for p in str(r.get("Dropdown row") or "").split(",") if p.strip()]
            item = [str(r["Field"]), str(r["Value"]), str(r["Type"] or "text")]
            if picks:
                item.append(picks)
            payload.append(item)
        env = {**run_env, "HEADER_FIELDS_JSON": json.dumps(payload)}
        if launch("header", SCRIPT_HEADER, [invoice_no.strip()], {}, env, len(payload)):
            st.rerun()

# ------------------------------------------------------------------ History
with tab_history:
    runs = list_runs()
    if not runs:
        st.info("No runs yet. Start one from the other tabs and it will appear here.")
    else:
        table = pd.DataFrame([
            {"Started": m.get("started", "").replace("T", " "),
             "Process": KIND_LABEL.get(m.get("kind"), m.get("kind")),
             "Result": STATUS_LABEL.get(run_status(m), run_status(m)),
             "Seconds": m.get("seconds"),
             "Run": d.name}
            for d, m in runs
        ])
        st.dataframe(table.drop(columns="Run"), hide_index=True, use_container_width=True)
        pick = st.selectbox("Open a run", [d.name for d, _ in runs], key="hist_pick")
        d = RUNS / pick
        meta = read_meta(d)
        if run_status(meta) == "running":
            st.info("This run is still in progress. Follow it in the panel at the top of the page.")
        else:
            (show_customs_result if meta.get("kind") == "customs" else show_header_result)(d, "hist")
