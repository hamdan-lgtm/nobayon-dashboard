"""Nobayon Traders - Monthly Business Performance dashboard (entrypoint).

Run:  dashboard shortcut  |  run_dashboard.bat  |  streamlit run streamlit_app.py
"""

from __future__ import annotations

import os
import subprocess
import sys
import time

import pandas as pd
import streamlit as st

from utils.data_pipeline import (
    MONTH_NAMES,
    ROOT,
    GOOGLE_SHEET_ID,
    google_auth_mode,
    available_districts,
    available_periods,
    available_products,
    available_warehouses,
    build_context,
    check_summary,
    csv_bytes,
    dataset_path,
    db_signature,
    default_period,
    excel_snapshot,
    get_theme,
    inject_css,
    load_dataset,
    load_database,
    month_bounds,
    overall_status,
    sales_export_frame,
    set_theme,
    status_pill_html,
    workbook_age_minutes,
)

st.set_page_config(
    page_title="Nobayon Traders · Business Performance",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="auto",
)

# Auto-reload cadence (seconds) - change here, used by the toggle and the stale warning.
REFRESH_SECONDS = 3600 * 2          # 2 hours
STALE_MINUTES = REFRESH_SECONDS // 60

# --------------------------------------------------------------------------- #
# Theme
# --------------------------------------------------------------------------- #

if "theme_mode" not in st.session_state:
    requested = str(st.query_params.get("theme", "")).lower()
    st.session_state["theme_mode"] = "light" if requested.startswith("light") else "brand"

active_theme = set_theme(st.session_state["theme_mode"])
inject_css(active_theme)

# --------------------------------------------------------------------------- #
# Data
# --------------------------------------------------------------------------- #

DB_FILE = dataset_path()
MTIME = db_signature(DB_FILE)

if not GOOGLE_SHEET_ID and not DB_FILE.exists():
    st.error(
        f"`{DB_FILE.name}` was not found.\n\n"
        "Expected in `data/`, next to the app, or in `..\\Business_dashboard`. "
        "For Streamlit Cloud, commit the workbook to `data/` or set `GOOGLE_SHEET_ID` in the app secrets."
    )
    st.stop()

with st.spinner("Loading the Power Query extract…"):
    try:
        dataset = load_database(str(DB_FILE), MTIME)
    except Exception as exc:
        st.error(f"Could not read `{DB_FILE.name}`.\n\n`{type(exc).__name__}: {exc}`")
        st.info("Close the workbook in Excel (or wait for Power Query to finish) and reload.")
        st.stop()

if dataset.get("sales", pd.DataFrame()).empty:
    st.error("`SO_DO_Database` returned no rows — the workbook may still be refreshing.")
    st.stop()

# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #

st.sidebar.markdown("### Nobayon Traders")
st.sidebar.caption("Monthly Business Performance")

st.sidebar.radio(
    "Appearance",
    options=["brand", "light"],
    format_func=lambda v: "🌿 Brand" if v == "brand" else "☀️ Light",
    horizontal=True,
    key="theme_mode",
)
st.query_params["theme"] = st.session_state["theme_mode"]

st.sidebar.divider()

# --- refresh: this button drives Power Query and updates every date/value --- #
if st.sidebar.button(
    "🔄 Refresh data",
    width="stretch",
    type="primary",
    help="Runs Data → Refresh All in Excel (Power Query), saves the workbook, then re-reads it so every date and value updates. "
         "On Google Sheets it simply re-reads the sheet.",
):
    if GOOGLE_SHEET_ID:
        st.cache_data.clear()
        st.session_state["refresh_result"] = {
            "status": 0, "output": "Re-read the Google Sheets data source.", "refresh": "refreshed from Google Sheets",
            "seconds": 0.0, "at": pd.Timestamp.now(),
        }
        st.rerun()

    started = time.time()
    with st.spinner("Refreshing Power Query in Excel… usually 1–3 minutes."):
        try:
            completed = subprocess.run(
                [sys.executable, str(ROOT / "refresh_pipeline.py"), "--no-git", "--workbook", str(DB_FILE)],
                cwd=str(ROOT), capture_output=True, text=True, timeout=1800,
            )
            output = (completed.stdout or "") + (completed.stderr or "")
            status = completed.returncode
        except subprocess.TimeoutExpired:
            output, status = "Refresh timed out after 30 minutes.", 1
        except Exception as exc:
            output, status = f"{type(exc).__name__}: {exc}", 1

    refresh_line = next(
        (line.split("REFRESH:", 1)[1].strip() for line in output.splitlines() if line.startswith("REFRESH:")), ""
    )
    st.session_state["refresh_result"] = {
        "status": status, "output": output[-6000:], "refresh": refresh_line,
        "seconds": time.time() - started, "at": pd.Timestamp.now(),
    }
    st.cache_data.clear()
    st.rerun()

result = st.session_state.get("refresh_result")
if result:
    did_refresh = result.get("refresh", "").startswith("refreshed")
    icon = "✅" if result["status"] == 0 and did_refresh else ("⚠️" if result["status"] == 0 else "❌")
    with st.sidebar.expander(f"{icon} Last refresh · {result['seconds']:.0f}s", expanded=not did_refresh):
        if result.get("refresh"):
            st.caption(result["refresh"])
        if result["status"] != 0:
            st.error("The refresh script failed — see the log below.")
        elif not did_refresh:
            st.warning("Power Query did not run (Excel/pywin32 unavailable). Data reloaded from disk only.")
        st.code(result["output"] or "(no output)", language="text")

auto_reload = st.sidebar.toggle(
    "⏱ Auto-reload (2 h)",
    value=False,
    help=f"Re-reads the workbook every {REFRESH_SECONDS // 3600} hours when its timestamp changes. "
         "It never runs Power Query - use Refresh data for that.",
)
if auto_reload:
    if "_db_signature" not in st.session_state:
        st.session_state["_db_signature"] = MTIME

    @st.fragment(run_every=REFRESH_SECONDS)
    def _watch() -> None:
        if db_signature(DB_FILE) != st.session_state.get("_db_signature"):
            st.cache_data.clear()
            st.rerun(scope="app")

    _watch()

age = workbook_age_minutes(MTIME)
loaded = pd.Timestamp.fromtimestamp(MTIME).strftime("%d %b %Y %H:%M") if MTIME else "unknown"
age_text = (
    f"{age:.0f} min ago"
    if age is not None and age < STALE_MINUTES
    else (f"{age / 60:.1f} h ago" if age else "")
)
if GOOGLE_SHEET_ID:
    st.sidebar.caption("Source: Google Sheets · reloads every 15 min, or press Refresh data")
else:
    st.sidebar.caption(f"Workbook updated {loaded} · {age_text}")

# --- period --- #
st.sidebar.divider()
periods = available_periods(dataset)
years = sorted(periods.keys(), reverse=True) or [pd.Timestamp.today().year]
default_year, default_month = default_period(dataset)

mode = st.sidebar.radio("Period", ["Month", "Custom range"], horizontal=True, key="period_mode")
if mode == "Month":
    year = st.sidebar.selectbox("Year", years, index=years.index(default_year) if default_year in years else 0, key="filter_year")
    months = periods.get(year) or list(range(1, 13))
    month = st.sidebar.selectbox(
        "Month", months, format_func=lambda m: MONTH_NAMES[int(m) - 1],
        index=months.index(default_month) if default_month in months else 0, key=f"filter_month_{year}",
    )
    start_dt, end_dt = month_bounds(year, month)
else:
    s0, e0 = month_bounds(default_year, default_month)
    picked = st.sidebar.date_input("Date range", value=(s0.date(), e0.date()), key="filter_dates")
    if isinstance(picked, (tuple, list)):
        start_dt, end_dt = pd.Timestamp(picked[0]), pd.Timestamp(picked[-1])
    else:
        start_dt = end_dt = pd.Timestamp(picked)
    start_dt, end_dt = start_dt.normalize(), end_dt.normalize()

# --- location & product --- #
districts = available_districts(dataset)
district_choice = st.sidebar.selectbox("District", ["All Locations"] + districts, key="filter_district")
district = None if district_choice == "All Locations" else district_choice

warehouses = available_warehouses(dataset, district)
warehouse_choice = st.sidebar.selectbox(
    "Warehouse", ["All Warehouses"] + warehouses, key=f"filter_warehouse_{district or 'all'}",
    help="Applies to stock. Sales are attributable at district level.",
)
warehouse = None if warehouse_choice == "All Warehouses" else warehouse_choice

products = available_products(dataset)
product_choice = st.sidebar.selectbox("Product", ["All Products"] + products, key="filter_product")
product = None if product_choice == "All Products" else product_choice

ctx = build_context(start_dt, end_dt, district, warehouse, product)
st.session_state["dataset"] = dataset
st.session_state["ctx"] = ctx

# --- downloads --- #
@st.cache_data(show_spinner=False, max_entries=6)
def _excel_export(path_str: str, mtime: float, start: str, end: str, district: str | None,
                  warehouse: str | None, product: str | None) -> bytes:
    from utils.data_pipeline import build_view

    export_ds = load_dataset(path_str, mtime)
    export_ctx = build_context(start, end, district, warehouse, product)
    return excel_snapshot(export_ds, export_ctx, build_view(export_ds, export_ctx))


with st.sidebar.expander("⬇ Download", expanded=False):
    st.caption(f"{ctx.month_label} · {ctx.location_label}")
    tag = f"{ctx.start:%Y%m}_{ctx.end:%d}"
    try:
        st.download_button(
            "⬇ Excel — full report",
            data=_excel_export(str(DB_FILE), MTIME, str(ctx.start.date()), str(ctx.end.date()), district, warehouse, product),
            file_name=f"nobayon_report_{tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    except Exception as exc:
        st.caption(f"Excel export unavailable: {type(exc).__name__}")

    from utils.data_pipeline import build_view

    st.download_button(
        "⬇ Sales detail — CSV",
        data=csv_bytes(sales_export_frame(build_view(dataset, ctx)["sales"])),
        file_name=f"nobayon_sales_{tag}.csv",
        mime="text/csv",
        width="stretch",
    )

# --- status --- #
checks = dataset.get("checks", [])
summary = check_summary(checks)
st.sidebar.divider()
st.sidebar.markdown(status_pill_html(overall_status(checks),
                                     f"{summary.get('PASS', 0)} pass · {summary.get('WARN', 0)} alert · {summary.get('FAIL', 0)} fail"),
                    unsafe_allow_html=True)
st.sidebar.caption(f"Source: {'Google Sheets · ' + GOOGLE_SHEET_ID if GOOGLE_SHEET_ID else DB_FILE.name}")
if GOOGLE_SHEET_ID:
    if google_auth_mode() == "service-account":
        st.sidebar.caption("Access: private · service account")
    else:
        st.sidebar.caption("Access: public link — add SHEETS_CREDENTIALS to make it private")
if age is not None and age > STALE_MINUTES:
    st.sidebar.warning(f"Data is {age / 60:.1f} h old, press Refresh data.", icon="⏳")

# --------------------------------------------------------------------------- #
# Navigation
# --------------------------------------------------------------------------- #

st.navigation(
    [
        st.Page("pages/1_overview.py", title="Performance Overview", icon="📊", default=True),
        st.Page("pages/2_analysis.py", title="Analysis & Insights", icon="🔍"),
        st.Page("pages/4_targets.py", title="Targets", icon="🎯"),
        st.Page("pages/3_detail.py", title="Detail & Audit", icon="📋"),
    ]
).run()
