"""Detail tables, exports and the data-pipeline audit."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.data_pipeline import (
    aging_detail,
    build_view,
    check_summary,
    collections_by_party,
    coverage_table,
    customer_performance,
    download_csv_button,
    fmt_cr,
    fmt_mt,
    get_state,
    inject_css,
    overall_status,
    product_performance,
    render_header,
    sales_export_frame,
    section,
    spacer,
    stock_detail,
    table_height,
)

dataset, ctx = get_state()
inject_css()
view = build_view(dataset, ctx)

render_header(dataset, ctx)
spacer(5)

TAB_SALES, TAB_COLL, TAB_STOCK, TAB_AGING, TAB_AUDIT = st.tabs(
    ["🧾 Sales", "💰 Collections", "🏬 Stock", "⏳ Aging", "✅ Audit"]
)

period_tag = f"{ctx.start:%Y%m}_{ctx.end:%d}"

with TAB_SALES:
    sales = sales_export_frame(view["sales"])
    left, middle, right = st.columns(3, gap="small")
    left.metric("Rows", f"{len(sales):,}")
    middle.metric("Sales (Cr)", fmt_cr(view["sales_cr"]))
    right.metric("GP (Cr)", fmt_cr(view["gp_cr"]))
    spacer(4)
    st.dataframe(
        sales,
        hide_index=True,
        width="stretch",
        height=table_height(sales, max_height=520),
        column_config={
            "Date": st.column_config.DateColumn("Date", format="DD-MMM-YYYY"),
            "SO Number": st.column_config.TextColumn("SO Number", width="small"),
            "Customer": st.column_config.TextColumn("Customer", width="medium"),
            "Product": st.column_config.TextColumn("Product", width="medium"),
            "Qty (MT)": st.column_config.NumberColumn("Qty (MT)", format="%.2f"),
            "Sales (BDT)": st.column_config.NumberColumn("Sales (BDT)", format="localized"),
            "Cost (BDT)": st.column_config.NumberColumn("Cost (BDT)", format="localized"),
            "GP (BDT)": st.column_config.NumberColumn("GP (BDT)", format="localized"),
        },
    )
    download_csv_button(sales, f"sales_detail_{period_tag}.csv", "dl_sales")

with TAB_COLL:
    collections = collections_by_party(view["collection"])
    left, middle, right = st.columns(3, gap="small")
    left.metric("Collection (Cr)", fmt_cr(view["collection_cr"]))
    middle.metric("Parties", f"{len(collections):,}")
    right.metric("Receipts", f"{int(view['collection']['Amount'].count()) if len(view['collection']) else 0:,}")
    spacer(4)
    st.dataframe(
        collections,
        hide_index=True,
        width="stretch",
        height=table_height(collections, max_height=520),
        column_config={
            "Party": st.column_config.TextColumn("Party", width="large"),
            "Group": st.column_config.TextColumn("Group", width="small"),
            "Receipts (Cr)": st.column_config.NumberColumn("Receipts (Cr)", format="%.2f"),
            "Receipts": st.column_config.NumberColumn("Receipts", format="%d"),
            "Main Bank": st.column_config.TextColumn("Main Bank", width="small"),
        },
    )
    download_csv_button(collections, f"collections_{period_tag}.csv", "dl_coll")

with TAB_STOCK:
    stock = stock_detail(view["stock"])
    left, middle, right = st.columns(3, gap="small")
    left.metric("Closing (MT)", fmt_mt(view["inventory_mt"]))
    middle.metric("Value (Cr)", fmt_cr(view["inventory_cr"]))
    right.metric("Snapshot", f"{view['stock_as_of']:%d-%b-%Y}" if pd.notna(view["stock_as_of"]) else "—")
    spacer(4)
    st.dataframe(
        stock,
        hide_index=True,
        width="stretch",
        height=table_height(stock, max_height=520),
        column_config={
            "District": st.column_config.TextColumn("District", width="small"),
            "Warehouse": st.column_config.TextColumn("Warehouse", width="medium"),
            "Product": st.column_config.TextColumn("Product", width="medium"),
            "Closing (MT)": st.column_config.NumberColumn("Closing (MT)", format="%.2f"),
            "Rate (BDT/MT)": st.column_config.NumberColumn("Rate (BDT/MT)", format="localized"),
            "Value (Cr)": st.column_config.NumberColumn("Value (Cr)", format="%.2f"),
        },
    )
    st.caption("Valued at the last COGS rate per product (fallback: last sale price, then family median).")
    download_csv_button(stock, f"stock_{period_tag}.csv", "dl_stock")

with TAB_AGING:
    aging = aging_detail(view["aging"])
    left, middle, right = st.columns(3, gap="small")
    left.metric("Receivable (Cr)", fmt_cr(view["receivable_cr"]))
    middle.metric("Overdue (Cr)", fmt_cr(view["overdue_cr"]))
    right.metric("Snapshot", f"{view['aging_as_of']:%d-%b-%Y}" if pd.notna(view["aging_as_of"]) else "—")
    spacer(4)
    st.dataframe(
        aging,
        hide_index=True,
        width="stretch",
        height=table_height(aging, max_height=460),
        column_config={
            "BP Code": st.column_config.TextColumn("BP Code", width="small"),
            "BP Name": st.column_config.TextColumn("BP Name", width="large"),
            **{col: st.column_config.NumberColumn(col, format="%.3f") for col in
               ["0-30", "31-60", "61-90", "91-180", "181-365", "365+", "Total (Cr)", "Overdue (Cr)"]},
        },
    )
    st.caption("Overdue = total minus the 0–30 day bucket (the source sheet has no Overdue column).")
    download_csv_button(aging, f"aging_{period_tag}.csv", "dl_aging")

with TAB_AUDIT:
    checks = dataset.get("checks", [])
    summary = check_summary(checks)
    left, middle, right = st.columns(3, gap="small")
    left.metric("Overall status", overall_status(checks))
    middle.metric("Checks passing", f"{summary.get('PASS', 0)} / {len(checks)}")
    right.metric("Alerts / failures", f"{summary.get('WARN', 0)} / {summary.get('FAIL', 0)}")
    spacer(4)

    checks_frame = pd.DataFrame(checks)[["Status", "Check", "Detail"]]
    styled = checks_frame.style.map(
        lambda v: ("color:#15803D;font-weight:700" if v == "PASS"
                   else ("color:#B45309;font-weight:700" if v == "WARN" else "color:#B91C1C;font-weight:700")),
        subset=["Status"],
    )
    st.dataframe(styled, hide_index=True, width="stretch", height=table_height(checks_frame, max_height=460))
    download_csv_button(checks_frame, "pipeline_audit.csv", "dl_audit")

    spacer(5)
    section("Sheet coverage")
    coverage = coverage_table(dataset)
    st.dataframe(
        coverage,
        hide_index=True,
        width="stretch",
        column_config={
            "Sheet": st.column_config.TextColumn("Sheet", width="medium"),
            "Rows": st.column_config.NumberColumn("Rows", format="%d"),
            "First Date": st.column_config.DateColumn("First Date", format="DD-MMM-YYYY"),
            "Latest Date": st.column_config.DateColumn("Latest Date", format="DD-MMM-YYYY"),
            "Null Dates": st.column_config.NumberColumn("Null Dates", format="%d"),
        },
    )
    st.caption(f"Workbook: {dataset.get('path', '')} · sheets: {', '.join(dataset.get('sheets', []))}")
