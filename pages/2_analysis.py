"""Analysis & Insights - what the numbers say, targets, mix and concentration."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.data_pipeline import (
    build_insights,
    build_view,
    customer_performance,
    download_csv_button,
    fmt_cr,
    fmt_pct,
    get_state,
    inject_css,
    insight_card,
    monthly_series,
    product_performance,
    render_header,
    section,
    spacer,
    table_height,
    target_achievement_table,
)

dataset, ctx = get_state()
inject_css()
view = build_view(dataset, ctx)
series = monthly_series(dataset, ctx, 12)
performance = product_performance(view["sales"], view["cogs"], dataset.get("product_targets"), ctx.month_key)
customers = customer_performance(view["sales"])
insights = build_insights(dataset, ctx, view, series, performance, customers)
targets = target_achievement_table(performance, dataset.get("product_targets"), ctx.month_key)

render_header(dataset, ctx)
spacer(5)

# --------------------------------------------------------------------------- #
# Insights
# --------------------------------------------------------------------------- #

section("What the numbers say", f"{ctx.month_label} · auto-generated from the current selection")
st.markdown(
    '<div class="insight-grid">' + "".join(insight_card(item) for item in insights) + "</div>",
    unsafe_allow_html=True,
)

spacer(6)

# --------------------------------------------------------------------------- #
# Targets and monthly performance
# --------------------------------------------------------------------------- #

left, right = st.columns([1, 1.1], gap="small")

with left:
    with st.container(border=True, key="card-targets"):
        section("Sales Target vs Achieved", "Product targets · BDT Crore")
        if targets.empty:
            st.info("No product targets for this month. Add rows to `product_targets.csv`.")
        else:
            st.dataframe(
                targets,
                hide_index=True,
                width="stretch",
                height=table_height(targets, max_height=320),
                column_config={
                    "Product": st.column_config.TextColumn("Product", width="medium"),
                    "Target (Cr)": st.column_config.NumberColumn("Target (Cr)", format="%.2f"),
                    "Achieved (Cr)": st.column_config.NumberColumn("Achieved (Cr)", format="%.2f"),
                    "Ach %": st.column_config.NumberColumn("Ach %", format="%.1f"),
                },
            )
            download_csv_button(targets, f"targets_{ctx.month_key}.csv", "dl_targets")

with right:
    with st.container(border=True, key="card-monthly"):
        section("Month-on-Month Performance", "Last 12 months · BDT Crore / MT")
        st.dataframe(
            series[["Month", "Sales (Cr)", "Collection (Cr)", "GP (Cr)", "GP %", "Qty (MT)"]],
            hide_index=True,
            width="stretch",
            height=table_height(series, max_height=320),
            column_config={
                "Month": st.column_config.TextColumn("Month", width="small"),
                "Sales (Cr)": st.column_config.NumberColumn("Sales (Cr)", format="%.2f"),
                "Collection (Cr)": st.column_config.NumberColumn("Collection (Cr)", format="%.2f"),
                "GP (Cr)": st.column_config.NumberColumn("GP (Cr)", format="%.2f"),
                "GP %": st.column_config.NumberColumn("GP %", format="%.1f"),
                "Qty (MT)": st.column_config.NumberColumn("Qty (MT)", format="%.0f"),
            },
        )
        download_csv_button(series, "monthly_performance.csv", "dl_series")

spacer(5)

# --------------------------------------------------------------------------- #
# Mix: products and customers
# --------------------------------------------------------------------------- #

left, right = st.columns(2, gap="small")

with left:
    with st.container(border=True, key="card-products"):
        section("Product Contribution", "Sales, cost and margin by product")
        st.dataframe(
            performance,
            hide_index=True,
            width="stretch",
            height=table_height(performance, max_height=460),
            column_config={
                "Product": st.column_config.TextColumn("Product", width="medium"),
                "Orders": st.column_config.NumberColumn("Orders", format="%d"),
                "Qty (MT)": st.column_config.NumberColumn("Qty (MT)", format="%.1f"),
                "Sales (Cr)": st.column_config.NumberColumn("Sales (Cr)", format="%.2f"),
                "COGS (Cr)": st.column_config.NumberColumn("COGS (Cr)", format="%.2f"),
                "GP (Cr)": st.column_config.NumberColumn("GP (Cr)", format="%.2f"),
                "GP %": st.column_config.NumberColumn("GP %", format="%.1f"),
                "Target (Cr)": st.column_config.NumberColumn("Target (Cr)", format="%.2f"),
                "Ach %": st.column_config.NumberColumn("Ach %", format="%.1f"),
            },
        )
        download_csv_button(performance, f"product_mix_{ctx.month_key}.csv", "dl_mix")

with right:
    with st.container(border=True, key="card-customers"):
        section("Customer Concentration", "Sales, invoice GP and share of total")
        st.dataframe(
            customers,
            hide_index=True,
            width="stretch",
            height=table_height(customers, max_height=460),
            column_config={
                "Customer": st.column_config.TextColumn("Customer", width="medium"),
                "Group": st.column_config.TextColumn("Group", width="small"),
                "Orders": st.column_config.NumberColumn("Orders", format="%d"),
                "Qty (MT)": st.column_config.NumberColumn("Qty (MT)", format="%.1f"),
                "Sales (Cr)": st.column_config.NumberColumn("Sales (Cr)", format="%.2f"),
                "Invoice GP (Cr)": st.column_config.NumberColumn("Invoice GP (Cr)", format="%.2f"),
                "GP %": st.column_config.NumberColumn("GP %", format="%.1f"),
                "Share %": st.column_config.NumberColumn("Share %", format="%.1f"),
            },
        )
        download_csv_button(customers, f"customer_mix_{ctx.month_key}.csv", "dl_cust")

spacer(5)

# --------------------------------------------------------------------------- #
# Reading guide
# --------------------------------------------------------------------------- #

with st.expander("How these numbers are calculated"):
    st.markdown(
        f"""
- **Sales** = `SUM(Sales Value)` for {ctx.month_label} (1–{ctx.end.day}). Volume is the line
  quantity in MT; `Qnt (Ton)` is ignored because its divisor is inconsistent across months.
- **Gross profit** = sales − `COGS_Database.Amount` for the same month. When a district or
  warehouse is selected the invoice cost on the sales line is used instead (COGS has no district).
- **Collection** = `SUM(Collection Amount)` for the month. *Sales vs Coll* is sales ÷ collection —
  above 1.00x means you billed more than you collected.
- **Receivables** = latest `Aging_Database` snapshot on or before {ctx.end:%d-%b-%Y}. The source has
  no Overdue column, so overdue = everything older than 30 days (total − 0–30 bucket).
- **Inventory** = latest stock snapshot per warehouse × product, valued at the last COGS rate
  (falling back to the last sale price, then the family median).
- **Targets** live in `targets.csv` (month totals) and `product_targets.csv` (product groups, use
  `SBM|Soybean (Local)` to combine products). Edit those files and press **R** to re-read them.
- **Prior MTD** = the previous month, 1st to the same day.
        """
    )
