"""Targets - product-wise target vs achieved (moved off the Overview page)."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.data_pipeline import (
    build_view,
    chart_target_achievement,
    check_summary,
    download_csv_button,
    fmt_cr,
    fmt_pct,
    get_state,
    inject_css,
    overall_status,
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
performance = product_performance(view["sales"], view["cogs"], dataset.get("product_targets"), ctx.month_key)
targets = target_achievement_table(performance, dataset.get("product_targets"), ctx.month_key)

render_header(dataset, ctx)
spacer(5)

# --------------------------------------------------------------------------- #
# Target summary
# --------------------------------------------------------------------------- #

summary = check_summary(dataset.get("checks", []))
left, middle, right, far = st.columns(4, gap="small")
left.metric("Sales target (Cr)", fmt_cr(view["sales_target"]))
middle.metric("Achieved (Cr)", fmt_cr(view["sales_cr"]), fmt_pct(view["sales_ach"]))
right.metric("Still to sell (Cr)", fmt_cr(view["asking_month"]))
far.markdown(
    f'<div class="metric-note">Asking per day</div>'
    f'<div class="metric-note-value">{fmt_cr(view["asking_day"])}</div>'
    f'<div class="metric-note">{ctx.days_remaining} day(s) left</div>',
    unsafe_allow_html=True,
)
spacer(5)

# --------------------------------------------------------------------------- #
# Chart + grouped table
# --------------------------------------------------------------------------- #

chart_col, table_col = st.columns([1.15, 1], gap="small")

with chart_col:
    with st.container(border=True, key="card-target-chart"):
        section("Target vs Achieved", "Product targets · BDT Crore · green = target met")
        st.plotly_chart(
            chart_target_achievement(performance),
            theme=None, config={"displayModeBar": False}, key="ch_target",
        )

with table_col:
    with st.container(border=True, key="card-target-table"):
        section("Sales Target vs Achieve", "Product groups · BDT Crore")
        if targets.empty:
            st.info("No product targets for this month. Add rows to `product_targets.csv`.")
        else:
            st.dataframe(
                targets,
                hide_index=True,
                width="stretch",
                height=table_height(targets, max_height=340),
                column_config={
                    "Product": st.column_config.TextColumn("Product", width="medium"),
                    "Target (Cr)": st.column_config.NumberColumn("Target (Cr)", format="%.2f"),
                    "Achieved (Cr)": st.column_config.NumberColumn("Achieved (Cr)", format="%.2f"),
                    "Ach %": st.column_config.NumberColumn("Ach %", format="%.1f"),
                },
            )
            download_csv_button(targets, f"targets_{ctx.month_key}.csv", "dl_targets")

spacer(5)

# --------------------------------------------------------------------------- #
# Product-level detail
# --------------------------------------------------------------------------- #

with st.container(border=True, key="card-target-products"):
    section("Product Detail", "Sales, cost and margin per product for the selected month")
    st.dataframe(
        performance,
        hide_index=True,
        width="stretch",
        height=table_height(performance, max_height=420),
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
    download_csv_button(performance, f"product_performance_{ctx.month_key}.csv", "dl_perf")

spacer(4)
st.caption(
    f"{ctx.month_label} · targets from targets.csv / product_targets.csv · "
    f"validation {overall_status(dataset.get('checks', []))} "
    f"({summary.get('PASS', 0)} pass / {summary.get('WARN', 0)} alert / {summary.get('FAIL', 0)} fail)"
)
