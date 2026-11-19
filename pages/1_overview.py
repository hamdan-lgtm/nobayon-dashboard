"""Performance Overview - the monthly executive dashboard."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.data_pipeline import (
    EMERALD,
    GOLD,
    GREEN,
    LIME,
    TEAL,
    build_view,
    chart_cumulative,
    chart_customer_profit,
    chart_inventory,
    chart_monthly_trend,
    chart_partner_split,
    chart_product_profit,
    chart_profit_trend,
    chart_receivables,
    check_summary,
    customer_performance,
    fmt_cr,
    fmt_mt,
    fmt_pct,
    get_state,
    inject_css,
    kpi_card,
    monthly_series,
    overall_status,
    product_performance,
    render_header,
    section,
    spacer,
)

dataset, ctx = get_state()
inject_css()
view = build_view(dataset, ctx)
series = monthly_series(dataset, ctx, 12)
performance = product_performance(view["sales"], view["cogs"], dataset.get("product_targets"), ctx.month_key)
customers = customer_performance(view["sales"])

render_header(dataset, ctx)
spacer(5)

# --------------------------------------------------------------------------- #
# KPI band
# --------------------------------------------------------------------------- #

ar = view["ar_split"]
sales_delta = view["sales_delta"]
collection_delta = view["collection_delta"]
gp_delta = view["gp_delta"]
inventory_delta = view["inventory_delta"]

cards = [
    kpi_card(
        "Sales Revenue", fmt_cr(view["sales_cr"]), "Cr",
        [
            ("Sales Target", fmt_cr(view["sales_target"])),
            ("Achieve %", fmt_pct(view["sales_ach"])),
            ("Prev MTD", fmt_cr(view["sales_prior_cr"])),
            ("Variance", f"{sales_delta:+.1f}%" if sales_delta is not None else "—"),
        ],
        LIME, "📈", f"{view['so_count']:,} SOs · {view['volume_mt']:,.0f} MT",
        delta=(f"{sales_delta:+.1f}%", sales_delta >= 0) if sales_delta is not None else None,
    ),
    kpi_card(
        "Collection", fmt_cr(view["collection_cr"]), "Cr",
        [
            ("Coll Target", fmt_cr(view["collection_target"])),
            ("Achieve %", fmt_pct(view["collection_ach"])),
            ("Sales vs Coll", f"{(view['collection_cr'] / view['sales_cr']):,.2f}x" if view["sales_cr"] else "—"),
            ("Prev MTD", fmt_cr(view["collection_prior_cr"])),
        ],
        EMERALD, "💰", "",
        delta=(f"{collection_delta:+.1f}%", collection_delta >= 0) if collection_delta is not None else None,
    ),
    kpi_card(
        "Gross Profit", fmt_cr(view["gp_cr"]), "Cr",
        [
            ("Profit %", fmt_pct(view["gp_pct"])),
            ("Prev MTD", fmt_cr(view["gp_prior_cr"])),
            ("Prev Profit %", fmt_pct(view["gp_prior_pct"])),
            ("Change", f"{view['gp_delta_pp']:+.2f} pp"),
        ],
        GOLD, "📊", view["gp_basis"],
        delta=(f"{gp_delta:+.1f}%", gp_delta >= 0) if gp_delta is not None else None,
    ),
    kpi_card(
        "Receivables", fmt_cr(view["receivable_cr"]), "Cr",
        [
            ("AR · AAFL", fmt_cr(ar.get("AAFL", 0))),
            ("AR · Customer", fmt_cr(ar.get("Customer", 0))),
            ("Overdue", f"{fmt_cr(view['overdue_cr'])} ({fmt_pct(view['overdue_pct'])})"),
            ("Change", f"{view['receivable_change']:+,.2f} Cr"),
        ],
        TEAL, "⏳",
        f"as at {view['aging_as_of']:%d-%b-%y}" if pd.notna(view["aging_as_of"]) else "",
        delta=(f"{view['receivable_change']:+,.2f} Cr", view["receivable_change"] <= 0),
    ),
    kpi_card(
        "Inventory", fmt_cr(view["inventory_cr"]), "Cr",
        [
            ("Closing", f"{fmt_mt(view['inventory_mt'])} MT"),
            ("SKUs", f"{view['sku_count']:,}"),
            ("Warehouses", f"{view['warehouse_count']:,}"),
            ("Vs prior", f"{inventory_delta:+.1f}%" if inventory_delta is not None else "—"),
        ],
        GREEN, "🏬", "valued at last cost/price",
        delta=(f"{inventory_delta:+.1f}%", inventory_delta <= 0) if inventory_delta is not None else None,
    ),
]
st.markdown('<div class="kpi-grid">' + "".join(cards) + "</div>", unsafe_allow_html=True)
spacer(6)

# --------------------------------------------------------------------------- #
# Symmetrical 2 x 4 card grid (equal widths, equal heights)
# --------------------------------------------------------------------------- #

row1_left, row1_right = st.columns(2, gap="small")
with row1_left:
    with st.container(border=True, key="card-trend"):
        section("Sales & Collection Trend", "BDT Crore · last 12 months")
        st.plotly_chart(chart_monthly_trend(series), theme=None, config={"displayModeBar": False}, key="ch_trend")
with row1_right:
    with st.container(border=True, key="card-profit"):
        section("Profit % Trend", "Gross margin by month · BDT Crore")
        st.plotly_chart(chart_profit_trend(series), theme=None, config={"displayModeBar": False}, key="ch_profit")

spacer(5)

row2_left, row2_right = st.columns(2, gap="small")
with row2_left:
    with st.container(border=True, key="card-partner"):
        section("In-House vs Customer", "Sales and collection by business partner · BDT Crore")
        st.plotly_chart(chart_partner_split(view["sales"], view["collection"]),
                        theme=None, config={"displayModeBar": False}, key="ch_partner")
with row2_right:
    with st.container(border=True, key="card-pace"):
        section("Sales Pace (MTD)", f"{ctx.month_label} vs {ctx.prior_label} · BDT Crore")
        st.plotly_chart(chart_cumulative(view["sales"], view["sales_prior"], ctx),
                        theme=None, config={"displayModeBar": False}, key="ch_pace")

spacer(5)

row3_left, row3_right = st.columns(2, gap="small")
with row3_left:
    with st.container(border=True, key="card-product"):
        section("Product Profitability", "Top 10 products by sales · BDT Crore")
        st.plotly_chart(chart_product_profit(performance), theme=None, config={"displayModeBar": False}, key="ch_product")
with row3_right:
    with st.container(border=True, key="card-customer"):
        section("Customer Profitability", "Top 10 customers by sales · BDT Crore")
        st.plotly_chart(chart_customer_profit(customers), theme=None, config={"displayModeBar": False}, key="ch_customer")

spacer(5)

row4_left, row4_right = st.columns(2, gap="small")
with row4_left:
    with st.container(border=True, key="card-aging"):
        section("Receivables Aging", "Top customers · due vs overdue · BDT Crore")
        st.plotly_chart(chart_receivables(view["aging"]), theme=None, config={"displayModeBar": False}, key="ch_aging")
with row4_right:
    with st.container(border=True, key="card-stock"):
        section("Inventory by Product", "Closing stock · Metric Tons")
        st.plotly_chart(chart_inventory(view["stock"]), theme=None, config={"displayModeBar": False}, key="ch_stock")

spacer(4)

# --------------------------------------------------------------------------- #
# Footer
# --------------------------------------------------------------------------- #

checks = dataset.get("checks", [])
summary = check_summary(checks)
st.caption(
    f"{ctx.month_label} · {ctx.period_label} · {ctx.location_label} · validation {overall_status(checks)} "
    f"({summary.get('PASS', 0)} pass / {summary.get('WARN', 0)} alert / {summary.get('FAIL', 0)} fail) · "
    "Sales/Collection/GP are month-to-date; targets from targets.csv"
)
