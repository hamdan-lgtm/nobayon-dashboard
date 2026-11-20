"""Nobayon Traders - data pipeline for the Monthly Business Performance dashboard.

Reads the Power Query workbook (``dashboard database.xlsx``), normalises the product
master, values closing stock, joins the monthly targets and exposes:
KPIs, tables, charts, insights, exports and the validation suite.

The module never writes to the source workbook.
"""

from __future__ import annotations

import os
import re
import textwrap
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------- #
# Paths and constants
# --------------------------------------------------------------------------- #

ROOT = Path(__file__).resolve().parents[1]
CRORE = 10_000_000

CORE_SHEETS = ("SO_DO_Database", "Collection", "COGS_Database", "Aging_Database", "Stock_Database")

BRAND_GREEN = "#0B5D3B"
BRAND_LIME = "#8CC63F"
BRAND_DEEP = "#06301F"

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# --------------------------------------------------------------------------- #
# Product master normalisation (identical rules across every sheet)
# --------------------------------------------------------------------------- #

PRODUCT_FAMILIES: dict[str, list[str]] = {
    "Maize (Local)": ["Maize (Spot Sale)", "Maize (Local)-(Old)"],
    "Maize (Import)": ["Maize(Import)"],
    "RSM": [
        "Rapeseed Meal/Extraction",
        "Rapseed Meal/Extraction",
        "Rapseed Meal/Extraction",
        "Rapeseed Meal/Extraction-(Old)",
        "Rapeseed Meal/Extraction(Virtual)",
        "Rapseed Meal",
        "Rapeseed Extraction",
    ],
    "DDGS": [
        "DISTILLERS DRIED GRAINS WITH SOLUBLES (DDGS)INDIA-Local",
        "DISTILLERS DRIED GRAINS WITH SOLUBLES (DDGS)- INDIA",
    ],
    "SBM": [
        "Soya Bean Meal (Brazil)",
        "Soybean Meal Hi-Pro Local",
        "Soyabean Extraction India",
        "Soyabeen Extraction (USA)",
        "SOYBEAN EXTRACTION (MEAL) Brazilian",
        "SOYBEAN EXTRACTION (MEAL) Brazilian-(Old)",
        "SOYBEAN EXTRACTION (MEAL) U.S.A",
        "SBM USA",
        "SBM-Brazil",
        "SBM Brazil",
        "SBM India",
        "SBM USA Hipro",
        "Soybeen Hi Pro Local",
    ],
    "Mustards": [
        "MUSTARD - RED (TRADING)",
        "MUSTARD - YELLOW (TRADING)",
        "Mustard Seed-Red (Local)",
        "Mustard Seed-Yellow (Local)",
        "Mustard Red Seed (Local)",
        "Mustard Yellow Seed (Local)",
        "Mustard Seed Red",
        "Mustard Seed Yellow",
    ],
    "Jute": ["Jute Tossa Average Chat", "Jute Tossa C-Grade", "Jute Tosha C Grade", "Jute Toshsha C Grade"],
    "Soybean (Local)": ["Soybean (Local)", "Soyabean (Buy & Sale)", "Soyabeen Local"],
    "DL-Methionine": ["DL Mithionine (Local)", "DL-Methionine, NHU", "DL Methionine NHU"],
    "Limestone": ["Limestone 2-3mm, Egypt", "Limestone 2-3mm Egypt", "LimeStone"],
    "DCP": ["DCP 18% Powder", "Dicalcium Phosphate (DCP)"],
    "Rice Bran Crude Oil": ["Rice Brun Crude Oil"],
    "Broken Rice": ["Broken Rice."],
    "Groundnut": ["Groundnut (local)", "Groundnut"],
    "Til (Sesame)": ["Sesame ( Til )", "Sesame (Til)", "Til"],
    "L-Lysine": ["L-Lysine HCL", "L- Lysine", "L Lysine", "L-Lysine"],
    "L-Threonine": ["L-Threonine", "L- Threonine"],
    "MCP": ["MONOCALCIUM PHOSPHATE 23%", "Monocalcium Phosphate 23%", "MCP"],
    "Choline Chloride": ["CHOLINE CHLORIDE 60%", "Choline Chloride"],
    "Palm Olein": ["Fortified Palm Olein (204 Liter-01 Drum)"],
    "Super Palm Olein": ["Fortified Super quality Palm Olein (204 Liter-01 Drum)"],
    "DORB": ["De-Oiled Rice Bran (DORB)", "DORB"],
    "Wheat Bran": ["Wheat Bran"],
    "Red Chili": ["Red Chili"],
}

_DASHES = re.compile(r"[\u2010-\u2015\u2212]")


def _key(text: Any) -> str:
    if text is None:
        return ""
    if isinstance(text, float) and np.isnan(text):
        return ""
    value = str(text).replace("\xa0", " ").replace("\u200b", "")
    value = _DASHES.sub("-", value)
    return re.sub(r"\s+", " ", value).strip().lower()


PRODUCT_LOOKUP: dict[str, str] = {}
for _canonical, _variants in PRODUCT_FAMILIES.items():
    PRODUCT_LOOKUP[_key(_canonical)] = _canonical
    for _variant in _variants:
        PRODUCT_LOOKUP[_key(_variant)] = _canonical
del _canonical, _variants, _variant


def normalize_product(name: Any, short_name: Any = None) -> tuple[str, bool]:
    raw = clean_text(name)
    key = _key(raw)
    if key and key in PRODUCT_LOOKUP:
        return PRODUCT_LOOKUP[key], True

    short = clean_text(short_name)
    short_key = _key(short)
    if short_key and short_key not in {"unmapped", "none", "nan", "n/a", "0"}:
        if short_key in PRODUCT_LOOKUP:
            return PRODUCT_LOOKUP[short_key], True
        if short_key == key:
            return raw, True
        if len(short) >= 4:
            return short, False
    return (raw or "Unspecified"), bool(key)


# --------------------------------------------------------------------------- #
# Cleaning helpers
# --------------------------------------------------------------------------- #


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    text = str(value).replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"[ \t]+", " ", text).strip()


def clean_series(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="object")
    return series.map(clean_text)


def to_number(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(series, errors="coerce").astype("float64")


def to_date(series: pd.Series | None) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def uom_to_mt(quantity: pd.Series, uom: pd.Series) -> pd.Series:
    text = uom.map(lambda v: _key(v))
    is_kg = text.str.contains("kg|kilogram", regex=True, na=False)
    return quantity.where(~is_kg, quantity / 1000.0)


# --------------------------------------------------------------------------- #
# Source workbook
# --------------------------------------------------------------------------- #

DB_FILENAME = "dashboard database.xlsx"
DB_CANDIDATES = (
    ROOT.parent / "Business_dashboard" / DB_FILENAME,  # live Power Query file (your PC)
    ROOT / DB_FILENAME,                                # next to the app
    ROOT / "data" / DB_FILENAME,                       # repo copy (Streamlit Cloud)
)


def _setting(name: str, default: str = "") -> str:
    """Read a setting from the environment first, then Streamlit secrets (Cloud)."""
    value = os.environ.get(name)
    if value:
        return str(value).strip()
    try:
        secret = st.secrets.get(name)  # type: ignore[union-attr]
        if secret:
            return str(secret).strip()
    except Exception:
        pass
    return default


GOOGLE_SHEET_ID = _setting("GOOGLE_SHEET_ID")


def dataset_path(path: str | Path | None = None) -> Path:
    """Explicit path → MASTER_EXCEL_PATH → local copy → the Business_dashboard source."""
    if path is not None:
        return Path(path)
    env = os.environ.get("MASTER_EXCEL_PATH")
    if env:
        return Path(env)
    for candidate in DB_CANDIDATES:
        if candidate.exists():
            return candidate
    return DB_CANDIDATES[0]


def db_signature(path: str | Path | None = None) -> float:
    """Cache key. Google Sheets has no file mtime, so roll the cache every 15 minutes."""
    if GOOGLE_SHEET_ID:
        return float(pd.Timestamp.now().floor("15min").timestamp())
    try:
        return float(dataset_path(path).stat().st_mtime)
    except OSError:
        return 0.0


def _read_sheet(path: Path, sheet: str) -> pd.DataFrame:
    frame = pd.read_excel(path, sheet_name=sheet, dtype=object)
    frame.columns = [clean_text(col) for col in frame.columns]
    return frame


def _read_google_sheet(sheet: str) -> pd.DataFrame:
    """Read one tab of a Google Sheet published as CSV (File → Share → Anyone with the link)."""
    from urllib.parse import quote

    url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv&sheet={quote(sheet)}"
    frame = pd.read_csv(url, dtype=object)
    frame.columns = [clean_text(col) for col in frame.columns]
    # gviz returns mixed date strings - normalise them once so the builders never guess.
    for column in frame.columns:
        if str(column).strip().lower() in {"date", "as_at"}:
            frame[column] = pd.to_datetime(frame[column], errors="coerce", format="mixed").dt.strftime("%Y-%m-%d")
    return frame


def _build_sales(raw: pd.DataFrame) -> pd.DataFrame:
    short = raw["Product Short_Name"] if "Product Short_Name" in raw else pd.Series([None] * len(raw))
    names = raw.get("Product Name", pd.Series([None] * len(raw)))
    normalised = [normalize_product(n, s) for n, s in zip(names, short)]

    frame = pd.DataFrame(
        {
            "Date": to_date(raw.get("Date")),
            "SO Number": clean_series(raw.get("SO Number")),
            "Challan No": clean_series(raw.get("Challan No")),
            "Customer": clean_series(raw.get("Customer")),
            "Shippoint": clean_series(raw.get("Shippoint")),
            "District": clean_series(raw.get("District")),
            "Company": clean_series(raw.get("Company")),
            "UoM": clean_series(raw.get("UoM")),
            "Quantity": to_number(raw.get("Quantity")),
            "Sales Value": to_number(raw.get("Sales Value")),
            "Cost Value": to_number(raw.get("Cost Value")),
            "SO Price": to_number(raw.get("SO Price")),
            "Product": [p for p, _ in normalised],
            "Product Mapped": [m for _, m in normalised],
        }
    )
    partner = clean_series(raw.get("Type."))
    frame["Partner"] = partner.where(partner.isin(["AAFL", "AEL", "Customer"]), "Customer")
    frame["Volume MT"] = uom_to_mt(frame["Quantity"], frame["UoM"])
    frame["Is Return"] = (frame["Quantity"] < 0) | (frame["Sales Value"] < 0) | raw.get("Return_Issue", pd.Series()).notna()
    frame["Year"] = frame["Date"].dt.year.astype("Int64")
    frame["Month"] = frame["Date"].dt.month.astype("Int64")
    return frame


def _build_collection(raw: pd.DataFrame) -> pd.DataFrame:
    frame = pd.DataFrame(
        {
            "Date": to_date(raw.get("Date")),
            "Party": clean_series(raw.get("Party Name")),
            "Customer Code": clean_series(raw.get("Customer Code")),
            "Amount": to_number(raw.get("Collection Amount")),
            "Bank": clean_series(raw.get("Bank Name")),
            "Salesman": clean_series(raw.get("Sales Man")),
            "Source": clean_series(raw.get("Source")),
        }
    )
    partner = clean_series(raw.get("Type."))
    frame["Partner"] = partner.where(partner.isin(["AAFL", "AEL", "Customer"]), "Customer")
    frame["Product"] = raw.get("Item Name", pd.Series([None] * len(raw))).map(
        lambda value: PRODUCT_LOOKUP.get(_key(value))
    )
    frame["Year"] = frame["Date"].dt.year.astype("Int64")
    frame["Month"] = frame["Date"].dt.month.astype("Int64")
    return frame


def _build_cogs(raw: pd.DataFrame) -> pd.DataFrame:
    names = raw.get("Item Name", pd.Series([None] * len(raw)))
    normalised = [normalize_product(n) for n in names]
    frame = pd.DataFrame(
        {
            "Date": to_date(raw.get("Date")),
            "Item Code": clean_series(raw.get("Item Code")),
            "Profit Center": clean_series(raw.get("Profit Center")),
            "Quantity": to_number(raw.get("Quantity")),
            "Rate": to_number(raw.get("Avg. COGS")),
            "Amount": to_number(raw.get("Amount")),
            "Product": [p for p, _ in normalised],
            "Product Mapped": [m for _, m in normalised],
        }
    )
    frame["Year"] = frame["Date"].dt.year.astype("Int64")
    frame["Month"] = frame["Date"].dt.month.astype("Int64")
    return frame


def _build_aging(raw: pd.DataFrame) -> pd.DataFrame:
    buckets = ["0-30 Days", "31-60 Days", "61-90 Days", "91-180 Days", "181-365 Days", "365+ Days"]
    date_column = "As_At" if "As_At" in raw.columns else "Date"
    frame = pd.DataFrame(
        {
            "Date": to_date(raw.get(date_column)),
            "GL": clean_series(raw.get("GL")),
            "BP Code": clean_series(raw.get("BP_Code")),
            "BP Name": clean_series(raw.get("BP_Name")),
        }
    )
    for bucket in buckets:
        frame[bucket] = to_number(raw.get(bucket))
    frame["Total"] = to_number(raw.get("Total"))
    frame["Buckets Sum"] = frame[buckets].sum(axis=1)
    # The source has no Overdue column: everything older than 30 days is overdue.
    overdue = raw.get("Overdue")
    frame["Overdue"] = to_number(overdue) if overdue is not None else frame["Total"] - frame["0-30 Days"]
    frame["Overdue"] = frame["Overdue"].fillna(0.0)
    return frame


def _build_stock(raw: pd.DataFrame) -> pd.DataFrame:
    names = raw.get("Item Name", pd.Series([None] * len(raw)))
    normalised = [normalize_product(n) for n in names]
    frame = pd.DataFrame(
        {
            "Date": to_date(raw.get("Date")),
            "Item Code": clean_series(raw.get("Item Code")),
            "UoM": clean_series(raw.get("UoM Name")),
            "Opening": to_number(raw.get("Open Qty")),
            "In Qty": to_number(raw.get("In Qty")),
            "Out Qty": to_number(raw.get("Out Qty")),
            "Closing": to_number(raw.get("Closing Qty")),
            "Warehouse": clean_series(raw.get("Warehouse")),
            "District": clean_series(raw.get("District")),
            "SBU": clean_series(raw.get("SBU")),
            "Product": [p for p, _ in normalised],
            "Product Mapped": [m for _, m in normalised],
        }
    )
    frame["Closing MT"] = uom_to_mt(frame["Closing"], frame["UoM"])
    frame["In MT"] = uom_to_mt(frame["In Qty"], frame["UoM"])
    frame["Out MT"] = uom_to_mt(frame["Out Qty"], frame["UoM"])
    return frame


# --------------------------------------------------------------------------- #
# Theme engine (deep-green brand + light)
# --------------------------------------------------------------------------- #

THEME_TOKENS: dict[str, dict[str, Any]] = {
    "brand": {
        "page": "#06251A",
        "card": "#0C3B2A",
        "sidebar": "#082D1F",
        "ink": "#EAF6EF",
        "muted": "#A7C4B4",
        "line": "#17553C",
        "grid": "rgba(167, 196, 180, 0.18)",
        "table_bg": "#FFFFFF",
        "hover": "#124B36",
        "accent": "#8CC63F",
        "palette": ["#8CC63F", "#4ADE80", "#FACC15", "#2DD4BF", "#A3E635", "#F59E0B", "#34D399", "#60A5FA"],
    },
    "light": {
        "page": "#F4F8F5",
        "card": "#FFFFFF",
        "sidebar": "#FFFFFF",
        "ink": "#0B3D2E",
        "muted": "#5B7A6B",
        "line": "#DCE9E1",
        "grid": "#E8F0EB",
        "table_bg": "#FFFFFF",
        "hover": "#EDF5F0",
        "accent": "#0B5D3B",
        "palette": ["#0B5D3B", "#8CC63F", "#B7791F", "#15803D", "#0EA5A0", "#65A30D", "#2563EB", "#C2410C"],
    },
}

_ACTIVE_THEME = "brand"
PALETTE = list(THEME_TOKENS["brand"]["palette"])
GREEN, LIME, GOLD, EMERALD, TEAL, OLIVE, BLUE, ORANGE = PALETTE
INK = THEME_TOKENS["brand"]["ink"]
MUTED = THEME_TOKENS["brand"]["muted"]
GRID = THEME_TOKENS["brand"]["grid"]


def set_theme(name: Any) -> str:
    global _ACTIVE_THEME, PALETTE, GREEN, LIME, GOLD, EMERALD, TEAL, OLIVE, BLUE, ORANGE, INK, MUTED, GRID
    key = "light" if str(name).strip().lower().startswith("light") else "brand"
    _ACTIVE_THEME = key
    tokens = THEME_TOKENS[key]
    PALETTE = list(tokens["palette"])
    GREEN, LIME, GOLD, EMERALD, TEAL, OLIVE, BLUE, ORANGE = PALETTE
    INK, MUTED, GRID = tokens["ink"], tokens["muted"], tokens["grid"]
    return key


def get_theme() -> str:
    return _ACTIVE_THEME


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #


@st.cache_data(show_spinner=False, max_entries=3)
def load_database(path_str: str, mtime: float) -> dict[str, Any]:
    return load_dataset(path_str, mtime)


def load_dataset(path_str: str, mtime: float = 0.0) -> dict[str, Any]:
    path = Path(path_str)
    use_google = bool(GOOGLE_SHEET_ID)

    if use_google:
        sheets: list[str] = list(CORE_SHEETS)
        missing: list[str] = []
        reader = _read_google_sheet
    else:
        sheets = pd.ExcelFile(path).sheet_names
        missing = [sheet for sheet in CORE_SHEETS if sheet not in sheets]
        reader = lambda name: _read_sheet(path, name)  # noqa: E731

    dataset: dict[str, Any] = {
        "path": f"Google Sheets · {GOOGLE_SHEET_ID}" if use_google else str(path),
        "source": "google" if use_google else "excel",
        "mtime": float(mtime),
        "sheets": list(sheets),
        "missing_sheets": missing,
    }

    dataset["sales"] = _build_sales(reader("SO_DO_Database")) if (use_google or "SO_DO_Database" in sheets) else pd.DataFrame()
    dataset["collection"] = _build_collection(reader("Collection")) if (use_google or "Collection" in sheets) else pd.DataFrame()
    dataset["cogs"] = _build_cogs(reader("COGS_Database")) if (use_google or "COGS_Database" in sheets) else pd.DataFrame()
    dataset["aging"] = _build_aging(reader("Aging_Database")) if (use_google or "Aging_Database" in sheets) else pd.DataFrame()
    dataset["stock"] = _build_stock(reader("Stock_Database")) if (use_google or "Stock_Database" in sheets) else pd.DataFrame()

    dataset["rates"] = build_rate_map(dataset)
    dataset["stock"] = value_stock(dataset["stock"], dataset["rates"])
    dataset["targets"] = load_targets()
    dataset["product_targets"] = load_product_targets()
    dataset["coverage"] = _coverage(dataset)
    dataset["checks"] = validate_pipeline(dataset)
    return dataset


def build_rate_map(dataset: dict[str, Any]) -> pd.Series:
    """BDT per MT for every product: last COGS rate, then last sale price, then family median."""
    cogs = dataset.get("cogs", pd.DataFrame())
    sales = dataset.get("sales", pd.DataFrame())

    rates: dict[str, float] = {}
    if isinstance(cogs, pd.DataFrame) and len(cogs):
        usable = cogs.dropna(subset=["Rate"])
        usable = usable[usable["Rate"] > 0]
        if len(usable):
            last = usable.sort_values("Date").groupby("Product")["Rate"].last()
            rates.update({k: float(v) for k, v in last.items()})

    if isinstance(sales, pd.DataFrame) and len(sales):
        priced = sales[sales["SO Price"] > 0]
        if len(priced):
            last_price = priced.sort_values("Date").groupby("Product")["SO Price"].last()
            for product, price in last_price.items():
                rates.setdefault(product, float(price))

    series = pd.Series(rates, dtype="float64")
    # Fallback for items that were never sold or costed: median of the same name family.
    if len(series):
        families: dict[str, float] = {}
        for product, rate in series.items():
            families.setdefault(str(product).split(" ")[0].lower(), []).append(rate)
        global_median = float(series.median())
        for product in dataset.get("stock", pd.DataFrame()).get("Product", pd.Series(dtype=object)).unique():
            if product in series.index:
                continue
            token = str(product).split(" ")[0].lower()
            family = families.get(token)
            series[product] = float(np.median(family)) if family else global_median
    return series


def value_stock(stock: pd.DataFrame, rates: pd.Series) -> pd.DataFrame:
    if stock is None or stock.empty:
        return stock
    frame = stock.copy()
    frame["Rate"] = frame["Product"].map(rates)
    frame["Value Cr"] = frame["Closing MT"].fillna(0) * frame["Rate"].fillna(0) / CRORE
    return frame


def _coverage(dataset: dict[str, Any]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key in ("sales", "cogs", "collection", "aging", "stock"):
        frame = dataset.get(key, pd.DataFrame())
        if isinstance(frame, pd.DataFrame) and len(frame) and "Date" in frame:
            dates = frame["Date"].dropna()
            out[key] = {
                "rows": int(len(frame)),
                "first": dates.min(),
                "last": dates.max(),
                "bad_dates": int(frame["Date"].isna().sum()),
            }
        else:
            out[key] = {"rows": 0, "first": pd.NaT, "last": pd.NaT, "bad_dates": 0}
    return out


# --------------------------------------------------------------------------- #
# Targets (editable CSVs shipped with the project)
# --------------------------------------------------------------------------- #


def load_targets() -> dict[str, dict[str, float]]:
    path = ROOT / "targets.csv"
    if not path.exists():
        return {}
    try:
        frame = pd.read_csv(path)
    except Exception:
        return {}
    frame.columns = [str(c).strip() for c in frame.columns]
    targets: dict[str, dict[str, float]] = {}
    for _, row in frame.iterrows():
        month = str(row.get("Month", "")).strip()
        if not month or month.lower() == "nan":
            continue
        targets[month] = {
            "sales": float(pd.to_numeric(row.get("Sales_Target_Cr"), errors="coerce") or 0),
            "collection": float(pd.to_numeric(row.get("Collection_Target_Cr"), errors="coerce") or 0),
        }
    return targets


def load_product_targets() -> pd.DataFrame:
    path = ROOT / "product_targets.csv"
    if not path.exists():
        return pd.DataFrame(columns=["Month", "Product", "Target_Cr"])
    try:
        frame = pd.read_csv(path)
    except Exception:
        return pd.DataFrame(columns=["Month", "Product", "Target_Cr"])
    frame["Month"] = frame["Month"].astype(str).str.strip()
    frame["Target_Cr"] = pd.to_numeric(frame["Target_Cr"], errors="coerce").fillna(0.0)
    return frame


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #


def _check(name: str, detail: str, status: str) -> dict[str, str]:
    return {"Check": name, "Detail": detail, "Status": status}


def _status(ok: bool, warn: bool = False) -> str:
    return "PASS" if ok else ("WARN" if warn else "FAIL")


def validate_pipeline(dataset: dict[str, Any]) -> list[dict[str, str]]:
    checks: list[dict[str, str]] = []
    sales = dataset.get("sales", pd.DataFrame())
    cogs = dataset.get("cogs", pd.DataFrame())
    collection = dataset.get("collection", pd.DataFrame())
    aging = dataset.get("aging", pd.DataFrame())
    stock = dataset.get("stock", pd.DataFrame())

    missing = dataset.get("missing_sheets") or []
    bits = [f"{sheet}={len(dataset.get(key, pd.DataFrame())):,}" for key, sheet in zip(
        ("sales", "cogs", "collection", "aging", "stock"), CORE_SHEETS)]
    checks.append(
        _check(
            "Core sheets loaded",
            ("missing: " + ", ".join(missing) + " | " if missing else "") + " · ".join(bits),
            _status(not missing and len(sales) > 0 and len(stock) > 0),
        )
    )

    bad_dates = sum(
        int(dataset.get(key, pd.DataFrame()).get("Date", pd.Series(dtype=float)).isna().sum())
        for key in ("sales", "cogs", "collection", "aging", "stock")
        if isinstance(dataset.get(key), pd.DataFrame) and len(dataset.get(key))
    )
    checks.append(_check("Date integrity", f"{bad_dates} rows with corrupt or null dates",
                         _status(bad_dates == 0, bad_dates <= 2)))

    if len(sales):
        total = float(sales["Sales Value"].sum())
        by_product = float(sales.groupby("Product")["Sales Value"].sum().sum())
        by_customer = float(sales.groupby("Customer")["Sales Value"].sum().sum())
        checks.append(_check("Product sales sum == total sales",
                             f"Δ {abs(by_product - total):,.2f} BDT", _status(abs(by_product - total) < 1)))
        checks.append(_check("Customer sales sum == total sales",
                             f"Δ {abs(by_customer - total):,.2f} BDT", _status(abs(by_customer - total) < 1)))

        returns = int(sales["Is Return"].sum())
        share = returns / len(sales) * 100
        checks.append(_check("Negative / return rows",
                             f"{returns:,} of {len(sales):,} ({share:.2f}%), net {sales.loc[sales['Is Return'], 'Sales Value'].sum() / CRORE:,.2f} Cr",
                             _status(share <= 2.0, share <= 5.0)))

        unmapped = sorted(sales.loc[~sales["Product Mapped"], "Product"].dropna().unique().tolist())
        checks.append(_check("Unclassified products",
                             f"{len(unmapped)} name(s)" + (": " + ", ".join(unmapped[:5]) if unmapped else ""),
                             _status(len(unmapped) == 0, len(unmapped) <= 5)))

        gap = (sales["Quantity"] * sales["SO Price"] - sales["Sales Value"]).abs()
        checks.append(_check("Quantity × rate == sales value",
                             f"p99 {float(gap.quantile(0.99)):,.2f} BDT", _status(float(gap.quantile(0.99)) < 1)))

        priced = sales[sales["Quantity"] != 0]
        unit = (priced["Sales Value"] / priced["Quantity"]).replace([np.inf, -np.inf], np.nan)
        odd = int(((unit < 8000) | (unit > 1_000_000)).sum())
        odd_share = odd / len(sales) * 100
        checks.append(_check("Unit-price sanity (BDT/MT)",
                             f"{odd:,} row(s) outside 8,000–1,000,000 ({odd_share:.2f}%)",
                             _status(odd_share <= 0.5, odd_share <= 2.0)))

    if len(aging):
        latest = aging[aging["Date"] == aging["Date"].max()]
        total = float(latest["Total"].sum())
        buckets = float(latest["Buckets Sum"].sum())
        overdue = float(latest["Overdue"].sum())
        checks.append(_check("Aging buckets sum == Total",
                             f"as at {aging['Date'].max():%d %b %Y}: Δ {abs(buckets - total):,.0f} BDT",
                             _status(abs(buckets - total) <= max(1.0, total * 1e-6))))
        checks.append(_check("Overdue ≤ receivable total",
                             f"overdue {overdue / CRORE:,.2f} Cr vs total {total / CRORE:,.2f} Cr",
                             _status(0 <= overdue <= total * 1.0001)))

    if len(stock):
        snapshot = stock_snapshot(stock, stock["Date"].max())
        unrated = int((snapshot["Rate"].fillna(0) <= 0).sum()) if len(snapshot) else 0
        checks.append(_check("Stock valuation",
                             f"{len(snapshot)} rows, {snapshot['Closing MT'].sum():,.0f} MT, "
                             f"{snapshot['Value Cr'].sum():,.2f} Cr, {unrated} row(s) without a rate",
                             _status(unrated == 0, unrated <= 2)))

    if len(sales) and len(cogs):
        last_month = sales["Date"].dt.to_period("M").max()
        invoice = float(sales.loc[sales["Date"].dt.to_period("M") == last_month, "Cost Value"].sum())
        cogs_amount = float(cogs.loc[cogs["Date"].dt.to_period("M") == last_month, "Amount"].sum())
        ratio = cogs_amount / invoice if invoice else 0.0
        checks.append(_check("COGS coverage (latest month)",
                             f"{last_month}: COGS {cogs_amount / CRORE:,.2f} Cr vs invoice cost {invoice / CRORE:,.2f} Cr (ratio {ratio:.2f})",
                             _status(0.85 <= ratio <= 1.15, 0.6 <= ratio <= 1.4)))

    if len(collection):
        bad = int(collection["Date"].isna().sum() + collection["Amount"].isna().sum())
        checks.append(_check("Collection completeness",
                             f"{bad:,} row(s) without date/amount · {collection['Amount'].sum() / CRORE:,.2f} Cr",
                             _status(bad == 0, bad <= 5)))

    targets = dataset.get("targets", {})
    checks.append(_check("Targets",
                         f"{len(targets)} month(s) in targets.csv" if targets else "targets.csv missing — target views hidden",
                         "PASS" if targets else "WARN"))
    return checks


def check_summary(checks: Iterable[dict[str, str]]) -> dict[str, int]:
    summary = {"PASS": 0, "WARN": 0, "FAIL": 0}
    for check in checks:
        status = check.get("Status", "WARN")
        summary[status] = summary.get(status, 0) + 1
    return summary


def overall_status(checks: Iterable[dict[str, str]]) -> str:
    summary = check_summary(checks)
    if summary.get("FAIL"):
        return "FAIL"
    return "WARN" if summary.get("WARN") else "PASS"


def coverage_table(dataset: dict[str, Any]) -> pd.DataFrame:
    labels = {
        "sales": "SO_DO_Database",
        "cogs": "COGS_Database",
        "collection": "Collection",
        "aging": "Aging_Database",
        "stock": "Stock_Database",
    }
    rows = []
    for key, label in labels.items():
        meta = dataset.get("coverage", {}).get(key, {})
        rows.append(
            {
                "Sheet": label,
                "Rows": int(meta.get("rows", 0)),
                "First Date": meta.get("first"),
                "Latest Date": meta.get("last"),
                "Null Dates": int(meta.get("bad_dates", 0)),
            }
        )
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Selection context
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class FilterContext:
    start: pd.Timestamp
    end: pd.Timestamp
    prior_start: pd.Timestamp
    prior_end: pd.Timestamp
    district: str | None = None
    warehouse: str | None = None
    product: str | None = None
    year: int = 0
    month: int = 0

    @property
    def month_key(self) -> str:
        return f"{self.year}-{self.month:02d}"

    @property
    def period_label(self) -> str:
        return f"{self.start:%d %b %Y} → {self.end:%d %b %Y}"

    @property
    def month_label(self) -> str:
        return f"{MONTH_NAMES[self.month - 1]} {self.year}" if self.month else self.period_label

    @property
    def prior_label(self) -> str:
        return f"{self.prior_start:%d %b} → {self.prior_end:%d %b %Y}"

    @property
    def days_in_month(self) -> int:
        return int((self.start + pd.offsets.MonthEnd(0) - self.start).days) + 1

    @property
    def days_remaining(self) -> int:
        month_end = self.start + pd.offsets.MonthEnd(0)
        return max(0, int((month_end - self.end).days))

    @property
    def elapsed_pct(self) -> float:
        return self.end.day / self.days_in_month * 100 if self.days_in_month else 0.0

    @property
    def location_label(self) -> str:
        return self.warehouse or (f"{self.district} district" if self.district else "All locations")

    @property
    def product_label(self) -> str:
        return self.product or "All products"


def build_context(
    start: date | datetime | pd.Timestamp,
    end: date | datetime | pd.Timestamp,
    district: str | None = None,
    warehouse: str | None = None,
    product: str | None = None,
) -> FilterContext:
    start_ts = pd.Timestamp(start).normalize()
    end_ts = pd.Timestamp(end).normalize()
    if end_ts < start_ts:
        start_ts, end_ts = end_ts, start_ts

    # Prior window: the previous month up to the same day (month-to-date comparison).
    if start_ts.day == 1:
        prev_month_end = start_ts - pd.Timedelta(days=1)
        prev_start = prev_month_end.replace(day=1)
        prev_end = min(prev_start + pd.Timedelta(days=end_ts.day - 1), prev_month_end)
    else:
        span = (end_ts - start_ts).days
        prev_end = start_ts - pd.Timedelta(days=1)
        prev_start = prev_end - pd.Timedelta(days=span)

    return FilterContext(
        start=start_ts,
        end=end_ts,
        prior_start=prev_start,
        prior_end=prev_end,
        district=district or None,
        warehouse=warehouse or None,
        product=product or None,
        year=int(start_ts.year),
        month=int(start_ts.month),
    )


def slice_dates(frame: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return frame.iloc[0:0] if isinstance(frame, pd.DataFrame) else pd.DataFrame()
    return frame[(frame["Date"] >= pd.Timestamp(start)) & (frame["Date"] <= pd.Timestamp(end))]


def filter_sales(sales: pd.DataFrame, ctx: FilterContext, prior: bool = False) -> pd.DataFrame:
    start, end = (ctx.prior_start, ctx.prior_end) if prior else (ctx.start, ctx.end)
    frame = slice_dates(sales, start, end)
    if frame.empty:
        return frame
    if ctx.district:
        frame = frame[frame["District"] == ctx.district]
    if ctx.product:
        frame = frame[frame["Product"] == ctx.product]
    return frame


def filter_cogs(cogs: pd.DataFrame, ctx: FilterContext, prior: bool = False) -> pd.DataFrame:
    start, end = (ctx.prior_start, ctx.prior_end) if prior else (ctx.start, ctx.end)
    frame = slice_dates(cogs, start, end)
    if frame.empty:
        return frame
    if ctx.product:
        frame = frame[frame["Product"] == ctx.product]
    return frame


def filter_collection(collection: pd.DataFrame, ctx: FilterContext, prior: bool = False) -> pd.DataFrame:
    start, end = (ctx.prior_start, ctx.prior_end) if prior else (ctx.start, ctx.end)
    frame = slice_dates(collection, start, end)
    if frame.empty:
        return frame
    if ctx.product:
        frame = frame[frame["Product"] == ctx.product]
    return frame


def stock_snapshot(
    stock: pd.DataFrame,
    end_dt: pd.Timestamp,
    district: str | None = None,
    warehouse: str | None = None,
    product: str | None = None,
) -> pd.DataFrame:
    if not isinstance(stock, pd.DataFrame) or stock.empty:
        return pd.DataFrame(columns=["Warehouse", "District", "Product", "Closing MT", "Value Cr", "Date"])
    frame = stock[stock["Date"] <= pd.Timestamp(end_dt)]
    if frame.empty:
        return frame.iloc[0:0]
    snap = frame[frame["Date"] == frame["Date"].max()]
    snap = snap.sort_values("Date").groupby(["Warehouse", "District", "Product"], dropna=False).tail(1)
    if district:
        snap = snap[snap["District"] == district]
    if warehouse:
        snap = snap[snap["Warehouse"] == warehouse]
    if product:
        snap = snap[snap["Product"] == product]
    return snap[snap["Closing MT"] > 0.01].copy()


def aging_snapshot(aging: pd.DataFrame, end_dt: pd.Timestamp) -> pd.DataFrame:
    if not isinstance(aging, pd.DataFrame) or aging.empty:
        return pd.DataFrame()
    frame = aging[aging["Date"] <= pd.Timestamp(end_dt)]
    if frame.empty:
        return frame.iloc[0:0]
    return frame[frame["Date"] == frame["Date"].max()].copy()


def aging_previous_snapshot(aging: pd.DataFrame, end_dt: pd.Timestamp) -> pd.DataFrame:
    """The snapshot immediately before the one used for the current period."""
    if not isinstance(aging, pd.DataFrame) or aging.empty:
        return pd.DataFrame()
    frame = aging[aging["Date"] <= pd.Timestamp(end_dt)]
    if frame.empty:
        return frame.iloc[0:0]
    dates = sorted(frame["Date"].dropna().unique())
    if len(dates) < 2:
        return frame.iloc[0:0]
    return frame[frame["Date"] == dates[-2]].copy()


# --------------------------------------------------------------------------- #
# Filter option lists
# --------------------------------------------------------------------------- #


def available_periods(dataset: dict[str, Any]) -> dict[int, list[int]]:
    sales = dataset.get("sales", pd.DataFrame())
    stock = dataset.get("stock", pd.DataFrame())
    dates = pd.concat([
        sales["Date"].dropna() if len(sales) else pd.Series(dtype="datetime64[ns]"),
        stock["Date"].dropna() if len(stock) else pd.Series(dtype="datetime64[ns]"),
    ])
    periods: dict[int, set[int]] = {}
    for stamp in dates:
        periods.setdefault(int(stamp.year), set()).add(int(stamp.month))
    return {year: sorted(months) for year, months in sorted(periods.items(), reverse=True)}


def default_period(dataset: dict[str, Any]) -> tuple[int, int]:
    periods = available_periods(dataset)
    today = pd.Timestamp.today()
    if today.year in periods and today.month in periods[today.year]:
        return today.year, today.month
    if periods:
        year = next(iter(periods))
        return year, max(periods[year])
    return today.year, today.month


def available_districts(dataset: dict[str, Any]) -> list[str]:
    values: set[str] = set()
    for key in ("sales", "stock"):
        frame = dataset.get(key, pd.DataFrame())
        if isinstance(frame, pd.DataFrame) and len(frame):
            values.update(v for v in frame["District"].unique() if v)
    return sorted(values)


def available_warehouses(dataset: dict[str, Any], district: str | None = None) -> list[str]:
    stock = dataset.get("stock", pd.DataFrame())
    if not isinstance(stock, pd.DataFrame) or stock.empty:
        return []
    frame = stock[stock["District"] == district] if district else stock
    return sorted(v for v in frame["Warehouse"].unique() if v)


def available_products(dataset: dict[str, Any]) -> list[str]:
    values: set[str] = set()
    for key in ("sales", "stock"):
        frame = dataset.get(key, pd.DataFrame())
        if isinstance(frame, pd.DataFrame) and len(frame) and "Product" in frame:
            values.update(v for v in frame["Product"].unique() if v)
    return sorted(values)


# --------------------------------------------------------------------------- #
# KPI bundle
# --------------------------------------------------------------------------- #


def pct(current: float | None, previous: float | None) -> float | None:
    if current is None or previous is None or pd.isna(current) or pd.isna(previous) or abs(previous) < 1e-9:
        return None
    return (current - previous) / abs(previous) * 100.0


def _gross_profit(sales_slice: pd.DataFrame, cogs_slice: pd.DataFrame, ctx: FilterContext) -> tuple[float, str]:
    sales_cr = float(sales_slice["Sales Value"].sum()) / CRORE if len(sales_slice) else 0.0
    invoice_cost = float(sales_slice["Cost Value"].sum()) / CRORE if len(sales_slice) else 0.0

    if ctx.district or ctx.warehouse:
        return sales_cr - invoice_cost, "Invoice cost"
    cogs_cr = float(cogs_slice["Amount"].sum()) / CRORE if isinstance(cogs_slice, pd.DataFrame) and len(cogs_slice) else 0.0
    if cogs_cr > 0:
        return sales_cr - cogs_cr, "COGS database"
    return sales_cr - invoice_cost, "Invoice cost"


def _ar_split(aging: pd.DataFrame) -> dict[str, float]:
    if aging is None or aging.empty:
        return {"AAFL": 0.0, "AEL": 0.0, "Customer": 0.0}
    names = aging["BP Name"].astype(str)
    group = np.where(names.str.startswith("Akij Agro Feed"), "AAFL",
                     np.where(names.str.startswith("Akij Essentials"), "AEL", "Customer"))
    return {k: float(v) / CRORE for k, v in aging.groupby(group)["Total"].sum().items()}


def build_view(dataset: dict[str, Any], ctx: FilterContext) -> dict[str, Any]:
    sales = dataset.get("sales", pd.DataFrame())
    cogs = dataset.get("cogs", pd.DataFrame())
    collection = dataset.get("collection", pd.DataFrame())
    aging = dataset.get("aging", pd.DataFrame())
    stock = dataset.get("stock", pd.DataFrame())

    sales_cur = filter_sales(sales, ctx)
    sales_prior = filter_sales(sales, ctx, prior=True)
    cogs_cur = filter_cogs(cogs, ctx)
    cogs_prior = filter_cogs(cogs, ctx, prior=True)
    coll_cur = filter_collection(collection, ctx)
    coll_prior = filter_collection(collection, ctx, prior=True)
    stock_cur = stock_snapshot(stock, ctx.end, ctx.district, ctx.warehouse, ctx.product)
    stock_prior = stock_snapshot(stock, ctx.prior_end, ctx.district, ctx.warehouse, ctx.product)
    aging_cur = aging_snapshot(aging, ctx.end)
    aging_prior = aging_previous_snapshot(aging, ctx.end)

    sales_cr = float(sales_cur["Sales Value"].sum()) / CRORE if len(sales_cur) else 0.0
    sales_prior_cr = float(sales_prior["Sales Value"].sum()) / CRORE if len(sales_prior) else 0.0
    volume_mt = float(sales_cur["Volume MT"].sum()) if len(sales_cur) else 0.0
    volume_prior = float(sales_prior["Volume MT"].sum()) if len(sales_prior) else 0.0

    gp_cr, gp_basis = _gross_profit(sales_cur, cogs_cur, ctx)
    gp_prior_cr, _ = _gross_profit(sales_prior, cogs_prior, ctx)
    gp_pct = gp_cr / sales_cr * 100 if sales_cr else 0.0
    gp_prior_pct = gp_prior_cr / sales_prior_cr * 100 if sales_prior_cr else 0.0

    collection_cr = float(coll_cur["Amount"].sum()) / CRORE if len(coll_cur) else 0.0
    collection_prior_cr = float(coll_prior["Amount"].sum()) / CRORE if len(coll_prior) else 0.0

    receivable_cr = float(aging_cur["Total"].sum()) / CRORE if len(aging_cur) else 0.0
    receivable_prior = float(aging_prior["Total"].sum()) / CRORE if len(aging_prior) else 0.0
    overdue_cr = float(aging_cur["Overdue"].sum()) / CRORE if len(aging_cur) else 0.0

    inventory_cr = float(stock_cur["Value Cr"].sum()) if len(stock_cur) else 0.0
    inventory_mt = float(stock_cur["Closing MT"].sum()) if len(stock_cur) else 0.0
    inventory_prior_cr = float(stock_prior["Value Cr"].sum()) if len(stock_prior) else 0.0

    targets = dataset.get("targets", {})
    target = targets.get(ctx.month_key, {})
    product_targets = dataset.get("product_targets", pd.DataFrame())
    if ctx.product and isinstance(product_targets, pd.DataFrame) and len(product_targets):
        match = product_targets[(product_targets["Month"] == ctx.month_key) & (product_targets["Product"] == ctx.product)]
        sales_target = float(match["Target_Cr"].sum()) if len(match) else None
    else:
        sales_target = target.get("sales") if target else None
    collection_target = target.get("collection") if target else None

    asking_month = (sales_target - sales_cr) if sales_target else None
    asking_day = (asking_month / ctx.days_remaining) if asking_month and ctx.days_remaining > 0 else None

    # Stock cover in months, using the trailing 3-month run-rate.
    recent = slice_dates(sales, ctx.end - pd.DateOffset(months=3), ctx.end)
    monthly_run = float(recent["Volume MT"].sum()) / 3 if len(recent) else 0.0
    stock_cover = (inventory_mt / monthly_run) if monthly_run else None

    return {
        "ctx": ctx,
        "sales": sales_cur,
        "sales_prior": sales_prior,
        "cogs": cogs_cur,
        "collection": coll_cur,
        "collection_prior": coll_prior,
        "stock": stock_cur,
        "stock_prior": stock_prior,
        "aging": aging_cur,
        "aging_prior": aging_prior,
        "sales_cr": sales_cr,
        "sales_prior_cr": sales_prior_cr,
        "sales_delta": pct(sales_cr, sales_prior_cr),
        "sales_target": sales_target,
        "sales_ach": (sales_cr / sales_target * 100) if sales_target else None,
        "sales_vs_target": (sales_cr - sales_target) if sales_target else None,
        "volume_mt": volume_mt,
        "volume_prior_mt": volume_prior,
        "so_count": int(sales_cur["SO Number"].nunique()) if len(sales_cur) else 0,
        "customer_count": int(sales_cur["Customer"].nunique()) if len(sales_cur) else 0,
        "gp_cr": gp_cr,
        "gp_prior_cr": gp_prior_cr,
        "gp_pct": gp_pct,
        "gp_prior_pct": gp_prior_pct,
        "gp_delta": pct(gp_cr, gp_prior_cr),
        "gp_delta_pp": gp_pct - gp_prior_pct,
        "gp_basis": gp_basis,
        "collection_cr": collection_cr,
        "collection_prior_cr": collection_prior_cr,
        "collection_delta": pct(collection_cr, collection_prior_cr),
        "collection_target": collection_target,
        "collection_ach": (collection_cr / collection_target * 100) if collection_target else None,
        "collection_vs_sales": (collection_cr / sales_cr) if sales_cr else None,
        "receivable_cr": receivable_cr,
        "receivable_prior_cr": receivable_prior,
        "receivable_change": receivable_cr - receivable_prior,
        "overdue_cr": overdue_cr,
        "overdue_pct": (overdue_cr / receivable_cr * 100) if receivable_cr else 0.0,
        "ar_split": _ar_split(aging_cur),
        "aging_as_of": aging_cur["Date"].max() if len(aging_cur) else pd.NaT,
        "inventory_cr": inventory_cr,
        "inventory_mt": inventory_mt,
        "inventory_prior_cr": inventory_prior_cr,
        "inventory_delta": pct(inventory_cr, inventory_prior_cr),
        "sku_count": int(stock_cur["Product"].nunique()) if len(stock_cur) else 0,
        "warehouse_count": int(stock_cur["Warehouse"].nunique()) if len(stock_cur) else 0,
        "stock_as_of": stock_cur["Date"].max() if len(stock_cur) else pd.NaT,
        "stock_cover": stock_cover,
        "asking_month": asking_month,
        "asking_day": asking_day,
        "days_remaining": ctx.days_remaining,
        "elapsed_pct": ctx.elapsed_pct,
    }


# --------------------------------------------------------------------------- #
# Tables
# --------------------------------------------------------------------------- #


def product_performance(
    sales: pd.DataFrame,
    cogs: pd.DataFrame,
    product_targets: pd.DataFrame | None = None,
    month_key: str | None = None,
) -> pd.DataFrame:
    columns = ["Product", "Orders", "Qty (MT)", "Sales (Cr)", "COGS (Cr)", "GP (Cr)", "GP %", "Target (Cr)", "Ach %"]
    if sales is None or sales.empty:
        return pd.DataFrame(columns=columns)

    grouped = sales.groupby("Product", dropna=False).agg(
        Orders=pd.NamedAgg(column="SO Number", aggfunc="nunique"),
        Qty=pd.NamedAgg(column="Volume MT", aggfunc="sum"),
        Sales=pd.NamedAgg(column="Sales Value", aggfunc="sum"),
        InvoiceCost=pd.NamedAgg(column="Cost Value", aggfunc="sum"),
    )
    sales_cr = grouped["Sales"] / CRORE
    cogs_map = (
        cogs.groupby("Product")["Amount"].sum() / CRORE
        if isinstance(cogs, pd.DataFrame) and len(cogs) else pd.Series(dtype=float)
    )
    cogs_cr = grouped.index.map(lambda name: float(cogs_map.get(name, 0.0)))
    cogs_cr = np.where(np.asarray(cogs_cr) > 0, cogs_cr, (grouped["InvoiceCost"] / CRORE).to_numpy())
    gp_cr = sales_cr.to_numpy() - cogs_cr
    gp_pct = np.where(sales_cr.to_numpy() != 0, gp_cr / sales_cr.to_numpy() * 100, 0.0)

    targets_map: dict[str, float] = {}
    if isinstance(product_targets, pd.DataFrame) and len(product_targets) and month_key:
        rows = product_targets[product_targets["Month"] == month_key]
        targets_map = {str(r["Product"]): float(r["Target_Cr"]) for _, r in rows.iterrows()}

    result = pd.DataFrame(
        {
            "Product": grouped.index,
            "Orders": grouped["Orders"].astype(int),
            "Qty (MT)": grouped["Qty"].round(2),
            "Sales (Cr)": sales_cr.round(3),
            "COGS (Cr)": np.round(cogs_cr, 3),
            "GP (Cr)": np.round(gp_cr, 3),
            "GP %": np.round(gp_pct, 1),
        }
    )
    result["Target (Cr)"] = result["Product"].map(targets_map).fillna(0.0)
    result["Ach %"] = np.where(result["Target (Cr)"] > 0, result["Sales (Cr)"] / result["Target (Cr)"] * 100, np.nan)
    result["Ach %"] = result["Ach %"].round(1)
    return result.sort_values("Sales (Cr)", ascending=False).reset_index(drop=True)


def target_achievement_table(
    performance: pd.DataFrame,
    product_targets: pd.DataFrame,
    month_key: str,
) -> pd.DataFrame:
    """Product targets vs achieved, supporting grouped products (``SBM|Soybean (Local)``)."""
    columns = ["Product", "Target (Cr)", "Achieved (Cr)", "Ach %"]
    if not isinstance(product_targets, pd.DataFrame) or product_targets.empty:
        return pd.DataFrame(columns=columns)
    rows = product_targets[product_targets["Month"] == month_key]
    if rows.empty:
        return pd.DataFrame(columns=columns)

    achieved_rows: list[dict[str, Any]] = []
    covered: set[str] = set()
    for _, row in rows.iterrows():
        label = str(row["Product"]).strip()
        members = [normalize_product(part.strip())[0] for part in label.split("|")]
        covered.update(members)
        subset = performance[performance["Product"].isin(members)] if len(performance) else performance
        achieved_rows.append(
            {
                "Product": label.split("|")[0].strip(),
                "Target (Cr)": float(row["Target_Cr"]),
                "Achieved (Cr)": float(subset["Sales (Cr)"].sum()) if len(subset) else 0.0,
            }
        )

    if len(performance):
        rest = performance[~performance["Product"].isin(covered)]
        if len(rest) and float(rest["Sales (Cr)"].sum()) > 0:
            achieved_rows.append(
                {"Product": "Others", "Target (Cr)": 0.0, "Achieved (Cr)": float(rest["Sales (Cr)"].sum())}
            )

    frame = pd.DataFrame(achieved_rows)
    frame["Ach %"] = np.where(frame["Target (Cr)"] > 0, frame["Achieved (Cr)"] / frame["Target (Cr)"] * 100, np.nan)
    frame = frame.sort_values("Target (Cr)", ascending=False).reset_index(drop=True)
    total_target = float(frame["Target (Cr)"].sum())
    total_achieved = float(frame["Achieved (Cr)"].sum())
    frame.loc[len(frame)] = {
        "Product": "Total",
        "Target (Cr)": total_target,
        "Achieved (Cr)": total_achieved,
        "Ach %": (total_achieved / total_target * 100) if total_target else np.nan,
    }
    frame["Target (Cr)"] = frame["Target (Cr)"].round(2)
    frame["Achieved (Cr)"] = frame["Achieved (Cr)"].round(2)
    frame["Ach %"] = frame["Ach %"].round(1)
    return frame


def customer_performance(sales: pd.DataFrame) -> pd.DataFrame:
    columns = ["Customer", "Group", "Orders", "Qty (MT)", "Sales (Cr)", "Invoice GP (Cr)", "GP %", "Share %"]
    if sales is None or sales.empty:
        return pd.DataFrame(columns=columns)
    grouped = sales.groupby("Customer", dropna=False).agg(
        Group=pd.NamedAgg(column="Partner", aggfunc="first"),
        Orders=pd.NamedAgg(column="SO Number", aggfunc="nunique"),
        Qty=pd.NamedAgg(column="Volume MT", aggfunc="sum"),
        Sales=pd.NamedAgg(column="Sales Value", aggfunc="sum"),
        Cost=pd.NamedAgg(column="Cost Value", aggfunc="sum"),
    )
    total = float(grouped["Sales"].sum())
    result = pd.DataFrame(
        {
            "Customer": grouped.index,
            "Group": grouped["Group"],
            "Orders": grouped["Orders"].astype(int),
            "Qty (MT)": grouped["Qty"].round(2),
            "Sales (Cr)": (grouped["Sales"] / CRORE).round(3),
            "Invoice GP (Cr)": ((grouped["Sales"] - grouped["Cost"]) / CRORE).round(3),
        }
    )
    result["GP %"] = np.where(grouped["Sales"].to_numpy() != 0,
                              (grouped["Sales"] - grouped["Cost"]).to_numpy() / grouped["Sales"].to_numpy() * 100, 0.0).round(1)
    result["Share %"] = ((grouped["Sales"] / total * 100) if total else 0).round(1).to_numpy()
    return result.sort_values("Sales (Cr)", ascending=False).reset_index(drop=True)


def collections_by_party(collection: pd.DataFrame) -> pd.DataFrame:
    columns = ["Party", "Group", "Receipts (Cr)", "Receipts", "Main Bank", "Salesman"]
    if collection is None or collection.empty:
        return pd.DataFrame(columns=columns)
    grouped = collection.groupby("Party", dropna=False).agg(
        Group=pd.NamedAgg(column="Partner", aggfunc="first"),
        Amount=pd.NamedAgg(column="Amount", aggfunc="sum"),
        Receipts=pd.NamedAgg(column="Amount", aggfunc="size"),
        Bank=pd.NamedAgg(column="Bank", aggfunc=lambda s: s.mode().iat[0] if len(s.mode()) else ""),
        Salesman=pd.NamedAgg(column="Salesman", aggfunc="first"),
    )
    result = pd.DataFrame(
        {
            "Party": grouped.index,
            "Group": grouped["Group"],
            "Receipts (Cr)": (grouped["Amount"] / CRORE).round(3),
            "Receipts": grouped["Receipts"].astype(int),
            "Main Bank": grouped["Bank"],
            "Salesman": grouped["Salesman"],
        }
    )
    return result.sort_values("Receipts (Cr)", ascending=False).reset_index(drop=True)


def stock_detail(stock: pd.DataFrame) -> pd.DataFrame:
    columns = ["District", "Warehouse", "Product", "Closing (MT)", "Rate (BDT/MT)", "Value (Cr)"]
    if stock is None or stock.empty:
        return pd.DataFrame(columns=columns)
    result = pd.DataFrame(
        {
            "District": stock["District"],
            "Warehouse": stock["Warehouse"],
            "Product": stock["Product"],
            "Closing (MT)": stock["Closing MT"].round(2),
            "Rate (BDT/MT)": stock["Rate"].round(0),
            "Value (Cr)": stock["Value Cr"].round(3),
        }
    )
    return result.sort_values("Value (Cr)", ascending=False).reset_index(drop=True)


def aging_detail(aging: pd.DataFrame) -> pd.DataFrame:
    columns = ["BP Code", "BP Name", "0-30", "31-60", "61-90", "91-180", "181-365", "365+", "Total (Cr)", "Overdue (Cr)"]
    if aging is None or aging.empty:
        return pd.DataFrame(columns=columns)
    result = pd.DataFrame(
        {
            "BP Code": aging["BP Code"],
            "BP Name": aging["BP Name"],
            "0-30": aging["0-30 Days"] / CRORE,
            "31-60": aging["31-60 Days"] / CRORE,
            "61-90": aging["61-90 Days"] / CRORE,
            "91-180": aging["91-180 Days"] / CRORE,
            "181-365": aging["181-365 Days"] / CRORE,
            "365+": aging["365+ Days"] / CRORE,
            "Total (Cr)": aging["Total"] / CRORE,
            "Overdue (Cr)": aging["Overdue"] / CRORE,
        }
    )
    return result.sort_values("Total (Cr)", ascending=False).reset_index(drop=True)


def aging_buckets(aging: pd.DataFrame) -> pd.DataFrame:
    buckets = ["0-30 Days", "31-60 Days", "61-90 Days", "91-180 Days", "181-365 Days", "365+ Days"]
    if aging is None or aging.empty:
        return pd.DataFrame(columns=["Bucket", "Cr"])
    return pd.DataFrame({"Bucket": [b.replace(" Days", "") for b in buckets],
                         "Cr": [float(aging[b].sum()) / CRORE for b in buckets]})


def monthly_series(dataset: dict[str, Any], ctx: FilterContext, months: int = 12) -> pd.DataFrame:
    sales = dataset.get("sales", pd.DataFrame())
    collection = dataset.get("collection", pd.DataFrame())
    cogs = dataset.get("cogs", pd.DataFrame())

    end = (ctx.start + pd.offsets.MonthEnd(0)).to_period("M")
    index = pd.period_range(end - (months - 1), end, freq="M")
    rows = []
    for period in index:
        start, stop = period.start_time, period.end_time
        s = slice_dates(sales, start, stop)
        c = slice_dates(collection, start, stop)
        g = slice_dates(cogs, start, stop)
        if ctx.product and len(s):
            s = s[s["Product"] == ctx.product]
        if ctx.product and len(g):
            g = g[g["Product"] == ctx.product]
        if ctx.district and len(s):
            s = s[s["District"] == ctx.district]
        sales_cr = float(s["Sales Value"].sum()) / CRORE if len(s) else 0.0
        coll_cr = float(c["Amount"].sum()) / CRORE if len(c) else 0.0
        cogs_cr = float(g["Amount"].sum()) / CRORE if len(g) else 0.0
        gp = sales_cr - (cogs_cr if cogs_cr > 0 else (float(s["Cost Value"].sum()) / CRORE if len(s) else 0.0))
        rows.append(
            {
                "Month": period.strftime("%b %y"),
                "Period": str(period),
                "Sales (Cr)": round(sales_cr, 2),
                "Collection (Cr)": round(coll_cr, 2),
                "GP (Cr)": round(gp, 2),
                "GP %": round((gp / sales_cr * 100) if sales_cr else 0.0, 1),
                "Qty (MT)": round(float(s["Volume MT"].sum()) if len(s) else 0.0, 1),
            }
        )
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Chart styling (single source of truth)
# --------------------------------------------------------------------------- #

# Categorical palette - colourblind safe, readable on both themes.
# Red / green are reserved for good-bad meaning only.
CATEGORICAL = ["#0F766E", "#F59E0B", "#6366F1", "#EC4899", "#0EA5E9", "#84CC16", "#F97316", "#8B5CF6"]
GOOD = "#16A34A"
BAD = "#DC2626"
WARN = "#F59E0B"
RISK = ["#16A34A", "#65A30D", "#CA8A04", "#F59E0B", "#EA580C", "#DC2626"]
DUE_COLOR = "#EAB308"       # 0-30 days
OVERDUE_COLOR = "#DC2626"   # > 30 days
TEXT_ON_DARK = "#FFFFFF"

# Business partners hidden from the receivables chart (kept in the database).
RECEIVABLE_EXCLUDED = (
    "BENGAL PROTIN",
    "Nobayon Traders Ltd. (FOC)",
    "Japfa Comfeed",
    "Homaira Enterprise",
    "M K Traders",
    "Akij Essentials",
)


def _rgb(hex_color: str, alpha: float) -> str:
    value = hex_color.lstrip("#")
    r, g, b = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def _mix(c1: str, c2: str, t: float) -> str:
    a, b = c1.lstrip("#"), c2.lstrip("#")
    out = []
    for i in (0, 2, 4):
        out.append(int(int(a[i:i + 2], 16) * (1 - t) + int(b[i:i + 2], 16) * t))
    return "#" + "".join(f"{c:02X}" for c in out)


def _gradient(n: int, light: str = "#99F6E4", dark: str = "#0F766E") -> list[str]:
    """Light → dark ramp; index 0 is the darkest so the largest bar stands out."""
    if n <= 0:
        return []
    if n == 1:
        return [dark]
    return [_mix(light, dark, 1 - i / (n - 1)) for i in range(n)]


def style_fig(
    fig: go.Figure,
    height: int = 360,
    theme: str | None = None,
    legend: bool = False,
    yaxis_title: str | None = None,
    xaxis_title: str | None = None,
    hovermode: str = "x unified",
    show_ygrid: bool = True,
    show_xgrid: bool = False,
) -> go.Figure:
    """The one place chart styling lives - every chart_* function ends here."""
    key = theme or get_theme()
    tokens = THEME_TOKENS.get(key, THEME_TOKENS["brand"])
    ink, muted = tokens["ink"], tokens["muted"]
    grid = "rgba(100,116,139,0.16)" if key == "light" else "rgba(226,232,240,0.12)"

    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Segoe UI, sans-serif", size=12, color=ink),
        title=dict(font=dict(size=15, color=ink), x=0, xanchor="left"),
        showlegend=legend,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.0, x=0, xanchor="left",
            font=dict(size=11, color=muted), bgcolor="rgba(0,0,0,0)",
            itemsizing="constant", itemwidth=30, traceorder="normal",
        ),
        hovermode=hovermode,
        hoverlabel=dict(bgcolor="#0F172A", bordercolor=CATEGORICAL[0], font=dict(color="#FFFFFF", size=12)),
        bargap=0.25,
        bargroupgap=0.08,
    )
    fig.update_xaxes(
        showgrid=show_xgrid, zeroline=False, showline=False, ticks="",
        tickfont=dict(size=12, color=muted), automargin=True, title_text=xaxis_title,
        title_font=dict(size=11, color=muted),
        gridcolor=grid if show_xgrid else "rgba(0,0,0,0)",
    )
    fig.update_yaxes(
        showgrid=show_ygrid, zeroline=False, showline=False, ticks="",
        tickfont=dict(size=12, color=muted), automargin=True, title_text=yaxis_title,
        title_font=dict(size=11, color=muted),
        gridcolor=grid if show_ygrid else "rgba(0,0,0,0)",
    )
    return fig


def _empty(message: str = "No data for the selected filters", height: int = 360) -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(text=message, showarrow=False, font=dict(size=13, color=MUTED), x=0.5, y=0.5,
                       xref="paper", yref="paper")
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return style_fig(fig, height, hovermode="closest")


def _label_points(n: int, values: list[float], fmt) -> list[str]:
    """All labels up to 15 points; beyond that only last / max / min (deduplicated)."""
    if n == 0:
        return []
    if n <= 15:
        indexes = range(n)
    else:
        indexes = {n - 1, int(np.argmax(values)), int(np.argmin(values))}
    labels = [""] * n
    seen: set[str] = set()
    for i in indexes:
        text = fmt(values[i])
        if text in seen:
            continue
        seen.add(text)
        labels[i] = text
    return labels


def _sparse_labels(n: int, values: list[float], fmt) -> list[str]:
    """Last, highest and lowest point only - for dense line charts."""
    if n == 0:
        return []
    labels = [""] * n
    for i in {n - 1, int(np.argmax(values)), int(np.argmin(values))}:
        labels[i] = fmt(values[i])
    return labels


def _alternating(n: int, values: list[float], fmt, parity: int) -> list[str]:
    """Label every other point so adjacent labels never collide."""
    return [fmt(values[i]) if i % 2 == parity else "" for i in range(n)]


def _every(n: int, values: list[float], fmt, step: int = 3, offset: int = 0) -> list[str]:
    """Label every `step`-th point - used on 12-month dual-series charts."""
    return [fmt(values[i]) if (i - offset) % step == 0 and i >= offset else "" for i in range(n)]


def _num(v: float) -> str:
    """Plain number - the unit lives in the chart subtitle."""
    return f"{v:,.1f}"


def _pct(v: float) -> str:
    return f"{v:,.1f}%"


def _mt(v: float) -> str:
    return f"{v:,.0f}"


def _hbar(
    labels: list[str],
    values: list[float],
    colors: list[str] | None = None,
    height: int = 360,
    fmt=_num,
    axis: str = "BDT Crore",
    suffix: str = "",
    show_grid: bool = False,
) -> go.Figure:
    """Ranked horizontal bars: names on the left, values outside, 15% headroom."""
    if not labels:
        return _empty("No data for the selected filters", height)

    labels = list(labels)
    values = [float(v) for v in values]
    colors = list(colors) if colors else _gradient(len(values))
    wrapped = [textwrap.fill(str(l), width=30, max_lines=2, placeholder="…") for l in labels]
    top = max(values) if values else 0.0
    texts = [fmt(v) for v in values]

    # Labels are drawn as a text layer so tiny bars still get a readable gap after the axis.
    label_x = [max(v, (top or 1) * 0.025) for v in values]
    fig = go.Figure(
        go.Bar(
            x=values,
            y=wrapped,
            orientation="h",
            marker=dict(color=colors, line=dict(width=0)),
            cliponaxis=False,
            hovertemplate="%{y}<br>%{x:,.2f}" + suffix + "<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=label_x, y=wrapped, mode="text", text=texts,
            textposition="middle right", cliponaxis=False,
            textfont=dict(size=12, color=INK, family="Inter, Segoe UI, sans-serif"),
            showlegend=False, hoverinfo="skip",
        )
    )
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_xaxes(range=[0, top * 1.15 if top > 0 else 1], title_text=axis)
    return style_fig(fig, height, show_xgrid=show_grid, show_ygrid=show_grid)


def chart_monthly_trend(series: pd.DataFrame) -> go.Figure:
    """Sales vs collection - two lines, markers with white outline, area under sales."""
    if series is None or series.empty:
        return _empty("No trend data")

    sales_color, coll_color = CATEGORICAL[0], CATEGORICAL[1]
    sales = series["Sales (Cr)"].astype(float).tolist()
    collection = series["Collection (Cr)"].astype(float).tolist()
    months = series["Month"].tolist()

    # Anchor labels only (peak + latest): 24 labels over 12 months cannot stay readable.
    sales_labels = [""] * len(sales)
    for idx in {len(sales) - 1, int(np.argmax(sales))}:
        sales_labels[idx] = _num(sales[idx])
    coll_labels = [""] * len(collection)
    coll_labels[len(collection) - 1] = _num(collection[len(collection) - 1])

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=months, y=sales, name="Sales", mode="lines+markers+text",
            line=dict(color=sales_color, width=3, shape="spline", smoothing=0.4),
            marker=dict(size=7, color=sales_color, line=dict(color=TEXT_ON_DARK, width=1.5)),
            fill="tozeroy", fillcolor=_rgb(sales_color, 0.10),
            text=sales_labels, textposition="top center",
            textfont=dict(size=12, color=INK), cliponaxis=False,
            hovertemplate="Sales %{y:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=months, y=collection, name="Collection", mode="lines+markers+text",
            line=dict(color=coll_color, width=3, shape="spline", smoothing=0.4),
            marker=dict(size=7, color=coll_color, line=dict(color=TEXT_ON_DARK, width=1.5)),
            text=coll_labels, textposition="bottom center",
            textfont=dict(size=12, color=INK), cliponaxis=False,
            hovertemplate="Collection %{y:,.2f} Cr<extra></extra>",
        )
    )
    top = max(sales + collection) if sales or collection else 0
    span = top * 0.16 + 2
    fig.update_yaxes(range=[-span * 0.55, top + span], title_text="BDT Crore")
    fig.update_xaxes(tickangle=0, nticks=8)
    return style_fig(fig, 360, legend=True)


def chart_profit_trend(series: pd.DataFrame) -> go.Figure:
    """Clean margin line - no markers, no gridlines, labels alternating above/below."""
    if series is None or series.empty:
        return _empty("No profit history")

    months = series["Month"].tolist()
    values = series["GP %"].astype(float).tolist()
    color = CATEGORICAL[2]

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=months, y=values, name="GP %", mode="lines",
            line=dict(color=color, width=3.5, shape="spline", smoothing=0.35),
            fill="tozeroy", fillcolor=_rgb(color, 0.07),
            hovertemplate="%{x}<br>GP %{y:,.2f}%<extra></extra>",
        )
    )
    # Labels are annotations so they can sit a fixed distance from the line.
    for i, (month, value) in enumerate(zip(months, values)):
        if value <= 0:
            continue
        fig.add_annotation(
            x=month, y=value, text=f"<b>{value:,.1f}%</b>", showarrow=False,
            yshift=26 if i % 2 == 0 else -20,
            font=dict(size=13, color=INK),
        )
    top = max(values) if values else 0
    fig.update_yaxes(range=[-3, max(top * 1.30, 8)], title_text="GP %")
    fig.update_xaxes(tickangle=0, nticks=6)
    return style_fig(fig, 360, show_ygrid=False)


def chart_partner_split(sales: pd.DataFrame, collection: pd.DataFrame | None = None) -> go.Figure:
    """In-house vs customer: sales and collection side by side per business partner."""
    if sales is None or sales.empty:
        return _empty("No sales in this period")

    partners = ["AAFL", "Customer", "AEL"]
    sales_by = (sales.groupby("Partner")["Sales Value"].sum() / CRORE).reindex(partners).fillna(0.0)
    if collection is not None and not collection.empty:
        coll_by = (collection.groupby("Partner")["Amount"].sum() / CRORE).reindex(partners).fillna(0.0)
    else:
        coll_by = pd.Series(0.0, index=partners)

    keep = [p for p in partners if sales_by.get(p, 0) > 0 or coll_by.get(p, 0) > 0]
    if not keep:
        return _empty("No sales in this period")
    sales_vals = [float(sales_by[p]) for p in keep]
    coll_vals = [float(coll_by[p]) for p in keep]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=keep, y=sales_vals, name="Sales",
            marker=dict(color=CATEGORICAL[0], line=dict(width=0)),
            text=[_num(v) for v in sales_vals], textposition="outside", cliponaxis=False,
            textfont=dict(size=12, color=INK),
            hovertemplate="%{x} · Sales %{y:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=keep, y=coll_vals, name="Collection",
            marker=dict(color=CATEGORICAL[1], line=dict(width=0)),
            text=[_num(v) for v in coll_vals], textposition="outside", cliponaxis=False,
            textfont=dict(size=12, color=INK),
            hovertemplate="%{x} · Collection %{y:,.2f} Cr<extra></extra>",
        )
    )
    top = max(sales_vals + coll_vals) if sales_vals or coll_vals else 0
    fig.update_layout(barmode="group")
    fig.update_yaxes(range=[0, top * 1.18 if top > 0 else 1], title_text="BDT Crore")
    fig.update_xaxes(tickangle=0, range=[-0.55, len(keep) - 0.45])
    return style_fig(fig, 360, legend=True, hovermode="x unified")


def chart_cumulative(sales: pd.DataFrame, prior: pd.DataFrame, ctx: FilterContext) -> go.Figure:
    """Month-to-date pace vs the prior month, with the gap shaded."""
    index = pd.date_range(ctx.start, ctx.end, freq="D")

    def cum(frame, idx):
        if frame is None or frame.empty:
            return pd.Series(0.0, index=idx)
        daily = frame.groupby(frame["Date"].dt.normalize())["Sales Value"].sum() / CRORE
        return daily.reindex(idx, fill_value=0.0).cumsum()

    current = cum(sales, index)
    prior_idx = pd.date_range(ctx.prior_start, ctx.prior_end, freq="D")
    previous = cum(prior, prior_idx).reindex(prior_idx, fill_value=0.0)
    previous_points = previous.to_numpy()[: len(index)]
    current_points = current.to_numpy()
    color = CATEGORICAL[0]

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=list(index), y=previous_points, name=f"Prior MTD ({ctx.prior_label})",
            mode="lines", line=dict(color=_rgb(color, 0.45), width=2, dash="dash"),
            hovertemplate="Prior %{y:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=index, y=current_points, name=f"{ctx.month_label} MTD", mode="lines+markers+text",
            line=dict(color=color, width=3, shape="spline", smoothing=0.4),
            marker=dict(size=6, color=color, line=dict(color=TEXT_ON_DARK, width=1.5)),
            fill="tonexty", fillcolor=_rgb(color, 0.10),
            text=_sparse_labels(len(current_points), [float(v) for v in current_points], _num),
            textposition="top center", textfont=dict(size=12, color=INK), cliponaxis=False,
            hovertemplate="Current %{y:,.2f} Cr<extra></extra>",
        )
    )
    top = float(max(current.max() if len(current) else 0, previous.max() if len(previous) else 0))
    fig.update_yaxes(range=[0, top * 1.18 if top > 0 else 1], title_text="BDT Crore")
    pad = pd.Timedelta(hours=8)
    fig.update_xaxes(tickformat="%d %b", nticks=8, range=[index[0] - pad, index[-1] + pad])
    return style_fig(fig, 360, legend=True)


def chart_target_achievement(performance: pd.DataFrame) -> go.Figure:
    """Target bar (light) beside achieved bar (green / amber / red) with ach. % labels."""
    if performance is None or performance.empty or float(performance["Target (Cr)"].sum()) <= 0:
        return _empty("Add product targets in product_targets.csv")

    frame = performance[performance["Target (Cr)"] > 0].copy()
    frame = frame.sort_values("Target (Cr)", ascending=False).head(8)
    labels = frame["Product"].tolist()
    targets = frame["Target (Cr)"].astype(float).tolist()
    achieved = frame["Sales (Cr)"].astype(float).tolist()
    ach_pct = [(a / t * 100) if t else 0.0 for a, t in zip(achieved, targets)]
    achieved_colors = [GOOD if p >= 100 else (WARN if p >= 50 else BAD) for p in ach_pct]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=labels, y=targets, name="Target",
            marker=dict(color="rgba(148,163,184,0.35)", line=dict(width=0)),
            text=[f"{v:,.1f}" for v in targets], textposition="outside", cliponaxis=False,
            textfont=dict(size=12, color=MUTED),
            hovertemplate="Target %{y:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=labels, y=achieved, name="Achieved",
            marker=dict(color=achieved_colors, line=dict(width=0)),
            text=[f"{v:,.1f} · {p:,.0f}%" for v, p in zip(achieved, ach_pct)],
            textposition="outside", cliponaxis=False,
            textfont=dict(size=12, color=INK),
            hovertemplate="Achieved %{y:,.2f} Cr · %{text}<extra></extra>",
        )
    )
    top = float(max(max(targets), max(achieved))) if targets or achieved else 0
    fig.update_layout(barmode="group")
    fig.update_yaxes(range=[0, top * 1.15 if top > 0 else 1], title_text="BDT Crore")
    fig.update_xaxes(tickangle=0, range=[-0.55, len(labels) - 0.45])
    return style_fig(fig, 380, legend=True)


def _ranked_profit_chart(labels, values, gp_values, height=360, axis="BDT Crore"):
    n = len(values)
    colors = _gradient(n)
    for i, gp in enumerate(gp_values):
        if gp < 0:
            colors[i] = BAD
    return _hbar(list(labels), list(values), colors, height, fmt=_num, axis=axis, suffix=" Cr")


def chart_product_profit(performance: pd.DataFrame, top_n: int = 10) -> go.Figure:
    if performance is None or performance.empty:
        return _empty("No sales in this period")
    frame = performance.head(top_n)
    return _ranked_profit_chart(
        frame["Product"].tolist(), frame["Sales (Cr)"].astype(float).tolist(),
        frame["GP (Cr)"].astype(float).tolist(), 360,
    )


def chart_customer_profit(customers: pd.DataFrame, top_n: int = 10) -> go.Figure:
    if customers is None or customers.empty:
        return _empty("No sales in this period")
    frame = customers.head(top_n)
    return _ranked_profit_chart(
        frame["Customer"].tolist(), frame["Sales (Cr)"].astype(float).tolist(),
        frame["Invoice GP (Cr)"].astype(float).tolist(), 360,
    )


def receivables_by_customer(aging: pd.DataFrame, top_n: int = 8) -> pd.DataFrame:
    """Due (0-30) vs overdue (>30) per business partner, minus the excluded list."""
    columns = ["Customer", "Due", "Overdue", "Total"]
    if not isinstance(aging, pd.DataFrame) or aging.empty:
        return pd.DataFrame(columns=columns)

    frame = aging.copy()
    names = frame["BP Name"].astype(str)
    mask = pd.Series(False, index=frame.index)
    for needle in RECEIVABLE_EXCLUDED:
        mask |= names.str.contains(needle, case=False, regex=False, na=False)
    frame = frame[~mask]

    due = frame["0-30 Days"].fillna(0) / CRORE
    overdue = (frame["Total"] - frame["0-30 Days"]).fillna(0) / CRORE
    grouped = pd.DataFrame({"Customer": frame["BP Name"], "Due": due, "Overdue": overdue})
    grouped = grouped.groupby("Customer", as_index=False)[["Due", "Overdue"]].sum()
    grouped["Total"] = grouped["Due"] + grouped["Overdue"]
    grouped = grouped[grouped["Total"] > 0.0001]
    grouped = grouped.sort_values("Total", ascending=False).head(top_n)
    return grouped.reset_index(drop=True)[columns].round(3)


def chart_receivables(aging: pd.DataFrame) -> go.Figure:
    """Top customers: yellow = due 0-30 days, red = overdue above 30 days."""
    frame = receivables_by_customer(aging)
    if frame is None or frame.empty:
        return _empty("No receivables snapshot")

    customers = (
        frame["Customer"].str.replace(r"\(\s+", "(", regex=True)
        .str.replace(r"\s+\)", ")", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .tolist()
    )
    due = frame["Due"].astype(float).tolist()
    overdue = frame["Overdue"].astype(float).tolist()
    totals = frame["Total"].astype(float).tolist()
    wrapped = [textwrap.fill(str(c), width=26, max_lines=2, placeholder="…") for c in customers]
    top = max(totals) if totals else 0
    threshold = top * 0.08

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=wrapped, x=due, orientation="h", name="Due (0–30 days)",
            marker=dict(color=DUE_COLOR, line=dict(width=0)),
            text=[f"{v:,.2f}" if v >= threshold else "" for v in due],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=11, color="#3F2E00"),
            hovertemplate="%{y}<br>Due (0–30): %{x:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            y=wrapped, x=overdue, orientation="h", name="Overdue (> 30 days)",
            marker=dict(color=OVERDUE_COLOR, line=dict(width=0)),
            text=[f"{v:,.2f}" if v >= threshold else "" for v in overdue],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=11, color="#FFFFFF"),
            hovertemplate="%{y}<br>Overdue (>30): %{x:,.2f} Cr<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=[max(v, (top or 1) * 0.03) for v in totals], y=wrapped, mode="text",
            text=[f"{v:,.2f}" for v in totals],
            textposition="middle right", textfont=dict(size=12, color=INK),
            cliponaxis=False, showlegend=False, hoverinfo="skip",
        )
    )
    fig.update_layout(barmode="stack")
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_xaxes(range=[0, top * 1.22 if top > 0 else 1], title_text="BDT Crore")
    return style_fig(fig, 380, legend=True, hovermode="closest")


def chart_inventory(stock: pd.DataFrame, top_n: int = 10) -> go.Figure:
    """One palette colour per product, MT values labelled outside."""
    if stock is None or stock.empty:
        return _empty("No closing stock")

    frame = (
        stock.groupby("Product")
        .agg(MT=("Closing MT", "sum"))
        .sort_values("MT", ascending=False)
        .head(top_n)
    )
    labels = frame.index.tolist()
    colours = [CATEGORICAL[i % len(CATEGORICAL)] for i in range(len(labels))]
    return _hbar(labels, frame["MT"].astype(float).tolist(), colours, 380, fmt=_mt,
                 axis="Metric Tons", suffix=" MT")


# --------------------------------------------------------------------------- #
# Insights (auto-generated analysis)
# --------------------------------------------------------------------------- #


def build_insights(
    dataset: dict[str, Any],
    ctx: FilterContext,
    view: dict[str, Any],
    series: pd.DataFrame,
    performance: pd.DataFrame,
    customers: pd.DataFrame,
) -> list[dict[str, str]]:
    """Rules-based observations that read like an analyst's commentary."""
    out: list[dict[str, str]] = []

    def add(tone: str, icon: str, title: str, text: str) -> None:
        out.append({"tone": tone, "icon": icon, "title": title, "text": text})

    delta = view["sales_delta"]
    if delta is not None:
        direction = "up" if delta >= 0 else "down"
        add(
            "good" if delta >= 0 else "warn",
            "📈" if delta >= 0 else "📉",
            f"Sales {direction} {abs(delta):.1f}% vs prior MTD",
            f"{view['sales_cr']:,.2f} Cr this period against {view['sales_prior_cr']:,.2f} Cr in "
            f"{ctx.prior_label}. {view['so_count']:,} SOs to {view['customer_count']:,} customers, "
            f"{view['volume_mt']:,.0f} MT despatched.",
        )

    if view["sales_target"]:
        ach = view["sales_ach"] or 0
        pacing = ach - view["elapsed_pct"]
        tone = "good" if pacing >= 0 else ("warn" if pacing > -15 else "bad")
        ask = view["asking_month"] or 0
        ask_day = view["asking_day"]
        add(
            tone,
            "🎯",
            f"Target achievement {ach:,.1f}% with {ctx.elapsed_pct:,.0f}% of the month gone",
            f"Target {view['sales_target']:,.2f} Cr. Still to sell {ask:,.2f} Cr over {ctx.days_remaining} "
            f"day(s)" + (f" — {ask_day:,.2f} Cr per day." if ask_day else ".")
            + (" Ahead of pace." if pacing >= 0 else f" Behind pace by {abs(pacing):,.1f} pts."),
        )

    coverage = view["collection_vs_sales"]
    if coverage:
        coverage_pct = coverage * 100
        tone = "good" if coverage_pct >= 100 else ("warn" if coverage_pct >= 70 else "bad")
        target_bit = (
            f" Collection target {view['collection_target']:,.2f} Cr ({view['collection_ach']:,.1f}% achieved)."
            if view["collection_target"] else ""
        )
        add(
            tone, "💰", f"Collection covers {coverage_pct:,.0f}% of sales",
            f"{view['collection_cr']:,.2f} Cr received vs {view['sales_cr']:,.2f} Cr invoiced."
            + target_bit
            + (" Cash is coming in faster than billing." if coverage_pct >= 100 else " Billing is ahead of cash — watch the receivable."),
        )

    add(
        "good" if view["gp_delta_pp"] >= 0 else "warn",
        "📊",
        f"Gross margin {view['gp_pct']:,.2f}% ({view['gp_delta_pp']:+,.2f} pp vs prior)",
        f"GP {view['gp_cr']:,.2f} Cr vs {view['gp_prior_cr']:,.2f} Cr prior MTD, based on the "
        f"{view['gp_basis'].lower()}.",
    )

    if performance is not None and len(performance):
        weak = performance[(performance["GP %"] < 0)]
        if len(weak):
            names = ", ".join(weak["Product"].head(3).tolist())
            add("bad", "⚠️", f"{len(weak)} product(s) sold below cost",
                f"{names} show negative GP in this period — review price or landing cost.")
        top = performance.iloc[0]
        share = top["Sales (Cr)"] / view["sales_cr"] * 100 if view["sales_cr"] else 0
        add("info", "🏷️", f"{top['Product']} is {share:,.1f}% of sales",
            f"{top['Sales (Cr)']:,.2f} Cr at {top['GP %']:,.1f}% GP ({top['Qty (MT)']:,.0f} MT). "
            f"Mix concentration of the top product drives the whole margin.")

    if customers is not None and len(customers):
        top5 = customers.head(5)["Sales (Cr)"].sum()
        share5 = top5 / view["sales_cr"] * 100 if view["sales_cr"] else 0
        tone = "warn" if share5 >= 70 else "info"
        add(tone, "🤝", f"Top 5 customers = {share5:,.1f}% of sales",
            f"{customers.iloc[0]['Customer']} leads with {customers.iloc[0]['Sales (Cr)']:,.2f} Cr "
            f"({customers.iloc[0]['Share %']:,.1f}% share). "
            + ("Concentration is high — a single lost account would hurt." if share5 >= 70 else "Spread is healthy."))

    if view["receivable_cr"]:
        split = view["ar_split"]
        add(
            "warn" if view["overdue_pct"] >= 10 else "good",
            "⏳",
            f"Receivables {view['receivable_cr']:,.2f} Cr, {view['overdue_pct']:,.1f}% overdue",
            f"AAFL {split.get('AAFL', 0):,.2f} Cr · Customer {split.get('Customer', 0):,.2f} Cr · "
            f"AEL {split.get('AEL', 0):,.2f} Cr. Overdue {view['overdue_cr']:,.2f} Cr; "
            f"balance {'up' if view['receivable_change'] >= 0 else 'down'} "
            f"{abs(view['receivable_change']):,.2f} Cr vs the prior snapshot.",
        )

    if view["inventory_cr"]:
        cover = view["stock_cover"]
        add(
            "warn" if cover and cover > 4 else "info",
            "🏬",
            f"Inventory {view['inventory_cr']:,.2f} Cr at cost ({view['inventory_mt']:,.0f} MT)",
            f"{view['sku_count']} SKUs across {view['warehouse_count']} warehouses"
            + (f", about {cover:,.1f} months of cover at the trailing 3-month run-rate." if cover else "."),
        )

    if series is not None and len(series) >= 2:
        previous = series.iloc[-2]
        add(
            "info", "🗓️", "Last month in context",
            f"{previous['Month']}: sales {previous['Sales (Cr)']:,.2f} Cr, collection "
            f"{previous['Collection (Cr)']:,.2f} Cr, GP {previous['GP %']:,.1f}%.",
        )

    last_sale = dataset.get("coverage", {}).get("sales", {}).get("last")
    last_stock = dataset.get("coverage", {}).get("stock", {}).get("last")
    if last_sale is not None and pd.notna(last_sale):
        add(
            "info", "🔄", "Data freshness",
            f"Sales to {last_sale:%d %b %Y}"
            + (f", stock to {last_stock:%d %b %Y}" if last_stock is not None and pd.notna(last_stock) else "")
            + ". Press Refresh data to pull the latest Power Query extract.",
        )
    return out


# --------------------------------------------------------------------------- #
# Exports
# --------------------------------------------------------------------------- #


def sales_export_frame(sales: pd.DataFrame) -> pd.DataFrame:
    columns = ["Date", "SO Number", "Challan No", "Customer", "Group", "Product", "District",
               "Shippoint", "Qty (MT)", "Sales (BDT)", "Cost (BDT)", "GP (BDT)"]
    if sales is None or sales.empty:
        return pd.DataFrame(columns=columns)
    frame = pd.DataFrame(
        {
            "Date": sales["Date"].dt.date,
            "SO Number": sales["SO Number"],
            "Challan No": sales["Challan No"],
            "Customer": sales["Customer"],
            "Group": sales["Partner"],
            "Product": sales["Product"],
            "District": sales["District"],
            "Shippoint": sales["Shippoint"],
            "Qty (MT)": sales["Volume MT"].round(3),
            "Sales (BDT)": sales["Sales Value"].round(2),
            "Cost (BDT)": sales["Cost Value"].round(2),
            "GP (BDT)": (sales["Sales Value"] - sales["Cost Value"]).round(2),
        }
    )
    return frame.sort_values(["Date", "SO Number"]).reset_index(drop=True)


def excel_snapshot(dataset: dict[str, Any], ctx: FilterContext, view: dict[str, Any]) -> bytes:
    from openpyxl.utils import get_column_letter

    performance = product_performance(view["sales"], view["cogs"], dataset.get("product_targets"), ctx.month_key)
    filters = pd.DataFrame(
        {
            "Setting": ["Generated at", "Month", "Period", "Prior MTD", "District", "Warehouse", "Product",
                        "Workbook", "Validation", "Currency", "Volume"],
            "Value": [
                f"{pd.Timestamp.now():%d %b %Y %H:%M}", ctx.month_label, ctx.period_label, ctx.prior_label,
                ctx.district or "All locations", ctx.warehouse or "All warehouses", ctx.product or "All products",
                Path(dataset.get("path", "")).name, overall_status(dataset.get("checks", [])),
                "BDT Crore (1 Cr = 10,000,000) except Sales Detail which is BDT", "Metric Tons",
            ],
        }
    )

    frames: dict[str, pd.DataFrame] = {
        "Summary": pd.DataFrame(
            {
                "KPI": ["Sales (Cr)", "Sales Target (Cr)", "Achievement %", "Collection (Cr)", "Collection Target (Cr)",
                        "GP (Cr)", "GP %", "Receivable (Cr)", "Overdue (Cr)", "Inventory (Cr)", "Inventory (MT)"],
                "Value": [view["sales_cr"], view["sales_target"], view["sales_ach"], view["collection_cr"],
                          view["collection_target"], view["gp_cr"], view["gp_pct"], view["receivable_cr"],
                          view["overdue_cr"], view["inventory_cr"], view["inventory_mt"]],
            }
        ),
        "Filters": filters,
        "Sales Detail": sales_export_frame(view["sales"]),
        "Product Performance": performance,
        "Customers": customer_performance(view["sales"]),
        "Collections": collections_by_party(view["collection"]),
        "Stock": stock_detail(view["stock"]),
        "Aging": aging_detail(view["aging"]),
        "Monthly Trend": monthly_series(dataset, ctx, 12),
    }

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in frames.items():
            sheet = name[:31]
            frame.to_excel(writer, sheet_name=sheet, index=False)
            worksheet = writer.sheets[sheet]
            for index, column in enumerate(frame.columns, start=1):
                worksheet.column_dimensions[get_column_letter(index)].width = min(40, max(12, len(str(column)) + 3))
            worksheet.freeze_panes = "A2"
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
# Formatting, cards and HTML blocks
# --------------------------------------------------------------------------- #


def fmt_cr(value: float | None, digits: int = 2) -> str:
    return "—" if value is None or (isinstance(value, float) and pd.isna(value)) else f"{value:,.{digits}f}"


def fmt_mt(value: float | None, digits: int = 0) -> str:
    return fmt_cr(value, digits)


def fmt_pct(value: float | None, digits: int = 1) -> str:
    return "—" if value is None or (isinstance(value, float) and pd.isna(value)) else f"{value:,.{digits}f}%"


def kpi_card(
    label: str,
    value: str,
    unit: str,
    rows: list[tuple[str, str]],
    accent: str,
    icon: str = "",
    note: str = "",
    delta: tuple[str, bool] | None = None,
) -> str:
    """KPI card: small label, big number, optional delta chip, detail rows."""
    body = "".join(f'<div class="kpi-row"><span>{name}</span><b>{val}</b></div>' for name, val in rows)
    note_html = f'<div class="kpi-note">{note}</div>' if note else ""
    delta_html = ""
    if delta:
        text, good = delta
        clean = text.strip()
        arrow = "▼" if clean.startswith("-") else "▲"
        cls = "up" if good else "down"
        delta_html = f'<span class="kpi-delta {cls}">{arrow} {clean.lstrip("+-")}</span>'
    return (
        f'<div class="kpi" style="--accent:{accent}">'
        f'<div class="kpi-head"><span class="kpi-label">{label}</span><span class="kpi-icon">{icon}</span></div>'
        f'<div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div>'
        f"{delta_html}{note_html}"
        f'<div class="kpi-rows">{body}</div>'
        f"</div>"
    )


def insight_card(insight: dict[str, str]) -> str:
    return (
        f'<div class="insight {insight.get("tone", "info")}">'
        f'<div class="insight-head"><span class="insight-icon">{insight.get("icon", "•")}</span>'
        f'<span class="insight-title">{insight.get("title", "")}</span></div>'
        f'<div class="insight-text">{insight.get("text", "")}</div>'
        f"</div>"
    )


def spacer(px: int = 22) -> None:
    st.markdown(f'<div style="height:{px}px"></div>', unsafe_allow_html=True)


def section(label: str, note: str = "") -> None:
    note_html = f'<p class="section-note">{note}</p>' if note else ""
    st.markdown(f'<div class="section-title">{label}</div>{note_html}', unsafe_allow_html=True)


def status_pill_html(status: str, detail: str) -> str:
    css = {"PASS": "pass", "WARN": "warn", "FAIL": "fail"}.get(status, "warn")
    return f'<div class="status-pill {css}">● <div>{status} · <small>{detail}</small></div></div>'


def brand_header(title: str, subtitle: str, update_note: str, logo_path: Path) -> None:
    """Deprecated single-block header kept for compatibility; use ``render_header``."""
    render_header_html(title, subtitle, update_note, logo_path)


@st.cache_data(show_spinner=False, max_entries=4)
def _logo_data_uri(path_str: str, mtime: float) -> str:
    import base64

    data = Path(path_str).read_bytes()
    return "data:image/png;base64," + base64.b64encode(data).decode("ascii")


def render_header_html(title: str, subtitle: str, meta_html: str, logo_path: Path) -> None:
    uri = _logo_data_uri(str(logo_path), logo_path.stat().st_mtime if logo_path.exists() else 0.0)
    st.markdown(
        f'<div class="nb-header">'
        f'<img class="nb-logo" src="{uri}" alt="Nobayon Traders">'
        f'<div class="nb-title-block"><div class="nb-title">{title}</div><div class="nb-sub">{subtitle}</div></div>'
        f'<div class="nb-meta">{meta_html}</div>'
        f"</div>",
        unsafe_allow_html=True,
    )


def render_header(dataset: dict[str, Any], ctx: FilterContext) -> None:
    """Branded header: logo, report title for the selected month, data-freshness meta."""
    theme = get_theme()
    logo = ROOT / "assets" / ("logo_white.png" if theme == "brand" else "logo_green.png")
    coverage = dataset.get("coverage", {})
    last_sale = coverage.get("sales", {}).get("last")
    last_stock = coverage.get("stock", {}).get("last")
    bits = []
    if last_sale is not None and pd.notna(last_sale):
        bits.append(f"Data up to <b>{last_sale:%d-%b-%y}</b>")
    if last_stock is not None and pd.notna(last_stock) and last_stock != last_sale:
        bits.append(f"Stock to <b>{last_stock:%d-%b-%y}</b>")
    if dataset.get("mtime"):
        bits.append(f"Extract <b>{pd.Timestamp.fromtimestamp(dataset['mtime']):%d-%b-%y %H:%M}</b>")
    render_header_html(
        f"Monthly Business Performance Dashboard : {ctx.month_label}",
        f"{ctx.location_label} · {ctx.product_label} · amounts in BDT Crore",
        "<br>".join(bits),
        logo,
    )


def inject_css(theme: str | None = None) -> None:
    """Theme tokens, card/button/control styling and responsive rules (both themes)."""
    requested = theme or st.session_state.get("theme_mode") or _ACTIVE_THEME
    key = set_theme(requested)
    t = THEME_TOKENS[key]
    dark = key == "brand"

    variables = f"""
  :root {{
    --page: {t['page']};
    --card: {t['card']};
    --sidebar: {t['sidebar']};
    --ink: {t['ink']};
    --muted: {t['muted']};
    --line: {t['line']};
    --table-bg: {t['table_bg']};
    --hover: {t['hover']};
    --accent: {t['accent']};
    --teal: #0F766E;
    --teal-dark: #115E59;
    --amber: #F59E0B;
    --up: {'#4ADE80' if dark else '#15803D'};
    --down: {'#FCA5A5' if dark else '#B91C1C'};
    --soft: {'rgba(148,163,184,0.16)' if dark else 'rgba(15,23,42,0.06)'};
  }}
"""

    base = """
  html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif; }
  .stApp, [data-testid="stAppViewContainer"] { background: var(--page) !important; color: var(--ink); }
  .block-container { padding: 1rem 2rem 1.6rem 2rem; max-width: 1660px; }
  header[data-testid="stHeader"] { background: transparent; }
  [data-testid="stSidebar"], [data-testid="stSidebarContent"] {
    background: var(--sidebar) !important; border-right: 1px solid var(--line);
  }
  [data-testid="stSidebar"] .block-container { padding-top: 0.9rem; }
  [data-testid="stVerticalBlock"] { gap: 0.6rem; }
  h1, h2, h3, h4 { color: var(--ink); margin: 0.3rem 0; letter-spacing: -0.02em; font-weight: 700; }

  .nb-header {
    display: flex; align-items: center; gap: 20px; flex-wrap: wrap;
    padding-bottom: 8px; border-bottom: 1px solid var(--line); margin-bottom: 2px;
  }
  .nb-logo { width: 225px; max-width: 34%; height: auto; }
  .nb-title-block { flex: 1 1 400px; min-width: 240px; }
  .nb-title { font-size: 20px; font-weight: 800; color: var(--ink); letter-spacing: -0.02em; line-height: 1.25; }
  .nb-sub { font-size: 11.5px; color: var(--muted); margin-top: 2px; }
  .nb-meta { margin-left: auto; text-align: right; font-size: 11px; color: var(--muted); line-height: 1.6; }

  /* Chart cards - applied through st.container(border=True, key="card-...") */
  [data-testid="stHorizontalBlock"] { align-items: stretch; }
  [class*="st-key-card-"] {
    height: 100%;
    background: var(--card) !important;
    border: 1px solid var(--line) !important;
    border-radius: 14px !important;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06), 0 10px 26px rgba(15, 23, 42, 0.06);
    padding: 12px !important;
  }

  .kpi-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; }
  .kpi {
    background: var(--card); border: 1px solid var(--line); border-left: 4px solid var(--accent);
    border-radius: 12px; padding: 11px 13px 9px 13px;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.10); min-width: 0;
  }
  .kpi-head { display: flex; justify-content: space-between; align-items: center; }
  .kpi-label { font-size: 10.5px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: var(--muted); }
  .kpi-icon { font-size: 13px; opacity: .85; }
  .kpi-value { font-size: 28px; font-weight: 800; color: var(--ink); line-height: 1.15; margin-top: 2px; letter-spacing: -0.02em; }
  .kpi-unit { font-size: 12px; font-weight: 600; color: var(--muted); margin-left: 4px; }
  .kpi-delta { display: inline-block; font-size: 11.5px; font-weight: 700; margin-top: 1px; }
  .kpi-delta.up { color: var(--up); }
  .kpi-delta.down { color: var(--down); }
  .kpi-note { font-size: 11px; color: var(--muted); margin-top: 1px; }
  .kpi-rows { margin-top: 6px; border-top: 1px dashed var(--line); padding-top: 5px; }
  .kpi-row { display: flex; justify-content: space-between; gap: 6px; font-size: 11.5px; padding: 1.5px 0; }
  .kpi-row span { color: var(--muted); }
  .kpi-row b { color: var(--ink); font-weight: 700; white-space: nowrap; }

  .insight {
    background: var(--card); border: 1px solid var(--line); border-left: 4px solid var(--accent);
    border-radius: 12px; padding: 12px 14px; height: 100%;
  }
  .insight.good { border-left-color: #16A34A; }
  .insight.warn { border-left-color: #F59E0B; }
  .insight.bad { border-left-color: #DC2626; }
  .insight.info { border-left-color: var(--teal); }
  .insight-head { display: flex; gap: 8px; align-items: baseline; }
  .insight-icon { font-size: 14px; }
  .insight-title { font-size: 13px; font-weight: 700; color: var(--ink); }
  .insight-text { font-size: 12px; color: var(--muted); margin-top: 3px; line-height: 1.5; }
  .insight-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(330px, 1fr)); gap: 12px; }

  .section-title { font-size: 13px; font-weight: 700; color: var(--ink); margin: 0 0 1px 0; }
  .section-note { font-size: 11.5px; color: var(--muted); margin: 0 0 2px 0; }
  .metric-note { font-size: 12px; color: var(--muted); font-weight: 600; }
  .metric-note-value { font-size: 22px; font-weight: 700; color: var(--ink); }

  /* Buttons */
  .stButton button[kind="primary"] {
    background: var(--teal) !important; color: #FFFFFF !important; border: 1px solid var(--teal) !important;
    border-radius: 10px !important; font-weight: 600 !important;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.18); transition: background .2s ease, border-color .2s ease;
  }
  .stButton button[kind="primary"]:hover { background: var(--teal-dark) !important; border-color: var(--teal-dark) !important; }
  .stButton button[kind="secondary"], .stButton button:not([kind]) {
    background: transparent !important; color: var(--teal) !important; border: 1px solid var(--teal) !important;
    border-radius: 10px !important; font-weight: 600 !important; transition: background .2s ease;
  }
  .stButton button[kind="secondary"]:hover, .stButton button:not([kind]):hover {
    background: rgba(15, 118, 110, 0.12) !important; color: var(--teal-dark) !important;
  }
  .stDownloadButton button {
    background: var(--amber) !important; color: #1F2937 !important; border: 1px solid var(--amber) !important;
    border-radius: 10px !important; font-weight: 600 !important;
  }
  .stDownloadButton button:hover { background: #D97706 !important; border-color: #D97706 !important; color: #111827 !important; }

  /* Tabs - pill style */
  .stTabs [data-baseweb="tab-list"] { gap: 6px; border-bottom: none; }
  .stTabs [data-baseweb="tab-highlight"] { display: none; }
  .stTabs [data-testid="stTab"], .stTabs [data-baseweb="tab"] {
    background: var(--soft); border-radius: 999px !important; padding: 4px 12px;
    color: var(--muted); font-size: 12.5px; font-weight: 600;
  }
  .stTabs [data-testid="stTab"][aria-selected="true"], .stTabs [aria-selected="true"] {
    background: var(--teal) !important; color: #FFFFFF !important;
  }
  .stTabs [data-testid="stTab"][aria-selected="true"] *, .stTabs [aria-selected="true"] * { color: #FFFFFF !important; }

  /* Radio - pill style */
  [data-testid="stRadio"] [role="radiogroup"] { gap: 6px; }
  [data-testid="stRadio"] label {
    background: var(--soft); border: 1px solid transparent; border-radius: 999px;
    padding: 3px 10px 3px 6px; transition: background .2s ease;
  }
  [data-testid="stRadio"] label:has(input:checked) { background: var(--teal); border-color: var(--teal); }
  [data-testid="stRadio"] label:has(input:checked) * { color: #FFFFFF !important; }

  /* Inputs */
  [data-baseweb="select"] > div, [data-baseweb="input"], [data-baseweb="base-input"], textarea {
    background: var(--card) !important; border-radius: 10px !important;
    border-color: var(--line) !important; color: var(--ink) !important;
  }
  [data-baseweb="select"] > div:focus-within, [data-baseweb="input"]:focus-within {
    border-color: var(--teal) !important; box-shadow: 0 0 0 1px var(--teal) !important;
  }
  [data-baseweb="checkbox"] div[role="checkbox"] { border-radius: 999px; }
  [data-baseweb="popover"] [role="listbox"], [data-baseweb="menu"] { background: var(--card) !important; }
  [role="option"] { color: var(--ink) !important; }
  [role="option"]:hover, [aria-selected="true"][role="option"] { background: var(--hover) !important; }

  [data-testid="stDataFrame"] { background: var(--table-bg); border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
  .status-pill { display: flex; align-items: center; gap: 8px; border-radius: 10px; padding: 8px 10px; font-size: 11.5px; font-weight: 600; }
  .status-pill small { font-weight: 500; opacity: .85; }
  footer { visibility: hidden; }
"""

    theme_css = (
        """
  [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p, [data-testid="stWidgetLabel"] label,
  [data-testid="stCaptionContainer"], [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
  [data-testid="stSidebarNav"] a, [data-testid="stExpander"] summary { color: var(--ink) !important; }
  [data-testid="stCaptionContainer"], [data-testid="stMetricLabel"] { color: var(--muted) !important; }
  [data-testid="stExpander"] { background: var(--card); border: 1px solid var(--line); border-radius: 10px; }
  [data-testid="stSidebarNav"] a:hover, [data-testid="stSidebarNav"] a[aria-current="page"] { background: var(--hover); }
  [data-testid="stSidebar"] hr { margin: 0.5rem 0; border-color: var(--line); }
  .status-pill.pass { background: rgba(74,222,128,.12); color: #86EFAC; border: 1px solid rgba(74,222,128,.35); }
  .status-pill.warn { background: rgba(245,158,11,.14); color: #FCD34D; border: 1px solid rgba(245,158,11,.38); }
  .status-pill.fail { background: rgba(220,38,38,.16); color: #FCA5A5; border: 1px solid rgba(220,38,38,.40); }
"""
        if dark
        else """
  .status-pill.pass { background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; }
  .status-pill.warn { background: #FFFBEB; color: #B45309; border: 1px solid #FDE68A; }
  .status-pill.fail { background: #FEF2F2; color: #B91C1C; border: 1px solid #FECACA; }
  [data-testid="stExpander"] { border: 1px solid var(--line); border-radius: 10px; }
  [data-testid="stSidebar"] hr { margin: 0.5rem 0; border-color: var(--line); }
"""
    )

    responsive = """
  @media (max-width: 1180px) { .kpi-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
  @media (max-width: 1200px) { .block-container { padding-left: 1rem; padding-right: 1rem; } }
  @media (max-width: 900px) {
    [data-testid="stHorizontalBlock"] { flex-wrap: wrap; align-items: stretch; }
    [data-testid="stHorizontalBlock"]:has([class*="st-key-card-"]) [data-testid="stColumn"] {
      flex: 1 1 100% !important; width: 100% !important; min-width: 100% !important;
    }
    [class*="st-key-card-"] { height: auto !important; }
    [data-testid="stPlotlyChart"], .js-plotly-plot { max-width: 100% !important; }
  }
  @media (max-width: 700px) {
    .block-container { padding: 0.6rem 0.5rem 1.2rem 0.5rem; }
    .nb-header { gap: 8px; }
    .nb-logo { width: 150px; max-width: 55%; }
    .nb-title { font-size: 15px; }
    .nb-sub { font-size: 10.5px; }
    .nb-meta { text-align: left; margin-left: 0; flex-basis: 100%; font-size: 10px; }
    .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
    .kpi { padding: 9px 10px; border-radius: 10px; }
    .kpi-value { font-size: 20px; }
    .kpi-label { font-size: 9px; }
    .kpi-row { font-size: 10px; }
    .insight-title { font-size: 12px; }
    .insight-text { font-size: 11px; }
    .section-title { font-size: 12px; }
    [data-testid="stHorizontalBlock"] [data-testid="stColumn"] { min-width: 100% !important; }
    .stTabs [data-baseweb="tab"] { font-size: 11.5px; padding: 3px 9px; }
  }
"""

    st.markdown(f"<style>{variables}{base}{theme_css}{responsive}</style>", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# State, downloads and small helpers
# --------------------------------------------------------------------------- #


def csv_bytes(frame: pd.DataFrame) -> bytes:
    return frame.to_csv(index=False).encode("utf-8-sig")


def download_csv_button(frame: pd.DataFrame, file_name: str, key: str, label: str = "⬇ Download CSV") -> None:
    st.download_button(label, data=csv_bytes(frame), file_name=file_name, mime="text/csv", key=key, width="content")


def table_height(frame: pd.DataFrame, max_height: int = 520, minimum: int = 150) -> int:
    rows = len(frame) if isinstance(frame, pd.DataFrame) else 0
    return int(min(max_height, max(minimum, 38 + 35 * rows)))


def get_state() -> tuple[dict[str, Any], FilterContext]:
    dataset = st.session_state.get("dataset")
    ctx = st.session_state.get("ctx")
    if dataset is None or ctx is None:
        try:
            st.switch_page("streamlit_app.py")
        except Exception:
            st.warning("Opening the dashboard…")
            st.stop()
        st.stop()
    return dataset, ctx


def get_checks() -> list[dict[str, str]]:
    return (st.session_state.get("dataset") or {}).get("checks", [])


def month_bounds(year: int, month: int, today: Any = None) -> tuple[pd.Timestamp, pd.Timestamp]:
    today_ts = pd.Timestamp.today().normalize() if today is None else pd.Timestamp(today).normalize()
    start = pd.Timestamp(year=int(year), month=int(month), day=1)
    end = start + pd.offsets.MonthEnd(0)
    if start <= today_ts <= end:
        end = today_ts
    return start, end


def workbook_age_minutes(mtime: float) -> float | None:
    if not mtime:
        return None
    return max(0.0, (pd.Timestamp.now() - pd.Timestamp.fromtimestamp(mtime)).total_seconds() / 60.0)

