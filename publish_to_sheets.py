#!/usr/bin/env python
"""Publish the refreshed workbook to Google Sheets (the live data source for Streamlit Cloud).

One-time setup (see README, "Option B"):
  1. Google Cloud console -> new project -> enable "Google Sheets API" and "Google Drive API".
  2. Create a service account -> Keys -> Add key (JSON) -> download it.
  3. Create the Google Sheet and share it with the service-account email as Editor.
  4. Set the id + key path:
         set GOOGLE_SHEET_ID=1AbC...
         set SHEETS_CREDENTIALS=C:\\path\\to\\service-account.json

Usage:
    python publish_to_sheets.py --dry-run          # show what would be written, no credentials
    python publish_to_sheets.py                    # push all five tabs
    python publish_to_sheets.py --workbook "..." --sheet-id "..." --credentials "...json"
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.data_pipeline import CORE_SHEETS, dataset_path  # noqa: E402

TABS = list(CORE_SHEETS)


def log(message: str = "") -> None:
    print(message, flush=True)


def to_values(frame: pd.DataFrame) -> list[list[object]]:
    """Header + rows, dates as ISO text so Google never re-formats them."""
    frame = frame.copy()
    for column in frame.columns:
        series = frame[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            frame[column] = series.dt.strftime("%Y-%m-%d")
    frame = frame.where(pd.notna(frame), "")
    rows: list[list[object]] = [list(frame.columns)]
    rows.extend(frame.astype(object).values.tolist())
    return rows


def read_sheets(workbook: Path) -> dict[str, list[list[object]]]:
    book = pd.ExcelFile(workbook)
    out: dict[str, list[list[object]]] = {}
    for tab in TABS:
        if tab not in book.sheet_names:
            log(f"  [WARN] {tab:<18} missing in the workbook - skipped")
            continue
        frame = book.parse(tab)
        out[tab] = to_values(frame)
        log(f"  {tab:<18} {len(frame):>7,} rows x {len(frame.columns):>2} cols")
    return out


def push(sheet_id: str, credentials: Path, payload: dict[str, list[list[object]]]) -> None:
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        log("  [FAIL] gspread / google-auth not installed:  pip install -r requirements-sync.txt")
        raise SystemExit(1)

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(str(credentials), scopes=scopes)
    client = gspread.authorize(creds)
    book = client.open_by_key(sheet_id)
    existing = {ws.title: ws for ws in book.worksheets()}

    for tab, values in payload.items():
        rows, cols = len(values), max(len(r) for r in values)
        if tab in existing:
            ws = existing[tab]
            ws.clear()
            if ws.row_count < rows or ws.col_count < cols:
                ws.resize(rows=rows + 20, cols=cols + 2)
        else:
            ws = book.add_worksheet(title=tab, rows=rows + 20, cols=cols + 2)
        try:
            ws.update(values=values, range_name="A1", value_input_option="RAW")
        except TypeError:  # older gspread signature
            ws.update(range_name="A1", values=values, value_input_option="RAW")
        log(f"  {tab:<18} written ({rows:,} rows)")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Publish the workbook to Google Sheets")
    parser.add_argument("--workbook", help="path to dashboard database.xlsx")
    parser.add_argument("--sheet-id", help="Google Sheet id (or GOOGLE_SHEET_ID)")
    parser.add_argument("--credentials", help="service-account JSON (or SHEETS_CREDENTIALS)")
    parser.add_argument("--dry-run", action="store_true", help="validate and summarise, write nothing")
    args = parser.parse_args(argv)

    workbook = Path(args.workbook) if args.workbook else dataset_path()
    log("=" * 70)
    log(f"Google Sheets publish · {workbook.name}")
    log("=" * 70)
    if not workbook.exists():
        log(f"[FAIL] workbook not found: {workbook}")
        return 1

    log("")
    log("Reading tabs")
    payload = read_sheets(workbook)
    if not payload:
        log("[FAIL] nothing to publish")
        return 1

    if args.dry_run:
        log("")
        log("[dry-run] nothing written. Remove --dry-run to publish.")
        return 0

    sheet_id = args.sheet_id or os.environ.get("GOOGLE_SHEET_ID", "")
    creds = args.credentials or os.environ.get("SHEETS_CREDENTIALS", "")
    if not sheet_id or not creds:
        log("")
        log("[FAIL] set GOOGLE_SHEET_ID and SHEETS_CREDENTIALS (or pass --sheet-id / --credentials)")
        return 1

    log("")
    log("Writing to Google Sheets")
    push(sheet_id, Path(creds), payload)
    log("")
    log("Done. The Streamlit app reads this sheet when GOOGLE_SHEET_ID is set.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
