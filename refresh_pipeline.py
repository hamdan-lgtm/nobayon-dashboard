#!/usr/bin/env python
"""Refresh, validate and optionally publish the Nobayon dashboard data.

    python refresh_pipeline.py                 # Power Query refresh + validate + git push
    python refresh_pipeline.py --no-git        # refresh + validate only (app button / CI)
    python refresh_pipeline.py --no-excel      # validate an already refreshed workbook
    python refresh_pipeline.py --export-csv    # also write sheets_export/*.csv (Google Sheets route)

Exit code 0 = valid, 1 = validation failed.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.data_pipeline import (  # noqa: E402
    CORE_SHEETS,
    check_summary,
    dataset_path,
    load_dataset,
    overall_status,
    validate_pipeline,
)

REQUIRED_COLUMNS: dict[str, tuple[str, ...]] = {
    "SO_DO_Database": ("Date", "SO Number", "Product Name", "Customer", "Quantity", "Sales Value", "Cost Value"),
    "Collection": ("Date", "Party Name", "Collection Amount"),
    "COGS_Database": ("Date", "Item Name", "Avg. COGS", "Amount"),
    "Aging_Database": ("As_At", "BP_Name", "Total"),
    "Stock_Database": ("Date", "Item Name", "Closing Qty", "Warehouse"),
}

SYMBOL = {"PASS": "[ OK ]", "WARN": "[WARN]", "FAIL": "[FAIL]"}


def log(message: str = "") -> None:
    print(message, flush=True)


def resolve_workbook(explicit: str | None = None) -> Path:
    return Path(explicit).expanduser().resolve() if explicit else dataset_path().resolve()


# --------------------------------------------------------------------------- #
# 1. Power Query refresh (Excel COM)
# --------------------------------------------------------------------------- #


def refresh_excel(workbook: Path, attempts: int = 3) -> str:
    """Refresh every Power Query connection and save the workbook.

    If the workbook is already open in Excel the open copy is refreshed in place;
    otherwise a private hidden Excel instance opens, refreshes, saves and quits
    (opening a file inside a busy instance can silently return None).
    """
    if os.name != "nt":
        return "skipped (Excel COM is Windows-only)"
    try:
        import win32com.client as win32  # type: ignore
    except Exception:
        return "skipped (pywin32 not installed)"

    excel = None
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        excel = None

    if excel is not None:
        book = None
        try:
            for candidate in excel.Workbooks:
                try:
                    if Path(candidate.FullName).resolve() == workbook:
                        book = candidate
                        break
                except Exception:
                    continue
        except Exception:
            book = None
        if book is not None:
            started = time.time()
            try:
                excel.DisplayAlerts = False
                book.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()
                book.Save()
                return f"refreshed and saved in {time.time() - started:,.0f}s (open Excel window)"
            except Exception as exc:
                return f"skipped while open in Excel ({type(exc).__name__}: {exc})"

    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        started = time.time()
        xl = None
        try:
            xl = win32.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AskToUpdateLinks = False
            book = xl.Workbooks.Open(str(workbook), UpdateLinks=0)
            if book is None:
                raise RuntimeError("Excel did not return a workbook object")
            if getattr(book, "ReadOnly", False):
                raise RuntimeError("workbook opened read-only (close it in Excel first)")
            book.RefreshAll()
            xl.CalculateUntilAsyncQueriesDone()
            book.Save()
            return f"refreshed and saved in {time.time() - started:,.0f}s"
        except Exception as exc:
            last_error = exc
            if attempt < attempts:
                time.sleep(3)
        finally:
            if xl is not None:
                try:
                    xl.Quit()
                except Exception:
                    pass
    return f"skipped ({type(last_error).__name__}: {last_error})"


# --------------------------------------------------------------------------- #
# 2. Validation
# --------------------------------------------------------------------------- #


def validate_workbook(workbook: Path) -> bool:
    log("")
    log(f"Workbook : {workbook}")
    if not workbook.exists():
        log("  [FAIL] file not found")
        return False
    log(f"  size    : {workbook.stat().st_size / 1024 / 1024:,.1f} MB")
    log(f"  updated : {datetime.fromtimestamp(workbook.stat().st_mtime):%d %b %Y %H:%M}")

    try:
        book = pd.ExcelFile(workbook)
    except Exception as exc:
        log(f"  [FAIL] cannot open workbook: {type(exc).__name__}: {exc}")
        return False

    log("")
    log("Sheets")
    failures = 0
    for sheet in CORE_SHEETS:
        if sheet not in book.sheet_names:
            log(f"  [FAIL] {sheet:<18} missing")
            failures += 1
            continue
        frame = book.parse(sheet)
        missing = [c for c in REQUIRED_COLUMNS[sheet] if c not in frame.columns]
        if frame.empty:
            log(f"  [FAIL] {sheet:<18} empty")
            failures += 1
        elif missing:
            log(f"  [WARN] {sheet:<18} {len(frame):>7,} rows · missing: {', '.join(missing)}")
        else:
            log(f"  [ OK ] {sheet:<18} {len(frame):>7,} rows · {len(frame.columns):>2} columns")
    if failures:
        log("")
        log("Validation aborted: core sheets missing or empty.")
        return False

    dataset = load_dataset(str(workbook), workbook.stat().st_mtime)
    checks = dataset.get("checks") or validate_pipeline(dataset)

    log("")
    log("Data checks")
    for check in checks:
        log(f"  {SYMBOL.get(check['Status'], '[????]')} {check['Check']:<32} {check['Detail']}")

    summary = check_summary(checks)
    status = overall_status(checks)
    log("")
    log(f"Result: {status} · {summary.get('PASS', 0)} pass, {summary.get('WARN', 0)} alert, {summary.get('FAIL', 0)} fail")

    log("")
    log("Coverage")
    for key in ("sales", "cogs", "collection", "aging", "stock"):
        meta = dataset.get("coverage", {}).get(key, {})
        first, last = meta.get("first"), meta.get("last")
        f_txt = f"{first:%d %b %Y}" if first is not None and pd.notna(first) else "-"
        l_txt = f"{last:%d %b %Y}" if last is not None and pd.notna(last) else "-"
        log(f"  {key:<12} {meta.get('rows', 0):>7,} rows · {f_txt} → {l_txt}")
    return status != "FAIL"


# --------------------------------------------------------------------------- #
# 3. Google Sheets route (optional)
# --------------------------------------------------------------------------- #


def export_csv(workbook: Path) -> None:
    """Write one CSV per sheet; upload these to Google Sheets for the web fallback."""
    out = ROOT / "sheets_export"
    out.mkdir(exist_ok=True)
    book = pd.ExcelFile(workbook)
    for sheet in CORE_SHEETS:
        if sheet not in book.sheet_names:
            continue
        frame = book.parse(sheet)
        path = out / f"{sheet}.csv"
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        log(f"  CSV {path.name:<26} {len(frame):>7,} rows")
    log(f"  → upload these to a Google Sheet and set GOOGLE_SHEET_ID to read from it")


# --------------------------------------------------------------------------- #
# 4. Git
# --------------------------------------------------------------------------- #


def sync_data_copy(workbook: Path) -> Path:
    """Copy the refreshed workbook into ./data so it can be committed for Streamlit Cloud."""
    target = ROOT / "data" / workbook.name
    target.parent.mkdir(exist_ok=True)
    shutil.copy2(workbook, target)
    log(f"  copied to {target.relative_to(ROOT)}")
    return target


def git_publish(paths: list[Path]) -> None:
    log("")
    log("Git")
    if shutil.which("git") is None:
        log("  [WARN] git not found - skipped")
        return
    try:
        inside = subprocess.run(["git", "rev-parse", "--is-inside-work-tree"], cwd=ROOT, capture_output=True, text=True)
        if inside.returncode != 0 or inside.stdout.strip() != "true":
            log("  [WARN] not a git repository - skipped")
            return
        staged_any = False
        for path in paths:
            try:
                relative = path.relative_to(ROOT)
            except ValueError:
                continue
            subprocess.run(["git", "add", str(relative)], cwd=ROOT, check=True)
            staged_any = True
        if not staged_any:
            log("  [WARN] nothing inside the repo to commit - use --sync-data")
            return
        if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=ROOT).returncode == 0:
            log("  no workbook changes to commit")
            return
        message = f"Refresh dashboard database ({datetime.now():%Y-%m-%d %H:%M})"
        subprocess.run(["git", "commit", "-m", message], cwd=ROOT, check=True)
        push = subprocess.run(["git", "push"], cwd=ROOT, capture_output=True, text=True)
        log(f"  {'pushed' if push.returncode == 0 else 'push failed: ' + (push.stderr or '')[:200]}")
    except subprocess.CalledProcessError as exc:
        log(f"  [WARN] git failed: {exc}")


# --------------------------------------------------------------------------- #
# Entrypoint
# --------------------------------------------------------------------------- #


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Refresh and validate the Nobayon dashboard data")
    parser.add_argument("--no-git", action="store_true")
    parser.add_argument("--no-excel", action="store_true")
    parser.add_argument("--export-csv", action="store_true", help="also write sheets_export/*.csv")
    parser.add_argument("--sync-data", action="store_true",
                        help="copy the refreshed workbook into data/ so git can commit it")
    parser.add_argument("--publish-sheets", action="store_true",
                        help="push the refreshed workbook to Google Sheets (needs GOOGLE_SHEET_ID + SHEETS_CREDENTIALS)")
    parser.add_argument("--workbook", help="path to the workbook")
    args = parser.parse_args(argv)

    workbook = resolve_workbook(args.workbook)
    log("=" * 78)
    log(f"Nobayon Traders · dashboard refresh {datetime.now():%d %b %Y %H:%M}")
    log("=" * 78)

    if args.no_excel:
        refresh_status = "skipped (--no-excel)"
        log("")
        log("Refresh: skipped (--no-excel)")
    else:
        log("")
        log("Refresh")
        refresh_status = refresh_excel(workbook)
        log(f"  {refresh_status}")
    log(f"REFRESH: {refresh_status}")

    ok = validate_workbook(workbook)

    if args.publish_sheets and ok:
        log("")
        log("Google Sheets")
        try:
            from publish_to_sheets import main as publish_main
            code = publish_main([])
            log(f"  publish exit: {code}")
        except Exception as exc:
            log(f"  [WARN] publish failed: {type(exc).__name__}: {exc}")

    if args.export_csv and ok:
        log("")
        log("CSV export")
        export_csv(workbook)

    publish_paths: list[Path] = []
    if args.sync_data and ok:
        log("")
        log("Sync to repo")
        publish_paths.append(sync_data_copy(workbook))

    if not args.no_git and ok:
        git_publish(publish_paths)

    log("")
    log("Done." if ok else "Done with validation failures - fix the data before publishing.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
