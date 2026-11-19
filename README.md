# Nobayon Traders — Monthly Business Performance Dashboard

Streamlit dashboard for the trading business, built on the Power Query extract
`dashboard database.xlsx`. Deep-green brand theme, the Nobayon logo in the header, monthly
performance against targets, an analysis/insights page and a full audit trail.

- **Performance Overview** — 5 KPI cards (sales, collection, gross profit, receivables, inventory) in a
  symmetrical 2 × 4 card grid: sales & collection trend, profit-% trend, in-house vs customer,
  sales pace, product/customer profitability, receivables (top customers, due vs overdue) and inventory.
- **Analysis & Insights** — auto-generated commentary from the current month, month-on-month table,
  product and customer contribution.
- **Targets** — product-wise target vs achieved (moved off the overview), grouped target table and the
  product detail grid.
- **Detail & Audit** — transaction and summary tables with CSV downloads plus the validation log.

---

## Folder map

```text
Nobayon_Dashboard/                     <- repository root (push this folder)
├── streamlit_app.py                   entrypoint: sidebar, refresh, theme, navigation
├── pages/
│   ├── 1_overview.py                  KPI band + 8 chart cards (2 x 4 grid)
│   ├── 2_analysis.py                  insights, mix and concentration
│   ├── 3_detail.py                    tables, CSV exports, audit log
│   └── 4_targets.py                   product-wise target vs achieved
├── utils/
│   ├── __init__.py
│   └── data_pipeline.py               ETL, KPIs, charts, insights, exports, theme
├── assets/                            logo (white / green), favicon
├── data/                              workbook copy used by Streamlit Cloud
│   └── .gitkeep                       (put "dashboard database.xlsx" here for the cloud build)
├── .streamlit/
│   ├── config.toml                    light base theme, teal accent
│   └── secrets.toml.example           template for GOOGLE_SHEET_ID
├── .github/workflows/validate_data.yml
├── targets.csv · product_targets.csv  editable targets
├── refresh_pipeline.py                Power Query refresh -> validate -> sync -> git push
├── run_dashboard.bat · Dashboard.lnk · Dashboard (open only).lnk
├── requirements.txt
└── README.md
```

---

## 1. Quick start

Double-click **`Dashboard.lnk`** — refreshes Power Query, validates, opens the app.
**`Dashboard (open only).lnk`** opens it with the data already on disk.

From a terminal:

```bat
run_dashboard.bat                 :: refresh + validate + open
run_dashboard.bat /norefresh      :: open only
run_dashboard.bat /port 8510      :: force a port
```

Inside the app the sidebar gives you: **🌿 Brand / ☀️ Light**, **🔄 Refresh data** (runs Power Query),
**⏱ Auto-reload**, **Period / District / Warehouse / Product**, **⬇ Download**, and the validation pill.

---

## 2. Where the data comes from

The app looks for the workbook in this order:

1. `MASTER_EXCEL_PATH` environment variable (or `GOOGLE_SHEET_ID`, see below)
2. `..\Business_dashboard\dashboard database.xlsx` ← **the live Power Query file on your PC**
3. `Nobayon_Dashboard\dashboard database.xlsx`
4. `Nobayon_Dashboard\data\dashboard database.xlsx` ← **the copy that lives in the GitHub repo**

On your PC it therefore always reads the live file; on Streamlit Cloud (where `..\Business_dashboard`
does not exist) it reads the copy in `data/`. Keep that copy current with:

```bat
python refresh_pipeline.py --sync-data
```

which refreshes the live workbook, copies it into `data/`, and (without `--no-git`) commits and
pushes it so the cloud app redeploys with the new data.

### Live data source — Google Sheets (recommended for the deployment)

A Google Sheet already exists and is shared read-only for this project:

**`1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY`** — “Nobayon Dashboard Data”
<https://docs.google.com/spreadsheets/d/1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY/edit>

It holds the five tabs (`SO_DO_Database`, `Collection`, `COGS_Database`, `Aging_Database`,
`Stock_Database`) and the app reads it through Google's CSV endpoint — no credentials on the cloud.

1. **Point the app at it** — locally:

   ```bat
   set GOOGLE_SHEET_ID=1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY
   ```

   On Streamlit Cloud → **Settings → Secrets**:

   ```toml
   GOOGLE_SHEET_ID = "1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY"
   ```

2. **Keep it fresh — option A: Apps Script (no keys, recommended)**

   Google converts the workbook itself, on a timer, inside your account:

   1. Open the live sheet → **Extensions → Apps Script**.
   2. Paste `apps_script/refresh_from_workbook.gs` over `Code.gs`.
   3. Left panel → **Services +** → **Drive API** → **Add**.
   4. Run **`refreshFromWorkbook`** once → accept the permissions.
   5. **Triggers** (clock icon) → **Add Trigger** → function `refreshFromWorkbook`, *Time-driven*,
      *Day timer*, 8am–9am. Add two more triggers for 1pm and 6pm if you want three runs a day.

   Nothing else to install: Power Query saves the workbook → Google Drive syncs it → the script
   converts it and copies the five tabs into the live sheet → the app picks it up.

3. **Keep it fresh — option B: service account (if you prefer running it from the PC)**

   1. <https://console.cloud.google.com> → new project → **APIs & Services → Library** → enable
      **Google Sheets API**.
   2. **Credentials → Create credentials → Service account** → create → open it → **Keys → Add key →
      Create new key → JSON** → save the file (e.g. `C:\secrets\nobayon-sheets.json`).
   3. Copy the service-account email (`…@….iam.gserviceaccount.com`).
   4. Open the sheet → **Share** → paste that email → **Editor** → Send.
   5. On the PC:

      ```bat
      pip install -r requirements-sync.txt
      set GOOGLE_SHEET_ID=1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY
      set SHEETS_CREDENTIALS=C:\secrets\nobayon-sheets.json
      python publish_to_sheets.py --dry-run
      python publish_to_sheets.py
      python refresh_pipeline.py --no-git --publish-sheets     :: refresh + validate + publish
      ```

4. **Result** — the cloud app shows new data within a minute of the refresh, with no redeploy and
   without growing the repository.

---

## 3. Refresh & daily schedule

**In-app** — `🔄 Refresh data` runs `refresh_pipeline.py`, which opens the workbook in Excel
(attaching to your open window when it is already open, otherwise a private hidden instance), runs
**Data → Refresh All**, waits for the async queries, saves the workbook, re-reads it and reloads the
dashboard. Every KPI, chart and the “Data up to” date update from the new extract. Expect 1–3 minutes.

**Daily automation** — Windows Task Scheduler, three runs a day:

```bat
schtasks /Create /TN "Nobayon Dashboard 08" /TR "\"C:\Python312\python.exe\" \"G:\...\Nobayon_Dashboard\refresh_pipeline.py\"" /SC DAILY /ST 08:00 /F
schtasks /Create /TN "Nobayon Dashboard 13" /TR "\"C:\Python312\python.exe\" \"G:\...\Nobayon_Dashboard\refresh_pipeline.py\"" /SC DAILY /ST 13:00 /F
schtasks /Create /TN "Nobayon Dashboard 18" /TR "\"C:\Python312\python.exe\" \"G:\...\Nobayon_Dashboard\refresh_pipeline.py\"" /SC DAILY /ST 18:00 /F
```

Tick *Run whether user is logged on or not* so Excel COM can attach. In the app, **⏱ Auto-reload
(60 s)** makes open dashboards pick up the refreshed file on their own.

> The workbook already contains the Power Query connections; the pipeline triggers them. To add a
> new query, edit it in Excel (**Data → Get Data**), then the same button/schedule refreshes it.

---

## 4. Targets

`targets.csv` — month totals:

```csv
Month,Sales_Target_Cr,Collection_Target_Cr
2026-11,51.58,45.10
```

`product_targets.csv` — product groups (`|` combines products into one target row):

```csv
Month,Product,Target_Cr
2026-11,SBM|Soybean (Local),12.50
2026-11,Jute,18.98
```

Both are plain CSVs in this folder — edit them and press **R** in the app (or refresh).

---

## 5. Definitions

- **Sales** = `SUM(Sales Value)` for the month. Volume is the line `Quantity` in MT;
  `Qnt (Ton)` is ignored (its divisor flips between 1 and 1,000 across months).
- **Collection** = `SUM(Collection Amount)`. *Sales vs Coll* = sales ÷ collection (above 1.00x means
  you billed more than you collected).
- **Gross profit** = sales − `COGS_Database.Amount`; invoice cost is used when a district/warehouse
  filter is on (COGS has no district dimension).
- **Receivables** = latest `Aging_Database` snapshot (`As_At`) on or before the month end.
  The sheet has **no Overdue column**, so overdue = total − 0–30 bucket.
- **Inventory** = latest stock snapshot per warehouse × product, valued at the last COGS rate
  (fallback: last sale price, then the family median). This is a **cost** valuation.
- **Prior MTD** = previous month, 1st to the same day.
- **Days remaining / asking per day** use the calendar month minus the data date.

---

## 6. Validation checks (Detail & Audit)

Core sheets loaded · date integrity · product and customer sums equal total sales · returns ·
unclassified products · quantity × rate = value · unit-price sanity · aging buckets = total ·
overdue ≤ total · stock valuation coverage · COGS coverage vs invoice cost · collection completeness ·
targets present.

---

## 7. Mobile & theme

Below 700 px the KPI grid becomes two columns, the logo/title shrink, chart rows stack full width
and the sidebar collapses. The theme switch is per session and is written to `?theme=light|brand`.

---

## 8. Deploy to GitHub + Streamlit Community Cloud (free)

**Important:** Streamlit Cloud cannot see your `G:` drive. It clones the GitHub repository and runs
what is inside it. So *adding values to your local workbook does nothing online* until the file is
inside the repo and pushed — or the app is pointed at Google Sheets instead.

### Option A — workbook inside the repo (simple)

1. Push the folder as the repo root:

   ```bat
   cd /d "G:\...\Nobayon_Dashboard"
   git init
   git add .
   git commit -m "Nobayon dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```

2. Put the workbook in `data/` and commit it (2.4 MB — well under GitHub's limits):

   ```bat
   python refresh_pipeline.py --sync-data
   git add data/"dashboard database.xlsx"
   git commit -m "Data refresh"
   git push
   ```

3. <https://share.streamlit.io> → **Create app** → pick the repo/branch → **Main file path**:
   `streamlit_app.py` → Python **3.12** → Deploy.

4. Daily updates: the Task Scheduler job runs `refresh_pipeline.py --sync-data` (without `--no-git`)
   → commit + push → Streamlit redeploys automatically (1–2 min).

**On the cloud the “Refresh data” button cannot run Power Query** (no Excel on Linux). It re-reads the
committed workbook, so the data only changes when you push. Everything else works normally.

### Option B — Google Sheets (live data, no redeploy)

1. `python refresh_pipeline.py --no-git --export-csv` → upload the five CSVs as tabs with the same names.
2. Share the sheet: *Anyone with the link → Viewer*.
3. Streamlit Cloud → **Settings → Secrets**:

   ```toml
   GOOGLE_SHEET_ID = "1AbC...your_sheet_id"
   ```

4. The app reads the sheet live; **Refresh data** re-reads it without Excel, and the cache refreshes
   hourly on its own. No redeploy is needed when the sheet changes.

### Free-tier notes

- 1 GB RAM per app — this app needs ~250 MB; the 2.4 MB workbook is fine.
- Apps sleep after ~12 h of inactivity; the first visit wakes them in ~30 s.
- Secrets are supported (used for `GOOGLE_SHEET_ID`).
- The GitHub Action compiles every module on each push; data validation runs only when the workbook
  is committed.

---

## 9. Troubleshooting

| Symptom | Fix |
|---|---|
| `dashboard database.xlsx was not found` | Put it in this folder or in `..\Business_dashboard`, or set `MASTER_EXCEL_PATH`. |
| Refresh says **Power Query did not run** | Excel/pywin32 unavailable (e.g. Streamlit Cloud). Run on the Windows machine or use the Google Sheets route. |
| Refresh says **read-only** | The workbook is locked by another Excel window — close it and retry. |
| Numbers look stale | Press **Refresh data**; check “Workbook updated … min ago” in the sidebar. |
| Targets missing | Add the month to `targets.csv` / `product_targets.csv` (format `YYYY-MM`). |
