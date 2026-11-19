/**
 * Nobayon Traders - refresh the live dashboard Google Sheet from the Power Query workbook.
 *
 * No API keys, no service account: this runs inside your own Google account.
 *
 * SETUP (once)
 *   1. Open the live sheet  ->  Extensions  ->  Apps Script
 *   2. Paste this whole file over Code.gs
 *   3. Left panel  ->  Services +  ->  Drive API  ->  Add   (needed for the xlsx -> Sheet conversion)
 *   4. Run  refreshFromWorkbook  once and accept the permissions
 *   5. Triggers (clock icon)  ->  Add Trigger
 *        function: refreshFromWorkbook
 *        event source: Time-driven
 *        type: Day timer  /  time: 8am-9am      (add more triggers for 1pm and 6pm)
 *
 * The workbook id below is your live "dashboard database.xlsx" in Google Drive; the sheet id is
 * "Nobayon Dashboard Data", which the Streamlit app reads.
 */

const SOURCE_XLSX_ID = '1me6vHlyHB503xt9AeOALQ6hw2vvYAsMP';
const TARGET_SHEET_ID = '1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY';
const TABS = ['SO_DO_Database', 'Collection', 'COGS_Database', 'Aging_Database', 'Stock_Database'];

function refreshFromWorkbook() {
  const started = new Date();

  // 1) Convert the .xlsx into a temporary Google Sheet (Drive keeps the original untouched)
  const temp = Drive.Files.copy(
    { title: 'tmp_nobayon_' + started.getTime(), mimeType: MimeType.GOOGLE_SHEETS },
    SOURCE_XLSX_ID,
    { convert: true }
  );
  const source = SpreadsheetApp.openById(temp.id);
  const target = SpreadsheetApp.openById(TARGET_SHEET_ID);

  // 2) Copy each tab, converting date cells to ISO text so the app parses them the same way
  TABS.forEach(function (tabName) {
    const src = source.getSheetByName(tabName);
    if (!src) {
      Logger.log('SKIP (missing in workbook): ' + tabName);
      return;
    }

    const values = src.getDataRange().getValues();
    if (!values.length) {
      Logger.log('SKIP (empty): ' + tabName);
      return;
    }

    const header = values[0];
    const dateCols = [];
    for (let c = 0; c < header.length; c++) {
      const h = String(header[c]).trim().toLowerCase();
      if (h === 'date' || h === 'as_at') dateCols.push(c);
    }
    if (dateCols.length) {
      for (let r = 1; r < values.length; r++) {
        for (let k = 0; k < dateCols.length; k++) {
          const cell = values[r][dateCols[k]];
          if (cell instanceof Date) {
            values[r][dateCols[k]] = Utilities.formatDate(cell, 'UTC', 'yyyy-MM-dd');
          }
        }
      }
    }

    let dst = target.getSheetByName(tabName);
    if (!dst) dst = target.insertSheet(tabName);

    dst.clearContents();
    if (dst.getMaxRows() < values.length) {
      dst.insertRowsAfter(dst.getMaxRows(), values.length - dst.getMaxRows());
    }
    if (dst.getMaxColumns() < values[0].length) {
      dst.insertColumnsAfter(dst.getMaxColumns(), values[0].length - dst.getMaxColumns());
    }
    dst.getRange(1, 1, values.length, values[0].length).setValues(values);
    Logger.log(tabName + ': ' + values.length + ' rows x ' + values[0].length + ' cols');
  });

  // 3) Clean up the temporary conversion
  DriveApp.getFileById(temp.id).setTrashed(true);

  Logger.log('Finished in ' + ((new Date() - started) / 1000).toFixed(1) + 's');
}
