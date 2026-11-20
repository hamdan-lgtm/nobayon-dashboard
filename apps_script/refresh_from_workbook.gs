/**
 * Nobayon Traders - refresh the live dashboard Google Sheet from the Power Query workbook.
 *
 * Runs entirely inside your Google account: no API keys, no service account.
 *
 * SETUP (once)
 *   1. Open the live sheet  ->  Extensions  ->  Apps Script
 *   2. Paste this file over Code.gs   (save)
 *   3. Services +  ->  Drive API  ->  Add        <-- required, the script copies the xlsx
 *   4. Run  refreshFromWorkbook  once, approve the permissions
 *   5. Triggers  ->  Add Trigger:
 *        function: refreshFromWorkbook
 *        event source: Time-driven
 *        type: Minutes timer  ->  every 5 minutes
 *
 * Why it is safe to run every 5 minutes:
 *   - it exits in ~1 s when the workbook has not changed since the last publish
 *   - a lock stops a second run from starting while one is still copying
 *   - a tab is only cleared once the replacement values are already in memory
 */

const SOURCE_XLSX_ID = '1me6vHlyHB503xt9AeOALQ6hw2vvYAsMP';
const TARGET_SHEET_ID = '1Q9BiOVrf3-vKWfeoGVWB_mfRK1-gZvgy1bLu-z-j4bY';
const TABS = ['SO_DO_Database', 'Collection', 'COGS_Database', 'Aging_Database', 'Stock_Database'];
const PROPS = PropertiesService.getScriptProperties();

function refreshFromWorkbook() {
  const started = new Date();
  const lock = LockService.getScriptLock();

  // A previous run is still copying - skip instead of queueing up (this is what produces
  // "Too many simultaneous invocations" and burns the daily quota).
  if (!lock.tryLock(0)) {
    Logger.log('Previous run still in progress - skipped');
    return;
  }

  try {
    // 1) Cheap change check - about 1 second, no conversion, no file created
    const sourceFile = DriveApp.getFileById(SOURCE_XLSX_ID);
    const stamp = sourceFile.getLastUpdated().toISOString();
    if (PROPS.getProperty('lastSynced') === stamp) {
      Logger.log('Workbook unchanged since ' + stamp + ' - nothing to do');
      return;
    }

    // 2) Convert the .xlsx into a temporary Google Sheet (the original stays untouched)
    const temp = Drive.Files.copy(
      { title: 'tmp_nobayon_' + started.getTime(), mimeType: MimeType.GOOGLE_SHEETS },
      SOURCE_XLSX_ID,
      { convert: true }
    );
    const source = SpreadsheetApp.openById(temp.id);
    const target = SpreadsheetApp.openById(TARGET_SHEET_ID);

    try {
      // 3) Copy each tab, turning date cells into ISO text
      TABS.forEach(function (tabName) {
        const src = source.getSheetByName(tabName);
        if (!src) {
          Logger.log('SKIP (missing): ' + tabName);
          return;
        }
        const values = src.getDataRange().getValues();
        if (!values.length || values[0].length === 0) {
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

        // Grow the grid first, then wipe, then write - never leave a tab half empty
        if (dst.getMaxRows() < values.length) {
          dst.insertRowsAfter(dst.getMaxRows(), values.length - dst.getMaxRows());
        }
        if (dst.getMaxColumns() < values[0].length) {
          dst.insertColumnsAfter(dst.getMaxColumns(), values[0].length - dst.getMaxColumns());
        }
        dst.clearContents();
        dst.getRange(1, 1, values.length, values[0].length).setValues(values);

        Logger.log(tabName + ': ' + values.length + ' rows x ' + values[0].length + ' cols');
      });
    } finally {
      // 4) Always drop the temporary copy, even if a tab failed
      try { DriveApp.getFileById(temp.id).setTrashed(true); } catch (e) { Logger.log('cleanup: ' + e); }
    }

    // 5) Only remember the version if the whole publish succeeded
    PROPS.setProperty('lastSynced', stamp);
    PROPS.setProperty('lastSuccess', started.toISOString());
    Logger.log('Finished in ' + ((new Date() - started) / 1000).toFixed(1) + 's');
  } catch (err) {
    PROPS.setProperty('lastError', started.toISOString() + ' :: ' + err.message);
    Logger.log('FAILED: ' + err.message + '\n' + (err.stack || ''));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Run once by hand (or from a button) to force a rebuild even if nothing changed. */
function forceRefresh() {
  PROPS.deleteProperty('lastSynced');
  refreshFromWorkbook();
}

/** Prints what the last runs did - handy when the trigger shows an error. */
function showStatus() {
  Logger.log('last success : ' + (PROPS.getProperty('lastSuccess') || 'never'));
  Logger.log('last synced  : ' + (PROPS.getProperty('lastSynced') || 'never'));
  Logger.log('last error   : ' + (PROPS.getProperty('lastError') || 'none'));
}
