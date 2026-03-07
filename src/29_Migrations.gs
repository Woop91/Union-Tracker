// ============================================================================
// MIGRATIONS — one-time data/schema fixes
// Run each function ONCE from the Apps Script editor. They are all idempotent
// (safe to re-run — they detect already-migrated state and no-op).
// ============================================================================

/**
 * MIGRATION v4.20.25 — Contact Log Folder URL → Member Admin Folder URL
 *
 * Background:
 *   v4.20.25 renamed the Member Directory column from 'Contact Log Folder URL'
 *   to 'Member Admin Folder URL'. createMemberDirectory() appends the new column
 *   if it is missing, but deployments that ran CREATE_DASHBOARD before this
 *   version end up with BOTH columns: the old one (has data) and the new one
 *   (empty). This migration:
 *     1. Finds the old column by header name (never by index).
 *     2. Finds the new column by header name.
 *     3. For each row where old has a value and new is empty, copies the value.
 *     4. Clears the old column header and all its data (replacing header with
 *        a blank so the column is inert but row count is preserved).
 *     5. Toasts a summary and logs every action.
 *
 * Safe to re-run:
 *   - If old column is absent → logs "already migrated" and exits.
 *   - If new column is absent → exits with error toast (run CREATE_DASHBOARD first).
 *   - If a row already has a value in the new column → skipped (never overwrites).
 *
 * Run from Apps Script editor: migrateContactLogFolderUrlColumn()
 */
function migrateContactLogFolderUrlColumn() {
  var OLD_HEADER = 'Contact Log Folder URL';
  var NEW_HEADER = 'Member Admin Folder URL';

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) {
      SpreadsheetApp.getActive().toast('Member Directory sheet not found.', '❌ Migration failed', 8);
      return;
    }

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol < 1) {
      SpreadsheetApp.getActive().toast('Member Directory is empty.', '❌ Migration failed', 8);
      return;
    }

    // Read header row
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    var headers     = headerRange.getValues()[0];

    var oldIdx = -1;
    var newIdx = -1;
    for (var h = 0; h < headers.length; h++) {
      var norm = String(headers[h]).trim().toLowerCase();
      if (norm === OLD_HEADER.toLowerCase()) oldIdx = h;
      if (norm === NEW_HEADER.toLowerCase()) newIdx = h;
    }

    // ── Already migrated? ───────────────────────────────────────────────────
    if (oldIdx === -1) {
      Logger.log('migrateContactLogFolderUrlColumn: old column "' + OLD_HEADER + '" not found — already migrated or never existed. Nothing to do.');
      SpreadsheetApp.getActive().toast('"' + OLD_HEADER + '" column not found — nothing to migrate.', '✅ Already clean', 5);
      return;
    }

    // ── New column must exist (run CREATE_DASHBOARD first) ──────────────────
    if (newIdx === -1) {
      Logger.log('migrateContactLogFolderUrlColumn: new column "' + NEW_HEADER + '" not found. Run CREATE_DASHBOARD first to add it.');
      SpreadsheetApp.getActive().toast('Run CREATE_DASHBOARD first to add the "' + NEW_HEADER + '" column, then re-run this migration.', '⚠ New column missing', 10);
      return;
    }

    Logger.log('migrateContactLogFolderUrlColumn: old col index=' + oldIdx + ', new col index=' + newIdx + ', rows=' + lastRow);

    // ── Read all data rows ──────────────────────────────────────────────────
    var copied   = 0;
    var skipped  = 0; // new already had a value

    if (lastRow > 1) {
      var dataRange  = sheet.getRange(2, 1, lastRow - 1, lastCol);
      var data       = dataRange.getValues();

      var newColCells = sheet.getRange(2, newIdx + 1, lastRow - 1, 1);
      var newColVals  = newColCells.getValues();

      for (var r = 0; r < data.length; r++) {
        var oldVal = String(data[r][oldIdx]).trim();
        var newVal = String(newColVals[r][0]).trim();

        if (oldVal && !newVal) {
          newColVals[r][0] = oldVal;
          copied++;
        } else if (oldVal && newVal) {
          skipped++;
          Logger.log('migrateContactLogFolderUrlColumn: row ' + (r + 2) + ' skipped — new column already has value: ' + newVal);
        }
      }

      if (copied > 0) {
        newColCells.setValues(newColVals);
        Logger.log('migrateContactLogFolderUrlColumn: copied ' + copied + ' URL(s) to new column.');
      }
    }

    // ── Clear old column (header + data) ────────────────────────────────────
    sheet.getRange(1, oldIdx + 1).setValue('');  // blank header — inert but col stays
    if (lastRow > 1) {
      sheet.getRange(2, oldIdx + 1, lastRow - 1, 1).clearContent();
    }
    Logger.log('migrateContactLogFolderUrlColumn: old column cleared.');

    // ── Summary ─────────────────────────────────────────────────────────────
    var msg = 'Copied ' + copied + ' URL(s). Skipped ' + skipped + ' (already set). Old column cleared.';
    Logger.log('migrateContactLogFolderUrlColumn: complete — ' + msg);
    SpreadsheetApp.getActive().toast(msg, '✅ Migration complete', 8);

  } catch (err) {
    Logger.log('migrateContactLogFolderUrlColumn ERROR: ' + err.message);
    SpreadsheetApp.getActive().toast('Migration error: ' + err.message, '❌ Migration failed', 10);
  }
}
