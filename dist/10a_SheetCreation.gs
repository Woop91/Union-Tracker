/**
 * ============================================================================
 * 10a_SheetCreation.gs — Individual Sheet Creation Functions
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Individual sheet creation functions for Config, Member Directory, and
 *   Grievance Log sheets. Each function creates a specific sheet with its
 *   headers, formatting, data validation rules, and protection settings.
 *   createConfigSheet() sets up the column-based configuration layout
 *   (row 1=section headers, row 2=column headers, row 3+=values).
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Split from 08a_SheetSetup.gs for maintainability. Each sheet creation
 *   function is standalone and idempotent. The Config sheet uses a column-based
 *   layout (not row-based) because it's easier to add new config options by
 *   adding columns than by managing row indices.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Individual sheets can't be created or repaired. CREATE_DASHBOARD will fail
 *   at the specific sheet creation step. Existing sheets are unaffected.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, CONFIG_COLS, MEMBER_COLS, GRIEVANCE_COLS, COLORS)
 *   Used by:    08a_SheetSetup.gs (CREATE_DASHBOARD), 06_Maintenance.gs (REPAIR_DASHBOARD)
 *
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// MENU SYSTEM
// ============================================================================

/**
 * Creates the menu system when the spreadsheet opens
 * Reorganized into 5 logical menus for easier navigation
 */
// Note: onOpen() defined in modular file - see respective module

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main setup function - creates the complete Dashboard
 * Creates the core sheets with proper structure and formatting
 */

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Create or recreate the Config sheet with dropdown values
 * Comprehensive configuration with section groupings and organization settings
 */
function createConfigSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG);
  var isExistingSheet = sheet.getLastRow() > 2;

  // Only clear if sheet is new or has no meaningful data (≤2 rows = headers only)
  if (!isExistingSheet) {
    sheet.clear();
  } else {
    log_('createConfigSheet', 'Config sheet has ' + sheet.getLastRow() + ' rows of data — updating headers while preserving settings');

    // Rename columns whose headers changed between versions (e.g. Managers → Directors).
    // Must run BEFORE orphan detection so the renamed header is recognized as valid.
    var CONFIG_HEADER_RENAMES_ = { 'Managers': 'Directors' };
    var renameLastCol = sheet.getLastColumn();
    if (renameLastCol > 0) {
      var renameRow = sheet.getRange(2, 1, 1, renameLastCol).getValues()[0];
      for (var rn = 0; rn < renameRow.length; rn++) {
        var oldName = String(renameRow[rn]).trim();
        if (CONFIG_HEADER_RENAMES_[oldName]) {
          sheet.getRange(2, rn + 1).setValue(CONFIG_HEADER_RENAMES_[oldName]);
          log_('createConfigSheet', 'renamed header "' + oldName + '" → "' + CONFIG_HEADER_RENAMES_[oldName] + '"');
        }
      }
    }

    // Migration: detect and remove orphaned columns left behind by past
    // CONFIG_HEADER_MAP_ removals (Yes/No, Satisfaction Form URL, etc.).
    // Deleting them shifts data left so it re-aligns with new headers.
    var _migrated = _migrateOrphanedColumns(sheet);
  }

  // Row 1: Section Headers (grouped categories) — must match CONFIG_HEADER_MAP_ length (81 cols)
  var sectionHeaders = [
    '── ORGANIZATION ──', '', '', '', '', '', '', '', '', '',     // cols 1-10 (10)
    '── CONTACT INFO ──', '', '', '', '', '',                     // cols 11-16 (6)
    '── EMPLOYMENT ──', '', '', '', '', '', '',                   // cols 17-23 (7)
    '── PEOPLE ──', '', '', '', '', '',                           // cols 24-29 (6)
    '── GRIEVANCE SETTINGS ──', '', '', '', '', '', '',           // cols 30-36 (7)
    '��─ DEADLINES ��─', '', '', '', '', '', '', '',                // cols 37-44 (8)
    '── NOTIFICATIONS ──', '', '',                                // cols 45-47 (3)
    '── LINKS ──', '', '', '', '', '', '',                        // cols 48-54 (7)
    '── DRIVE & CALENDAR ��─', '', '', '', '', '', '', '', '', '', '', // cols 55-65 (11)
    '── SURVEY ──', '', '',                                       // cols 66-68 (3)
    '── BRANDING & UX ──', '', '', '', '', '',                    // cols 69-74 (6)
    '── FEATURE TOGGLES ──', '', '', '', '',                      // cols 75-79 (5)
    '── RETENTION ──', ''                                         // cols 80-81 (2)
  ];

  // Row 2: Column Headers — auto-derived from CONFIG_HEADER_MAP_
  var columnHeaders = getHeadersFromMap_(CONFIG_HEADER_MAP_);

  // Ensure sheet has enough columns for all headers (handles sheets created before new columns were added)
  ensureMinimumColumns(sheet, columnHeaders.length);

  // Quick-exit for migration: if existing headers already match expected order, skip the
  // expensive data-remap block below (still apply headers, validations, and defaults).
  var _skipMigration = false;
  if (isExistingSheet) {
    var existingLastCol = sheet.getLastColumn();
    if (existingLastCol >= columnHeaders.length) {
      var existingHeaders = sheet.getRange(2, 1, 1, existingLastCol).getValues()[0];
      _skipMigration = columnHeaders.every(function(h, i) { return existingHeaders[i] === h; });
      if (_skipMigration) {
        log_('createConfigSheet', 'headers already match — skipping migration');
      }
    }
  }

  // ── Data migration: remap row 3+ data when headers are reordered ──
  // If the existing sheet has headers in a different order than CONFIG_HEADER_MAP_,
  // we must move the data to match the new header positions before overwriting headers.
  // Without this, header reorder causes data misalignment (e.g. org name reads "yes"
  // because the boolean toggle data stayed in the old column).
  if (isExistingSheet && !_skipMigration) {
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var oldHeaders = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      var lastDataRow = sheet.getLastRow();

      // Build old header text → 0-indexed column map
      var oldHeaderToCol = {};
      for (var oh = 0; oh < oldHeaders.length; oh++) {
        var hText = String(oldHeaders[oh]).trim();
        if (hText) oldHeaderToCol[hText] = oh;
      }

      // Check if any headers moved position
      var needsMigration = false;
      for (var nh = 0; nh < columnHeaders.length; nh++) {
        var oldIdx = oldHeaderToCol[columnHeaders[nh]];
        if (oldIdx !== undefined && oldIdx !== nh) {
          needsMigration = true;
          break;
        }
      }

      if (needsMigration && lastDataRow >= 3) {
        log_('createConfigSheet', 'headers reordered — migrating row 3+ data to match new layout');
        var dataRows = sheet.getRange(3, 1, lastDataRow - 2, lastCol).getValues();

        // Remap each data row: new column order pulls from old column positions
        var remapped = [];
        for (var dr = 0; dr < dataRows.length; dr++) {
          var newRow = [];
          for (var nc = 0; nc < columnHeaders.length; nc++) {
            var srcIdx = oldHeaderToCol[columnHeaders[nc]];
            newRow.push(srcIdx !== undefined && srcIdx < dataRows[dr].length ? dataRows[dr][srcIdx] : '');
          }
          remapped.push(newRow);
        }

        // Clear old data first to avoid stale values in columns beyond new width
        sheet.getRange(3, 1, lastDataRow - 2, lastCol).clearContent();
        sheet.getRange(3, 1, remapped.length, columnHeaders.length).setValues(remapped);
        log_('createConfigSheet', 'migrated ' + remapped.length + ' data row(s) across ' + columnHeaders.length + ' columns');
      }
    }
  }

  // Always apply headers (Row 1 & 2) — safe to overwrite since these are structure, not data.
  // This ensures existing sheets pick up newly-added columns beyond AZ.

  // Apply section headers (Row 1)
  sheet.getRange(1, 1, 1, sectionHeaders.length).setValues([sectionHeaders])
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK)
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Apply column headers (Row 2)
  sheet.getRange(2, 1, 1, columnHeaders.length).setValues([columnHeaders])
    .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Re-sync CONFIG_COLS from the freshly-written headers.  Without this,
  // seedConfigDefault_ and _applyYesNoValidation use stale positions resolved
  // from the PRE-migration layout (e.g. SHOW_GRIEVANCES → old col 70 instead
  // of new col 75), causing yes/no dropdowns to land on Branding columns.
  syncColumnMaps();

  // Repair stale yes/no validations and misaligned data left on non-toggle
  // columns by prior runs that used stale CONFIG_COLS (pre-v4.50.5 bug).
  // The old bug applied yes/no dropdowns to Branding columns (Logo Initials,
  // Steward Label, etc.) instead of Feature Toggle columns, letting users
  // accidentally set branding values to "yes"/"no".
  if (isExistingSheet) {
    // Every non-toggle Config column that the pre-v4.50.5 stale-CONFIG_COLS
    // bug could have smeared yes/no dropdowns onto.  Covers Branding & UX,
    // the numeric Feature Toggle fields, and Retention.
    var nonToggleRepairCols = [
      CONFIG_COLS.ACCENT_HUE, CONFIG_COLS.LOGO_INITIALS,
      CONFIG_COLS.STEWARD_LABEL, CONFIG_COLS.MEMBER_LABEL,
      CONFIG_COLS.MAGIC_LINK_EXPIRY_DAYS, CONFIG_COLS.COOKIE_DURATION_DAYS,
      CONFIG_COLS.INSIGHTS_CACHE_TTL_MIN,
      CONFIG_COLS.GRIEVANCE_ARCHIVE_DAYS, CONFIG_COLS.AUDIT_ARCHIVE_DAYS
    ];
    for (var bc = 0; bc < nonToggleRepairCols.length; bc++) {
      if (!nonToggleRepairCols[bc]) continue;
      try {
        var repairCell = sheet.getRange(3, nonToggleRepairCols[bc]);
        repairCell.clearDataValidations();
        // Clear yes/no values from non-boolean columns — these are toggle values
        // that landed here due to the stale CONFIG_COLS bug.
        var repairVal = String(repairCell.getValue() || '').toLowerCase().trim();
        if (repairVal === 'yes' || repairVal === 'no') {
          repairCell.clearContent();
          log_('createConfigSheet', 'cleared misaligned toggle value "' + repairVal + '" from non-toggle col ' + nonToggleRepairCols[bc]);
        }
      } catch (_v) {}
    }
  }

  // Add default dropdown values (Row 3+)
  // For existing sheets, seedConfigDefault_ only writes to columns that are currently empty.
  // This ensures re-running CREATE_DASHBOARD fills in newly-added columns without
  // overwriting user-customized values.

  // ── Organization (A–J)
  seedConfigDefault_(sheet, CONFIG_COLS.ORG_NAME, ['Your Union Name'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ORG_ABBREV, ['DDS'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.LOCAL_NUMBER, ['000'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.UNION_PARENT, ['Your Parent Union'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STATE_REGION, ['Your State'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ORG_WEBSITE, ['https://www.example.org/'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CONTRACT_NAME, ['Current CBA'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CONTRACT_GRIEVANCE, ['Article XX'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CONTRACT_DISCIPLINE, ['Article YY'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CONTRACT_WORKLOAD, ['Article ZZ'], isExistingSheet);

  // ── Contact Info (K–P)
  seedConfigDefault_(sheet, CONFIG_COLS.MAIN_ADDRESS, ['123 Main Street, Suite 100, City, ST 00000'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MAIN_PHONE, ['(000) 000-0000'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MAIN_FAX, ['(000) 000-0000'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MAIN_CONTACT_NAME, ['Your Contact Name'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MAIN_CONTACT_EMAIL, ['your-email@your-org.org'], isExistingSheet);

  // ── Employment (Q–W)
  seedConfigDefault_(sheet, CONFIG_COLS.OFFICE_DAYS, DEFAULT_CONFIG.OFFICE_DAYS, isExistingSheet);

  // ── People (X–AC)
  var committees = ['Grievance Committee', 'Bargaining Committee', 'Health & Safety Committee',
                    'Political Action Committee', 'Membership Committee', 'Executive Board'];
  seedConfigDefault_(sheet, CONFIG_COLS.STEWARD_COMMITTEES, committees, isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.COMM_METHODS, DEFAULT_CONFIG.COMM_METHODS, isExistingSheet);
  var bestTimes = ['Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Weekends', 'Flexible'];
  seedConfigDefault_(sheet, CONFIG_COLS.BEST_TIMES, bestTimes, isExistingSheet);

  // ── Grievance Settings (AD–AJ)
  seedConfigDefault_(sheet, CONFIG_COLS.GRIEVANCE_STATUS, DEFAULT_CONFIG.GRIEVANCE_STATUS, isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.GRIEVANCE_STEP, DEFAULT_CONFIG.GRIEVANCE_STEP, isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ISSUE_CATEGORY, DEFAULT_CONFIG.ISSUE_CATEGORY, isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ARTICLES, DEFAULT_CONFIG.ARTICLES, isExistingSheet);
  var escalationStatuses = COMMAND_CONFIG.ESCALATION_STATUSES || ['In Arbitration', 'Appealed'];
  seedConfigDefault_(sheet, CONFIG_COLS.ESCALATION_STATUSES, escalationStatuses, isExistingSheet);
  var escalationSteps = COMMAND_CONFIG.ESCALATION_STEPS || ['Step II', 'Step III', 'Arbitration'];
  seedConfigDefault_(sheet, CONFIG_COLS.ESCALATION_STEPS, escalationSteps, isExistingSheet);

  // ── Deadlines (AK–AR) — values from DEADLINE_DEFAULTS (01_Core.gs)
  seedConfigDefault_(sheet, CONFIG_COLS.FILING_DEADLINE_DAYS, [DEADLINE_DEFAULTS.FILING_DAYS], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEP1_RESPONSE_DAYS, [DEADLINE_DEFAULTS.STEP_1_RESPONSE], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEP2_APPEAL_DAYS, [DEADLINE_DEFAULTS.STEP_2_APPEAL], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEP2_RESPONSE_DAYS, [DEADLINE_DEFAULTS.STEP_2_RESPONSE], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEP3_APPEAL_DAYS, [DEADLINE_DEFAULTS.STEP_3_APPEAL], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEP3_RESPONSE_DAYS, [DEADLINE_DEFAULTS.STEP_3_RESPONSE], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ARBITRATION_DEMAND_DAYS, [DEADLINE_DEFAULTS.ARBITRATION_DEMAND], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.ALERT_DAYS, ['3, 7, 14'], isExistingSheet);

  // ── Links (AV–BB)
  seedConfigDefault_(sheet, CONFIG_COLS.CUSTOM_LINK_1_NAME, ['Resources'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CUSTOM_LINK_1_URL, ['https://www.example.org/resources'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CUSTOM_LINK_2_NAME, ['Help'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.CUSTOM_LINK_2_URL, ['https://www.example.org/help'], isExistingSheet);

  // ── Survey (BN–BP)
  seedConfigDefault_(sheet, CONFIG_COLS.SURVEY_PRIORITY_OPTIONS, [
    'Contract Enforcement',
    'Workload & Staffing',
    'Scheduling & Office Days',
    'Pay & Benefits',
    'Health & Safety',
    'Training & Development',
    'Equity & Inclusion',
    'Communication',
    'Steward Support & Access',
    'Member Organizing',
    'Other'
  ], isExistingSheet);

  // ── Branding & UX (BQ–BV)
  seedConfigDefault_(sheet, CONFIG_COLS.ACCENT_HUE, [30], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.LOGO_INITIALS, ['DDS'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.STEWARD_LABEL, ['Steward'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MEMBER_LABEL, ['Member'], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.MAGIC_LINK_EXPIRY_DAYS, [7], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.COOKIE_DURATION_DAYS, [30], isExistingSheet);

  // ── Feature Toggles (BW–CA) — yes/no dropdown validation
  seedConfigDefault_(sheet, CONFIG_COLS.INSIGHTS_CACHE_TTL_MIN, [5], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.SHOW_GRIEVANCES, ['yes'], isExistingSheet);
  _applyYesNoValidation(sheet, CONFIG_COLS.SHOW_GRIEVANCES,
    'yes = show grievance tracking features. no = hide all grievance UI and endpoints.');

  seedConfigDefault_(sheet, CONFIG_COLS.BROADCAST_SCOPE_ALL, ['yes'], isExistingSheet);
  _applyYesNoValidation(sheet, CONFIG_COLS.BROADCAST_SCOPE_ALL,
    'yes = stewards can broadcast to all members. no = stewards can only broadcast to their assigned members.');

  seedConfigDefault_(sheet, CONFIG_COLS.ENABLE_CORRELATION, ['yes'], isExistingSheet);
  _applyYesNoValidation(sheet, CONFIG_COLS.ENABLE_CORRELATION,
    'yes = enable correlation analysis on Insights page. no = disable (reduces compute).');

  seedConfigDefault_(sheet, CONFIG_COLS.ENABLE_TAB_MODALS, ['yes'], isExistingSheet);
  _applyYesNoValidation(sheet, CONFIG_COLS.ENABLE_TAB_MODALS,
    'yes = show contextual modals when navigating to sheet tabs. no = disable tab modals.');

  // ── Retention (CB–CC)
  seedConfigDefault_(sheet, CONFIG_COLS.GRIEVANCE_ARCHIVE_DAYS, [90], isExistingSheet);
  seedConfigDefault_(sheet, CONFIG_COLS.AUDIT_ARCHIVE_DAYS, [90], isExistingSheet);

  // Freeze header rows (1 and 2)
  sheet.setFrozenRows(2);

  // Auto-resize all columns
  sheet.autoResizeColumns(1, columnHeaders.length);

  // Set minimum column widths for readability
  for (var i = 1; i <= columnHeaders.length; i++) {
    if (sheet.getColumnWidth(i) < 100) {
      sheet.setColumnWidth(i, 100);
    }
  }

  // Apply full Config sheet styling
  applyConfigSheetStyling(sheet);

  // Set tab color
  sheet.setTabColor(COLORS.PRIMARY_PURPLE);
}

/**
 * General-purpose migration: detects and removes orphaned columns in the Config sheet.
 *
 * When a column is removed from CONFIG_HEADER_MAP_ (e.g. Yes/No, Satisfaction Form URL),
 * existing sheets retain the physical column. This shifts all data to the right of it,
 * causing headers and data to fall out of alignment.
 *
 * Strategy: walk actual row-2 headers alongside expected headers. When a mismatch is
 * found (actual ≠ expected), the actual column is orphaned — delete it. Deleting
 * shifts subsequent data left, restoring alignment.
 *
 * Works in two scenarios:
 *  A) Old headers still present — orphan detected by header text mismatch.
 *  B) Headers already overwritten — first 81 cols match, extras past col 81 are deleted.
 *
 * @param {Sheet} sheet - The Config sheet
 * @returns {number} Number of orphaned columns deleted
 * @private
 */
function _migrateOrphanedColumns(sheet) {
  var expected = getHeadersFromMap_(CONFIG_HEADER_MAP_);
  var maxCol = sheet.getMaxColumns();
  if (maxCol <= expected.length) return 0;

  var row2 = sheet.getRange(2, 1, 1, maxCol).getValues()[0];

  // Build a set of all expected header names for safe lookup.
  // Only delete columns whose header text does NOT appear anywhere in
  // CONFIG_HEADER_MAP_. The old two-pointer algorithm assumed sequential
  // order and would incorrectly mark reordered columns as orphans.
  var expectedSet = {};
  for (var e = 0; e < expected.length; e++) {
    expectedSet[expected[e]] = true;
  }

  var toDelete = [];
  for (var ci = 0; ci < row2.length; ci++) {
    var actual = String(row2[ci]).trim();
    if (actual === '') continue; // skip blank headers
    if (!expectedSet[actual]) {
      toDelete.push(ci + 1); // truly orphaned — not in expected set
    }
  }

  // Also mark trailing blank columns beyond the expected count
  if (row2.length > expected.length) {
    for (var ti = expected.length; ti < row2.length; ti++) {
      var trailing = String(row2[ti]).trim();
      if (trailing === '' && toDelete.indexOf(ti + 1) === -1) {
        toDelete.push(ti + 1);
      }
    }
  }

  if (toDelete.length === 0) return 0;

  // Group contiguous columns for batch deletion (right-to-left)
  toDelete.sort(function(a, b) { return b - a; });
  var i = 0;
  while (i < toDelete.length) {
    var start = toDelete[i];
    var count = 1;
    // Check for contiguous columns (sorted descending, so contiguous = consecutive decreasing by 1)
    while (i + count < toDelete.length && toDelete[i + count] === start - count) {
      count++;
    }
    if (count > 1) {
      sheet.deleteColumns(start - count + 1, count); // deleteColumns(startCol, numCols)
    } else {
      sheet.deleteColumn(start);
    }
    i += count;
  }
  log_('_migrateOrphanedColumns', 'deleted ' + toDelete.length +
    ' orphaned column(s) — sheet now has ' + sheet.getMaxColumns() + ' columns');
  return toDelete.length;
}

/**
 * Repairs Config sheet data alignment.
 * Clears all data rows (3+) and re-seeds defaults. Call from menu after
 * orphaned-column migration if data is still misaligned.
 *
 * Safe to call multiple times — only writes default values.
 * User-customized values in affected columns WILL be lost.
 */
function repairConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Config sheet not found. Run CREATE_DASHBOARD first.');
    return;
  }

  var expected = getHeadersFromMap_(CONFIG_HEADER_MAP_);
  var maxCol = sheet.getMaxColumns();

  // Step 1: if sheet still has extra columns, run migration
  if (maxCol > expected.length) {
    var deleted = _migrateOrphanedColumns(sheet);
    if (deleted > 0) {
      ss.toast('Removed ' + deleted + ' orphaned column(s).', 'Migration', 3);
    }
  }

  // Step 2: clear ALL data rows (row 3+) — a full reset
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, Math.max(expected.length, maxCol)).clearContent();
  }

  // Step 3: rebuild Config from scratch (headers + default seeds + toggle validations)
  createConfigSheet(ss);
  SpreadsheetApp.flush();  // ensure writes are visible to subsequent reads

  // Step 4: backfill Config dropdown columns from existing sheet data
  // (Job Titles, Locations, Stewards, Statuses, etc. from Member Dir & Grievance Log)
  ss.toast('Populating Config from existing sheet data...', '🔧 Repair', 3);
  populateConfigFromSheetData();
  SpreadsheetApp.flush();

  // Step 5: reapply data validations so dropdowns reflect new Config values
  ss.toast('Reapplying dropdown validations...', '🔧 Repair', 3);
  setupDataValidations();

  ss.toast('Config fully repaired — defaults seeded, sheet data imported, validations applied.', '✅ Repair Complete', 5);
}

/**
 * Seeds default values into a Config column only if the column is currently empty.
 * For new sheets (isExisting=false), always writes. For existing sheets,
 * checks whether the column already has data and skips if so.
 * @param {Sheet} sheet - The Config sheet
 * @param {number} col - 1-indexed column number
 * @param {Array} values - Array of default values to write
 * @param {boolean} isExisting - Whether the sheet already had data
 * @private
 */
function seedConfigDefault_(sheet, col, values, isExisting) {
  if (isExisting) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      var existing = sheet.getRange(3, col, lastRow - 2, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        // Truthiness check covers '', null, undefined, and 0 (no valid config default is 0)
        if (existing[i][0]) {
          return; // Column already has data, don't overwrite
        }
      }
    }
  }
  sheet.getRange(3, col, values.length, 1)
    .setValues(values.map(function(v) { return [v]; }));
}

/**
 * Safely applies a yes/no data-validation dropdown to a Config toggle cell.
 * Normalizes the cell value first, then applies the validation rule.
 * Entire operation is wrapped in try/catch — if GAS throws (known quirk where
 * setDataValidation echoes helpText as the error), we log and continue.
 * The dropdown is cosmetic; config reads work fine without it.
 * @param {Sheet} sheet - The Config sheet
 * @param {number} col - 1-indexed column number
 * @param {string} helpText - Hover help text for the dropdown
 * @private
 */
function _applyYesNoValidation(sheet, col, helpText) {
  if (!col) return;
  try {
    var cell = sheet.getRange(3, col);
    var raw = cell.getValue();
    var val = (raw === '' || raw == null) ? '' : String(raw).toLowerCase().trim();
    if (val === '') {
      cell.setValue('yes');
    } else if (val !== 'yes' && val !== 'no') {
      cell.setValue('yes');
      log_('_applyYesNoValidation', '_applyYesNoValidation col ' + col + ': replaced unexpected value "' + raw + '" with "yes"');
    }
    // Try with helpText first; GAS has a known quirk where setDataValidation
    // can throw when helpText is set.  Fall back to no helpText.
    var rule;
    try {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['yes', 'no'], true)
        .setAllowInvalid(true)
        .setHelpText(helpText)
        .build();
      cell.setDataValidation(rule);
    } catch (_ht) {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['yes', 'no'], true)
        .setAllowInvalid(true)
        .build();
      cell.setDataValidation(rule);
    }
  } catch (e) {
    log_('_applyYesNoValidation', '_applyYesNoValidation col ' + col + ' FAILED: ' + e.message);
  }
}

/**
 * Scans Grievance Log and Member Directory for unique values in dropdown
 * columns and backfills them into the Config sheet. This handles the case
 * where data was entered before dropdown sync was active, or where bulk
 * imports/pastes bypassed the onEdit trigger.
 *
 * Safe to run multiple times — skips values already present in Config.
 * Callable from menu: Union Hub > Admin > Populate Config from Sheet Data
 */
function populateConfigFromSheetData() {
  // Always resolve fresh column positions from actual sheet headers.
  // Prevents stale MEMBER_COLS/CONFIG_COLS after header renames.
  try { syncColumnMaps(); } catch (_e) { log_('populateConfig syncColumnMaps', (_e.message || _e)); }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) {
    try { SpreadsheetApp.getUi().alert('Config sheet not found. Run CREATE_DASHBOARD first.'); } catch (_) { log_('_', (_.message || _)); }
    return;
  }

  var added = 0;

  // Maps configCol → { existing: {}, newValues: [] }
  var colBuckets = {};

  // Pre-read ALL Config data in one batch
  var configLastRow = configSheet.getLastRow();
  var configDataRows = configLastRow >= 3 ? configLastRow - 2 : 0;
  var configMaxCol = configSheet.getLastColumn();
  var allConfigData = configDataRows > 0 && configMaxCol > 0
    ? configSheet.getRange(3, 1, configDataRows, configMaxCol).getValues()
    : [];

  // Helper: init a bucket for a Config column
  function initBucket_(configCol) {
    if (colBuckets[configCol]) return colBuckets[configCol];
    var existingSet = {};
    for (var c = 0; c < allConfigData.length; c++) {
      var cv = (allConfigData[c][configCol - 1] || '').toString().trim();
      if (cv) existingSet[cv] = true;
    }
    colBuckets[configCol] = { existing: existingSet, newValues: [] };
    return colBuckets[configCol];
  }

  // Helper: add a value to bucket if new
  function addValue_(bucket, val) {
    if (val && !bucket.existing[val]) {
      bucket.existing[val] = true;
      bucket.newValues.push(val);
      added++;
    }
  }

  // ── Phase 1a: Single-select columns (DROPDOWN_MAP) ──
  // NO comma-split — the cell value IS the value (e.g. "Unit 8", "Boston, MA")
  // NO numeric filter — single-digit values like "8" are valid unit codes
  var singleSources = [
    { sheetName: SHEETS.MEMBER_DIR, maps: DROPDOWN_MAP.MEMBER_DIR },
    { sheetName: SHEETS.GRIEVANCE_LOG, maps: DROPDOWN_MAP.GRIEVANCE_LOG }
  ];
  for (var s1 = 0; s1 < singleSources.length; s1++) {
    var sheet1 = ss.getSheetByName(singleSources[s1].sheetName);
    if (!sheet1 || sheet1.getLastRow() < 2) continue;
    var data1 = sheet1.getDataRange().getValues();
    for (var m1 = 0; m1 < singleSources[s1].maps.length; m1++) {
      var map1 = singleSources[s1].maps[m1];
      var bucket1 = initBucket_(map1.configCol);
      for (var r1 = 1; r1 < data1.length; r1++) {
        var val1 = (data1[r1][map1.col - 1] || '').toString().trim();
        addValue_(bucket1, val1);
      }
    }
  }

  // ── Phase 1b: Multi-select columns (MULTI_SELECT_COLS) ──
  // Comma-split — values like "Email, Phone" become separate entries
  // Filter pure-digit values — "1","2","3" are dropdown indices, not real data
  var multiSources = [
    { sheetName: SHEETS.MEMBER_DIR, maps: MULTI_SELECT_COLS.MEMBER_DIR },
    { sheetName: SHEETS.GRIEVANCE_LOG, maps: MULTI_SELECT_COLS.GRIEVANCE_LOG }
  ];
  for (var s2 = 0; s2 < multiSources.length; s2++) {
    var sheet2 = ss.getSheetByName(multiSources[s2].sheetName);
    if (!sheet2 || sheet2.getLastRow() < 2) continue;
    var data2 = sheet2.getDataRange().getValues();
    for (var m2 = 0; m2 < multiSources[s2].maps.length; m2++) {
      var map2 = multiSources[s2].maps[m2];
      var bucket2 = initBucket_(map2.configCol);
      for (var r2 = 1; r2 < data2.length; r2++) {
        var cellVal2 = (data2[r2][map2.col - 1] || '').toString().trim();
        if (!cellVal2) continue;
        var parts2 = cellVal2.split(',');
        for (var p2 = 0; p2 < parts2.length; p2++) {
          var val2 = parts2[p2].trim();
          // Filter pure-digit values — they're dropdown indices, not labels
          if (val2 && !/^\d+$/.test(val2)) {
            addValue_(bucket2, val2);
          }
        }
      }
    }
  }

  // ── Phase 1c: Steward sync from IS_STEWARD column ──
  // Stewards are primarily marked via IS_STEWARD=Yes, not the dropdown columns.
  // Build full names from First+Last where IS_STEWARD is truthy.
  try {
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (memberSheet && memberSheet.getLastRow() >= 2) {
      var mData = memberSheet.getDataRange().getValues();
      var stewardBucket = initBucket_(CONFIG_COLS.STEWARDS);
      for (var sr = 1; sr < mData.length; sr++) {
        var isSteward = (col_(mData[sr], MEMBER_COLS.IS_STEWARD) || '').toString().trim();
        if (isSteward.toLowerCase() === 'yes' || isSteward === '1' || isSteward.toLowerCase() === 'true') {
          var fName = (col_(mData[sr], MEMBER_COLS.FIRST_NAME) || '').toString().trim();
          var lName = (col_(mData[sr], MEMBER_COLS.LAST_NAME) || '').toString().trim();
          var fullName = (fName + ' ' + lName).trim();
          if (fullName) addValue_(stewardBucket, fullName);
        }
      }
    }
  } catch (_stErr) { log_('populateConfig steward sync', _stErr.message); }

  // ── Phase 2: Batch-write all columns ──
  for (var col in colBuckets) {
    col = parseInt(col, 10);
    var bkt = colBuckets[col];

    // Merge existing + new
    var allValues = [];
    for (var er = 0; er < allConfigData.length; er++) {
      var ev = (allConfigData[er][col - 1] || '').toString().trim();
      if (ev) allValues.push(ev);
    }
    for (var nv = 0; nv < bkt.newValues.length; nv++) {
      allValues.push(bkt.newValues[nv]);
    }

    // Deduplicate + sort
    var seen = {};
    var unique = [];
    for (var u = 0; u < allValues.length; u++) {
      if (!seen[allValues[u]]) {
        seen[allValues[u]] = true;
        unique.push(allValues[u]);
      }
    }
    unique.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });

    var rowsNeeded = Math.max(unique.length, configDataRows);
    if (rowsNeeded === 0) continue;

    var writeData = [];
    for (var w = 0; w < rowsNeeded; w++) {
      writeData.push([w < unique.length ? unique[w] : '']);
    }
    configSheet.getRange(3, col, rowsNeeded, 1).setValues(writeData);
  }

  ss.toast('Added ' + added + ' new values to Config from existing sheet data.', 'Config Sync', 5);
}

/**
 * Deduplicates and alphabetically sorts all dropdown/multi-select Config columns.
 * Preserves rows 1-2 (section/column headers). Only touches row 3+.
 * @param {Sheet} configSheet - The Config sheet
 * @private
 */
function deduplicateAndSortConfigColumns_(configSheet) {
  if (!configSheet) return;

  // Gather all Config columns that are populated by dropdown/multi-select sync
  var configColsToClean = {};
  var ddMember = DROPDOWN_MAP.MEMBER_DIR;
  var ddGriev = DROPDOWN_MAP.GRIEVANCE_LOG;
  var msMember = MULTI_SELECT_COLS.MEMBER_DIR;
  var msGriev = MULTI_SELECT_COLS.GRIEVANCE_LOG;

  var allMaps = ddMember.concat(ddGriev, msMember, msGriev);
  for (var i = 0; i < allMaps.length; i++) {
    var cc = allMaps[i].configCol;
    if (cc) {
      configColsToClean[cc] = true;
    }
  }

  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return;
  var dataRows = lastRow - 2;

  for (var colNum in configColsToClean) {
    colNum = parseInt(colNum, 10);
    var colData = configSheet.getRange(3, colNum, dataRows, 1).getValues();

    // Collect unique non-empty values
    var seen = {};
    var unique = [];
    for (var r = 0; r < colData.length; r++) {
      var v = (colData[r][0] || '').toString().trim();
      if (v && !seen[v]) {
        // Also reject pure-numeric values during cleanup
        if (/^\d+$/.test(v)) continue;
        seen[v] = true;
        unique.push(v);
      }
    }

    // Sort alphabetically (case-insensitive)
    unique.sort(function(a, b) {
      return a.toLowerCase().localeCompare(b.toLowerCase());
    });

    // Write back: sorted values + blank fill for remaining rows
    var writeData = [];
    for (var w = 0; w < dataRows; w++) {
      writeData.push([w < unique.length ? unique[w] : '']);
    }
    configSheet.getRange(3, colNum, dataRows, 1).setValues(writeData);
  }
}

/**
 * Applies consistent styling to the entire Config sheet
 * - Row 1 (Section Headers): Dark slate background
 * - Row 2 (Column Headers): Purple background
 * - Row 3+ (Data Entry): Light background for easy editing
 * Can be called from menu to restyle existing Config sheets
 * @param {Sheet} sheet - The Config sheet (optional)
 */
function applyConfigSheetStyling(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = sheet || ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    ss.toast('Config sheet not found', 'Error', 3);
    return;
  }

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastCol < 1) {
    ss.toast('Config sheet is empty', 'Error', 3);
    return;
  }

  // Ensure we have at least 50 rows for data entry
  var maxRows = Math.max(lastRow, 50);

  // ═══ ROW 1: Section Headers - Dark Slate ═══
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(SHEET_COLORS.HEADER_SLATE)  // Dark slate
    .setFontColor(SHEET_COLORS.BG_OFF_WHITE)   // Light text
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);

  // ═══ ROW 2: Column Headers - Purple ═══
  sheet.getRange(2, 1, 1, lastCol)
    .setBackground(SHEET_COLORS.STATUS_PURPLE)  // Primary purple
    .setFontColor(SHEET_COLORS.TEXT_WHITE)   // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  // ═══ ROWS 3+: Data Entry Area - Light Background ═══
  if (maxRows >= 3) {
    var dataRange = sheet.getRange(3, 1, maxRows - 2, lastCol);
    dataRange
      .setBackground(SHEET_COLORS.BG_OFF_WHITE)  // Very light gray/white
      .setFontColor(SHEET_COLORS.HEADER_SLATE)   // Dark text
      .setFontWeight('normal')
      .setVerticalAlignment('middle');

    // Apply alternating row colors (zebra stripes) — single setBackgrounds() call
    // instead of one setBackground() per row for a large performance improvement.
    var bgColors = [];
    for (var row = 3; row <= maxRows; row++) {
      var rowColor = (row % 2 === 0) ? SHEET_COLORS.BG_SLATE_LIGHT : SHEET_COLORS.BG_WHITE;
      bgColors.push(new Array(lastCol).fill(rowColor));
    }
    sheet.getRange(3, 1, maxRows - 2, lastCol).setBackgrounds(bgColors);
  }

  // Apply section-specific column colors to headers
  applySectionColors_(sheet, lastCol);

  ss.toast('Config sheet styling applied!', 'Theme Applied', 3);
  log_('applyConfigSheetStyling', 'Config sheet styling applied to ' + lastCol + ' columns');
}

/**
 * Applies section-specific colors to column headers in Config sheet
 * Each section gets a distinct color for easy identification
 * @param {Sheet} sheet - The Config sheet
 * @param {number} lastCol - Last column number
 * @private
 */
function applySectionColors_(sheet, lastCol) {
  // Section color definitions (13 sections, columns A-CC = 81 cols)
  var SECTION_COLORS = {
    ORGANIZATION:    { bg: '#22c55e', text: SHEET_COLORS.TEXT_WHITE },     // Green
    CONTACT:         { bg: '#0ea5e9', text: SHEET_COLORS.TEXT_WHITE },     // Sky
    EMPLOYMENT:      { bg: '#3b82f6', text: SHEET_COLORS.TEXT_WHITE },     // Blue
    PEOPLE:          { bg: '#8b5cf6', text: SHEET_COLORS.TEXT_WHITE },     // Violet
    GRIEVANCE:       { bg: '#ef4444', text: SHEET_COLORS.TEXT_WHITE },     // Red
    DEADLINES:       { bg: '#ec4899', text: SHEET_COLORS.TEXT_WHITE },     // Pink
    NOTIFICATIONS:   { bg: '#eab308', text: SHEET_COLORS.HEADER_SLATE },  // Yellow
    LINKS:           { bg: '#f97316', text: SHEET_COLORS.TEXT_WHITE },     // Orange
    DRIVE_CALENDAR:  { bg: '#0d9488', text: SHEET_COLORS.TEXT_WHITE },     // Teal
    SURVEY:          { bg: '#7c3aed', text: SHEET_COLORS.TEXT_WHITE },     // Violet-dark
    BRANDING:        { bg: '#10b981', text: SHEET_COLORS.TEXT_WHITE },     // Emerald
    FEATURE_TOGGLES: { bg: '#ea580c', text: SHEET_COLORS.TEXT_WHITE },     // Orange-dark
    RETENTION:       { bg: '#dc2626', text: SHEET_COLORS.TEXT_WHITE }      // Red-dark
  };

  // Apply colors by column ranges (both row 1 section header and row 2 column header)
  // Total: 81 columns (A-CC) — must match CONFIG_HEADER_MAP_ order
  var sections = [
    { start: 1,  end: 10, color: SECTION_COLORS.ORGANIZATION },     // A–J   Organization
    { start: 11, end: 16, color: SECTION_COLORS.CONTACT },          // K–P   Contact Info
    { start: 17, end: 23, color: SECTION_COLORS.EMPLOYMENT },       // Q–W   Employment
    { start: 24, end: 29, color: SECTION_COLORS.PEOPLE },           // X–AC  People
    { start: 30, end: 36, color: SECTION_COLORS.GRIEVANCE },        // AD–AJ Grievance Settings
    { start: 37, end: 44, color: SECTION_COLORS.DEADLINES },        // AK–AR Deadlines
    { start: 45, end: 47, color: SECTION_COLORS.NOTIFICATIONS },    // AS–AU Notifications
    { start: 48, end: 54, color: SECTION_COLORS.LINKS },            // AV–BB Links
    { start: 55, end: 65, color: SECTION_COLORS.DRIVE_CALENDAR },   // BC–BM Drive & Calendar
    { start: 66, end: 68, color: SECTION_COLORS.SURVEY },           // BN–BP Survey
    { start: 69, end: 74, color: SECTION_COLORS.BRANDING },         // BQ–BV Branding & UX
    { start: 75, end: 79, color: SECTION_COLORS.FEATURE_TOGGLES },  // BW–CA Feature Toggles
    { start: 80, end: 81, color: SECTION_COLORS.RETENTION }         // CB–CC Retention
  ];

  sections.forEach(function(section) {
    if (section.start <= lastCol) {
      var endCol = Math.min(section.end, lastCol);
      var colCount = endCol - section.start + 1;

      // Row 1 - Section header
      sheet.getRange(1, section.start, 1, colCount)
        .setBackground(section.color.bg)
        .setFontColor(section.color.text);

      // Row 2 - Column header (slightly lighter version)
      sheet.getRange(2, section.start, 1, colCount)
        .setBackground(section.color.bg)
        .setFontColor(section.color.text);
    }
  });
}

/**
 * Menu function to apply Config sheet styling
 * Call this from menu to restyle the Config sheet
 */
function applyConfigStyling() {
  applyConfigSheetStyling();
}

/**
 * Creates the Config Guide sheet - a dedicated tab explaining how to use the Config tab
 */
function createConfigGuideSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.CONFIG_GUIDE || '📖 Config Guide';

  // CR-11: This sheet is fully system-generated (no user data), so clearing is safe.
  var sheet = ss.getSheetByName(sheetName);
  var _isNew = !sheet;
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define guide colors
  var headerBg = COLORS.STATUS_BLUE;       // Blue header
  var sectionBg = '#E8F4FD';     // Light blue section (no exact SHEET_COLORS match)
  var tipBg = '#FFF9E6';         // Light yellow for tips (no exact SHEET_COLORS match)
  var warningBg = SHEET_COLORS.BG_LIGHT_RED;     // Light red for warnings
  var successBg = '#DCFCE7';     // Light green for success (no exact SHEET_COLORS match)
  var textColor = '#1F2937';     // Dark text (no exact SHEET_COLORS match)

  var row = 1;

  // ═══ HEADER ═══
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📖 CONFIG TAB USER GUIDE')
    .setBackground(headerBg)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  // ═══ INTRO SECTION ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('🎯 What is the Config tab for?')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(textColor);
  sheet.setRowHeight(row, 35);

  row++;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('The Config tab is the control center for your dashboard. All dropdown options throughout the system pull from these columns. When you add a value to Config, it becomes available as a dropdown option everywhere.')
    .setFontColor(textColor)
    .setWrap(true);
  sheet.setRowHeight(row, 50);

  // ═══ HOW TO USE ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📝 How to Add/Edit Dropdown Options')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(textColor);
  sheet.setRowHeight(row, 35);

  var howToSteps = [
    ['Step 1:', 'Go to the Config tab and find the column you want to modify (e.g., "Job Titles" in Column A)'],
    ['Step 2:', 'Add new values in empty cells below the existing values - NO GAPS allowed!'],
    ['Step 3:', 'The dropdown will automatically include your new values throughout the system'],
    ['Step 4:', 'To remove a value, delete the cell and shift cells up (don\'t leave blanks)']
  ];

  for (var i = 0; i < howToSteps.length; i++) {
    row++;
    sheet.getRange(row, 1).setValue(howToSteps[i][0]).setFontWeight('bold').setFontColor(COLORS.STATUS_BLUE);
    sheet.getRange(row, 2, 1, 5).merge().setValue(howToSteps[i][1]).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 30);
  }

  // ═══ COLUMN GUIDE ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📊 Config Column Quick Reference')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(textColor);
  sheet.setRowHeight(row, 35);

  row++;
  // Table header
  sheet.getRange(row, 1).setValue('Col').setFontWeight('bold').setBackground('#E5E7EB').setHorizontalAlignment('center');
  sheet.getRange(row, 2).setValue('Name').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 3, 1, 2).merge().setValue('Used In').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 5, 1, 2).merge().setValue('Example Values').setFontWeight('bold').setBackground('#E5E7EB');

  var columnData = [
    ['A', 'Organization Name', 'Dashboard-wide', 'Your Union Name'],
    ['Q', 'Job Titles', 'Member Directory', 'Case Worker, Supervisor, Manager...'],
    ['R', 'Office Locations', 'Member Dir & Grievance Log', 'Boston Office, Springfield Office...'],
    ['U', 'Units', 'Member Dir & Grievance Log', 'Unit 1, Unit 2, Unit 3...'],
    ['X', 'Supervisors', 'Member Directory', 'Names of supervisors'],
    ['Y', 'Directors', 'Member Directory', 'Names of directors'],
    ['Z', 'Stewards', 'Member Dir & Grievance Log', 'Names of union stewards'],
    ['AD', 'Grievance Status', 'Grievance Log', 'Open, Pending Info, Settled, Won...'],
    ['AE', 'Grievance Step', 'Grievance Log', 'Informal, Step I, Step II, Step III...'],
    ['AF', 'Issue Category', 'Grievance Log', 'Discipline, Workload, Pay, Benefits...'],
    ['AG', 'Articles Violated', 'Grievance Log', 'Article 12, Article 23A...']
  ];

  for (var j = 0; j < columnData.length; j++) {
    row++;
    sheet.getRange(row, 1).setValue(columnData[j][0]).setFontColor(COLORS.STATUS_BLUE).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row, 2).setValue(columnData[j][1]).setFontColor(textColor);
    sheet.getRange(row, 3, 1, 2).merge().setValue(columnData[j][2]).setFontColor(SHEET_COLORS.TEXT_GRAY);
    sheet.getRange(row, 5, 1, 2).merge().setValue(columnData[j][3]).setFontColor(SHEET_COLORS.TEXT_GRAY).setFontStyle('italic');
  }

  // ═══ TIPS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('💡 Pro Tips')
    .setBackground(tipBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(SHEET_COLORS.TEXT_DARK_ORANGE);
  sheet.setRowHeight(row, 35);

  var tips = [
    '✓ Keep dropdown lists in alphabetical order for easier selection',
    '✓ Use consistent naming conventions (e.g., "Boston Office" not "boston office")',
    '✓ Add your organization\'s specific values before entering member/grievance data',
    '✓ The system pre-fills some default values - modify them to match your organization'
  ];

  for (var k = 0; k < tips.length; k++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(tips[k]).setBackground(tipBg).setFontColor(SHEET_COLORS.TEXT_DARK_ORANGE);
  }

  // ═══ WARNINGS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('⚠️ Important Warnings')
    .setBackground(warningBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(SHEET_COLORS.STATUS_ERROR);
  sheet.setRowHeight(row, 35);

  var warnings = [
    '⚠ Do NOT delete values that are already in use in Member Directory or Grievance Log',
    '⚠ Do NOT leave blank cells in the middle of a column - this breaks the dropdowns',
    '⚠ Do NOT modify the Section Headers (Row 1) or Column Headers (Row 2) in Config'
  ];

  for (var m = 0; m < warnings.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(warnings[m]).setBackground(warningBg).setFontColor(SHEET_COLORS.STATUS_ERROR);
  }

  // ═══ NEED HELP ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('🆘 Need More Help?')
    .setBackground(successBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor(SHEET_COLORS.TEXT_GREEN_ALT);
  sheet.setRowHeight(row, 35);

  row++;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Check the "📚 Getting Started" tab for full setup instructions, or the "❓ FAQ" tab for common questions.')
    .setBackground(successBg)
    .setFontColor(SHEET_COLORS.TEXT_GREEN_ALT);

  // Set column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 6) {
    sheet.deleteColumns(7, maxCols - 6);
  }

  // Freeze header
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor(COLORS.STATUS_BLUE);

  return sheet;
}

/**
 * Adds any columns defined in MEMBER_HEADER_MAP_ that are missing from the
 * Member Directory sheet, without touching existing data.
 *
 * Pattern matches createConfigSheet: headers are always authoritative;
 * data columns are never deleted or moved.
 *
 * For each newly-added column this function also applies column-specific
 * formatting (checkboxes, conditional rules) so the sheet is fully set up
 * without requiring a manual step.
 *
 * @param {Sheet} sheet - The Member Directory sheet object
 * @returns {string[]} Array of header names that were added (empty if none)
 */
function _addMissingMemberHeaders_(sheet) {
  var allHeaders = getMemberHeaders(); // ordered list from MEMBER_HEADER_MAP_
  var lastCol = sheet.getLastColumn();

  // Read current headers — only as many columns as the sheet has
  var existingHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim().toLowerCase(); })
    : [];

  var added = [];

  for (var i = 0; i < allHeaders.length; i++) {
    var header = allHeaders[i];
    var normalised = header.toLowerCase();

    // Skip if header already present anywhere in row 1
    if (existingHeaders.indexOf(normalised) !== -1) continue;

    // Append to the next empty column
    var targetCol = lastCol + added.length + 1;
    var cell = sheet.getRange(1, targetCol);
    cell.setValue(header)
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    added.push(header);
    log_('_addMissingMemberHeaders_', 'added column "' + header + '" at col ' + targetCol);

    // ── Per-column post-setup ───────────────────────────────────────────────
    // Dues Paying: checkbox + conditional formatting
    if (normalised === 'dues paying') {
      sheet.getRange(2, targetCol, 4999, 1).insertCheckboxes();
      sheet.setColumnWidth(targetCol, 100);
      var colLetter = getColumnLetter(targetCol);
      var duesRange = sheet.getRange(2, targetCol, 4999, 1);
      var existingRules = sheet.getConditionalFormatRules();
      var duesTrueRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + colLetter + '2=TRUE')
        .setBackground(SHEET_COLORS.BG_LIGHT_GREEN).setFontColor(SHEET_COLORS.TEXT_GREEN)
        .setRanges([duesRange]).build();
      var duesFalseRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + colLetter + '2=FALSE')
        .setBackground(SHEET_COLORS.BG_CREAM).setFontColor(SHEET_COLORS.TEXT_YELLOW_DARK)
        .setRanges([duesRange]).build();
      sheet.setConditionalFormatRules(existingRules.concat([duesTrueRule, duesFalseRule]));
    }

    // Share Phone: Yes/No dropdown (steward opt-in for member phone visibility)
    if (normalised === 'share phone') {
      sheet.setColumnWidth(targetCol, 110);
      var sharePhoneMissingRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true)
        .setAllowInvalid(false)
        .setHelpText('Yes = members can see this steward\'s phone number in the directory. No = phone is hidden from members.')
        .build();
      // Seed 'No' into all existing data rows — default is opt-out.
      var lastDataRow = Math.max(sheet.getLastRow(), 1);
      if (lastDataRow > 1) {
        var noValues = [];
        for (var r = 0; r < lastDataRow - 1; r++) noValues.push(['No']);
        sheet.getRange(2, targetCol, lastDataRow - 1, 1).setValues(noValues);
      }
      sheet.getRange(2, targetCol, 4999, 1).setDataValidation(sharePhoneMissingRule);
    }
  }

  return added;
}

/**
 * Create or recreate the Member Directory sheet.
 * On new sheets: writes all headers from MEMBER_HEADER_MAP_ and applies full formatting.
 * On existing sheets: appends any headers not yet present without touching existing data.
 */
function createMemberDirectory(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEMBER_DIR);
  var headers = getMemberHeaders();

  // Ensure sheet has enough columns before any column operations.
  // When an existing sheet has data, header writing is skipped (which would
  // auto-expand columns), so the sheet may still have fewer columns than needed.
  ensureMinimumColumns(sheet, headers.length);

  // getOrCreateSheet now preserves data - only set headers on empty sheets
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold');
    // New sheet: apply per-column setup that requires knowing the column position.
    // syncColumnMaps has not run yet, so re-resolve DUES_PAYING position from the
    // header we just wrote rather than relying on the global MEMBER_COLS constant.
    var duesPayingColIdx = headers.indexOf('Dues Paying');
    if (duesPayingColIdx !== -1) {
      var dpCol = duesPayingColIdx + 1;
      sheet.getRange(2, dpCol, 4999, 1).insertCheckboxes();
      sheet.setColumnWidth(dpCol, 100);
      var dpLetter = getColumnLetter(dpCol);
      var dpRange = sheet.getRange(2, dpCol, 4999, 1);
      var dpTrueRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + dpLetter + '2=TRUE')
        .setBackground(SHEET_COLORS.BG_LIGHT_GREEN).setFontColor(SHEET_COLORS.TEXT_GREEN)
        .setRanges([dpRange]).build();
      var dpFalseRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + dpLetter + '2=FALSE')
        .setBackground(SHEET_COLORS.BG_CREAM).setFontColor(SHEET_COLORS.TEXT_YELLOW_DARK)
        .setRanges([dpRange]).build();
      // Rules array will be extended below with the rest of the conditional formatting
      // Store these for later application (setConditionalFormatRules is called once at end)
      sheet._duesCfRules = [dpTrueRule, dpFalseRule];
    }
  } else {
    // Existing sheet: rename any columns whose headers changed between versions.
    // Data is preserved — only the header cell text is updated.
    // NOTE: syncColumnMaps() also handles renames via headerRenames, but this
    // serves as a safety net for flows that don't call syncColumnMaps first.
    var HEADER_RENAMES_ = { 'Manager': 'Director' };
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var ri = 0; ri < existingHeaders.length; ri++) {
        var oldH = String(existingHeaders[ri]).trim();
        if (HEADER_RENAMES_[oldH]) {
          sheet.getRange(1, ri + 1).setValue(HEADER_RENAMES_[oldH]);
          existingHeaders[ri] = HEADER_RENAMES_[oldH];
          log_('createMemberDirectory', 'renamed header "' + oldH + '" → "' + HEADER_RENAMES_[oldH] + '" in col ' + (ri + 1));
        }
      }

      // Remove duplicate columns created by a previous backfill that ran
      // before the rename was in place (e.g. both "Manager" and "Director" existed).
      var headerCount = {};
      var dupeColsToDelete = [];
      for (var di = 0; di < existingHeaders.length; di++) {
        var hdr = String(existingHeaders[di]).trim();
        if (!hdr) continue;
        if (headerCount[hdr]) {
          // Keep the FIRST occurrence (has data), delete later duplicates (empty backfills)
          dupeColsToDelete.push(di + 1);
        } else {
          headerCount[hdr] = true;
        }
      }
      // Delete right-to-left to avoid shifting issues
      for (var dd = dupeColsToDelete.length - 1; dd >= 0; dd--) {
        sheet.deleteColumn(dupeColsToDelete[dd]);
        log_('createMemberDirectory', 'deleted duplicate column at col ' + dupeColsToDelete[dd]);
      }
    }

    // Existing sheet: append any columns from MEMBER_HEADER_MAP_ not yet present.
    // Data in existing columns is never touched.
    var added = _addMissingMemberHeaders_(sheet);
    if (added.length > 0) {
      log_('createMemberDirectory', 'added ' + added.length + ' missing column(s): ' + added.join(', '));
      SpreadsheetApp.getActive().toast(
        'Added ' + added.length + ' new column(s): ' + added.join(', '),
        '\uD83D\uDCCB Member Directory', 5
      );
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(MEMBER_COLS.MEMBER_ID, 100);
  sheet.setColumnWidth(MEMBER_COLS.FIRST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.LAST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.EMAIL, 200);
  sheet.setColumnWidth(MEMBER_COLS.CONTACT_NOTES, 250);

  // Hide Cubicle column by default
  sheet.hideColumns(MEMBER_COLS.CUBICLE);

  // Add checkbox for Start Grievance column (pre-allocate for future rows)
  sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, 4999, 1).insertCheckboxes();

  // Add checkbox for Quick Actions column (opens quick actions dialog when checked)
  sheet.getRange(2, MEMBER_COLS.QUICK_ACTIONS, 4999, 1).insertCheckboxes();

  // Dues Paying checkbox and conditional formatting are applied by _addMissingMemberHeaders_
  // when the column is first added to an existing sheet, or during new-sheet creation below.
  // MEMBER_COLS.DUES_PAYING is resolved by syncColumnMaps after headers are written.

  // Format date columns (MM/dd/yyyy)
  var dateColumns = [
    MEMBER_COLS.LAST_VIRTUAL_MTG,
    MEMBER_COLS.LAST_INPERSON_MTG,
    MEMBER_COLS.RECENT_CONTACT_DATE
  ];
  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, 998, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Format numeric columns with comma separators
  sheet.getRange(2, MEMBER_COLS.OPEN_RATE, 998, 1).setNumberFormat('#,##0.0');       // T - Open Rate %
  sheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, 998, 1).setNumberFormat('#,##0');  // U - Volunteer Hours

  // Columns AB-AD (Has Open Grievance?, Grievance Status, Days to Deadline)
  // are populated by syncGrievanceToMemberDirectory() with STATIC values
  // No formulas in visible sheets - all calculations done by script

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN GROUPS: Group and hide optional columns for cleaner view
  // ═══════════════════════════════════════════════════════════════════════════
  try {
    var maxRows = sheet.getMaxRows();
    var memberGroupRanges = [
      sheet.getRange(1, MEMBER_COLS.LAST_VIRTUAL_MTG, maxRows, 4),
      sheet.getRange(1, MEMBER_COLS.INTEREST_LOCAL, maxRows, 4),
      sheet.getRange(1, MEMBER_COLS.STREET_ADDRESS, maxRows, 4)
    ];

    // Clear any existing column groups first to prevent duplicates on re-run
    memberGroupRanges.forEach(function(range) {
      for (var d = 0; d < 8; d++) {
        try { range.shiftColumnGroupDepth(-1); } catch(_e) { break; }
      }
    });

    // Group 1: Engagement Metrics (Q-T, columns 17-20) - Hidden by default
    memberGroupRanges[0].shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    // Group 2: Member Interests (U-X, columns 21-24) - Hidden by default
    memberGroupRanges[1].shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    // Group 3: Mailing Address / PII (AK-AN, columns 37-40) - Hidden by default, PII
    memberGroupRanges[2].shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    log_('Member Directory column group setup skipped', e.toString());
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // EMPLOYMENT & PII COLUMN SETUP: Widths, date format, and hide PIN Hash
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.setColumnWidth(MEMBER_COLS.EMPLOYEE_ID, 120);
  sheet.setColumnWidth(MEMBER_COLS.DEPARTMENT, 140);
  sheet.setColumnWidth(MEMBER_COLS.HIRE_DATE, 110);
  sheet.setColumnWidth(MEMBER_COLS.STREET_ADDRESS, 200);
  sheet.setColumnWidth(MEMBER_COLS.CITY, 120);
  sheet.setColumnWidth(MEMBER_COLS.STATE, 80);

  // Format Hire Date column as date
  sheet.getRange(2, MEMBER_COLS.HIRE_DATE, 998, 1).setNumberFormat('MM/dd/yyyy');

  // Hide PIN Hash column (sensitive data — should not be visible in the sheet)
  sheet.hideColumns(MEMBER_COLS.PIN_HASH, 1);

  // Share Phone column — Yes/No dropdown (steward opt-in for member phone visibility)
  if (MEMBER_COLS.SHARE_PHONE) {
    sheet.setColumnWidth(MEMBER_COLS.SHARE_PHONE, 110);
    var sharePhoneRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .setHelpText('Yes = members can see this steward\'s phone number in the directory. No = phone is hidden from members.')
      .build();
    sheet.getRange(2, MEMBER_COLS.SHARE_PHONE, 4999, 1).setDataValidation(sharePhoneRule);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CONDITIONAL FORMATTING: Highlight members with open grievances
  // ═══════════════════════════════════════════════════════════════════════════
  var _lastRow = Math.max(sheet.getLastRow(), 2);
  var hasOpenGrievanceRange = sheet.getRange(2, MEMBER_COLS.HAS_OPEN_GRIEVANCE, 4999, 1);

  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#ffebee')  // Light red background (no exact SHEET_COLORS match)
    .setFontColor(SHEET_COLORS.TEXT_RED)   // Dark red text
    .setBold(true)
    .setRanges([hasOpenGrievanceRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // VALIDATION HIGHLIGHTING: Red background for empty Email and Phone fields
  // ═══════════════════════════════════════════════════════════════════════════
  var emailRange = sheet.getRange(2, MEMBER_COLS.EMAIL, 4999, 1);
  var phoneRange = sheet.getRange(2, MEMBER_COLS.PHONE, 4999, 1);

  // Dynamic column letters from constants
  var colId = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var colEmail = getColumnLetter(MEMBER_COLS.EMAIL);
  var colPhone = getColumnLetter(MEMBER_COLS.PHONE);
  var colDeadline = getColumnLetter(MEMBER_COLS.NEXT_DEADLINE);

  // Rule: Red background for empty Email
  var emptyEmailRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + colId + '2<>"",ISBLANK($' + colEmail + '2))')
    .setBackground(SHEET_COLORS.BG_LIGHT_RED_ALT)
    .setRanges([emailRange])
    .build();

  // Rule: Red background for empty Phone
  var emptyPhoneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + colId + '2<>"",ISBLANK($' + colPhone + '2))')
    .setBackground(SHEET_COLORS.BG_LIGHT_RED_ALT)
    .setRanges([phoneRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // DEADLINE HEATMAP: Color-coded Days to Deadline
  // ═══════════════════════════════════════════════════════════════════════════
  var daysDeadlineRange = sheet.getRange(2, MEMBER_COLS.NEXT_DEADLINE, 4999, 1);

  // Rule: Red - Overdue (shows "Overdue" or negative/0 days)
  var deadlineOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + colDeadline + '2="Overdue",AND(ISNUMBER($' + colDeadline + '2),$' + colDeadline + '2<=0))')
    .setBackground('#ffebee')
    .setFontColor(SHEET_COLORS.TEXT_RED)
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Orange - Due in 1-3 days
  var deadline1to3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + colDeadline + '2),$' + colDeadline + '2>=1,$' + colDeadline + '2<=3)')
    .setBackground(SHEET_COLORS.BG_WARM)
    .setFontColor(SHEET_COLORS.TEXT_ORANGE)
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Yellow - Due in 4-7 days
  var deadline4to7Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + colDeadline + '2),$' + colDeadline + '2>=4,$' + colDeadline + '2<=7)')
    .setBackground(SHEET_COLORS.BG_EXTRA_PALE_YELLOW)
    .setFontColor(SHEET_COLORS.TEXT_YELLOW_DARK)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Green - On Track (more than 7 days remaining)
  var deadlineOnTrackRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + colDeadline + '2),$' + colDeadline + '2>7)')
    .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
    .setFontColor(SHEET_COLORS.TEXT_GREEN)
    .setRanges([daysDeadlineRange])
    .build();

  var rules = [redRule, emptyEmailRule, emptyPhoneRule, deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule];

  // Dues Paying conditional formatting — green = paying, amber = not paying
  // New sheets: rules built inline above (stored in sheet._duesCfRules) because
  //   MEMBER_COLS.DUES_PAYING isn't resolved until syncColumnMaps runs.
  // Existing sheets: MEMBER_COLS.DUES_PAYING is already resolved from the live sheet.
  if (sheet._duesCfRules) {
    rules = rules.concat(sheet._duesCfRules);
    delete sheet._duesCfRules;
  } else if (MEMBER_COLS.DUES_PAYING) {
    var duesPayingRange = sheet.getRange(2, MEMBER_COLS.DUES_PAYING, 4999, 1);
    var duesPayingTrueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + getColumnLetter(MEMBER_COLS.DUES_PAYING) + '2=TRUE')
      .setBackground(SHEET_COLORS.BG_LIGHT_GREEN).setFontColor(SHEET_COLORS.TEXT_GREEN)
      .setRanges([duesPayingRange]).build();
    var duesPayingFalseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + getColumnLetter(MEMBER_COLS.DUES_PAYING) + '2=FALSE')
      .setBackground(SHEET_COLORS.BG_CREAM).setFontColor(SHEET_COLORS.TEXT_YELLOW_DARK)
      .setRanges([duesPayingRange]).build();
    rules = rules.concat([duesPayingTrueRule, duesPayingFalseRule]);
  }

  sheet.setConditionalFormatRules(rules);

  // ═══════════════════════════════════════════════════════════════════════════
  // FILTER: Enable sorting on all columns via filter dropdown
  // ═══════════════════════════════════════════════════════════════════════════
  // Remove existing filter if any
  var existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Create filter on entire data range (all columns)
  // This enables sorting via dropdown on: Last Name, Job Title, Work Location, Unit,
  // Office Days, Preferred Communication, Best Time to Contact, Supervisor, Manager,
  // Committees, Assigned Steward, Last Virtual Mtg, Last In-Person Mtg, Open Rate %,
  // Volunteer Hours, Interest: Local/Chapter/Allied, Recent Contact Date,
  // Contact Steward, Contact Notes, Has Open Grievance?, Grievance Status, Days to Deadline
  var filterRange = sheet.getRange(1, 1, 5000, headers.length);
  filterRange.createFilter();

  // Set tab color
  sheet.setTabColor(COLORS.UNION_GREEN);
}

/**
 * Adds any columns defined in GRIEVANCE_HEADER_MAP_ that are missing from the
 * Grievance Log sheet, without touching existing data.
 *
 * Parallel to _addMissingMemberHeaders_ — same pattern, same rules:
 *  - Header name matching only (case-insensitive), never by column index
 *  - Data in existing columns is never deleted, moved, or overwritten
 *  - Per-column post-setup block handles column-specific formatting on first-add
 *
 * @param {Sheet} sheet - The Grievance Log sheet object
 * @returns {string[]} Array of header names that were added (empty if none)
 */
function _addMissingGrievanceHeaders_(sheet) {
  var allHeaders = getGrievanceHeaders(); // ordered list from GRIEVANCE_HEADER_MAP_
  var lastCol = sheet.getLastColumn();

  var existingHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim().toLowerCase(); })
    : [];

  var added = [];

  for (var i = 0; i < allHeaders.length; i++) {
    var header = allHeaders[i];
    var normalised = header.toLowerCase();

    if (existingHeaders.indexOf(normalised) !== -1) continue;

    var targetCol = lastCol + added.length + 1;
    var cell = sheet.getRange(1, targetCol);
    cell.setValue(header)
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    added.push(header);
    log_('_addMissingGrievanceHeaders_', 'added column "' + header + '" at col ' + targetCol);

    // ── Per-column post-setup ───────────────────────────────────────────────
    // Add cases here for any Grievance column that needs special formatting on first-add.
    // Example: if ('message alert' === normalised) { sheet.getRange(2,targetCol,4999,1).insertCheckboxes(); }
    // (MESSAGE_ALERT and QUICK_ACTIONS already exist in all live sheets so no case needed today)
  }

  return added;
}

/**
 * Create or recreate the Grievance Log sheet
 * NOTE: Calculated columns (First Name, Last Name, Email, Deadlines, Days Open, etc.)
 * are managed by the hidden _Grievance_Formulas sheet for self-healing capability.
 * Users can't accidentally erase formulas because they're in the hidden sheet.
 */
function createGrievanceLog(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.GRIEVANCE_LOG);
  var headers = getGrievanceHeaders();

  // Ensure sheet has enough columns before any column operations.
  // When an existing sheet has data, header writing is skipped (which would
  // auto-expand columns), so the sheet may still have fewer columns than needed.
  ensureMinimumColumns(sheet, headers.length);

  // getOrCreateSheet now preserves data - only set headers on empty sheets
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold');
  } else {
    // Existing sheet: append any columns from GRIEVANCE_HEADER_MAP_ not yet present.
    // Data in existing columns is never touched.
    var added = _addMissingGrievanceHeaders_(sheet);
    if (added.length > 0) {
      log_('createGrievanceLog', 'added ' + added.length + ' missing column(s): ' + added.join(', '));
      SpreadsheetApp.getActive().toast(
        'Added ' + added.length + ' new column(s): ' + added.join(', '),
        '\uD83D\uDCCB Grievance Log', 5
      );
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(GRIEVANCE_COLS.GRIEVANCE_ID, 100);
  sheet.setColumnWidth(GRIEVANCE_COLS.RESOLUTION, 250);
  sheet.setColumnWidth(GRIEVANCE_COLS.COORDINATOR_MESSAGE, 250);

  // Add checkbox for Message Alert column (pre-allocate for future rows)
  sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, 4999, 1).insertCheckboxes();

  // Add checkbox for Quick Actions column (opens quick actions dialog when checked)
  sheet.getRange(2, GRIEVANCE_COLS.QUICK_ACTIONS, 4999, 1).insertCheckboxes();

  // Format date columns
  var dateColumns = [
    GRIEVANCE_COLS.INCIDENT_DATE,
    GRIEVANCE_COLS.FILING_DEADLINE,
    GRIEVANCE_COLS.DATE_FILED,
    GRIEVANCE_COLS.STEP1_DUE,
    GRIEVANCE_COLS.STEP1_RCVD,
    GRIEVANCE_COLS.STEP2_APPEAL_DUE,
    GRIEVANCE_COLS.STEP2_APPEAL_FILED,
    GRIEVANCE_COLS.STEP2_DUE,
    GRIEVANCE_COLS.STEP2_RCVD,
    GRIEVANCE_COLS.STEP3_APPEAL_DUE,
    GRIEVANCE_COLS.STEP3_APPEAL_FILED,
    GRIEVANCE_COLS.DATE_CLOSED,
    GRIEVANCE_COLS.NEXT_ACTION_DUE,
    GRIEVANCE_COLS.LAST_UPDATED
  ];

  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, 998, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Format Days Open (S) and Days to Deadline (U) as whole numbers with comma separators
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, 998, 1).setNumberFormat('#,##0');
  // Days to Deadline can show "Overdue" text, so use General format that handles both
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 998, 1).setNumberFormat('#,##0');

  // Format Last Updated (AP) as date-time
  sheet.getRange(2, GRIEVANCE_COLS.LAST_UPDATED, 998, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // Setup column groups for timeline (Step I, II, III collapsible)
  try {
    var grMaxRows = sheet.getMaxRows();
    var grievanceGroupRanges = [
      sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, grMaxRows, 2),
      sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, grMaxRows, 4),
      sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, grMaxRows, 2),
      sheet.getRange(1, GRIEVANCE_COLS.MESSAGE_ALERT, grMaxRows, 4)
    ];

    // Clear any existing column groups first to prevent duplicates on re-run
    grievanceGroupRanges.forEach(function(range) {
      for (var d = 0; d < 8; d++) {
        try { range.shiftColumnGroupDepth(-1); } catch(_e) { break; }
      }
    });

    grievanceGroupRanges[0].shiftColumnGroupDepth(1);
    grievanceGroupRanges[1].shiftColumnGroupDepth(1);
    grievanceGroupRanges[2].shiftColumnGroupDepth(1);
    // Group Coordinator columns AC-AF (Message Alert, Coordinator Message, Acknowledged By, Acknowledged Date)
    grievanceGroupRanges[3].shiftColumnGroupDepth(1);
    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
    // Collapse all groups by default (including coordinator columns AC-AF)
    sheet.collapseAllColumnGroups();
    // Hide Drive Folder ID column (AG) - internal use only
    sheet.hideColumns(GRIEVANCE_COLS.DRIVE_FOLDER_ID, 1);
  } catch (e) {
    log_('Column group setup skipped', e.toString());
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DAYS TO DEADLINE HEATMAP
  // ═══════════════════════════════════════════════════════════════════════════
  var daysDeadlineRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 4999, 1);

  // Dynamic column letters from constants
  var grColId = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var grColStatus = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var grColStep = getColumnLetter(GRIEVANCE_COLS.CURRENT_STEP);
  var grColDaysDeadline = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);

  // Rule: Red - Overdue (shows "Overdue" or negative/0 days)
  var deadlineOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + grColDaysDeadline + '2="Overdue",AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2<=0))')
    .setBackground('#ffebee')
    .setFontColor(SHEET_COLORS.TEXT_RED)
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Orange - Due in 1-3 days
  var deadline1to3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>=1,$' + grColDaysDeadline + '2<=3)')
    .setBackground(SHEET_COLORS.BG_WARM)
    .setFontColor(SHEET_COLORS.TEXT_ORANGE)
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Yellow - Due in 4-7 days
  var deadline4to7Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>=4,$' + grColDaysDeadline + '2<=7)')
    .setBackground(SHEET_COLORS.BG_EXTRA_PALE_YELLOW)
    .setFontColor(SHEET_COLORS.TEXT_YELLOW_DARK)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Green - On Track (more than 7 days remaining)
  var deadlineOnTrackRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>7)')
    .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
    .setFontColor(SHEET_COLORS.TEXT_GREEN)
    .setRanges([daysDeadlineRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // PROGRESS BAR: Colored backgrounds showing grievance stage
  // Based on Current Step, highlights completed stages
  // ═══════════════════════════════════════════════════════════════════════════

  // Progress bar spans: Step I, Step II, Step III, Date Closed
  var step1Range = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 2);
  var allStepsRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 9);

  // Completed cases: All columns green (Closed, Won, Denied, Settled, Withdrawn)
  var completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + grColStatus + '2="Closed",$' + grColStatus + '2="Won",$' + grColStatus + '2="Denied",$' + grColStatus + '2="Settled",$' + grColStatus + '2="Withdrawn")')
    .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
    .setRanges([allStepsRange])
    .build();

  // Step III in progress
  var step3ProgressRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 8);
  var step3ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + grColStep + '2="Step III"')
    .setBackground(SHEET_COLORS.BG_LIGHT_BLUE_ALT)
    .setRanges([step3ProgressRange])
    .build();

  // Step II in progress
  var step2ProgressRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 6);
  var step2ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + grColStep + '2="Step II"')
    .setBackground(SHEET_COLORS.BG_LIGHT_BLUE_ALT)
    .setRanges([step2ProgressRange])
    .build();

  // Step I in progress
  var step1ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + grColStep + '2="Step I"')
    .setBackground(SHEET_COLORS.BG_LIGHT_BLUE_ALT)
    .setRanges([step1Range])
    .build();

  // Gray out columns not yet reached
  var notReachedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + grColId + '2<>"",$' + grColStep + '2<>"")')
    .setBackground('#fafafa')
    .setRanges([allStepsRange])
    .build();

  // Apply all rules (order matters - more specific rules first)
  var rules = [
    deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule,
    completedRule, step3ProgressRule, step2ProgressRule, step1ProgressRule, notReachedRule
  ];
  sheet.setConditionalFormatRules(rules);

  // Set tab color
  sheet.setTabColor(COLORS.SOLIDARITY_RED);
}
// ============================================================================
// CHART FUNCTIONS - MOVED TO 08d_ChartBuilder.gs
// ============================================================================
// The following functions have been moved to 08d_ChartBuilder.gs:
// - generateSelectedChart()
// - createGaugeStyleChart_()
// - createScorecardChart_()
// - createTrendLineChart_()
// - createAreaChart_()
// - createComboChart_()
// - createSummaryTableChart_()
// - createStewardLeaderboardChart_()
// - padRight()
// ============================================================================
// ============================================================================
// VOLUNTEER HOURS & MEETING ATTENDANCE SHEETS
// ============================================================================

/**
 * Creates the Volunteer Hours tracking sheet
 * Records volunteer activities and auto-calculates totals for Member Directory
 * @param {Spreadsheet} ss - The spreadsheet
 * @returns {void}
 */
function createVolunteerHoursSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.VOLUNTEER_HOURS);

  var headers = [
    'Entry ID',           // A - Auto-generated
    'Member ID',          // B - Dropdown from Member Directory
    'Member Name',        // C - Auto-lookup from Member Directory
    'Activity Date',      // D - Date of volunteer activity
    'Activity Type',      // E - Type of volunteer work
    'Hours',              // F - Number of hours volunteered
    'Description',        // G - Brief description
    'Verified By',        // H - Who verified the hours
    'Notes'               // I - Additional notes
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center')
    .setWrap(false);

  // Add data type hints (row 2)
  var hints = [
    'Auto-ID', 'Text', 'Auto-lookup', 'MM/DD/YYYY', 'Dropdown', 'Number', 'Text', 'Text', 'Text'
  ];

  sheet.getRange(2, 1, 1, hints.length).setValues([hints])
    .setFontStyle('italic')
    .setFontSize(9)
    .setBackground(SHEET_COLORS.BG_PALE_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 100);  // A - Entry ID
  sheet.setColumnWidth(2, 100);  // B - Member ID
  sheet.setColumnWidth(3, 150);  // C - Member Name
  sheet.setColumnWidth(4, 110);  // D - Activity Date
  sheet.setColumnWidth(5, 150);  // E - Activity Type
  sheet.setColumnWidth(6, 80);   // F - Hours
  sheet.setColumnWidth(7, 250);  // G - Description
  sheet.setColumnWidth(8, 130);  // H - Verified By
  sheet.setColumnWidth(9, 200);  // I - Notes

  // Format columns
  sheet.getRange(3, 4, 998, 1).setNumberFormat('MM/DD/YYYY');  // D - Activity Date
  sheet.getRange(3, 6, 998, 1).setNumberFormat('#,##0.0');     // F - Hours

  // Auto-ID formula for Entry ID (column A)
  var idFormula = '=IF(B3<>"", "VOL-" & TEXT(ROW()-2, "0000"), "")';
  sheet.getRange('A3').setFormula(idFormula);
  sheet.getRange('A3').copyTo(sheet.getRange('A3:A1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Member Name lookup formula (column C) - VLOOKUP from Member Directory
  // Use dynamic column letters/indices so column reordering doesn't break lookups.
  var mIdColLetter = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mLastColLetter = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var mRange = "'" + SHEETS.MEMBER_DIR + "'!" + mIdColLetter + ":" + mLastColLetter;
  var fnIdx = MEMBER_COLS.FIRST_NAME - MEMBER_COLS.MEMBER_ID + 1;
  var lnIdx = MEMBER_COLS.LAST_NAME - MEMBER_COLS.MEMBER_ID + 1;
  var nameLookupFormula = '=IF(B3<>"", IFERROR(VLOOKUP(B3, ' + mRange + ', ' + fnIdx + ', FALSE) & " " & VLOOKUP(B3, ' + mRange + ', ' + lnIdx + ', FALSE), "Not Found"), "")';
  sheet.getRange('C3').setFormula(nameLookupFormula);
  sheet.getRange('C3').copyTo(sheet.getRange('C3:C1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Freeze header rows
  sheet.setFrozenRows(2);

  // Set tab color
  sheet.setTabColor('#8B5CF6');  // Purple for volunteer hours

  log_('createVolunteerHoursSheet', 'Volunteer Hours sheet created');
}

/**
 * Creates the Meeting Attendance tracking sheet
 * Records meeting attendance and auto-updates Member Directory
 * @param {Spreadsheet} ss - The spreadsheet
 * @returns {void}
 */
function createMeetingAttendanceSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEETING_ATTENDANCE);

  var headers = [
    'Entry ID',           // A - Auto-generated
    'Meeting Date',       // B - Date of meeting
    'Meeting Type',       // C - Virtual or In-Person
    'Meeting Name',       // D - Name/description of meeting
    'Member ID',          // E - Dropdown from Member Directory
    'Member Name',        // F - Auto-lookup from Member Directory
    'Attended',           // G - Yes/No checkbox
    'Notes'               // H - Additional notes
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center')
    .setWrap(false);

  // Add data type hints (row 2)
  var hints = [
    'Auto-ID', 'MM/DD/YYYY', 'Dropdown', 'Text', 'Text', 'Auto-lookup', 'Checkbox', 'Text'
  ];

  sheet.getRange(2, 1, 1, hints.length).setValues([hints])
    .setFontStyle('italic')
    .setFontSize(9)
    .setBackground(SHEET_COLORS.BG_PALE_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 100);  // A - Entry ID
  sheet.setColumnWidth(2, 110);  // B - Meeting Date
  sheet.setColumnWidth(3, 120);  // C - Meeting Type
  sheet.setColumnWidth(4, 200);  // D - Meeting Name
  sheet.setColumnWidth(5, 100);  // E - Member ID
  sheet.setColumnWidth(6, 150);  // F - Member Name
  sheet.setColumnWidth(7, 90);   // G - Attended
  sheet.setColumnWidth(8, 200);  // H - Notes

  // Format columns
  sheet.getRange(3, 2, 998, 1).setNumberFormat('MM/DD/YYYY');  // B - Meeting Date

  // Meeting Type dropdown (column C)
  var meetingTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Virtual', 'In-Person'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C3:C1000').setDataValidation(meetingTypeRule);

  // Add checkboxes for Attended column
  sheet.getRange('G3:G1000').insertCheckboxes();

  // Auto-ID formula for Entry ID (column A)
  var idFormula = '=IF(E3<>"", "MTG-" & TEXT(ROW()-2, "0000"), "")';
  sheet.getRange('A3').setFormula(idFormula);
  sheet.getRange('A3').copyTo(sheet.getRange('A3:A1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Member Name lookup formula (column F) - VLOOKUP from Member Directory
  // Use dynamic column letters/indices so column reordering doesn't break lookups.
  var mIdColLetter = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mLastColLetter = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var mRange = "'" + SHEETS.MEMBER_DIR + "'!" + mIdColLetter + ":" + mLastColLetter;
  var fnIdx = MEMBER_COLS.FIRST_NAME - MEMBER_COLS.MEMBER_ID + 1;
  var lnIdx = MEMBER_COLS.LAST_NAME - MEMBER_COLS.MEMBER_ID + 1;
  var nameLookupFormula = '=IF(E3<>"", IFERROR(VLOOKUP(E3, ' + mRange + ', ' + fnIdx + ', FALSE) & " " & VLOOKUP(E3, ' + mRange + ', ' + lnIdx + ', FALSE), "Not Found"), "")';
  sheet.getRange('F3').setFormula(nameLookupFormula);
  sheet.getRange('F3').copyTo(sheet.getRange('F3:F1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Freeze header rows
  sheet.setFrozenRows(2);

  // Set tab color
  sheet.setTabColor('#10B981');  // Green for meeting attendance

  log_('createMeetingAttendanceSheet', 'Meeting Attendance sheet created');
}

/**
 * Create the Meeting Check-In Log sheet
 * Stewards create meetings here; members check in via modal with email + PIN
 * @param {Spreadsheet} ss - The active spreadsheet
 */
function createMeetingCheckInLogSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEETING_CHECKIN_LOG);

  // Headers — auto-derived from MEETING_CHECKIN_HEADER_MAP_
  var headers = getHeadersFromMap_(MEETING_CHECKIN_HEADER_MAP_);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center')
    .setWrap(false);

  // Set column widths
  sheet.setColumnWidth(1, 130);  // A - Meeting ID
  sheet.setColumnWidth(2, 220);  // B - Meeting Name
  sheet.setColumnWidth(3, 120);  // C - Meeting Date
  sheet.setColumnWidth(4, 120);  // D - Meeting Type
  sheet.setColumnWidth(5, 110);  // E - Member ID
  sheet.setColumnWidth(6, 170);  // F - Member Name
  sheet.setColumnWidth(7, 170);  // G - Check-In Time
  sheet.setColumnWidth(8, 200);  // H - Email
  sheet.setColumnWidth(9, 100);  // I - Start Time
  sheet.setColumnWidth(10, 110); // J - Duration
  sheet.setColumnWidth(11, 110); // K - Event Status
  sheet.setColumnWidth(12, 220); // L - Notify Stewards
  sheet.setColumnWidth(13, 150); // M - Calendar Event ID
  sheet.setColumnWidth(14, 250); // N - Notes Doc URL
  sheet.setColumnWidth(15, 250); // O - Agenda Doc URL
  sheet.setColumnWidth(16, 250); // P - Agenda Stewards

  // Format date columns — column numbers from MEETING_CHECKIN_COLS
  sheet.getRange(2, MEETING_CHECKIN_COLS.MEETING_DATE, 999, 1).setNumberFormat('MM/DD/YYYY');
  sheet.getRange(2, MEETING_CHECKIN_COLS.CHECKIN_TIME, 999, 1).setNumberFormat('MM/DD/YYYY HH:mm:ss');
  sheet.getRange(2, MEETING_CHECKIN_COLS.MEETING_TIME, 999, 1).setNumberFormat('HH:mm');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor('#8B5CF6');  // Purple for check-in

  log_('createMeetingCheckInLogSheet', 'Meeting Check-In Log sheet created');
}

/**
 * Menu-callable wrapper to create the Meeting Check-In Log sheet
 */
function setupMeetingCheckInSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createMeetingCheckInLogSheet(ss);
  ss.toast('Meeting Check-In Log sheet created', '📝 Meeting Check-In', 5);
}

// ============================================================================
// GRIEVANCE ARCHIVE SHEET (v4.30.0)
// ============================================================================

/**
 * Ensures the _Archive_Grievances sheet exists. Copies headers from
 * Grievance Log and sets the sheet to "very hidden" (only visible via
 * Apps Script, not the Sheets UI).
 *
 * @param {Spreadsheet} [ss] - Optional spreadsheet reference
 * @returns {Sheet} The archive sheet
 */
function ensureGrievanceArchiveSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var archive = ss.getSheetByName(SHEETS.GRIEVANCE_ARCHIVE);
  if (archive) return archive;

  archive = ss.insertSheet(SHEETS.GRIEVANCE_ARCHIVE);
  // Copy header row from active Grievance Log
  var headers = getGrievanceHeaders();
  if (headers && headers.length > 0) {
    archive.getRange(1, 1, 1, headers.length).setValues([headers]);
    archive.setFrozenRows(1);
  }
  // Very hidden — only accessible via script
  setSheetVeryHidden_(archive);
  log_('Created grievance archive sheet', SHEETS.GRIEVANCE_ARCHIVE);
  return archive;
}

/**
 * Ensures the Non-Member Contacts sheet exists. Creates with headers if missing.
 * @param {Spreadsheet} [ss] - Spreadsheet instance (defaults to active)
 * @returns {Sheet} The Non-Member Contacts sheet
 */
function ensureNonMemberContactsSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS) || ss.getSheetByName('Non member contacts');
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEETS.NON_MEMBER_CONTACTS);
  var headers = getNMCHeaders();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // Category validation dropdown
  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Management', 'Legal', 'HR', 'Union Rep', 'Ally', 'Other'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NMC_COLS.CATEGORY, 500, 1).setDataValidation(catRule);

  // Shirt Size validation dropdown
  var shirtRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NMC_COLS.SHIRT_SIZE, 500, 1).setDataValidation(shirtRule);

  // Steward Yes/No validation dropdown
  var stewardRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NMC_COLS.IS_STEWARD, 500, 1).setDataValidation(stewardRule);

  log_('Created Non-Member Contacts sheet', SHEETS.NON_MEMBER_CONTACTS);
  return sheet;
}
