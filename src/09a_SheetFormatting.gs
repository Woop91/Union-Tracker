/**
 * ============================================================================
 * 09a_SheetFormatting.gs — Union Brand Theme & Tab Formatting
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Applies a consistent Union brand visual theme across all visible
 *   spreadsheet tabs. Includes reusable formatting helpers, individual
 *   per-tab formatters (20 tabs), tab-bar color grouping, and a master
 *   applyUnionThemeToAllTabs() function callable from the Admin menu.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Before this file, each sheet creation function applied its own ad-hoc
 *   colors. This centralises the Union brand palette so a single menu click
 *   re-themes every visible tab. Formatters are idempotent and non-destructive
 *   — they never insert or delete rows, only restyle existing structure.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   No data loss. Sheets keep their current (pre-theme) formatting.
 *   The "Apply Union Theme" menu item fails but nothing else is affected.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, SHEET_COLORS)
 *   Used by:    03_UIComponents.gs (Admin menu item)
 *
 * @fileoverview Union brand theme and tab formatting
 * @requires 01_Core.gs
 */

// ============================================================================
// TAB COLOR ASSIGNMENTS — which sheet gets which tab-bar color
// ============================================================================

/**
 * Map of sheet names to tab-bar color groups.
 * Blue  = data entry, Green = engagement/survey,
 * Gold  = documentation/guide, Red = admin/technical.
 * @private
 */
var TAB_COLOR_MAP_ = {};
// Populated lazily by applyTabBarColors_() because SHEETS may not yet be
// initialised when this file first loads in the GAS V8 runtime.

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Safely gets a sheet by name, returns null if not found.
 * @param {Spreadsheet} ss
 * @param {string} name
 * @returns {Sheet|null}
 * @private
 */
function getSheetSafe_(ss, name) {
  if (!name) return null;
  return ss.getSheetByName(name) || null;
}

/**
 * Applies the Union brand header style to row 1 of a sheet.
 * Navy background, white bold text, 14pt, centered, 40px tall.
 * @param {Sheet} sheet
 * @param {number} numCols - Number of columns to span
 * @private
 */
function applyBrandHeader_(sheet, numCols) {
  if (numCols < 1) return;
  var range = sheet.getRange(1, 1, 1, numCols);
  range
    .setBackground(SHEET_COLORS.THEME_NAVY)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
}

/**
 * Applies a subtitle/instruction row (italic gray, 10pt).
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @param {string} text - Subtitle text to set (if falsy, only styles existing text)
 * @private
 */
function applySubtitleRow_(sheet, row, numCols, text) {
  if (numCols < 1) return;
  var range = sheet.getRange(row, 1, 1, numCols);
  if (text) {
    sheet.getRange(row, 1, 1, 1).setValue(text);
  }
  range
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setFontStyle('italic')
    .setFontSize(10)
    .setBackground(SHEET_COLORS.BG_WHITE);
}

/**
 * Applies alternating row banding (Light Sky / white) to a data range.
 * Uses a single setBackgrounds() call for performance.
 * @param {Sheet} sheet
 * @param {number} startRow - First data row
 * @param {number} numCols
 * @param {number} [endRow] - Last row (defaults to sheet lastRow or startRow+50)
 * @private
 */
function applyRowBanding_(sheet, startRow, numCols, endRow) {
  var lastRow = endRow || Math.max(sheet.getLastRow(), startRow + 50);
  var rowCount = lastRow - startRow + 1;
  if (rowCount < 1 || numCols < 1) return;

  var bgColors = [];
  for (var r = 0; r < rowCount; r++) {
    var color = (r % 2 === 0) ? SHEET_COLORS.BG_WHITE : SHEET_COLORS.THEME_LIGHT_SKY;
    bgColors.push(new Array(numCols).fill(color));
  }
  sheet.getRange(startRow, 1, rowCount, numCols).setBackgrounds(bgColors);
}

/**
 * Applies a section divider row (full-width colored bar with centered text).
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @param {string} bgColor
 * @private
 */
function applySectionDivider_(sheet, row, numCols, bgColor) {
  if (numCols < 1) return;
  sheet.getRange(row, 1, 1, numCols)
    .setBackground(bgColor)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);
}

/**
 * Sets an empty-state placeholder message in italic gray.
 * Only writes if the sheet has no data below the header rows.
 * @param {Sheet} sheet
 * @param {number} dataStartRow - First expected data row
 * @param {number} numCols
 * @param {string} message
 * @private
 */
function applyEmptyState_(sheet, dataStartRow, numCols, message) {
  if (sheet.getLastRow() < dataStartRow) {
    sheet.getRange(dataStartRow, 1, 1, numCols).merge()
      .setValue(message)
      .setFontColor(SHEET_COLORS.TEXT_LIGHT_GRAY)
      .setFontStyle('italic')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setBackground(SHEET_COLORS.BG_WHITE);
  }
}

/**
 * Applies a secondary header row style (Steel Blue bg, white bold text).
 * Useful for column header rows below a banner.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @private
 */
function applyColumnHeaderRow_(sheet, row, numCols) {
  if (numCols < 1) return;
  sheet.getRange(row, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_STEEL_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 32);
}

/**
 * Clears existing conditional formatting rules for a sheet.
 * Called before adding new theme rules to avoid duplicates.
 * @param {Sheet} sheet
 * @private
 */
function clearConditionalFormats_(sheet) {
  sheet.setConditionalFormatRules([]);
}

/**
 * Standard column width presets.
 * @param {Sheet} sheet
 * @param {Object<number,number>} widthMap - {colNumber: widthPx}
 * @private
 */
function applyColumnWidths_(sheet, widthMap) {
  var keys = Object.keys(widthMap);
  for (var i = 0; i < keys.length; i++) {
    var col = parseInt(keys[i], 10);
    if (col > 0 && col <= sheet.getMaxColumns()) {
      sheet.setColumnWidth(col, widthMap[keys[i]]);
    }
  }
}

// ============================================================================
// INDIVIDUAL TAB FORMATTERS
// ============================================================================

// ─── Tab 1: Getting Started ─────────────────────────────────────────────────

function formatGettingStartedTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.GETTING_STARTED);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 7);

  // Row 1 branded header
  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Quick Navigation panel in columns H-J
  var navCol = 8; // column H
  var maxCols = sheet.getMaxColumns();
  if (maxCols < navCol + 2) {
    sheet.insertColumnsAfter(maxCols, (navCol + 2) - maxCols);
  }
  numCols = Math.max(numCols, navCol + 2);
  applyColumnWidths_(sheet, { 8: 160, 9: 160, 10: 160 });

  // Panel title
  sheet.getRange(2, navCol, 1, 3).merge()
    .setValue('QUICK NAVIGATION')
    .setBackground(SHEET_COLORS.THEME_NAVY)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  // Category groups with tab names
  var navGroups = [
    { label: 'CORE DATA', color: '#6A1B9A', tabs: [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG, SHEETS.CASE_CHECKLIST] },
    { label: 'REFERENCE', color: SHEET_COLORS.THEME_GOLD, tabs: [SHEETS.FAQ, SHEETS.RESOURCES, SHEETS.FEATURES_REFERENCE] },
    { label: 'ENGAGEMENT', color: SHEET_COLORS.TAB_GREEN, tabs: [SHEETS.MEETING_ATTENDANCE, SHEETS.FEEDBACK, SHEETS.SATISFACTION] },
    { label: 'ADMIN', color: SHEET_COLORS.TAB_RED_ORANGE, tabs: [SHEETS.FUNCTION_CHECKLIST, SHEETS.CONFIG] }
  ];

  var navRow = 3;
  for (var g = 0; g < navGroups.length; g++) {
    var group = navGroups[g];
    // Category header spans cols H-J
    sheet.getRange(navRow, navCol, 1, 3).merge()
      .setValue(group.label)
      .setBackground(group.color)
      .setFontColor(SHEET_COLORS.TEXT_WHITE)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    navRow++;

    // Tab names as navigable links across columns
    for (var t = 0; t < group.tabs.length; t++) {
      var tabName = group.tabs[t];
      var targetSheet = getSheetSafe_(ss, tabName);
      if (targetSheet) {
        var gid = targetSheet.getSheetId();
        sheet.getRange(navRow, navCol + (t % 3))
          .setFormula('=HYPERLINK("#gid=' + gid + '","' + tabName.replace(/"/g, '""') + '")')
          .setFontColor(SHEET_COLORS.LINK_PRIMARY)
          .setFontSize(11);
      } else {
        sheet.getRange(navRow, navCol + (t % 3))
          .setValue(tabName)
          .setFontColor(SHEET_COLORS.TEXT_GRAY)
          .setFontSize(11);
      }
      if (t % 3 === 2 || t === group.tabs.length - 1) navRow++;
    }
  }

  // Light border around the nav panel
  var panelRange = sheet.getRange(2, navCol, navRow - 2, 3);
  panelRange.setBorder(true, true, true, true, false, false,
    SHEET_COLORS.THEME_NAVY, SpreadsheetApp.BorderStyle.SOLID);

  // Color-code step sections by scanning column A for "Step" keywords
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var stepColors = {
    '1': SHEET_COLORS.THEME_GOLD,
    '2': SHEET_COLORS.THEME_STEEL_BLUE,
    '3': SHEET_COLORS.THEME_AMBER,
    '4': SHEET_COLORS.THEME_GREEN
  };

  for (var r = 0; r < data.length; r++) {
    var val = String(data[r][0]).trim();
    // Match "Step 1", "Step 2", etc.
    var match = val.match(/^Step\s+(\d)/i);
    if (match && stepColors[match[1]]) {
      // Limit to columns A-G so we don't overwrite the nav panel in H-J
      applySectionDivider_(sheet, r + 2, Math.min(numCols, 7), stepColors[match[1]]);
    }
  }

  Logger.log('Formatted: Getting Started');
}

// ─── Tab 2: FAQ ─────────────────────────────────────────────────────────────

function formatFAQTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FAQ);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Apply alternating Q/A backgrounds: scan for question vs answer rows
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var sectionColors = [
    SHEET_COLORS.THEME_NAVY,
    SHEET_COLORS.THEME_STEEL_BLUE,
    SHEET_COLORS.THEME_GREEN,
    SHEET_COLORS.THEME_AMBER,
    SHEET_COLORS.THEME_RED
  ];
  var sectionIdx = 0;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < data.length; r++) {
    var val = String(data[r][0]).trim().toUpperCase();
    var row = r + 2;

    // Detect section headers (all-caps text like "GETTING STARTED", "MEMBER DIRECTORY")
    var rawVal = String(data[r][0]).trim();
    if (rawVal.length > 3 && rawVal === rawVal.toUpperCase() &&
        !val.match(/^\d/) && !val.match(/^Q[:.]/) && val.indexOf('?') === -1 &&
        rawVal.length < 50) {
      var color = sectionColors[sectionIdx % sectionColors.length];
      applySectionDivider_(sheet, row, numCols, color);
      sectionIdx++;
    }
  }

  Logger.log('Formatted: FAQ');
}

// ─── Tab 3: Survey Questions ────────────────────────────────────────────────

function formatSurveyQuestionsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SURVEY_QUESTIONS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 10);

  // Row 1 = column headers → brand header style
  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Column widths: Question Text wide, IDs narrow
  applyColumnWidths_(sheet, { 1: 80, 2: 100, 3: 100, 4: 140, 5: 300, 6: 100, 7: 80, 8: 80, 9: 200, 10: 100 });

  // Text wrap on Question Text column (5) and Options column (9)
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1) {
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);
    if (numCols >= 9) sheet.getRange(2, 9, lastRow - 1, 1).setWrap(true);
  }

  // Conditional formatting: Type column (6) color-coding
  clearConditionalFormats_(sheet);
  var typeRange = sheet.getRange(2, 6, Math.max(lastRow - 1, 50), 1);
  var activeRange = sheet.getRange(2, 8, Math.max(lastRow - 1, 50), 1);
  var rules = [];

  var typeColors = [
    { text: 'dropdown', bg: '#DBEAFE' },     // blue
    { text: 'slider-10', bg: '#F3E8FF' },    // purple
    { text: 'radio', bg: '#FFEDD5' },         // orange
    { text: 'paragraph', bg: '#CCFBF1' }     // teal
  ];
  for (var t = 0; t < typeColors.length; t++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(typeColors[t].text)
      .setBackground(typeColors[t].bg)
      .setRanges([typeRange])
      .build());
  }

  // Active column: Y = green, N = gray
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Y')
    .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
    .setRanges([activeRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N')
    .setBackground(SHEET_COLORS.BG_VERY_LIGHT_GRAY)
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setRanges([activeRange])
    .build());

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Survey Questions');
}

// ─── Tab 4: Notifications ───────────────────────────────────────────────────

function formatNotificationsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.NOTIFICATIONS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No notifications yet. Use 📢 Notifications menu to create in-app messages.');

  Logger.log('Formatted: Notifications');
}

// ─── Tab 5: Resources ───────────────────────────────────────────────────────

function formatResourcesTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.RESOURCES);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 6);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Column widths: Content column wide, text wrap
  applyColumnWidths_(sheet, { 1: 80, 2: 160, 3: 120, 4: 200, 5: 300, 6: 100 });
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1 && numCols >= 5) {
    sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true);  // Summary
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);  // Content
  }

  // Category color-coding via conditional formatting
  clearConditionalFormats_(sheet);
  var catRange = sheet.getRange(2, 3, Math.max(lastRow - 1, 50), 1);
  var catColors = [
    { text: 'Grievance Process', bg: '#FEE2E2' },
    { text: 'Know Your Rights', bg: SHEET_COLORS.BG_LIGHT_GREEN },
    { text: 'Forms & Templates', bg: '#DBEAFE' },
    { text: 'Contact Info', bg: '#CCFBF1' },
    { text: 'Guide', bg: SHEET_COLORS.BG_LIGHT_YELLOW },
    { text: 'FAQ', bg: '#F3E8FF' }
  ];
  var rules = [];
  for (var c = 0; c < catColors.length; c++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(catColors[c].text)
      .setBackground(catColors[c].bg)
      .setRanges([catRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Resources');
}

// ─── Tab 6: Resource Config ─────────────────────────────────────────────────

function formatResourceConfigTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.RESOURCE_CONFIG);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 4);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Resource Config');
}

// ─── Tab 7: Feedback & Development ──────────────────────────────────────────

function formatFeedbackTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FEEDBACK);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 140, 2: 140, 3: 110, 4: 90, 5: 200, 6: 300, 7: 100, 8: 120 });

  // Conditional formatting: Priority + Status
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 50);
  var priorityRange = sheet.getRange(2, 4, dataRows, 1);
  var statusRange = sheet.getRange(2, 7, dataRows, 1);
  var rules = [];

  // Priority: Critical=deep red, High=orange, Medium=amber, Low=gray
  var priorities = [
    { text: 'Critical', bg: SHEET_COLORS.THEME_RED, font: SHEET_COLORS.TEXT_WHITE },
    { text: 'High', bg: '#FED7AA', font: '#9A3412' },
    { text: 'Medium', bg: SHEET_COLORS.BG_LIGHT_YELLOW, font: '#92400E' },
    { text: 'Low', bg: SHEET_COLORS.BG_VERY_LIGHT_GRAY, font: SHEET_COLORS.TEXT_GRAY }
  ];
  for (var p = 0; p < priorities.length; p++) {
    var builder = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(priorities[p].text)
      .setBackground(priorities[p].bg)
      .setRanges([priorityRange]);
    if (priorities[p].font) builder = builder.setFontColor(priorities[p].font);
    rules.push(builder.build());
  }

  // Status: New=blue, Planned=purple, In Progress=amber, Resolved=green, Wont Fix=gray
  var statuses = [
    { text: 'New', bg: '#DBEAFE' },
    { text: 'Planned', bg: '#F3E8FF' },
    { text: 'In Progress', bg: '#FEF3C7' },
    { text: 'Resolved', bg: SHEET_COLORS.BG_LIGHT_GREEN },
    { text: 'Wont Fix', bg: SHEET_COLORS.BG_VERY_LIGHT_GRAY }
  ];
  for (var s = 0; s < statuses.length; s++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(statuses[s].text)
      .setBackground(statuses[s].bg)
      .setRanges([statusRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Feedback & Development');
}

// ─── Tab 8: Function Checklist ──────────────────────────────────────────────

function formatFunctionChecklistTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FUNCTION_CHECKLIST);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 6);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 50, 2: 120, 3: 200, 4: 160, 5: 300, 6: 200 });

  // Conditional formatting: checked rows (col A = TRUE) get green tint
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 100);
  var fullRange = sheet.getRange(2, 1, dataRows, numCols);

  var rules = [];
  // Whole-row green tint when checkbox is checked
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setBackground(SHEET_COLORS.BG_PALE_GREEN)
    .setRanges([fullRange])
    .build());
  // Whole-row highlight when Notes (col F) is non-empty
  if (numCols >= 6) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>TRUE, LEN($F2)>0)')
      .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
      .setRanges([fullRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Function Checklist');
}

// ─── Tab 9: Settings Overview ───────────────────────────────────────────────

function formatSettingsOverviewTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SETTINGS_OVERVIEW);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  // Steel blue header for admin tabs
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_STEEL_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);

  // Conditional formatting: empty Current Value cells = amber warning
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 20);
  if (numCols >= 2) {
    var valueRange = sheet.getRange(2, 2, dataRows, 1);
    var rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenCellEmpty()
        .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
        .setRanges([valueRange])
        .build()
    ];
    sheet.setConditionalFormatRules(rules);
  }

  applyRowBanding_(sheet, 2, numCols);
  Logger.log('Formatted: Settings Overview');
}

// ─── Tab 10: Events ─────────────────────────────────────────────────────────

function formatEventsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.PORTAL_EVENTS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 10);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 60, 2: 200, 3: 100, 4: 140, 5: 140, 6: 160, 7: 300, 8: 200, 9: 140, 10: 120 });

  // Conditional formatting: event Type color-coding
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 50);
  var typeRange = sheet.getRange(2, 3, dataRows, 1);
  var typeColors = [
    { text: 'Meeting', bg: '#DBEAFE' },
    { text: 'Negotiation', bg: '#FEE2E2' },
    { text: 'Training', bg: SHEET_COLORS.BG_LIGHT_GREEN },
    { text: 'Social', bg: '#F3E8FF' },
    { text: 'Community', bg: SHEET_COLORS.BG_LIGHT_YELLOW }
  ];
  var rules = [];
  for (var i = 0; i < typeColors.length; i++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(typeColors[i].text)
      .setBackground(typeColors[i].bg)
      .setRanges([typeRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No events yet. Use 📅 Events menu → Add New Event to create your first event.');

  Logger.log('Formatted: Events');
}

// ─── Tab 11: MeetingMinutes ─────────────────────────────────────────────────

function formatMeetingMinutesTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.PORTAL_MINUTES);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 60, 2: 120, 3: 200, 4: 300, 5: 300, 6: 140, 7: 120, 8: 200 });

  // Text wrap on Bullets (4) and FullMinutes (5)
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1) {
    sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true);
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);
  }

  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No meeting minutes yet. Use 📝 Meeting Minutes menu to add new minutes.');

  Logger.log('Formatted: MeetingMinutes');
}

// ─── Tab 12: Workload Reporting ─────────────────────────────────────────────

function formatWorkloadReportingTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.WORKLOAD_REPORTING);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 13);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Conditional formatting: Priority Cases column (3) — 0=green, 1-3=yellow, 4+=red
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 50);
  var priorityRange = sheet.getRange(2, 3, dataRows, 1);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
      .setRanges([priorityRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 3)
      .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
      .setRanges([priorityRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(4)
      .setBackground(SHEET_COLORS.BG_LIGHT_RED)
      .setRanges([priorityRange])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No workload data yet. Members submit weekly reports via the Workload tab in the web portal.');

  Logger.log('Formatted: Workload Reporting');
}

// ─── Tab 13: Non-Member Contacts ────────────────────────────────────────────

function formatNonMemberContactsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.NON_MEMBER_CONTACTS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 9);

  // Distinct amber-tinted header to differentiate from Member Directory
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_AMBER)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);

  applyColumnWidths_(sheet, { 1: 120, 2: 120, 3: 160, 4: 140, 5: 80, 6: 100, 7: 120, 8: 200, 9: 140 });
  applyRowBanding_(sheet, 2, numCols);

  Logger.log('Formatted: Non-Member Contacts');
}

// ─── Tab 14: Case Checklist ─────────────────────────────────────────────────

function formatCaseChecklistTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.CASE_CHECKLIST);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 9);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 80, 3: 120, 4: 250, 5: 120, 6: 80, 7: 80, 8: 120, 9: 120 });

  // Conditional formatting
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 100);
  var fullRange = sheet.getRange(2, 1, dataRows, numCols);
  var rules = [];

  // Completed (col G=TRUE) → green tint + strikethrough
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2=TRUE')
    .setBackground(SHEET_COLORS.BG_PALE_GREEN)
    .setStrikethrough(true)
    .setRanges([fullRange])
    .build());

  // Required + NOT completed → red tint warning
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2="Y", $G2<>TRUE)')
    .setBackground(SHEET_COLORS.BG_LIGHT_RED)
    .setRanges([fullRange])
    .build());

  // Action Type color-coding (col C)
  var actionRange = sheet.getRange(2, 3, dataRows, 1);
  var actionColors = [
    { text: 'Filing', bg: SHEET_COLORS.THEME_NAVY, font: SHEET_COLORS.TEXT_WHITE },
    { text: 'Documentation', bg: '#DBEAFE', font: '#1E40AF' },
    { text: 'Hearing Prep', bg: '#FFEDD5', font: '#9A3412' },
    { text: 'Follow-Up', bg: SHEET_COLORS.BG_LIGHT_GREEN, font: SHEET_COLORS.TEXT_DARK_GREEN },
    { text: 'Admin', bg: SHEET_COLORS.BG_VERY_LIGHT_GRAY, font: SHEET_COLORS.TEXT_GRAY }
  ];
  for (var a = 0; a < actionColors.length; a++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(actionColors[a].text)
      .setBackground(actionColors[a].bg)
      .setFontColor(actionColors[a].font)
      .setRanges([actionRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No checklist items yet. Items are auto-generated when a new case is created.');

  Logger.log('Formatted: Case Checklist');
}

// ─── Tab 15: Member Satisfaction ────────────────────────────────────────────

function formatMemberSatisfactionTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SATISFACTION);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 20);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Freeze first 4 columns (Timestamp + demographics) for horizontal scroll
  sheet.setFrozenColumns(4);

  // Uniform column widths: Timestamp wide, Q columns 90px
  var widths = { 1: 150 };
  for (var col = 2; col <= numCols; col++) {
    widths[col] = (col <= 5) ? 120 : 90;
  }
  applyColumnWidths_(sheet, widths);

  // Set uniform row height for readability
  var lastRow = sheet.getLastRow();
  for (var row = 1; row <= lastRow; row++) {
    sheet.setRowHeight(row, 30);
  }

  // Text wrap on header row so long question text is readable
  sheet.getRange(1, 1, 1, numCols).setWrap(true);

  // Color scale on numeric response columns (typically cols 6-20: slider 1-10)
  // Red→Yellow→Green gradient
  clearConditionalFormats_(sheet);
  if (lastRow > 1 && numCols >= 6) {
    var dataRows = lastRow - 1;
    var rules = [];
    var sliderStart = Math.min(6, numCols);
    var sliderEnd = Math.min(20, numCols);
    var sliderRange = sheet.getRange(2, sliderStart, dataRows, sliderEnd - sliderStart + 1);

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#FCA5A5', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMidpointWithValue('#FDE68A', SpreadsheetApp.InterpolationType.NUMBER, '5')
      .setGradientMaxpointWithValue('#6EE7B7', SpreadsheetApp.InterpolationType.NUMBER, '10')
      .setRanges([sliderRange])
      .build());

    sheet.setConditionalFormatRules(rules);
  }

  applyRowBanding_(sheet, 2, numCols);
  Logger.log('Formatted: Member Satisfaction');
}

// ─── Tab 16: Features Reference ─────────────────────────────────────────────

function formatFeaturesReferenceTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FEATURES_REFERENCE);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  // This sheet already has good formatting from createFeaturesReferenceSheet.
  // Just re-theme the main header row to match Union brand.
  var row1Val = String(sheet.getRange(1, 1).getValue());
  if (row1Val.indexOf('FEATURES REFERENCE') !== -1) {
    sheet.getRange(1, 1, 1, numCols)
      .setBackground(SHEET_COLORS.THEME_NAVY)
      .setFontColor(SHEET_COLORS.TEXT_WHITE);
  }

  // Re-theme the column header row (usually row 5)
  var lastRow = sheet.getLastRow();
  for (var r = 2; r <= Math.min(lastRow, 10); r++) {
    var val = String(sheet.getRange(r, 1).getValue()).trim();
    if (val === 'Category') {
      applyColumnHeaderRow_(sheet, r, numCols);
      sheet.setFrozenRows(r);
      break;
    }
  }

  Logger.log('Formatted: Features Reference');
}

// ─── Tab 17: Volunteer Hours ────────────────────────────────────────────────

function formatVolunteerHoursTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.VOLUNTEER_HOURS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 9);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 80, 3: 140, 4: 120, 5: 120, 6: 70, 7: 250, 8: 120, 9: 200 });

  // Type hint row (row 2) styling if present
  if (sheet.getLastRow() >= 2) {
    var row2Val = String(sheet.getRange(2, 1).getValue()).trim().toLowerCase();
    if (row2Val === 'auto-id' || row2Val.indexOf('auto') !== -1) {
      applySubtitleRow_(sheet, 2, numCols);
    }
  }

  // Conditional formatting: Unverified rows (Verified By empty)
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 2, 50);
  var fullRange = sheet.getRange(3, 1, dataRows, numCols);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(LEN($A3)>0, LEN($H3)=0)')
      .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
      .setRanges([fullRange])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 3, numCols);
  applyEmptyState_(sheet, 3, numCols,
    'No volunteer hours logged yet. Use 🤝 Volunteer menu → Log Hours to add entries.');

  Logger.log('Formatted: Volunteer Hours');
}

// ─── Tab 18: Meeting Attendance ─────────────────────────────────────────────

function formatMeetingAttendanceTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.MEETING_ATTENDANCE);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 120, 3: 120, 4: 200, 5: 80, 6: 140, 7: 80, 8: 200 });

  // Type hint row (row 2) if present
  if (sheet.getLastRow() >= 2) {
    var row2Val = String(sheet.getRange(2, 1).getValue()).trim().toLowerCase();
    if (row2Val === 'auto-id' || row2Val.indexOf('auto') !== -1) {
      applySubtitleRow_(sheet, 2, numCols);
    }
  }

  // Conditional formatting: Attended checkbox
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 2, 50);
  var fullRange = sheet.getRange(3, 1, dataRows, numCols);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3=TRUE')
      .setBackground(SHEET_COLORS.BG_PALE_GREEN)
      .setRanges([fullRange])
      .build()
  ];

  // Meeting Type color-coding (col C)
  var typeRange = sheet.getRange(3, 3, dataRows, 1);
  var mtgColors = [
    { text: 'Regular', bg: '#DBEAFE' },
    { text: 'Special', bg: '#FFEDD5' },
    { text: 'Committee', bg: '#CCFBF1' },
    { text: 'Emergency', bg: '#FEE2E2' }
  ];
  for (var m = 0; m < mtgColors.length; m++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mtgColors[m].text)
      .setBackground(mtgColors[m].bg)
      .setRanges([typeRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);

  applyRowBanding_(sheet, 3, numCols);
  applyEmptyState_(sheet, 3, numCols,
    'No attendance records yet. Use 📅 Meetings menu → Take Attendance to log participation.');

  Logger.log('Formatted: Meeting Attendance');
}

// ─── Tab 19: Meeting Check-In Log ───────────────────────────────────────────

function formatMeetingCheckInLogTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.MEETING_CHECKIN_LOG);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No check-ins yet. Check-ins are auto-populated when members sign in via the web portal.');

  Logger.log('Formatted: Meeting Check-In Log');
}

// ============================================================================
// TAB BAR COLORS
// ============================================================================

/**
 * Applies tab-bar color grouping to all visible sheets.
 * Blue = data entry, Green = engagement, Gold = documentation, Red = admin.
 * @param {Spreadsheet} ss
 * @private
 */
function applyTabBarColors_(ss) {
  var blue   = SHEET_COLORS.TAB_BLUE;
  var green  = SHEET_COLORS.TAB_GREEN;
  var gold   = SHEET_COLORS.TAB_GOLD;
  var red    = SHEET_COLORS.TAB_RED_ORANGE;

  // Blue — Data entry tabs
  var blueSheets = [
    SHEETS.PORTAL_EVENTS, SHEETS.PORTAL_MINUTES, SHEETS.MEETING_ATTENDANCE,
    SHEETS.MEETING_CHECKIN_LOG, SHEETS.VOLUNTEER_HOURS, SHEETS.WORKLOAD_REPORTING,
    SHEETS.NON_MEMBER_CONTACTS, SHEETS.CASE_CHECKLIST
  ];

  // Green — Engagement / survey tabs
  var greenSheets = [
    SHEETS.SURVEY_QUESTIONS, SHEETS.SATISFACTION, SHEETS.NOTIFICATIONS,
    SHEETS.FEEDBACK
  ];

  // Gold — Documentation / guide tabs
  var goldSheets = [
    SHEETS.GETTING_STARTED, SHEETS.FAQ, SHEETS.RESOURCES,
    SHEETS.RESOURCE_CONFIG, SHEETS.FEATURES_REFERENCE, SHEETS.CONFIG_GUIDE
  ];

  // Red-Orange — Admin / technical tabs
  var redSheets = [
    SHEETS.FUNCTION_CHECKLIST, SHEETS.SETTINGS_OVERVIEW, SHEETS.CONFIG
  ];

  var groups = [
    { names: blueSheets, color: blue },
    { names: greenSheets, color: green },
    { names: goldSheets, color: gold },
    { names: redSheets, color: red }
  ];

  for (var g = 0; g < groups.length; g++) {
    for (var n = 0; n < groups[g].names.length; n++) {
      var sheet = getSheetSafe_(ss, groups[g].names[n]);
      if (sheet) {
        sheet.setTabColor(groups[g].color);
      }
    }
  }

  // Core data sheets get a distinct purple
  var coreSheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];
  for (var c = 0; c < coreSheets.length; c++) {
    var coreSheet = getSheetSafe_(ss, coreSheets[c]);
    if (coreSheet) coreSheet.setTabColor('#6A1B9A');
  }

  Logger.log('Tab bar colors applied');
}

// ============================================================================
// MASTER FUNCTION — Apply Union Theme to All Tabs
// ============================================================================

/**
 * Applies the Union brand theme to all visible tabs.
 * Callable from Admin menu: 🎨 Apply Union Theme.
 * Idempotent — safe to run multiple times.
 */
function applyUnionThemeToAllTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Applying Union brand theme to all tabs...', '🎨 Theme', 10);

  try {
    // Phase 1: Format each tab
    formatGettingStartedTab_(ss);
    formatFAQTab_(ss);
    formatSurveyQuestionsTab_(ss);
    formatNotificationsTab_(ss);
    formatResourcesTab_(ss);
    formatResourceConfigTab_(ss);
    formatFeedbackTab_(ss);
    formatFunctionChecklistTab_(ss);
    formatSettingsOverviewTab_(ss);
    formatEventsTab_(ss);
    formatMeetingMinutesTab_(ss);
    formatWorkloadReportingTab_(ss);
    formatNonMemberContactsTab_(ss);
    formatCaseChecklistTab_(ss);
    formatMemberSatisfactionTab_(ss);
    formatFeaturesReferenceTab_(ss);
    formatVolunteerHoursTab_(ss);
    formatMeetingAttendanceTab_(ss);
    formatMeetingCheckInLogTab_(ss);

    // Phase 2: Tab bar colors
    applyTabBarColors_(ss);

    ss.toast('Union brand theme applied to all tabs!', '✅ Theme Complete', 5);
    Logger.log('applyUnionThemeToAllTabs: completed successfully');

  } catch (error) {
    Logger.log('applyUnionThemeToAllTabs error: ' + error.message);
    ss.toast('Theme error: ' + error.message, '❌ Error', 5);
  }
}

/**
 * Format a single tab by name (for debugging or selective formatting).
 * @param {string} sheetName - The tab name to format
 */
function formatSingleTab(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formatters = {};
  formatters[SHEETS.GETTING_STARTED] = formatGettingStartedTab_;
  formatters[SHEETS.FAQ] = formatFAQTab_;
  formatters[SHEETS.SURVEY_QUESTIONS] = formatSurveyQuestionsTab_;
  formatters[SHEETS.NOTIFICATIONS] = formatNotificationsTab_;
  formatters[SHEETS.RESOURCES] = formatResourcesTab_;
  formatters[SHEETS.RESOURCE_CONFIG] = formatResourceConfigTab_;
  formatters[SHEETS.FEEDBACK] = formatFeedbackTab_;
  formatters[SHEETS.FUNCTION_CHECKLIST] = formatFunctionChecklistTab_;
  formatters[SHEETS.SETTINGS_OVERVIEW] = formatSettingsOverviewTab_;
  formatters[SHEETS.PORTAL_EVENTS] = formatEventsTab_;
  formatters[SHEETS.PORTAL_MINUTES] = formatMeetingMinutesTab_;
  formatters[SHEETS.WORKLOAD_REPORTING] = formatWorkloadReportingTab_;
  formatters[SHEETS.NON_MEMBER_CONTACTS] = formatNonMemberContactsTab_;
  formatters[SHEETS.CASE_CHECKLIST] = formatCaseChecklistTab_;
  formatters[SHEETS.SATISFACTION] = formatMemberSatisfactionTab_;
  formatters[SHEETS.FEATURES_REFERENCE] = formatFeaturesReferenceTab_;
  formatters[SHEETS.VOLUNTEER_HOURS] = formatVolunteerHoursTab_;
  formatters[SHEETS.MEETING_ATTENDANCE] = formatMeetingAttendanceTab_;
  formatters[SHEETS.MEETING_CHECKIN_LOG] = formatMeetingCheckInLogTab_;

  var fn = formatters[sheetName];
  if (fn) {
    fn(ss);
    ss.toast('Formatted: ' + sheetName, '✅', 3);
  } else {
    ss.toast('No formatter found for: ' + sheetName, '⚠️', 3);
  }
}
