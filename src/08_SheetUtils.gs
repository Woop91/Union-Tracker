/**
 * ============================================================================
 * SheetCreation.gs - Sheet Creation and Management
 * ============================================================================
 *
 * This module handles all sheet creation functions including:
 * - Main dashboard setup (CREATE_509_DASHBOARD)
 * - Individual sheet creation (Config, Member Directory, Grievance Log, etc.)
 * - Sheet ordering and organization
 * - Hidden sheet setup
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Sheet creation and management functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main setup function - creates the complete 509 Dashboard
 * Creates the core sheets with proper structure and formatting
 * @returns {void}
 */
function CREATE_509_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('UI not available, proceeding without confirmation: ' + e.message);
  }

  if (ui) {
    var response = ui.alert(
      '🏗️ Create 509 Dashboard',
      'This will create the 509 Dashboard with:\n\n' +
      '• Config (dropdown sources)\n' +
      '• Member Directory\n' +
      '• Grievance Log (with Action Type dropdown)\n' +
      '• ✅ Case Checklist (track grievance tasks)\n' +
      '• 📊 Member Satisfaction (Survey tracking)\n' +
      '• 💡 Feedback & Development (Bug/feature tracking)\n' +
      '• ✅ Function Checklist (function reference)\n' +
      '• 📚 Getting Started (setup instructions)\n' +
      '• ❓ FAQ (common questions)\n' +
      '• 📖 Config Guide (how to use Config tab)\n\n' +
      'Note: All dashboards are now modal-based (popup windows).\n' +
      'Access them via: Union Hub > Dashboards menu.\n\n' +
      'Plus 6 hidden calculation sheets for self-healing formulas.\n\n' +
      'Existing sheets with matching names will be recreated.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Setup cancelled.');
      return;
    }
  }

  ss.toast('Starting dashboard creation...', '🏗️ Setup', 5);

  try {
    createConfigSheet(ss);
    ss.toast('Created Config sheet', '🏗️ Progress', 2);

    createMemberDirectory(ss);
    ss.toast('Created Member Directory', '🏗️ Progress', 2);

    createGrievanceLog(ss);
    ss.toast('Created Grievance Log', '🏗️ Progress', 2);

    setupActionTypeColumn();
    ss.toast('Setup Action Type dropdown', '🏗️ Progress', 2);

    getOrCreateChecklistSheet();
    ss.toast('Created Case Checklist sheet', '🏗️ Progress', 2);

    ss.toast('Setting up hidden sheets...', '🏗️ Progress', 3);
    setupHiddenSheets(ss);

    createSatisfactionSheet(ss);
    ss.toast('Created Member Satisfaction', '🏗️ Progress', 2);

    createFeedbackSheet(ss);
    ss.toast('Created Feedback & Development', '🏗️ Progress', 2);

    createFunctionChecklistSheet_();
    ss.toast('Created Function Checklist', '🏗️ Progress', 2);

    createGettingStartedSheet(ss);
    ss.toast('Created Getting Started', '🏗️ Progress', 2);

    createFAQSheet(ss);
    ss.toast('Created FAQ', '🏗️ Progress', 2);

    createConfigGuideSheet(ss);
    ss.toast('Created Config Guide', '🏗️ Progress', 2);

    saveFormUrlsToConfig_silent(ss);
    ss.toast('Saved form URLs to Config', '🏗️ Progress', 2);

    ss.toast('Setting up validations...', '🏗️ Progress', 3);
    setupDataValidations();

    reorderSheetsToStandard(ss);
    ss.toast('Sheets reordered', '🏗️ Progress', 2);

    ss.toast('Dashboard creation complete!', '✅ Success', 5);
    if (ui) {
      ui.alert('✅ Success', '509 Dashboard has been created successfully!\n\n' +
        '10 sheets created:\n' +
        '• Config, Member Directory, Grievance Log (data)\n' +
        '• ✅ Case Checklist (track grievance tasks)\n' +
        '• 📊 Member Satisfaction, 💡 Feedback (tracking)\n' +
        '• ✅ Function Checklist (function reference)\n' +
        '• 📚 Getting Started, ❓ FAQ, 📖 Config Guide (help)\n\n' +
        '📋 Action Type dropdown configured with 8 case types.\n' +
        '📊 Dashboards are now modal-based (popup windows).\n' +
        'Access via: Union Hub > Dashboards menu.\n\n' +
        'Plus 6 hidden calculation sheets with self-healing formulas.\n\n' +
        '⚡ Auto-sync trigger installed - dates and deadlines will\n' +
        'update automatically when you edit the sheets.\n\n' +
        'Use the Demo menu to seed sample data.', ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log('Error in CREATE_509_DASHBOARD: ' + error.message);
    if (ui) {
      ui.alert('❌ Error', 'An error occurred: ' + error.message, ui.ButtonSet.OK);
    }
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets or creates a sheet by name
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} name - Sheet name
 * @returns {Sheet} The sheet
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Reorders sheets to standard layout
 * @param {Spreadsheet} ss - The spreadsheet
 * @returns {void}
 */
function reorderSheetsToStandard(ss) {
  var desiredOrder = [
    SHEETS.GETTING_STARTED,
    SHEETS.FAQ,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.FEEDBACK,
    SHEETS.FUNCTION_CHECKLIST,
    SHEETS.CONFIG_GUIDE,
    'Config',
    SHEETS.DASHBOARD,
    SHEETS.SATISFACTION
  ];

  var position = 1;

  for (var i = 0; i < desiredOrder.length; i++) {
    var sheetName = desiredOrder[i];
    var sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++;
    }
  }

  var firstSheet = ss.getSheetByName(SHEETS.GETTING_STARTED) ||
                   ss.getSheetByName(SHEETS.MEMBER_DIR) ||
                   ss.getSheets()[0];
  if (firstSheet) {
    ss.setActiveSheet(firstSheet);
  }

  Logger.log('Sheets reordered to standard layout');
}

/**
 * Sets up hidden calculation sheets
 * @param {Spreadsheet} ss - The spreadsheet
 * @returns {void}
 */
function setupHiddenSheets(ss) {
  var hiddenSheets = [
    { name: HIDDEN_SHEETS.CALC_MEMBERS, setup: setupCalcMembersSheet },
    { name: HIDDEN_SHEETS.CALC_GRIEVANCES, setup: setupCalcGrievancesSheet },
    { name: HIDDEN_SHEETS.CALC_DEADLINES, setup: setupCalcDeadlinesSheet },
    { name: HIDDEN_SHEETS.CALC_STATS, setup: setupCalcStatsSheet },
    { name: HIDDEN_SHEETS.CALC_SYNC, setup: setupCalcSyncSheet },
    { name: HIDDEN_SHEETS.CALC_FORMULAS, setup: setupCalcFormulasSheet }
  ];

  hiddenSheets.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
    }
    sheet.clear();
    config.setup(sheet);
    sheet.hideSheet();
  });
}



/**
 * ============================================================================
 * DataValidation.gs - Data Validation and Multi-Select
 * ============================================================================
 *
 * This module handles all data validation functions including:
 * - Dropdown validations from Config sheet
 * - Multi-select functionality
 * - Validation triggers
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Data validation and multi-select functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// DATA VALIDATION SETUP
// ============================================================================

/**
 * Sets up all data validations for Member Directory and Grievance Log
 * Configures dropdowns from Config sheet values
 * @returns {void}
 */
function setupDataValidations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet || !memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found. Please run CREATE_509_DASHBOARD first.');
    return;
  }

  // Member Directory Validations - Single-select dropdowns
  setDropdownValidation(memberSheet, MEMBER_COLS.JOB_TITLE, configSheet, CONFIG_COLS.JOB_TITLES);
  setDropdownValidation(memberSheet, MEMBER_COLS.WORK_LOCATION, configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  setDropdownValidation(memberSheet, MEMBER_COLS.UNIT, configSheet, CONFIG_COLS.UNITS);
  setDropdownValidation(memberSheet, MEMBER_COLS.IS_STEWARD, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.SUPERVISOR, configSheet, CONFIG_COLS.SUPERVISORS);
  setDropdownValidation(memberSheet, MEMBER_COLS.MANAGER, configSheet, CONFIG_COLS.MANAGERS);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_LOCAL, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_CHAPTER, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_ALLIED, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.HOME_TOWN, configSheet, CONFIG_COLS.HOME_TOWNS);
  setDropdownValidation(memberSheet, MEMBER_COLS.CONTACT_STEWARD, configSheet, CONFIG_COLS.STEWARDS);

  // Member Directory Validations - Multi-select dropdowns
  setMultiSelectValidation(memberSheet, MEMBER_COLS.OFFICE_DAYS, configSheet, CONFIG_COLS.OFFICE_DAYS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.PREFERRED_COMM, configSheet, CONFIG_COLS.COMM_METHODS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.BEST_TIME, configSheet, CONFIG_COLS.BEST_TIMES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.COMMITTEES, configSheet, CONFIG_COLS.STEWARD_COMMITTEES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.ASSIGNED_STEWARD, configSheet, CONFIG_COLS.STEWARDS);

  // Grievance Log Validations
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.STATUS, configSheet, CONFIG_COLS.GRIEVANCE_STATUS);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.CURRENT_STEP, configSheet, CONFIG_COLS.GRIEVANCE_STEP);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ISSUE_CATEGORY, configSheet, CONFIG_COLS.ISSUE_CATEGORY);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ARTICLES, configSheet, CONFIG_COLS.ARTICLES);

  SpreadsheetApp.getActiveSpreadsheet().toast('Data validations applied successfully!', '✅ Success', 3);
}

/**
 * Sets Member ID validation dropdown from Member Directory
 * @param {Sheet} grievanceSheet - Grievance Log sheet
 * @param {Sheet} memberSheet - Member Directory sheet
 * @returns {void}
 */
function setMemberIdValidation(grievanceSheet, memberSheet) {
  var memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var sourceRange = memberSheet.getRange(memberIdCol + '2:' + memberIdCol + '1000');

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Sets dropdown validation from Config sheet
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 * @returns {void}
 */
function setDropdownValidation(targetSheet, targetCol, configSheet, sourceCol) {
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  var rowCount = Math.max(10, actualRows + 10);
  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Sets multi-select validation (allows comma-separated values)
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 * @returns {void}
 */
function setMultiSelectValidation(targetSheet, targetCol, configSheet, sourceCol) {
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  var rowCount = Math.max(10, actualRows + 10);
  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true)  // Allow comma-separated values
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

// ============================================================================
// MULTI-SELECT FUNCTIONALITY
// ============================================================================

/**
 * Gets values from a Config sheet column
 * @param {Sheet} configSheet - The Config sheet
 * @param {number} col - Column number
 * @returns {Array<string>} Array of non-empty values
 */
function getConfigValues(configSheet, col) {
  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return [];

  var data = configSheet.getRange(3, col, lastRow - 2, 1).getValues();
  var values = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '') {
      values.push(data[i][0].toString());
    }
  }
  return values;
}

/**
 * Applies the multi-select value to the stored cell
 * @param {string} value - Comma-separated selected values
 * @returns {void}
 */
function applyMultiSelectValue(value) {
  var props = PropertiesService.getDocumentProperties();
  var row = parseInt(props.getProperty('multiSelectRow'), 10);
  var col = parseInt(props.getProperty('multiSelectCol'), 10);

  if (!row || !col) {
    throw new Error('Target cell not found. Please try again.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  sheet.getRange(row, col).setValue(value);

  props.deleteProperty('multiSelectRow');
  props.deleteProperty('multiSelectCol');
}

/**
 * Handles edit events to trigger multi-select dialog
 * @param {Object} e - Edit event object
 * @returns {void}
 */
function onEditMultiSelect(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (row < 2) return;

  var config = getMultiSelectConfig(col);
  if (!config) return;

  var newValue = e.value || '';

  if (newValue === '' || newValue.indexOf(',') !== -1) return;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tip: Use Dashboard menu > "Multi-Select Editor" for easier selection of multiple values.',
    config.label,
    5
  );
}

/**
 * Handles selection change to auto-open multi-select dialog
 * @param {Object} e - Selection change event object
 * @returns {void}
 */
function onSelectionChangeMultiSelect(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  var config = getMultiSelectConfig(col);
  if (!config) return;

  var props = PropertiesService.getDocumentProperties();
  var lastCell = props.getProperty('lastMultiSelectCell');
  var currentCell = row + ',' + col;

  if (lastCell === currentCell) return;

  props.setProperty('lastMultiSelectCell', currentCell);
  openCellMultiSelectEditor();
}

/**
 * Installs the multi-select auto-open trigger
 * @returns {void}
 */
function installMultiSelectTrigger() {
  var ui = SpreadsheetApp.getUi();

  ui.alert('☑️ Multi-Select Auto-Open Setup',
    'To enable auto-open for multi-select cells:\n\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click the clock icon (Triggers) in the left sidebar\n' +
    '3. Click "+ Add Trigger"\n' +
    '4. Choose function: onSelectionChangeMultiSelect\n' +
    '5. Select event type: "On change" or "On edit"\n' +
    '6. Click Save\n\n' +
    'Alternatively, use the manual method:\n' +
    '• Select a multi-select cell\n' +
    '• Go to Tools → Multi-Select → Open Editor',
    ui.ButtonSet.OK);
}

/**
 * Removes the multi-select auto-open trigger
 * @returns {void}
 */
function removeMultiSelectTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  var removed = false;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSelectionChangeMultiSelect') {
      ScriptApp.deleteTrigger(trigger);
      removed = true;
    }
  });

  if (removed) {
    SpreadsheetApp.getUi().alert('Multi-Select auto-open has been disabled.');
  } else {
    SpreadsheetApp.getUi().alert('No multi-select trigger was found.');
  }
}

/**
 * Sets dropdown validation dynamically
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 * @returns {void}
 */
function setDropdownValidationDynamic(targetSheet, targetCol, configSheet, sourceCol) {
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  var rowCount = Math.max(10, actualRows + 10);
  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Sets multi-select validation dynamically
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 * @returns {void}
 */
function setMultiSelectValidationDynamic(targetSheet, targetCol, configSheet, sourceCol) {
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  var rowCount = Math.max(10, actualRows + 10);
  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true)
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}



/**
 * ============================================================================
 * SearchEngine.gs - Search and Navigation Functions
 * ============================================================================
 *
 * This module handles all search-related functions including:
 * - Desktop search
 * - Quick search
 * - Advanced search with filters
 * - Navigation to search results
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Search and navigation functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// SEARCH FUNCTIONS
// ============================================================================

/**
 * Gets locations for desktop search filter dropdown
 * @returns {Array<string>} Array of unique locations sorted alphabetically
 */
function getDesktopSearchLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = [];

  var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (mSheet && mSheet.getLastRow() > 1) {
    var mData = mSheet.getRange(2, MEMBER_COLS.WORK_LOCATION, mSheet.getLastRow() - 1, 1).getValues();
    mData.forEach(function(row) {
      var loc = row[0];
      if (loc && locations.indexOf(loc) === -1) {
        locations.push(loc);
      }
    });
  }

  return locations.sort();
}

/**
 * Gets search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array<Object>} Array of search results
 */
function getDesktopSearchData(query, tab, filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = (query || '').toLowerCase();
  filters = filters || {};

  // Search Members
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var lastCol = Math.max(MEMBER_COLS.IS_STEWARD, MEMBER_COLS.WORK_LOCATION, MEMBER_COLS.JOB_TITLE, MEMBER_COLS.EMAIL);
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, lastCol).getValues();

      mData.forEach(function(row, index) {
        var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var firstName = row[MEMBER_COLS.FIRST_NAME - 1] || '';
        var lastName = row[MEMBER_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
        var jobTitle = row[MEMBER_COLS.JOB_TITLE - 1] || '';
        var location = row[MEMBER_COLS.WORK_LOCATION - 1] || '';
        var isSteward = row[MEMBER_COLS.IS_STEWARD - 1] || '';

        // Apply filters
        if (filters.location && location !== filters.location) return;
        if (filters.isSteward && isSteward !== filters.isSteward) return;

        // Search across fields
        var searchable = (memberId + ' ' + fullName + ' ' + email + ' ' + jobTitle + ' ' + location).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.location && !filters.isSteward) return;

        results.push({
          type: 'member',
          id: memberId,
          title: fullName.trim() || 'Unnamed Member',
          email: email,
          jobTitle: jobTitle,
          location: location,
          isSteward: isSteward,
          row: index + 2
        });
      });
    }
  }

  // Search Grievances
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var lastGCol = Math.max(GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.ISSUE_CATEGORY, GRIEVANCE_COLS.LOCATION, GRIEVANCE_COLS.STEWARD, GRIEVANCE_COLS.DATE_FILED);
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, lastGCol).getValues();

      gData.forEach(function(row, index) {
        var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
        var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
        var issueType = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '';
        var location = row[GRIEVANCE_COLS.LOCATION - 1] || '';
        var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
        var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1] || '';

        // Apply filters
        if (filters.status && status !== filters.status) return;
        if (filters.location && location !== filters.location) return;

        // Search across fields
        var searchable = (grievanceId + ' ' + fullName + ' ' + status + ' ' + issueType + ' ' + location + ' ' + steward).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.status && !filters.location) return;

        // Format date
        var filedDateStr = '';
        if (dateFiled) {
          try {
            filedDateStr = Utilities.formatDate(new Date(dateFiled), Session.getScriptTimeZone(), 'MM/dd/yyyy');
          } catch(e) {
            filedDateStr = dateFiled.toString();
          }
        }

        results.push({
          type: 'grievance',
          id: grievanceId,
          title: fullName.trim() || 'Unknown Member',
          status: status,
          issueType: issueType,
          location: location,
          steward: steward,
          filedDate: filedDateStr,
          row: index + 2
        });
      });
    }
  }

  return results.slice(0, 50);
}

/**
 * Navigates to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
 * @returns {void}
 */
function navigateToSearchResult(type, id, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = type === 'member' ? SHEETS.MEMBER_DIR : SHEETS.GRIEVANCE_LOG;
  var sheet = ss.getSheetByName(sheetName);

  if (sheet && row) {
    ss.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(row, 1));
    SpreadsheetApp.flush();
  }
}

/**
 * Navigates to active grievances view
 * @returns {void}
 */
function viewActiveGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * Searches the dashboard for matching records
 * @param {string} query - Search query
 * @param {string} searchType - 'all', 'members', or 'grievances'
 * @param {Object} filters - Additional filters
 * @returns {Array<Object>} Search results
 */
function searchDashboard(query, searchType, filters) {
  var results = [];
  var queryLower = query.toLowerCase();

  // Search members
  if (searchType === 'all' || searchType === 'members') {
    var members = getMemberList();
    members.forEach(function(m) {
      if (m.name.toLowerCase().indexOf(queryLower) !== -1 ||
          m.id.toLowerCase().indexOf(queryLower) !== -1 ||
          m.department.toLowerCase().indexOf(queryLower) !== -1) {
        results.push({
          id: m.id,
          type: 'member',
          title: m.name,
          subtitle: m.department + ' - ID: ' + m.id
        });
      }
    });
  }

  // Search grievances
  if (searchType === 'all' || searchType === 'grievances') {
    var grievances = getOpenGrievances();
    grievances.forEach(function(g) {
      var grievanceId = g['Grievance ID'] || '';
      var memberName = g['Member Name'] || '';
      var description = g['Description'] || '';
      var status = g['Status'] || '';

      if (grievanceId.toLowerCase().indexOf(queryLower) !== -1 ||
          memberName.toLowerCase().indexOf(queryLower) !== -1 ||
          description.toLowerCase().indexOf(queryLower) !== -1) {

        // Apply status filter
        if (filters.status && status.toLowerCase() !== filters.status.toLowerCase()) {
          return;
        }

        results.push({
          id: grievanceId,
          type: 'grievance',
          title: grievanceId + ' - ' + memberName,
          subtitle: status + ' - Step ' + (g['Current Step'] || 1)
        });
      }
    });
  }

  return results.slice(0, 50);
}

/**
 * Quick search for instant results
 * @param {string} query - Search query
 * @returns {Array<Object>} Quick search results (max 10)
 */
function quickSearchDashboard(query) {
  if (!query || query.length < 2) return [];

  var results = searchDashboard(query, 'all', {});
  return results.slice(0, 10);
}

/**
 * Advanced search with complex filters
 * @param {Object} filters - Search filters
 * @returns {Array<Object>} Search results
 */
function advancedSearch(filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search members if included
  if (filters.includeMembers) {
    var memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
    if (memberSheet) {
      var data = memberSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var matches = true;

        // Apply department filter
        if (filters.department && row[MEMBER_COLUMNS.DEPARTMENT] !== filters.department) {
          matches = false;
        }

        // Apply name filter
        if (filters.name && matches) {
          var fullName = (row[MEMBER_COLUMNS.FIRST_NAME] + ' ' + row[MEMBER_COLUMNS.LAST_NAME]).toLowerCase();
          if (fullName.indexOf(filters.name.toLowerCase()) === -1) {
            matches = false;
          }
        }

        // Apply steward filter
        if (filters.stewardOnly && matches) {
          if (row[MEMBER_COLUMNS.IS_STEWARD] !== 'Yes') {
            matches = false;
          }
        }

        if (matches && row[MEMBER_COLUMNS.ID]) {
          results.push({
            id: row[MEMBER_COLUMNS.ID],
            type: 'member',
            title: row[MEMBER_COLUMNS.FIRST_NAME] + ' ' + row[MEMBER_COLUMNS.LAST_NAME],
            subtitle: row[MEMBER_COLUMNS.DEPARTMENT],
            row: i + 1
          });
        }
      }
    }
  }

  // Search grievances if included
  if (filters.includeGrievances) {
    var grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    if (grievanceSheet) {
      var gData = grievanceSheet.getDataRange().getValues();

      for (var j = 1; j < gData.length; j++) {
        var gRow = gData[j];
        var gMatches = true;

        // Apply status filter
        if (filters.status && gRow[GRIEVANCE_COLUMNS.STATUS] !== filters.status) {
          gMatches = false;
        }

        // Apply date range filter
        if (filters.startDate && gMatches) {
          var filedDate = gRow[GRIEVANCE_COLUMNS.FILING_DATE];
          if (filedDate && new Date(filedDate) < new Date(filters.startDate)) {
            gMatches = false;
          }
        }

        if (filters.endDate && gMatches) {
          var endFiledDate = gRow[GRIEVANCE_COLUMNS.FILING_DATE];
          if (endFiledDate && new Date(endFiledDate) > new Date(filters.endDate)) {
            gMatches = false;
          }
        }

        if (gMatches && gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID]) {
          results.push({
            id: gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
            type: 'grievance',
            title: gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
            subtitle: gRow[GRIEVANCE_COLUMNS.STATUS] + ' - ' + gRow[GRIEVANCE_COLUMNS.TYPE],
            row: j + 1
          });
        }
      }
    }
  }

  return results.slice(0, 100);
}

/**
 * Gets department list from calc sheet or direct query
 * @returns {Array<string>} Array of department names
 */
function getDepartmentList() {
  var formulaSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HIDDEN_SHEETS.CALC_FORMULAS);

  if (!formulaSheet) {
    var memberSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

    if (!memberSheet) return [];

    var data = memberSheet.getRange(2, MEMBER_COLUMNS.DEPARTMENT + 1,
      memberSheet.getLastRow() - 1, 1).getValues();

    var depts = {};
    data.forEach(function(row) {
      if (row[0]) depts[row[0]] = true;
    });

    return Object.keys(depts).sort();
  }

  var deptData = formulaSheet.getRange('B4:B').getValues();
  return deptData.filter(function(row) { return row[0]; }).map(function(row) { return row[0]; });
}

/**
 * Gets member list for dropdowns
 * @returns {Array<Object>} Array of member objects with id, name, department
 */
function getMemberList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLUMNS.ID]) {
      members.push({
        id: data[i][MEMBER_COLUMNS.ID],
        name: data[i][MEMBER_COLUMNS.FIRST_NAME] + ' ' + data[i][MEMBER_COLUMNS.LAST_NAME],
        department: data[i][MEMBER_COLUMNS.DEPARTMENT]
      });
    }
  }

  return members;
}

/**
 * Gets a value from a hidden calculation sheet
 * @param {string} sheetName - The hidden sheet name
 * @param {string} cellRef - The cell reference (e.g., 'B4')
 * @returns {*} The cell value or null if not found
 */
function getCalcValue(sheetName, cellRef) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Hidden sheet ' + sheetName + ' not found');
    return null;
  }

  return sheet.getRange(cellRef).getValue();
}



/**
 * ============================================================================
 * ChartBuilder.gs - Dashboard Chart Generation
 * ============================================================================
 *
 * This module handles all chart-related functions including:
 * - Chart options display
 * - Chart generation (bar, pie, line, gauge, etc.)
 * - Chart styling and customization
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Chart generation and display functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// CHART GENERATION
// ============================================================================

/**
 * Generates the selected chart based on user's choice in cell G120
 * Call this from the menu after user selects a chart number
 */
function generateSelectedChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!sheet) {
    ss.toast('Dashboard sheet not found', '❌ Error', 3);
    return;
  }

  var chartNum = sheet.getRange('G120').getValue();
  if (!chartNum || chartNum < 1 || chartNum > 15) {
    ss.toast('Please enter a valid chart number (1-15) in cell G120', '⚠️ Invalid Selection', 5);
    return;
  }

  ss.toast('Generating chart #' + chartNum + '...', '📊 Creating Chart', 2);

  // Remove existing charts in the display area
  var existingCharts = sheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    sheet.removeChart(existingCharts[i]);
  }

  var chart;
  var chartBuilder;

  switch (parseInt(chartNum)) {
    case 1: // Gauge Chart - Win Rate
      createGaugeStyleChart_(sheet);
      break;

    case 2: // Bar Chart - Status Distribution
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange('A15:A16'))
        .addRange(sheet.getRange('A16:F16'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Grievance Status Distribution')
        .setOption('legend', {position: 'bottom'})
        .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.ACCENT_ORANGE, COLORS.STATUS_BLUE, COLORS.UNION_GREEN, '#9CA3AF', '#6B7280'])
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 3: // Pie Chart - Issue Categories
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange('A26:B30'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Grievances by Issue Category')
        .setOption('pieHole', 0)
        .setOption('legend', {position: 'right'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 4: // Line Chart - Trends
      createTrendLineChart_(sheet);
      break;

    case 5: // Column Chart - Location
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(sheet.getRange('A35:B39'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Members by Location')
        .setOption('legend', {position: 'none'})
        .setOption('colors', [COLORS.CHART_CYAN])
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 6: // Scorecard
      createScorecardChart_(sheet);
      break;

    case 7: // Heatmap Table
      ss.toast('Heatmap applied! Check "Apply Gradient Heatmaps" in Admin menu for color scales.', '🔥 Heatmap', 5);
      if (typeof applyWinRateGradients === 'function') {
        applyWinRateGradients();
      }
      break;

    case 8: // Stacked Bar
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange('A26:D30'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Cases by Issue Category (Open vs Resolved)')
        .setOption('isStacked', true)
        .setOption('legend', {position: 'bottom'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 9: // Donut Chart
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange('A26:B30'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Issue Category Distribution')
        .setOption('pieHole', 0.4)
        .setOption('legend', {position: 'right'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 10: // Area Chart
      createAreaChart_(sheet);
      break;

    case 11: // Combo Chart
      createComboChart_(sheet);
      break;

    case 12: // Scatter Plot
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(sheet.getRange('A35:C39'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Response Time vs Outcome Analysis')
        .setOption('legend', {position: 'bottom'})
        .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.UNION_GREEN])
        .setOption('width', 600)
        .setOption('height', 300)
        .setOption('pointSize', 8);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 13: // Histogram
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.HISTOGRAM)
        .addRange(sheet.getRange('B26:B30'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Case Duration Distribution')
        .setOption('legend', {position: 'none'})
        .setOption('colors', [COLORS.CHART_CYAN])
        .setOption('width', 600)
        .setOption('height', 300)
        .setOption('histogram', {bucketSize: 5});
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 14: // Summary Table
      createSummaryTableChart_(sheet);
      break;

    case 15: // Steward Leaderboard
      createStewardLeaderboardChart_(sheet);
      break;

    default:
      ss.toast('Enter 1-15 in cell G120 to select a chart type. See options table above.', 'ℹ️ Chart Help', 5);
  }

  ss.toast('Chart generated! Scroll down to "Chart Display Area" to view.', '✅ Done', 5);
}

// ============================================================================
// CHART HELPER FUNCTIONS
// ============================================================================

/**
 * Creates a gauge-style display (Google Sheets doesn't have native gauge)
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createGaugeStyleChart_(sheet) {
  var winRate = sheet.getRange('D6').getValue() || '0%';
  var gaugeText = '═══════════════════════════════════════\n' +
                  '           🎯 WIN RATE GAUGE\n' +
                  '═══════════════════════════════════════\n\n' +
                  '                 ' + winRate + '\n\n' +
                  '    ◀━━━━━━━━━━━━━━━━━━━━━━━━━━━▶\n' +
                  '    0%        50%        100%\n\n' +
                  '═══════════════════════════════════════';

  sheet.getRange('A135').setValue(gaugeText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F0FDF4');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates a scorecard-style display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createScorecardChart_(sheet) {
  var openCases = sheet.getRange('A16').getValue() || 0;
  var scorecardText = '┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓\n' +
                      '┃         📊 OPEN GRIEVANCES          ┃\n' +
                      '┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫\n' +
                      '┃                                     ┃\n' +
                      '┃              ' + openCases + '                    ┃\n' +
                      '┃                                     ┃\n' +
                      '┃           ▲ Active Cases            ┃\n' +
                      '┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛';

  sheet.getRange('A135').setValue(scorecardText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FEF3C7');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates a trend line chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createTrendLineChart_(sheet) {
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('A44:C46'))
    .setPosition(135, 1, 0, 0)
    .setOption('title', 'Month-Over-Month Trends')
    .setOption('legend', {position: 'bottom'})
    .setOption('curveType', 'function')
    .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.CHART_BLUE])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates an area chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createAreaChart_(sheet) {
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(sheet.getRange('A44:C46'))
    .setPosition(135, 1, 0, 0)
    .setOption('title', 'Cumulative Trend Analysis')
    .setOption('legend', {position: 'bottom'})
    .setOption('colors', [COLORS.CHART_PURPLE, COLORS.CHART_BLUE])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates a combo chart (bars with line overlay)
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createComboChart_(sheet) {
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange('A26:C30'))
    .setPosition(135, 1, 0, 0)
    .setOption('title', 'Cases by Category with Trend Line')
    .setOption('legend', {position: 'bottom'})
    .setOption('seriesType', 'bars')
    .setOption('series', {1: {type: 'line'}})
    .setOption('colors', [COLORS.CHART_BLUE, COLORS.SOLIDARITY_RED])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates a summary table display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createSummaryTableChart_(sheet) {
  var openCases = sheet.getRange('A16').getValue() || 0;
  var resolvedCases = sheet.getRange('D16').getValue() || 0;
  var winRate = sheet.getRange('D6').getValue() || '0%';
  var avgDays = sheet.getRange('E6').getValue() || 'N/A';

  var tableText = '╔═══════════════════════════════════════════════════════╗\n' +
                  '║            📋 KPI SUMMARY TABLE                       ║\n' +
                  '╠═══════════════════════════════════════════════════════╣\n' +
                  '║  Metric                    │  Value                   ║\n' +
                  '╠═══════════════════════════════════════════════════════╣\n' +
                  '║  Open Cases                │  ' + padRight(String(openCases), 23) + '║\n' +
                  '║  Resolved Cases            │  ' + padRight(String(resolvedCases), 23) + '║\n' +
                  '║  Win Rate                  │  ' + padRight(String(winRate), 23) + '║\n' +
                  '║  Avg Resolution Time       │  ' + padRight(String(avgDays), 23) + '║\n' +
                  '╚═══════════════════════════════════════════════════════╝';

  sheet.getRange('A135').setValue(tableText)
    .setFontFamily('Courier New')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F3F4F6');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates steward leaderboard chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createStewardLeaderboardChart_(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    sheet.getRange('A135').setValue('Member Directory not found');
    return;
  }

  var data = memberSheet.getDataRange().getValues();

  // Build case counts from Grievance Log since MEMBER_COLS doesn't have TOTAL_CASES/WINS
  var caseCounts = {};
  var winCounts = {};
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var gData = grievanceSheet.getDataRange().getValues();
    for (var g = 1; g < gData.length; g++) {
      var stewardName = gData[g][GRIEVANCE_COLS.STEWARD - 1] || '';
      if (stewardName) {
        caseCounts[stewardName] = (caseCounts[stewardName] || 0) + 1;
        var gStatus = gData[g][GRIEVANCE_COLS.STATUS - 1] || '';
        if (gStatus === 'Won' || gStatus === 'Settled') {
          winCounts[stewardName] = (winCounts[stewardName] || 0) + 1;
        }
      }
    }
  }

  var stewards = [];

  // Find stewards and look up their case counts
  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      var fullName = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      stewards.push({
        name: fullName,
        cases: caseCounts[fullName] || 0,
        wins: winCounts[fullName] || 0
      });
    }
  }

  // Sort by cases descending
  stewards.sort(function(a, b) { return b.cases - a.cases; });

  var leaderboardText = '╔═══════════════════════════════════════════════════════╗\n' +
                        '║           🏆 STEWARD LEADERBOARD                      ║\n' +
                        '╠═══════════════════════════════════════════════════════╣\n' +
                        '║  Rank │ Name                      │ Cases │ Wins     ║\n' +
                        '╠═══════════════════════════════════════════════════════╣\n';

  var medals = ['🥇', '🥈', '🥉'];
  for (var j = 0; j < Math.min(5, stewards.length); j++) {
    var rank = j < 3 ? medals[j] : (j + 1) + '.';
    leaderboardText += '║  ' + padRight(rank, 4) + ' │ ' +
                       padRight(stewards[j].name, 25) + ' │ ' +
                       padRight(String(stewards[j].cases), 5) + ' │ ' +
                       padRight(String(stewards[j].wins), 8) + '║\n';
  }

  leaderboardText += '╚═══════════════════════════════════════════════════════╝';

  sheet.getRange('A135').setValue(leaderboardText)
    .setFontFamily('Courier New')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#EFF6FF');
  sheet.getRange('A135:G145').merge();
}

/**
 * Pads a string to a specified length
 * @param {string} str - String to pad
 * @param {number} len - Target length
 * @returns {string} Padded string
 */
function padRight(str, len) {
  str = String(str);
  while (str.length < len) {
    str += ' ';
  }
  return str.substring(0, len);
}



/**
 * ============================================================================
 * 08e_FormHandlers.gs - Form Handlers Module
 * ============================================================================
 *
 * This module contains all form-related functions for the 509 Steward Dashboard:
 * - Form submission handlers (Contact, Satisfaction, Grievance)
 * - Form trigger setup functions
 * - Form URL handling and configuration
 * - Form value parsing utilities
 *
 * Dependencies:
 * - SHEETS, CONFIG_COLS constants from 00_Constants.gs
 * - GRIEVANCE_FORM_CONFIG, CONTACT_FORM_CONFIG, SATISFACTION_FORM_CONFIG
 * - MEMBER_COLS, SATISFACTION_COLS from column configuration
 * - Helper functions: findExistingMember, generateNameBasedId, validateMemberEmail
 * - Helper functions: getCurrentQuarter, computeSatisfactionRowAverages, syncSatisfactionValues
 *
 * Note: onGrievanceFormSubmit() is defined in 05_Integrations.gs
 *
 * @author SEIU Local 509 Development Team
 * @version 1.0.0
 */

// ============================================================================
// FORM URL CONFIGURATION
// ============================================================================

/**
 * Get form URL from Config sheet, with fallback to hardcoded defaults
 * @param {string} formType - Type of form: 'grievance', 'contact', or 'satisfaction'
 * @returns {string} The form URL
 */
function getFormUrlFromConfig(formType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var configCol, defaultUrl;

  switch(formType.toLowerCase()) {
    case 'grievance':
      configCol = CONFIG_COLS.GRIEVANCE_FORM_URL;
      defaultUrl = GRIEVANCE_FORM_CONFIG.FORM_URL;
      break;
    case 'contact':
      configCol = CONFIG_COLS.CONTACT_FORM_URL;
      defaultUrl = CONTACT_FORM_CONFIG.FORM_URL;
      break;
    case 'satisfaction':
      configCol = CONFIG_COLS.SATISFACTION_FORM_URL;
      defaultUrl = SATISFACTION_FORM_CONFIG.FORM_URL;
      break;
    default:
      Logger.log('Unknown form type: ' + formType);
      return '';
  }

  // Try to get from Config sheet (row 2 contains data, row 3 for newer format)
  if (configSheet) {
    // Check row 2 first (original format)
    var url = configSheet.getRange(2, configCol).getValue();
    if (!url || url === '') {
      // Check row 3 (newer format with section headers)
      url = configSheet.getRange(3, configCol).getValue();
    }
    if (url && url !== '' && url.indexOf('http') === 0) {
      return url;
    }
  }

  // Fall back to hardcoded default
  return defaultUrl;
}

/**
 * Build pre-filled grievance form URL
 * @param {Object} memberData - Member information object
 * @param {Object} stewardData - Steward information object
 * @returns {string} Pre-filled form URL
 * @private
 */
function buildGrievanceFormUrl_(memberData, stewardData) {
  // Get form URL from Config (allows admin to update without code changes)
  var baseUrl = getFormUrlFromConfig('grievance');
  var fields = GRIEVANCE_FORM_CONFIG.FIELD_IDS;

  var params = [];

  // Member info
  if (memberData.memberId) params.push(fields.MEMBER_ID + '=' + encodeURIComponent(memberData.memberId));
  if (memberData.firstName) params.push(fields.MEMBER_FIRST_NAME + '=' + encodeURIComponent(memberData.firstName));
  if (memberData.lastName) params.push(fields.MEMBER_LAST_NAME + '=' + encodeURIComponent(memberData.lastName));
  if (memberData.jobTitle) params.push(fields.JOB_TITLE + '=' + encodeURIComponent(memberData.jobTitle));
  if (memberData.unit) params.push(fields.AGENCY_DEPARTMENT + '=' + encodeURIComponent(memberData.unit));
  if (memberData.workLocation) {
    params.push(fields.REGION + '=' + encodeURIComponent(memberData.workLocation));
    params.push(fields.WORK_LOCATION + '=' + encodeURIComponent(memberData.workLocation));
  }
  if (memberData.manager) params.push(fields.MANAGERS + '=' + encodeURIComponent(memberData.manager));
  if (memberData.email) params.push(fields.MEMBER_EMAIL + '=' + encodeURIComponent(memberData.email));

  // Steward info
  if (stewardData.firstName) params.push(fields.STEWARD_FIRST_NAME + '=' + encodeURIComponent(stewardData.firstName));
  if (stewardData.lastName) params.push(fields.STEWARD_LAST_NAME + '=' + encodeURIComponent(stewardData.lastName));
  if (stewardData.email) params.push(fields.STEWARD_EMAIL + '=' + encodeURIComponent(stewardData.email));

  // Default values
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  params.push(fields.DATE_FILED + '=' + encodeURIComponent(today));
  params.push(fields.STEP + '=' + encodeURIComponent('I'));

  return baseUrl + '?usp=pp_url&' + params.join('&');
}

/**
 * Save form URLs to the Config tab for easy reference and updating
 * Writes Grievance Form, Contact Form, and Satisfaction Survey URLs to Config columns P, Q, AR
 */
function saveFormUrlsToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  saveFormUrlsToConfig_silent(ss);
  ss.toast('Form URLs saved to Config tab (columns P, Q, AR)', 'Saved', 3);
}

/**
 * Silent version - used during CREATE_509_DASHBOARD setup
 * @param {Spreadsheet} ss - The spreadsheet object
 * @private
 */
function saveFormUrlsToConfig_silent(ss) {
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    Logger.log('Config sheet not found - cannot save form URLs');
    return;
  }

  // Config layout: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
  // Set column headers in row 2
  configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue('Grievance Form URL');
  configSheet.getRange(2, CONFIG_COLS.CONTACT_FORM_URL).setValue('Contact Form URL');
  configSheet.getRange(2, CONFIG_COLS.SATISFACTION_FORM_URL).setValue('Satisfaction Survey URL');

  // Set form URLs in row 3 (data row)
  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue(GRIEVANCE_FORM_CONFIG.FORM_URL);
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setValue(CONTACT_FORM_CONFIG.FORM_URL);
  configSheet.getRange(3, CONFIG_COLS.SATISFACTION_FORM_URL).setValue(SATISFACTION_FORM_CONFIG.FORM_URL);

  // Format as links
  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(3, CONFIG_COLS.SATISFACTION_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
}

// ============================================================================
// FORM VALUE PARSING UTILITIES
// ============================================================================

/**
 * Get a value from form named responses
 * @param {Object} responses - Form named values object
 * @param {string} fieldName - Name of the field to retrieve
 * @returns {string} The field value or empty string
 * @private
 */
function getFormValue_(responses, fieldName) {
  if (responses[fieldName] && responses[fieldName].length > 0) {
    return responses[fieldName][0];
  }
  return '';
}

/**
 * Get multiple values from form response (for checkbox questions)
 * Returns comma-separated string
 * @param {Object} responses - Form named values object
 * @param {string} fieldName - Name of the field to retrieve
 * @returns {string} Comma-separated values or empty string
 * @private
 */
function getFormMultiValue_(responses, fieldName) {
  if (responses[fieldName] && responses[fieldName].length > 0) {
    // Filter out empty values and join with comma
    var values = responses[fieldName].filter(function(v) { return v && v.trim() !== ''; });
    return values.join(', ');
  }
  return '';
}

/**
 * Parse a date string from form submission
 * @param {string} dateStr - Date string to parse
 * @returns {Date|string} Parsed Date object or original string if parsing fails
 * @private
 */
function parseFormDate_(dateStr) {
  if (!dateStr) return '';

  try {
    var date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      return dateStr; // Return as-is if can't parse
    }
    return date;
  } catch (e) {
    return dateStr;
  }
}

// ============================================================================
// CONTACT FORM HANDLER
// ============================================================================

/**
 * Show the Personal Contact Info form link
 * Members fill out the blank form and data is written to Member Directory on submit
 */
function sendContactInfoForm() {
  var ui = SpreadsheetApp.getUi();
  // Get form URL from Config (allows admin to update without code changes)
  var formUrl = getFormUrlFromConfig('contact');

  // Show dialog with form link options
  var response = ui.alert('Personal Contact Info Form',
    'Share this form with members to collect their contact information.\n\n' +
    'When submitted, the data will be written to the Member Directory:\n' +
    '- Existing members (matched by name) will be updated\n' +
    '- New members will be added automatically\n\n' +
    '- Click YES to open the form\n' +
    '- Click NO to copy the link',
    ui.ButtonSet.YES_NO_CANCEL);

  if (response === ui.Button.YES) {
    // Open form in new window
    var html = HtmlService.createHtmlOutput(
      '<script>window.open("' + formUrl + '", "_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    ui.showModalDialog(html, 'Opening form...');
  } else if (response === ui.Button.NO) {
    // Show link to copy
    var copyHtml = HtmlService.createHtmlOutput(
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + formUrl + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  ta.select();' +
      '  document.execCommand("copy");' +
      '  document.getElementById("copied").style.display = "inline";' +
      '}' +
      '</script>' +
      '</div></body></html>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, 'Contact Form Link');
  }
}

/**
 * Handle contact form submission
 * Writes member data to Member Directory (updates existing or creates new)
 *
 * @param {Object} e - Form submission event object
 */
function onContactFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    Logger.log('Member Directory sheet not found');
    return;
  }

  try {
    // Get form responses from event
    var responses = e.namedValues || {};

    // Extract form data
    var firstName = getFormValue_(responses, 'First Name');
    var lastName = getFormValue_(responses, 'Last Name');
    var jobTitle = getFormValue_(responses, 'Job Title / Position');
    var unit = getFormValue_(responses, 'Department / Unit');
    var workLocation = getFormValue_(responses, 'Worksite / Office Location');
    var officeDays = getFormMultiValue_(responses, 'Work Schedule / Office Days');
    var preferredComm = getFormMultiValue_(responses, 'Please select your preferred communication methods (check all that apply):');
    var bestTime = getFormMultiValue_(responses, 'What time(s) are best for us to reach you? (check all that apply)');
    var supervisor = getFormValue_(responses, 'Immediate Supervisor');
    var manager = getFormValue_(responses, 'Manager / Program Director');
    var email = getFormValue_(responses, 'Personal Email');
    var phone = getFormValue_(responses, 'Personal Phone Number');
    var interestAllied = getFormValue_(responses, 'Willing to support other chapters (DDS, DCF, Public Sector, etc.)?');
    var interestChapter = getFormValue_(responses, 'Willing to be active in sub-chapter (at other worksites within your agency of employment)?');
    var interestLocal = getFormValue_(responses, 'Willing to join direct actions (e.g., at your place of employment)?');

    // Require at least first and last name
    if (!firstName || !lastName) {
      Logger.log('Contact form submission missing name: ' + firstName + ' ' + lastName);
      return;
    }

    // Multi-Key Smart Match: Check ID, Email, then Name (hierarchical)
    var data = memberSheet.getDataRange().getValues();
    var memberId = getFormValue_(responses, 'Member ID');  // Optional field from form

    var match = findExistingMember({
      memberId: memberId,
      email: email,
      firstName: firstName,
      lastName: lastName
    }, data);

    var memberRow = match ? match.row : -1;

    if (match) {
      Logger.log('Found existing member via ' + match.matchType + ' match (confidence: ' + match.confidence + ') at row ' + match.row);
    }

    if (memberRow === -1) {
      // Member not found - create new member
      // Mask name in logs for privacy
      var maskedName = typeof maskName === 'function' ? maskName(firstName + ' ' + lastName) : '[REDACTED]';
      Logger.log('Creating new member: ' + maskedName);

      // Generate Member ID
      var existingIds = {};
      for (var k = 1; k < data.length; k++) {
        var id = data[k][MEMBER_COLS.MEMBER_ID - 1];
        if (id) existingIds[id] = true;
      }
      var memberId = generateNameBasedId('M', firstName, lastName, existingIds);

      // Build new row array
      var newRow = [];
      newRow[MEMBER_COLS.MEMBER_ID - 1] = memberId;
      newRow[MEMBER_COLS.FIRST_NAME - 1] = firstName;
      newRow[MEMBER_COLS.LAST_NAME - 1] = lastName;
      newRow[MEMBER_COLS.JOB_TITLE - 1] = jobTitle || '';
      newRow[MEMBER_COLS.WORK_LOCATION - 1] = workLocation || '';
      newRow[MEMBER_COLS.UNIT - 1] = unit || '';
      newRow[MEMBER_COLS.OFFICE_DAYS - 1] = officeDays || '';
      newRow[MEMBER_COLS.EMAIL - 1] = email || '';
      newRow[MEMBER_COLS.PHONE - 1] = phone || '';
      newRow[MEMBER_COLS.PREFERRED_COMM - 1] = preferredComm || '';
      newRow[MEMBER_COLS.BEST_TIME - 1] = bestTime || '';
      newRow[MEMBER_COLS.SUPERVISOR - 1] = supervisor || '';
      newRow[MEMBER_COLS.MANAGER - 1] = manager || '';
      newRow[MEMBER_COLS.IS_STEWARD - 1] = 'No';
      newRow[MEMBER_COLS.INTEREST_LOCAL - 1] = interestLocal || '';
      newRow[MEMBER_COLS.INTEREST_CHAPTER - 1] = interestChapter || '';
      newRow[MEMBER_COLS.INTEREST_ALLIED - 1] = interestAllied || '';

      // Append new member row
      memberSheet.appendRow(newRow);
      // Log with masked name for privacy
      Logger.log('Created new member ' + memberId + ': ' + maskedName);

    } else {
      // Update existing member record with form data
      var updates = [];

      // Update all fields from form (even if they change existing values)
      if (jobTitle) updates.push({ col: MEMBER_COLS.JOB_TITLE, value: jobTitle });
      if (unit) updates.push({ col: MEMBER_COLS.UNIT, value: unit });
      if (workLocation) updates.push({ col: MEMBER_COLS.WORK_LOCATION, value: workLocation });
      if (officeDays) updates.push({ col: MEMBER_COLS.OFFICE_DAYS, value: officeDays });
      if (preferredComm) updates.push({ col: MEMBER_COLS.PREFERRED_COMM, value: preferredComm });
      if (bestTime) updates.push({ col: MEMBER_COLS.BEST_TIME, value: bestTime });
      if (supervisor) updates.push({ col: MEMBER_COLS.SUPERVISOR, value: supervisor });
      if (manager) updates.push({ col: MEMBER_COLS.MANAGER, value: manager });
      if (email) updates.push({ col: MEMBER_COLS.EMAIL, value: email });
      if (phone) updates.push({ col: MEMBER_COLS.PHONE, value: phone });
      if (interestLocal) updates.push({ col: MEMBER_COLS.INTEREST_LOCAL, value: interestLocal });
      if (interestChapter) updates.push({ col: MEMBER_COLS.INTEREST_CHAPTER, value: interestChapter });
      if (interestAllied) updates.push({ col: MEMBER_COLS.INTEREST_ALLIED, value: interestAllied });

      // Apply updates
      for (var j = 0; j < updates.length; j++) {
        memberSheet.getRange(memberRow, updates[j].col).setValue(updates[j].value);
      }

      // Mask name in logs for privacy
      var maskedUpdateName = typeof maskName === 'function' ? maskName(firstName + ' ' + lastName) : '[REDACTED]';
      Logger.log('Updated contact info for ' + maskedUpdateName + ' (row ' + memberRow + ')');
    }

  } catch (error) {
    Logger.log('Error processing contact form submission: ' + error.message);
    throw error;
  }
}

/**
 * Set up the contact form submission trigger
 * Run this once to enable automatic processing of form submissions
 */
function setupContactFormTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Check for existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  var hasContactTrigger = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onContactFormSubmit') {
      hasContactTrigger = true;
      break;
    }
  }

  if (hasContactTrigger) {
    ui.alert('Trigger Exists',
      'A contact form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Contact Form Trigger',
    'This will set up automatic processing of contact info form submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  if (!formUrl) {
    ui.alert('No URL', 'Please provide the form edit URL.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Extract form ID from URL
    var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onContactFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('Trigger Created',
      'Contact form trigger has been set up!\n\n' +
      'When a contact form is submitted:\n' +
      '- The member\'s record will be updated in Member Directory\n' +
      '- Contact info, preferences, and interests will be saved',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// GRIEVANCE FORM TRIGGER SETUP
// ============================================================================

/**
 * Set up the grievance form submission trigger
 * Run this once to enable automatic processing of form submissions
 *
 * Note: The actual handler onGrievanceFormSubmit() is defined in 05_Integrations.gs
 */
function setupGrievanceFormTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Check for existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  var hasGrievanceTrigger = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onGrievanceFormSubmit') {
      hasGrievanceTrigger = true;
      break;
    }
  }

  if (hasGrievanceTrigger) {
    ui.alert('Trigger Exists',
      'A grievance form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Grievance Form Trigger',
    'This will set up automatic processing of grievance form submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):\n' +
    '(Leave blank to use the configured form)',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  try {
    var formId;

    if (formUrl) {
      // Extract form ID from URL
      var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
        return;
      }
      formId = match[1];
    } else {
      // Use configured form
      var configFormUrl = GRIEVANCE_FORM_CONFIG.FORM_URL;
      var match = configFormUrl.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        ui.alert('No Form Configured',
          'No form URL provided and could not extract ID from config.\n\n' +
          'Please provide the form edit URL.',
          ui.ButtonSet.OK);
        return;
      }
      // Note: The /e/ URL is the published version, we need the actual form ID
      ui.alert('Form URL Needed',
        'Please provide the form edit URL (the one ending in /edit).\n\n' +
        'You can find this by opening the form in edit mode.',
        ui.ButtonSet.OK);
      return;
    }

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onGrievanceFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('Trigger Created',
      'Grievance form trigger has been set up!\n\n' +
      'When a grievance form is submitted:\n' +
      '- A new row will be added to Grievance Log\n' +
      '- A Drive folder will be created automatically\n' +
      '- Deadlines will be calculated\n' +
      '- Member Directory will be updated',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// SATISFACTION SURVEY HANDLER
// ============================================================================

/**
 * Show the Member Satisfaction Survey form link
 * Survey responses are written to the Member Satisfaction sheet
 */
function getSatisfactionSurveyLink() {
  var ui = SpreadsheetApp.getUi();
  // Get form URL from Config (allows admin to update without code changes)
  var formUrl = getFormUrlFromConfig('satisfaction');

  // Show dialog with form link options
  var response = ui.alert('Member Satisfaction Survey',
    'Share this survey with members to collect feedback.\n\n' +
    'When submitted, responses will be written to the\n' +
    'Member Satisfaction sheet.\n\n' +
    '- Click YES to open the survey\n' +
    '- Click NO to copy the link',
    ui.ButtonSet.YES_NO_CANCEL);

  if (response === ui.Button.YES) {
    // Open form in new window
    var html = HtmlService.createHtmlOutput(
      '<script>window.open("' + formUrl + '", "_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    ui.showModalDialog(html, 'Opening survey...');
  } else if (response === ui.Button.NO) {
    // Show link to copy
    var copyHtml = HtmlService.createHtmlOutput(
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + formUrl + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  ta.select();' +
      '  document.execCommand("copy");' +
      '  document.getElementById("copied").style.display = "inline";' +
      '}' +
      '</script>' +
      '</div></body></html>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, 'Survey Link');
  }
}

/**
 * Handle satisfaction survey form submission
 * Writes survey responses to the Member Satisfaction sheet
 *
 * @param {Object} e - Form submission event object
 */
function onSatisfactionFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet) {
    Logger.log('Member Satisfaction sheet not found');
    return;
  }

  try {
    // Get form responses from event
    var responses = e.namedValues || {};

    // Build row data array matching SATISFACTION_COLS order
    var newRow = [];

    // Timestamp
    newRow[SATISFACTION_COLS.TIMESTAMP - 1] = new Date();

    // Work Context (Q1-5) - Note: Q3_SHIFT not in form, column left empty
    newRow[SATISFACTION_COLS.Q1_WORKSITE - 1] = getFormValue_(responses, 'Worksite / Program / Region');
    newRow[SATISFACTION_COLS.Q2_ROLE - 1] = getFormValue_(responses, 'Role / Job Group');
    // Q3_SHIFT skipped - form does not have this question
    newRow[SATISFACTION_COLS.Q4_TIME_IN_ROLE - 1] = getFormValue_(responses, 'Time in current role');
    newRow[SATISFACTION_COLS.Q5_STEWARD_CONTACT - 1] = getFormValue_(responses, 'Contact with steward in past 12 months?');

    // Overall Satisfaction (Q6-9)
    newRow[SATISFACTION_COLS.Q6_SATISFIED_REP - 1] = getFormValue_(responses, 'Satisfied with union representation');
    newRow[SATISFACTION_COLS.Q7_TRUST_UNION - 1] = getFormValue_(responses, 'Trust union to act in best interests');
    newRow[SATISFACTION_COLS.Q8_FEEL_PROTECTED - 1] = getFormValue_(responses, 'Feel more protected at work');
    newRow[SATISFACTION_COLS.Q9_RECOMMEND - 1] = getFormValue_(responses, 'Would recommend membership');

    // Steward Ratings 3A (Q10-17)
    newRow[SATISFACTION_COLS.Q10_TIMELY_RESPONSE - 1] = getFormValue_(responses, 'Responded in timely manner');
    newRow[SATISFACTION_COLS.Q11_TREATED_RESPECT - 1] = getFormValue_(responses, 'Treated me with respect');
    newRow[SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS - 1] = getFormValue_(responses, 'Explained options clearly');
    newRow[SATISFACTION_COLS.Q13_FOLLOWED_THROUGH - 1] = getFormValue_(responses, 'Followed through on commitments');
    newRow[SATISFACTION_COLS.Q14_ADVOCATED - 1] = getFormValue_(responses, 'Advocated effectively');
    newRow[SATISFACTION_COLS.Q15_SAFE_CONCERNS - 1] = getFormValue_(responses, 'Felt safe raising concerns');
    newRow[SATISFACTION_COLS.Q16_CONFIDENTIALITY - 1] = getFormValue_(responses, 'Handled confidentiality appropriately');
    newRow[SATISFACTION_COLS.Q17_STEWARD_IMPROVE - 1] = getFormValue_(responses, 'What should stewards improve?');

    // Steward Access 3B (Q18-20)
    newRow[SATISFACTION_COLS.Q18_KNOW_CONTACT - 1] = getFormValue_(responses, 'Know how to contact steward/rep');
    newRow[SATISFACTION_COLS.Q19_CONFIDENT_HELP - 1] = getFormValue_(responses, 'Confident I would get help');
    newRow[SATISFACTION_COLS.Q20_EASY_FIND - 1] = getFormValue_(responses, 'Easy to figure out who to contact');

    // Chapter Effectiveness (Q21-25)
    newRow[SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES - 1] = getFormValue_(responses, 'Reps understand my workplace issues');
    newRow[SATISFACTION_COLS.Q22_CHAPTER_COMM - 1] = getFormValue_(responses, 'Chapter communication is regular and clear');
    newRow[SATISFACTION_COLS.Q23_ORGANIZES - 1] = getFormValue_(responses, 'Chapter organizes members effectively');
    newRow[SATISFACTION_COLS.Q24_REACH_CHAPTER - 1] = getFormValue_(responses, 'Know how to reach chapter contact');
    newRow[SATISFACTION_COLS.Q25_FAIR_REP - 1] = getFormValue_(responses, 'Representation is fair across roles/shifts');

    // Local Leadership (Q26-31)
    newRow[SATISFACTION_COLS.Q26_DECISIONS_CLEAR - 1] = getFormValue_(responses, 'Leadership communicates decisions clearly');
    newRow[SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS - 1] = getFormValue_(responses, 'Understand how decisions are made');
    newRow[SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE - 1] = getFormValue_(responses, 'Union is transparent about finances');
    newRow[SATISFACTION_COLS.Q29_ACCOUNTABLE - 1] = getFormValue_(responses, 'Leadership is accountable to feedback');
    newRow[SATISFACTION_COLS.Q30_FAIR_PROCESSES - 1] = getFormValue_(responses, 'Internal processes feel fair');
    newRow[SATISFACTION_COLS.Q31_WELCOMES_OPINIONS - 1] = getFormValue_(responses, 'Union welcomes differing opinions');

    // Contract Enforcement (Q32-36)
    newRow[SATISFACTION_COLS.Q32_ENFORCES_CONTRACT - 1] = getFormValue_(responses, 'Union enforces contract effectively');
    newRow[SATISFACTION_COLS.Q33_REALISTIC_TIMELINES - 1] = getFormValue_(responses, 'Communicates realistic timelines');
    newRow[SATISFACTION_COLS.Q34_CLEAR_UPDATES - 1] = getFormValue_(responses, 'Provides clear updates on issues');
    newRow[SATISFACTION_COLS.Q35_FRONTLINE_PRIORITY - 1] = getFormValue_(responses, 'Prioritizes frontline conditions');
    newRow[SATISFACTION_COLS.Q36_FILED_GRIEVANCE - 1] = getFormValue_(responses, 'Filed grievance in past 24 months?');

    // Representation Process 6A (Q37-40)
    newRow[SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS - 1] = getFormValue_(responses, 'Understood steps and timeline');
    newRow[SATISFACTION_COLS.Q38_FELT_SUPPORTED - 1] = getFormValue_(responses, 'Felt supported throughout');
    newRow[SATISFACTION_COLS.Q39_UPDATES_OFTEN - 1] = getFormValue_(responses, 'Received updates often enough');
    newRow[SATISFACTION_COLS.Q40_OUTCOME_JUSTIFIED - 1] = getFormValue_(responses, 'Outcome feels justified');

    // Communication Quality (Q41-45)
    newRow[SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE - 1] = getFormValue_(responses, 'Communications are clear and actionable');
    newRow[SATISFACTION_COLS.Q42_ENOUGH_INFO - 1] = getFormValue_(responses, 'Receive enough information');
    newRow[SATISFACTION_COLS.Q43_FIND_EASILY - 1] = getFormValue_(responses, 'Can find information easily');
    newRow[SATISFACTION_COLS.Q44_ALL_SHIFTS - 1] = getFormValue_(responses, 'Communications reach all locations');
    newRow[SATISFACTION_COLS.Q45_MEETINGS_WORTH - 1] = getFormValue_(responses, 'Meetings are worth attending');

    // Member Voice & Culture (Q46-50)
    newRow[SATISFACTION_COLS.Q46_VOICE_MATTERS - 1] = getFormValue_(responses, 'My voice matters in the union');
    newRow[SATISFACTION_COLS.Q47_SEEKS_INPUT - 1] = getFormValue_(responses, 'Union actively seeks input');
    newRow[SATISFACTION_COLS.Q48_DIGNITY - 1] = getFormValue_(responses, 'Members treated with dignity');
    newRow[SATISFACTION_COLS.Q49_NEWER_SUPPORTED - 1] = getFormValue_(responses, 'Newer members are supported');
    newRow[SATISFACTION_COLS.Q50_CONFLICT_RESPECT - 1] = getFormValue_(responses, 'Internal conflict handled respectfully');

    // Value & Collective Action (Q51-55)
    newRow[SATISFACTION_COLS.Q51_GOOD_VALUE - 1] = getFormValue_(responses, 'Union provides good value for dues');
    newRow[SATISFACTION_COLS.Q52_PRIORITIES_NEEDS - 1] = getFormValue_(responses, 'Priorities reflect member needs');
    newRow[SATISFACTION_COLS.Q53_PREPARED_MOBILIZE - 1] = getFormValue_(responses, 'Union prepared to mobilize');
    newRow[SATISFACTION_COLS.Q54_HOW_INVOLVED - 1] = getFormValue_(responses, 'Understand how to get involved');
    newRow[SATISFACTION_COLS.Q55_WIN_TOGETHER - 1] = getFormValue_(responses, 'Acting together, we can win improvements');

    // Scheduling/Office Days (Q56-63)
    newRow[SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES - 1] = getFormValue_(responses, 'Understand proposed changes');
    newRow[SATISFACTION_COLS.Q57_ADEQUATELY_INFORMED - 1] = getFormValue_(responses, 'Feel adequately informed');
    newRow[SATISFACTION_COLS.Q58_CLEAR_CRITERIA - 1] = getFormValue_(responses, 'Decisions use clear criteria');
    newRow[SATISFACTION_COLS.Q59_WORK_EXPECTATIONS - 1] = getFormValue_(responses, 'Work can be done under expectations');
    newRow[SATISFACTION_COLS.Q60_EFFECTIVE_OUTCOMES - 1] = getFormValue_(responses, 'Approach supports effective outcomes');
    newRow[SATISFACTION_COLS.Q61_SUPPORTS_WELLBEING - 1] = getFormValue_(responses, 'Approach supports my wellbeing');
    newRow[SATISFACTION_COLS.Q62_CONCERNS_SERIOUS - 1] = getFormValue_(responses, 'My concerns would be taken seriously');
    newRow[SATISFACTION_COLS.Q63_SCHEDULING_CHALLENGE - 1] = getFormValue_(responses, 'Biggest scheduling challenge?');

    // Priorities & Close (Q64-67)
    newRow[SATISFACTION_COLS.Q64_TOP_PRIORITIES - 1] = getFormMultiValue_(responses, 'Top 3 priorities (6-12 mo)');
    newRow[SATISFACTION_COLS.Q65_ONE_CHANGE - 1] = getFormValue_(responses, '#1 change union should make');
    newRow[SATISFACTION_COLS.Q66_KEEP_DOING - 1] = getFormValue_(responses, 'One thing union should keep doing');
    newRow[SATISFACTION_COLS.Q67_ADDITIONAL - 1] = getFormValue_(responses, 'Additional comments (no names)');

    // ========================================================================
    // EMAIL VERIFICATION & QUARTERLY TRACKING
    // ========================================================================

    // Get email from form (try multiple common field names)
    var email = getFormValue_(responses, 'Email Address') ||
                getFormValue_(responses, 'Email') ||
                getFormValue_(responses, 'email') ||
                (e.response ? e.response.getRespondentEmail() : '') || '';
    email = email.toString().toLowerCase().trim();

    newRow[SATISFACTION_COLS.EMAIL - 1] = email;

    // Get current quarter
    var currentQuarter = getCurrentQuarter();
    newRow[SATISFACTION_COLS.QUARTER - 1] = currentQuarter;

    // Validate email against Member Directory
    var memberMatch = validateMemberEmail(email);

    if (memberMatch) {
      // Email matches a member - mark as verified
      newRow[SATISFACTION_COLS.VERIFIED - 1] = 'Yes';
      newRow[SATISFACTION_COLS.MATCHED_MEMBER_ID - 1] = memberMatch.memberId;
      newRow[SATISFACTION_COLS.IS_LATEST - 1] = 'Yes';

      // Check for existing responses from this member in same quarter
      var existingData = satSheet.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) {
        var rowEmail = (existingData[i][SATISFACTION_COLS.EMAIL - 1] || '').toString().toLowerCase().trim();
        var rowQuarter = existingData[i][SATISFACTION_COLS.QUARTER - 1];
        var rowIsLatest = existingData[i][SATISFACTION_COLS.IS_LATEST - 1];

        // If same email, same quarter, and currently marked as latest
        if (rowEmail === email && rowQuarter === currentQuarter && rowIsLatest === 'Yes') {
          // Mark the old row as superseded (row index is i+1 because of 0-indexing and header)
          var oldRowNum = i + 1;
          satSheet.getRange(oldRowNum, SATISFACTION_COLS.IS_LATEST).setValue('No');
          satSheet.getRange(oldRowNum, SATISFACTION_COLS.SUPERSEDED_BY).setValue(satSheet.getLastRow() + 1);
          Logger.log('Marked row ' + oldRowNum + ' as superseded by new submission');
        }
      }
    } else {
      // Email doesn't match - flag for review
      newRow[SATISFACTION_COLS.VERIFIED - 1] = 'Pending Review';
      newRow[SATISFACTION_COLS.MATCHED_MEMBER_ID - 1] = '';
      newRow[SATISFACTION_COLS.IS_LATEST - 1] = 'Yes';
    }

    newRow[SATISFACTION_COLS.REVIEWER_NOTES - 1] = '';

    // Append row to satisfaction sheet
    satSheet.appendRow(newRow);

    // Compute section averages for the new row (no formulas in visible sheet)
    var newRowNum = satSheet.getLastRow();
    computeSatisfactionRowAverages(newRowNum);

    // Update dashboard summary values
    syncSatisfactionValues();

    Logger.log('Satisfaction survey response recorded at ' + new Date() + ' | Verified: ' + newRow[SATISFACTION_COLS.VERIFIED - 1]);

  } catch (error) {
    Logger.log('Error processing satisfaction survey submission: ' + error.message);
    throw error;
  }
}

/**
 * Set up the satisfaction survey form submission trigger
 * Run this once to enable automatic processing of survey submissions
 */
function setupSatisfactionFormTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Check for existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  var hasSatisfactionTrigger = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onSatisfactionFormSubmit') {
      hasSatisfactionTrigger = true;
      break;
    }
  }

  if (hasSatisfactionTrigger) {
    ui.alert('Trigger Exists',
      'A satisfaction survey trigger already exists.\n\n' +
      'Survey submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Satisfaction Survey Trigger',
    'This will set up automatic processing of survey submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  if (!formUrl) {
    ui.alert('No URL', 'Please provide the form edit URL.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Extract form ID from URL
    var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onSatisfactionFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('Trigger Created',
      'Satisfaction survey trigger has been set up!\n\n' +
      'When a survey is submitted:\n' +
      '- Response will be added to Member Satisfaction sheet\n' +
      '- All 68 questions will be recorded\n' +
      '- Dashboard will reflect new data',
      ui.ButtonSet.OK);

    ss.toast('Survey trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}



/**
 * ============================================================================
 * 08h_NotificationEngine.gs - Notification and Alert System
 * ============================================================================
 *
 * This module handles all notification and alert functionality for the
 * SEIU Local 509 Dashboard including:
 * - Deadline notification settings and triggers
 * - Steward deadline alerts
 * - Survey email distribution
 * - Member email validation
 * - Quarter utilities for notifications
 *
 * Dependencies:
 * - SHEETS constant (from 08_Code.gs)
 * - MEMBER_COLS constant (from 08_Code.gs)
 * - GRIEVANCE_COLS constant (from 08_Code.gs)
 * - CONFIG_COLS constant (from 08_Code.gs)
 * - SATISFACTION_FORM_CONFIG constant (from 08_Code.gs)
 *
 * @author SEIU Local 509
 * @version 1.0.0
 */

// ============================================================================
// NOTIFICATION SETTINGS AND TRIGGERS
// ============================================================================

/**
 * Show notification settings dialog
 * Allows user to enable/disable daily deadline notifications
 */
function showNotificationSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var enabled = props.getProperty('notifications_enabled') === 'true';
  var email = props.getProperty('notification_email') || Session.getEffectiveUser().getEmail();

  var response = ui.alert('Notification Settings',
    'Daily deadline notifications: ' + (enabled ? 'ENABLED' : 'DISABLED') + '\n' +
    'Email: ' + email + '\n\n' +
    'Notifications are sent daily at 8 AM for grievances due within 3 days.\n\n' +
    'Would you like to ' + (enabled ? 'DISABLE' : 'ENABLE') + ' notifications?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    if (enabled) {
      // Disable
      props.setProperty('notifications_enabled', 'false');
      removeDailyTrigger_();
      ui.alert('Notifications Disabled', 'Daily deadline notifications have been turned off.', ui.ButtonSet.OK);
    } else {
      // Enable
      props.setProperty('notifications_enabled', 'true');
      props.setProperty('notification_email', email);
      installDailyTrigger_();
      ui.alert('Notifications Enabled',
        'Daily notifications enabled!\n\n' +
        'You will receive an email at 8 AM when grievances are due within 3 days.\n\n' +
        'Email: ' + email, ui.ButtonSet.OK);
    }
  }
}

/**
 * Install daily trigger for notifications
 */
function installDailyTrigger_() {
  // Remove existing triggers
  removeDailyTrigger_();

  // Create new daily trigger at 8 AM
  ScriptApp.newTrigger('checkDeadlinesAndNotify_')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

/**
 * Remove daily notification trigger
 */
function removeDailyTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDeadlinesAndNotify_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Check deadlines and send notification email (called by trigger)
 */
function checkDeadlinesAndNotify_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('notifications_enabled') !== 'true') return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return;

  var email = props.getProperty('notification_email');
  if (!email) return;

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var threeDaysAhead = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  var urgent = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    var daysToDeadline = data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var currentStep = data[i][GRIEVANCE_COLS.CURRENT_STEP - 1];

    var closedStatuses = ['Closed', 'Settled', 'Won', 'Denied', 'Withdrawn'];
    if (closedStatuses.indexOf(status) !== -1) continue;

    if (daysToDeadline === 'Overdue' || (daysToDeadline !== '' && typeof daysToDeadline === 'number' && daysToDeadline <= 3)) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });
    }
  }

  if (urgent.length === 0) return;

  var subject = '509 Dashboard: ' + urgent.length + ' Grievance Deadline(s) Approaching';
  var body = 'The following grievances have deadlines within 3 days:\n\n';

  for (var j = 0; j < urgent.length; j++) {
    var g = urgent[j];
    body += '* ' + g.id + ' (' + g.step + ') - ' +
      (g.days <= 0 ? 'OVERDUE!' : g.days + ' day(s) remaining') + '\n';
  }

  body += '\n\nView your dashboard: ' + ss.getUrl();

  MailApp.sendEmail(email, subject, body);
}

/**
 * Test the notification system
 */
function testDeadlineNotifications() {
  var ui = SpreadsheetApp.getUi();
  var email = Session.getEffectiveUser().getEmail();

  var response = ui.alert('Test Notifications',
    'This will send a test notification email to:\n' + email + '\n\nSend test email?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    MailApp.sendEmail(email,
      '509 Dashboard Test Notification',
      'This is a test notification from your 509 Dashboard.\n\n' +
      'If you received this email, notifications are working correctly!\n\n' +
      'Dashboard: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    );
    ui.alert('Test Sent', 'Test email sent to ' + email + '\n\nCheck your inbox!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to send test email: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// STEWARD DEADLINE ALERTS
// ============================================================================

/**
 * Send daily digest to all stewards with their assigned grievance deadlines
 * Each steward gets their own personalized email
 */
function sendStewardDeadlineAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || !memberSheet) {
    Logger.log('Required sheets not found for steward alerts');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var alertDays = parseInt(props.getProperty('alert_days') || '7', 10);

  var grievanceData = sheet.getDataRange().getValues();
  var memberData = memberSheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // Build member lookup for steward emails
  var memberLookup = {};
  for (var m = 1; m < memberData.length; m++) {
    var memberId = memberData[m][MEMBER_COLS.MEMBER_ID - 1];
    if (memberId) {
      memberLookup[memberId] = {
        name: (memberData[m][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (memberData[m][MEMBER_COLS.LAST_NAME - 1] || ''),
        steward: memberData[m][MEMBER_COLS.ASSIGNED_STEWARD - 1] || ''
      };
    }
  }

  // Group grievances by steward
  var stewardGrievances = {};
  var closedStatuses = ['Closed', 'Settled', 'Won', 'Denied', 'Withdrawn'];

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var steward = row[GRIEVANCE_COLS.ASSIGNED_STEWARD - 1] || '';

    // Skip closed grievances
    if (closedStatuses.indexOf(status) !== -1) continue;
    if (!grievanceId) continue;

    // Check if deadline is within alert window
    var daysRemaining = null;
    if (daysToDeadline === 'Overdue') {
      daysRemaining = -1;
    } else if (typeof daysToDeadline === 'number') {
      daysRemaining = daysToDeadline;
    } else {
      continue; // No deadline
    }

    if (daysRemaining > alertDays) continue;

    // Get member info
    var memberInfo = memberLookup[memberId] || { name: 'Unknown', steward: '' };
    var assignedSteward = steward || memberInfo.steward || 'Unassigned';

    if (!stewardGrievances[assignedSteward]) {
      stewardGrievances[assignedSteward] = [];
    }

    stewardGrievances[assignedSteward].push({
      id: grievanceId,
      memberName: memberInfo.name,
      step: currentStep,
      status: status,
      daysRemaining: daysRemaining,
      nextDue: nextDue
    });
  }

  // Get steward emails from Config sheet
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var stewardEmails = {};
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    // Look for Steward Emails column (assume it's after Stewards column)
    for (var c = 1; c < configData.length; c++) {
      var stewardName = configData[c][CONFIG_COLS.STEWARDS - 1];
      var stewardEmail = configData[c][CONFIG_COLS.STEWARDS]; // Next column
      if (stewardName && stewardEmail && stewardEmail.indexOf('@') !== -1) {
        stewardEmails[stewardName] = stewardEmail;
      }
    }
  }

  // Send emails to each steward
  var emailsSent = 0;
  var adminEmail = Session.getEffectiveUser().getEmail();

  for (var stewardName in stewardGrievances) {
    var grievances = stewardGrievances[stewardName];
    if (grievances.length === 0) continue;

    // Sort by days remaining (most urgent first)
    grievances.sort(function(a, b) { return a.daysRemaining - b.daysRemaining; });

    var email = stewardEmails[stewardName] || adminEmail;

    // Build email body
    var overdue = grievances.filter(function(g) { return g.daysRemaining < 0; });
    var urgent = grievances.filter(function(g) { return g.daysRemaining >= 0 && g.daysRemaining <= 3; });
    var upcoming = grievances.filter(function(g) { return g.daysRemaining > 3; });

    var body = '509 GRIEVANCE DEADLINE ALERT\n';
    body += '====================================\n\n';
    body += 'Steward: ' + stewardName + '\n';
    body += 'Date: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') + '\n\n';

    if (overdue.length > 0) {
      body += 'OVERDUE (' + overdue.length + ')\n';
      body += '---------------------\n';
      for (var o = 0; o < overdue.length; o++) {
        body += '  [!] ' + overdue[o].id + ' - ' + overdue[o].memberName + '\n';
        body += '     Step: ' + overdue[o].step + ' | Status: ' + overdue[o].status + '\n';
        body += '     OVERDUE by ' + Math.abs(overdue[o].daysRemaining) + ' day(s)\n\n';
      }
    }

    if (urgent.length > 0) {
      body += 'URGENT - Due within 3 days (' + urgent.length + ')\n';
      body += '---------------------\n';
      for (var u = 0; u < urgent.length; u++) {
        body += '  [*] ' + urgent[u].id + ' - ' + urgent[u].memberName + '\n';
        body += '     Step: ' + urgent[u].step + ' | Status: ' + urgent[u].status + '\n';
        body += '     Due in ' + urgent[u].daysRemaining + ' day(s)\n\n';
      }
    }

    if (upcoming.length > 0) {
      body += 'UPCOMING - Due within ' + alertDays + ' days (' + upcoming.length + ')\n';
      body += '---------------------\n';
      for (var up = 0; up < upcoming.length; up++) {
        body += '  [-] ' + upcoming[up].id + ' - ' + upcoming[up].memberName + '\n';
        body += '     Step: ' + upcoming[up].step + ' | Due in ' + upcoming[up].daysRemaining + ' day(s)\n\n';
      }
    }

    body += '====================================\n';
    body += 'Dashboard: ' + ss.getUrl() + '\n';
    body += 'Total grievances requiring attention: ' + grievances.length + '\n';

    var subject = (overdue.length > 0 ? 'OVERDUE: ' : '') +
      grievances.length + ' Grievance Deadline(s) - ' + stewardName;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: 'SEIU Local 509 Dashboard'
      });
      emailsSent++;
      // Mask name in logs for privacy
      var maskedSteward = typeof maskName === 'function' ? maskName(stewardName) : '[REDACTED]';
      Logger.log('Sent alert to ' + maskedSteward + ': ' + grievances.length + ' grievances');
    } catch (e) {
      Logger.log('Failed to send steward alert: ' + e.message);
    }
  }

  Logger.log('Steward deadline alerts complete. Sent ' + emailsSent + ' emails.');
  return emailsSent;
}

/**
 * Manual trigger to send steward alerts now
 */
function sendStewardAlertsNow() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('Send Steward Alerts',
    'This will send deadline alert emails to all stewards with upcoming deadlines.\n\n' +
    'Each steward will receive their own personalized digest.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var emailsSent = sendStewardDeadlineAlerts();

  ui.alert('Alerts Sent',
    'Sent ' + emailsSent + ' steward alert email(s).\n\n' +
    'Check the Logs for details.',
    ui.ButtonSet.OK);
}

/**
 * Configure alert settings
 */
function configureAlertSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var currentDays = props.getProperty('alert_days') || '7';
  var stewardAlerts = props.getProperty('steward_alerts_enabled') === 'true';

  var response = ui.prompt('Alert Settings',
    'Current settings:\n' +
    '* Alert window: ' + currentDays + ' days before deadline\n' +
    '* Per-steward alerts: ' + (stewardAlerts ? 'ENABLED' : 'DISABLED') + '\n\n' +
    'Enter new alert window (days before deadline):\n' +
    '(Enter 3, 7, 14, or 30)',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var newDays = parseInt(response.getResponseText(), 10);
  if (isNaN(newDays) || newDays < 1 || newDays > 30) {
    ui.alert('Invalid input. Please enter a number between 1 and 30.');
    return;
  }

  props.setProperty('alert_days', newDays.toString());

  // Ask about per-steward alerts
  var stewardResponse = ui.alert('Per-Steward Alerts',
    'Enable per-steward email alerts?\n\n' +
    'When enabled, each steward receives their own personalized deadline digest.\n\n' +
    'Enable per-steward alerts?',
    ui.ButtonSet.YES_NO);

  props.setProperty('steward_alerts_enabled', stewardResponse === ui.Button.YES ? 'true' : 'false');

  ui.alert('Settings Saved',
    'Alert window: ' + newDays + ' days\n' +
    'Per-steward alerts: ' + (stewardResponse === ui.Button.YES ? 'ENABLED' : 'DISABLED'),
    ui.ButtonSet.OK);
}

// ============================================================================
// SURVEY EMAIL DISTRIBUTION
// ============================================================================

/**
 * Show dialog for sending random survey emails to members
 */
function sendRandomSurveyEmails() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show configuration dialog
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px;max-width:450px}' +
    'h2{color:#5B4B9E;margin-top:0}' +
    '.form-group{margin-bottom:15px}' +
    'label{display:block;font-weight:bold;margin-bottom:5px}' +
    'input,select{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box}' +
    '.info{background:#e8f4fd;padding:12px;border-radius:8px;margin-bottom:15px;font-size:13px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 24px;border:none;border-radius:4px;cursor:pointer;font-size:14px;flex:1}' +
    '.primary{background:#5B4B9E;color:white}' +
    '.secondary{background:#e0e0e0;color:#333}' +
    '</style></head><body><div class="container">' +
    '<h2>Send Survey to Random Members</h2>' +
    '<div class="info">Select how many random members to email. Each member will receive a personalized survey link.</div>' +
    '<div class="form-group"><label>Number of Members to Email</label>' +
    '<select id="count"><option value="5">5 members</option><option value="10" selected>10 members</option>' +
    '<option value="20">20 members</option><option value="50">50 members</option><option value="100">100 members</option></select></div>' +
    '<div class="form-group"><label>Email Subject</label>' +
    '<input type="text" id="subject" value="SEIU Local 509 - Member Satisfaction Survey"></div>' +
    '<div class="form-group"><label>Exclude members emailed in last (days)</label>' +
    '<select id="excludeDays"><option value="0">No exclusion</option><option value="30" selected>30 days</option>' +
    '<option value="60">60 days</option><option value="90">90 days</option></select></div>' +
    '<div class="buttons">' +
    '<button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="primary" onclick="send()">Send Surveys</button></div></div>' +
    '<script>function send(){var opts={count:parseInt(document.getElementById("count").value),' +
    'subject:document.getElementById("subject").value,excludeDays:parseInt(document.getElementById("excludeDays").value)};' +
    'google.script.run.withSuccessHandler(function(r){alert(r);google.script.host.close()})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message)}).executeSendRandomSurveyEmails(opts)}</script></body></html>'
  ).setWidth(500).setHeight(450);

  ui.showModalDialog(html, 'Send Random Survey Emails');
}

/**
 * Execute sending random survey emails
 * @param {Object} opts - Options {count, subject, excludeDays}
 * @returns {string} Result message
 */
function executeSendRandomSurveyEmails(opts) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!memberSheet) throw new Error('Member Directory not found');

  // Get all members with valid emails
  var memberData = memberSheet.getDataRange().getValues();
  var headers = memberData[0];
  var emailCol = MEMBER_COLS.EMAIL - 1;
  var memberIdCol = MEMBER_COLS.MEMBER_ID - 1;
  var firstNameCol = MEMBER_COLS.FIRST_NAME - 1;
  var lastNameCol = MEMBER_COLS.LAST_NAME - 1;

  // Get survey email log from Config (if exists)
  var surveyLogCol = 50; // Column AX for survey email log
  var surveyLog = {};
  try {
    var logData = configSheet.getRange(2, surveyLogCol, configSheet.getLastRow() - 1, 2).getValues();
    logData.forEach(function(row) {
      if (row[0]) surveyLog[row[0]] = new Date(row[1]);
    });
  } catch(e) { /* No log yet */ }

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - opts.excludeDays);

  // Build list of eligible members
  var eligibleMembers = [];
  for (var i = 1; i < memberData.length; i++) {
    var row = memberData[i];
    var memberId = row[memberIdCol];
    var email = row[emailCol];
    var firstName = row[firstNameCol];

    // Skip if no valid member ID or email
    if (!memberId || !email || !email.toString().includes('@')) continue;

    // Skip if recently emailed
    if (opts.excludeDays > 0 && surveyLog[memberId] && surveyLog[memberId] > cutoffDate) continue;

    eligibleMembers.push({
      memberId: memberId,
      email: email,
      firstName: firstName,
      lastName: row[lastNameCol]
    });
  }

  if (eligibleMembers.length === 0) {
    return 'No eligible members found. All members may have been recently emailed.';
  }

  // Shuffle and select random members
  var shuffled = eligibleMembers.sort(function() { return 0.5 - Math.random(); });
  var selected = shuffled.slice(0, Math.min(opts.count, shuffled.length));

  // Send emails
  var sent = 0;
  var errors = [];
  var formUrl = SATISFACTION_FORM_CONFIG.FORM_URL;
  var newLogEntries = [];

  selected.forEach(function(member) {
    try {
      var personalizedUrl = formUrl + '?memberId=' + encodeURIComponent(member.memberId);
      var body = 'Dear ' + member.firstName + ',\n\n' +
        'We value your feedback! Please take a few minutes to complete our Member Satisfaction Survey.\n\n' +
        'Your responses help us improve union services and representation.\n\n' +
        'Survey Link: ' + personalizedUrl + '\n\n' +
        'Your Member ID: ' + member.memberId + '\n' +
        '(You will need this to verify your membership when submitting)\n\n' +
        'Thank you for being a member!\n\n' +
        'SEIU Local 509';

      MailApp.sendEmail({
        to: member.email,
        subject: opts.subject,
        body: body,
        name: 'SEIU Local 509 Dashboard'
      });

      sent++;
      newLogEntries.push([member.memberId, new Date()]);
    } catch(e) {
      errors.push(member.firstName + ' ' + member.lastName + ': ' + e.message);
    }
  });

  // Update survey email log
  if (newLogEntries.length > 0) {
    var nextRow = Object.keys(surveyLog).length + 2;
    configSheet.getRange(nextRow, surveyLogCol, newLogEntries.length, 2).setValues(newLogEntries);
  }

  var result = 'Sent ' + sent + ' survey emails';
  if (errors.length > 0) {
    result += '\n\n' + errors.length + ' errors:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) result += '\n...and ' + (errors.length - 5) + ' more';
  }

  return result;
}

// ============================================================================
// MEMBER VALIDATION UTILITIES
// ============================================================================

/**
 * Validate that an email belongs to a member in the directory
 * @param {string} email - Email to validate
 * @returns {Object|null} Member info if valid, null otherwise
 */
function validateMemberEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || !email) return null;

  var data = memberSheet.getDataRange().getValues();
  var emailCol = MEMBER_COLS.EMAIL - 1;
  var memberIdCol = MEMBER_COLS.MEMBER_ID - 1;
  var firstNameCol = MEMBER_COLS.FIRST_NAME - 1;
  var lastNameCol = MEMBER_COLS.LAST_NAME - 1;

  email = email.toString().toLowerCase().trim();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][emailCol] || '').toString().toLowerCase().trim();
    if (rowEmail === email) {
      return {
        memberId: data[i][memberIdCol],
        firstName: data[i][firstNameCol],
        lastName: data[i][lastNameCol],
        email: rowEmail
      };
    }
  }

  return null;
}

// ============================================================================
// QUARTER UTILITIES FOR NOTIFICATIONS
// ============================================================================

/**
 * Get the current quarter string (e.g., "2026-Q1")
 * @returns {string} Quarter string
 */
function getCurrentQuarter() {
  var now = new Date();
  var quarter = Math.floor(now.getMonth() / 3) + 1;
  return now.getFullYear() + '-Q' + quarter;
}

/**
 * Get quarter string from a date
 * @param {Date} date - Date to get quarter from
 * @returns {string} Quarter string
 */
function getQuarterFromDate(date) {
  var d = new Date(date);
  var quarter = Math.floor(d.getMonth() / 3) + 1;
  return d.getFullYear() + '-Q' + quarter;
}



/**
 * ============================================================================
 * AUDIT LOG MODULE (08i_AuditLog.gs)
 * ============================================================================
 *
 * This module provides audit logging functionality for tracking changes
 * to the Member Directory and Grievance Log sheets.
 *
 * Features:
 * - Automatic change tracking via onEdit trigger
 * - Audit log sheet setup and management
 * - History viewing and cleanup utilities
 * - Record-specific audit trail retrieval
 *
 * Dependencies:
 * - SHEETS constant (for sheet names)
 * - COLORS constant (for formatting)
 * - logAuditEvent() function (from core module)
 *
 * @author Union Membership System
 * @version 1.0.0
 * ============================================================================
 */

// ============================================================================
// AUDIT LOG SHEET SETUP
// ============================================================================

/**
 * Sets up the Audit Log sheet with proper headers and formatting
 * Creates the sheet if it doesn't exist, or clears and reformats if it does
 */
function setupAuditLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.AUDIT_LOG);
  }

  sheet.clear();

  // Headers
  var headers = [
    'Timestamp',
    'User Email',
    'Sheet',
    'Row',
    'Column',
    'Field Name',
    'Old Value',
    'New Value',
    'Record ID',
    'Action Type'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(COLORS.PRIMARY_PURPLE);
  headerRange.setFontColor(COLORS.WHITE);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 160); // Timestamp
  sheet.setColumnWidth(2, 200); // User Email
  sheet.setColumnWidth(3, 120); // Sheet
  sheet.setColumnWidth(4, 50);  // Row
  sheet.setColumnWidth(5, 50);  // Column
  sheet.setColumnWidth(6, 150); // Field Name
  sheet.setColumnWidth(7, 200); // Old Value
  sheet.setColumnWidth(8, 200); // New Value
  sheet.setColumnWidth(9, 100); // Record ID
  sheet.setColumnWidth(10, 100); // Action Type

  // Hide the sheet
  sheet.hideSheet();

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit log sheet created and hidden.', '✅ Setup Complete', 3);
}

// ============================================================================
// AUDIT TRIGGER FUNCTIONS
// ============================================================================

/**
 * onEdit trigger for audit logging
 * Tracks changes to Member Directory and Grievance Log
 * @param {Object} e - The edit event object
 */
function onEditAudit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only track changes to Member Directory and Grievance Log
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Skip header row
  if (row < 2) return;

  var oldValue = e.oldValue || '';
  var newValue = e.value || '';

  // Skip if no actual change
  if (oldValue === newValue) return;

  // Get field name from header
  var fieldName = sheet.getRange(1, col).getValue() || ('Column ' + col);

  // Get record ID (column A for both sheets)
  var recordId = sheet.getRange(row, 1).getValue() || '';

  // Determine action type
  var actionType = 'Edit';
  if (!oldValue && newValue) {
    actionType = 'Create';
  } else if (oldValue && !newValue) {
    actionType = 'Delete';
  }

  logAuditEvent(actionType + '_' + sheetName.toUpperCase().replace(/\s+/g, '_'), {
    sheet: sheetName,
    row: row,
    col: col,
    field: fieldName,
    oldValue: oldValue,
    newValue: newValue,
    recordId: recordId
  });
}

/**
 * Install the audit trigger
 * Sets up automatic change tracking for Member Directory and Grievance Log
 */
function installAuditTrigger() {
  // Remove existing audit triggers
  removeAuditTrigger();

  // Create new onEdit trigger
  ScriptApp.newTrigger('onEditAudit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  // Ensure audit sheet exists
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.AUDIT_LOG)) {
    setupAuditLogSheet();
  }

  SpreadsheetApp.getUi().alert('✅ Audit Tracking Enabled',
    'All changes to Member Directory and Grievance Log will now be logged.\n\n' +
    'View the audit log via:\n⚙️ Administrator > 📋 Audit Log > 📋 View Audit Log',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Remove the audit trigger
 * Disables automatic change tracking
 */
function removeAuditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditAudit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit tracking disabled.', '🚫 Disabled', 3);
}

// ============================================================================
// AUDIT LOG VIEWING AND MANAGEMENT
// ============================================================================

/**
 * View the audit log sheet
 * Shows the hidden audit log and sorts entries by newest first
 */
function viewAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    var response = SpreadsheetApp.getUi().alert('📋 Audit Log Not Found',
      'The audit log sheet does not exist yet.\n\nWould you like to create it now?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (response === SpreadsheetApp.getUi().Button.YES) {
      setupAuditLogSheet();
      sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    } else {
      return;
    }
  }

  // Show the hidden sheet temporarily
  sheet.showSheet();
  ss.setActiveSheet(sheet);

  // Sort by timestamp descending (newest first)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).sort({column: 1, ascending: false});
  }

  SpreadsheetApp.getUi().alert('📋 Audit Log',
    'Viewing audit log.\n\n' +
    'Total entries: ' + Math.max(0, sheet.getLastRow() - 1) + '\n\n' +
    'The sheet will be hidden again when you navigate away.\n' +
    'To keep it visible, right-click the tab and select "Unhide".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear audit entries older than 30 days
 * Prompts for confirmation before deleting old entries
 */
function clearOldAuditEntries() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('🗑️ Clear Old Audit Entries',
    'This will delete all audit entries older than 30 days.\n\n' +
    'This action cannot be undone.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No audit entries to clear.');
    return;
  }

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Find rows older than 30 days (skip header)
  for (var i = data.length - 1; i >= 1; i--) {
    var timestamp = data[i][0];
    if (timestamp instanceof Date && timestamp < cutoffDate) {
      rowsToDelete.push(i + 1); // +1 for 1-indexed rows
    }
  }

  // Delete rows from bottom to top to maintain correct indices
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  ui.alert('✅ Cleanup Complete',
    'Deleted ' + rowsToDelete.length + ' entries older than 30 days.\n\n' +
    'Remaining entries: ' + Math.max(0, sheet.getLastRow() - 1),
    ui.ButtonSet.OK);
}

// ============================================================================
// AUDIT HISTORY RETRIEVAL
// ============================================================================

/**
 * Get audit summary for a specific record
 * @param {string} recordId - Member ID or Grievance ID
 * @returns {Array} Array of audit entries for this record
 */
function getAuditHistory(recordId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === recordId) { // Column I is Record ID
      history.push({
        timestamp: data[i][0],
        user: data[i][1],
        field: data[i][5],
        oldValue: data[i][6],
        newValue: data[i][7],
        action: data[i][9]
      });
    }
  }

  return history;
}



/**
 * ============================================================================
 * 08j_CalcSheets.gs - Hidden Calculation Sheet Setup & Management
 * ============================================================================
 *
 * This module contains all functions related to setting up and managing
 * the hidden calculation sheets that power the dashboard's "self-healing"
 * formula system. These sheets contain complex formulas that aggregate,
 * calculate, and cross-reference data across the dashboard.
 *
 * Hidden Sheets Managed:
 * - _Grievance_Calc: Grievance -> Member Directory data sync
 * - _Grievance_Formulas: Self-healing Grievance Log formulas
 * - _Member_Lookup: Member -> Grievance Log data sync
 * - _Steward_Contact_Calc: Steward contact tracking metrics
 * - _Steward_Performance_Calc: Steward performance scores
 * - _Dashboard_Calc: Dashboard summary statistics
 * - _CalcMembers: Member statistics and lookups
 * - _CalcGrievances: Grievance aggregations
 * - _CalcDeadlines: Deadline calculations and alerts
 * - _CalcStats: Dashboard-wide statistics
 * - _CalcSync: Cross-sheet synchronization
 * - _CalcFormulas: Named formula references
 *
 * @fileoverview Hidden sheet and formula management for 509 Dashboard
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// LIVE FORMULA SYNC FUNCTIONS
// ============================================================================

/**
 * Syncs grievance data to Member Directory using static values
 * Menu Location: Dashboard menu
 */
function setupLiveGrievanceFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory not found.');
    return;
  }

  ss.toast('Syncing grievance data to Member Directory...', '🔄 Sync', 3);

  // Use sync function to populate with static values (no formulas)
  syncGrievanceToMemberDirectory();

  ss.toast('Grievance data synced! Columns AB-AD updated with static values.', '✅ Success', 3);
}

/**
 * Remove Member ID dropdown from Grievance Log
 * Clears any existing data validation to allow free text entry
 */
function setupGrievanceMemberDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Grievance Log not found.');
    return;
  }

  ss.toast('Removing Member ID dropdown...', '🔄 Setup', 3);

  // Clear any existing data validation from Member ID column (column B, rows 2-1000)
  // This allows free text entry for Member ID
  grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, 998, 1).clearDataValidations();

  ss.toast('Member ID dropdown removed - free text entry enabled!', '✅ Success', 3);
}

// ============================================================================
// HIDDEN SHEET 1: _Grievance_Calc
// Source: Grievance Log -> Destination: Member Directory (AB-AD)
// ============================================================================

/**
 * Setup the _Grievance_Calc hidden sheet with self-healing formulas
 * Calculates: Has Open Grievance, Grievance Status, Next Deadline per member
 */
function setupGrievanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'Has Open Grievance', 'Grievance Status', 'Days to Deadline', 'Total Count', 'Win Rate %', 'Last Grievance Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for dynamic formulas
  var memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gNextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);

  // Formula for Member IDs (Column A) - pulls unique member IDs from Member Directory
  var memberIdFormula = '=IFERROR(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + '<>"Member ID"),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // Formulas for calculations (using ARRAYFORMULA for efficiency)
  // Column B: Has Open Grievance
  var hasOpenFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Yes","No")))';
  sheet.getRange('B2').setFormula(hasOpenFormula);

  // Column C: Grievance Status (most urgent: Open > Pending Info, blank if all closed)
  var statusFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")>0,"Open",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Pending Info",""))))';
  sheet.getRange('C2').setFormula(statusFormula);

  // Column D: Days to Deadline (minimum/most urgent deadline for open grievances only)
  // Excludes all closed statuses: Closed, Settled, Withdrawn, Denied, Won
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var deadlineFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MINIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Closed",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Settled",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Withdrawn",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Denied",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Won"),"")))';
  sheet.getRange('D2').setFormula(deadlineFormula);

  // Column E: Total Grievance Count
  var countFormula = '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)))';
  sheet.getRange('E2').setFormula(countFormula);

  // Column F: Win Rate %
  var winRateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)*100,0)))';
  sheet.getRange('F2').setFormula(winRateFormula);

  // Column G: Last Grievance Date
  var lastDateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MAXIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A),"")))';
  sheet.getRange('G2').setFormula(lastDateFormula);

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Calc sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 2: _Grievance_Formulas (SELF-HEALING)
// Source: Grievance Log -> Destination: Grievance Log (calculated columns)
// This sheet contains all auto-calculated formulas and syncs them back
// ============================================================================

/**
 * Setup the _Grievance_Formulas hidden sheet with self-healing formulas
 * Calculates: First Name, Last Name, Email, Unit, Location, Steward (from Member Dir)
 *            Filing Deadline, Step I-III dates, Days Open, Next Action Due, Days to Deadline
 */
function setupGrievanceFormulasSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FORMULAS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_FORMULAS);
  }

  sheet.clear();

  // Headers matching Grievance Log columns that need formulas
  var headers = [
    'Row Index',           // A - For tracking which row in Grievance Log
    'Member ID',           // B - From Grievance Log
    'First Name',          // C - Lookup from Member Directory
    'Last Name',           // D - Lookup from Member Directory
    'Incident Date',       // E - From Grievance Log
    'Date Filed',          // F - From Grievance Log
    'Step I Rcvd',         // G - From Grievance Log
    'Step II Appeal Filed',// H - From Grievance Log
    'Step II Rcvd',        // I - From Grievance Log
    'Status',              // J - From Grievance Log
    'Current Step',        // K - From Grievance Log
    'Date Closed',         // L - From Grievance Log
    'Filing Deadline',     // M - CALCULATED
    'Step I Due',          // N - CALCULATED
    'Step II Appeal Due',  // O - CALCULATED
    'Step II Due',         // P - CALCULATED
    'Step III Appeal Due', // Q - CALCULATED
    'Days Open',           // R - CALCULATED
    'Next Action Due',     // S - CALCULATED
    'Days to Deadline',    // T - CALCULATED
    'Member Email',        // U - Lookup from Member Directory
    'Unit',                // V - Lookup from Member Directory
    'Location',            // W - Lookup from Member Directory
    'Steward'              // X - Lookup from Member Directory
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for Grievance Log source data
  var gGrievanceIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);     // A
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);           // B
  var gIncidentDateCol = getColumnLetter(GRIEVANCE_COLS.INCIDENT_DATE);   // G
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);         // I
  var gStep1RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP1_RCVD);         // K
  var gStep2AppealFiledCol = getColumnLetter(GRIEVANCE_COLS.STEP2_APPEAL_FILED); // M
  var gStep2RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP2_RCVD);         // O
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);                // E
  var gCurrentStepCol = getColumnLetter(GRIEVANCE_COLS.CURRENT_STEP);     // F
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);       // R

  // Member Directory columns for lookups
  var mMemberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);
  var memberRange = "'" + SHEETS.MEMBER_DIR + "'!" + mMemberIdCol + ":" + mStewardCol;

  // Column A: Row Index (ROW()-1 to match Grievance Log rows)
  // Pull unique grievance IDs to create row mapping
  sheet.getRange('A2').setFormula(
    '=IFERROR(FILTER(ROW(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + ')-1,' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + '<>""),"")'
  );

  // Column B: Member ID (from Grievance Log)
  sheet.getRange('B2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A+1)))'
  );

  // Column C: First Name (VLOOKUP from Member Directory)
  sheet.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.FIRST_NAME + ',FALSE),"")))'
  );

  // Column D: Last Name (VLOOKUP from Member Directory)
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.LAST_NAME + ',FALSE),"")))'
  );

  // Column E: Incident Date (from Grievance Log)
  sheet.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIncidentDateCol + ':' + gIncidentDateCol + ',A2:A+1)))'
  );

  // Column F: Date Filed (from Grievance Log)
  sheet.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',A2:A+1)))'
  );

  // Column G: Step I Rcvd (from Grievance Log)
  sheet.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep1RcvdCol + ':' + gStep1RcvdCol + ',A2:A+1)))'
  );

  // Column H: Step II Appeal Filed (from Grievance Log)
  sheet.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2AppealFiledCol + ':' + gStep2AppealFiledCol + ',A2:A+1)))'
  );

  // Column I: Step II Rcvd (from Grievance Log)
  sheet.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2RcvdCol + ':' + gStep2RcvdCol + ',A2:A+1)))'
  );

  // Column J: Status (from Grievance Log)
  sheet.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',A2:A+1)))'
  );

  // Column K: Current Step (from Grievance Log)
  sheet.getRange('K2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gCurrentStepCol + ':' + gCurrentStepCol + ',A2:A+1)))'
  );

  // Column L: Date Closed (from Grievance Log)
  sheet.getRange('L2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',A2:A+1)))'
  );

  // =========== CALCULATED COLUMNS ===========

  // Column M: Filing Deadline = Incident Date + 21 days
  sheet.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(E2:E="","",E2:E+21))'
  );

  // Column N: Step I Due = Date Filed + 30 days
  sheet.getRange('N2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",F2:F+30))'
  );

  // Column O: Step II Appeal Due = Step I Rcvd + 10 days
  sheet.getRange('O2').setFormula(
    '=ARRAYFORMULA(IF(G2:G="","",G2:G+10))'
  );

  // Column P: Step II Due = Step II Appeal Filed + 30 days
  sheet.getRange('P2').setFormula(
    '=ARRAYFORMULA(IF(H2:H="","",H2:H+30))'
  );

  // Column Q: Step III Appeal Due = Step II Rcvd + 30 days
  sheet.getRange('Q2').setFormula(
    '=ARRAYFORMULA(IF(I2:I="","",I2:I+30))'
  );

  // Column R: Days Open = IF closed: Date Closed - Date Filed, ELSE: Today - Date Filed
  sheet.getRange('R2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",IF(L2:L<>"",L2:L-F2:F,TODAY()-F2:F)))'
  );

  // Column S: Next Action Due = Based on current step and status
  // If closed status, leave blank; otherwise return appropriate deadline
  sheet.getRange('S2').setFormula(
    '=ARRAYFORMULA(IF(J2:J="","",' +
    'IF(OR(J2:J="Settled",J2:J="Withdrawn",J2:J="Denied",J2:J="Won",J2:J="Closed"),"",' +
    'IF(K2:K="Informal",M2:M,' +
    'IF(K2:K="Step I",N2:N,' +
    'IF(K2:K="Step II",P2:P,' +
    'Q2:Q))))))'
  );

  // Column T: Days to Deadline = Next Action Due - Today
  sheet.getRange('T2').setFormula(
    '=ARRAYFORMULA(IF(S2:S="","",S2:S-TODAY()))'
  );

  // =========== MEMBER LOOKUP COLUMNS ===========

  // Column U: Member Email (VLOOKUP from Member Directory)
  sheet.getRange('U2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.EMAIL + ',FALSE),"")))'
  );

  // Column V: Unit (VLOOKUP from Member Directory)
  sheet.getRange('V2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.UNIT + ',FALSE),"")))'
  );

  // Column W: Location (VLOOKUP from Member Directory)
  sheet.getRange('W2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.WORK_LOCATION + ',FALSE),"")))'
  );

  // Column X: Steward (VLOOKUP from Member Directory)
  sheet.getRange('X2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.ASSIGNED_STEWARD + ',FALSE),"")))'
  );

  // Format date columns (MM/dd/yyyy)
  sheet.getRange('E:E').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('F:F').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('G:G').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('H:H').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('I:I').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('L:L').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('M:M').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('N:N').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('O:O').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('P:P').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('Q:Q').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('S:S').setNumberFormat('MM/dd/yyyy');

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Formulas sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 3: _Member_Lookup
// Source: Member Directory -> Destination: Grievance Log (C,D,X-AA)
// ============================================================================

/**
 * Setup the _Member_Lookup hidden sheet with self-healing formulas
 * Looks up: First Name, Last Name, Email, Unit, Location, Steward from Member Directory
 */
function setupMemberLookupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_LOOKUP);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MEMBER_LOOKUP);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'First Name', 'Last Name', 'Email', 'Unit', 'Location', 'Assigned Steward'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mFirstCol = getColumnLetter(MEMBER_COLS.FIRST_NAME);
  var mLastCol = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var mEmailCol = getColumnLetter(MEMBER_COLS.EMAIL);
  var mUnitCol = getColumnLetter(MEMBER_COLS.UNIT);
  var mLocCol = getColumnLetter(MEMBER_COLS.WORK_LOCATION);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);

  // Formula to get unique member IDs from Grievance Log
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var memberIdFormula = '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"Member ID",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"")),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // VLOOKUP formulas for member data
  var vlookupBase = 'VLOOKUP(A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mStewardCol + ',';

  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '2,FALSE),"")))'); // First Name
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '3,FALSE),"")))'); // Last Name
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '8,FALSE),"")))'); // Email
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '6,FALSE),"")))'); // Unit
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '5,FALSE),"")))'); // Location
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '16,FALSE),"")))'); // Steward

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Member_Lookup sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 4: _Steward_Contact_Calc
// Source: Member Directory (Y-AA) -> Aggregates steward contact tracking metrics
// ============================================================================

/**
 * Setup the _Steward_Contact_Calc hidden sheet with self-healing formulas
 * Tracks and aggregates steward contact data from Member Directory
 */
function setupStewardContactCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_CONTACT_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_CONTACT_CALC);
  }

  sheet.clear();

  // Headers for steward contact summary (5 columns)
  var headers = ['Steward Name', 'Total Contacts', 'Contacts This Month', 'Contacts Last 7 Days', 'Last Contact Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for formulas
  var mContactStewardCol = getColumnLetter(MEMBER_COLS.CONTACT_STEWARD);
  var mContactDateCol = getColumnLetter(MEMBER_COLS.RECENT_CONTACT_DATE);

  // Column A: Unique steward names who have made contacts
  sheet.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + '<>""))),)');

  // Column B: Total contacts per steward
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A)))');

  // Column C: Contacts this month
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))))');

  // Column D: Contacts last 7 days
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&(TODAY()-7))))');

  // Column E: Most recent contact date for this steward
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(TEXT(MAXIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A),"MM/dd/yyyy"),"-")))');

  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  sheet.hideSheet();
  Logger.log('_Steward_Contact_Calc sheet setup complete with live formulas');
}

// ============================================================================
// HIDDEN SHEET 5: _Dashboard_Calc
// Source: Member Directory + Grievance Log -> Dashboard Summary Statistics
// ============================================================================

/**
 * Setup the _Dashboard_Calc hidden sheet with self-healing formulas
 * Calculates key dashboard metrics that auto-update
 */
function setupDashboardCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.DASHBOARD_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Column references
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);

  // Metrics with formulas (15 key metrics)
  // Note: Using COUNTIF with "M*" and "G*" patterns to only count valid IDs (ignores blank rows)
  var metrics = [
    ['Total Members', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ',"M*")', 'Total union members in directory'],
    ['Active Stewards', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")', 'Members marked as stewards'],
    ['Total Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ',"G*")', 'All grievances filed'],
    ['Open Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")', 'Currently open cases'],
    ['Pending Info', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")', 'Cases awaiting information'],
    ['Settled', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")', 'Cases settled'],
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")', 'Cases won (full or partial)'],
    ['Denied', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")', 'Cases denied'],
    ['Withdrawn', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Withdrawn")', 'Cases withdrawn'],
    ['Win Rate %', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*"))*100,1),0)', 'Wins / (Wins + Settled + Denied)'],
    ['Avg Days to Resolution', '=IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<>",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)', 'Average days for closed cases'],
    ['Overdue Cases', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")', 'Cases past deadline'],
    ['Due This Week', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")', 'Cases due in next 7 days'],
    ['Filed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())', 'Grievances filed this month'],
    ['Closed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())', 'Grievances closed this month']
  ];

  for (var i = 0; i < metrics.length; i++) {
    sheet.getRange(i + 2, 1).setValue(metrics[i][0]);
    sheet.getRange(i + 2, 2).setFormula(metrics[i][1]);
    sheet.getRange(i + 2, 3).setValue(metrics[i][2]);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 300);

  sheet.hideSheet();
  Logger.log('_Dashboard_Calc sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 6: _Steward_Performance_Calc
// Source: Grievance Log -> Steward Performance Metrics
// ============================================================================

/**
 * Setup the _Steward_Performance_Calc hidden sheet
 * Calculates detailed steward performance metrics
 */
function setupStewardPerformanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_PERFORMANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Steward', 'Total Cases', 'Active', 'Closed', 'Won', 'Win Rate %', 'Avg Days', 'Overdue', 'Due This Week', 'Performance Score'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  var gStewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);

  // Get unique stewards
  sheet.getRange('A2').setFormula(
    '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"Assigned Steward",' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"")),"")'
  );

  // Total Cases
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A)))');

  // Active Cases (Open + Pending Info)
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")))');

  // Closed Cases
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))');

  // Won Cases
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")))');

  // Win Rate
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(E2:E/D2:D*100,1),0)))');

  // Avg Days
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)))');

  // Overdue
  sheet.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")))');

  // Due This Week
  sheet.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")))');

  // Performance Score (weighted: Win Rate * 0.4 + (100 - Overdue%) * 0.3 + (100 - AvgDays/60*100) * 0.3)
  sheet.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(F2:F*0.4 + (100-IFERROR(H2:H/C2:C*100,0))*0.3 + MAX(0,100-G2:G/60*100)*0.3,1)))');

  sheet.hideSheet();
  Logger.log('_Steward_Performance_Calc sheet setup complete');
}

// ============================================================================
// MASTER SETUP & REPAIR FUNCTIONS
// ============================================================================

/**
 * Setup all hidden calculation sheets
 * @returns {Object} Result object with created and repaired counts
 */
function setupAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Setting up hidden calculation sheets...', '🔧 Setup', 3);

  var created = 0;
  var repaired = 0;

  // Core grievance/member calculation sheets (7 total)
  // Each function creates the sheet if missing or updates if exists
  try { setupGrievanceCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupGrievanceFormulasSheet(); created++; } catch (e) { repaired++; }
  try { setupMemberLookupSheet(); created++; } catch (e) { repaired++; }
  try { setupStewardContactCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupDashboardCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupStewardPerformanceCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupChecklistCalcSheet(); created++; } catch (e) { repaired++; }

  ss.toast('All 7 hidden sheets created!', '✅ Success', 3);

  return { created: created, repaired: repaired, success: true };
}

/**
 * Repair all hidden sheets - recreates formulas and syncs data
 */
function repairAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ss.toast('Repairing hidden sheets...', '🔧 Repair', 3);

  // Recreate all hidden sheets with formulas
  setupAllHiddenSheets();

  // Install trigger (quick mode - no dialog)
  installAutoSyncTriggerQuick();

  // Run initial sync
  ss.toast('Running initial data sync...', '🔧 Sync', 3);
  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();
  syncChecklistCalcToGrievanceLog();

  // Repair checkboxes
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  ss.toast('Hidden sheets repaired and synced!', '✅ Success', 5);
  ui.alert('✅ Repair Complete',
    'Hidden calculation sheets have been repaired:\n\n' +
    '• 7 hidden sheets recreated with self-healing formulas\n' +
    '• Auto-sync trigger installed\n' +
    '• All data synced (grievances, members, dashboard, checklists)\n' +
    '• Checkboxes repaired in Grievance Log and Member Directory\n\n' +
    'Data will now auto-sync when you edit Member Directory or Grievance Log.\n' +
    'Formulas cannot be accidentally erased - they are stored in hidden sheets.',
    ui.ButtonSet.OK);

  return { repaired: 7, success: true };
}

/**
 * Verify all hidden sheets and triggers
 */
function verifyHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var report = [];

  report.push('🔍 HIDDEN SHEET VERIFICATION');
  report.push('============================');
  report.push('');

  // Check each hidden sheet (7 hidden sheets)
  var hiddenSheets = [
    {name: SHEETS.GRIEVANCE_CALC, purpose: 'Grievance -> Member Directory'},
    {name: SHEETS.GRIEVANCE_FORMULAS, purpose: 'Self-healing Grievance formulas'},
    {name: SHEETS.MEMBER_LOOKUP, purpose: 'Member -> Grievance Log'},
    {name: SHEETS.STEWARD_CONTACT_CALC, purpose: 'Steward contact tracking'},
    {name: SHEETS.DASHBOARD_CALC, purpose: 'Dashboard summary metrics'},
    {name: SHEETS.STEWARD_PERFORMANCE_CALC, purpose: 'Steward performance scores'},
    {name: HIDDEN_SHEETS.CHECKLIST_CALC, purpose: 'Case Checklist progress calculations'}
  ];

  report.push('📋 HIDDEN SHEETS:');
  hiddenSheets.forEach(function(hs) {
    var sheet = ss.getSheetByName(hs.name);
    if (sheet) {
      var isHidden = sheet.isSheetHidden();
      var hasData = sheet.getLastRow() > 1;
      var status = isHidden && hasData ? '✅' : (sheet ? '⚠️' : '❌');
      report.push('  ' + status + ' ' + hs.name);
      report.push('      Hidden: ' + (isHidden ? 'Yes' : 'NO - Should be hidden'));
      report.push('      Has formulas: ' + (hasData ? 'Yes' : 'No'));
    } else {
      report.push('  ❌ ' + hs.name + ' - NOT FOUND');
    }
  });

  report.push('');

  // Check triggers
  report.push('⚡ AUTO-SYNC TRIGGER:');
  var triggers = ScriptApp.getProjectTriggers();
  var hasAutoSync = false;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      hasAutoSync = true;
      report.push('  ✅ onEditAutoSync trigger installed');
    }
  });
  if (!hasAutoSync) {
    report.push('  ❌ onEditAutoSync trigger NOT installed');
    report.push('     Run: installAutoSyncTrigger()');
  }

  report.push('');
  report.push('============================');

  ui.alert('Hidden Sheet Verification', report.join('\n'), ui.ButtonSet.OK);
  Logger.log(report.join('\n'));
}

/**
 * Refresh all formulas (force recalculation)
 */
function refreshAllHiddenFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Touch each hidden sheet to force recalc (6 hidden sheets)
  var hiddenSheetNames = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_CONTACT_CALC,
    SHEETS.DASHBOARD_CALC,
    SHEETS.STEWARD_PERFORMANCE_CALC
  ];

  hiddenSheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      // Force recalc by getting values
      sheet.getDataRange().getValues();
    }
  });

  // Then sync
  syncAllData();

  // Repair checkboxes
  repairGrievanceCheckboxes();

  ss.toast('Formulas refreshed and data synced!', '✅ Success', 3);
}

// ============================================================================
// FORMULA SERVICE - INDIVIDUAL CALCULATION SHEET SETUP
// These functions set up additional calculation sheets for specialized purposes
// ============================================================================

/**
 * Sets up the _CalcMembers hidden sheet
 * Contains member statistics and lookup tables
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcMembersSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;

  // Header row
  sheet.getRange('A1').setValue('Member Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Total Members
  sheet.getRange('A2').setValue('Total Members');
  sheet.getRange('B2').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  // Active Members
  sheet.getRange('A3').setValue('Active Members');
  sheet.getRange('B3').setFormula(
    `=COUNTIF('${memberSheetName}'!K:K,"Active")`
  );

  // Members by Department (dynamic list)
  sheet.getRange('A5').setValue('Department');
  sheet.getRange('B5').setValue('Count');
  sheet.getRange('A5:B5').setFontWeight('bold');

  sheet.getRange('A6').setFormula(
    `=UNIQUE(FILTER('${memberSheetName}'!E:E,'${memberSheetName}'!E:E<>"Department",'${memberSheetName}'!E:E<>""))`
  );

  sheet.getRange('B6').setFormula(
    `=ARRAYFORMULA(IF(A6:A<>"",COUNTIF('${memberSheetName}'!E:E,A6:A),""))`
  );

  // Union Status breakdown
  sheet.getRange('D2').setValue('Union Status');
  sheet.getRange('E2').setValue('Count');
  sheet.getRange('D2:E2').setFontWeight('bold');

  sheet.getRange('D3').setValue('Full Member');
  sheet.getRange('E3').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Full Member")`
  );

  sheet.getRange('D4').setValue('Agency Fee');
  sheet.getRange('E4').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Agency Fee")`
  );

  sheet.getRange('D5').setValue('Non-Member');
  sheet.getRange('E5').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Non-Member")`
  );

  // Lookup helper for member names
  sheet.getRange('G1').setValue('ID->Name Lookup');
  sheet.getRange('G1').setFontWeight('bold');
  sheet.getRange('G2').setFormula(
    `=ARRAYFORMULA(IF('${memberSheetName}'!A2:A<>"",` +
    `'${memberSheetName}'!A2:A&"|"&'${memberSheetName}'!B2:B&" "&'${memberSheetName}'!C2:C,""))`
  );
}

/**
 * Sets up the _CalcGrievances hidden sheet
 * Contains grievance aggregations and summaries
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcGrievancesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Grievance Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Status counts
  sheet.getRange('A3').setValue('Status');
  sheet.getRange('B3').setValue('Count');
  sheet.getRange('A3:B3').setFontWeight('bold');

  const statuses = Object.values(GRIEVANCE_STATUS);
  statuses.forEach((status, index) => {
    sheet.getRange(4 + index, 1).setValue(status);
    sheet.getRange(4 + index, 2).setFormula(
      `=COUNTIF('${grievanceSheetName}'!W:W,"${status}")`
    );
  });

  // Grievances by Type
  sheet.getRange('D3').setValue('Type');
  sheet.getRange('E3').setValue('Count');
  sheet.getRange('D3:E3').setFontWeight('bold');

  sheet.getRange('D4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!E:E,` +
    `'${grievanceSheetName}'!E:E<>"Grievance Type",'${grievanceSheetName}'!E:E<>""))`
  );

  sheet.getRange('E4').setFormula(
    `=ARRAYFORMULA(IF(D4:D<>"",COUNTIF('${grievanceSheetName}'!E:E,D4:D),""))`
  );

  // Grievances by Current Step
  sheet.getRange('G3').setValue('Step');
  sheet.getRange('H3').setValue('Count');
  sheet.getRange('G3:H3').setFontWeight('bold');

  for (let step = 1; step <= 4; step++) {
    sheet.getRange(3 + step, 7).setValue(`Step ${step}`);
    sheet.getRange(3 + step, 8).setFormula(
      `=COUNTIF('${grievanceSheetName}'!H:H,${step})`
    );
  }

  // Monthly filing trend (last 12 months)
  sheet.getRange('A15').setValue('Monthly Filings');
  sheet.getRange('A15').setFontWeight('bold');

  sheet.getRange('A16').setValue('Month');
  sheet.getRange('B16').setValue('Filings');
  sheet.getRange('A16:B16').setFontWeight('bold');

  // Generate last 12 months
  for (let i = 0; i < 12; i++) {
    sheet.getRange(17 + i, 1).setFormula(
      `=EOMONTH(TODAY(),-${i})`
    );
    sheet.getRange(17 + i, 2).setFormula(
      `=SUMPRODUCT((MONTH('${grievanceSheetName}'!D:D)=MONTH(A${17 + i}))*` +
      `(YEAR('${grievanceSheetName}'!D:D)=YEAR(A${17 + i})))`
    );
  }
}

/**
 * Sets up the _CalcDeadlines hidden sheet
 * Contains deadline calculations and alert logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcDeadlinesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Deadline Calculations');
  sheet.getRange('A1').setFontWeight('bold');

  // Configuration reference
  sheet.getRange('A3').setValue('Deadline Rules (Days)');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setValue('Step 1 Response');
  sheet.getRange('B4').setValue(DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);

  sheet.getRange('A5').setValue('Step 2 Appeal');
  sheet.getRange('B5').setValue(DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL);

  sheet.getRange('A6').setValue('Step 2 Response');
  sheet.getRange('B6').setValue(DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE);

  sheet.getRange('A7').setValue('Step 3 Appeal');
  sheet.getRange('B7').setValue(DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL);

  sheet.getRange('A8').setValue('Step 3 Response');
  sheet.getRange('B8').setValue(DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE);

  sheet.getRange('A9').setValue('Arbitration Demand');
  sheet.getRange('B9').setValue(DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND);

  // Upcoming deadlines calculation
  sheet.getRange('D1').setValue('Upcoming Deadlines (Next 14 Days)');
  sheet.getRange('D1').setFontWeight('bold');

  sheet.getRange('D2').setValue('Grievance ID');
  sheet.getRange('E2').setValue('Step');
  sheet.getRange('F2').setValue('Due Date');
  sheet.getRange('G2').setValue('Days Left');
  sheet.getRange('D2:G2').setFontWeight('bold');

  // Complex formula to extract upcoming deadlines
  // This uses FILTER to get open grievances and calculate their current deadline
  sheet.getRange('D3').setFormula(
    `=IFERROR(FILTER('${grievanceSheetName}'!A:A,` +
    `('${grievanceSheetName}'!W:W="Open")+('${grievanceSheetName}'!W:W="Pending Response")+` +
    `('${grievanceSheetName}'!W:W="Appealed")),"")`
  );

  // Overdue grievances
  sheet.getRange('I1').setValue('Overdue Grievances');
  sheet.getRange('I1').setFontWeight('bold');
  sheet.getRange('I2').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"<>Resolved",` +
    `'${grievanceSheetName}'!W:W,"<>Closed",` +
    `'${grievanceSheetName}'!W:W,"<>Withdrawn",` +
    `'${grievanceSheetName}'!J:J,"<"&TODAY())`
  );

  // Alert thresholds
  sheet.getRange('I4').setValue('Alert Thresholds');
  sheet.getRange('I4').setFontWeight('bold');

  sheet.getRange('I5').setValue('Critical (<=3 days)');
  sheet.getRange('J5').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">="&TODAY(),` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+3)`
  );

  sheet.getRange('I6').setValue('Warning (4-7 days)');
  sheet.getRange('J6').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">"&TODAY()+3,` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+7)`
  );
}

/**
 * Sets up the _CalcStats hidden sheet
 * Contains dashboard-wide statistics
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcStatsSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Dashboard Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Quick stats for sidebar
  sheet.getRange('A3').setValue('Sidebar Stats');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setValue('open_grievances');
  sheet.getRange('B4').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Open")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Appealed")`
  );

  sheet.getRange('A5').setValue('pending_response');
  sheet.getRange('B5').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")`
  );

  sheet.getRange('A6').setValue('total_members');
  sheet.getRange('B6').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  sheet.getRange('A7').setValue('resolved_ytd');
  sheet.getRange('B7').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Resolved",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))+` +
    `COUNTIFS('${grievanceSheetName}'!W:W,"Closed",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))`
  );

  // Win rate calculation
  sheet.getRange('A9').setValue('Performance Metrics');
  sheet.getRange('A9').setFontWeight('bold');

  sheet.getRange('A10').setValue('total_resolved');
  sheet.getRange('B10').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Resolved")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Closed")`
  );

  sheet.getRange('A11').setValue('sustained_count');
  sheet.getRange('B11').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Sustained")`
  );

  sheet.getRange('A12').setValue('settled_count');
  sheet.getRange('B12').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Settled")`
  );

  sheet.getRange('A13').setValue('denied_count');
  sheet.getRange('B13').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Denied")`
  );

  sheet.getRange('A14').setValue('win_rate');
  sheet.getRange('B14').setFormula(
    `=IFERROR((B11+B12)/B10*100,0)`
  );

  // Average time to resolution
  sheet.getRange('A16').setValue('avg_days_to_resolve');
  sheet.getRange('B16').setFormula(
    `=IFERROR(AVERAGEIFS('${grievanceSheetName}'!X:X-'${grievanceSheetName}'!D:D,` +
    `'${grievanceSheetName}'!W:W,"Resolved"),0)`
  );
}

/**
 * Sets up the _CalcSync hidden sheet
 * Contains cross-sheet synchronization logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcSyncSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Cross-Sheet Sync');
  sheet.getRange('A1').setFontWeight('bold');

  // Member ID validation list
  sheet.getRange('A3').setValue('Valid Member IDs');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setFormula(
    `=FILTER('${memberSheetName}'!A:A,'${memberSheetName}'!A:A<>"ID",'${memberSheetName}'!A:A<>"")`
  );

  // Grievances per member count
  sheet.getRange('C3').setValue('Member ID');
  sheet.getRange('D3').setValue('Grievance Count');
  sheet.getRange('C3:D3').setFontWeight('bold');

  sheet.getRange('C4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!B:B,'${grievanceSheetName}'!B:B<>"Member ID",'${grievanceSheetName}'!B:B<>""))`
  );

  sheet.getRange('D4').setFormula(
    `=ARRAYFORMULA(IF(C4:C<>"",COUNTIF('${grievanceSheetName}'!B:B,C4:C),""))`
  );

  // Data consistency checks
  sheet.getRange('F3').setValue('Data Consistency');
  sheet.getRange('F3').setFontWeight('bold');

  sheet.getRange('F4').setValue('Orphaned Grievances');
  sheet.getRange('G4').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!B:B,"<>",'${grievanceSheetName}'!B:B,"<>Member ID")-` +
    `SUMPRODUCT(COUNTIF(A4:A,'${grievanceSheetName}'!B2:B))`
  );

  sheet.getRange('F5').setValue('Members with Grievances');
  sheet.getRange('G5').setFormula(
    `=COUNTA(C4:C)`
  );

  // Last sync timestamp
  sheet.getRange('F7').setValue('Last Formula Update');
  sheet.getRange('G7').setFormula('=NOW()');
}

/**
 * Sets up the _CalcFormulas hidden sheet
 * Contains named formula references for use in other sheets
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcFormulasSheet(sheet) {
  // Header
  sheet.getRange('A1').setValue('Named Formula References');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A2').setValue('Use these formulas via indirect references');

  // Department list formula
  sheet.getRange('A4').setValue('DEPARTMENT_LIST');
  sheet.getRange('B4').setFormula(
    `=SORT(UNIQUE(FILTER('${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E,` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"Department",` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"")))`
  );

  // Status list formula
  sheet.getRange('A6').setValue('STATUS_LIST');
  sheet.getRange('B6').setValue(Object.values(GRIEVANCE_STATUS).join(','));

  // Outcome list formula
  sheet.getRange('A8').setValue('OUTCOME_LIST');
  sheet.getRange('B8').setValue(Object.values(GRIEVANCE_OUTCOMES).join(','));

  // Grievance type list
  sheet.getRange('A10').setValue('GRIEVANCE_TYPES');
  sheet.getRange('B10').setValue('Contract Violation,Discipline,Discharge,Working Conditions,Safety,Other');

  // Date formatting formula
  sheet.getRange('A12').setValue('TODAY_FORMATTED');
  sheet.getRange('B12').setFormula('=TEXT(TODAY(),"MMMM D, YYYY")');

  // Year calculation
  sheet.getRange('A14').setValue('CURRENT_YEAR');
  sheet.getRange('B14').setFormula('=YEAR(TODAY())');

  // Next grievance ID prefix
  sheet.getRange('A16').setValue('GRIEVANCE_ID_PREFIX');
  sheet.getRange('B16').setFormula('="GRV-"&YEAR(TODAY())&"-"');
}
