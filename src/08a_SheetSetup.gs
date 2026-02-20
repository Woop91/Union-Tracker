/**
 * ============================================================================
 * 08a_SheetSetup.gs - Sheet Creation and Management
 * ============================================================================
 *
 * This module handles all sheet creation functions including:
 * - Main dashboard setup (CREATE_DASHBOARD)
 * - Individual sheet creation (Config, Member Directory, Grievance Log, etc.)
 * - Sheet ordering and organization
 * - Hidden sheet setup
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Sheet creation and management functions
 * @version 4.7.0
 * @requires 01_Core.gs
 */

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main setup function - creates the complete Dashboard
 * Creates the core sheets with proper structure and formatting
 * @returns {void}
 */
function CREATE_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('UI not available, proceeding without confirmation: ' + e.message);
  }

  if (ui) {
    var response = ui.alert(
      '🏗️ Create Dashboard',
      'This will create the Dashboard with:\n\n' +
      '• Config (dropdown sources)\n' +
      '• Member Directory\n' +
      '• Grievance Log (with Action Type dropdown)\n' +
      '• ✅ Case Checklist (track grievance tasks)\n' +
      '• 📊 Member Satisfaction (Survey tracking)\n' +
      '• 💡 Feedback & Development (Bug/feature tracking)\n' +
      '• 🤝 Volunteer Hours (track volunteer activities)\n' +
      '• 📅 Meeting Attendance (track meeting participation)\n' +
      '• 📝 Meeting Check-In Log (email+PIN check-in)\n' +
      '• ✅ Function Checklist (function reference)\n' +
      '• 📋 Features Reference (complete feature list)\n' +
      '• 📚 Getting Started (setup instructions)\n' +
      '• ❓ FAQ (common questions)\n' +
      '• 📖 Config Guide (how to use Config tab)\n\n' +
      'Note: All dashboards are now modal-based (popup windows).\n' +
      'Access them via: Union Hub > Dashboards menu.\n\n' +
      'Plus 6 hidden calculation sheets for self-healing formulas.\n\n' +
      'Existing sheets with data will be preserved (headers updated only).\n\n' +
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

    createFeaturesReferenceSheet(ss);
    ss.toast('Created Features Reference', '🏗️ Progress', 2);

    createVolunteerHoursSheet(ss);
    ss.toast('Created Volunteer Hours tracking', '🏗️ Progress', 2);

    createMeetingAttendanceSheet(ss);
    ss.toast('Created Meeting Attendance tracking', '🏗️ Progress', 2);

    createMeetingCheckInLogSheet(ss);
    ss.toast('Created Meeting Check-In Log', '🏗️ Progress', 2);

    saveFormUrlsToConfig_silent(ss);
    ss.toast('Saved form URLs to Config', '🏗️ Progress', 2);

    ss.toast('Setting up validations...', '🏗️ Progress', 3);
    setupDataValidations();

    reorderSheetsToStandard(ss);
    ss.toast('Sheets reordered', '🏗️ Progress', 2);

    // Install hourly trigger to keep hidden sheets hidden on mobile
    installHiddenSheetEnforcerTrigger();
    ss.toast('Hidden sheet enforcer installed', '🏗️ Progress', 2);

    ss.toast('Dashboard creation complete!', '✅ Success', 5);
    if (ui) {
      ui.alert('✅ Success', 'Dashboard has been created successfully!\n\n' +
        '14 sheets created:\n' +
        '• Config, Member Directory, Grievance Log (data)\n' +
        '• ✅ Case Checklist (track grievance tasks)\n' +
        '• 📊 Member Satisfaction, 💡 Feedback (tracking)\n' +
        '• 🤝 Volunteer Hours, 📅 Meeting Attendance, 📝 Meeting Check-In Log\n' +
        '• ✅ Function Checklist, 📋 Features Reference (references)\n' +
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
    Logger.log('Error in CREATE_DASHBOARD: ' + error.message);
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
    // CRITICAL: Never clear sheets that contain user data.
    // Only clear if the sheet is empty (no data rows beyond header).
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Sheet has data - preserve it. Only update headers if needed.
      Logger.log('Sheet "' + name + '" has ' + (lastRow - 1) + ' data rows - preserving existing data');
      return sheet;
    }
    // Sheet is empty or has only headers - safe to clear and rebuild
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
  // Dashboard and Satisfaction sheets are deprecated (v4.3.2/v4.3.8) - excluded from ordering
  var desiredOrder = [
    SHEETS.GETTING_STARTED,
    SHEETS.FAQ,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.FEEDBACK,
    SHEETS.FUNCTION_CHECKLIST,
    SHEETS.CONFIG_GUIDE,
    'Config'
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
 * Sets up hidden calculation sheets.
 * Creates (or recreates) each hidden sheet, runs its setup function, and hides it.
 *
 * Includes the _Survey_Tracking sheet which tracks per-member survey completion.
 * Survey tracking detection flow:
 *   Google Form submit -> onSatisfactionFormSubmit() -> validateMemberEmail()
 *   -> updateSurveyTrackingOnSubmit_() in 08c_FormsAndNotifications.gs
 * See SURVEY_TRACKING_COLS in 01_Core.gs for full documentation.
 *
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
    { name: HIDDEN_SHEETS.CALC_FORMULAS, setup: setupCalcFormulasSheet },
    { name: HIDDEN_SHEETS.SURVEY_TRACKING, setup: setupSurveyTrackingSheet }
  ];

  hiddenSheets.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
    }
    sheet.clear();
    config.setup(sheet);
    setSheetVeryHidden_(sheet);
  });

  // Self-contained hidden sheet setups (each creates/hides its own sheet)
  setupGrievanceCalcSheet();
  setupMemberLookupSheet();
  setupStewardContactCalcSheet();
  setupStewardPerformanceCalcSheet();
  setupAuditLogSheet();
  setupChecklistCalcSheet();

  // Survey Vault uses its own setup (includes sheet protection)
  setupSurveyVaultSheet();
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
  // Re-sync column maps from actual sheet headers before applying validations.
  // This guarantees dropdowns land on the correct columns even if the layout changed.
  try { syncColumnMaps(); } catch (_e) { /* proceed with defaults */ }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet || !memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found. Please run CREATE_DASHBOARD first.');
    return;
  }

  // Member Directory Validations — driven by DROPDOWN_MAP (single-select)
  var memberDD = DROPDOWN_MAP.MEMBER_DIR;
  for (var m = 0; m < memberDD.length; m++) {
    setDropdownValidation(memberSheet, memberDD[m].col, configSheet, memberDD[m].configCol);
  }

  // Member Directory Validations — driven by MULTI_SELECT_COLS
  var memberMS = MULTI_SELECT_COLS.MEMBER_DIR;
  for (var mm = 0; mm < memberMS.length; mm++) {
    setMultiSelectValidation(memberSheet, memberMS[mm].col, configSheet, memberMS[mm].configCol);
  }

  // Grievance Log Validations — driven by DROPDOWN_MAP (single-select)
  var grievDD = DROPDOWN_MAP.GRIEVANCE_LOG;
  for (var g = 0; g < grievDD.length; g++) {
    setDropdownValidation(grievanceSheet, grievDD[g].col, configSheet, grievDD[g].configCol);
  }

  // Grievance Log Validations — driven by MULTI_SELECT_COLS
  var grievMS = MULTI_SELECT_COLS.GRIEVANCE_LOG;
  for (var gm = 0; gm < grievMS.length; gm++) {
    setMultiSelectValidation(grievanceSheet, grievMS[gm].col, configSheet, grievMS[gm].configCol);
  }

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

  var targetRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, Math.max(1, grievanceSheet.getMaxRows() - 1), 1);
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
  // Get actual values from Config column (skip blanks) and use requireValueInList
  // This avoids the 500-row fixed range issue that can cause empty entries in dropdowns
  var lastRow = configSheet.getLastRow();
  var values = [];
  if (lastRow >= 3) {
    var data = configSheet.getRange(3, sourceCol, lastRow - 2, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        values.push(data[i][0].toString().trim());
      }
    }
  }

  if (values.length === 0) return; // No values to validate against

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)  // Allow custom entries for bidirectional sync with Config
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, Math.max(1, targetSheet.getMaxRows() - 1), 1);
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
  // Get actual values from Config column (skip blanks) and use requireValueInList
  var lastRow = configSheet.getLastRow();
  var values = [];
  if (lastRow >= 3) {
    var data = configSheet.getRange(3, sourceCol, lastRow - 2, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        values.push(data[i][0].toString().trim());
      }
    }
  }

  if (values.length === 0) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)  // Allow comma-separated values from multi-select dialog
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, Math.max(1, targetSheet.getMaxRows() - 1), 1);
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
  // The inline dialog (getMultiSelectHtml) passes an array of selected IDs,
  // while MultiSelectDialog.html passes a pre-joined comma string.
  // Normalise to a comma-separated string so all selections are saved.
  if (Array.isArray(value)) {
    value = value.join(', ');
  }

  var props = PropertiesService.getUserProperties();
  var row = parseInt(props.getProperty('multiSelectRow'), 10);
  var col = parseInt(props.getProperty('multiSelectCol'), 10);
  var sheetName = props.getProperty('multiSelectSheet') || SHEETS.MEMBER_DIR;

  if (!row || !col) {
    throw new Error('Target cell not found. Please try again.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  sheet.getRange(row, col).setValue(value);

  props.deleteProperty('multiSelectRow');
  props.deleteProperty('multiSelectCol');
  props.deleteProperty('multiSelectSheet');
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

  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (row < 2) return;

  var config = getMultiSelectConfig(col, sheetName);
  if (!config) return;

  var newValue = e.value || '';

  if (newValue === '' || newValue.indexOf(',') !== -1) return;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tip: Use Tools menu > "Multi-Select Editor" for easier selection of multiple values.',
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

  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  var config = getMultiSelectConfig(col, sheetName);
  if (!config) return;

  var props = PropertiesService.getUserProperties();
  var lastCell = props.getProperty('lastMultiSelectCell');
  var currentCell = row + ',' + col;

  if (lastCell === currentCell) return;

  props.setProperty('lastMultiSelectCell', currentCell);
  openCellMultiSelectEditor();
}

/**
 * Enables the multi-select auto-open feature.
 * Uses a user property flag checked by the simple onSelectionChange trigger.
 * (onSelectionChange cannot be installed via ScriptApp.newTrigger; it only
 * works as a simple trigger defined in code.)
 * @returns {void}
 */
function installMultiSelectTrigger() {
  var props = PropertiesService.getUserProperties();
  if (props.getProperty('multiSelectAutoOpen') === 'true') {
    SpreadsheetApp.getUi().alert('Multi-Select auto-open is already enabled.');
    return;
  }

  props.setProperty('multiSelectAutoOpen', 'true');

  // Clean up any legacy .onChange() triggers that were incorrectly installed
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSelectionChangeMultiSelect') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  SpreadsheetApp.getUi().alert('Multi-Select auto-open has been enabled!\n\n' +
    'The multi-select dialog will now open automatically when you click a multi-select cell.');
}

/**
 * Disables the multi-select auto-open feature.
 * Clears the user property flag and removes any legacy installable triggers.
 * @returns {void}
 */
function removeMultiSelectTrigger() {
  var props = PropertiesService.getUserProperties();
  var wasEnabled = props.getProperty('multiSelectAutoOpen') === 'true';
  props.deleteProperty('multiSelectAutoOpen');
  props.deleteProperty('lastMultiSelectCell');

  // Also clean up any legacy .onChange() triggers
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSelectionChangeMultiSelect') {
      ScriptApp.deleteTrigger(trigger);
      wasEnabled = true;
    }
  });

  if (wasEnabled) {
    SpreadsheetApp.getUi().alert('Multi-Select auto-open has been disabled.');
  } else {
    SpreadsheetApp.getUi().alert('Multi-Select auto-open was not enabled.');
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
  // Delegate to the main function which now uses dynamic value lists
  setDropdownValidation(targetSheet, targetCol, configSheet, sourceCol);
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
  // Delegate to the main function which now uses dynamic value lists
  setMultiSelectValidation(targetSheet, targetCol, configSheet, sourceCol);
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

