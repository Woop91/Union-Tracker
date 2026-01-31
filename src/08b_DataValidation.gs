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
