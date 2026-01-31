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
    SHEETS.INTERACTIVE,
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
