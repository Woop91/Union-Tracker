/**
 * ============================================================================
 * 11_CommandHub.gs - STRATEGIC COMMAND CENTER
 * ============================================================================
 * STATUS: Production Ready / Harmonized / High-Performance
 *
 * v4.0 UNIFIED ARCHITECTURE:
 * - Single-File Modular Build (Virtual Files: Constants, UI, Performance, Security, DevTools)
 * - Global Scope Rule: All functions share CONFIG object
 *
 * FEATURES:
 * - Security: Audit Log & Sabotage Protection (>15 cells)
 * - Performance: Batch Array Processing (No-Lag Architecture up to 5,000 members)
 * - Workflow: Stage-Gate Case Tracking & Auto-PDF Generation
 * - Production: Nuke/Seed Isolation & UI Self-Hiding (PRODUCTION_MODE)
 * - Accessibility: Mobile/Pocket View & Search Engine
 * - Legal: Signature-Ready PDF Merge & Auto-Drive Archiving
 *
 * CURRENT PHASE: Pre-Production (verifying ID generation & Mobile View logic)
 *
 * NOTE: This file provides consolidated access to Strategic Command Center
 * features. Core implementations are in their respective module files.
 *
 * @version 4.7.0
 * ============================================================================
 */

// ACCEPTABLE: 60+ commands reflect feature depth; collision managed by verb-noun naming convention
// ============================================================================
// COMMAND CENTER CONFIGURATION
// ============================================================================
/**
 * Get COMMAND_CENTER_CONFIG lazily to avoid load-order issues.
 * All properties are resolved on demand — no eager initialization.
 * @returns {Object} Command center configuration object
 */
function getCommandCenterConfig() {
  return {
    SYSTEM_NAME: "Strategic Command Center",
    LOG_SHEET_NAME: SHEETS.GRIEVANCE_LOG,
    DIR_SHEET_NAME: SHEETS.MEMBER_DIR,
    AUDIT_SHEET_NAME: SHEETS.AUDIT_LOG,
    TEMPLATE_ID: COMMAND_CONFIG.TEMPLATE_ID,
    ARCHIVE_FOLDER_ID: COMMAND_CONFIG.ARCHIVE_FOLDER_ID,
    CHIEF_STEWARD_EMAIL: COMMAND_CONFIG.CHIEF_STEWARD_EMAIL,
    UNIT_CODES: COMMAND_CONFIG.UNIT_CODES,
    THEME: COMMAND_CONFIG.THEME
  };
}

// Legacy COMMAND_CENTER_CONFIG — uses lazy getters so help content and
// dynamic fields are only resolved when first accessed (not at load time).
// @deprecated Use getCommandCenterConfig() or COMMAND_CONFIG instead. Kept for backward compatibility.
var COMMAND_CENTER_CONFIG = {
  SYSTEM_NAME: "Strategic Command Center",
  get LOG_SHEET_NAME()    { return typeof SHEETS !== 'undefined' && SHEETS.GRIEVANCE_LOG ? SHEETS.GRIEVANCE_LOG : 'Grievance Log'; },
  get DIR_SHEET_NAME()    { return typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR ? SHEETS.MEMBER_DIR : 'Member Directory'; },
  get AUDIT_SHEET_NAME()  { return typeof SHEETS !== 'undefined' && SHEETS.AUDIT_LOG ? SHEETS.AUDIT_LOG : '_Audit_Log'; },
  get TEMPLATE_ID()       { return COMMAND_CONFIG.TEMPLATE_ID; },
  get ARCHIVE_FOLDER_ID() { return COMMAND_CONFIG.ARCHIVE_FOLDER_ID; },
  get CHIEF_STEWARD_EMAIL() { return COMMAND_CONFIG.CHIEF_STEWARD_EMAIL; },
  get UNIT_CODES()        { return COMMAND_CONFIG.UNIT_CODES; },
  THEME: {
    HEADER_BG: '#1e293b',
    HEADER_TEXT: '#ffffff',
    ALT_ROW: '#f8fafc',
    FONT: 'Roboto',
    FONT_SIZE: 10
  }
};

// ============================================================================
// COMMAND CENTER MENU (DEPRECATED - v4.4.0)
// ============================================================================

/**
 * @deprecated v4.4.0 - Menu consolidated into createDashboardMenu() in UIService.gs
 * The Field Portal menu has been merged into:
 * - 📊 Union Hub (primary operations)
 * - 🔧 Tools (supporting features)
 * - 🛠️ Admin (system administration)
 *
 * This function is kept for backward compatibility but is no longer called.
 * All menu items have been reorganized into the new consolidated structure.
 */
function createCommandCenterMenu() {
  // v4.4.0: This function is deprecated - menu consolidated into UIService.gs
  // Keeping the function stub to prevent errors if called from legacy code
  console.log('createCommandCenterMenu() is deprecated in v4.4.0 - menus consolidated');
}

// ============================================================================
// NAVIGATION SHORTCUTS (v4.0)
// ============================================================================

/**
 * Quick navigation to Executive Dashboard
 */
function navigateToDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found. Use the Dashboards menu for modal dashboards.');
  }
}

/**
 * Quick navigation to Mobile View
 */
function navigateToMobileView() {
  navToMobile();
}

/**
 * v4.6 Pocket View Navigation
 * Detects current sheet (Member Directory or Grievance Log) and applies
 * appropriate pocket view by hiding non-essential columns.
 *
 * Member Directory pocket view HIDES: PIN Hash, Employee ID, Hire Date, Department
 * Member Directory pocket view SHOWS: First Name, Last Name, Email, Phone,
 *   Recent Contact Date, Contact Steward, Contact Notes, and more
 *
 * Grievance Log pocket view HIDES: Grievance ID, Member ID, Articles Violated,
 *   Issue Category, Work Location, Resolution, Checklist Progress,
 *   Reminder 1 Date, Reminder 1 Note, Reminder 2 Date, Reminder 2 Note, Last Updated
 *
 * Call showAllColumns() to restore full view.
 */
function navToMobile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();

  // Determine which sheet to apply pocket view to
  if (activeSheetName === SHEETS.GRIEVANCE_LOG) {
    applyGrievancePocketView_();
  } else {
    applyMemberPocketView_();
  }
}

/**
 * Applies pocket view to Member Directory
 * Hides: PIN Hash, Employee ID, Hire Date, Department
 * Shows: Name, Email, Phone, Recent Contact Date, Contact Steward, Contact Notes
 * @private
 */
function applyMemberPocketView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  try {
    // First show all columns, then hide specific ones
    sheet.showColumns(1, sheet.getMaxColumns());

    // Columns to HIDE in pocket view:
    // PIN Hash, Employee ID, Hire Date, Department
    var colsToHide = [
      MEMBER_COLS.PIN_HASH,       // AG - PIN Hash
      MEMBER_COLS.EMPLOYEE_ID,    // AH - Employee ID
      MEMBER_COLS.HIRE_DATE,      // AJ - Hire Date
      MEMBER_COLS.DEPARTMENT      // AI - Department
    ];

    // Also hide: Cubicle (already hidden), Engagement metrics, Interests, PII
    colsToHide.push(MEMBER_COLS.CUBICLE);           // Cubicle
    colsToHide.push(MEMBER_COLS.LAST_VIRTUAL_MTG);  // Last Virtual Mtg
    colsToHide.push(MEMBER_COLS.LAST_INPERSON_MTG); // Last In-Person Mtg
    colsToHide.push(MEMBER_COLS.OPEN_RATE);          // Open Rate
    colsToHide.push(MEMBER_COLS.VOLUNTEER_HOURS);    // Volunteer Hours
    colsToHide.push(MEMBER_COLS.INTEREST_LOCAL);     // Interest: Local
    colsToHide.push(MEMBER_COLS.INTEREST_CHAPTER);   // Interest: Chapter
    colsToHide.push(MEMBER_COLS.INTEREST_ALLIED);    // Interest: Allied
    colsToHide.push(MEMBER_COLS.STREET_ADDRESS);     // Street Address (PII)
    colsToHide.push(MEMBER_COLS.CITY);               // City (PII)
    colsToHide.push(MEMBER_COLS.STATE);              // State (PII)

    // Hide each column individually
    for (var i = 0; i < colsToHide.length; i++) {
      if (colsToHide[i] <= sheet.getMaxColumns()) {
        sheet.hideColumns(colsToHide[i]);
      }
    }

    sheet.activate();
    ss.toast('Pocket View: Member Directory. Key columns visible for quick access.', COMMAND_CONFIG.SYSTEM_NAME, 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error enabling pocket view: ' + e.message);
  }
}

/**
 * Applies pocket view to Grievance Log
 * Hides: Grievance ID, Member ID, Articles Violated, Issue Category,
 *   Work Location, Resolution, Checklist Progress,
 *   Reminder 1 Date, Reminder 1 Note, Reminder 2 Date, Reminder 2 Note, Last Updated
 * @private
 */
function applyGrievancePocketView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log not found');
    return;
  }

  try {
    // First show all columns, then hide specific ones
    sheet.showColumns(1, sheet.getMaxColumns());

    // Columns to HIDE in grievance pocket view:
    var colsToHide = [
      GRIEVANCE_COLS.GRIEVANCE_ID,       // Grievance ID
      GRIEVANCE_COLS.MEMBER_ID,          // Member ID
      GRIEVANCE_COLS.ARTICLES,           // Articles Violated
      GRIEVANCE_COLS.ISSUE_CATEGORY,     // Issue Category
      GRIEVANCE_COLS.LOCATION,           // Work Location
      GRIEVANCE_COLS.RESOLUTION,         // Resolution
      GRIEVANCE_COLS.CHECKLIST_PROGRESS, // Checklist Progress
      GRIEVANCE_COLS.REMINDER_1_DATE,    // Reminder 1 Date
      GRIEVANCE_COLS.REMINDER_1_NOTE,    // Reminder 1 Note
      GRIEVANCE_COLS.REMINDER_2_DATE,    // Reminder 2 Date
      GRIEVANCE_COLS.REMINDER_2_NOTE,    // Reminder 2 Note
      GRIEVANCE_COLS.LAST_UPDATED        // Last Updated
    ];

    for (var i = 0; i < colsToHide.length; i++) {
      if (colsToHide[i] <= sheet.getMaxColumns()) {
        sheet.hideColumns(colsToHide[i]);
      }
    }

    sheet.activate();
    ss.toast('Pocket View: Grievance Log. Essential columns visible.', COMMAND_CONFIG.SYSTEM_NAME, 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error enabling pocket view: ' + e.message);
  }
}

/**
 * Restore Full View
 * Shows all columns on the active sheet (Member Directory or Grievance Log).
 */
function showAllMemberColumns() {
  showAllColumns();
}

/**
 * Restore all columns for the active sheet
 */
function showAllColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();

  // Restore on the active sheet
  var sheet = activeSheet;
  if (activeSheetName !== SHEETS.MEMBER_DIR && activeSheetName !== SHEETS.GRIEVANCE_LOG) {
    // Default to Member Directory
    sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  }

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet not found');
    return;
  }

  try {
    sheet.showColumns(1, sheet.getMaxColumns());

    // Re-hide Cubicle column in Member Directory (always hidden)
    if (sheet.getName() === SHEETS.MEMBER_DIR) {
      sheet.hideColumns(MEMBER_COLS.CUBICLE);
    }

    ss.toast('All columns restored.', COMMAND_CONFIG.SYSTEM_NAME, 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error showing columns: ' + e.message);
  }
}

/**
 * Toggle Mobile View - Easy one-click toggle between mobile and full view
 * Checks current state and switches to the opposite view
 * Accessible from top-level Union Hub menu for mobile users
 */
function toggleMobileView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  // Check if we're currently in mobile view by checking if column D is hidden
  // Column D (Job Title) is always hidden in mobile view
  try {
    // Use properties to track mobile view state
    var props = PropertiesService.getUserProperties();
    var isMobileView = props.getProperty('MOBILE_VIEW_ENABLED') === 'true';

    if (isMobileView) {
      // Switch to full view
      showAllMemberColumns();
      props.setProperty('MOBILE_VIEW_ENABLED', 'false');
    } else {
      // Switch to mobile view
      navToMobile();
      props.setProperty('MOBILE_VIEW_ENABLED', 'true');
    }
  } catch (_e) {
    // Fallback: just try to show all columns (safe default)
    showAllMemberColumns();
    SpreadsheetApp.getUi().alert('Restored to full view. Use this toggle to switch between mobile and desktop views.');
  }
}

// ============================================================================
// v4.0 HIGH-PERFORMANCE DATA ENGINE
// ============================================================================


/**
 * Batch Member ID Generator (High-Performance)
 * Uses batch array processing to generate IDs for up to 5,000 members without lag.
 * Format: M + first 2 letters of first name + first 2 letters of last name + 3 random digits
 */
function generateMissingMemberIDsBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || sheet.getLastRow() < 2) {
    ss.toast('No members found to process.', COMMAND_CONFIG.SYSTEM_NAME, 3);
    return;
  }

  // Batch read all data
  var data = sheet.getDataRange().getValues();
  var countAdded = 0;

  // Collect existing IDs for collision detection
  var existingIds = {};
  for (var j = 1; j < data.length; j++) {
    var id = data[j][MEMBER_COLS.MEMBER_ID - 1];
    if (id) existingIds[id] = true;
  }

  // Process in memory, tracking which rows were modified
  var changedRows = []; // { row: sheetRow (1-indexed), id: newId }
  for (var i = 1; i < data.length; i++) {
    var firstName = data[i][MEMBER_COLS.FIRST_NAME - 1];
    var lastName = data[i][MEMBER_COLS.LAST_NAME - 1];

    // Check if Member ID is empty and member has name data
    if (!data[i][MEMBER_COLS.MEMBER_ID - 1] && (firstName || lastName)) {
      var newId = generateNameBasedId('M', firstName, lastName, existingIds);
      existingIds[newId] = true;
      countAdded++;
      changedRows.push({ row: i + 1, id: newId });
    }
  }

  // Batch write only the Member ID column cells that were actually changed
  if (countAdded > 0) {
    for (var c = 0; c < changedRows.length; c++) {
      sheet.getRange(changedRows[c].row, MEMBER_COLS.MEMBER_ID).setValue(changedRows[c].id);
    }
    SpreadsheetApp.flush();
  }

  ss.toast(countAdded + ' IDs generated.', COMMAND_CONFIG.SYSTEM_NAME, 3);

  return {
    generated: countAdded,
    total: data.length - 1
  };
}

/**
 * v4.0 Refresh Member View
 * Reloads the Member Directory view without changing column visibility.
 */
function refreshMemberView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  sheet.activate();
  SpreadsheetApp.flush();
  ss.toast('✅ View refreshed.', COMMAND_CONFIG.SYSTEM_NAME, 2);
}

/**
 * ID Generation Engine Verification
 * Tests the name-based ID generation system and reports Member Directory statistics.
 */
function verifyIDGenerationEngine() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var report = '🔍 ID GENERATION ENGINE VERIFICATION\n';
  report += '=' .repeat(45) + '\n\n';

  // Test name-based ID generation
  report += '🧪 TEST ID GENERATION (Name-Based):\n';
  report += '  Format: M + 2 chars first name + 2 chars last name + 3 digits\n';
  var testId = generateNameBasedId('M', 'Jane', 'Smith', {});
  report += '  Test (Jane Smith): ' + testId + '\n';
  var formatValid = /^M[A-Z]{4}\d{3}$/.test(testId);
  report += '  Format Valid: ' + (formatValid ? '✅ Yes' : '❌ No') + '\n';

  // Check Member Directory for ID statistics
  report += '\n📈 MEMBER ID STATISTICS:\n';
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
    var withID = 0;
    var withoutID = 0;
    var nameBasedCount = 0;
    var legacyCount = 0;

    data.forEach(function(row) {
      if (row[0]) {
        withID++;
        var idStr = String(row[0]);
        if (/^M[A-Z]{4}\d{3}/.test(idStr)) {
          nameBasedCount++;
        } else {
          legacyCount++;
        }
      } else {
        withoutID++;
      }
    });

    report += '  Members with ID: ' + withID + '\n';
    report += '  Members without ID: ' + withoutID + '\n';
    report += '  Name-based IDs (current format): ' + nameBasedCount + '\n';
    if (legacyCount > 0) {
      report += '  Legacy IDs (old format): ' + legacyCount + '\n';
    }
  } else {
    report += '  No member data found.\n';
  }

  report += '\n' + '=' .repeat(45) + '\n';
  report += '✅ ID Engine Verification Complete\n';

  ui.alert('ID Engine Report', report, ui.ButtonSet.OK);

  return {
    success: true,
    report: report
  };
}

/**
 * v4.0 Status Report
 * Displays comprehensive system status for the Unified Master Engine.
 */
function showV4StatusReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var report = '📊 STRATEGIC COMMAND CENTER\n';
  report += 'v4.0 UNIFIED MASTER ENGINE STATUS\n';
  report += '=' .repeat(45) + '\n\n';

  // System Identity
  report += '🏷️ SYSTEM IDENTITY:\n';
  report += '  Name: ' + COMMAND_CONFIG.SYSTEM_NAME + '\n';
  report += '  Version: ' + COMMAND_CONFIG.VERSION + ' (' + VERSION_INFO.BUILD_DATE + ')\n';
  report += '  Codename: ' + VERSION_INFO.CODENAME + '\n';
  report += '  Architecture: Single-File Modular (10 Virtual Files)\n\n';

  // Production Status
  report += '🔒 PRODUCTION STATUS:\n';
  var prodMode = isProductionMode();
  report += '  Mode: ' + (prodMode ? '🔴 PRODUCTION' : '🟢 DEVELOPMENT') + '\n';
  report += '  Demo Menu: ' + (prodMode ? 'Hidden' : 'Visible') + '\n\n';

  // Feature Status
  report += '⚡ v4.0 FEATURES:\n';
  report += '  ✅ Security Fortress (Audit Log + Sabotage Alert)\n';
  report += '  ✅ High-Performance Engine (Batch Array Processing)\n';
  report += '  ✅ Mobile/Pocket View (Field Accessibility)\n';
  report += '  ✅ Stage-Gate Workflow (Escalation Alerts)\n';
  report += '  ✅ Production Mode (UI Self-Hiding)\n';
  report += '  ✅ Search Engine (Member Lookup)\n\n';

  // Data Summary
  report += '📈 DATA SUMMARY:\n';
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (memberSheet) {
    report += '  Members: ' + Math.max(0, memberSheet.getLastRow() - 1) + '\n';
  }
  if (grievanceSheet) {
    report += '  Grievances: ' + Math.max(0, grievanceSheet.getLastRow() - 1) + '\n';
  }

  // Configuration Check
  report += '\n⚙️ CONFIGURATION:\n';
  try {
    var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
    report += '  Chief Steward Email: ' + (chiefEmail ? '✅ Set' : '⚠️ Not configured') + '\n';
  } catch (_e) {
    report += '  Chief Steward Email: ⚠️ Unable to check\n';
  }

  report += '\n' + '=' .repeat(45) + '\n';
  report += 'Status: ✅ All Systems Operational\n';

  ui.alert('v4.0 Status Report', report, ui.ButtonSet.OK);

  return {
    success: true,
    productionMode: prodMode,
    version: COMMAND_CONFIG.VERSION
  };
}

// ============================================================================
// PRODUCTION MODE & NUKE FUNCTIONS
// ============================================================================

/**
 * Checks if system is in production mode
 * @returns {boolean} True if production mode is enabled
 */
function isProductionMode() {
  return PropertiesService.getScriptProperties().getProperty('PRODUCTION_MODE') === 'true';
}

/**
 * Enables production mode (hides demo tools)
 */
function enableProductionMode() {
  PropertiesService.getScriptProperties().setProperty('PRODUCTION_MODE', 'true');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Production Mode enabled. Reload the spreadsheet to see changes.',
    COMMAND_CONFIG.SYSTEM_NAME,
    5
  );
}

/**
 * Disables production mode (shows demo tools)
 */
function disableProductionMode() {
  PropertiesService.getScriptProperties().setProperty('PRODUCTION_MODE', 'false');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Production Mode disabled. Demo tools will be visible on reload.',
    COMMAND_CONFIG.SYSTEM_NAME,
    5
  );
}

/**
 * THE NUCLEAR OPTION - Wipes all data and enables production mode
 * More aggressive than NUKE_SEEDED_DATA - clears ALL data, not just seeded
 */
function NUKE_DATABASE() {
  var isDev = typeof IS_DEV_ENVIRONMENT !== 'undefined' && IS_DEV_ENVIRONMENT === true;
  if (!isDev) {
    var _ui = SpreadsheetApp.getUi();
    var _response = _ui.alert('Production Safety Check',
      'This action is intended for development environments only. Are you SURE you want to proceed?',
      _ui.ButtonSet.YES_NO);
    if (_response !== _ui.Button.YES) return;
  }
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '☢️ NUCLEAR OPTION',
    'This will:\n\n' +
    '• DELETE ALL members from Member Directory\n' +
    '• DELETE ALL grievances from Grievance Log\n' +
    '• CLEAR Config dropdown values\n' +
    '• DELETE Function Checklist sheet\n' +
    '• DELETE _Audit_Log hidden sheet\n' +
    '• REMOVE demo references from documentation tabs\n' +
    '• APPLY professional tab colors\n' +
    '• ENABLE Production Mode (hide Demo menu)\n\n' +
    '⏱️ IMPORTANT: This process takes approximately 3-5 MINUTES.\n' +
    '⚠️ WAIT until the "Running script" dialog disappears!\n\n' +
    'This action CANNOT be undone!\n\n' +
    'Are you absolutely sure?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert('Cancelled', 'Nuclear operation cancelled.', ui.ButtonSet.OK);
    return;
  }

  // Second confirmation
  var confirm2 = ui.alert(
    '⚠️ FINAL WARNING',
    'You are about to permanently delete ALL data.\n\n' +
    '⏱️ This will take 3-5 MINUTES. DO NOT close the tab!\n' +
    'Wait for the "Running script" dialog to disappear.\n\n' +
    'Type YES to confirm this is intentional.',
    ui.ButtonSet.YES_NO
  );

  if (confirm2 !== ui.Button.YES) {
    ui.alert('Cancelled', 'Nuclear operation cancelled.', ui.ButtonSet.OK);
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Nuking database... This will take 3-5 minutes. Please wait!', '☢️ NUKE', 10);

  try {
    // Clear Member Directory (preserve header row)
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (memberSheet && memberSheet.getLastRow() > 1) {
      memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn())
        .clearContent()
        .setBackground(null)
        .clearNote();
    }

    // Clear Grievance Log (preserve header row)
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, grievanceSheet.getLastColumn())
        .clearContent()
        .setBackground(null)
        .clearNote();
    }

    // Clear Config dropdown values (rows 3+, columns A-E typical dropdowns)
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && configSheet.getLastRow() > 2) {
      configSheet.getRange(3, 1, Math.max(1, configSheet.getLastRow() - 2), 5)
        .clearContent();
    }

    // Delete Function Checklist sheet
    var functionChecklistSheet = ss.getSheetByName(SHEETS.FUNCTION_CHECKLIST);
    if (functionChecklistSheet) {
      try { ss.deleteSheet(functionChecklistSheet); } catch (e) { Logger.log('Could not delete Function Checklist: ' + e.message); }
    }

    // Log the nuclear wipe to both audit sheet AND ScriptProperties (durable record)
    var wipeRecord = {
      event: 'NUCLEAR_WIPE',
      performedBy: (function() { try { return Session.getActiveUser().getEmail(); } catch (_e) { return 'Unknown'; } })(),
      timestamp: new Date().toISOString()
    };
    logAuditEvent('NUCLEAR_WIPE', wipeRecord);
    // Persist to ScriptProperties so the record survives audit sheet deletion
    PropertiesService.getScriptProperties().setProperty('LAST_NUCLEAR_WIPE', JSON.stringify(wipeRecord));

    // Archive _Audit_Log before deletion for compliance
    var auditLogSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    if (auditLogSheet) {
      try {
        var auditData = auditLogSheet.getDataRange().getValues();
        if (auditData.length > 1) {
          var csv = auditData.map(function(row) {
            return row.map(function(val) {
              var s = String(val);
              return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1) ? '"' + s.replace(/"/g, '""') + '"' : s;
            }).join(',');
          }).join('\n');
          DriveApp.createFile(Utilities.newBlob(csv, 'text/csv', 'AUDIT_LOG_BACKUP_' + new Date().getTime() + '.csv'));
        }
        ss.deleteSheet(auditLogSheet);
      } catch (e) { Logger.log('Could not archive/delete _Audit_Log: ' + e.message); }
    }

    // Clean up demo references from documentation tabs
    cleanupDocumentationTabs_(ss);

    // Apply professional tab colors
    applyTabColors_(ss);

    // Enable Production Mode
    PropertiesService.getScriptProperties().setProperty('PRODUCTION_MODE', 'true');

    // Clear demo tracking
    PropertiesService.getScriptProperties().deleteProperty('SEEDED_MEMBER_IDS');
    PropertiesService.getScriptProperties().deleteProperty('SEEDED_GRIEVANCE_IDS');

    ui.alert(
      '✅ Nuclear Operation Complete',
      'All data has been wiped.\n\n' +
      'Production Mode has been enabled.\n' +
      'The Demo Data menu will be hidden on reload.\n' +
      'Tab colors have been applied.\n\n' +
      'Please reload the spreadsheet to see changes.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error', 'Nuclear operation failed: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// NUKE HELPER FUNCTIONS - Documentation & Tab Colors
// ============================================================================

/**
 * Cleans up demo/seed/nuke references from documentation tabs
 * Called during NUKE to remove developer-focused content
 * Also adds a repo link to FAQ for users wanting seed/nuke features
 * @param {Spreadsheet} ss - The active spreadsheet
 * @private
 */
function cleanupDocumentationTabs_(ss) {
  var docSheets = [
    SHEETS.FAQ,
    SHEETS.CONFIG_GUIDE,
    SHEETS.GETTING_STARTED
  ];

  // Patterns to identify rows mentioning seed/nuke/demo features
  var demoPatterns = [
    /\bseed\b/i,
    /\bnuke\b/i,
    /\bdemo data\b/i,
    /\bsample data\b/i,
    /\btest data\b/i,
    /SEED_/i,
    /NUKE_/i,
    /🎭.*demo/i
  ];

  docSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    try {
      var data = sheet.getDataRange().getValues();
      var rowsToUpdate = [];

      // Find rows that contain demo-related content
      for (var i = 0; i < data.length; i++) {
        var rowText = data[i].join(' ');
        var shouldClean = demoPatterns.some(function(pattern) {
          return pattern.test(rowText);
        });

        if (shouldClean) {
          // Clear this row's content (but preserve row structure)
          rowsToUpdate.push(i + 1); // 1-indexed
        }
      }

      // Clear identified rows
      rowsToUpdate.forEach(function(row) {
        try {
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
        } catch (e) {
          Logger.log('Could not clear row ' + row + ' in ' + sheetName + ': ' + e.message);
        }
      });

      Logger.log('Cleaned ' + rowsToUpdate.length + ' demo-related rows from ' + sheetName);
    } catch (e) {
      Logger.log('Error cleaning ' + sheetName + ': ' + e.message);
    }
  });

  // Add repo link to FAQ for users who want to start fresh with seed/nuke
  addRepoLinkToFAQ_(ss);
}

/**
 * Adds a link to the GitHub repo in the FAQ sheet
 * For users who want to create their own copy with seed/nuke features
 * @param {Spreadsheet} ss - The active spreadsheet
 * @private
 */
function addRepoLinkToFAQ_(ss) {
  var faqSheet = ss.getSheetByName(SHEETS.FAQ);
  if (!faqSheet) return;

  try {
    // Find the last row with content
    var lastRow = faqSheet.getLastRow();

    // Add some spacing and the repo link section
    var linkRow = lastRow + 2;

    // Add header for the section
    faqSheet.getRange(linkRow, 1).setValue('📦 Want to Create Your Own Copy?');
    faqSheet.getRange(linkRow, 1).setFontWeight('bold').setFontSize(12);

    // Add description
    faqSheet.getRange(linkRow + 1, 1).setValue(
      'To create a new spreadsheet with seed data and demo features, visit our GitHub repository:'
    );

    // Add the repo link
    var repoUrl = getConfigValue_(CONFIG_COLS.ORG_WEBSITE) || '';
    faqSheet.getRange(linkRow + 2, 1).setValue(repoUrl);
    faqSheet.getRange(linkRow + 2, 1)
      .setFontColor('#1a73e8')
      .setFontWeight('bold');

    // Add instructions
    faqSheet.getRange(linkRow + 3, 1).setValue(
      'The repository includes instructions for setting up a fresh copy with sample data generation (Seed) and data reset (Nuke) features.'
    );

    Logger.log('Added repo link to FAQ sheet');
  } catch (e) {
    Logger.log('Error adding repo link to FAQ: ' + e.message);
  }
}

/**
 * Applies professional tab colors to spreadsheet sheets
 * - Data sheets (Grievance Log, Member Directory): Blue
 * - Documentation (Getting Started, FAQ, Config Guide): Green
 * @param {Spreadsheet} ss - The active spreadsheet
 * @private
 */
function applyTabColors_(ss) {
  // Tab color definitions - Updated per user requirements
  var TAB_COLORS = {
    RED: '#ea4335',         // Red - Getting Started, FAQ
    PURPLE: '#7c3aed',      // Purple - Member Directory, Grievance Log
    YELLOW: '#fbbc04',      // Yellow - Feedback and Development, Function Checklist
    ORANGE: '#ff9800',      // Orange - Config, Config Guide
    BLUE: '#1a73e8',        // Blue - Member Satisfaction
    HIDDEN: '#9e9e9e'       // Gray - for hidden/calc sheets
  };

  // Helper function to find sheet by name (tries exact match and variations)
  function findSheet(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) return sheet;

    // Try without emoji prefix (if name starts with emoji)
    if (name && name.length > 2) {
      var noEmoji = name.replace(/^[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]\s*/u, '');
      if (noEmoji !== name) {
        sheet = ss.getSheetByName(noEmoji);
        if (sheet) return sheet;
      }
    }
    return null;
  }

  // Helper to apply color safely
  function applyColor(sheetName, color) {
    var sheet = findSheet(sheetName);
    if (sheet) {
      try {
        sheet.setTabColor(color);
        Logger.log('Tab color applied to: ' + sheet.getName());
      } catch (e) {
        Logger.log('Tab color error for ' + sheetName + ': ' + e.message);
      }
    } else {
      Logger.log('Sheet not found for tab color: ' + sheetName);
    }
  }

  // Red tabs - Getting Started, FAQ
  applyColor(SHEETS.GETTING_STARTED, TAB_COLORS.RED);
  applyColor(SHEETS.FAQ, TAB_COLORS.RED);

  // Purple tabs - Member Directory, Grievance Log, Case Checklist
  applyColor(SHEETS.MEMBER_DIR, TAB_COLORS.PURPLE);
  applyColor('Member Directory', TAB_COLORS.PURPLE);  // Fallback without constant
  applyColor(SHEETS.GRIEVANCE_LOG, TAB_COLORS.PURPLE);
  applyColor('Grievance Log', TAB_COLORS.PURPLE);  // Fallback without constant
  applyColor(SHEETS.CASE_CHECKLIST, TAB_COLORS.PURPLE);
  applyColor('Case Checklist', TAB_COLORS.PURPLE);  // Fallback without constant

  // Yellow tabs - Feedback and Development, Function Checklist
  applyColor(SHEETS.FEEDBACK, TAB_COLORS.YELLOW);
  applyColor('Feedback & Development', TAB_COLORS.YELLOW);  // Without emoji
  applyColor('Feedback and Development', TAB_COLORS.YELLOW);  // Alternate spelling
  applyColor(SHEETS.FUNCTION_CHECKLIST, TAB_COLORS.YELLOW);

  // Orange tabs - Config, Config Guide
  applyColor(SHEETS.CONFIG, TAB_COLORS.ORANGE);
  applyColor('Config', TAB_COLORS.ORANGE);  // Fallback
  applyColor(SHEETS.CONFIG_GUIDE, TAB_COLORS.ORANGE);

  // Blue tabs - Member Satisfaction
  applyColor(SHEETS.SATISFACTION, TAB_COLORS.BLUE);
  applyColor('Member Satisfaction', TAB_COLORS.BLUE);  // Without emoji

  Logger.log('Tab colors applied successfully');
}

/**
 * Manually applies tab colors (can be called from menu)
 */
function applyTabColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  applyTabColors_(ss);
  ss.toast('Tab colors applied!', 'Success', 3);
}

/**
 * Applies different background colors to Config sheet sections for easy identification
 * Each section gets a distinct color to make navigation easier
 */
function applyConfigSectionColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Config sheet not found!');
    return;
  }

  // Section definitions with colors — derive column ranges from CONFIG_COLS constants
  var CC = typeof CONFIG_COLS !== 'undefined' ? CONFIG_COLS : {};
  var sections = [
    { name: 'Employment Info', startCol: CC.JOB_TITLES || 1, endCol: CC.SUPERVISORS || 5, color: '#e3f2fd' },      // Light Blue
    { name: 'Supervision', startCol: CC.SUPERVISORS || 6, endCol: CC.MANAGERS || 7, color: '#fff3e0' },             // Light Orange
    { name: 'Steward Info', startCol: CC.STEWARDS || 8, endCol: CC.STEWARD_COMMITTEES || 9, color: '#e8f5e9' },     // Light Green
    { name: 'Grievance Settings', startCol: CC.GRIEVANCE_STATUS || 10, endCol: CC.ARTICLES || 13, color: '#fce4ec' }, // Light Pink
    { name: 'Links & Coordinators', startCol: CC.COMM_METHODS || 14, endCol: CC.CONTACT_FORM_URL || 17, color: '#f3e5f5' }, // Light Purple
    { name: 'Notifications', startCol: CC.ADMIN_EMAILS || 18, endCol: CC.NOTIFICATION_RECIPIENTS || 20, color: '#fff8e1' }, // Light Amber
    { name: 'Organization', startCol: CC.ORG_NAME || 21, endCol: CC.MAIN_PHONE || 24, color: '#e0f7fa' },           // Light Cyan
    { name: 'Integration', startCol: CC.DRIVE_FOLDER_ID || 25, endCol: CC.CALENDAR_ID || 26, color: '#fbe9e7' },    // Light Deep Orange
    { name: 'Deadlines', startCol: CC.FILING_DEADLINE_DAYS || 27, endCol: CC.BEST_TIMES || 30, color: '#ffebee' },  // Light Red
    { name: 'Multi-Select Options', startCol: CC.CONTRACT_GRIEVANCE || 31, endCol: CC.CONTRACT_DISCIPLINE || 32, color: '#e8eaf6' }, // Light Indigo
    { name: 'Contract & Legal', startCol: CC.CONTRACT_WORKLOAD || 33, endCol: CC.CONTRACT_NAME || 36, color: '#f1f8e9' }, // Light Light Green
    { name: 'Org Identity', startCol: CC.UNION_PARENT || 37, endCol: CC.ORG_WEBSITE || 39, color: '#eceff1' },      // Light Blue Grey
    { name: 'Extended Contact', startCol: CC.OFFICE_ADDRESSES || 40, endCol: CC.MAIN_CONTACT_EMAIL || 43, color: '#ede7f6' }, // Light Deep Purple
    { name: 'Form Links', startCol: CC.SATISFACTION_FORM_URL || 44, endCol: CC.SATISFACTION_FORM_URL || 44, color: '#e1f5fe' }, // Light Light Blue
    { name: 'Strategic Command Center', startCol: CC.CHIEF_STEWARD_EMAIL || 45, endCol: CC.PDF_FOLDER_ID || 51, color: '#f9fbe7' } // Light Lime
  ];

  var lastRow = Math.max(configSheet.getLastRow(), 50); // At least 50 rows for visibility

  // Apply colors to each section (all rows)
  sections.forEach(function(section) {
    var numCols = section.endCol - section.startCol + 1;
    var range = configSheet.getRange(1, section.startCol, lastRow, numCols);
    range.setBackground(section.color);
  });

  // Make header row (row 1) bold with darker background
  sections.forEach(function(section) {
    var numCols = section.endCol - section.startCol + 1;
    var headerRange = configSheet.getRange(1, section.startCol, 1, numCols);
    headerRange.setFontWeight('bold');
    // Darken the header slightly
    var darkerColor = darkenColor_(section.color, 15);
    headerRange.setBackground(darkerColor);
  });

  // Make column header row (row 2) slightly different
  sections.forEach(function(section) {
    var numCols = section.endCol - section.startCol + 1;
    var colHeaderRange = configSheet.getRange(2, section.startCol, 1, numCols);
    colHeaderRange.setFontWeight('bold');
    colHeaderRange.setFontStyle('italic');
  });

  ss.toast('Config section colors applied!', 'Success', 3);
  Logger.log('Config section colors applied successfully');
}

/**
 * Darkens a hex color by a percentage
 * @param {string} color - Hex color (e.g., '#e3f2fd')
 * @param {number} percent - Percentage to darken (0-100)
 * @returns {string} Darkened hex color
 * @private
 */
function darkenColor_(color, percent) {
  var num = parseInt(color.replace('#', ''), 16);
  var amt = Math.round(2.55 * percent);
  var R = (num >> 16) - amt;
  var G = (num >> 8 & 0x00FF) - amt;
  var B = (num & 0x0000FF) - amt;
  R = Math.max(0, R);
  G = Math.max(0, G);
  B = Math.max(0, B);
  return '#' + (0x1000000 + R * 0x10000 + G * 0x100 + B).toString(16).slice(1);
}

// ============================================================================
// DIAGNOSTIC & REPAIR FUNCTIONS
// ============================================================================

/**
 * Runs a comprehensive diagnostic check and shows UI report
 * NOTE: Renamed from DIAGNOSE_SETUP to avoid duplicate with 06_Maintenance.gs
 * This version shows a UI alert; use DIAGNOSE_SETUP() from 06_Maintenance.gs for programmatic results
 */
function showDiagnosticReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var report = '🔍 COMMAND CENTER DIAGNOSTIC REPORT\n';
  report += '=' .repeat(50) + '\n\n';

  // Check required sheets
  var requiredSheets = [
    SHEETS.CONFIG,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.DASHBOARD
  ];

  report += '📋 SHEET STATUS:\n';
  var allSheetsOK = true;

  requiredSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var rows = sheet.getLastRow();
      report += '  ✅ ' + sheetName + ' (' + rows + ' rows)\n';
    } else {
      report += '  ❌ ' + sheetName + ' (MISSING)\n';
      allSheetsOK = false;
    }
  });

  // Check hidden calculation sheets
  report += '\n📊 HIDDEN CALCULATION SHEETS:\n';
  var hiddenSheets = [
    typeof HIDDEN_SHEETS !== 'undefined' && HIDDEN_SHEETS.CALC_STATS ? HIDDEN_SHEETS.CALC_STATS : '_Dashboard_Calc',
    typeof HIDDEN_SHEETS !== 'undefined' && HIDDEN_SHEETS.GRIEVANCE_CALC ? HIDDEN_SHEETS.GRIEVANCE_CALC : '_Grievance_Calc',
    typeof HIDDEN_SHEETS !== 'undefined' && HIDDEN_SHEETS.MEMBER_LOOKUP ? HIDDEN_SHEETS.MEMBER_LOOKUP : '_Member_Lookup'
  ];

  hiddenSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      report += '  ✅ ' + sheetName + '\n';
    } else {
      report += '  ⚠️ ' + sheetName + ' (not found - may need rebuild)\n';
    }
  });

  // Check configuration
  report += '\n⚙️ CONFIGURATION:\n';
  report += '  Production Mode: ' + (isProductionMode() ? 'ENABLED' : 'DISABLED') + '\n';

  try {
    var templateId = getConfigValue_(CONFIG_COLS.TEMPLATE_ID);
    report += '  PDF Template: ' + (templateId ? '✅ Configured' : '⚠️ Not set') + '\n';
  } catch (_e) {
    report += '  PDF Template: ⚠️ Unable to check\n';
  }

  try {
    var archiveId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID);
    report += '  Archive Folder: ' + (archiveId ? '✅ Configured' : '⚠️ Not set') + '\n';
  } catch (_e) {
    report += '  Archive Folder: ⚠️ Unable to check\n';
  }

  try {
    var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
    report += '  Chief Steward Email: ' + (chiefEmail ? '✅ ' + chiefEmail : '⚠️ Not set') + '\n';
  } catch (_e) {
    report += '  Chief Steward Email: ⚠️ Unable to check\n';
  }

  // Check triggers
  report += '\n⏰ TRIGGERS:\n';
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    report += '  ⚠️ No triggers installed\n';
  } else {
    triggers.forEach(function(trigger) {
      report += '  ✅ ' + trigger.getHandlerFunction() + ' (' + trigger.getEventType() + ')\n';
    });
  }

  // Summary
  report += '\n' + '=' .repeat(50) + '\n';
  report += allSheetsOK ? '✅ All required sheets present\n' : '❌ Some sheets are missing\n';
  report += '\nRun REPAIR_DASHBOARD() to fix issues.\n';

  ui.alert('Diagnostic Results', report, ui.ButtonSet.OK);

  return {
    allSheetsOK: allSheetsOK,
    report: report
  };
}

/**
 * Repairs the dashboard with UI confirmation dialog
 * NOTE: Renamed from REPAIR_DASHBOARD to avoid conflict with 06_Maintenance.gs version
 * This version shows UI; use REPAIR_DASHBOARD() from 06_Maintenance.gs for programmatic repair
 */
function repairDashboardWithUI() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var response = ui.alert(
    '🛠️ Repair Dashboard',
    'This will:\n\n' +
    '• Rebuild missing hidden calculation sheets\n' +
    '• Repair broken formulas\n' +
    '• Refresh all visual styling\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  ss.toast('Starting repair...', COMMAND_CONFIG.SYSTEM_NAME, 10);

  var repairLog = [];

  try {
    // Rebuild hidden sheets if missing
    if (typeof setupAllHiddenSheets === 'function') {
      var result = setupAllHiddenSheets();
      repairLog.push('Hidden sheets: ' + (result.created || 0) + ' created');
    }
  } catch (e) {
    repairLog.push('Hidden sheets: Error - ' + e.message);
  }

  try {
    // Apply theme
    APPLY_SYSTEM_THEME();
    repairLog.push('Theme: Applied successfully');
  } catch (e) {
    repairLog.push('Theme: Error - ' + e.message);
  }

  try {
    // Apply traffic lights
    applyTrafficLightIndicators();
    repairLog.push('Traffic lights: Applied successfully');
  } catch (e) {
    repairLog.push('Traffic lights: Error - ' + e.message);
  }

  try {
    // Sync member data
    if (typeof syncMemberGrievanceData === 'function') {
      syncMemberGrievanceData();
      repairLog.push('Member sync: Completed');
    }
  } catch (e) {
    repairLog.push('Member sync: Error - ' + e.message);
  }

  // Log the repair
  logAuditEvent('DASHBOARD_REPAIR', {
    performedBy: Session.getActiveUser().getEmail(),
    results: repairLog
  });

  ui.alert(
    '🛠️ Repair Complete',
    'Repair Results:\n\n' + repairLog.join('\n'),
    ui.ButtonSet.OK
  );
}

/**
 * Syncs grievance deadlines to Google Calendar
 */
function syncToCalendar() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    '📅 Sync to Calendar',
    'This will sync all open grievance deadlines to your Google Calendar.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    var result = syncDeadlinesToCalendar();

    if (result.success) {
      ui.alert(
        '✅ Calendar Sync Complete',
        'Synced ' + result.synced + ' grievances to calendar.\n' +
        'Skipped ' + result.skipped + ' (no applicable deadlines).',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', 'Sync failed: ' + result.error, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Calendar sync failed: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// BATCH PROCESSING WRAPPERS
// ============================================================================

/**
 * Batch update member (wrapper for updateMemberDataBatch)
 * @param {string} memberId - Member ID
 * @param {Object} updateObj - Fields to update
 */
function updateMemberBatch(memberId, updateObj) {
  return updateMemberDataBatch(memberId, updateObj);
}

// ============================================================================
// QUICK ACTION FUNCTIONS
// ============================================================================

/**
 * Shows quick member search dialog
 */
function showSearchDialog() {
  var html = HtmlService.createHtmlOutput(getSearchDialogHtml_())
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Member Search');
}

/**
 * Generates HTML for search dialog
 * @private
 */
function getSearchDialogHtml_() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      ${getMobileOptimizedHead()}
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
          font-family: 'Google Sans', 'Roboto', sans-serif;
          padding: 20px;
          background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
          min-height: 100%;
          color: #F8FAFC;
        }
        .search-container { margin-bottom: 20px; }
        .search-input {
          width: 100%;
          padding: 12px 15px;
          font-size: 16px;
          border: 2px solid #334155;
          border-radius: 8px;
          background: #1E293B;
          color: #F8FAFC;
          outline: none;
        }
        .search-input:focus { border-color: #7C3AED; }
        .search-input::placeholder { color: #64748B; }
        .btn {
          padding: 10px 20px;
          border: none;
          border-radius: 6px;
          cursor: pointer;
          font-size: 14px;
          font-weight: 500;
          margin-right: 8px;
          transition: all 0.2s;
        }
        .btn-primary {
          background: linear-gradient(135deg, #7C3AED, #5B21B6);
          color: white;
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(124,58,237,0.4); }
        .btn-secondary {
          background: #334155;
          color: #F8FAFC;
        }
        .results {
          margin-top: 20px;
          max-height: 250px;
          overflow-y: auto;
          background: rgba(255,255,255,0.05);
          border-radius: 8px;
          padding: 10px;
        }
        .result-item {
          padding: 10px;
          border-bottom: 1px solid #334155;
          cursor: pointer;
        }
        .result-item:hover { background: rgba(124,58,237,0.2); }
        .result-item:last-child { border-bottom: none; }
        .no-results { text-align: center; color: #64748B; padding: 20px; }
      </style>
    </head>
    <body>
      <div class="search-container">
        <input type="text" class="search-input" id="searchQuery"
               placeholder="Search by name, ID, or email..." autofocus>
      </div>
      <div>
        <button class="btn btn-primary" onclick="doSearch()">🔍 Search</button>
        <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
      </div>
      <div class="results" id="results">
        <div class="no-results">Enter a search term and click Search</div>
      </div>

      <script>
        ${getClientSideEscapeHtml()}
        document.getElementById('searchQuery').addEventListener('keypress', function(e) {
          if (e.key === 'Enter') doSearch();
        });

        function doSearch() {
          var query = document.getElementById('searchQuery').value.trim();
          if (!query) {
            document.getElementById('results').innerHTML = '<div class="no-results">Please enter a search term</div>';
            return;
          }

          document.getElementById('results').innerHTML = '<div class="no-results">Searching...</div>';

          google.script.run
            .withSuccessHandler(function(results) {
              displayResults(results);
            })
            .withFailureHandler(function(e) {
              document.getElementById('results').innerHTML = '<div class="no-results">Error: ' + escapeHtml(e.message) + '</div>';
            })
            .searchMembers(query);
        }

        function displayResults(results) {
          var container = document.getElementById('results');

          if (!results || results.length === 0) {
            container.innerHTML = '<div class="no-results">No members found</div>';
            return;
          }

          var html = '';
          results.forEach(function(member) {
            var name = escapeHtml((member['First Name'] || '') + ' ' + (member['Last Name'] || ''));
            var id = escapeHtml(member['Member ID'] || 'N/A');
            var email = escapeHtml(member['Email'] || '');
            html += '<div class="result-item" data-id="' + id + '" onclick="selectMember(this.dataset.id)">';
            html += '<strong>' + name + '</strong> (' + id + ')';
            if (email) html += '<br><small>' + email + '</small>';
            html += '</div>';
          });

          container.innerHTML = html;
        }

        function selectMember(memberId) {
          google.script.run.navigateToMember(memberId);
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * Navigates to a specific member row in Member Directory
 * @param {string} memberId - The member ID to find
 */
function navigateToMember(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      sheet.activate();
      sheet.setActiveRange(sheet.getRange(i + 1, 1));
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Found: ' + memberId,
        COMMAND_CONFIG.SYSTEM_NAME,
        3
      );
      return;
    }
  }

  SpreadsheetApp.getUi().alert('Member not found: ' + memberId);
}

// ============================================================================
// GEMINI v4.0 UNIFIED MASTER ENGINE - LEGACY CONFIG MAPPING
// ============================================================================

/**
 * Gemini v4.0 Legacy CONFIG Object
 * Maps emoji-prefixed sheet names for backwards compatibility with
 * standalone single-file deployments.
 *
 * This CONFIG mirrors the Gemini v4.0 unified architecture while
 * maintaining compatibility with the modular SHEETS/COMMAND_CONFIG constants.
 */
// @deprecated Use COMMAND_CONFIG (01_Core.gs) instead. Kept for backward compatibility with standalone deployments.
var GEMINI_CONFIG = {
  SYSTEM_NAME: "Strategic Command Center",
  // Legacy emoji-prefixed sheet names (for standalone deployments)
  LOG_SHEET_NAME: "📋 Grievance Log",
  DIR_SHEET_NAME: "👤 Member Directory",
  AUDIT_SHEET_NAME: "🛡️ Audit Log",
  CONFIG_SHEET_NAME: "⚙️ Config",
  DASHBOARD_NAME: "📊 Dashboard",
  MOBILE_VIEW_NAME: "📱 Mobile View",
  // These are read from Config sheet in modular build
  TEMPLATE_ID: '',
  ARCHIVE_FOLDER_ID: '',
  CHIEF_STEWARD_EMAIL: '',
  // Default unit codes (can be overridden in Config sheet)
  UNIT_CODES: { "Main Station": "MS", "Field Ops": "FO", "Health": "HC" },
  THEME: {
    HEADER_BG: '#1e293b',
    HEADER_TEXT: '#ffffff',
    ALT_ROW: '#f8fafc',
    FONT: 'Roboto'
  },
  // Production mode check
  get PRODUCTION_MODE() {
    return PropertiesService.getScriptProperties().getProperty('PRODUCTION_MODE') === 'true';
  }
};

// ============================================================================
// GEMINI v4.0 LEGAL & PDF SIGNATURE ENGINE
// ============================================================================

// Dead code removed: onGrievanceFormSubmit_Legacy_ — replaced by onGrievanceFormSubmit in 05_Integrations.gs

/**
 * Gemini v4.0 Signature-Ready PDF Generator
 * Creates a PDF from template with signature blocks for legal filing.
 *
 * @param {Folder} folder - Google Drive folder for the member
 * @param {Object} data - Grievance data object
 * @returns {File} The generated PDF file
 */
function createGrievancePDF(folder, data) {
  // Get template ID from Config or COMMAND_CONFIG
  var templateId = getConfigValue_(CONFIG_COLS.TEMPLATE_ID) || COMMAND_CONFIG.TEMPLATE_ID;

  if (!templateId) {
    throw new Error('PDF Template ID not configured. Set it in Config sheet column AV.');
  }

  // Make a copy of the template
  var temp = DriveApp.getFileById(templateId).makeCopy('SIGN_REQ_' + data.name, folder);
  var doc = DocumentApp.openById(temp.getId());
  var body = doc.getBody();

  // Replace placeholders with data
  body.replaceText('{{MemberName}}', data.name);
  body.replaceText('{{MemberID}}', data.id);
  body.replaceText('{{Date}}', new Date().toLocaleDateString());
  body.replaceText('{{Details}}', data.details);

  // Add signature blocks
  body.appendParagraph('\n\n' + COMMAND_CONFIG.PDF.SIGNATURE_BLOCK);

  doc.saveAndClose();

  // Convert to PDF
  var pdf = folder.createFile(temp.getAs(MimeType.PDF))
                  .setName('Grievance_UNSIGNED_' + data.name + '_' + new Date().toISOString().split('T')[0] + '.pdf');

  // Trash the temporary doc copy
  temp.setTrashed(true);

  return pdf;
}

/**
 * Gemini v4.0 Member Folder Creator (Legacy)
 * NOTE: Renamed to avoid duplicate - use getOrCreateMemberFolder() in 05_Integrations.gs instead
 *
 * @param {string} name - Member name
 * @param {string} id - Member ID
 * @returns {Folder} Google Drive folder for the member
 * @deprecated Use getOrCreateMemberFolder in 05_Integrations.gs (has better error handling)
 */
function getOrCreateMemberFolder_Legacy_(name, id) {
  // Get archive folder ID from Config
  var archiveFolderId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID) ||
                        getConfigValue_(CONFIG_COLS.DRIVE_FOLDER_ID) ||
                        COMMAND_CONFIG.ARCHIVE_FOLDER_ID;

  if (!archiveFolderId) {
    throw new Error('Archive Folder ID not configured. Set it in Config sheet.');
  }

  var parent = DriveApp.getFolderById(archiveFolderId);
  var folderName = name + ' (' + id + ')';

  // Check if folder already exists
  var iter = parent.getFoldersByName(folderName);
  if (iter.hasNext()) {
    return iter.next();
  }

  // Create new folder
  return parent.createFolder(folderName);
}

/**
 * Gemini v4.0 Enhanced Escalation Alert
 * Sends formatted escalation email to Chief Steward.
 *
 * @param {string} member - Member name
 * @param {string} caseID - Grievance case ID
 * @param {string} status - New status/step
 */
function sendGeminiEscalationAlert(member, caseID, status) {
  var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL) ||
                   COMMAND_CONFIG.CHIEF_STEWARD_EMAIL;

  if (!chiefEmail) {
    console.log('Chief Steward email not configured - skipping escalation alert');
    return;
  }

  try {
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' 🚨 Escalation: ' + caseID;
    var body = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
               '🚨 GRIEVANCE ESCALATION ALERT\n' +
               '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
               'Case ID: ' + caseID + '\n' +
               'Member: ' + member + '\n' +
               'New Status: ' + status + '\n' +
               'Timestamp: ' + new Date().toLocaleString() + '\n\n' +
               'IMMEDIATE ACTION REQUIRED\n' +
               '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━' +
               COMMAND_CONFIG.EMAIL.FOOTER;

    MailApp.sendEmail(chiefEmail, subject, body);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '🚨 Escalation alert sent for ' + caseID,
      COMMAND_CONFIG.SYSTEM_NAME,
      3
    );
  } catch (e) {
    console.error('Escalation email error: ' + e.message);
  }
}

// ============================================================================
// GEMINI v4.0 SCALING MODULES - OCR & SENTIMENT HOOKS
// ============================================================================

/**
 * WGER OCR FUNCTION
 * Main OCR function for transcribing handwritten forms and documents.
 * Uses Google Cloud Vision API for text detection and extraction.
 *
 * Supported formats: PNG, JPG, JPEG, GIF, BMP, WEBP, PDF (first page)
 *
 * @param {string} fileId - Google Drive file ID of the image/document to transcribe
 * @param {Object} options - Optional configuration object
 * @param {string} options.mode - OCR mode: 'TEXT' (default), 'DOCUMENT', 'HANDWRITING'
 * @param {string} options.language - Language hint for OCR: 'en' (default), 'es', 'fr', etc.
 * @param {boolean} options.autoPopulate - If true, attempts to auto-populate grievance fields
 * @param {string} options.targetGrievanceId - Grievance ID to populate (if autoPopulate is true)
 * @returns {Object} Result object with extracted text and metadata
 */
function wger(fileId, options) {
  options = options || {};
  var mode = options.mode || 'TEXT';
  var language = options.language || 'en';
  var autoPopulate = options.autoPopulate || false;
  var targetGrievanceId = options.targetGrievanceId || null;

  try {
    // Validate file ID
    if (!fileId || typeof fileId !== 'string') {
      return {
        success: false,
        status: 'ERROR',
        message: 'Invalid file ID provided',
        text: ''
      };
    }

    var file = DriveApp.getFileById(fileId);
    var fileName = file.getName();
    var mimeType = file.getMimeType();
    var imageBlob = file.getBlob();

    // Validate file type
    var supportedTypes = [
      'image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp',
      'application/pdf'
    ];

    if (supportedTypes.indexOf(mimeType) === -1) {
      return {
        success: false,
        status: 'UNSUPPORTED_FORMAT',
        message: 'Unsupported file format: ' + mimeType + '. Supported: PNG, JPG, GIF, BMP, WEBP, PDF',
        fileName: fileName,
        text: ''
      };
    }

    // Log OCR request for audit
    logAuditEvent('OCR_REQUEST', {
      fileId: fileId,
      fileName: fileName,
      mimeType: mimeType,
      mode: mode,
      requestedBy: Session.getActiveUser().getEmail()
    });

    // Perform OCR using Google Cloud Vision API
    var ocrResult = performCloudVisionOCR_(imageBlob, mode, { language: language });

    if (!ocrResult.success) {
      return {
        success: false,
        status: 'OCR_FAILED',
        message: ocrResult.message,
        fileName: fileName,
        text: ''
      };
    }

    var extractedText = ocrResult.text;

    // Auto-populate grievance fields if requested
    var populationResult = null;
    if (autoPopulate && targetGrievanceId && extractedText) {
      populationResult = autoPopulateGrievanceFromOCR_(extractedText, targetGrievanceId);
    }

    // Log successful OCR
    logAuditEvent('OCR_COMPLETED', {
      fileId: fileId,
      fileName: fileName,
      textLength: extractedText.length,
      autoPopulated: autoPopulate && populationResult && populationResult.success
    });

    return {
      success: true,
      status: 'SUCCESS',
      message: 'OCR completed successfully',
      fileId: fileId,
      fileName: fileName,
      mimeType: mimeType,
      mode: mode,
      text: extractedText,
      textLength: extractedText.length,
      wordCount: extractedText.split(/\s+/).filter(function(w) { return w.length > 0; }).length,
      populationResult: populationResult,
      timestamp: new Date().toISOString()
    };

  } catch (e) {
    console.error('WGER OCR error: ' + e.message);
    logAuditEvent('OCR_ERROR', {
      fileId: fileId,
      error: e.message
    });

    return {
      success: false,
      status: 'ERROR',
      message: 'OCR failed: ' + e.message,
      text: ''
    };
  }
}

/**
 * Performs OCR using Google Cloud Vision API
 * @param {Blob} imageBlob - Image blob to process
 * @param {string} mode - OCR mode (TEXT, DOCUMENT, HANDWRITING)
 * @param {Object} options - Optional configuration
 * @param {string} options.language - Language hint for OCR (default: 'en')
 * @returns {Object} Result with success status and extracted text
 * @private
 */
function performCloudVisionOCR_(imageBlob, mode, options) {
  options = options || {};
  var language = options.language || 'en';

  try {
    // Get API key from script properties (set via Script Properties or Config)
    var apiKey = PropertiesService.getScriptProperties().getProperty('CLOUD_VISION_API_KEY');

    if (!apiKey) {
      // Fallback: Try to use built-in Cloud Vision (requires Cloud project linking)
      return performBuiltInOCR_(imageBlob, mode);
    }

    // Validate file size (Cloud Vision limit is 10MB for base64 encoded images)
    var MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
    var imageBytes = imageBlob.getBytes();
    if (imageBytes.length > MAX_FILE_SIZE) {
      return {
        success: false,
        message: 'File too large (' + Math.round(imageBytes.length / 1024 / 1024 * 10) / 10 + 'MB). Maximum size is 10MB.'
      };
    }

    // Prepare Cloud Vision API request
    var base64Image = Utilities.base64Encode(imageBytes);

    // Select feature type based on mode
    var featureType = 'TEXT_DETECTION';
    if (mode === 'DOCUMENT') {
      featureType = 'DOCUMENT_TEXT_DETECTION';
    } else if (mode === 'HANDWRITING') {
      featureType = 'DOCUMENT_TEXT_DETECTION'; // Best for handwriting
    }

    var requestBody = {
      requests: [{
        image: {
          content: base64Image
        },
        features: [{
          type: featureType,
          maxResults: 1
        }],
        imageContext: {
          languageHints: [language]
        }
      }]
    };

    // API request with explicit timeout (30 seconds)
    // Use X-Goog-Api-Key header instead of URL query string for security
    // (prevents API key exposure in logs, browser history, and referrer headers)
    var response = UrlFetchApp.fetch(
      'https://vision.googleapis.com/v1/images:annotate',
      {
        method: 'POST',
        contentType: 'application/json',
        headers: {
          'X-Goog-Api-Key': apiKey
        },
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true,
        timeout: 30000
      }
    );

    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      console.error('Cloud Vision API error: ' + responseText);
      return {
        success: false,
        message: 'Cloud Vision API returned error code: ' + responseCode
      };
    }

    var result = JSON.parse(responseText);

    // Extract text from response
    var extractedText = '';
    if (result.responses && result.responses[0]) {
      var annotations = result.responses[0];

      if (featureType === 'DOCUMENT_TEXT_DETECTION' && annotations.fullTextAnnotation) {
        extractedText = annotations.fullTextAnnotation.text || '';
      } else if (annotations.textAnnotations && annotations.textAnnotations.length > 0) {
        extractedText = annotations.textAnnotations[0].description || '';
      }
    }

    return {
      success: true,
      text: extractedText.trim()
    };

  } catch (e) {
    console.error('Cloud Vision OCR error: ' + e.message);
    return {
      success: false,
      message: 'Cloud Vision API call failed: ' + e.message
    };
  }
}

/**
 * Fallback OCR using Google's built-in capabilities
 * Attempts to use Google Drive's built-in OCR for PDFs
 * @param {Blob} imageBlob - Image blob to process
 * @param {string} mode - OCR mode
 * @returns {Object} Result with success status and extracted text
 * @private
 */
function performBuiltInOCR_(imageBlob, mode) {
  try {
    var mimeType = imageBlob.getContentType();

    // For images, we need Cloud Vision API
    if (mimeType.indexOf('image/') === 0) {
      return {
        success: false,
        message: 'Cloud Vision API key not configured. Set CLOUD_VISION_API_KEY in Script Properties to enable image OCR.'
      };
    }

    // For PDFs, try Google Drive's OCR conversion
    if (mimeType === 'application/pdf') {
      // Create a temporary file in Drive with OCR enabled
      var resource = {
        title: 'OCR_TEMP_' + new Date().getTime(),
        mimeType: 'application/pdf'
      };

      // Upload and convert to Google Doc (which applies OCR)
      var tempFile = Drive.Files.insert(resource, imageBlob, {
        ocr: true,
        ocrLanguage: 'en'
      });

      // Get the converted document content
      var doc = DocumentApp.openById(tempFile.id);
      var extractedText = doc.getBody().getText();

      // Clean up temp file
      DriveApp.getFileById(tempFile.id).setTrashed(true);

      return {
        success: true,
        text: extractedText.trim()
      };
    }

    return {
      success: false,
      message: 'Could not perform OCR. Please configure Cloud Vision API key.'
    };

  } catch (e) {
    return {
      success: false,
      message: 'Built-in OCR failed: ' + e.message + '. Configure Cloud Vision API key for better results.'
    };
  }
}

/**
 * Attempts to auto-populate grievance fields from OCR text
 * Uses pattern matching to identify common grievance form fields
 * @param {string} text - Extracted OCR text
 * @param {string} grievanceId - Target grievance ID
 * @returns {Object} Result indicating what fields were populated
 * @private
 */
function autoPopulateGrievanceFromOCR_(text, grievanceId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!sheet) {
      return errorResponse('Grievance Log not found');
    }

    // Find the grievance row
    var data = sheet.getDataRange().getValues();
    var grievanceRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
        grievanceRow = i + 1;
        break;
      }
    }

    if (grievanceRow === -1) {
      return errorResponse('Grievance not found: ' + grievanceId);
    }

    var fieldsPopulated = [];
    var _textLower = text.toLowerCase();

    // Pattern matching for common grievance form fields
    var patterns = {
      incidentDate: /(?:incident\s*date|date\s*of\s*incident|occurred\s*on)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
      articles: /(?:article|art\.?)\s*(\d+)/gi,
      issueCategory: /(?:discipline|workload|scheduling|pay|benefits|safety|harassment|discrimination)/i
    };

    // Try to extract incident date with explicit format handling
    var dateMatch = text.match(patterns.incidentDate);
    if (dateMatch) {
      try {
        var dateStr = dateMatch[1];
        var parts = dateStr.split(/[\/\-]/);
        var parsedDate = null;

        if (parts.length === 3) {
          var p0 = parseInt(parts[0], 10);
          var p1 = parseInt(parts[1], 10);
          var p2 = parseInt(parts[2], 10);

          // Handle 2-digit years
          if (p2 < 100) {
            p2 = p2 + (p2 < 50 ? 2000 : 1900);
          }

          if (parts[0].length === 4) {
            // YYYY-MM-DD format
            parsedDate = new Date(p0, p1 - 1, p2);
          } else if (p0 > 12) {
            // DD-MM-YYYY format (day > 12 indicates day-first)
            parsedDate = new Date(p2, p1 - 1, p0);
          } else if (p1 > 12) {
            // MM-DD-YYYY format (second part > 12 indicates month-first)
            parsedDate = new Date(p2, p0 - 1, p1);
          } else {
            // Ambiguous - default to MM-DD-YYYY (US format)
            parsedDate = new Date(p2, p0 - 1, p1);
          }
        }

        if (parsedDate && !isNaN(parsedDate.getTime()) && parsedDate.getFullYear() > 1900) {
          sheet.getRange(grievanceRow, GRIEVANCE_COLS.INCIDENT_DATE).setValue(parsedDate);
          fieldsPopulated.push('Incident Date');
        }
      } catch (e) {
        // Date parsing failed, skip
        console.log('Date parsing failed for: ' + dateMatch[1] + ' - ' + e.message);
      }
    }

    // Try to extract articles violated
    var articleMatches = [];
    var articleMatch;
    while ((articleMatch = patterns.articles.exec(text)) !== null) {
      articleMatches.push('Art. ' + articleMatch[1]);
    }
    if (articleMatches.length > 0) {
      var uniqueArticles = articleMatches.filter(function(v, i, a) { return a.indexOf(v) === i; });
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.ARTICLES).setValue(escapeForFormula(uniqueArticles.join(', ')));
      fieldsPopulated.push('Articles');
    }

    // Try to identify issue category
    var categoryMatch = text.match(patterns.issueCategory);
    if (categoryMatch) {
      var category = categoryMatch[0].charAt(0).toUpperCase() + categoryMatch[0].slice(1).toLowerCase();
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.ISSUE_CATEGORY).setValue(escapeForFormula(category));
      fieldsPopulated.push('Issue Category');
    }

    // Store the full OCR text in resolution notes for reference
    var existingResolution = sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).getValue() || '';
    var ocrNote = '[OCR Extract ' + new Date().toLocaleDateString() + ']: ' + text.substring(0, 500);
    if (text.length > 500) ocrNote += '...';

    if (existingResolution) {
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).setValue(escapeForFormula(existingResolution + '\n\n' + ocrNote));
    } else {
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).setValue(escapeForFormula(ocrNote));
    }
    fieldsPopulated.push('Resolution Notes (OCR text)');

    return {
      success: true,
      message: 'Populated ' + fieldsPopulated.length + ' fields',
      fieldsPopulated: fieldsPopulated
    };

  } catch (e) {
    return {
      success: false,
      message: 'Auto-populate failed: ' + e.message
    };
  }
}

// Dead code removed: wgerQuick, wgerHandwriting, wgerDocument, transcribeHandwrittenForm
// — unused convenience wrappers. Use wger(fileId, {mode:'TEXT'|'HANDWRITING'|'DOCUMENT'}) directly.

/**
 * EXTENSION: SENTIMENT CORRELATION HOOK
 * Compares grievance activity to survey results for unit health analysis.
 * Flags units with high dissatisfaction but low representation.
 *
 * @param {string} unitName - The unit to analyze
 * @returns {Object} Unit health analysis result
 */
function calculateUnitHealth(unitName) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  // Count grievances for this unit
  var grievanceCount = getGrievanceCountForUnit(unitName);

  // Get recent survey average (placeholder - connect to Typeform/SurveyMonkey)
  var surveyScore = getRecentSurveyAverage(unitName);

  var result = {
    unit: unitName,
    grievanceCount: grievanceCount,
    surveyScore: surveyScore,
    status: '',
    recommendation: ''
  };

  // Sentiment correlation logic
  if (surveyScore < 3 && grievanceCount === 0) {
    result.status = '🚩 RED FLAG';
    result.recommendation = 'High Dissatisfaction / Low Representation - Investigate immediately';
  } else if (surveyScore < 5 && grievanceCount < 2) {
    result.status = '⚠️ WARNING';
    result.recommendation = 'Moderate dissatisfaction with limited grievance activity';
  } else if (surveyScore >= 7) {
    result.status = '✅ HEALTHY';
    result.recommendation = 'Unit appears stable with adequate representation';
  } else {
    result.status = '📊 MONITORING';
    result.recommendation = 'Standard activity levels - continue monitoring';
  }

  return result;
}

/**
 * Helper: Count grievances for a specific unit
 * @param {string} unitName - Unit name to count
 * @returns {number} Number of grievances
 */
function getGrievanceCountForUnit(unitName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) return 0;

  var data = sheet.getRange(2, GRIEVANCE_COLS.LOCATION, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;

  data.forEach(function(row) {
    if (row[0] === unitName) count++;
  });

  return count;
}

/**
 * Helper: Get recent survey average for a unit
 * Reads from Member Satisfaction sheet, filtering by worksite/unit.
 *
 * @param {string} unitName - Unit name to filter by
 * @returns {number} Average survey score (1-10), or 5 if no data
 */
function getRecentSurveyAverage(unitName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet || satSheet.getLastRow() < 2) {
    return 5; // Neutral score if no survey data
  }

  var lastRow = satSheet.getLastRow();
  var lastCol = (typeof SATISFACTION_COLS !== 'undefined' && SATISFACTION_COLS.AVG_SCHEDULING) ? SATISFACTION_COLS.AVG_SCHEDULING : satSheet.getLastColumn();
  var data = satSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Load vault data to check verified/isLatest status (PII stays in vault)
  var vaultMap = getVaultDataMap_();

  var scores = [];
  var unitLower = unitName.toString().trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var satRow = i + 2; // 1-indexed sheet row

    // Only include verified, latest responses (checked via vault)
    var vEntry = vaultMap[satRow];
    if (!vEntry || !isTruthyValue(vEntry.verified) || !isTruthyValue(vEntry.isLatest)) continue;

    // Check if worksite matches unit (partial match for flexibility)
    var worksite = (row[SATISFACTION_COLS.Q1_WORKSITE - 1] || '').toString().trim().toLowerCase();

    if (worksite.indexOf(unitLower) !== -1 || unitLower.indexOf(worksite) !== -1) {
      // Get overall satisfaction average (pre-calculated column)
      var avgScore = parseFloat(row[SATISFACTION_COLS.AVG_OVERALL_SAT - 1]);

      if (!isNaN(avgScore) && avgScore > 0) {
        scores.push(avgScore);
      } else {
        // Fallback: calculate from individual questions Q6-Q9
        var q6 = parseFloat(row[SATISFACTION_COLS.Q6_SATISFIED_REP - 1]) || 0;
        var q7 = parseFloat(row[SATISFACTION_COLS.Q7_TRUST_UNION - 1]) || 0;
        var q8 = parseFloat(row[SATISFACTION_COLS.Q8_FEEL_PROTECTED - 1]) || 0;
        var q9 = parseFloat(row[SATISFACTION_COLS.Q9_RECOMMEND - 1]) || 0;

        var count = (q6 > 0 ? 1 : 0) + (q7 > 0 ? 1 : 0) + (q8 > 0 ? 1 : 0) + (q9 > 0 ? 1 : 0);
        if (count > 0) {
          scores.push((q6 + q7 + q8 + q9) / count);
        }
      }
    }
  }

  if (scores.length === 0) {
    return 5; // Neutral if no matching responses
  }

  // Calculate average of all matching scores
  var total = scores.reduce(function(sum, val) { return sum + val; }, 0);
  return Math.round((total / scores.length) * 10) / 10; // Round to 1 decimal
}

/**
 * Show Unit Health Report Dialog
 * Displays sentiment analysis for all units using Member Satisfaction data.
 */
function showUnitHealthReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get all unique units from Config
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var units = [];

  if (configSheet) {
    var unitData = configSheet.getRange(3, CONFIG_COLS.UNITS, 50, 1).getValues();
    unitData.forEach(function(row) {
      if (row[0]) units.push(row[0]);
    });
  }

  if (units.length === 0) {
    units = Object.keys(GEMINI_CONFIG.UNIT_CODES);
  }

  var report = '📊 UNIT HEALTH ANALYSIS REPORT\n';
  report += '═'.repeat(45) + '\n\n';

  var redFlags = [];
  var warnings = [];
  var healthy = [];

  units.forEach(function(unit) {
    var health = calculateUnitHealth(unit);
    var entry = health.status + ' ' + health.unit + '\n';
    entry += '  Grievances: ' + health.grievanceCount + ' | Survey: ' + health.surveyScore + '/10\n';
    entry += '  → ' + health.recommendation + '\n';

    if (health.status.indexOf('RED FLAG') !== -1) {
      redFlags.push(entry);
    } else if (health.status.indexOf('WARNING') !== -1) {
      warnings.push(entry);
    } else {
      healthy.push(entry);
    }
  });

  // Show red flags first, then warnings, then healthy
  if (redFlags.length > 0) {
    report += '⚠️ REQUIRES ATTENTION:\n' + redFlags.join('\n') + '\n';
  }
  if (warnings.length > 0) {
    report += '📋 MONITORING:\n' + warnings.join('\n') + '\n';
  }
  if (healthy.length > 0) {
    report += '✅ STABLE:\n' + healthy.join('\n');
  }

  report += '\n' + '═'.repeat(45) + '\n';
  report += 'Data source: Member Satisfaction Survey + Grievance Log';

  ui.alert('Unit Health Report', report, ui.ButtonSet.OK);
}

// ============================================================================
// GEMINI v4.0 APPLY SYSTEM THEME (UI Refresh)
// ============================================================================

/**
 * Gemini v4.0 System Theme Application
 * Wrapper for APPLY_SYSTEM_THEME that ensures Gemini compatibility.
 */
function APPLY_GEMINI_THEME() {
  // Use existing APPLY_SYSTEM_THEME if available
  if (typeof APPLY_SYSTEM_THEME === 'function') {
    APPLY_SYSTEM_THEME();
    return;
  }

  // Fallback: Apply basic Roboto theme
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];

  sheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
           .setFontFamily(GEMINI_CONFIG.THEME.FONT)
           .setFontSize(10);
    }
  });

  ss.toast('✅ Theme applied.', COMMAND_CONFIG.SYSTEM_NAME, 3);
}

// ============================================================================
// ANALYTICS & INSIGHTS FUNCTIONS (v4.0 Scaling)
// ============================================================================

/**
 * Shows grievance trends over time
 * Displays counts per month and identifies patterns
 */
function showGrievanceTrends() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No Grievance Data', 'No grievances found to analyze.', ui.ButtonSet.OK);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var header = data[0];

  // Use GRIEVANCE_COLS constant for date column (subtract 1 for 0-indexed array access)
  var dateCol = (typeof GRIEVANCE_COLS !== 'undefined' && GRIEVANCE_COLS.DATE_FILED) ? GRIEVANCE_COLS.DATE_FILED - 1 : header.indexOf('Date Filed');
  if (dateCol === -1) dateCol = 1; // Default to column B

  // Aggregate by month
  var monthCounts = {};

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][dateCol];
    if (dateVal) {
      var date = new Date(dateVal);
      if (!isNaN(date.getTime())) {
        var monthKey = date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
        monthCounts[monthKey] = (monthCounts[monthKey] || 0) + 1;
      }
    }
  }

  // Build report
  var report = '📈 GRIEVANCE TRENDS\n';
  report += '═'.repeat(40) + '\n\n';

  var months = Object.keys(monthCounts).sort();
  if (months.length === 0) {
    report += 'No dated grievances found.\n';
  } else {
    report += 'Month          | Count | Trend\n';
    report += '───────────────|───────|──────\n';

    var prevCount = 0;
    months.forEach(function(month) {
      var count = monthCounts[month];
      var trend = count > prevCount ? '📈 Up' : (count < prevCount ? '📉 Down' : '➡️ Flat');
      report += month.padEnd(14) + ' | ' + String(count).padStart(5) + ' | ' + trend + '\n';
      prevCount = count;
    });

    // Summary
    var total = months.reduce(function(sum, m) { return sum + monthCounts[m]; }, 0);
    var avg = (total / months.length).toFixed(1);
    report += '\n───────────────────────────────\n';
    report += 'Total: ' + total + ' grievances over ' + months.length + ' months\n';
    report += 'Average: ' + avg + ' per month\n';
  }

  ui.alert('Grievance Trends', report, ui.ButtonSet.OK);
}

/**
 * Shows OCR transcription dialog
 * Uses the wger OCR function for text extraction
 */
function showOCRDialog() {
  var ui = SpreadsheetApp.getUi();

  var html = HtmlService.createHtmlOutput(
    '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
    '<style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: "Google Sans", Roboto, Arial, sans-serif; padding: 20px; margin: 0; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; }' +
    'h3 { color: #F8FAFC; margin: 0 0 16px 0; display: flex; align-items: center; gap: 8px; }' +
    '.input-group { margin-bottom: 16px; }' +
    '.input-group label { display: block; font-size: 13px; color: #94A3B8; margin-bottom: 6px; }' +
    'input, select { width: 100%; padding: 12px; border: 2px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-size: 14px; outline: none; }' +
    'input:focus, select:focus { border-color: #7C3AED; }' +
    'input::placeholder { color: #64748B; }' +
    '.btn-row { display: flex; gap: 10px; margin-top: 16px; }' +
    'button { flex: 1; padding: 12px 20px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; }' +
    '.btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }' +
    '.btn-primary:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(124,58,237,0.4); }' +
    '.btn-secondary { background: #334155; color: #F8FAFC; }' +
    '.btn-secondary:hover { background: #475569; }' +
    '.result-box { margin-top: 16px; padding: 16px; background: rgba(255,255,255,0.05); border-radius: 8px; border: 1px solid #334155; max-height: 200px; overflow-y: auto; }' +
    '.result-box.success { border-color: #10B981; }' +
    '.result-box.error { border-color: #EF4444; }' +
    '.result-header { font-size: 12px; color: #94A3B8; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }' +
    '.result-text { font-size: 14px; white-space: pre-wrap; word-break: break-word; }' +
    '.stats { display: flex; gap: 16px; margin-top: 12px; font-size: 12px; color: #64748B; }' +
    '.loading { text-align: center; padding: 20px; }' +
    '.spinner { display: inline-block; width: 24px; height: 24px; border: 3px solid #334155; border-top-color: #7C3AED; border-radius: 50%; animation: spin 1s linear infinite; }' +
    '@keyframes spin { to { transform: rotate(360deg); } }' +
    '.help-text { font-size: 12px; color: #64748B; margin-top: 4px; }' +
    '</style>' +
    '<h3>📝 WGER OCR Transcription</h3>' +
    '<div class="input-group">' +
    '  <label>Google Drive File ID or URL</label>' +
    '  <input type="text" id="fileId" placeholder="Paste file ID or Drive URL...">' +
    '  <div class="help-text">Supports: PNG, JPG, GIF, BMP, WEBP, PDF</div>' +
    '</div>' +
    '<div class="input-group">' +
    '  <label>OCR Mode</label>' +
    '  <select id="ocrMode">' +
    '    <option value="TEXT">Standard Text Detection</option>' +
    '    <option value="HANDWRITING">Handwriting Optimized</option>' +
    '    <option value="DOCUMENT">Document Layout Preservation</option>' +
    '  </select>' +
    '</div>' +
    '<div class="input-group">' +
    '  <label>Language Hint</label>' +
    '  <select id="ocrLanguage">' +
    '    <option value="en">English</option>' +
    '    <option value="es">Spanish (Español)</option>' +
    '    <option value="fr">French (Français)</option>' +
    '    <option value="de">German (Deutsch)</option>' +
    '    <option value="pt">Portuguese (Português)</option>' +
    '    <option value="zh">Chinese (中文)</option>' +
    '    <option value="ja">Japanese (日本語)</option>' +
    '    <option value="ko">Korean (한국어)</option>' +
    '  </select>' +
    '</div>' +
    '<div class="btn-row">' +
    '  <button class="btn-primary" onclick="runOCR()">🔍 Extract Text</button>' +
    '  <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div id="resultContainer"></div>' +
    '<script>' +
    'function extractFileId(input) {' +
    '  input = input.trim();' +
    '  // Handle various Google Drive URL formats' +
    '  var patterns = [' +
    '    /\\/d\\/([a-zA-Z0-9_-]{25,})/,           ' + // /d/{fileId}/ format
    '    /[?&]id=([a-zA-Z0-9_-]{25,})/,          ' + // ?id={fileId} format
    '    /\\/file\\/d\\/([a-zA-Z0-9_-]{25,})/,   ' + // /file/d/{fileId}/ format
    '    /\\/open\\?id=([a-zA-Z0-9_-]{25,})/,    ' + // /open?id={fileId} format
    '    /^([a-zA-Z0-9_-]{25,})$/                ' + // Plain file ID
    '  ];' +
    '  for (var i = 0; i < patterns.length; i++) {' +
    '    var match = input.match(patterns[i]);' +
    '    if (match && match[1]) {' +
    '      return match[1];' +
    '    }' +
    '  }' +
    '  // Fallback: try to find any 25+ char alphanumeric string' +
    '  var fallbackMatch = input.match(/[a-zA-Z0-9_-]{25,}/);' +
    '  return fallbackMatch ? fallbackMatch[0] : input;' +
    '}' +
    'function runOCR() {' +
    '  var rawInput = document.getElementById("fileId").value;' +
    '  var fileId = extractFileId(rawInput);' +
    '  var mode = document.getElementById("ocrMode").value;' +
    '  var language = document.getElementById("ocrLanguage").value;' +
    '  if (!fileId) {' +
    '    showResult({ success: false, message: "Please enter a valid file ID or URL" });' +
    '    return;' +
    '  }' +
    '  document.getElementById("resultContainer").innerHTML = "<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:12px\\">Processing OCR...</div></div>";' +
    '  google.script.run.withSuccessHandler(showResult).withFailureHandler(function(e) {' +
    '    showResult({ success: false, message: e.message });' +
    '  }).wger(fileId, { mode: mode, language: language });' +
    '}' +
    'function showResult(result) {' +
    '  var container = document.getElementById("resultContainer");' +
    '  if (result.success) {' +
    '    container.innerHTML = "<div class=\\"result-box success\\">" +' +
    '      "<div class=\\"result-header\\">✅ Extracted Text</div>" +' +
    '      "<div class=\\"result-text\\">" + escapeHtml(result.text || "(No text found)") + "</div>" +' +
    '      "<div class=\\"stats\\"><span>📄 " + (result.wordCount || 0) + " words</span><span>📏 " + (result.textLength || 0) + " chars</span></div>" +' +
    '    "</div>";' +
    '  } else {' +
    '    container.innerHTML = "<div class=\\"result-box error\\">" +' +
    '      "<div class=\\"result-header\\">❌ OCR Failed</div>" +' +
    '      "<div class=\\"result-text\\">" + escapeHtml(result.message || "Unknown error") + "</div>" +' +
    '    "</div>";' +
    '  }' +
    '}' +
    getClientSideEscapeHtml() +
    'document.getElementById("fileId").addEventListener("keypress", function(e) {' +
    '  if (e.key === "Enter") runOCR();' +
    '});' +
    '</script>'
  )
  .setWidth(500)
  .setHeight(540);

  ui.showModalDialog(html, '📝 WGER OCR Transcription');
}

/**
 * OCR Setup Helper - Guides users through Cloud Vision API configuration
 * Checks current status, provides instructions, and allows API key entry
 */
function setupOCRApiKey() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var currentKey = props.getProperty('CLOUD_VISION_API_KEY');

  // Check current status
  var statusMessage = currentKey
    ? '✅ API Key Configured (ends with: ...' + currentKey.slice(-6) + ')'
    : '❌ API Key Not Configured';

  var html = HtmlService.createHtmlOutput(
    '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
    '<style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: "Google Sans", Roboto, Arial, sans-serif; padding: 24px; margin: 0; background: #F8FAFC; color: #1E293B; }' +
    'h2 { margin: 0 0 8px 0; color: #1E293B; }' +
    '.status { padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }' +
    '.status.configured { background: #DCFCE7; color: #166534; border: 1px solid #86EFAC; }' +
    '.status.not-configured { background: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; }' +
    '.section { background: white; border-radius: 12px; padding: 20px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    '.section h3 { margin: 0 0 12px 0; font-size: 16px; color: #334155; }' +
    '.step { display: flex; gap: 12px; margin-bottom: 12px; }' +
    '.step-num { width: 28px; height: 28px; background: #7C3AED; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 14px; flex-shrink: 0; }' +
    '.step-text { font-size: 14px; line-height: 1.5; color: #475569; }' +
    '.step-text a { color: #7C3AED; }' +
    '.input-group { margin-top: 16px; }' +
    '.input-group label { display: block; font-size: 13px; font-weight: 500; color: #64748B; margin-bottom: 6px; }' +
    'input[type="text"] { width: 100%; padding: 12px; border: 2px solid #E2E8F0; border-radius: 8px; font-size: 14px; font-family: monospace; }' +
    'input[type="text"]:focus { outline: none; border-color: #7C3AED; }' +
    '.btn-row { display: flex; gap: 10px; margin-top: 20px; }' +
    'button { flex: 1; padding: 12px 20px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; }' +
    '.btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }' +
    '.btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }' +
    '.btn-secondary { background: #E2E8F0; color: #475569; }' +
    '.btn-secondary:hover { background: #CBD5E1; }' +
    '.btn-test { background: #0EA5E9; color: white; }' +
    '.btn-test:hover { background: #0284C7; }' +
    '.result { margin-top: 16px; padding: 12px; border-radius: 8px; font-size: 14px; display: none; }' +
    '.result.success { display: block; background: #DCFCE7; color: #166534; border: 1px solid #86EFAC; }' +
    '.result.error { display: block; background: #FEE2E2; color: #991B1B; border: 1px solid #FECACA; }' +
    '.result.info { display: block; background: #DBEAFE; color: #1E40AF; border: 1px solid #93C5FD; }' +
    '.free-tier { background: #F0FDF4; border: 1px solid #86EFAC; border-radius: 8px; padding: 12px; margin-top: 12px; font-size: 13px; color: #166534; }' +
    '</style>' +
    '<h2>🔧 OCR Setup</h2>' +
    '<div class="status ' + (currentKey ? 'configured' : 'not-configured') + '">' + statusMessage + '</div>' +

    '<div class="section">' +
    '<h3>📋 Setup Instructions</h3>' +
    '<div class="step"><div class="step-num">1</div><div class="step-text">Go to <a href="https://console.cloud.google.com/apis/library/vision.googleapis.com" target="_blank">Google Cloud Console → Vision API</a></div></div>' +
    '<div class="step"><div class="step-num">2</div><div class="step-text">Click <strong>"Enable"</strong> to activate Cloud Vision API for your project</div></div>' +
    '<div class="step"><div class="step-num">3</div><div class="step-text">Go to <a href="https://console.cloud.google.com/apis/credentials" target="_blank">APIs & Services → Credentials</a></div></div>' +
    '<div class="step"><div class="step-num">4</div><div class="step-text">Click <strong>"+ CREATE CREDENTIALS"</strong> → <strong>"API key"</strong></div></div>' +
    '<div class="step"><div class="step-num">5</div><div class="step-text">Copy the API key and paste it below</div></div>' +
    '<div class="free-tier">💰 <strong>Free Tier:</strong> 1,000 OCR requests/month at no cost. Most unions will never exceed this.</div>' +
    '</div>' +

    '<div class="section">' +
    '<h3>🔑 Enter API Key</h3>' +
    '<div class="input-group">' +
    '<label>Cloud Vision API Key' + (currentKey ? ' <span style="color:#10B981;font-weight:normal;">(currently configured)</span>' : '') + '</label>' +
    '<input type="password" id="apiKey" placeholder="AIzaSy..." autocomplete="off">' +
    '<div style="margin-top:8px;font-size:12px;color:#64748B;">' +
    '<label style="cursor:pointer;display:inline-flex;align-items:center;gap:4px;">' +
    '<input type="checkbox" id="showKey" onchange="toggleKeyVisibility()"> Show API key' +
    '</label>' +
    '</div>' +
    '</div>' +
    '<div class="btn-row">' +
    '<button class="btn-primary" onclick="saveKey()">💾 Save Key</button>' +
    '<button class="btn-test" onclick="testKey()">🧪 Test OCR</button>' +
    '</div>' +
    '<div id="result" class="result"></div>' +
    '</div>' +

    '<div class="btn-row">' +
    '<button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +

    '<script>' +
    'function toggleKeyVisibility() {' +
    '  var input = document.getElementById("apiKey");' +
    '  var checkbox = document.getElementById("showKey");' +
    '  input.type = checkbox.checked ? "text" : "password";' +
    '}' +
    'function saveKey() {' +
    '  var key = document.getElementById("apiKey").value.trim();' +
    '  if (!key) {' +
    '    showResult("error", "Please enter an API key");' +
    '    return;' +
    '  }' +
    '  showResult("info", "Saving...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.success) {' +
    '        showResult("success", "✅ " + result.message);' +
    '      } else {' +
    '        showResult("error", "❌ " + result.message);' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      showResult("error", "❌ Error: " + e.message);' +
    '    })' +
    '    .saveOCRApiKey(key);' +
    '}' +
    'function testKey() {' +
    '  showResult("info", "🔄 Testing OCR... (this may take a few seconds)");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.success) {' +
    '        showResult("success", "✅ " + result.message);' +
    '      } else {' +
    '        showResult("error", "❌ " + result.message);' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      showResult("error", "❌ Error: " + e.message);' +
    '    })' +
    '    .testOCRConnection();' +
    '}' +
    'function showResult(type, message) {' +
    '  var el = document.getElementById("result");' +
    '  el.className = "result " + type;' +
    '  el.textContent = message;' +
    '}' +
    '</script>'
  )
  .setWidth(520)
  .setHeight(620);

  ui.showModalDialog(html, '🔧 OCR Setup - Cloud Vision API');
}

/**
 * Saves the Cloud Vision API key to Script Properties
 * @param {string} apiKey - The API key to save
 * @returns {Object} Result with success status
 */
function saveOCRApiKey(apiKey) {
  try {
    if (!apiKey || apiKey.trim().length < 10) {
      return errorResponse('Invalid API key format');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty('CLOUD_VISION_API_KEY', apiKey.trim());

    // Log the configuration for audit
    logAuditEvent('OCR_API_KEY_CONFIGURED', {
      configuredBy: Session.getActiveUser().getEmail(),
      keyPreview: '...' + apiKey.slice(-6)
    });

    return { success: true, message: 'API key saved successfully! You can now use OCR features.' };
  } catch (e) {
    return errorResponse('Failed to save: ' + e.message);
  }
}

/**
 * Tests the OCR connection by making a simple API validation call
 * @returns {Object} Result with success status
 */
function testOCRConnection() {
  try {
    var props = PropertiesService.getScriptProperties();
    var apiKey = props.getProperty('CLOUD_VISION_API_KEY');

    if (!apiKey) {
      return errorResponse('No API key configured. Please save an API key first.');
    }

    // Make a minimal test request to validate the API key
    var testRequest = {
      requests: [{
        image: {
          content: Utilities.base64Encode(Utilities.newBlob('Test').getBytes())
        },
        features: [{
          type: 'TEXT_DETECTION',
          maxResults: 1
        }]
      }]
    };

    var response = UrlFetchApp.fetch(
      'https://vision.googleapis.com/v1/images:annotate',
      {
        method: 'POST',
        contentType: 'application/json',
        headers: { 'X-Goog-Api-Key': apiKey },
        payload: JSON.stringify(testRequest),
        muteHttpExceptions: true
      }
    );

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      return { success: true, message: 'OCR connection successful! API key is valid and working.' };
    } else if (responseCode === 400) {
      // 400 with valid key often means the test image was invalid, but key works
      var errorData = JSON.parse(responseBody);
      if (errorData.error && errorData.error.message.indexOf('image') !== -1) {
        return { success: true, message: 'API key is valid! OCR is ready to use.' };
      }
      return errorResponse('API Error: ' + (errorData.error ? errorData.error.message : responseBody));
    } else if (responseCode === 403) {
      return errorResponse('API key invalid or Cloud Vision API not enabled. Check your Google Cloud Console.');
    } else {
      return errorResponse('Unexpected response (' + responseCode + '): ' + responseBody.substring(0, 200));
    }

  } catch (e) {
    return errorResponse('Connection test failed: ' + e.message);
  }
}

/**
 * Checks if OCR is configured and ready to use
 * @returns {Object} Status object
 */
function getOCRStatus() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('CLOUD_VISION_API_KEY');

  return {
    configured: !!apiKey,
    message: apiKey ? 'OCR is configured and ready' : 'OCR requires setup - run setupOCRApiKey()'
  };
}

// ============================================================================
// SEARCH PRECEDENTS (v4.1 - Historical Grievance Outcomes)
// ============================================================================

/**
 * Shows Search Precedents dialog for finding historical grievance outcomes.
 * Helps stewards cite "Past Practice" during Step 1 meetings.
 */
function showSearchPrecedents() {
  var ui = SpreadsheetApp.getUi();

  var html = HtmlService.createHtmlOutput(
    '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
    '<style>' +
    'body { font-family: Roboto, Arial, sans-serif; padding: 16px; margin: 0; }' +
    'h3 { color: #1e293b; margin: 0 0 16px 0; display: flex; align-items: center; gap: 8px; }' +
    '.search-box { display: flex; gap: 8px; margin-bottom: 16px; }' +
    'input, select { padding: 10px; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 14px; }' +
    'input[type="text"] { flex: 1; }' +
    'select { min-width: 140px; }' +
    'button { background: #1e293b; color: white; border: none; padding: 10px 16px; border-radius: 6px; cursor: pointer; font-weight: 500; }' +
    'button:hover { background: #334155; }' +
    '.results { max-height: 320px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px; }' +
    '.result-item { padding: 12px; border-bottom: 1px solid #f1f5f9; }' +
    '.result-item:last-child { border-bottom: none; }' +
    '.result-item:hover { background: #f8fafc; }' +
    '.result-header { display: flex; justify-content: space-between; margin-bottom: 6px; }' +
    '.result-id { font-weight: 600; color: #1e293b; }' +
    '.result-status { font-size: 12px; padding: 2px 8px; border-radius: 12px; }' +
    '.status-won { background: #dcfce7; color: #166534; }' +
    '.status-lost { background: #fee2e2; color: #991b1b; }' +
    '.status-settled { background: #dbeafe; color: #1e40af; }' +
    '.status-withdrawn { background: #f3f4f6; color: #6b7280; }' +
    '.result-meta { font-size: 13px; color: #64748b; margin-bottom: 4px; }' +
    '.result-outcome { font-size: 13px; color: #334155; background: #f8fafc; padding: 8px; border-radius: 4px; margin-top: 6px; }' +
    '.empty { text-align: center; padding: 40px; color: #94a3b8; }' +
    '.help-text { font-size: 12px; color: #64748b; margin-bottom: 12px; }' +
    '.copy-btn { font-size: 11px; padding: 4px 8px; background: #e2e8f0; color: #475569; margin-left: 8px; }' +
    '.copy-btn:hover { background: #cbd5e1; }' +
    '</style>' +
    '<h3>📚 Search Precedents</h3>' +
    '<p class="help-text">Find resolved grievances to cite as "Past Practice" in Step 1 meetings.</p>' +
    '<div class="search-box">' +
    '  <input type="text" id="searchQuery" placeholder="Search by keyword, article, or issue...">' +
    '  <select id="filterOutcome">' +
    '    <option value="">All Outcomes</option>' +
    '    <option value="won">Won / Sustained</option>' +
    '    <option value="settled">Settled</option>' +
    '    <option value="lost">Denied / Lost</option>' +
    '    <option value="withdrawn">Withdrawn</option>' +
    '  </select>' +
    '  <button onclick="searchPrecedents()">Search</button>' +
    '</div>' +
    '<div id="results" class="results">' +
    '  <div class="empty">Enter search terms to find historical grievances</div>' +
    '</div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'function searchPrecedents() {' +
    '  var query = document.getElementById("searchQuery").value;' +
    '  var outcomeFilter = document.getElementById("filterOutcome").value;' +
    '  document.getElementById("results").innerHTML = "<div class=\\"empty\\">Searching...</div>";' +
    '  google.script.run.withSuccessHandler(displayResults).searchPrecedentsData(query, outcomeFilter);' +
    '}' +
    'function displayResults(data) {' +
    '  var container = document.getElementById("results");' +
    '  if (!data || data.length === 0) {' +
    '    container.innerHTML = "<div class=\\"empty\\">No matching precedents found</div>";' +
    '    return;' +
    '  }' +
    '  var html = "";' +
    '  data.forEach(function(item) {' +
    '    var statusClass = "status-" + escapeHtml(item.outcomeClass);' +
    '    html += "<div class=\\"result-item\\">" +' +
    '      "<div class=\\"result-header\\">" +' +
    '        "<span class=\\"result-id\\">" + escapeHtml(item.id) + "</span>" +' +
    '        "<span class=\\"result-status " + statusClass + "\\">" + escapeHtml(item.outcome) + "</span>" +' +
    '      "</div>" +' +
    '      "<div class=\\"result-meta\\"><strong>" + escapeHtml(item.issueCategory) + "</strong> | " + escapeHtml(item.article) + "</div>" +' +
    '      "<div class=\\"result-meta\\">" + escapeHtml(item.memberName) + " • " + escapeHtml(item.location) + " • " + escapeHtml(item.dateResolved) + "</div>" +' +
    '      (item.resolution ? "<div class=\\"result-outcome\\"><strong>Resolution:</strong> " + escapeHtml(item.resolution) + "<button class=\\"copy-btn\\" onclick=\\"copyText(\'" + escapeHtml(item.id).replace(/\'/g,"") + ": " + escapeHtml(item.resolution).replace(/\'/g, "") + "\')\\">Copy</button></div>" : "") +' +
    '    "</div>";' +
    '  });' +
    '  container.innerHTML = html;' +
    '}' +
    'function copyText(text) {' +
    '  navigator.clipboard.writeText(text).then(function() {' +
    '    alert("Copied to clipboard!");' +
    '  });' +
    '}' +
    'document.getElementById("searchQuery").addEventListener("keypress", function(e) {' +
    '  if (e.key === "Enter") searchPrecedents();' +
    '});' +
    '</script>'
  )
  .setWidth(600)
  .setHeight(520);

  ui.showModalDialog(html, '📚 Search Precedents - Past Practice');
}

/**
 * Backend function to search resolved grievances for precedent data.
 * Called from the Search Precedents dialog.
 *
 * @param {string} query - Search query (keyword, article, issue)
 * @param {string} outcomeFilter - Filter by outcome type (won, lost, settled, withdrawn, or empty for all)
 * @returns {Array} Array of matching grievance objects
 */
function searchPrecedentsData(query, outcomeFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();

  var results = [];
  var queryLower = (query || '').toString().toLowerCase().trim();

  // Closed/resolved statuses to include
  var closedStatuses = ['closed', 'resolved', 'won', 'lost', 'denied', 'sustained', 'settled', 'withdrawn', 'dismissed'];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    // Get status and check if closed/resolved
    var status = (row[GRIEVANCE_COLS.STATUS - 1] || '').toString().toLowerCase();
    var isClosed = closedStatuses.some(function(s) { return status.indexOf(s) !== -1; });

    if (!isClosed) continue;

    // Get grievance data
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
    var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
    var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';
    var issueCategory = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '';
    var article = row[GRIEVANCE_COLS.ARTICLES - 1] || '';
    var location = row[GRIEVANCE_COLS.LOCATION - 1] || '';
    var resolution = row[GRIEVANCE_COLS.RESOLUTION - 1] || '';
    var dateResolved = '';

    // Try to get resolution date from Step III received or last updated field
    var step3Rcvd = row[GRIEVANCE_COLS.STEP3_RCVD - 1];
    if (step3Rcvd) {
      dateResolved = Utilities.formatDate(new Date(step3Rcvd), Session.getScriptTimeZone(), 'MMM yyyy');
    }

    // Determine outcome class for styling
    var outcomeClass = 'settled';
    var outcomeDisplay = status;

    if (status.indexOf('won') !== -1 || status.indexOf('sustained') !== -1) {
      outcomeClass = 'won';
      outcomeDisplay = 'Won';
    } else if (status.indexOf('lost') !== -1 || status.indexOf('denied') !== -1 || status.indexOf('dismissed') !== -1) {
      outcomeClass = 'lost';
      outcomeDisplay = 'Denied';
    } else if (status.indexOf('settled') !== -1) {
      outcomeClass = 'settled';
      outcomeDisplay = 'Settled';
    } else if (status.indexOf('withdrawn') !== -1) {
      outcomeClass = 'withdrawn';
      outcomeDisplay = 'Withdrawn';
    }

    // Apply outcome filter
    if (outcomeFilter) {
      if (outcomeFilter === 'won' && outcomeClass !== 'won') continue;
      if (outcomeFilter === 'lost' && outcomeClass !== 'lost') continue;
      if (outcomeFilter === 'settled' && outcomeClass !== 'settled') continue;
      if (outcomeFilter === 'withdrawn' && outcomeClass !== 'withdrawn') continue;
    }

    // Apply search query filter
    if (queryLower) {
      var searchable = [
        grievanceId, firstName, lastName, issueCategory, article, location, resolution, status
      ].join(' ').toLowerCase();

      if (searchable.indexOf(queryLower) === -1) continue;
    }

    results.push({
      id: grievanceId,
      memberName: firstName + ' ' + lastName,
      issueCategory: issueCategory || 'N/A',
      article: article || 'N/A',
      location: location || 'N/A',
      outcome: outcomeDisplay,
      outcomeClass: outcomeClass,
      resolution: resolution,
      dateResolved: dateResolved || 'N/A'
    });
  }

  // Sort by most recent first (by ID, assuming IDs are chronological)
  results.sort(function(a, b) {
    return b.id.localeCompare(a.id);
  });

  // Limit to 50 results
  return results.slice(0, 50);
}



/**
 * ============================================================================
 * STRATEGIC COMMAND CENTER - MEMBER PORTAL SERVICE (v4.5.0)
 * ============================================================================
 * Personalized member portal with Weingarten Rights and steward access.
 *
 * NOTE: The main dashboard functionality has been consolidated into
 * 04_UIService.gs (buildUnifiedDashboard). This file now focuses on:
 * - Personalized member portals (accessed via ?id=MEMBER_ID)
 * - Public portal without member ID
 * - PII safety utilities
 *
 * @fileoverview Member portal service with PII protection
 * @version 4.5.0
 * @requires 01_Constants.gs
 * ============================================================================
 */

// ============================================================================
// SAFETY VALVE - PII AUTO-REDACTION
// ============================================================================

/**
 * Scans and masks PII patterns (Phone numbers, SSNs) from strings.
 * Used to ensure accidental data entry doesn't leak to the Member Dashboard.
 *
 * @param {*} data - Input data to scrub
 * @returns {*} Scrubbed data with PII masked
 */
function safetyValveScrub(data) {
  if (typeof data !== 'string') return data;

  // Mask Phone Numbers: (123) 456-7890 or 123-456-7890 or +1 123-456-7890
  var phoneRegex = /(\+?\d{1,2}\s?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/g;

  // Mask SSN-like patterns: 000-00-0000
  var ssnRegex = /\b\d{3}-\d{2}-\d{4}\b/g;

  // Mask email addresses
  var emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  return data.replace(phoneRegex, "[REDACTED CONTACT]")
             .replace(ssnRegex, "[REDACTED ID]")
             .replace(emailRegex, "[REDACTED EMAIL]");
}

/**
 * Scrubs an object's string values for PII
 * @param {Object} obj - Object to scrub
 * @returns {Object} Object with scrubbed values
 */
function scrubObjectPII(obj) {
  if (!obj || typeof obj !== 'object') return obj;

  var scrubbed = {};
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      scrubbed[key] = safetyValveScrub(obj[key]);
    }
  }
  return scrubbed;
}

// ============================================================================
// DATA FETCHING FUNCTIONS - Portal Support
// ============================================================================

/**
 * Gets grievance statistics for portal display (no PII)
 * @returns {Object} Statistics object with winRate
 */
function getSecureGrievanceStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { total: 0, open: 0, won: 0, pending: 0, resolved: 0, winRate: 0 };
  }

  var data = sheet.getDataRange().getValues();
  var stats = { total: 0, open: 0, won: 0, pending: 0, resolved: 0, denied: 0 };

  for (var i = 1; i < data.length; i++) {
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    if (!status) continue;

    stats.total++;

    switch(status) {
      case 'Open':
        stats.open++;
        break;
      case 'Won':
        stats.won++;
        stats.resolved++;
        break;
      case 'Denied':
        stats.denied++;
        stats.resolved++;
        break;
      case 'Settled':
        stats.won++;
        stats.resolved++;
        break;
      case 'Withdrawn':
        stats.resolved++;
        break;
      default:
        if (status.indexOf('Pending') >= 0) {
          stats.pending++;
        }
    }
  }

  stats.winRate = stats.resolved > 0 ? Math.round((stats.won / stats.resolved) * 100) : 0;
  return stats;
}

/**
 * Gets all stewards for portal display (public info only)
 * @returns {Array} Array of steward objects
 */
function getSecureAllStewards_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var stewards = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      stewards.push({
        'First Name': data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        'Last Name': data[i][MEMBER_COLS.LAST_NAME - 1] || '',
        'Unit': data[i][MEMBER_COLS.UNIT - 1] || 'General',
        'Work Location': data[i][MEMBER_COLS.WORK_LOCATION - 1] || ''
      });
    }
  }

  return stewards;
}

/**
 * Gets satisfaction statistics for portal display
 * @returns {Object} Satisfaction stats with avgTrust
 */
function getSecureSatisfactionStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    avgTrust: 0,
    avgSatisfaction: 0,
    responseCount: 0,
    recentTrend: 'stable'
  };

  if (!sheet || sheet.getLastRow() < 2) return result;

  try {
    var data = sheet.getDataRange().getValues();
    var trustScores = [];
    var satScores = [];

    for (var i = 1; i < data.length; i++) {
      var trustVal = parseFloat(data[i][SATISFACTION_COLS.Q7_TRUST_UNION - 1]);
      var satVal = parseFloat(data[i][SATISFACTION_COLS.Q6_SATISFIED_REP - 1]);

      if (!isNaN(trustVal) && trustVal >= 1 && trustVal <= 10) {
        trustScores.push(trustVal);
      }
      if (!isNaN(satVal) && satVal >= 1 && satVal <= 10) {
        satScores.push(satVal);
      }
    }

    if (trustScores.length > 0) {
      result.avgTrust = Math.round((trustScores.reduce(function(a,b){return a+b;}, 0) / trustScores.length) * 10) / 10;
    }
    if (satScores.length > 0) {
      result.avgSatisfaction = Math.round((satScores.reduce(function(a,b){return a+b;}, 0) / satScores.length) * 10) / 10;
    }
    result.responseCount = Math.max(trustScores.length, satScores.length);

    // Trend analysis (compare last 10 vs previous 10)
    if (trustScores.length >= 20) {
      var recent = trustScores.slice(-10).reduce(function(a,b){return a+b;}, 0) / 10;
      var previous = trustScores.slice(-20, -10).reduce(function(a,b){return a+b;}, 0) / 10;
      result.recentTrend = recent > previous + 0.2 ? 'improving' : (recent < previous - 0.2 ? 'declining' : 'stable');
    }
  } catch (e) {
    Logger.log('Error in getSecureSatisfactionStats_: ' + e.message);
  }

  return result;
}

/**
 * Gets steward workload data
 * @returns {Array} Array of workload objects
 */
function getStewardWorkload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  var workload = {};

  // Get all stewards first
  var memberData = memberSheet.getDataRange().getValues();
  for (var i = 1; i < memberData.length; i++) {
    if (memberData[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      var name = (memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                 (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      workload[name.trim()] = { name: name.trim(), openCases: 0, totalCases: 0 };
    }
  }

  // Count grievances per steward
  var gData = grievanceSheet.getDataRange().getValues();
  for (var g = 1; g < gData.length; g++) {
    var steward = gData[g][GRIEVANCE_COLS.STEWARD - 1];
    var status = gData[g][GRIEVANCE_COLS.STATUS - 1];

    if (steward && workload[steward]) {
      workload[steward].totalCases++;
      if (status === 'Open' || (status && status.indexOf('Pending') >= 0)) {
        workload[steward].openCases++;
      }
    }
  }

  return Object.keys(workload).map(function(key) { return workload[key]; });
}

/**
 * Gets contract PDF URL from config
 * @returns {string} Contract PDF URL
 */
function getContractPdfUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var url = configSheet.getRange(3, CONFIG_COLS.ORG_WEBSITE).getValue();
      if (url) return url;
    }
  } catch (_e) { /* Config sheet may not exist yet; fall back to '#' */ }
  return '#';
}

/**
 * Gets resource Drive folder URL from config
 * @returns {string} Resource folder URL
 */
function getResourceDriveUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var folderId = configSheet.getRange(3, CONFIG_COLS.ARCHIVE_FOLDER_ID).getValue();
      if (folderId) return 'https://drive.google.com/drive/folders/' + folderId;
    }
  } catch (_e) { /* Config sheet may not exist yet; fall back to '#' */ }
  return '#';
}

// ============================================================================
// PORTAL MENU WRAPPERS
// ============================================================================

/**
 * Menu wrapper: Build portal for the currently selected member
 * Gets the member ID from the active row in Member Directory
 */
function buildPortalForSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Build Member Portal',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Build Member Portal',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();

  if (!memberId) {
    ui.alert('Build Member Portal',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  var portal = buildMemberPortal(memberId);
  ui.showModalDialog(portal, 'Member Portal');
}

/**
 * Menu wrapper: Send portal email to the currently selected member
 */
function sendPortalEmailToSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Send Portal Email',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Send Portal Email',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();
  var memberEmail = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var _firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();

  if (!memberId) {
    ui.alert('Send Portal Email',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  if (!memberEmail) {
    ui.alert('Send Portal Email',
      'No email address found for this member.',
      ui.ButtonSet.OK);
    return;
  }

  sendMemberDashboardEmail(memberId);
}

// ============================================================================
// PORTAL BUILDERS
// ============================================================================

/**
 * Builds personalized member portal for specific member ID
 * @param {string} memberId - The member ID to look up
 * @returns {HtmlOutput} Personalized portal HTML
 */
function buildMemberPortal(memberId) {
  var profile = getMemberProfile(memberId);

  if (!profile) {
    return HtmlService.createHtmlOutput(getErrorPageHtml_('Member not found'))
      .setTitle('Member Portal - Error');
  }

  var html = getMemberPortalHtml_(profile);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Member Portal - ' + profile.firstName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Builds public portal (no member ID required)
 * @returns {HtmlOutput} Public portal HTML
 */
function buildPublicPortal() {
  var stats = getSecureGrievanceStats_();
  var stewards = getSecureAllStewards_();
  var satisfaction = getSecureSatisfactionStats_();

  var html = getPublicPortalHtml_(stats, stewards, satisfaction);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Union Member Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Gets member profile by ID (with PII scrubbing for sensitive fields)
 * @param {string} memberId - The member ID to look up
 * @returns {Object|null} Member profile object or null if not found
 */
function getMemberProfile(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || memberSheet.getLastRow() < 2) return null;

  var data = memberSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      return {
        memberId: memberId,
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: data[i][MEMBER_COLS.LAST_NAME - 1] || '',
        unit: data[i][MEMBER_COLS.UNIT - 1] || 'General',
        workLocation: data[i][MEMBER_COLS.WORK_LOCATION - 1] || '',
        duesPaying: true,
        isSteward: data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
        volunteerHours: parseFloat(data[i][MEMBER_COLS.VOLUNTEER_HOURS - 1]) || 0
      };
    }
  }

  return null;
}

/**
 * Sends member dashboard email with personalized portal link
 * @param {string} memberId - The member ID to send link to
 */
function sendMemberDashboardEmail(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found.');
    return;
  }

  var data = memberSheet.getDataRange().getValues();
  var member = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      member = {
        email: data[i][MEMBER_COLS.EMAIL - 1],
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1]
      };
      break;
    }
  }

  if (!member || !member.email) {
    SpreadsheetApp.getUi().alert('Member email not found.');
    return;
  }

  var webAppUrl = ScriptApp.getService().getUrl();
  var portalUrl = webAppUrl + '?id=' + memberId;

  var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Your Personal Union Portal';
  var body = 'Hello ' + member.firstName + ',\n\n' +
    'Access your personalized Union Member Portal:\n\n' +
    portalUrl + '\n\n' +
    'Your portal includes:\n' +
    '- Emergency Weingarten Rights card\n' +
    '- Your steward contact information\n' +
    '- Union resources and contract links\n' +
    '- Satisfaction survey link\n\n' +
    'Keep this link private - it is personalized for you.\n\n' +
    'In Solidarity,\n' +
    'Strategic Command Center';

  try {
    MailApp.sendEmail(member.email, subject, body);
    SpreadsheetApp.getUi().alert('Portal link sent to ' + member.email);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + e.message);
  }
}

// ============================================================================
// PORTAL HTML GENERATORS
// ============================================================================

/**
 * Generates personalized member portal HTML
 * @param {Object} profile - Member profile object
 * @returns {string} Complete HTML for the portal
 * @private
 */
function getMemberPortalHtml_(profile) {
  var stewards = getSecureAllStewards_();
  var stats = getSecureGrievanceStats_();
  var CONTRACT_PDF_URL = getContractPdfUrl_();
  var RESOURCE_DRIVE_URL = getResourceDriveUrl_();

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; }' +
    '    .container { max-width: 600px; margin: 0 auto; padding: 20px; }' +
    '    .header { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); padding: 24px; border-radius: 16px; margin-bottom: 20px; text-align: center; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; margin-bottom: 8px; }' +
    '    .header .welcome { font-size: 24px; font-weight: 700; }' +
    '    .member-badge { display: inline-flex; align-items: center; gap: 6px; background: rgba(255,255,255,0.2); padding: 6px 12px; border-radius: 20px; font-size: 12px; margin-top: 12px; }' +
    '    .rights-box { background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%); border: 2px solid #dc2626; border-radius: 12px; padding: 20px; margin-bottom: 20px; }' +
    '    .rights-header { display: flex; align-items: center; gap: 8px; font-weight: 700; color: #fecaca; margin-bottom: 12px; font-size: 14px; }' +
    '    .rights-text { font-size: 15px; line-height: 1.6; font-weight: 500; }' +
    '    .card { background: rgba(255,255,255,0.05); border-radius: 12px; padding: 16px; margin-bottom: 12px; border: 1px solid rgba(255,255,255,0.1); }' +
    '    .card-title { display: flex; align-items: center; gap: 8px; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; margin-bottom: 12px; }' +
    '    .btn-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }' +
    '    .btn { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 14px; border-radius: 10px; text-decoration: none; font-weight: 600; font-size: 13px; transition: all 0.2s; }' +
    '    .btn-purple { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; }' +
    '    .btn-green { background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; }' +
    '    .btn:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.3); }' +
    '    .steward-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.1); }' +
    '    .steward-item:last-child { border-bottom: none; }' +
    '    .pill { background: rgba(124,58,237,0.2); color: #A78BFA; padding: 4px 10px; border-radius: 20px; font-size: 11px; }' +
    '    .stats-row { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }' +
    '    .stat { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .stat-value { font-size: 24px; font-weight: 700; color: #60A5FA; }' +
    '    .stat-label { font-size: 10px; color: #94A3B8; text-transform: uppercase; }' +
    '    .footer { text-align: center; font-size: 10px; color: #64748B; padding: 20px 0; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="header">' +
    '      <h1>MEMBER PORTAL</h1>' +
    '      <div class="welcome">Welcome, ' + escapeHtml(profile.firstName) + '!</div>' +
    '      <div class="member-badge">' +
    '        <i class="material-icons" style="font-size:16px">' + (profile.isSteward ? 'verified' : 'person') + '</i>' +
    '        ' + (profile.isSteward ? 'Union Steward' : 'Union Member') + ' - ' + escapeHtml(profile.unit) +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="rights-box">' +
    '      <div class="rights-header"><i class="material-icons">gavel</i> WEINGARTEN RIGHTS</div>' +
    '      <div class="rights-text">"If this discussion could in any way lead to my being disciplined or terminated, I respectfully request that my union representative be present. Until my representative arrives, I choose not to answer questions or write any statements."</div>' +
    '    </div>' +
    '    ' +
    '    <div class="btn-grid">' +
    '      <a href="' + (/^https:\/\//.test(CONTRACT_PDF_URL) ? escapeHtml(CONTRACT_PDF_URL) : '#') + '" target="_blank" class="btn btn-purple"><i class="material-icons">description</i> Contract</a>' +
    '      <a href="' + (/^https:\/\//.test(RESOURCE_DRIVE_URL) ? escapeHtml(RESOURCE_DRIVE_URL) : '#') + '" target="_blank" class="btn btn-green"><i class="material-icons">folder</i> Resources</a>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">people</i> Your Stewards</div>' +
    stewards.slice(0, 5).map(function(s) {
      return '<div class="steward-item"><span>' + escapeHtml(s['First Name']) + ' ' + escapeHtml(s['Last Name']) + '</span><span class="pill">' + escapeHtml(s['Unit']) + '</span></div>';
    }).join('') +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">bar_chart</i> Union Stats</div>' +
    '      <div class="stats-row">' +
    '        <div class="stat"><div class="stat-value">' + stats.winRate + '%</div><div class="stat-label">Win Rate</div></div>' +
    '        <div class="stat"><div class="stat-value">' + stats.open + '</div><div class="stat-label">Active Cases</div></div>' +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="footer">' +
    '      <i class="material-icons" style="font-size:12px;vertical-align:middle">lock</i> ' +
    '      Secure Member Portal | Member ID: ' + escapeHtml(profile.memberId) +
    '    </div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}

/**
 * Generates public portal HTML (no member ID required)
 * @param {Object} stats - Grievance statistics
 * @param {Array} stewards - List of stewards
 * @param {Object} satisfaction - Satisfaction stats
 * @returns {string} Complete HTML for public portal
 * @private
 */
function getPublicPortalHtml_(stats, stewards, satisfaction) {
  var CONTRACT_PDF_URL = getContractPdfUrl_();
  var RESOURCE_DRIVE_URL = getResourceDriveUrl_();

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; }' +
    '    .container { max-width: 600px; margin: 0 auto; padding: 20px; }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 2px solid rgba(124,58,237,0.3); }' +
    '    .header h1 { font-size: 20px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 28px; }' +
    '    .rights-box { background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%); border: 2px solid #dc2626; border-radius: 12px; padding: 20px; margin-bottom: 20px; }' +
    '    .rights-header { display: flex; align-items: center; gap: 8px; font-weight: 700; color: #fecaca; margin-bottom: 12px; }' +
    '    .rights-text { font-size: 14px; line-height: 1.6; font-weight: 500; }' +
    '    .btn-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }' +
    '    .btn { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 14px; border-radius: 10px; text-decoration: none; font-weight: 600; font-size: 13px; }' +
    '    .btn-purple { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; }' +
    '    .btn-green { background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; }' +
    '    .card { background: rgba(255,255,255,0.05); border-radius: 12px; padding: 16px; margin-bottom: 12px; border: 1px solid rgba(255,255,255,0.1); }' +
    '    .card-title { display: flex; align-items: center; gap: 8px; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; margin-bottom: 12px; }' +
    '    .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }' +
    '    .stat { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .stat-value { font-size: 24px; font-weight: 700; }' +
    '    .stat-value.green { color: #10B981; }' +
    '    .stat-value.blue { color: #60A5FA; }' +
    '    .stat-label { font-size: 10px; color: #94A3B8; text-transform: uppercase; }' +
    '    .steward-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.1); }' +
    '    .steward-item:last-child { border-bottom: none; }' +
    '    .pill { background: rgba(124,58,237,0.2); color: #A78BFA; padding: 4px 10px; border-radius: 20px; font-size: 11px; }' +
    '    .footer { text-align: center; font-size: 10px; color: #64748B; padding: 20px 0; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="header"><i class="material-icons">shield</i><h1>Union Member Portal</h1></div>' +
    '    ' +
    '    <div class="rights-box">' +
    '      <div class="rights-header"><i class="material-icons">gavel</i> WEINGARTEN RIGHTS</div>' +
    '      <div class="rights-text">"If this discussion could in any way lead to my being disciplined or terminated, I respectfully request that my union representative be present."</div>' +
    '    </div>' +
    '    ' +
    '    <div class="btn-grid">' +
    '      <a href="' + (/^https:\/\//.test(CONTRACT_PDF_URL) ? escapeHtml(CONTRACT_PDF_URL) : '#') + '" target="_blank" class="btn btn-purple"><i class="material-icons">description</i> Contract</a>' +
    '      <a href="' + (/^https:\/\//.test(RESOURCE_DRIVE_URL) ? escapeHtml(RESOURCE_DRIVE_URL) : '#') + '" target="_blank" class="btn btn-green"><i class="material-icons">folder</i> Resources</a>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">bar_chart</i> Union Statistics</div>' +
    '      <div class="stats-grid">' +
    '        <div class="stat"><div class="stat-value green">' + satisfaction.avgTrust + '/10</div><div class="stat-label">Trust Score</div></div>' +
    '        <div class="stat"><div class="stat-value blue">' + stats.winRate + '%</div><div class="stat-label">Win Rate</div></div>' +
    '        <div class="stat"><div class="stat-value blue">' + stats.open + '</div><div class="stat-label">Active Cases</div></div>' +
    '        <div class="stat"><div class="stat-value green">' + stats.total + '</div><div class="stat-label">Total Cases</div></div>' +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">people</i> Find a Steward</div>' +
    stewards.map(function(s) {
      return '<div class="steward-item"><span>' + escapeHtml(s['First Name']) + ' ' + escapeHtml(s['Last Name']) + '</span><span class="pill">' + escapeHtml(s['Unit']) + '</span></div>';
    }).join('') +
    '    </div>' +
    '    ' +
    '    <div class="footer"><i class="material-icons" style="font-size:12px;vertical-align:middle">lock</i> No PII Visible | Public Portal v' + VERSION_INFO.CURRENT + ' (' + VERSION_INFO.BUILD_DATE + ')' + '</div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}

/**
 * Generates error page HTML
 * @param {string} message - Error message to display
 * @returns {string} Error page HTML
 * @private
 */
function getErrorPageHtml_(message) {
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    body { font-family: "Roboto", sans-serif; background: #0F172A; min-height: 100vh; display: flex; align-items: center; justify-content: center; color: #F8FAFC; }' +
    '    .error-box { text-align: center; padding: 40px; background: rgba(239,68,68,0.1); border: 1px solid rgba(239,68,68,0.3); border-radius: 16px; max-width: 400px; }' +
    '    .error-icon { font-size: 48px; color: #ef4444; margin-bottom: 16px; }' +
    '    .error-title { font-size: 20px; font-weight: 700; margin-bottom: 8px; }' +
    '    .error-msg { color: #94A3B8; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="error-box">' +
    '    <i class="material-icons error-icon">error_outline</i>' +
    '    <div class="error-title">Access Error</div>' +
    '    <div class="error-msg">' + escapeHtml(message) + '</div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}
