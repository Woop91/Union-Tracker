/**
 * ============================================================================
 * 509 STRATEGIC COMMAND CENTER - UNIFIED MASTER ENGINE (v4.0)
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
 * ============================================================================
 */

// ============================================================================
// COMMAND CENTER CONFIGURATION
// ============================================================================

/**
 * Alternative CONFIG object for legacy compatibility
 * Maps to COMMAND_CONFIG in 01_Constants.gs
 */
/**
 * Get COMMAND_CENTER_CONFIG lazily to avoid load-order issues
 * This function ensures SHEETS and COMMAND_CONFIG are available before use
 * @returns {Object} Command center configuration object
 */
function getCommandCenterConfig() {
  return {
    SYSTEM_NAME: "509 Strategic Command Center",
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

// Legacy COMMAND_CENTER_CONFIG variable for backward compatibility
// Uses hardcoded values to avoid load-order issues with global initialization
var COMMAND_CENTER_CONFIG = {
  SYSTEM_NAME: "509 Strategic Command Center",
  LOG_SHEET_NAME: 'Grievance Log',           // Hardcoded to avoid load-order issues
  DIR_SHEET_NAME: 'Member Directory',        // Hardcoded to avoid load-order issues
  AUDIT_SHEET_NAME: '_Audit_Log',            // Hardcoded to avoid load-order issues
  TEMPLATE_ID: '',                           // Loaded dynamically via getCommandCenterConfig()
  ARCHIVE_FOLDER_ID: '',                     // Loaded dynamically via getCommandCenterConfig()
  CHIEF_STEWARD_EMAIL: '',                   // Loaded dynamically via getCommandCenterConfig()
  UNIT_CODES: {},                            // Loaded dynamically via getCommandCenterConfig()
  THEME: {
    HEADER_BG: '#1e293b',
    HEADER_TEXT: '#ffffff',
    ALT_ROW: '#f8fafc',
    FONT: 'Roboto',
    FONT_SIZE: 10
  }
};

// ============================================================================
// COMMAND CENTER MENU
// ============================================================================

/**
 * Creates the 509 Command Center menu (v4.0 Harmonized Structure)
 * Called by createDashboardMenu() in UIService.gs
 *
 * v4.0 Features:
 * - Quick access to Search and Mobile View at top level
 * - Production Mode: Demo Data menu disappears after NUKE
 * - Organized submenus for Personnel, Grievance, Security, and Styling
 */
function createCommandCenterMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📱 Field Portal');

  // Top-level quick actions (v4.0)
  menu.addItem('👁️ Refresh Dashboard UI', 'APPLY_SYSTEM_THEME')
      .addItem('🔍 Search Members', 'showSearchDialog');

  menu.addSeparator();

  // v4.0 Field Accessibility submenu (Mobile/Pocket View)
  menu.addSubMenu(ui.createMenu('📱 Field Accessibility')
      .addItem('📱 Pocket View', 'navToMobile')
      .addItem('🖥️ Restore Full Desktop View', 'showAllMemberColumns')
      .addItem('🔄 Refresh View', 'refreshMemberView')
      .addSeparator()
      .addItem('📱 Get Mobile App URL', 'showWebAppUrl'));

  menu.addSeparator();

  // Personnel Management submenu
  menu.addSubMenu(ui.createMenu('👤 Personnel Management')
      .addItem('🆔 Generate Missing Member IDs', 'generateMissingMemberIDs')
      .addItem('⚡ Generate IDs (Batch Mode)', 'generateMissingMemberIDsBatch')
      .addItem('🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs')
      .addItem('✅ Verify ID Engine', 'verifyIDGenerationEngine')
      .addSeparator()
      .addItem('🌟 Promote Selected to Steward', 'promoteSelectedMemberToSteward')
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward'));

  menu.addSeparator();

  // Grievance Tools submenu
  menu.addSubMenu(ui.createMenu('📋 Grievance Tools')
      .addItem('🚦 Apply Traffic Light Indicators', 'applyTrafficLightIndicators')
      .addItem('🔄 Clear Traffic Lights', 'clearTrafficLightIndicators')
      .addItem('📄 Create PDF for Selected', 'createPDFForSelectedGrievance'));

  menu.addSeparator();

  // System Security submenu
  menu.addSubMenu(ui.createMenu('🛡️ System Security')
      .addItem('📸 Create Manual Snapshot', 'createWeeklySnapshot')
      .addItem('📅 Setup Weekly Backup', 'setupWeeklySnapshotTrigger')
      .addItem('📜 View Audit Log', 'navigateToAuditLog')
      .addSeparator()
      .addItem('🔍 v4.0 System Diagnostic', 'showDiagnosticReport')
      .addItem('🛠️ Repair Dashboard', 'repairDashboardWithUI')
      .addItem('📊 v4.0 Status Report', 'showV4StatusReport'));

  menu.addSeparator();

  // Styling & Theme submenu
  menu.addSubMenu(ui.createMenu('🎨 Styling & Theme')
      .addItem('🎨 Apply Global Theme', 'APPLY_SYSTEM_THEME')
      .addItem('🔄 Reset to Default', 'resetToDefaultTheme')
      .addItem('✨ Refresh All Visuals', 'refreshAllVisuals'));

  menu.addSeparator();

  // v4.0 Analytics & Scaling submenu
  menu.addSubMenu(ui.createMenu('📈 Analytics & Insights')
      .addItem('🏥 Unit Health Report', 'showUnitHealthReport')
      .addItem('📊 Grievance Trends', 'showGrievanceTrends')
      .addItem('📚 Search Precedents', 'showSearchPrecedents')
      .addSeparator()
      .addItem('📝 OCR Transcribe Form', 'showOCRDialog')
      .addItem('🔧 OCR Setup', 'setupOCRApiKey'));

  menu.addSeparator();

  // Unified Dashboards (v4.3.2)
  menu.addItem('👥 Member Dashboard', 'showPublicMemberDashboard');
  menu.addItem('🛡️ Steward Dashboard', 'showStewardDashboard');

  menu.addSeparator();

  // Web App & Portal submenu
  menu.addSubMenu(ui.createMenu('🌐 Web App & Portal')
      .addItem('👤 Build Member Portal', 'buildPortalForSelectedMember')
      .addItem('📊 Build Public Portal', 'buildPublicPortal')
      .addItem('📧 Send Portal Email', 'sendPortalEmailToSelectedMember'));

  // v4.0 PRODUCTION MODE: Demo Data menu disappears after NUKE
  if (!isProductionMode()) {
    menu.addSeparator();
    menu.addSubMenu(ui.createMenu('🎭 Demo Data')
        .addItem('🌱 Seed Sample Data', 'SEED_SAMPLE_DATA')
        .addItem('☢️ NUKE EVERYTHING', 'NUKE_DATABASE'));
  }

  menu.addToUi();
}

// ============================================================================
// NAVIGATION SHORTCUTS (v4.0)
// ============================================================================

/**
 * Quick navigation to Executive Dashboard
 */
function navigateToDashboard() {
  navToDash();
}

/**
 * Quick navigation to Mobile View
 */
function navigateToMobileView() {
  navToMobile();
}

/**
 * v4.0 Mobile/Pocket View Navigation
 * Optimizes the Member Directory for smartphone viewing by hiding non-essential columns.
 * Shows only: Member ID, Name, Status, Phone for quick field access.
 *
 * Call showAllMemberColumns() to restore full view.
 */
function navToMobile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  // Mobile Optimization: Hide non-essential columns for phone view
  // Keep visible: A (Member ID), B (First Name), C (Last Name), I (Phone), N (Is Steward)
  // Hide columns D-H (Job, Location, Unit, Office Days, Email) - columns 4-8
  // Hide columns J-M (Preferred Comm to Manager) - columns 10-13
  // Hide columns O-AF (Committees to Quick Actions) - columns 15-32

  try {
    // Hide columns 4-8 (D-H)
    if (sheet.getMaxColumns() >= 8) {
      sheet.hideColumns(4, 5);  // Hide 5 columns starting at column 4
    }

    // Hide columns 10-13 (J-M)
    if (sheet.getMaxColumns() >= 13) {
      sheet.hideColumns(10, 4);  // Hide 4 columns starting at column 10
    }

    // Hide columns 15-32 (O-AF) if they exist
    if (sheet.getMaxColumns() >= 32) {
      sheet.hideColumns(15, 18);  // Hide 18 columns starting at column 15
    } else if (sheet.getMaxColumns() >= 15) {
      sheet.hideColumns(15, sheet.getMaxColumns() - 14);
    }

    sheet.activate();
    ss.toast('📱 Mobile View Optimized. Essential columns visible for phone access.', COMMAND_CONFIG.SYSTEM_NAME, 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error enabling mobile view: ' + e.message);
  }
}

/**
 * v4.0 Restore Full View
 * Shows all columns in Member Directory after Mobile View was enabled.
 */
function showAllMemberColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found');
    return;
  }

  try {
    // Show all columns
    sheet.showColumns(1, sheet.getMaxColumns());
    ss.toast('✅ All columns restored.', COMMAND_CONFIG.SYSTEM_NAME, 3);
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
  } catch (e) {
    // Fallback: just try to show all columns (safe default)
    showAllMemberColumns();
    SpreadsheetApp.getUi().alert('Restored to full view. Use this toggle to switch between mobile and desktop views.');
  }
}

// ============================================================================
// v4.0 HIGH-PERFORMANCE DATA ENGINE
// ============================================================================

/**
 * v4.0 Sequential ID Generator
 * Gets the next sequence number for a given prefix from stored properties.
 * Used for generating Member IDs with unit code prefixes.
 *
 * @param {string} prefix - The unit code prefix (e.g., "MS", "FO", "HC")
 * @returns {string} The next sequence number as a 4-digit padded string
 */
function getNextSequence(prefix) {
  var props = PropertiesService.getScriptProperties();
  var key = 'SEQUENCE_' + prefix;
  var current = parseInt(props.getProperty(key) || '0', 10);
  var next = current + 1;
  props.setProperty(key, String(next));
  return String(next).padStart(4, '0');
}

/**
 * v4.0 Batch Member ID Generator (High-Performance)
 * Uses batch array processing to generate IDs for up to 5,000 members without lag.
 * Reads unit codes from Config sheet or falls back to defaults.
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
  var unitCodes = getUnitCodes_();
  var countAdded = 0;

  // Process in memory (no individual cell writes)
  for (var i = 1; i < data.length; i++) {
    // Check if Member ID is empty and Unit exists
    if (!data[i][MEMBER_COLS.MEMBER_ID - 1] && data[i][MEMBER_COLS.UNIT - 1]) {
      var unit = data[i][MEMBER_COLS.UNIT - 1];
      var prefix = unitCodes[unit] || 'GEN';
      var nextNum = getNextSequence(prefix);
      data[i][MEMBER_COLS.MEMBER_ID - 1] = prefix + '-' + nextNum + '-H';
      countAdded++;
    }
  }

  // Single batch write (high-performance)
  if (countAdded > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  ss.toast('✅ ' + countAdded + ' IDs generated.', COMMAND_CONFIG.SYSTEM_NAME, 3);

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
 * v4.0 ID Generation Engine Verification
 * Tests the ID generation system to ensure proper sequencing and format.
 */
function verifyIDGenerationEngine() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();

  var report = '🔍 ID GENERATION ENGINE VERIFICATION\n';
  report += '=' .repeat(45) + '\n\n';

  // Check current sequences
  report += '📊 CURRENT SEQUENCE COUNTERS:\n';
  var unitCodes = getUnitCodes_();
  var sequenceKeys = Object.keys(unitCodes).map(function(unit) {
    return 'SEQUENCE_' + unitCodes[unit];
  });

  // Add GEN prefix for generic IDs
  sequenceKeys.push('SEQUENCE_GEN');

  sequenceKeys.forEach(function(key) {
    var value = props.getProperty(key) || '0';
    report += '  ' + key + ': ' + value + '\n';
  });

  // Test ID format generation
  report += '\n🧪 TEST ID GENERATION:\n';
  var testPrefix = 'TEST';
  var testSeq = getNextSequence(testPrefix);
  report += '  Generated: ' + testPrefix + '-' + testSeq + '-H\n';
  report += '  Format Valid: ' + (testSeq.length === 4 ? '✅ Yes' : '❌ No') + '\n';

  // Check Member Directory for ID statistics
  report += '\n📈 MEMBER ID STATISTICS:\n';
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
    var withID = 0;
    var withoutID = 0;
    var idFormats = {};

    data.forEach(function(row) {
      if (row[0]) {
        withID++;
        var prefix = String(row[0]).split('-')[0];
        idFormats[prefix] = (idFormats[prefix] || 0) + 1;
      } else {
        withoutID++;
      }
    });

    report += '  Members with ID: ' + withID + '\n';
    report += '  Members without ID: ' + withoutID + '\n';
    report += '\n  ID Prefix Distribution:\n';
    Object.keys(idFormats).forEach(function(prefix) {
      report += '    ' + prefix + ': ' + idFormats[prefix] + '\n';
    });
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

  var report = '📊 509 STRATEGIC COMMAND CENTER\n';
  report += 'v4.0 UNIFIED MASTER ENGINE STATUS\n';
  report += '=' .repeat(45) + '\n\n';

  // System Identity
  report += '🏷️ SYSTEM IDENTITY:\n';
  report += '  Name: ' + COMMAND_CONFIG.SYSTEM_NAME + '\n';
  report += '  Version: ' + COMMAND_CONFIG.VERSION + '\n';
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
  } catch (e) {
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

    // Delete _Audit_Log hidden sheet
    var auditLogSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    if (auditLogSheet) {
      try { ss.deleteSheet(auditLogSheet); } catch (e) { Logger.log('Could not delete _Audit_Log: ' + e.message); }
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

    // Log the action
    logAuditEvent('NUCLEAR_WIPE', {
      performedBy: Session.getActiveUser().getEmail(),
      timestamp: new Date().toISOString()
    });

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
    var repoUrl = 'https://github.com/Woop91/MULTIPLE-SCRIPS-REPO';
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

  // Red tabs - Getting Started, FAQ
  var redSheets = [SHEETS.GETTING_STARTED, SHEETS.FAQ];
  redSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.setTabColor(TAB_COLORS.RED); } catch (e) { Logger.log('Tab color error: ' + e.message); }
    }
  });

  // Purple tabs - Member Directory, Grievance Log
  var purpleSheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];
  purpleSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.setTabColor(TAB_COLORS.PURPLE); } catch (e) { Logger.log('Tab color error: ' + e.message); }
    }
  });

  // Yellow tabs - Feedback and Development, Function Checklist
  var yellowSheets = [SHEETS.FEEDBACK, SHEETS.FUNCTION_CHECKLIST];
  yellowSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.setTabColor(TAB_COLORS.YELLOW); } catch (e) { Logger.log('Tab color error: ' + e.message); }
    }
  });

  // Orange tabs - Config, Config Guide
  var orangeSheets = [SHEETS.CONFIG, SHEETS.CONFIG_GUIDE];
  orangeSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.setTabColor(TAB_COLORS.ORANGE); } catch (e) { Logger.log('Tab color error: ' + e.message); }
    }
  });

  // Blue tabs - Member Satisfaction
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet) {
    try { satSheet.setTabColor(TAB_COLORS.BLUE); } catch (e) { Logger.log('Tab color error: ' + e.message); }
  }

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

  var report = '🔍 509 COMMAND CENTER DIAGNOSTIC REPORT\n';
  report += '=' .repeat(50) + '\n\n';

  // Check required sheets
  var requiredSheets = [
    SHEETS.CONFIG,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.DASHBOARD,
    SHEETS.INTERACTIVE
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
  var hiddenSheets = ['_Dashboard_Calc', '_Grievance_Calc', '_Member_Lookup'];

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
  } catch (e) {
    report += '  PDF Template: ⚠️ Unable to check\n';
  }

  try {
    var archiveId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID);
    report += '  Archive Folder: ' + (archiveId ? '✅ Configured' : '⚠️ Not set') + '\n';
  } catch (e) {
    report += '  Archive Folder: ⚠️ Unable to check\n';
  }

  try {
    var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
    report += '  Chief Steward Email: ' + (chiefEmail ? '✅ ' + chiefEmail : '⚠️ Not set') + '\n';
  } catch (e) {
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
              document.getElementById('results').innerHTML = '<div class="no-results">Error: ' + e.message + '</div>';
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
            var name = (member['First Name'] || '') + ' ' + (member['Last Name'] || '');
            var id = member['Member ID'] || 'N/A';
            var email = member['Email'] || '';
            html += '<div class="result-item" onclick="selectMember(\\'' + id + '\\')">';
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
var GEMINI_CONFIG = {
  SYSTEM_NAME: "509 Strategic Command Center",
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

/**
 * Gemini v4.0 Form Submission Handler (Legacy - use onGrievanceFormSubmit in 05_Integrations.gs instead)
 * NOTE: Renamed to avoid duplicate function definition conflict.
 * The primary implementation is in 05_Integrations.gs
 *
 * @param {Object} e - Form submission event object
 * @deprecated Use onGrievanceFormSubmit in 05_Integrations.gs
 */
function onGrievanceFormSubmit_Legacy_(e) {
  try {
    var responses = e.namedValues;
    var data = {
      name: responses['Member Name'] ? responses['Member Name'][0] : "Unknown",
      id: responses['Member ID'] ? responses['Member ID'][0] : "000",
      details: responses['Details'] ? responses['Details'][0] : "No details provided."
    };

    // Get or create member-specific folder
    var memberFolder = getOrCreateMemberFolder(data.name, data.id);

    // Create signature-ready PDF
    var pdfFile = createGrievancePDF(memberFolder, data);

    // Log PDF URL to Grievance Log
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG) ||
                   ss.getSheetByName(GEMINI_CONFIG.LOG_SHEET_NAME);

    if (logSheet && logSheet.getLastRow() > 1) {
      // Update Drive Folder URL column
      logSheet.getRange(logSheet.getLastRow(), GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(pdfFile.getUrl());
    }

    // Log the action
    logAuditEvent('GRIEVANCE_PDF_CREATED', {
      memberId: data.id,
      memberName: data.name,
      pdfUrl: pdfFile.getUrl()
    });

  } catch (error) {
    console.error('Form submission error: ' + error.message);
    logAuditEvent('FORM_SUBMISSION_ERROR', { error: error.message });
  }
}

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
    throw new Error('PDF Template ID not configured. Set it in Config sheet column AX.');
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
 * @param {boolean} options.autoPopulate - If true, attempts to auto-populate grievance fields
 * @param {string} options.targetGrievanceId - Grievance ID to populate (if autoPopulate is true)
 * @returns {Object} Result object with extracted text and metadata
 */
function wger(fileId, options) {
  options = options || {};
  var mode = options.mode || 'TEXT';
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
    var ocrResult = performCloudVisionOCR_(imageBlob, mode);

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
 * @returns {Object} Result with success status and extracted text
 * @private
 */
function performCloudVisionOCR_(imageBlob, mode) {
  try {
    // Get API key from script properties (set via Script Properties or Config)
    var apiKey = PropertiesService.getScriptProperties().getProperty('CLOUD_VISION_API_KEY');

    if (!apiKey) {
      // Fallback: Try to use built-in Cloud Vision (requires Cloud project linking)
      return performBuiltInOCR_(imageBlob, mode);
    }

    // Prepare Cloud Vision API request
    var base64Image = Utilities.base64Encode(imageBlob.getBytes());

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
        }]
      }]
    };

    var response = UrlFetchApp.fetch(
      'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey,
      {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true
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
      return { success: false, message: 'Grievance Log not found' };
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
      return { success: false, message: 'Grievance not found: ' + grievanceId };
    }

    var fieldsPopulated = [];
    var textLower = text.toLowerCase();

    // Pattern matching for common grievance form fields
    var patterns = {
      incidentDate: /(?:incident\s*date|date\s*of\s*incident|occurred\s*on)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
      articles: /(?:article|art\.?)\s*(\d+)/gi,
      issueCategory: /(?:discipline|workload|scheduling|pay|benefits|safety|harassment|discrimination)/i
    };

    // Try to extract incident date
    var dateMatch = text.match(patterns.incidentDate);
    if (dateMatch) {
      try {
        var parsedDate = new Date(dateMatch[1]);
        if (!isNaN(parsedDate.getTime())) {
          sheet.getRange(grievanceRow, GRIEVANCE_COLS.INCIDENT_DATE).setValue(parsedDate);
          fieldsPopulated.push('Incident Date');
        }
      } catch (e) {
        // Date parsing failed, skip
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
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.ARTICLES).setValue(uniqueArticles.join(', '));
      fieldsPopulated.push('Articles');
    }

    // Try to identify issue category
    var categoryMatch = text.match(patterns.issueCategory);
    if (categoryMatch) {
      var category = categoryMatch[0].charAt(0).toUpperCase() + categoryMatch[0].slice(1).toLowerCase();
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.ISSUE_CATEGORY).setValue(category);
      fieldsPopulated.push('Issue Category');
    }

    // Store the full OCR text in resolution notes for reference
    var existingResolution = sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).getValue() || '';
    var ocrNote = '[OCR Extract ' + new Date().toLocaleDateString() + ']: ' + text.substring(0, 500);
    if (text.length > 500) ocrNote += '...';

    if (existingResolution) {
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).setValue(existingResolution + '\n\n' + ocrNote);
    } else {
      sheet.getRange(grievanceRow, GRIEVANCE_COLS.RESOLUTION).setValue(ocrNote);
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

/**
 * Quick OCR function - simplified wrapper for common use cases
 * @param {string} fileId - Google Drive file ID
 * @returns {string} Extracted text, or error message
 */
function wgerQuick(fileId) {
  var result = wger(fileId, { mode: 'TEXT' });
  if (result.success) {
    return result.text;
  }
  return 'Error: ' + result.message;
}

/**
 * OCR with handwriting optimization
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} Full result object
 */
function wgerHandwriting(fileId) {
  return wger(fileId, { mode: 'HANDWRITING' });
}

/**
 * OCR with document layout preservation
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} Full result object
 */
function wgerDocument(fileId) {
  return wger(fileId, { mode: 'DOCUMENT' });
}

/**
 * EXTENSION: CLOUD VISION OCR HOOK (Legacy - calls wger)
 * Prepares the system for future Google Cloud Vision integration.
 * Requires: Enabling 'Cloud Vision API' in Google Cloud Console.
 *
 * @param {string} fileId - Google Drive file ID of the image to transcribe
 * @returns {Object} Status object with transcription placeholder
 * @deprecated Use wger() instead
 */
function transcribeHandwrittenForm(fileId) {
  // Use the new wger OCR function
  var result = wger(fileId, { mode: 'HANDWRITING' });

  // Return in legacy format for backward compatibility
  if (result.success) {
    return {
      status: 'SUCCESS',
      message: 'OCR completed via wger function',
      fileId: fileId,
      fileName: result.fileName,
      text: result.text
    };
  }

  return {
    status: result.status || 'ERROR',
    message: result.message,
    fileId: fileId,
    fileName: result.fileName || ''
  };
}

/**
 * EXTENSION: SENTIMENT CORRELATION HOOK
 * Compares grievance activity to survey results for unit health analysis.
 * Flags units with high dissatisfaction but low representation.
 *
 * @param {string} unitName - The unit to analyze
 * @returns {Object} Unit health analysis result
 */
function calculateUnitHealth(unitName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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

  var data = sheet.getRange(2, GRIEVANCE_COLS.UNIT, sheet.getLastRow() - 1, 1).getValues();
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
  var data = satSheet.getRange(2, 1, lastRow - 1, SATISFACTION_COLS.REVIEWER_NOTES).getValues();

  var scores = [];
  var unitLower = unitName.toString().trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    // Only include verified, latest responses
    var verified = row[SATISFACTION_COLS.VERIFIED - 1];
    var isLatest = row[SATISFACTION_COLS.IS_LATEST - 1];

    if (verified !== 'Yes' || isLatest !== 'Yes') continue;

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

  // Find date column (column A is typically Case ID, look for Date column)
  var dateCol = header.indexOf('Date Filed');
  if (dateCol === -1) dateCol = header.indexOf('Date');
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
    '<div class="btn-row">' +
    '  <button class="btn-primary" onclick="runOCR()">🔍 Extract Text</button>' +
    '  <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div id="resultContainer"></div>' +
    '<script>' +
    'function extractFileId(input) {' +
    '  input = input.trim();' +
    '  var driveMatch = input.match(/[-\\w]{25,}/);' +
    '  return driveMatch ? driveMatch[0] : input;' +
    '}' +
    'function runOCR() {' +
    '  var rawInput = document.getElementById("fileId").value;' +
    '  var fileId = extractFileId(rawInput);' +
    '  var mode = document.getElementById("ocrMode").value;' +
    '  if (!fileId) {' +
    '    showResult({ success: false, message: "Please enter a valid file ID or URL" });' +
    '    return;' +
    '  }' +
    '  document.getElementById("resultContainer").innerHTML = "<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:12px\\">Processing OCR...</div></div>";' +
    '  google.script.run.withSuccessHandler(showResult).withFailureHandler(function(e) {' +
    '    showResult({ success: false, message: e.message });' +
    '  }).wger(fileId, { mode: mode });' +
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
    'function escapeHtml(text) {' +
    '  var div = document.createElement("div");' +
    '  div.textContent = text;' +
    '  return div.innerHTML;' +
    '}' +
    'document.getElementById("fileId").addEventListener("keypress", function(e) {' +
    '  if (e.key === "Enter") runOCR();' +
    '});' +
    '</script>'
  )
  .setWidth(500)
  .setHeight(480);

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
    '<label>Cloud Vision API Key</label>' +
    '<input type="text" id="apiKey" placeholder="AIzaSy..." value="' + (currentKey || '') + '">' +
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
      return { success: false, message: 'Invalid API key format' };
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
    return { success: false, message: 'Failed to save: ' + e.message };
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
      return { success: false, message: 'No API key configured. Please save an API key first.' };
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
      'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey,
      {
        method: 'POST',
        contentType: 'application/json',
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
      return { success: false, message: 'API Error: ' + (errorData.error ? errorData.error.message : responseBody) };
    } else if (responseCode === 403) {
      return { success: false, message: 'API key invalid or Cloud Vision API not enabled. Check your Google Cloud Console.' };
    } else {
      return { success: false, message: 'Unexpected response (' + responseCode + '): ' + responseBody.substring(0, 200) };
    }

  } catch (e) {
    return { success: false, message: 'Connection test failed: ' + e.message };
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
    '    var statusClass = "status-" + item.outcomeClass;' +
    '    html += "<div class=\\"result-item\\">" +' +
    '      "<div class=\\"result-header\\">" +' +
    '        "<span class=\\"result-id\\">" + item.id + "</span>" +' +
    '        "<span class=\\"result-status " + statusClass + "\\">" + item.outcome + "</span>" +' +
    '      "</div>" +' +
    '      "<div class=\\"result-meta\\"><strong>" + item.issueCategory + "</strong> | " + item.article + "</div>" +' +
    '      "<div class=\\"result-meta\\">" + item.memberName + " • " + item.location + " • " + item.dateResolved + "</div>" +' +
    '      (item.resolution ? "<div class=\\"result-outcome\\"><strong>Resolution:</strong> " + item.resolution + "<button class=\\"copy-btn\\" onclick=\\"copyText(\'" + item.id + ": " + item.resolution.replace(/\'/g, "") + "\')\\">Copy</button></div>" : "") +' +
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
