/**
 * ============================================================================
 * 509 STRATEGIC COMMAND CENTER - CONSOLIDATED ENGINE (v3.6.5)
 * ============================================================================
 * AUTHOR: Gemini Assistant / Claude Assistant
 * STATUS: Fully Integrated / Production Ready
 *
 * FEATURES:
 * - Security: Audit Log & Sabotage Protection (>15 cells)
 * - Workflow: Stage-Gate Case Tracking (Step 1 -> Arbitration)
 * - Automation: Auto-ID Generator & Duplicate Prevention
 * - UI/UX: Fully Automated "Roboto" Theme & Zebra Striping
 * - Legal: Signature-Ready PDF Merge & Auto-Drive Archiving
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
var COMMAND_CENTER_CONFIG = {
  SYSTEM_NAME: "509 Strategic Command Center",
  LOG_SHEET_NAME: SHEETS.GRIEVANCE_LOG,      // Maps to existing SHEETS constant
  DIR_SHEET_NAME: SHEETS.MEMBER_DIR,         // Maps to existing SHEETS constant
  AUDIT_SHEET_NAME: SHEETS.AUDIT_LOG,        // Maps to existing SHEETS constant
  TEMPLATE_ID: COMMAND_CONFIG.TEMPLATE_ID,
  ARCHIVE_FOLDER_ID: COMMAND_CONFIG.ARCHIVE_FOLDER_ID,
  CHIEF_STEWARD_EMAIL: COMMAND_CONFIG.CHIEF_STEWARD_EMAIL,
  UNIT_CODES: COMMAND_CONFIG.UNIT_CODES,
  THEME: COMMAND_CONFIG.THEME
};

// ============================================================================
// COMMAND CENTER MENU
// ============================================================================

/**
 * Creates the 509 Command Center menu
 * Called by createDashboardMenu() in UIService.gs
 */
function createCommandCenterMenu() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('📊 509 COMMAND CENTER')
    .addItem('👁️ Refresh Dashboard UI', 'APPLY_SYSTEM_THEME')
    .addSeparator()
    .addSubMenu(ui.createMenu('👤 Personnel Management')
      .addItem('🆔 Generate Missing Member IDs', 'generateMissingMemberIDs')
      .addItem('🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs')
      .addItem('🌟 Promote Selected to Steward', 'promoteSelectedMemberToSteward')
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Grievance Tools')
      .addItem('🚦 Apply Traffic Light Indicators', 'applyTrafficLightIndicators')
      .addItem('🔄 Clear Traffic Lights', 'clearTrafficLightIndicators')
      .addItem('📄 Create PDF for Selected', 'createPDFForSelectedGrievance'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛡️ System Security')
      .addItem('📸 Create Manual Snapshot', 'createWeeklySnapshot')
      .addItem('📅 Setup Weekly Backup', 'setupWeeklySnapshotTrigger')
      .addItem('📜 View Audit Log', 'navigateToAuditLog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🎨 Styling & Theme')
      .addItem('🎨 Apply Global Theme', 'APPLY_SYSTEM_THEME')
      .addItem('🔄 Reset to Default', 'resetToDefaultTheme')
      .addItem('✨ Refresh All Visuals', 'refreshAllVisuals'))
    .addToUi();
}

// ============================================================================
// NAVIGATION SHORTCUTS
// ============================================================================

/**
 * Quick navigation to Executive Dashboard
 */
function navigateToDashboard() {
  navToDash();
}

/**
 * Quick navigation to Custom View
 */
function navigateToCustomView() {
  navToCustom();
}

/**
 * Quick navigation to Mobile View
 */
function navigateToMobileView() {
  navToMobile();
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
    '• ENABLE Production Mode (hide Demo menu)\n\n' +
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
    'Type YES to confirm this is intentional.',
    ui.ButtonSet.YES_NO
  );

  if (confirm2 !== ui.Button.YES) {
    ui.alert('Cancelled', 'Nuclear operation cancelled.', ui.ButtonSet.OK);
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
      'The Demo Data menu will be hidden on reload.\n\n' +
      'Please reload the spreadsheet to see changes.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error', 'Nuclear operation failed: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// DIAGNOSTIC & REPAIR FUNCTIONS
// ============================================================================

/**
 * Runs a comprehensive diagnostic check on the dashboard setup
 */
function DIAGNOSE_SETUP() {
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
 * Repairs the dashboard by rebuilding missing components
 */
function REPAIR_DASHBOARD() {
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
