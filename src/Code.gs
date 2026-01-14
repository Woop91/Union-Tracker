/**
 * 509 Dashboard - Main Entry Point
 *
 * Core setup functions, menu system, and sheet creation.
 *
 * ⚠️ WARNING: DO NOT DEPLOY THIS FILE DIRECTLY
 * This is a source file used to generate ConsolidatedDashboard.gs.
 * Deploy ONLY ConsolidatedDashboard.gs to avoid function conflicts.
 *
 * @version 1.0.0
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
 * Main setup function - creates the complete 509 Dashboard
 * Creates the core sheets with proper structure and formatting
 */
function CREATE_509_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Confirm with user
  var response = ui.alert(
    '🏗️ Create 509 Dashboard',
    'This will create the 509 Dashboard with:\n\n' +
    '• Config (dropdown sources)\n' +
    '• Member Directory\n' +
    '• Grievance Log\n' +
    '• 💼 Dashboard (Executive metrics)\n' +
    '• 🎯 Custom View (Customizable metrics)\n' +
    '• 📊 Member Satisfaction (Survey tracking)\n' +
    '• 💡 Feedback & Development (Bug/feature tracking)\n' +
    '• ✅ Function Checklist (function reference)\n' +
    '• 📚 Getting Started (setup instructions)\n' +
    '• ❓ FAQ (common questions)\n' +
    '• 📖 Config Guide (how to use Config tab)\n\n' +
    'Plus 6 hidden calculation sheets for self-healing formulas.\n\n' +
    'Existing sheets with matching names will be recreated.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Setup cancelled.');
    return;
  }

  ss.toast('Starting dashboard creation...', '🏗️ Setup', 5);

  // Track created sheets for potential rollback
  var createdSheets = [];
  var setupState = {
    step: '',
    success: false,
    sheetsCreated: 0
  };

  try {
    // Phase 1: Core data sheets (most critical)
    setupState.step = 'Config sheet';
    createConfigSheet(ss);
    createdSheets.push(SHEETS.CONFIG);
    setupState.sheetsCreated++;
    ss.toast('Created Config sheet (1/11)', '🏗️ Progress', 2);

    setupState.step = 'Member Directory';
    createMemberDirectory(ss);
    createdSheets.push(SHEETS.MEMBER_DIR);
    setupState.sheetsCreated++;
    ss.toast('Created Member Directory (2/11)', '🏗️ Progress', 2);

    setupState.step = 'Grievance Log';
    createGrievanceLog(ss);
    createdSheets.push(SHEETS.GRIEVANCE_LOG);
    setupState.sheetsCreated++;
    ss.toast('Created Grievance Log (3/11)', '🏗️ Progress', 2);

    // Phase 2: Hidden calculation sheets
    setupState.step = 'Hidden sheets';
    ss.toast('Setting up hidden sheets...', '🏗️ Progress', 3);
    setupHiddenSheets(ss);

    // Phase 3: Dashboard and view sheets
    setupState.step = 'Dashboard';
    createDashboard(ss);
    createdSheets.push(SHEETS.DASHBOARD);
    setupState.sheetsCreated++;
    ss.toast('Created Dashboard (4/11)', '🏗️ Progress', 2);

    setupState.step = 'Custom View';
    createInteractiveDashboard(ss);
    createdSheets.push(SHEETS.INTERACTIVE);
    setupState.sheetsCreated++;
    ss.toast('Created Custom View (5/11)', '🏗️ Progress', 2);

    setupState.step = 'Member Satisfaction';
    createSatisfactionSheet(ss);
    createdSheets.push(SHEETS.SATISFACTION);
    setupState.sheetsCreated++;
    ss.toast('Created Member Satisfaction (6/11)', '🏗️ Progress', 2);

    setupState.step = 'Feedback sheet';
    createFeedbackSheet(ss);
    createdSheets.push(SHEETS.FEEDBACK);
    setupState.sheetsCreated++;
    ss.toast('Created Feedback & Development (7/11)', '🏗️ Progress', 2);

    // Phase 4: Documentation sheets (non-critical)
    setupState.step = 'Function Checklist';
    createFunctionChecklistSheet_();
    createdSheets.push(SHEETS.FUNCTION_CHECKLIST);
    setupState.sheetsCreated++;
    ss.toast('Created Function Checklist (8/11)', '🏗️ Progress', 2);

    setupState.step = 'Getting Started';
    createGettingStartedSheet(ss);
    createdSheets.push(SHEETS.GETTING_STARTED);
    setupState.sheetsCreated++;
    ss.toast('Created Getting Started (9/11)', '🏗️ Progress', 2);

    setupState.step = 'FAQ';
    createFAQSheet(ss);
    createdSheets.push(SHEETS.FAQ);
    setupState.sheetsCreated++;
    ss.toast('Created FAQ (10/11)', '🏗️ Progress', 2);

    setupState.step = 'Config Guide';
    createConfigGuideSheet(ss);
    createdSheets.push(SHEETS.CONFIG_GUIDE);
    setupState.sheetsCreated++;
    ss.toast('Created Config Guide (11/11)', '🏗️ Progress', 2);

    // Phase 5: Final setup
    setupState.step = 'Form URLs';
    saveFormUrlsToConfig_silent(ss);
    ss.toast('Saved form URLs to Config', '🏗️ Progress', 2);

    setupState.step = 'Data validations';
    ss.toast('Setting up validations...', '🏗️ Progress', 3);
    setupDataValidations();

    // Move Config to first position
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      ss.setActiveSheet(configSheet);
      ss.moveActiveSheet(1);
    }

    setupState.success = true;
    ss.toast('Dashboard creation complete!', '✅ Success', 5);
    ui.alert('✅ Success', '509 Dashboard has been created successfully!\n\n' +
      '11 sheets created:\n' +
      '• Config, Member Directory, Grievance Log (data)\n' +
      '• 💼 Dashboard, 🎯 Custom View (views)\n' +
      '• 📊 Member Satisfaction, 💡 Feedback (tracking)\n' +
      '• ✅ Function Checklist (function reference)\n' +
      '• 📚 Getting Started, ❓ FAQ, 📖 Config Guide (help)\n\n' +
      'Plus 6 hidden calculation sheets with self-healing formulas.\n\n' +
      '⚡ Auto-sync trigger installed - dates and deadlines will\n' +
      'update automatically when you edit the sheets.\n\n' +
      'Use the Demo menu to seed sample data.', ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in CREATE_509_DASHBOARD at step "' + setupState.step + '": ' + error.message);

    // Offer rollback for partial failures
    var rollbackResponse = ui.alert(
      '❌ Setup Failed',
      'An error occurred while creating: ' + setupState.step + '\n\n' +
      'Error: ' + error.message + '\n\n' +
      'Sheets created before error: ' + setupState.sheetsCreated + '\n\n' +
      'Would you like to roll back and delete the partially created sheets?\n\n' +
      '• YES = Remove all sheets created during this attempt\n' +
      '• NO = Keep partial setup (you can repair manually)',
      ui.ButtonSet.YES_NO
    );

    if (rollbackResponse === ui.Button.YES) {
      rollbackDashboardCreation_(ss, createdSheets);
      ui.alert('🔄 Rollback Complete',
        'Removed ' + createdSheets.length + ' sheet(s) created during the failed setup.\n\n' +
        'Please check the script logs for more details about the error.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('ℹ️ Partial Setup Retained',
        setupState.sheetsCreated + ' sheet(s) were created before the error.\n\n' +
        'You can try running REPAIR_DASHBOARD to fix the incomplete setup,\n' +
        'or manually delete the partial sheets and start over.',
        ui.ButtonSet.OK);
    }
  }
}

/**
 * Rollback helper - removes sheets created during a failed setup attempt
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Array<string>} sheetNames - Names of sheets to remove
 * @private
 */
function rollbackDashboardCreation_(ss, sheetNames) {
  var removed = 0;
  for (var i = 0; i < sheetNames.length; i++) {
    try {
      var sheet = ss.getSheetByName(sheetNames[i]);
      if (sheet) {
        ss.deleteSheet(sheet);
        removed++;
        Logger.log('Rollback: Removed sheet "' + sheetNames[i] + '"');
      }
    } catch (e) {
      Logger.log('Rollback: Could not remove sheet "' + sheetNames[i] + '": ' + e.message);
    }
  }

  // Also try to remove hidden sheets that may have been created
  var hiddenSheets = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_CONTACT_CALC,
    SHEETS.DASHBOARD_CALC,
    SHEETS.STEWARD_PERFORMANCE_CALC
  ];

  for (var j = 0; j < hiddenSheets.length; j++) {
    try {
      var hiddenSheet = ss.getSheetByName(hiddenSheets[j]);
      if (hiddenSheet) {
        ss.deleteSheet(hiddenSheet);
        Logger.log('Rollback: Removed hidden sheet "' + hiddenSheets[j] + '"');
      }
    } catch (e) {
      // Hidden sheet may not exist, ignore
    }
  }

  Logger.log('Rollback complete: Removed ' + removed + ' visible sheets');
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Create or recreate the Config sheet with dropdown values
 * Comprehensive configuration with section groupings and organization settings
 */
function createConfigSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG);
  sheet.clear();

  // Row 1: Section Headers (grouped categories)
  var sectionHeaders = [
    '── EMPLOYMENT INFO ──', '', '', '', '',           // A-E (5 cols)
    '── SUPERVISION ──', '',                            // F-G (2 cols)
    '── STEWARD INFO ──', '',                           // H-I (2 cols)
    '── GRIEVANCE SETTINGS ──', '', '', '',             // J-M (4 cols)
    '── LINKS & COORDINATORS ──', '', '', '',           // N-Q (4 cols)
    '── NOTIFICATIONS ──', '', '',                      // R-T (3 cols)
    '── ORGANIZATION ──', '', '', '',                   // U-X (4 cols)
    '── INTEGRATION ──', '',                            // Y-Z (2 cols)
    '── DEADLINES ──', '', '', '',                      // AA-AD (4 cols)
    '── MULTI-SELECT OPTIONS ──', '',                   // AE-AF (2 cols)
    '── CONTRACT & LEGAL ──', '', '', '',               // AG-AJ (4 cols)
    '── ORG IDENTITY ──', '', '',                       // AK-AM (3 cols)
    '── EXTENDED CONTACT ──', '', '', ''                // AN-AQ (4 cols)
  ];

  // Row 2: Column Headers
  var columnHeaders = [
    'Job Titles', 'Office Locations', 'Units', 'Office Days', 'Yes/No (Dropdowns)',
    'Supervisors', 'Managers',
    'Stewards', 'Steward Committees',
    'Grievance Status', 'Grievance Step', 'Issue Category', 'Articles Violated',
    'Communication Methods', 'Grievance Coordinators', 'Grievance Form URL', 'Contact Form URL',
    'Admin Emails', 'Alert Days Before Deadline', 'Notification Recipients',
    'Organization Name', 'Local Number', 'Main Office Address', 'Main Phone',
    'Google Drive Folder ID', 'Google Calendar ID',
    'Filing Deadline Days', 'Step I Response Days', 'Step II Appeal Days', 'Step II Response Days',
    'Best Times to Contact', 'Home Towns',
    'Contract Article (Grievance)', 'Contract Article (Discipline)', 'Contract Article (Workload)', 'Contract Name',
    'Union Parent', 'State/Region', 'Organization Website',
    'Office Addresses', 'Main Fax', 'Main Contact Name', 'Main Contact Email'
  ];

  // Apply section headers (Row 1)
  sheet.getRange(1, 1, 1, sectionHeaders.length).setValues([sectionHeaders])
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK)
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Apply column headers (Row 2)
  sheet.getRange(2, 1, 1, columnHeaders.length).setValues([columnHeaders])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Add default dropdown values (Row 3+)
  // Office Days (D)
  sheet.getRange(3, CONFIG_COLS.OFFICE_DAYS, DEFAULT_CONFIG.OFFICE_DAYS.length, 1)
    .setValues(DEFAULT_CONFIG.OFFICE_DAYS.map(function(v) { return [v]; }));

  // Yes/No (E)
  sheet.getRange(3, CONFIG_COLS.YES_NO, DEFAULT_CONFIG.YES_NO.length, 1)
    .setValues(DEFAULT_CONFIG.YES_NO.map(function(v) { return [v]; }));

  // Steward Committees (I)
  var committees = ['Grievance Committee', 'Bargaining Committee', 'Health & Safety Committee',
                    'Political Action Committee', 'Membership Committee', 'Executive Board'];
  sheet.getRange(3, CONFIG_COLS.STEWARD_COMMITTEES, committees.length, 1)
    .setValues(committees.map(function(v) { return [v]; }));

  // Grievance Status (J)
  sheet.getRange(3, CONFIG_COLS.GRIEVANCE_STATUS, DEFAULT_CONFIG.GRIEVANCE_STATUS.length, 1)
    .setValues(DEFAULT_CONFIG.GRIEVANCE_STATUS.map(function(v) { return [v]; }));

  // Grievance Step (K)
  sheet.getRange(3, CONFIG_COLS.GRIEVANCE_STEP, DEFAULT_CONFIG.GRIEVANCE_STEP.length, 1)
    .setValues(DEFAULT_CONFIG.GRIEVANCE_STEP.map(function(v) { return [v]; }));

  // Issue Category (L)
  sheet.getRange(3, CONFIG_COLS.ISSUE_CATEGORY, DEFAULT_CONFIG.ISSUE_CATEGORY.length, 1)
    .setValues(DEFAULT_CONFIG.ISSUE_CATEGORY.map(function(v) { return [v]; }));

  // Articles (M)
  sheet.getRange(3, CONFIG_COLS.ARTICLES, DEFAULT_CONFIG.ARTICLES.length, 1)
    .setValues(DEFAULT_CONFIG.ARTICLES.map(function(v) { return [v]; }));

  // Communication Methods (N)
  sheet.getRange(3, CONFIG_COLS.COMM_METHODS, DEFAULT_CONFIG.COMM_METHODS.length, 1)
    .setValues(DEFAULT_CONFIG.COMM_METHODS.map(function(v) { return [v]; }));

  // Alert Days (S) - default notification intervals
  sheet.getRange(3, CONFIG_COLS.ALERT_DAYS, 1, 1).setValue('3, 7, 14');

  // Organization defaults (SEIU Local 509)
  sheet.getRange(3, CONFIG_COLS.ORG_NAME, 1, 1).setValue('SEIU Local 509');
  sheet.getRange(3, CONFIG_COLS.LOCAL_NUMBER, 1, 1).setValue('509');
  sheet.getRange(3, CONFIG_COLS.MAIN_ADDRESS, 1, 1).setValue('293 Boston Post Road West, 4th Floor, Marlborough, MA 01752');
  sheet.getRange(3, CONFIG_COLS.MAIN_PHONE, 1, 1).setValue('774-843-7509');

  // Deadline defaults (in days)
  sheet.getRange(3, CONFIG_COLS.FILING_DEADLINE_DAYS, 1, 1).setValue(21);
  sheet.getRange(3, CONFIG_COLS.STEP1_RESPONSE_DAYS, 1, 1).setValue(30);
  sheet.getRange(3, CONFIG_COLS.STEP2_APPEAL_DAYS, 1, 1).setValue(10);
  sheet.getRange(3, CONFIG_COLS.STEP2_RESPONSE_DAYS, 1, 1).setValue(30);

  // Best Times to Contact (AE)
  var bestTimes = ['Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Weekends', 'Flexible'];
  sheet.getRange(3, CONFIG_COLS.BEST_TIMES, bestTimes.length, 1)
    .setValues(bestTimes.map(function(v) { return [v]; }));

  // Contract articles
  sheet.getRange(3, CONFIG_COLS.CONTRACT_GRIEVANCE, 1, 1).setValue('Article 23A');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_DISCIPLINE, 1, 1).setValue('Article 12');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_WORKLOAD, 1, 1).setValue('Article 15');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_NAME, 1, 1).setValue('2023-2026 CBA');

  // Org identity
  sheet.getRange(3, CONFIG_COLS.UNION_PARENT, 1, 1).setValue('SEIU');
  sheet.getRange(3, CONFIG_COLS.STATE_REGION, 1, 1).setValue('Massachusetts');
  sheet.getRange(3, CONFIG_COLS.ORG_WEBSITE, 1, 1).setValue('https://www.seiu509.org/');

  // Extended contact
  sheet.getRange(3, CONFIG_COLS.MAIN_FAX, 1, 1).setValue('508-485-8529');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_NAME, 1, 1).setValue('Marc');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_EMAIL, 1, 1).setValue('marc@seiu509.org');

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
}

/**
 * Creates the Config Guide sheet - a dedicated tab explaining how to use the Config tab
 */
function createConfigGuideSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.CONFIG_GUIDE || '📖 Config Guide';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define guide colors
  var headerBg = '#4A90D9';      // Blue header
  var sectionBg = '#E8F4FD';     // Light blue section
  var tipBg = '#FFF9E6';         // Light yellow for tips
  var warningBg = '#FEE2E2';     // Light red for warnings
  var successBg = '#DCFCE7';     // Light green for success
  var textColor = '#1F2937';     // Dark text

  var row = 1;

  // ═══ HEADER ═══
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📖 CONFIG TAB USER GUIDE')
    .setBackground(headerBg)
    .setFontColor('#FFFFFF')
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
    sheet.getRange(row, 1).setValue(howToSteps[i][0]).setFontWeight('bold').setFontColor('#4A90D9');
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
    ['A', 'Job Titles', 'Member Directory', 'Case Worker, Supervisor, Manager...'],
    ['B', 'Office Locations', 'Member Dir & Grievance Log', 'Boston Office, Springfield Office...'],
    ['C', 'Units', 'Member Dir & Grievance Log', 'Unit 1, Unit 2, Unit 3...'],
    ['F', 'Supervisors', 'Member Directory', 'Names of supervisors'],
    ['G', 'Managers', 'Member Directory', 'Names of managers'],
    ['H', 'Stewards', 'Member Dir & Grievance Log', 'Names of union stewards'],
    ['J', 'Grievance Status', 'Grievance Log', 'Open, Pending Info, Settled, Won...'],
    ['K', 'Grievance Step', 'Grievance Log', 'Informal, Step I, Step II, Step III...'],
    ['L', 'Issue Category', 'Grievance Log', 'Discipline, Workload, Pay, Benefits...'],
    ['M', 'Articles Violated', 'Grievance Log', 'Article 12, Article 23A...']
  ];

  for (var j = 0; j < columnData.length; j++) {
    row++;
    sheet.getRange(row, 1).setValue(columnData[j][0]).setFontColor('#4A90D9').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row, 2).setValue(columnData[j][1]).setFontColor(textColor);
    sheet.getRange(row, 3, 1, 2).merge().setValue(columnData[j][2]).setFontColor('#6B7280');
    sheet.getRange(row, 5, 1, 2).merge().setValue(columnData[j][3]).setFontColor('#6B7280').setFontStyle('italic');
  }

  // ═══ TIPS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('💡 Pro Tips')
    .setBackground(tipBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#92400E');
  sheet.setRowHeight(row, 35);

  var tips = [
    '✓ Keep dropdown lists in alphabetical order for easier selection',
    '✓ Use consistent naming conventions (e.g., "Boston Office" not "boston office")',
    '✓ Add your organization\'s specific values before entering member/grievance data',
    '✓ The system pre-fills some default values - modify them to match your organization'
  ];

  for (var k = 0; k < tips.length; k++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(tips[k]).setBackground(tipBg).setFontColor('#92400E');
  }

  // ═══ WARNINGS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('⚠️ Important Warnings')
    .setBackground(warningBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#DC2626');
  sheet.setRowHeight(row, 35);

  var warnings = [
    '⚠ Do NOT delete values that are already in use in Member Directory or Grievance Log',
    '⚠ Do NOT leave blank cells in the middle of a column - this breaks the dropdowns',
    '⚠ Do NOT modify the Section Headers (Row 1) or Column Headers (Row 2) in Config'
  ];

  for (var m = 0; m < warnings.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(warnings[m]).setBackground(warningBg).setFontColor('#DC2626');
  }

  // ═══ NEED HELP ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('🆘 Need More Help?')
    .setBackground(successBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#166534');
  sheet.setRowHeight(row, 35);

  row++;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Check the "📚 Getting Started" tab for full setup instructions, or the "❓ FAQ" tab for common questions.')
    .setBackground(successBg)
    .setFontColor('#166534');

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

  return sheet;
}

/**
 * Create or recreate the Member Directory sheet
 */
function createMemberDirectory(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEMBER_DIR);
  sheet.clear();

  var headers = getMemberHeaders();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(MEMBER_COLS.MEMBER_ID, 100);
  sheet.setColumnWidth(MEMBER_COLS.FIRST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.LAST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.EMAIL, 200);
  sheet.setColumnWidth(MEMBER_COLS.CONTACT_NOTES, 250);

  // Add checkbox for Start Grievance column (pre-allocate for future rows)
  sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, 4999, 1).insertCheckboxes();

  // Add checkbox for Quick Actions column (opens quick actions dialog when checked)
  sheet.getRange(2, MEMBER_COLS.QUICK_ACTIONS, 4999, 1).insertCheckboxes();

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
  sheet.getRange(2, MEMBER_COLS.OPEN_RATE, 998, 1).setNumberFormat('#,##0.0');       // S - Open Rate %
  sheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, 998, 1).setNumberFormat('#,##0');  // T - Volunteer Hours

  // Columns AB-AD (Has Open Grievance?, Grievance Status, Days to Deadline)
  // are populated by syncGrievanceToMemberDirectory() with STATIC values
  // No formulas in visible sheets - all calculations done by script

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN GROUPS: Group and hide optional columns for cleaner view
  // ═══════════════════════════════════════════════════════════════════════════
  try {
    // Group 1: Engagement Metrics (Q-T, columns 17-20) - Hidden by default
    sheet.getRange(1, MEMBER_COLS.LAST_VIRTUAL_MTG, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    // Group 2: Member Interests (U-X, columns 21-24) - Hidden by default
    sheet.getRange(1, MEMBER_COLS.INTEREST_LOCAL, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    Logger.log('Member Directory column group setup skipped: ' + e.toString());
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CONDITIONAL FORMATTING: Highlight members with open grievances
  // ═══════════════════════════════════════════════════════════════════════════
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var hasOpenGrievanceRange = sheet.getRange(2, MEMBER_COLS.HAS_OPEN_GRIEVANCE, 4999, 1);

  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#ffebee')  // Light red background
    .setFontColor('#c62828')   // Dark red text
    .setBold(true)
    .setRanges([hasOpenGrievanceRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // VALIDATION HIGHLIGHTING: Red background for empty Email and Phone fields
  // ═══════════════════════════════════════════════════════════════════════════
  var emailRange = sheet.getRange(2, MEMBER_COLS.EMAIL, 4999, 1);
  var phoneRange = sheet.getRange(2, MEMBER_COLS.PHONE, 4999, 1);

  // Rule: Red background for empty Email
  var emptyEmailRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",ISBLANK($H2))')
    .setBackground('#ffcdd2')  // Red background for missing email
    .setRanges([emailRange])
    .build();

  // Rule: Red background for empty Phone
  var emptyPhoneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",ISBLANK($I2))')
    .setBackground('#ffcdd2')  // Red background for missing phone
    .setRanges([phoneRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // DEADLINE HEATMAP: Color-coded Days to Deadline (Column AD)
  // ═══════════════════════════════════════════════════════════════════════════
  var daysDeadlineRange = sheet.getRange(2, MEMBER_COLS.NEXT_DEADLINE, 4999, 1);

  // Rule: Red - Overdue (shows "Overdue" or negative/0 days)
  var deadlineOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($AD2="Overdue",AND(ISNUMBER($AD2),$AD2<=0))')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Orange - Due in 1-3 days
  var deadline1to3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($AD2),$AD2>=1,$AD2<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Yellow - Due in 4-7 days
  var deadline4to7Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($AD2),$AD2>=4,$AD2<=7)')
    .setBackground('#fffde7')
    .setFontColor('#f57f17')
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Green - On Track (more than 7 days remaining)
  var deadlineOnTrackRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($AD2),$AD2>7)')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setRanges([daysDeadlineRange])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(redRule, emptyEmailRule, emptyPhoneRule, deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule);
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
  // Volunteer Hours, Interest: Local/Chapter/Allied, Home Town, Recent Contact Date,
  // Contact Steward, Contact Notes, Has Open Grievance?, Grievance Status, Days to Deadline
  var filterRange = sheet.getRange(1, 1, 5000, headers.length);
  filterRange.createFilter();
}

/**
 * Create or recreate the Grievance Log sheet
 * NOTE: Calculated columns (First Name, Last Name, Email, Deadlines, Days Open, etc.)
 * are managed by the hidden _Grievance_Formulas sheet for self-healing capability.
 * Users can't accidentally erase formulas because they're in the hidden sheet.
 */
function createGrievanceLog(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.GRIEVANCE_LOG);
  sheet.clear();

  var headers = getGrievanceHeaders();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');

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
    GRIEVANCE_COLS.NEXT_ACTION_DUE
  ];

  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, 998, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Format Days Open (S) and Days to Deadline (U) as whole numbers with comma separators
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, 998, 1).setNumberFormat('#,##0');
  // Days to Deadline can show "Overdue" text, so use General format that handles both
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 998, 1).setNumberFormat('#,##0');

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // Setup column groups for timeline (Step I, II, III collapsible)
  try {
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    Logger.log('Column group setup skipped: ' + e.toString());
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DAYS TO DEADLINE HEATMAP (Column U)
  // ═══════════════════════════════════════════════════════════════════════════
  var daysDeadlineRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 4999, 1);

  // Rule: Red - Overdue (shows "Overdue" or negative/0 days)
  var deadlineOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($U2="Overdue",AND(ISNUMBER($U2),$U2<=0))')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Orange - Due in 1-3 days
  var deadline1to3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=1,$U2<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Yellow - Due in 4-7 days
  var deadline4to7Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=4,$U2<=7)')
    .setBackground('#fffde7')
    .setFontColor('#f57f17')
    .setRanges([daysDeadlineRange])
    .build();

  // Rule: Green - On Track (more than 7 days remaining)
  var deadlineOnTrackRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>7)')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setRanges([daysDeadlineRange])
    .build();

  // ═══════════════════════════════════════════════════════════════════════════
  // PROGRESS BAR: Colored backgrounds showing grievance stage (Columns J-R)
  // Based on Current Step (Column F), highlights completed stages
  // ═══════════════════════════════════════════════════════════════════════════

  // Progress bar spans: Step I (J-K), Step II (L-O), Step III (P-Q), Date Closed (R)
  var step1Range = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 2);         // J-K
  var step2Range = sheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, 4999, 4);  // L-O
  var step3Range = sheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, 4999, 2);  // P-Q
  var closedRange = sheet.getRange(2, GRIEVANCE_COLS.DATE_CLOSED, 4999, 1);      // R
  var allStepsRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 9);      // J-R (all 9 columns)

  // Completed cases: All columns green (Closed, Won, Denied, Settled, Withdrawn)
  var completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($E2="Closed",$E2="Won",$E2="Denied",$E2="Settled",$E2="Withdrawn")')
    .setBackground('#e8f5e9')  // Soft green
    .setRanges([allStepsRange])
    .build();

  // Step III in progress: J-Q highlighted (all except Date Closed)
  var step3ProgressRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 8);  // J-Q
  var step3ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Step III"')
    .setBackground('#e3f2fd')  // Soft blue
    .setRanges([step3ProgressRange])
    .build();

  // Step II in progress: J-O highlighted
  var step2ProgressRange = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, 4999, 6);  // J-O
  var step2ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Step II"')
    .setBackground('#e3f2fd')  // Soft blue
    .setRanges([step2ProgressRange])
    .build();

  // Step I in progress: J-K highlighted
  var step1ProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Step I"')
    .setBackground('#e3f2fd')  // Soft blue
    .setRanges([step1Range])
    .build();

  // Gray out columns not yet reached (applies to all step columns by default)
  var notReachedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$F2<>"")')
    .setBackground('#fafafa')  // Very light gray (default for uncolored)
    .setRanges([allStepsRange])
    .build();

  // Apply all rules (order matters - more specific rules first)
  var rules = sheet.getConditionalFormatRules();
  rules.push(
    deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule,
    completedRule, step3ProgressRule, step2ProgressRule, step1ProgressRule, notReachedRule
  );
  sheet.setConditionalFormatRules(rules);
}


/**
 * Create or recreate the unified Dashboard sheet (Executive Dashboard theme)
 * Combines member metrics, grievance metrics, and key performance indicators
 * Uses Executive Dashboard's green QUICK STATS theme
 */
function createDashboard(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.DASHBOARD);
  sheet.clear();

  // Title - Executive Dashboard style
  sheet.getRange('A1').setValue('💼 Executive Dashboard')
    .setFontSize(24)
    .setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.getRange('A1:F1').merge();

  // Subtitle with last refresh
  sheet.getRange('A2').setValue('Real-time metrics from Member Directory & Grievance Log')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontColor(COLORS.TEXT_DARK);
  sheet.getRange('A2:F2').merge();

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 1: QUICK STATS (Executive Dashboard green theme)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A4').setValue('QUICK STATS')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A4:F4').merge();

  // Quick stats row - pulls from hidden _Dashboard_Calc sheet
  var quickStatsLabels = [
    ['Total Members', 'Active Stewards', 'Active Grievances', 'Win Rate', 'Overdue Cases', 'Due This Week']
  ];
  sheet.getRange('A5:F5').setValues(quickStatsLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Quick stats values (populated by syncDashboardValues)
  sheet.getRange('A6:F6')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 2: MEMBER METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A8').setValue('MEMBER METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor(COLORS.TEXT_DARK);
  sheet.getRange('A8:D8').merge();

  var memberMetricLabels = [['Total Members', 'Active Stewards', 'Avg Open Rate', 'YTD Vol. Hours']];
  sheet.getRange('A9:D9').setValues(memberMetricLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Member metric values (populated by syncDashboardValues)
  sheet.getRange('A10:D10')
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 3: GRIEVANCE METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A12').setValue('GRIEVANCE METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A12:F12').merge();

  var grievanceLabels = [['Open', 'Pending Info', 'Settled', 'Won', 'Denied', 'Withdrawn']];
  sheet.getRange('A13:F13').setValues(grievanceLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Grievance metric values (populated by syncDashboardValues)
  sheet.getRange('A14:F14')
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 4: TIMELINE METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A16').setValue('TIMELINE & PERFORMANCE')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A16:D16').merge();

  var timelineLabels = [['Avg Days Open', 'Filed This Month', 'Closed This Month', 'Avg Resolution Days']];
  sheet.getRange('A17:D17').setValues(timelineLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Timeline metric values (populated by syncDashboardValues)
  sheet.getRange('A18:D18')
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 5: TYPE ANALYSIS (by Issue Category)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A20').setValue('📊 TYPE ANALYSIS (by Issue Category)')
    .setFontWeight('bold')
    .setBackground('#6366F1')  // Indigo
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A20:F20').merge();

  var typeLabels = [['Issue Category', 'Total Cases', 'Open', 'Resolved', 'Win Rate', 'Avg Days']];
  sheet.getRange('A21:F21').setValues(typeLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Type analysis values (populated by syncDashboardValues)
  sheet.getRange('A22:F26').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 6: LOCATION BREAKDOWN
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A28').setValue('🗺️ LOCATION BREAKDOWN')
    .setFontWeight('bold')
    .setBackground('#0891B2')  // Cyan
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A28:F28').merge();

  var locationLabels = [['Location', 'Members', 'Grievances', 'Open Cases', 'Win Rate', 'Avg Satisfaction']];
  sheet.getRange('A29:F29').setValues(locationLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Location values (populated by syncDashboardValues)
  sheet.getRange('A30:F34').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 7: MONTH-OVER-MONTH TRENDS (moved up from rows 40-44)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A36').setValue('📈 MONTH-OVER-MONTH TRENDS')
    .setFontWeight('bold')
    .setBackground('#DC2626')  // Red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A36:F36').merge();

  var trendLabels = [['Metric', 'This Month', 'Last Month', 'Change', '% Change', 'Trend']];
  sheet.getRange('A37:F37').setValues(trendLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Trend values (populated by syncDashboardValues)
  sheet.getRange('A38:F40').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 8: STATUS LEGEND (compact - fits in A-F only)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A42').setValue('STATUS LEGEND')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange('A42:F42').merge();

  // Deadline status - compact version
  var legendDeadline = [
    ['🟢 >7d', '🟡 4-7d', '🟠 1-3d', '🔴 Overdue', '✅ Won', '❌ Denied']
  ];
  sheet.getRange('A43:F43').setValues(legendDeadline)
    .setHorizontalAlignment('center')
    .setFontSize(9);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 9: STEWARD PERFORMANCE SUMMARY (moved to near bottom)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A45').setValue('👨‍⚖️ STEWARD PERFORMANCE SUMMARY')
    .setFontWeight('bold')
    .setBackground('#7C3AED')  // Purple
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A45:F45').merge();

  var stewardLabels = [['Total Stewards', 'Active (w/ Cases)', 'Avg Cases/Steward', 'Total Vol Hours', 'Contacts This Month', '']];
  sheet.getRange('A46:F46').setValues(stewardLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Steward summary values (populated by syncDashboardValues)
  sheet.getRange('A47:F47')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 10: TOP 30 BUSIEST STEWARDS (by active grievance count)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A49').setValue('📊 TOP 30 BUSIEST STEWARDS (Active Cases)')
    .setFontWeight('bold')
    .setBackground('#B91C1C')  // Dark red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A49:F49').merge();

  var busiestLabels = [['Rank', 'Steward Name', 'Active Cases', 'Open', 'Pending Info', 'Total Ever']];
  sheet.getRange('A50:F50').setValues(busiestLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Busiest stewards values (populated by syncDashboardValues)
  sheet.getRange('A51:F80').setHorizontalAlignment('center');

  // Alternate row coloring for busiest stewards list
  for (var r = 51; r <= 80; r++) {
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#F9FAFB');
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 11: TOP 10 PERFORMERS BY SCORE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A82').setValue('🏆 TOP 10 PERFORMERS BY SCORE')
    .setFontWeight('bold')
    .setBackground('#059669')  // Green
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A82:F82').merge();

  var topPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A83:F83').setValues(topPerfLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Top performers values (populated by syncDashboardValues)
  // Alternate row coloring for top performers
  for (var r = 84; r <= 93; r++) {
    if (r % 2 === 0) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#ECFDF5');  // Light green
    }
  }
  sheet.getRange('A84:F93').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 12: STEWARDS NEEDING SUPPORT (Bottom 10 by Score)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A95').setValue('⚠️ STEWARDS NEEDING SUPPORT (Lowest Scores)')
    .setFontWeight('bold')
    .setBackground('#DC2626')  // Red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A95:F95').merge();

  var lowPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A96:F96').setValues(lowPerfLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Stewards needing support values (populated by syncDashboardValues)
  // Alternate row coloring for bottom performers
  for (var r = 97; r <= 106; r++) {
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#FEF2F2');  // Light red
    }
  }
  sheet.getRange('A97:F106').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // FORMATTING AND CLEANUP
  // ═══════════════════════════════════════════════════════════════════════════

  // Auto-resize and format
  sheet.autoResizeColumns(1, 6);
  sheet.setFrozenRows(3);

  // Set minimum column widths
  for (var i = 1; i <= 6; i++) {
    if (sheet.getColumnWidth(i) < 120) {
      sheet.setColumnWidth(i, 120);
    }
  }

  // Hide column G onwards (if any exist)
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 6) {
      sheet.deleteColumns(7, maxCols - 6);
    }
  } catch (e) {
    Logger.log('Could not delete extra columns: ' + e.toString());
  }

  // NUMBER FORMATTING: Use comma separators for all numeric values (1,000)
  var numberFormat = '#,##0';
  // Quick Stats row (row 6)
  sheet.getRange('A6:C6').setNumberFormat(numberFormat);
  sheet.getRange('E6:F6').setNumberFormat(numberFormat);
  // Member Metrics row (row 10)
  sheet.getRange('A10:B10').setNumberFormat(numberFormat);
  sheet.getRange('D10').setNumberFormat(numberFormat);
  // Grievance Metrics row (row 14)
  sheet.getRange('A14:F14').setNumberFormat(numberFormat);
  // Timeline row (row 18)
  sheet.getRange('A18:D18').setNumberFormat(numberFormat);
  // Type Analysis rows (rows 22-26)
  sheet.getRange('B22:D26').setNumberFormat(numberFormat);
  // Location Breakdown rows (rows 30-34)
  sheet.getRange('B30:D34').setNumberFormat(numberFormat);
  // Trends rows (rows 38-40)
  sheet.getRange('B38:D40').setNumberFormat(numberFormat);
  // Steward Performance row (row 47)
  sheet.getRange('A47:E47').setNumberFormat(numberFormat);
  // Top 30 Busiest Stewards (rows 51-80)
  sheet.getRange('C51:F80').setNumberFormat(numberFormat);
  // Top 10 Performers (rows 84-93) - Score and Win Rate have decimals
  var decimalFormat = '#,##0.0';
  sheet.getRange('C84:D93').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E84:F93').setNumberFormat(numberFormat);   // Avg Days, Overdue
  // Stewards Needing Support (rows 97-106)
  sheet.getRange('C97:D106').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E97:F106').setNumberFormat(numberFormat);   // Avg Days, Overdue

  // Populate all values using JavaScript-computed metrics (no formulas in visible sheet)
  syncDashboardValues();
}

/**
 * Create or recreate the Custom View sheet
 * Allows users to select metrics and visualization preferences
 */
function createInteractiveDashboard(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.INTERACTIVE);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('🎯 Custom View')
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.getRange('A1:F1').merge();

  // Instructions
  sheet.getRange('A3').setValue('Select metrics and chart types using the dropdowns below. Metrics auto-update from live data.')
    .setFontStyle('italic');
  sheet.getRange('A3:F3').merge();

  // ═══════════════════════════════════════════════════════════════════════════
  // METRIC SELECTION ROW
  // ═══════════════════════════════════════════════════════════════════════════
  var controlLabels = [['Metric 1', 'Metric 2', 'Metric 3', 'Time Range', 'Show Trend', 'Theme']];
  sheet.getRange('A5:F5').setValues(controlLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Default selections
  var defaultSelections = [['Total Members', 'Open Grievances', 'Win Rate', 'All Time', 'Yes', 'Default']];
  sheet.getRange('A6:F6').setValues(defaultSelections);

  // ═══════════════════════════════════════════════════════════════════════════
  // SELECTED METRICS DISPLAY
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A8').setValue('SELECTED METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A8:F8').merge();

  // Headers for metrics
  sheet.getRange('A9:C9').setValues([['Metric', 'Current Value', 'Description']])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Dynamic metric formulas based on selection
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);

  // Metric lookup table (row 10-17)
  var metricData = [
    ['Total Members', '=COUNTA(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ')-1', 'Total union members in directory'],
    ['Active Stewards', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")', 'Members marked as stewards'],
    ['Total Grievances', '=COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1', 'All grievances filed'],
    ['Open Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")', 'Currently open cases'],
    ['Pending Info', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")', 'Cases awaiting information'],
    ['Settled', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")', 'Cases settled'],
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")', 'Cases won'],
    ['Win Rate', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")/(COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1)*100,1)&"%","0%")', 'Win percentage of all cases']
  ];

  for (var i = 0; i < metricData.length; i++) {
    sheet.getRange(10 + i, 1).setValue(metricData[i][0]);
    sheet.getRange(10 + i, 2).setFormula(metricData[i][1]);
    sheet.getRange(10 + i, 3).setValue(metricData[i][2]);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DROPDOWN VALIDATIONS
  // ═══════════════════════════════════════════════════════════════════════════
  // Metric dropdown options
  var metricOptions = ['Total Members', 'Active Stewards', 'Total Grievances', 'Open Grievances', 'Pending Info', 'Settled', 'Won', 'Win Rate'];
  var metricRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(metricOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A6').setDataValidation(metricRule);
  sheet.getRange('B6').setDataValidation(metricRule);
  sheet.getRange('C6').setDataValidation(metricRule);

  // Time range options
  var timeOptions = ['All Time', 'This Month', 'This Quarter', 'This Year', 'Last 30 Days', 'Last 90 Days'];
  var timeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(timeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D6').setDataValidation(timeRule);

  // Yes/No options
  var yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E6').setDataValidation(yesNoRule);

  // Theme options
  var themeOptions = ['Default', 'Dark', 'High Contrast', 'Print Friendly'];
  var themeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(themeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F6').setDataValidation(themeRule);

  // Format
  sheet.autoResizeColumns(1, 6);
  sheet.setColumnWidth(3, 250);

  // Delete excess columns after F (column 6)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 6) {
    sheet.deleteColumns(7, maxCols - 6);
  }
}

// ============================================================================
// MEMBER SATISFACTION SHEET
// ============================================================================

/**
 * Create the Member Satisfaction Tracking sheet
 * Tracks survey responses and calculates satisfaction metrics
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function createSatisfactionSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.SATISFACTION);
  sheet.clear();

  // Ensure sheet has enough columns (need 100+ for dashboard area)
  var requiredCols = 100;
  var currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // GOOGLE FORM RESPONSE HEADERS (68 questions + timestamp)
  // When you link a Google Form, these columns receive form responses
  // ═══════════════════════════════════════════════════════════════════════════

  var headers = [
    'Timestamp',                                    // A - Auto by Google Forms
    // Work Context (Q1-5)
    'Q1: Worksite/Program/Region',                  // B
    'Q2: Role/Job Group',                           // C
    'Q3: Shift',                                    // D
    'Q4: Time in current role',                     // E
    'Q5: Steward contact (12 mo)?',                 // F
    // Overall Satisfaction (Q6-9)
    'Q6: Satisfied with representation',            // G
    'Q7: Trust union acts in best interest',        // H
    'Q8: Feel protected at work',                   // I
    'Q9: Would recommend membership',               // J
    // Steward Ratings 3A (Q10-17)
    'Q10: Timely response',                         // K
    'Q11: Treated with respect',                    // L
    'Q12: Explained options clearly',               // M
    'Q13: Followed through',                        // N
    'Q14: Advocated effectively',                   // O
    'Q15: Safe raising concerns',                   // P
    'Q16: Confidentiality',                         // Q
    'Q17: Steward improvement suggestions',         // R
    // Steward Access 3B (Q18-20)
    'Q18: Know how to contact steward',             // S
    'Q19: Confident would get help',                // T
    'Q20: Easy to find who to contact',             // U
    // Chapter Effectiveness (Q21-25)
    'Q21: Reps understand issues',                  // V
    'Q22: Chapter communication',                   // W
    'Q23: Organizes effectively',                   // X
    'Q24: Know chapter contact',                    // Y
    'Q25: Fair representation',                     // Z
    // Local Leadership (Q26-31)
    'Q26: Decisions communicated clearly',          // AA
    'Q27: Understand decision process',             // AB
    'Q28: Transparent finances',                    // AC
    'Q29: Leadership accountable',                  // AD
    'Q30: Fair internal processes',                 // AE
    'Q31: Welcomes differing opinions',             // AF
    // Contract Enforcement (Q32-36)
    'Q32: Enforces contract',                       // AG
    'Q33: Realistic timelines',                     // AH
    'Q34: Clear updates',                           // AI
    'Q35: Frontline priority',                      // AJ
    'Q36: Filed grievance (24 mo)?',                // AK
    // Representation 6A (Q37-40)
    'Q37: Understood steps/timeline',               // AL
    'Q38: Felt supported',                          // AM
    'Q39: Updates often enough',                    // AN
    'Q40: Outcome justified',                       // AO
    // Communication (Q41-45)
    'Q41: Clear & actionable',                      // AP
    'Q42: Enough information',                      // AQ
    'Q43: Find info easily',                        // AR
    'Q44: Reaches all shifts',                      // AS
    'Q45: Meetings worth attending',                // AT
    // Member Voice (Q46-50)
    'Q46: Voice matters',                           // AU
    'Q47: Seeks input',                             // AV
    'Q48: Dignity',                                 // AW
    'Q49: Newer members supported',                 // AX
    'Q50: Conflict handled respectfully',           // AY
    // Value & Action (Q51-55)
    'Q51: Good value for dues',                     // AZ
    'Q52: Priorities reflect needs',                // BA
    'Q53: Prepared to mobilize',                    // BB
    'Q54: Know how to get involved',                // BC
    'Q55: Can win together',                        // BD
    // Scheduling (Q56-63)
    'Q56: Understand changes',                      // BE
    'Q57: Adequately informed',                     // BF
    'Q58: Clear criteria',                          // BG
    'Q59: Work under expectations',                 // BH
    'Q60: Effective outcomes',                      // BI
    'Q61: Supports wellbeing',                      // BJ
    'Q62: Concerns taken seriously',                // BK
    'Q63: Scheduling challenge',                    // BL
    // Priorities & Close (Q64-68)
    'Q64: Top 3 priorities',                        // BM
    'Q65: #1 change to make',                       // BN
    'Q66: Keep doing',                              // BO
    'Q67: Additional comments'                      // BP
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setWrap(true);

  // Set column widths for response data
  sheet.setColumnWidth(1, 140);  // Timestamp
  for (var c = 2; c <= 69; c++) {
    // Wider for paragraph columns (R, BL, BM, BN, BO, BP)
    if (c === 18 || c === 64 || c === 65 || c === 66 || c === 67 || c === 68) {
      sheet.setColumnWidth(c, 250);
    } else {
      sheet.setColumnWidth(c, 45);
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION AVERAGE FORMULAS (Columns BT-CD) - For Charts
  // These calculate section averages for each response row
  // ═══════════════════════════════════════════════════════════════════════════

  // Section average headers
  var sectionHeaders = [
    'Avg: Overall Sat',     // BT (72)
    'Avg: Steward Rating',  // BU (73)
    'Avg: Steward Access',  // BV (74)
    'Avg: Chapter',         // BW (75)
    'Avg: Leadership',      // BX (76)
    'Avg: Contract',        // BY (77)
    'Avg: Representation',  // BZ (78)
    'Avg: Communication',   // CA (79)
    'Avg: Member Voice',    // CB (80)
    'Avg: Value/Action',    // CC (81)
    'Avg: Scheduling'       // CD (82)
  ];

  sheet.getRange(1, SATISFACTION_COLS.SUMMARY_START, 1, sectionHeaders.length)
    .setValues([sectionHeaders])
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE)
    .setWrap(true);

  // Set widths for section columns
  for (var sc = 72; sc <= 82; sc++) {
    sheet.setColumnWidth(sc, 65);
  }

  // Section averages (columns BT-CD) are computed by syncSatisfactionValues()
  // No formulas in visible sheet - values are written by JavaScript

  // ═══════════════════════════════════════════════════════════════════════════
  // DASHBOARD / CHART DATA AREA (Columns CF onwards - col 84+)
  // Summary metrics for creating charts
  // ═══════════════════════════════════════════════════════════════════════════

  var dashStart = 84; // Column CF

  // Dashboard Header
  sheet.getRange(1, dashStart).setValue('📊 SURVEY DASHBOARD')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#4285F4')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(1, dashStart, 1, 4).merge();

  // Response Summary
  sheet.getRange(3, dashStart).setValue('📈 RESPONSE SUMMARY')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(3, dashStart, 1, 2).merge();

  // Response summary labels (values populated by syncSatisfactionValues)
  var responseSummary = [
    ['Total Responses', ''],
    ['Response Period', ''],
    ['', ''],
    ['📊 SECTION SCORES', ''],
    ['Section', 'Avg Score'],
    ['Overall Satisfaction', ''],
    ['Steward Rating', ''],
    ['Steward Access', ''],
    ['Chapter Effectiveness', ''],
    ['Local Leadership', ''],
    ['Contract Enforcement', ''],
    ['Representation', ''],
    ['Communication', ''],
    ['Member Voice', ''],
    ['Value & Action', ''],
    ['Scheduling', '']
  ];
  sheet.getRange(4, dashStart, responseSummary.length, 2).setValues(responseSummary);

  // Format headers
  sheet.getRange(7, dashStart, 1, 2).setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(8, dashStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');

  // Column widths for dashboard
  sheet.setColumnWidth(dashStart, 170);
  sheet.setColumnWidth(dashStart + 1, 100);

  // ═══════════════════════════════════════════════════════════════════════════
  // DEMOGRAPHICS BREAKDOWN (Column CH onwards)
  // ═══════════════════════════════════════════════════════════════════════════

  var demoStart = dashStart + 3; // Column CH (87)

  sheet.getRange(3, demoStart).setValue('👥 DEMOGRAPHICS')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(3, demoStart, 1, 2).merge();

  // Demographics labels (values populated by syncSatisfactionValues)
  var demographics = [
    ['Shift Breakdown', ''],
    ['Day', ''],
    ['Evening', ''],
    ['Night', ''],
    ['Rotating', ''],
    ['', ''],
    ['Tenure', ''],
    ['<1 year', ''],
    ['1-3 years', ''],
    ['4-7 years', ''],
    ['8-15 years', ''],
    ['15+ years', ''],
    ['', ''],
    ['Steward Contact', ''],
    ['Yes (12 mo)', ''],
    ['No', ''],
    ['', ''],
    ['Filed Grievance', ''],
    ['Yes (24 mo)', ''],
    ['No', '']
  ];
  sheet.getRange(4, demoStart, demographics.length, 2).setValues(demographics);

  // Format demographic headers
  sheet.getRange(4, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(10, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(17, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(21, demoStart, 1, 2).setFontWeight('bold').setBackground('#E8F0FE');

  sheet.setColumnWidth(demoStart, 120);
  sheet.setColumnWidth(demoStart + 1, 60);

  // ═══════════════════════════════════════════════════════════════════════════
  // CHART DATA TABLE (for bar/column charts)
  // ═══════════════════════════════════════════════════════════════════════════

  var chartStart = dashStart + 6; // Column CK (90)

  sheet.getRange(3, chartStart).setValue('📉 CHART DATA (Select for Charts)')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange(3, chartStart, 1, 2).merge();

  // Chart data labels (values populated by syncSatisfactionValues)
  var chartData = [
    ['Section', 'Score'],
    ['Overall Satisfaction', 0],
    ['Steward Rating', 0],
    ['Steward Access', 0],
    ['Chapter', 0],
    ['Leadership', 0],
    ['Contract', 0],
    ['Representation', 0],
    ['Communication', 0],
    ['Member Voice', 0],
    ['Value & Action', 0],
    ['Scheduling', 0]
  ];
  sheet.getRange(4, chartStart, chartData.length, 2).setValues(chartData);

  sheet.getRange(4, chartStart, 1, 2).setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.setColumnWidth(chartStart, 140);
  sheet.setColumnWidth(chartStart + 1, 60);

  // Add border around chart data for easy selection
  sheet.getRange(4, chartStart, chartData.length, 2).setBorder(true, true, true, true, false, false);

  // ═══════════════════════════════════════════════════════════════════════════
  // GOOGLE FORM SETUP INSTRUCTIONS
  // ═══════════════════════════════════════════════════════════════════════════

  var instrStart = 26; // Row 26

  sheet.getRange(instrStart, dashStart).setValue('🔗 GOOGLE FORM SETUP')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#4285F4')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(instrStart, dashStart, 1, 4).merge();

  var instructions = [
    ['1. Create Form:', 'https://docs.google.com/forms/create'],
    ['2. Add 68 questions per outline below', ''],
    ['3. Link form to this sheet:', ''],
    ['   - In Form: Responses tab → Link to Sheets', ''],
    ['   - Select this spreadsheet & sheet', ''],
    ['4. Form responses auto-populate cols A-BP', ''],
    ['5. Section averages auto-calculate in BT-CD', ''],
    ['6. Create charts from Chart Data (col CK-CL)', '']
  ];

  for (var ins = 0; ins < instructions.length; ins++) {
    sheet.getRange(instrStart + 1 + ins, dashStart).setValue(instructions[ins][0]);
    if (instructions[ins][1]) {
      sheet.getRange(instrStart + 1 + ins, dashStart + 1).setValue(instructions[ins][1])
        .setFontColor('#1155CC');
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // 68-QUESTION SURVEY OUTLINE (Reference for Google Form creation)
  // ═══════════════════════════════════════════════════════════════════════════

  var surveyStartRow = 38;

  sheet.getRange(surveyStartRow, dashStart).setValue('📋 68-QUESTION SURVEY OUTLINE')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#34A853')
    .setFontColor(COLORS.WHITE);
  sheet.getRange(surveyStartRow, dashStart, 1, 5).merge();

  // Survey sections with all 68 questions
  var surveyOutline = [
    ['SECTION', 'Q#', 'QUESTION', 'TYPE', 'OPTIONS'],
    ['', '', '', '', ''],
    ['INTRO', '', 'Union Member Satisfaction Survey', 'Description', 'Anonymous, reported in aggregate'],
    ['', '', '', '', ''],
    ['1: WORK CONTEXT', '1', 'Worksite / Program / Region', 'Dropdown', '(Your worksites)'],
    ['', '2', 'Role / Job Group', 'Dropdown', '(Your roles)'],
    ['', '3', 'Shift', 'Multiple choice', 'Day | Evening | Night | Rotating'],
    ['', '4', 'Time in current role', 'Multiple choice', '<1 yr | 1-3 yrs | 4-7 yrs | 8-15 yrs | 15+ yrs'],
    ['', '5', 'Contact with steward in past 12 months?', 'MC + Branch', 'Yes → 3A | No → 3B'],
    ['', '', '', '', ''],
    ['2: OVERALL SAT', '6', 'Satisfied with union representation', '1-10 scale', ''],
    ['', '7', 'Trust union to act in best interests', '1-10 scale', ''],
    ['', '8', 'Feel more protected at work', '1-10 scale', ''],
    ['', '9', 'Would recommend membership to coworker', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['3A: STEWARD (Contact)', '10', 'Responded in timely manner', '1-10 scale', ''],
    ['', '11', 'Treated me with respect', '1-10 scale', ''],
    ['', '12', 'Explained options clearly', '1-10 scale', ''],
    ['', '13', 'Followed through on commitments', '1-10 scale', ''],
    ['', '14', 'Advocated effectively', '1-10 scale', ''],
    ['', '15', 'Felt safe raising concerns', '1-10 scale', ''],
    ['', '16', 'Handled confidentiality appropriately', '1-10 scale', ''],
    ['', '17', 'What should stewards improve?', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['3B: STEWARD ACCESS', '18', 'Know how to contact steward/rep', '1-10 scale', ''],
    ['', '19', 'Confident I would get help', '1-10 scale', ''],
    ['', '20', 'Easy to figure out who to contact', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['4: CHAPTER', '21', 'Reps understand my workplace issues', '1-10 scale', ''],
    ['', '22', 'Chapter communication is regular and clear', '1-10 scale', ''],
    ['', '23', 'Chapter organizes members effectively', '1-10 scale', ''],
    ['', '24', 'Know how to reach chapter contact', '1-10 scale', ''],
    ['', '25', 'Representation is fair across roles/shifts', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['5: LEADERSHIP', '26', 'Leadership communicates decisions clearly', '1-10 scale', ''],
    ['', '27', 'Understand how decisions are made', '1-10 scale', ''],
    ['', '28', 'Union is transparent about finances', '1-10 scale', ''],
    ['', '29', 'Leadership is accountable to feedback', '1-10 scale', ''],
    ['', '30', 'Internal processes feel fair', '1-10 scale', ''],
    ['', '31', 'Union welcomes differing opinions', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['6: CONTRACT', '32', 'Union enforces contract effectively', '1-10 scale', ''],
    ['', '33', 'Communicates realistic timelines', '1-10 scale', ''],
    ['', '34', 'Provides clear updates on issues', '1-10 scale', ''],
    ['', '35', 'Prioritizes frontline conditions', '1-10 scale', ''],
    ['', '36', 'Filed grievance in past 24 months?', 'MC + Branch', 'Yes → 6A | No → 7'],
    ['', '', '', '', ''],
    ['6A: REPRESENTATION', '37', 'Understood steps and timeline', '1-10 scale', ''],
    ['', '38', 'Felt supported throughout', '1-10 scale', ''],
    ['', '39', 'Received updates often enough', '1-10 scale', ''],
    ['', '40', 'Outcome feels justified', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['7: COMMUNICATION', '41', 'Communications are clear and actionable', '1-10 scale', ''],
    ['', '42', 'Receive enough information', '1-10 scale', ''],
    ['', '43', 'Can find information easily', '1-10 scale', ''],
    ['', '44', 'Communications reach all shifts/locations', '1-10 scale', ''],
    ['', '45', 'Meetings are worth attending', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['8: MEMBER VOICE', '46', 'My voice matters in the union', '1-10 scale', ''],
    ['', '47', 'Union actively seeks input', '1-10 scale', ''],
    ['', '48', 'Members treated with dignity', '1-10 scale', ''],
    ['', '49', 'Newer members are supported', '1-10 scale', ''],
    ['', '50', 'Internal conflict handled respectfully', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['9: VALUE & ACTION', '51', 'Union provides good value for dues', '1-10 scale', ''],
    ['', '52', 'Priorities reflect member needs', '1-10 scale', ''],
    ['', '53', 'Union prepared to mobilize', '1-10 scale', ''],
    ['', '54', 'Understand how to get involved', '1-10 scale', ''],
    ['', '55', 'Acting together, we can win improvements', '1-10 scale', ''],
    ['', '', '', '', ''],
    ['10: SCHEDULING', '56', 'Understand proposed changes', '1-10 scale', ''],
    ['', '57', 'Feel adequately informed', '1-10 scale', ''],
    ['', '58', 'Decisions use clear criteria', '1-10 scale', ''],
    ['', '59', 'Work can be done under expectations', '1-10 scale', ''],
    ['', '60', 'Approach supports effective outcomes', '1-10 scale', ''],
    ['', '61', 'Approach supports my wellbeing', '1-10 scale', ''],
    ['', '62', 'My concerns would be taken seriously', '1-10 scale', ''],
    ['', '63', 'Biggest scheduling challenge?', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['11: PRIORITIES', '64', 'Top 3 priorities (6-12 mo)', 'Checkboxes', 'Contract | Workload | Scheduling | Pay | Safety | Training | Equity | Comm | Steward | Organizing | Other'],
    ['', '65', '#1 change union should make', 'Paragraph', ''],
    ['', '66', 'One thing union should keep doing', 'Paragraph', ''],
    ['', '67', 'Additional comments (no names)', 'Paragraph', 'Optional'],
    ['', '', '', '', ''],
    ['BRANCHING:', '', 'Q5: Yes → Section 3A, No → Section 3B', '', ''],
    ['', '', 'Q36: Yes → Section 6A, No → Section 7', '', '']
  ];

  sheet.getRange(surveyStartRow + 1, dashStart, surveyOutline.length, 5).setValues(surveyOutline);

  // Format survey outline
  sheet.getRange(surveyStartRow + 1, dashStart, 1, 5)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Format section headers
  for (var so = surveyStartRow + 2; so <= surveyStartRow + surveyOutline.length; so++) {
    var sectionVal = sheet.getRange(so, dashStart).getValue();
    if (sectionVal && sectionVal.toString().includes(':')) {
      sheet.getRange(so, dashStart, 1, 5).setBackground('#E8F0FE').setFontWeight('bold');
    }
  }

  // Column widths for survey outline
  sheet.setColumnWidth(dashStart + 2, 280);  // Question column
  sheet.setColumnWidth(dashStart + 3, 80);   // Type
  sheet.setColumnWidth(dashStart + 4, 200);  // Options

  // Delete excess columns after CJ (column 88)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 88) {
    sheet.deleteColumns(89, maxCols - 88);
  }

  // Populate computed values (no formulas in visible sheet)
  syncSatisfactionValues();

  Logger.log('Member Satisfaction sheet created with 68-question survey, dashboard, and chart data');
}

// ============================================================================
// FEEDBACK & DEVELOPMENT SHEET
// ============================================================================

/**
 * Create the Feedback & Development tracking sheet
 * Tracks bugs, feature requests, and system improvements
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function createFeedbackSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.FEEDBACK);
  sheet.clear();

  // Headers
  var headers = [
    'Timestamp',       // A - Auto-generated
    'Submitted By',    // B
    'Category',        // C
    'Type',            // D
    'Priority',        // E
    'Title',           // F
    'Description',     // G
    'Status',          // H
    'Assigned To',     // I
    'Resolution',      // J
    'Notes'            // K
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);

  // Column widths
  sheet.setColumnWidth(FEEDBACK_COLS.TIMESTAMP, 140);
  sheet.setColumnWidth(FEEDBACK_COLS.SUBMITTED_BY, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.CATEGORY, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.TYPE, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.PRIORITY, 80);
  sheet.setColumnWidth(FEEDBACK_COLS.TITLE, 200);
  sheet.setColumnWidth(FEEDBACK_COLS.DESCRIPTION, 350);
  sheet.setColumnWidth(FEEDBACK_COLS.STATUS, 100);
  sheet.setColumnWidth(FEEDBACK_COLS.ASSIGNED_TO, 120);
  sheet.setColumnWidth(FEEDBACK_COLS.RESOLUTION, 250);
  sheet.setColumnWidth(FEEDBACK_COLS.NOTES, 200);

  // Category dropdown
  var categoryOptions = ['Dashboard', 'Member Directory', 'Grievance Log', 'Config', 'Search', 'Mobile', 'Reports', 'Performance', 'UI/UX', 'Other'];
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.CATEGORY, 998, 1).setDataValidation(categoryRule);

  // Type dropdown
  var typeOptions = ['Bug', 'Feature Request', 'Improvement', 'Documentation', 'Question'];
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.TYPE, 998, 1).setDataValidation(typeRule);

  // Priority dropdown with conditional formatting
  var priorityOptions = ['Low', 'Medium', 'High', 'Critical'];
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(priorityOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.PRIORITY, 998, 1).setDataValidation(priorityRule);

  // Status dropdown
  var statusOptions = ['New', 'In Progress', 'On Hold', 'Resolved', 'Won\'t Fix', 'Duplicate'];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, FEEDBACK_COLS.STATUS, 998, 1).setDataValidation(statusRule);

  // Conditional formatting for Priority
  var priorityRange = sheet.getRange(2, FEEDBACK_COLS.PRIORITY, 998, 1);

  // Critical = Red
  var criticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Critical')
    .setBackground('#FFCDD2')
    .setFontColor('#B71C1C')
    .setRanges([priorityRange])
    .build();

  // High = Orange
  var highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('High')
    .setBackground('#FFE0B2')
    .setFontColor('#E65100')
    .setRanges([priorityRange])
    .build();

  // Medium = Yellow
  var mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#FFF9C4')
    .setFontColor('#F57F17')
    .setRanges([priorityRange])
    .build();

  // Low = Green
  var lowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Low')
    .setBackground('#C8E6C9')
    .setFontColor('#1B5E20')
    .setRanges([priorityRange])
    .build();

  // Conditional formatting for Status
  var statusRange = sheet.getRange(2, FEEDBACK_COLS.STATUS, 998, 1);

  // Resolved = Green
  var resolvedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Resolved')
    .setBackground('#C8E6C9')
    .setFontColor('#1B5E20')
    .setRanges([statusRange])
    .build();

  // In Progress = Blue
  var inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground('#BBDEFB')
    .setFontColor('#0D47A1')
    .setRanges([statusRange])
    .build();

  // Apply all conditional formatting rules
  var rules = sheet.getConditionalFormatRules();
  rules.push(criticalRule, highRule, mediumRule, lowRule, resolvedRule, inProgressRule);
  sheet.setConditionalFormatRules(rules);

  // Timestamp format
  sheet.getRange(2, FEEDBACK_COLS.TIMESTAMP, 998, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  // ═══════════════════════════════════════════════════════════════════════════
  // SUMMARY METRICS SECTION
  // ═══════════════════════════════════════════════════════════════════════════

  sheet.getRange('M1').setValue('📊 FEEDBACK METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('M1:O1').merge();

  // Feedback metrics labels (values populated by syncFeedbackValues)
  var feedbackMetrics = [
    ['Metric', 'Value', 'Description'],
    ['Total Items', 0, 'All feedback items'],
    ['Bugs', 0, 'Bug reports'],
    ['Feature Requests', 0, 'New feature asks'],
    ['Improvements', 0, 'Enhancement suggestions'],
    ['New/Open', 0, 'Unresolved items'],
    ['Resolved', 0, 'Completed items'],
    ['Critical Priority', 0, 'Urgent items'],
    ['Resolution Rate', '0%', 'Percentage resolved']
  ];
  sheet.getRange(2, 13, feedbackMetrics.length, 3).setValues(feedbackMetrics);

  // Format metrics header
  sheet.getRange('M2:O2').setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.setColumnWidth(13, 140);
  sheet.setColumnWidth(14, 80);
  sheet.setColumnWidth(15, 150);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Delete excess columns after O (column 15)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 15) {
    sheet.deleteColumns(16, maxCols - 15);
  }

  // Populate computed values (no formulas in visible sheet)
  syncFeedbackValues();

  Logger.log('Feedback & Development sheet created');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get existing sheet or create new one
 * WARNING: This function deletes existing sheets without confirmation.
 * For user-facing operations, use getOrCreateSheetSafe() from DataIntegrity.gs instead.
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name
 * @returns {Sheet} Sheet object
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  return ss.insertSheet(name);
}

/**
 * Get existing sheet or create new one with user confirmation
 * Shows a warning dialog before deleting sheets that contain data.
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name
 * @param {boolean} skipConfirmation - Skip the confirmation dialog (for batch operations)
 * @returns {Sheet|null} Sheet object or null if user cancelled
 */
function getOrCreateSheetWithConfirmation(ss, name, skipConfirmation) {
  var sheet = ss.getSheetByName(name);

  if (sheet) {
    // Check if sheet has data beyond the header
    var hasData = sheet.getLastRow() > 1;

    if (hasData && !skipConfirmation) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        '⚠️ Sheet Contains Data',
        'The sheet "' + name + '" already exists and contains data.\n\n' +
        'All existing data in this sheet will be PERMANENTLY DELETED.\n\n' +
        'Do you want to continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        Logger.log('User cancelled recreation of sheet: ' + name);
        return null;
      }
    }

    ss.deleteSheet(sheet);
    Logger.log('Deleted existing sheet: ' + name);
  }

  return ss.insertSheet(name);
}

/**
 * Setup hidden calculation sheets for cross-sheet data sync
 * Calls the full implementation in HiddenSheets.gs
 */
function setupHiddenSheets(ss) {
  // Call the full hidden sheet setup from HiddenSheets.gs
  setupAllHiddenSheets();

  // Install the auto-sync trigger
  installAutoSyncTrigger();
}

// ============================================================================
// DATA VALIDATION
// ============================================================================

/**
 * Setup all data validations for Member Directory and Grievance Log
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

  // Member Directory Validations
  // Single-select dropdowns (strict validation)
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

  // Multi-select dropdowns (allow comma-separated values)
  setMultiSelectValidation(memberSheet, MEMBER_COLS.OFFICE_DAYS, configSheet, CONFIG_COLS.OFFICE_DAYS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.PREFERRED_COMM, configSheet, CONFIG_COLS.COMM_METHODS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.BEST_TIME, configSheet, CONFIG_COLS.BEST_TIMES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.COMMITTEES, configSheet, CONFIG_COLS.STEWARD_COMMITTEES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.ASSIGNED_STEWARD, configSheet, CONFIG_COLS.STEWARDS);

  // Grievance Log Validations
  // Note: Member ID does NOT have dropdown - allows free text entry for flexibility
  // setMemberIdValidation(grievanceSheet, memberSheet);  // REMOVED: Member ID should not have dropdown

  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.STATUS, configSheet, CONFIG_COLS.GRIEVANCE_STATUS);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.CURRENT_STEP, configSheet, CONFIG_COLS.GRIEVANCE_STEP);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ISSUE_CATEGORY, configSheet, CONFIG_COLS.ISSUE_CATEGORY);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ARTICLES, configSheet, CONFIG_COLS.ARTICLES);

  // Note: Unit, Location, Steward columns now use formulas for auto-lookup
  // No need for manual dropdown validation on those columns

  SpreadsheetApp.getActiveSpreadsheet().toast('Data validations applied successfully!', '✅ Success', 3);
}

/**
 * Set Member ID validation dropdown from Member Directory
 * @param {Sheet} grievanceSheet - Grievance Log sheet
 * @param {Sheet} memberSheet - Member Directory sheet
 */
function setMemberIdValidation(grievanceSheet, memberSheet) {
  // Get the Member ID column from Member Directory
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
 * Set dropdown validation from Config sheet
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setDropdownValidation(targetSheet, targetCol, configSheet, sourceCol) {
  // Config data starts at row 3 (row 1 = section headers, row 2 = column headers)
  // Use dynamic row count instead of fixed 100 to handle large Config lists
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  // Use at least 10 rows, or actual count + buffer for growth
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
 * Set multi-select validation (allows comma-separated values)
 * Shows dropdown for convenience but accepts any text
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setMultiSelectValidation(targetSheet, targetCol, configSheet, sourceCol) {
  // Config data starts at row 3
  // Use dynamic row count instead of fixed 100 to handle large Config lists
  var configLastRow = configSheet.getLastRow();
  var configData = configSheet.getRange(3, sourceCol, Math.max(1, configLastRow - 2), 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  // Use at least 10 rows, or actual count + buffer for growth
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

// Store the target cell for multi-select dialog
var multiSelectTarget_ = null;

/**
 * Show multi-select dialog for the current cell
 * Called from menu or double-click on multi-select column
 */
// Note: showMultiSelectDialog() defined in modular file - see respective module

/**
 * Get values from a Config sheet column
 * Note: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
 * @param {Sheet} configSheet - The Config sheet
 * @param {number} col - Column number
 * @returns {Array} Array of non-empty values
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
 * Apply the multi-select value to the stored cell
 * Called from the dialog
 * @param {string} value - Comma-separated selected values
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

  // Clear stored properties
  props.deleteProperty('multiSelectRow');
  props.deleteProperty('multiSelectCol');
}

/**
 * Handle edit events to trigger multi-select dialog
 * This is installed as an onEdit trigger
 */
function onEditMultiSelect(e) {
  // Only process single cell edits
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only Member Directory
  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Skip header row
  if (row < 2) return;

  // Check if this is a multi-select column
  var config = getMultiSelectConfig(col);
  if (!config) return;

  // If user typed something, show the dialog to help them select properly
  // Only trigger if the new value isn't already comma-separated (user might be pasting)
  var newValue = e.value || '';
  var oldValue = e.oldValue || '';

  // If user cleared the cell or pasted valid data, don't interrupt
  if (newValue === '' || newValue.indexOf(',') !== -1) return;

  // Show helpful toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tip: Use Dashboard menu > "Multi-Select Editor" for easier selection of multiple values.',
    config.label,
    5
  );
}

/**
 * Handle selection change to auto-open multi-select dialog
 * This is installed as an onSelectionChange trigger
 */
function onSelectionChangeMultiSelect(e) {
  // Only process if we have a valid range
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only Member Directory
  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Skip header row and multi-cell selections
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  // Check if this is a multi-select column
  var config = getMultiSelectConfig(col);
  if (!config) return;

  // Check if we already showed dialog for this cell (avoid repeated opens)
  var props = PropertiesService.getDocumentProperties();
  var lastCell = props.getProperty('lastMultiSelectCell');
  var currentCell = row + ',' + col;

  if (lastCell === currentCell) return;

  // Store current cell
  props.setProperty('lastMultiSelectCell', currentCell);

  // Auto-open the multi-select dialog
  showMultiSelectDialog();
}

/**
 * Install the multi-select auto-open trigger
 * Run this once to enable auto-open on cell selection
 */
function installMultiSelectTrigger() {
  var ui = SpreadsheetApp.getUi();

  // Note: onSelectionChange triggers cannot be created programmatically
  // User must set this up manually in Apps Script editor
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
 * Remove the multi-select auto-open trigger
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

// ============================================================================
// DIAGNOSE FUNCTION
// ============================================================================

/**
 * System health check - validates sheets and column counts
 */
// Note: DIAGNOSE_SETUP() defined in modular file - see respective module

// ============================================================================
// REPAIR FUNCTION
// ============================================================================

/**
 * Repair dashboard - recreates hidden sheets, triggers, and syncs data
 */
// Note: REPAIR_DASHBOARD() defined in modular file - see respective module

/**
 * Creates the Menu Checklist sheet with all menu items
 * Called automatically during dashboard repair/creation
 * @private
 */
function createFunctionChecklistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.FUNCTION_CHECKLIST || 'Menu Checklist';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Menu items organized by optimal testing order: [Phase, Menu, Item, Function, Description]
  var menuItems = [
    // ═══ PHASE 1: Foundation & Setup (Test these first!) ═══
    ['1️⃣ Foundation', '🏗️ Setup', '🔧 REPAIR DASHBOARD', 'REPAIR_DASHBOARD', 'Repairs all hidden sheets, reapplies formulas, fixes broken references'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 DIAGNOSE SETUP', 'DIAGNOSE_SETUP', 'Checks sheet structure, triggers, and configuration for issues'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 Verify Hidden Sheets', 'verifyHiddenSheets', 'Validates all 6 hidden calculation sheets exist and have correct formulas'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets', 'Creates/recreates all hidden sheets with self-healing formulas'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets', 'Fixes broken formulas in hidden sheets without recreating them'],
    ['1️⃣ Foundation', '🏗️ Setup', '⚙️ Setup Data Validations', 'setupDataValidations', 'Applies dropdown validations to Member Directory and Grievance Log'],
    ['1️⃣ Foundation', '🏗️ Setup', '🎨 Setup Comfort View', 'setupADHDDefaults', 'Configures default accessibility-friendly visual settings'],

    // ═══ PHASE 2: Triggers & Data Sync ═══
    ['2️⃣ Sync', '⚙️ Admin > Setup', '⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger', 'Creates edit trigger to auto-sync data between sheets'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync All Data Now', 'syncAllData', 'Manually syncs all data between Member Directory and Grievance Log'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory', 'Updates Member Directory with grievance counts and status'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog', 'Updates Grievance Log with member names and contact info'],
    ['2️⃣ Sync', '⚙️ Admin > Setup', '🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger', 'Removes the automatic sync trigger (manual sync still works)'],

    // ═══ PHASE 3: Core Dashboards ═══
    ['3️⃣ Dashboards', '👤 Dashboard', '🎯 Custom View', 'showInteractiveDashboardTab', 'Opens the Custom View sheet with configurable metrics'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📋 View Active Grievances', 'viewActiveGrievances', 'Shows filtered list of all open/pending grievances'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Mobile Dashboard', 'showMobileDashboard', 'Touch-friendly dashboard for phones and tablets'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Get Mobile App URL', 'showWebAppUrl', 'Displays the web app URL for mobile bookmarking'],
    ['3️⃣ Dashboards', '👤 Dashboard', '⚡ Quick Actions', 'showQuickActionsMenu', 'Popup menu for common actions (add member, new grievance, etc.)'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '📊 Rebuild Dashboard', 'rebuildDashboard', 'Recreates the Dashboard sheet with fresh formulas'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '📈 Refresh Interactive Charts', 'refreshInteractiveCharts', 'Updates all charts in Custom View with current data'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '🔄 Refresh All Formulas', 'refreshAllFormulas', 'Recalculates all formulas across all sheets'],

    // ═══ PHASE 4: Search ═══
    ['4️⃣ Search', '🔍 Search', '🔍 Search Members', 'searchMembers', 'Opens search dialog to find members by name, ID, email, or location'],

    // ═══ PHASE 5: Grievance Management ═══
    ['5️⃣ Grievances', '👤 Grievance Tools', '➕ Start New Grievance', 'startNewGrievance', 'Opens form to create new grievance with auto-generated ID'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched', 'Recalculates deadline and status formulas for all grievances'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas', 'Updates calculated columns in Member Directory'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔗 Setup Live Grievance Links', 'setupLiveGrievanceFormulas', 'Creates formulas linking grievances to member data'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '👤 Clear Member ID Validation', 'setupGrievanceMemberDropdown', 'Removes dropdown from Member ID column to allow free text entry'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔧 Fix Overdue Text Data', 'fixOverdueTextToNumbers', 'Converts text dates to proper date format for calculations'],

    // ═══ PHASE 6: Google Drive ═══
    ['6️⃣ Drive', '📊 Google Drive', '📁 Setup Folder for Grievance', 'setupDriveFolderForGrievance', 'Creates organized folder structure for grievance documents'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 View Grievance Files', 'showGrievanceFiles', 'Shows all files associated with selected grievance'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 Batch Create Folders', 'batchCreateGrievanceFolders', 'Creates folders for multiple grievances at once'],

    // ═══ PHASE 7: Calendar ═══
    ['7️⃣ Calendar', '📊 Calendar', '📅 Sync Deadlines to Calendar', 'syncDeadlinesToCalendar', 'Adds grievance deadlines to Google Calendar with reminders'],
    ['7️⃣ Calendar', '📊 Calendar', '📅 View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar', 'Shows next 30 days of deadlines from calendar'],
    ['7️⃣ Calendar', '📊 Calendar', '🗑️ Clear Calendar Events', 'clearAllCalendarEvents', 'Removes all grievance events from calendar (use with caution)'],

    // ═══ PHASE 8: Notifications ═══
    ['8️⃣ Notify', '📊 Notifications', '⚙️ Notification Settings', 'showNotificationSettings', 'Configure email notification preferences and timing'],
    ['8️⃣ Notify', '📊 Notifications', '🧪 Test Notifications', 'testDeadlineNotifications', 'Sends test email to verify notification setup'],

    // ═══ PHASE 9: Accessibility & Theming ═══
    ['9️⃣ Access', '♿ Comfort View', '♿ Comfort View Panel', 'showADHDControlPanel', 'Central hub for all accessibility-friendly features and settings'],
    ['9️⃣ Access', '♿ Comfort View', '🎯 Focus Mode', 'activateFocusMode', 'Highlights current row, dims distractions, reduces visual noise'],
    ['9️⃣ Access', '♿ Comfort View', '🔲 Toggle Zebra Stripes', 'toggleZebraStripes', 'Alternating row colors for easier row tracking'],
    ['9️⃣ Access', '♿ Comfort View', '📝 Quick Capture', 'showQuickCaptureNotepad', 'Fast notepad for capturing thoughts without losing focus'],
    ['9️⃣ Access', '♿ Comfort View', '🍅 Pomodoro Timer', 'startPomodoroTimer', '25-minute focus timer with break reminders'],
    ['9️⃣ Access', '🔧 Theming', '🎨 Theme Manager', 'showThemeManager', 'Choose from preset themes or customize colors'],
    ['9️⃣ Access', '🔧 Theming', '🌙 Toggle Dark Mode', 'quickToggleDarkMode', 'Switch between light and dark color schemes'],
    ['9️⃣ Access', '🔧 Theming', '🔄 Reset Theme', 'resetToDefaultTheme', 'Restores default purple/green color scheme'],

    // ═══ PHASE 10: Productivity Tools ═══
    ['🔟 Tools', '🔧 Multi-Select', '📝 Open Editor', 'showMultiSelectDialog', 'Select multiple values for multi-select columns'],
    ['🔟 Tools', '🔧 Multi-Select', '⚡ Enable Auto-Open', 'installMultiSelectTrigger', 'Auto-opens multi-select dialog when clicking multi-select cells'],
    ['🔟 Tools', '🔧 Multi-Select', '🚫 Disable Auto-Open', 'removeMultiSelectTrigger', 'Stops auto-opening multi-select dialog'],

    // ═══ PHASE 11: Performance & Cache ═══
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗄️ Cache Status', 'showCacheStatusDashboard', 'Shows what data is cached and cache hit/miss rates'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🔥 Warm Up Caches', 'warmUpCaches', 'Pre-loads frequently used data into cache for faster access'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗑️ Clear All Caches', 'invalidateAllCaches', 'Clears all cached data (forces fresh data on next load)'],

    // ═══ PHASE 12: Validation ═══
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🔍 Run Bulk Validation', 'runBulkValidation', 'Checks all data for errors, duplicates, and missing values'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚙️ Validation Settings', 'showValidationSettings', 'Configure which validations run and error thresholds'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🧹 Clear Indicators', 'clearValidationIndicators', 'Removes error highlighting from cells'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚡ Install Validation Trigger', 'installValidationTrigger', 'Enables automatic validation on data entry'],

    // ═══ PHASE 13: Testing (Run last to verify everything) ═══
    ['1️⃣3️⃣ Test', '🧪 Testing', '🧪 Run All Tests', 'runAllTests', 'Executes full test suite for all functions (takes 2-3 minutes)'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '⚡ Run Quick Tests', 'runQuickTests', 'Runs essential tests only (30 seconds)'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '📊 View Test Results', 'viewTestResults', 'Shows results from last test run with pass/fail details']
  ];

  // Build rows with header
  var rows = [['✓', 'Phase', 'Menu', 'Item', 'Function', 'Description', 'Notes']];
  for (var i = 0; i < menuItems.length; i++) {
    rows.push([false, menuItems[i][0], menuItems[i][1], menuItems[i][2], menuItems[i][3], menuItems[i][4], '']);
  }

  // Write all data
  sheet.getRange(1, 1, rows.length, 7).setValues(rows);

  // Format header
  sheet.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE || '#7C3AED')
    .setFontColor(COLORS.WHITE || '#FFFFFF')
    .setHorizontalAlignment('center');

  // Add checkboxes
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  }

  // Set column widths
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 350);
  sheet.setColumnWidth(7, 250);

  // Freeze header
  sheet.setFrozenRows(1);

  // Alternating colors
  for (var r = 2; r <= rows.length; r++) {
    if (r % 2 === 0) {
      sheet.getRange(r, 1, 1, 7).setBackground('#F9FAFB');
    }
  }

  // Conditional formatting for checked items
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setBackground('#E8F5E9')
    .setRanges([sheet.getRange(2, 1, rows.length - 1, 7)])
    .build();
  sheet.setConditionalFormatRules([rule]);

  // Delete excess columns after G (column 7)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 7) {
    sheet.deleteColumns(8, maxCols - 7);
  }

  return sheet;
}

/**
 * Creates the Getting Started sheet with setup instructions
 */
function createGettingStartedSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.GETTING_STARTED || '📚 Getting Started';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define colors
  var headerBg = '#7C3AED';       // Purple header
  var sectionBg = '#F3E8FF';      // Light purple section
  var stepBg = '#ECFDF5';         // Light green for steps
  var tipBg = '#FEF3C7';          // Light yellow for tips
  var textColor = '#1F2937';
  var white = '#FFFFFF';

  var row = 1;

  // ═══ MAIN HEADER ═══
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📚 GETTING STARTED WITH 509 DASHBOARD')
    .setBackground(headerBg)
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Welcome! This guide will help you set up and use the 509 Dashboard effectively.')
    .setFontSize(12)
    .setFontColor('#6B7280')
    .setHorizontalAlignment('center');

  // ═══ SECTION 1: FIRST-TIME SETUP ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('🚀 STEP 1: First-Time Setup (5 minutes)')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#7C3AED');
  sheet.setRowHeight(row, 35);

  var setupSteps = [
    ['1.1', 'Open the Config tab and customize dropdown values for your organization'],
    ['1.2', 'Add your Job Titles in Column A (e.g., Case Worker, Supervisor, Manager)'],
    ['1.3', 'Add your Office Locations in Column B'],
    ['1.4', 'Add your Steward names in Column H'],
    ['1.5', 'Run the diagnostic: Admin menu → DIAGNOSE SETUP to verify everything is working']
  ];

  for (var i = 0; i < setupSteps.length; i++) {
    row++;
    sheet.getRange(row, 1).setValue(setupSteps[i][0]).setFontWeight('bold').setFontColor('#7C3AED').setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(setupSteps[i][1]).setFontColor(textColor).setWrap(true);
    sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
  }

  // ═══ SECTION 2: ADDING MEMBERS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('👥 STEP 2: Adding Members')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#7C3AED');
  sheet.setRowHeight(row, 35);

  var memberSteps = [
    ['2.1', 'Go to the Member Directory tab'],
    ['2.2', 'Click on the first empty row (row 2 if empty)'],
    ['2.3', 'Enter a Member ID (format: MJOHN123 - M + first 2 letters of first/last name + 3 digits)'],
    ['2.4', 'Fill in First Name, Last Name, and Email (required fields)'],
    ['2.5', 'Use the dropdowns for Job Title, Location, and other fields'],
    ['2.6', 'TIP: Columns AB-AD auto-populate from Grievance Log - don\'t edit them manually!']
  ];

  for (var j = 0; j < memberSteps.length; j++) {
    row++;
    sheet.getRange(row, 1).setValue(memberSteps[j][0]).setFontWeight('bold').setFontColor('#7C3AED').setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(memberSteps[j][1]).setFontColor(textColor).setWrap(true);
    if (memberSteps[j][1].indexOf('TIP:') === 0) {
      sheet.getRange(row, 1, 1, 6).setBackground(tipBg);
    } else {
      sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
    }
  }

  // ═══ SECTION 3: FILING GRIEVANCES ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📋 STEP 3: Filing a Grievance')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#7C3AED');
  sheet.setRowHeight(row, 35);

  var grievanceSteps = [
    ['3.1', 'Go to the Grievance Log tab'],
    ['3.2', 'Enter a Grievance ID (format: GJOHN456 - G + first 2 letters of first/last name + 3 digits)'],
    ['3.3', 'Enter the Member ID (must match a member in Member Directory)'],
    ['3.4', 'Set Status to "Open" and Current Step to "Informal" or "Step I"'],
    ['3.5', 'Enter the Incident Date (when the issue occurred)'],
    ['3.6', 'Enter the Date Filed (when the grievance was submitted)'],
    ['3.7', 'The system auto-calculates: Filing Deadline, Step I Due, Days to Deadline, etc.'],
    ['3.8', 'TIP: Use Grievances menu → Sort by Status Priority to organize by urgency']
  ];

  for (var k = 0; k < grievanceSteps.length; k++) {
    row++;
    sheet.getRange(row, 1).setValue(grievanceSteps[k][0]).setFontWeight('bold').setFontColor('#7C3AED').setHorizontalAlignment('center');
    sheet.getRange(row, 2, 1, 5).merge().setValue(grievanceSteps[k][1]).setFontColor(textColor).setWrap(true);
    if (grievanceSteps[k][1].indexOf('TIP:') === 0) {
      sheet.getRange(row, 1, 1, 6).setBackground(tipBg);
    } else {
      sheet.getRange(row, 1, 1, 6).setBackground(stepBg);
    }
  }

  // ═══ SECTION 4: USING DASHBOARDS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📊 STEP 4: Using Dashboards')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#7C3AED');
  sheet.setRowHeight(row, 35);

  var dashboardInfo = [
    ['💼 Dashboard', 'Executive overview with key metrics, steward performance, and trends'],
    ['🎯 Custom View', 'Interactive popup with tabbed views - Overview, Members, Grievances, Analytics'],
    ['📊 Member Satisfaction', 'Survey results dashboard (requires linked Google Form)'],
    ['📱 Mobile Dashboard', 'Touch-friendly view for phones and tablets']
  ];

  row++;
  sheet.getRange(row, 1, 1, 2).merge().setValue('Dashboard').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 3, 1, 4).merge().setValue('Description').setFontWeight('bold').setBackground('#E5E7EB');

  for (var m = 0; m < dashboardInfo.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 2).merge().setValue(dashboardInfo[m][0]).setFontColor('#7C3AED').setFontWeight('bold');
    sheet.getRange(row, 3, 1, 4).merge().setValue(dashboardInfo[m][1]).setFontColor(textColor).setWrap(true);
  }

  // ═══ SECTION 5: MENU OVERVIEW ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('📌 MENU QUICK REFERENCE')
    .setBackground(sectionBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#7C3AED');
  sheet.setRowHeight(row, 35);

  var menuInfo = [
    ['📊 509 Dashboard', 'Main dashboards, search, quick actions, mobile access'],
    ['📋 Grievances', 'New grievances, folder management, calendar, notifications'],
    ['👁️ View', 'Comfort View settings, themes, timeline options'],
    ['⚙️ Settings', 'Repair dashboard, validations, triggers, formulas'],
    ['🔧 Admin', 'Diagnostics, testing, data sync, demo/seed functions']
  ];

  row++;
  sheet.getRange(row, 1, 1, 2).merge().setValue('Menu').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange(row, 3, 1, 4).merge().setValue('Contains').setFontWeight('bold').setBackground('#E5E7EB');

  for (var n = 0; n < menuInfo.length; n++) {
    row++;
    sheet.getRange(row, 1, 1, 2).merge().setValue(menuInfo[n][0]).setFontColor('#7C3AED').setFontWeight('bold');
    sheet.getRange(row, 3, 1, 4).merge().setValue(menuInfo[n][1]).setFontColor(textColor);
  }

  // ═══ SECTION 6: TIPS FOR SUCCESS ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('💡 TIPS FOR SUCCESS')
    .setBackground(tipBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#92400E');
  sheet.setRowHeight(row, 35);

  var tips = [
    '✓ Always use dropdowns instead of typing - this ensures consistency',
    '✓ Check the Dashboard daily to monitor deadlines and overdue items',
    '✓ Use the Search function (509 Dashboard menu) to quickly find members or grievances',
    '✓ Run DIAGNOSE SETUP (Admin menu) if something seems wrong',
    '✓ Back up your data regularly (File → Download → Excel)',
    '✓ Use the Message Alert checkbox to flag urgent grievances (they\'ll move to top)'
  ];

  for (var p = 0; p < tips.length; p++) {
    row++;
    sheet.getRange(row, 1, 1, 6).merge().setValue(tips[p]).setFontColor('#92400E').setBackground(tipBg);
  }

  // ═══ FOOTER ═══
  row += 2;
  sheet.getRange(row, 1, 1, 6).merge()
    .setValue('Need more help? Check the ❓ FAQ tab or the Config tab\'s User Guide section.')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 6) {
    sheet.deleteColumns(7, maxCols - 6);
  }

  return sheet;
}

/**
 * Creates the FAQ sheet with common questions and answers
 */
function createFAQSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.FAQ || '❓ FAQ';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define colors
  var headerBg = '#059669';       // Green header
  var questionBg = '#ECFDF5';     // Light green for questions
  var answerBg = '#FFFFFF';       // White for answers
  var categoryBg = '#D1FAE5';     // Medium green for categories
  var textColor = '#1F2937';
  var white = '#FFFFFF';

  var row = 1;

  // ═══ MAIN HEADER ═══
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('❓ FREQUENTLY ASKED QUESTIONS')
    .setBackground(headerBg)
    .setFontColor(white)
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);

  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Find answers to common questions about using the 509 Dashboard')
    .setFontSize(12)
    .setFontColor('#6B7280')
    .setHorizontalAlignment('center');

  // ═══ CATEGORY: GETTING STARTED ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('🚀 GETTING STARTED')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var gettingStartedFAQs = [
    ['Q: How do I set up the dashboard for the first time?',
     'A: Go to Admin menu → DIAGNOSE SETUP to check your system, then customize the Config tab with your organization\'s dropdown values (job titles, locations, stewards, etc.).'],
    ['Q: Can I use this with existing member data?',
     'A: Yes! You can paste member data into the Member Directory tab. Just make sure the columns match and Member IDs follow the format (MJOHN123).'],
    ['Q: How do I test the system without real data?',
     'A: Use Admin → Demo → Seed All Sample Data to generate 1,000 test members and 300 grievances. Use NUKE SEEDED DATA when done testing.']
  ];

  for (var i = 0; i < gettingStartedFAQs.length; i++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(gettingStartedFAQs[i][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(gettingStartedFAQs[i][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: MEMBER DIRECTORY ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('👥 MEMBER DIRECTORY')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var memberFAQs = [
    ['Q: What format should Member IDs use?',
     'A: Format is M + first 2 letters of first name + first 2 letters of last name + 3 random digits. Example: John Smith → MJOSM123'],
    ['Q: Why are columns AB-AD not editable?',
     'A: These columns are auto-calculated from the Grievance Log. Has Open Grievance, Grievance Status, and Days to Deadline update automatically.'],
    ['Q: How do I assign a steward to multiple members?',
     'A: Use the Assigned Steward dropdown in column P. You can select multiple stewards using the multi-select editor.'],
    ['Q: What does the "Start Grievance" checkbox do?',
     'A: Checking this opens a pre-filled grievance form for that member. The checkbox auto-resets after use.']
  ];

  for (var j = 0; j < memberFAQs.length; j++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(memberFAQs[j][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(memberFAQs[j][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: GRIEVANCES ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('📋 GRIEVANCES')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var grievanceFAQs = [
    ['Q: How are deadlines calculated?',
     'A: Based on Article 23A: Filing = Incident + 21 days, Step I = Filed + 30 days, Step II Appeal = Step I Decision + 10 days, Step II Decision = Appeal + 30 days.'],
    ['Q: What does "Message Alert" do?',
     'A: When checked, the row is highlighted yellow and moves to the top of the list when sorted. Use it to flag urgent cases.'],
    ['Q: Why does Days to Deadline show "Overdue"?',
     'A: This means the next deadline has passed. Check the Next Action Due column to see which deadline is overdue.'],
    ['Q: How do I create a folder for grievance documents?',
     'A: Select the grievance row, then go to Grievances → Drive Folders → Setup Folder. This creates a Google Drive folder with subfolders.'],
    ['Q: Can I sync deadlines to my calendar?',
     'A: Yes! Go to Grievances → Calendar → Sync Deadlines to Calendar. You\'ll need to grant calendar access the first time.']
  ];

  for (var k = 0; k < grievanceFAQs.length; k++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(grievanceFAQs[k][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(grievanceFAQs[k][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: TROUBLESHOOTING ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('🔧 TROUBLESHOOTING')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var troubleshootingFAQs = [
    ['Q: Dropdowns are empty or not working',
     'A: Check the Config tab - the corresponding column may be empty. Also run Settings → Setup Data Validations to reapply dropdowns.'],
    ['Q: Data isn\'t syncing between sheets',
     'A: Run Settings → Triggers → Install Auto-Sync Trigger. Also try Admin → Data Sync → Sync All Data Now.'],
    ['Q: The dashboard shows wrong numbers',
     'A: Try Settings → Refresh All Formulas. If issues persist, run Settings → REPAIR DASHBOARD.'],
    ['Q: I accidentally deleted data - can I undo?',
     'A: Use Ctrl+Z (or Cmd+Z on Mac) immediately. For older changes, go to File → Version history → See version history.'],
    ['Q: Menus are not appearing',
     'A: Close and reopen the spreadsheet. If still missing, go to Extensions → Apps Script and run the onOpen function manually.']
  ];

  for (var m = 0; m < troubleshootingFAQs.length; m++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(troubleshootingFAQs[m][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(troubleshootingFAQs[m][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ CATEGORY: ADVANCED ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('⚡ ADVANCED')
    .setBackground(categoryBg)
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#065F46');
  sheet.setRowHeight(row, 35);

  var advancedFAQs = [
    ['Q: How do I link a Google Form for grievances?',
     'A: Create a Google Form with matching fields, then use Admin → Setup & Triggers → Setup Grievance Form Trigger.'],
    ['Q: Can multiple people use this at the same time?',
     'A: Yes! Google Sheets supports real-time collaboration. Changes sync automatically between users.'],
    ['Q: How do I customize the deadline days?',
     'A: The default deadlines (21, 30, 10 days) are set in the Config tab columns AA-AD. You can modify these values.']
  ];

  for (var n = 0; n < advancedFAQs.length; n++) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(advancedFAQs[n][0])
      .setBackground(questionBg).setFontWeight('bold').setFontColor('#065F46').setWrap(true);
    sheet.setRowHeight(row, 30);
    row++;
    sheet.getRange(row, 1, 1, 5).merge().setValue(advancedFAQs[n][1])
      .setBackground(answerBg).setFontColor(textColor).setWrap(true);
    sheet.setRowHeight(row, 45);
  }

  // ═══ FOOTER ═══
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Can\'t find your answer? Check the 📚 Getting Started tab or ask your administrator.')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 180);

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 5) {
    sheet.deleteColumns(6, maxCols - 5);
  }

  // Freeze header
  sheet.setFrozenRows(1);

  return sheet;
}

// ============================================================================
// MENU HANDLER FUNCTIONS
// ============================================================================

/**
 * Show desktop search modal - comprehensive search for members and grievances
 * Enhanced version of mobile search with more fields and filtering options
 */
function searchMembers() {
  showDesktopSearch();
}

/**
 * Show the desktop unified search dialog
 * Optimized for larger screens with advanced filtering
 */
// Note: showDesktopSearch() defined in modular file - see respective module

/**
 * Get locations for desktop search filter dropdown
 * @returns {Array} Array of unique locations
 */
function getDesktopSearchLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = [];

  // Get locations from Member Directory
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
 * Get search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array} Array of search results
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
          row: index + 2 // 1-indexed + header row
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

  // Limit results
  return results.slice(0, 50);
}

/**
 * Navigate to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
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

function viewActiveGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * Grievance Form Configuration
 * Maps form entry IDs to Member Directory fields for pre-filling
 */
var GRIEVANCE_FORM_CONFIG = {
  // Google Form URL (viewform version for pre-filling)
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLSedX8nf_xXeLe2sCL9MpjkEEmSuSPbjn3fNxMaMNaPlD0H5lA/viewform',

  // Form field entry IDs mapped to their purpose
  FIELD_IDS: {
    MEMBER_ID: 'entry.272049116',
    MEMBER_FIRST_NAME: 'entry.736822578',
    MEMBER_LAST_NAME: 'entry.694440931',
    JOB_TITLE: 'entry.286226203',
    AGENCY_DEPARTMENT: 'entry.2025752361',
    REGION: 'entry.352196859',
    WORK_LOCATION: 'entry.413952220',
    MANAGERS: 'entry.417314483',
    MEMBER_EMAIL: 'entry.710401757',
    STEWARD_FIRST_NAME: 'entry.84740378',
    STEWARD_LAST_NAME: 'entry.1254106933',
    STEWARD_EMAIL: 'entry.732806953',
    DATE_OF_INCIDENT: 'entry.1797903534',
    ARTICLES_VIOLATED: 'entry.1969613230',
    REMEDY_SOUGHT: 'entry.1234608137',
    DATE_FILED: 'entry.361538394',
    STEP: 'entry.2060308142',
    CONFIDENTIAL_WAIVER: 'entry.473442818'
  }
};

/**
 * Personal Contact Info Form Configuration
 * Maps form entry IDs to Member Directory fields for updating member contact info
 */
var CONTACT_FORM_CONFIG = {
  // Google Form URL - members fill out blank form, data written to Member Directory on submit
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLSeOs6Kxqca85DYRF1wTP634gMNdEirZdi5mg7aUIY5q7dIfRg/viewform',

  // Form field entry IDs mapped to Member Directory columns
  FIELD_IDS: {
    FIRST_NAME: 'entry.1970622040',
    LAST_NAME: 'entry.1536025015',
    JOB_TITLE: 'entry.1856093463',
    UNIT: 'entry.290280210',
    WORK_LOCATION: 'entry.776695410',
    OFFICE_DAYS: 'entry.1779089574',           // Multi-select
    PREFERRED_COMM: 'entry.1201030790',        // Multi-select
    BEST_TIME: 'entry.1790968369',             // Multi-select
    SUPERVISOR: 'entry.781564445',
    MANAGER: 'entry.236404577',
    EMAIL: 'entry.736229769',
    PHONE: 'entry.1824028805',
    INTEREST_ALLIED: 'entry.919302622',        // Willing to support other chapters
    INTEREST_CHAPTER: 'entry.513494211',       // Willing to be active in sub-chapter
    INTEREST_LOCAL: 'entry.1902862430'         // Willing to join direct actions
  }
};

// ============================================================================
// FORM URL CONFIGURATION HELPERS
// ============================================================================

/**
 * Get form URL from Config sheet, falling back to hardcoded default
 * This allows admins to update form links without touching code
 * @param {string} formType - 'grievance', 'contact', or 'satisfaction'
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
 * Start a new grievance for a member
 * Opens pre-filled Google Form with member info from Member Directory
 * Can be triggered from Member Directory "Start Grievance" checkbox or menu
 */
// Note: startNewGrievance() defined in modular file - see respective module

/**
 * Get current user's steward info from Member Directory
 * @private
 */
function getCurrentStewardInfo_(ss) {
  var currentUserEmail = Session.getActiveUser().getEmail();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || !currentUserEmail) {
    return { firstName: '', lastName: '', email: currentUserEmail || '' };
  }

  var data = memberSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var email = data[i][MEMBER_COLS.EMAIL - 1];
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];

    if (email && email.toLowerCase() === currentUserEmail.toLowerCase() && isSteward === 'Yes') {
      return {
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: data[i][MEMBER_COLS.LAST_NAME - 1] || '',
        email: email
      };
    }
  }

  // Return email only if not found as steward
  return { firstName: '', lastName: '', email: currentUserEmail };
}

/**
 * Build pre-filled grievance form URL
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

// ============================================================================
// GRIEVANCE FORM SUBMISSION HANDLER
// ============================================================================

/**
 * Handle grievance form submission
 * This function is triggered when a grievance form is submitted.
 * It adds the grievance to the Grievance Log and creates a Drive folder.
 *
 * To set up: Run setupGrievanceFormTrigger() once, or manually add an
 * installable trigger for this function on the form.
 *
 * @param {Object} e - Form submission event object
 */
// Note: onGrievanceFormSubmit() defined in modular file - see respective module

/**
 * Get a value from form named responses
 * @private
 */
function getFormValue_(responses, fieldName) {
  if (responses[fieldName] && responses[fieldName].length > 0) {
    return responses[fieldName][0];
  }
  return '';
}

/**
 * Parse a date string from form submission
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

/**
 * Get existing grievance IDs for collision detection
 * @private
 */
function getExistingGrievanceIds_(sheet) {
  var ids = {};
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var id = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (id) {
      ids[id] = true;
    }
  }

  return ids;
}

/**
 * Create a Drive folder for a grievance from form data
 * @private
 */
function createGrievanceFolderFromData_(grievanceId, memberId, firstName, lastName) {
  try {
    // Get or create root folder
    var rootFolder = getOrCreateDashboardFolder_();

    // Create folder name: GXXX123 - FirstName LastName (MemberID)
    var memberName = ((firstName || '') + ' ' + (lastName || '')).trim() || 'Unknown';
    var folderName = grievanceId + ' - ' + memberName;
    if (memberId) {
      folderName += ' (' + memberId + ')';
    }

    // Check if folder already exists
    var folders = rootFolder.getFoldersByName(folderName);
    var folder;

    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = rootFolder.createFolder(folderName);

      // Create subfolders for organization
      folder.createFolder('📄 Documents');
      folder.createFolder('📧 Correspondence');
      folder.createFolder('📝 Notes');
    }

    // Share with grievance coordinators from Config
    shareWithCoordinators_(folder);

    return {
      id: folder.getId(),
      url: folder.getUrl()
    };

  } catch (e) {
    Logger.log('Error creating grievance folder: ' + e.message);
    return { id: '', url: '' };
  }
}

/**
 * Share folder with grievance coordinators from Config sheet
 * @private
 */
function shareWithCoordinators_(folder) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!configSheet) return;

    // Get coordinator emails from Config (column O = GRIEVANCE_COORDINATORS)
    var coordData = configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_COORDINATORS,
                                          configSheet.getLastRow() - 1, 1).getValues();

    for (var i = 0; i < coordData.length; i++) {
      var email = coordData[i][0];
      if (email && email.toString().trim() !== '') {
        try {
          folder.addEditor(email.toString().trim());
        } catch (shareError) {
          Logger.log('Could not share with ' + email + ': ' + shareError.message);
        }
      }
    }
  } catch (e) {
    Logger.log('Error sharing with coordinators: ' + e.message);
  }
}

/**
 * Set up the grievance form submission trigger
 * Run this once to enable automatic processing of form submissions
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
    ui.alert('ℹ️ Trigger Exists',
      'A grievance form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('📋 Setup Grievance Form Trigger',
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
        ui.alert('❌ Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
        return;
      }
      formId = match[1];
    } else {
      // Use configured form
      var configFormUrl = GRIEVANCE_FORM_CONFIG.FORM_URL;
      var match = configFormUrl.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        ui.alert('❌ No Form Configured',
          'No form URL provided and could not extract ID from config.\n\n' +
          'Please provide the form edit URL.',
          ui.ButtonSet.OK);
        return;
      }
      // Note: The /e/ URL is the published version, we need the actual form ID
      ui.alert('ℹ️ Form URL Needed',
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

    ui.alert('✅ Trigger Created',
      'Grievance form trigger has been set up!\n\n' +
      'When a grievance form is submitted:\n' +
      '• A new row will be added to Grievance Log\n' +
      '• A Drive folder will be created automatically\n' +
      '• Deadlines will be calculated\n' +
      '• Member Directory will be updated',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', '✅ Success', 3);

  } catch (e) {
    ui.alert('❌ Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

/**
 * Manually process a grievance from form data (for testing or re-processing)
 * Call this with test data to verify the form submission handler works
 */
function testGrievanceFormSubmission() {
  var testEvent = {
    namedValues: {
      'Member ID': ['TEST001'],
      'Member First Name': ['Test'],
      'Member Last Name': ['Member'],
      'Job Title': ['Test Position'],
      'Agency/Department': ['Test Unit'],
      'Region': ['Test Location'],
      'Work Location': ['Test Location'],
      'Manager(s)': ['Test Manager'],
      'Member Email': ['test@example.com'],
      'Steward First Name': ['Test'],
      'Steward Last Name': ['Steward'],
      'Steward Email': ['steward@example.com'],
      'Date of Incident': [new Date().toISOString()],
      'Articles Violated': ['Art. 6 - Hours of Work'],
      'Remedy Sought': ['Test remedy'],
      'Date Filed': [new Date().toISOString()],
      'Step (I/II/III)': ['Step I'],
      'Confidential Waiver Attached?': ['Yes']
    }
  };

  onGrievanceFormSubmit(testEvent);
  SpreadsheetApp.getActiveSpreadsheet().toast('Test grievance created!', '✅ Test Complete', 3);
}

// ============================================================================
// PERSONAL CONTACT INFO FORM HANDLER
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
  var response = ui.alert('📋 Personal Contact Info Form',
    'Share this form with members to collect their contact information.\n\n' +
    'When submitted, the data will be written to the Member Directory:\n' +
    '• Existing members (matched by name) will be updated\n' +
    '• New members will be added automatically\n\n' +
    '• Click YES to open the form\n' +
    '• Click NO to copy the link',
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
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + formUrl + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">📋 Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  ta.select();' +
      '  document.execCommand("copy");' +
      '  document.getElementById("copied").style.display = "inline";' +
      '}' +
      '</script>' +
      '</div>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, '📋 Contact Form Link');
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

    // Find the member by first name + last name
    var data = memberSheet.getDataRange().getValues();
    var memberRow = -1;

    for (var i = 1; i < data.length; i++) {
      var rowFirstName = (data[i][MEMBER_COLS.FIRST_NAME - 1] || '').toString().trim().toLowerCase();
      var rowLastName = (data[i][MEMBER_COLS.LAST_NAME - 1] || '').toString().trim().toLowerCase();

      if (rowFirstName === firstName.toLowerCase().trim() &&
          rowLastName === lastName.toLowerCase().trim()) {
        memberRow = i + 1; // Convert to 1-indexed row number
        break;
      }
    }

    if (memberRow === -1) {
      // Member not found - create new member
      Logger.log('Creating new member: ' + firstName + ' ' + lastName);

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
      Logger.log('Created new member ' + memberId + ': ' + firstName + ' ' + lastName);

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

      Logger.log('Updated contact info for ' + firstName + ' ' + lastName + ' (row ' + memberRow + ')');
    }

  } catch (error) {
    Logger.log('Error processing contact form submission: ' + error.message);
    throw error;
  }
}

/**
 * Get multiple values from form response (for checkbox questions)
 * Returns comma-separated string
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
    ui.alert('ℹ️ Trigger Exists',
      'A contact form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('📋 Setup Contact Form Trigger',
    'This will set up automatic processing of contact info form submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  if (!formUrl) {
    ui.alert('❌ No URL', 'Please provide the form edit URL.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Extract form ID from URL
    var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      ui.alert('❌ Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onContactFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('✅ Trigger Created',
      'Contact form trigger has been set up!\n\n' +
      'When a contact form is submitted:\n' +
      '• The member\'s record will be updated in Member Directory\n' +
      '• Contact info, preferences, and interests will be saved',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', '✅ Success', 3);

  } catch (e) {
    ui.alert('❌ Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// MEMBER SATISFACTION SURVEY FORM HANDLER
// ============================================================================

/**
 * Member Satisfaction Survey Form Configuration
 */
var SATISFACTION_FORM_CONFIG = {
  // Google Form URLs
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLSeR4VxrGTEvK-PaQP2S8JXn6xwTwp-vkR9tI5c3PRvfhr75nA/viewform',
  EDIT_URL: 'https://docs.google.com/forms/d/10irg3mZ4kPShcJ5gFHuMoTxvTeZmo_cBs6HGvfasbL0/edit',

  // Form field entry IDs (from pre-filled URL)
  FIELD_IDS: {
    WORKSITE: 'entry.829990399',
    TOP_PRIORITIES: 'entry.1290096581',      // Multi-select checkboxes
    ONE_CHANGE: 'entry.1926319061',
    KEEP_DOING: 'entry.1554906279',
    ADDITIONAL_COMMENTS: 'entry.650574503'
  }
};

/**
 * Show the Member Satisfaction Survey form link
 * Survey responses are written to the Member Satisfaction sheet
 */
function getSatisfactionSurveyLink() {
  var ui = SpreadsheetApp.getUi();
  // Get form URL from Config (allows admin to update without code changes)
  var formUrl = getFormUrlFromConfig('satisfaction');

  // Show dialog with form link options
  var response = ui.alert('📊 Member Satisfaction Survey',
    'Share this survey with members to collect feedback.\n\n' +
    'When submitted, responses will be written to the\n' +
    '📊 Member Satisfaction sheet.\n\n' +
    '• Click YES to open the survey\n' +
    '• Click NO to copy the link',
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
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + formUrl + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">📋 Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  ta.select();' +
      '  document.execCommand("copy");' +
      '  document.getElementById("copied").style.display = "inline";' +
      '}' +
      '</script>' +
      '</div>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, '📊 Survey Link');
  }
}

/**
 * Save form URLs to the Config tab for easy reference and updating
 * Writes Grievance Form, Contact Form, and Satisfaction Survey URLs to Config columns P, Q, AR
 */
function saveFormUrlsToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  saveFormUrlsToConfig_silent(ss);
  ss.toast('Form URLs saved to Config tab (columns P, Q, AR)', '✅ Saved', 3);
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

  // Set headers in row 1
  configSheet.getRange(1, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue('Grievance Form URL');
  configSheet.getRange(1, CONFIG_COLS.CONTACT_FORM_URL).setValue('Contact Form URL');
  configSheet.getRange(1, CONFIG_COLS.SATISFACTION_FORM_URL).setValue('Satisfaction Survey URL');

  // Set form URLs in row 2
  configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue(GRIEVANCE_FORM_CONFIG.FORM_URL);
  configSheet.getRange(2, CONFIG_COLS.CONTACT_FORM_URL).setValue(CONTACT_FORM_CONFIG.FORM_URL);
  configSheet.getRange(2, CONFIG_COLS.SATISFACTION_FORM_URL).setValue(SATISFACTION_FORM_CONFIG.FORM_URL);

  // Format as links
  configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(2, CONFIG_COLS.CONTACT_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(2, CONFIG_COLS.SATISFACTION_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
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
    newRow[SATISFACTION_COLS.Q9_RECOMMEND - 1] = getFormValue_(responses, 'Voted during the last election');

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

    // ══════════════════════════════════════════════════════════════════════════
    // EMAIL VERIFICATION & QUARTERLY TRACKING
    // ══════════════════════════════════════════════════════════════════════════

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
    ui.alert('ℹ️ Trigger Exists',
      'A satisfaction survey trigger already exists.\n\n' +
      'Survey submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('📊 Setup Satisfaction Survey Trigger',
    'This will set up automatic processing of survey submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  if (!formUrl) {
    ui.alert('❌ No URL', 'Please provide the form edit URL.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Extract form ID from URL
    var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      ui.alert('❌ Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onSatisfactionFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('✅ Trigger Created',
      'Satisfaction survey trigger has been set up!\n\n' +
      'When a survey is submitted:\n' +
      '• Response will be added to 📊 Member Satisfaction sheet\n' +
      '• All 68 questions will be recorded\n' +
      '• Dashboard will reflect new data',
      ui.ButtonSet.OK);

    ss.toast('Survey trigger created successfully!', '✅ Success', 3);

  } catch (e) {
    ui.alert('❌ Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// SURVEY ENHANCEMENTS - Auto-Email, Quarterly Tracking, Member Auth
// ============================================================================

/**
 * Send satisfaction survey emails to random members
 * Allows stewards to email a configurable number of random members
 */
function sendRandomSurveyEmails() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show configuration dialog
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
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
    '<h2>📧 Send Survey to Random Members</h2>' +
    '<div class="info">💡 Select how many random members to email. Each member will receive a personalized survey link.</div>' +
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
    '<button class="primary" onclick="send()">📧 Send Surveys</button></div></div>' +
    '<script>function send(){var opts={count:parseInt(document.getElementById("count").value),' +
    'subject:document.getElementById("subject").value,excludeDays:parseInt(document.getElementById("excludeDays").value)};' +
    'google.script.run.withSuccessHandler(function(r){alert(r);google.script.host.close()})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message)}).executeSendRandomSurveyEmails(opts)}</script></body></html>'
  ).setWidth(500).setHeight(450);

  ui.showModalDialog(html, '📧 Send Random Survey Emails');
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

  var result = '✅ Sent ' + sent + ' survey emails';
  if (errors.length > 0) {
    result += '\n\n⚠️ ' + errors.length + ' errors:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) result += '\n...and ' + (errors.length - 5) + ' more';
  }

  return result;
}

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

// ============================================================================
// FLAGGED SUBMISSIONS REVIEW - Admin interface for pending survey responses
// ============================================================================

/**
 * Show the flagged submissions review interface
 * Displays count and email addresses of Pending Review submissions
 * Protects actual survey answers - only shows metadata
 */
function showFlaggedSubmissionsReview() {
  var html = HtmlService.createHtmlOutput(getFlaggedSubmissionsHtml())
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Flagged Survey Submissions Review');
}

/**
 * Get HTML for flagged submissions review interface
 * @returns {string} HTML content
 */
function getFlaggedSubmissionsHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    ':root{--purple:#5B4B9E;--green:#059669;--red:#DC2626;--orange:#F97316}' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f5f5f5;padding:20px}' +
    '.container{max-width:650px;margin:0 auto}' +
    '.stats-row{display:flex;gap:15px;margin-bottom:20px}' +
    '.stat-card{flex:1;background:white;padding:20px;border-radius:12px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.stat-card.pending{border-left:4px solid var(--orange)}' +
    '.stat-card.verified{border-left:4px solid var(--green)}' +
    '.stat-value{font-size:32px;font-weight:bold;color:#333}' +
    '.stat-label{font-size:13px;color:#666;margin-top:5px}' +
    '.section{background:white;border-radius:12px;padding:20px;margin-bottom:15px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.section-title{font-size:16px;font-weight:600;color:#333;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #eee}' +
    '.email-list{max-height:250px;overflow-y:auto}' +
    '.email-item{display:flex;align-items:center;justify-content:space-between;padding:12px;background:#f8f9fa;border-radius:8px;margin-bottom:8px}' +
    '.email-info{display:flex;align-items:center;gap:10px}' +
    '.email-text{font-size:14px;color:#333}' +
    '.email-date{font-size:12px;color:#666}' +
    '.actions{display:flex;gap:8px}' +
    '.btn{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:500}' +
    '.btn-approve{background:#059669;color:white}' +
    '.btn-reject{background:#DC2626;color:white}' +
    '.empty-state{text-align:center;padding:40px;color:#666}' +
    '.info-box{background:#E8F4FD;padding:15px;border-radius:8px;margin-bottom:15px;font-size:13px;color:#1E40AF}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div id="content"><div class="empty-state">Loading...</div></div>' +
    '</div>' +
    '<script>' +
    'function load(){google.script.run.withSuccessHandler(render).getFlaggedSubmissionsData()}' +
    'function render(d){' +
    '  var h="<div class=\\"stats-row\\">";' +
    '  h+="<div class=\\"stat-card pending\\"><div class=\\"stat-value\\">"+d.pendingCount+"</div><div class=\\"stat-label\\">Pending Review</div></div>";' +
    '  h+="<div class=\\"stat-card verified\\"><div class=\\"stat-value\\">"+d.verifiedCount+"</div><div class=\\"stat-label\\">Verified Responses</div></div>";' +
    '  h+="</div>";' +
    '  h+="<div class=\\"info-box\\">⚠️ These submissions could not be matched to a member email. Survey answers are protected and not shown here.</div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📧 Pending Review Emails ("+d.pendingCount+")</div>";' +
    '  if(d.pendingEmails.length===0){' +
    '    h+="<div class=\\"empty-state\\">✅ No submissions pending review</div>";' +
    '  }else{' +
    '    h+="<div class=\\"email-list\\">";' +
    '    d.pendingEmails.forEach(function(e){' +
    '      h+="<div class=\\"email-item\\"><div class=\\"email-info\\">";' +
    '      h+="<span class=\\"email-text\\">"+e.email+"</span>";' +
    '      h+="<span class=\\"email-date\\">"+e.date+" | "+e.quarter+"</span></div>";' +
    '      h+="<div class=\\"actions\\">";' +
    '      h+="<button class=\\"btn btn-approve\\" onclick=\\"approve("+e.row+\")\\">✓ Approve</button>";' +
    '      h+="<button class=\\"btn btn-reject\\" onclick=\\"reject("+e.row+\")\\">✗ Reject</button>";' +
    '      h+="</div></div>";' +
    '    });' +
    '    h+="</div>";' +
    '  }' +
    '  h+="</div>";' +
    '  document.getElementById("content").innerHTML=h;' +
    '}' +
    'function approve(row){' +
    '  if(confirm("Mark this submission as verified? This will include it in statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).approveFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'function reject(row){' +
    '  if(confirm("Reject this submission? It will be excluded from all statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).rejectFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'load();' +
    '</script></body></html>';
}

/**
 * Get data for flagged submissions review
 * @returns {Object} Pending submissions data (email, date, row number - NO survey answers)
 */
function getFlaggedSubmissionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    pendingCount: 0,
    verifiedCount: 0,
    pendingEmails: []
  };

  if (!satSheet) return result;

  var data = satSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var verified = data[i][SATISFACTION_COLS.VERIFIED - 1];
    var email = data[i][SATISFACTION_COLS.EMAIL - 1] || '(no email provided)';
    var timestamp = data[i][SATISFACTION_COLS.TIMESTAMP - 1];
    var quarter = data[i][SATISFACTION_COLS.QUARTER - 1] || '';

    if (verified === 'Yes') {
      result.verifiedCount++;
    } else if (verified === 'Pending Review') {
      result.pendingCount++;
      result.pendingEmails.push({
        email: email.toString(),
        date: timestamp ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'MMM d, yyyy') : 'Unknown',
        quarter: quarter,
        row: i + 1  // 1-indexed row number for editing
      });
    }
    // Rejected submissions are counted but not shown
  }

  // Sort by most recent first
  result.pendingEmails.sort(function(a, b) { return b.row - a.row; });

  return result;
}

/**
 * Approve a flagged submission - mark as Verified
 * @param {number} rowNum - Row number (1-indexed)
 */
function approveFlaggedSubmission(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet || rowNum < 2) return;

  satSheet.getRange(rowNum, SATISFACTION_COLS.VERIFIED).setValue('Yes');
  satSheet.getRange(rowNum, SATISFACTION_COLS.REVIEWER_NOTES).setValue('Manually approved on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));

  // Update dashboard
  syncSatisfactionValues();
}

/**
 * Reject a flagged submission - mark as Rejected
 * @param {number} rowNum - Row number (1-indexed)
 */
function rejectFlaggedSubmission(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet || rowNum < 2) return;

  satSheet.getRange(rowNum, SATISFACTION_COLS.VERIFIED).setValue('Rejected');
  satSheet.getRange(rowNum, SATISFACTION_COLS.IS_LATEST).setValue('No');
  satSheet.getRange(rowNum, SATISFACTION_COLS.REVIEWER_NOTES).setValue('Rejected on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));

  // Update dashboard
  syncSatisfactionValues();
}

// ============================================================================
// PUBLIC MEMBER DASHBOARD - Stats without PII
// ============================================================================

/**
 * Show the public member dashboard
 * Displays anonymized statistics accessible to all members
 */
function showPublicMemberDashboard() {
  var html = HtmlService.createHtmlOutput(getPublicMemberDashboardHtml())
    .setWidth(950)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Member Dashboard - Union Statistics');
}

/**
 * Get HTML for public member dashboard
 * @returns {string} HTML content
 */
function getPublicMemberDashboardHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    ':root{--purple:#5B4B9E;--green:#059669;--blue:#1a73e8}' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f5f5f5;padding:20px}' +
    '.dashboard{max-width:900px;margin:0 auto}' +
    '.tabs{display:flex;gap:5px;margin-bottom:20px;border-bottom:2px solid #ddd;padding-bottom:10px}' +
    '.tab{padding:12px 20px;border:none;background:none;cursor:pointer;font-size:14px;font-weight:500;color:#666;border-radius:8px 8px 0 0;transition:all 0.2s}' +
    '.tab:hover{background:#e8e8e8}' +
    '.tab.active{background:var(--purple);color:white}' +
    '.tab-content{display:none;background:white;border-radius:12px;padding:25px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.tab-content.active{display:block}' +
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;margin-bottom:25px}' +
    '.stat-card{background:linear-gradient(135deg,var(--purple),#7B6BB8);color:white;padding:20px;border-radius:12px;text-align:center}' +
    '.stat-card.green{background:linear-gradient(135deg,var(--green),#10B981)}' +
    '.stat-card.blue{background:linear-gradient(135deg,var(--blue),#3B82F6)}' +
    '.stat-value{font-size:36px;font-weight:bold}' +
    '.stat-label{font-size:13px;opacity:0.9;margin-top:5px}' +
    '.section{margin-bottom:25px}' +
    '.section-title{font-size:18px;font-weight:600;color:#333;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #eee}' +
    '.bar-chart{margin:15px 0}' +
    '.bar-row{display:flex;align-items:center;margin-bottom:10px}' +
    '.bar-label{width:150px;font-size:13px;color:#555}' +
    '.bar-container{flex:1;height:24px;background:#eee;border-radius:12px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:12px;transition:width 0.5s}' +
    '.bar-value{width:50px;text-align:right;font-size:13px;font-weight:500;margin-left:10px}' +
    '.steward-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:15px}' +
    '.steward-card{background:#f8f9fa;border-radius:12px;padding:20px;border-left:4px solid var(--purple)}' +
    '.steward-name{font-weight:600;font-size:16px;color:#333}' +
    '.steward-info{font-size:13px;color:#666;margin-top:8px}' +
    '.steward-info div{margin-bottom:5px}' +
    '.loading{text-align:center;padding:40px;color:#666}' +
    '</style></head><body>' +
    '<div class="dashboard">' +
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\',this)">📊 Overview</button>' +
    '<button class="tab" onclick="switchTab(\'survey\',this)">📋 Survey Results</button>' +
    '<button class="tab" onclick="switchTab(\'grievances\',this)">⚖️ Grievance Stats</button>' +
    '<button class="tab" onclick="switchTab(\'stewards\',this)">👥 Steward Directory</button>' +
    '</div>' +
    '<div id="overview" class="tab-content active"><div class="loading">Loading...</div></div>' +
    '<div id="survey" class="tab-content"><div class="loading">Loading...</div></div>' +
    '<div id="grievances" class="tab-content"><div class="loading">Loading...</div></div>' +
    '<div id="stewards" class="tab-content"><div class="loading">Loading...</div></div>' +
    '</div>' +
    '<script>' +
    'var loadedTabs={};' +
    'function switchTab(id,btn){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  document.getElementById(id).classList.add("active");' +
    '  if(!loadedTabs[id])loadTab(id);' +
    '}' +
    'function loadTab(id){' +
    '  loadedTabs[id]=true;' +
    '  if(id==="overview")google.script.run.withSuccessHandler(renderOverview).getPublicOverviewData();' +
    '  else if(id==="survey")google.script.run.withSuccessHandler(renderSurvey).getPublicSurveyData(includeHistory);' +
    '  else if(id==="grievances")google.script.run.withSuccessHandler(renderGrievances).getPublicGrievanceData();' +
    '  else if(id==="stewards")google.script.run.withSuccessHandler(renderStewards).getPublicStewardData();' +
    '}' +
    'function renderOverview(d){' +
    '  var h="<div class=\\"stats-grid\\">";' +
    '  h+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+d.totalMembers+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  h+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+d.totalStewards+"</div><div class=\\"stat-label\\">Active Stewards</div></div>";' +
    '  h+="<div class=\\"stat-card blue\\"><div class=\\"stat-value\\">"+d.totalGrievances+"</div><div class=\\"stat-label\\">Total Grievances</div></div>";' +
    '  h+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+d.winRate+"%</div><div class=\\"stat-label\\">Grievance Win Rate</div></div>";' +
    '  h+="</div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📍 Members by Location</div><div class=\\"bar-chart\\">";' +
    '  d.locationBreakdown.forEach(function(l){' +
    '    var pct=d.totalMembers>0?Math.round(l.count/d.totalMembers*100):0;' +
    '    h+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+l.location+"</div>";' +
    '    h+="<div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:#5B4B9E\\"></div></div>";' +
    '    h+="<div class=\\"bar-value\\">"+l.count+"</div></div>";' +
    '  });' +
    '  h+="</div></div>";' +
    '  document.getElementById("overview").innerHTML=h;' +
    '}' +
    'var includeHistory=false;' +
    'function renderSurvey(d){' +
    '  var h="<div class=\\"toggle-row\\" style=\\"display:flex;align-items:center;justify-content:flex-end;margin-bottom:15px;gap:10px\\"><label style=\\"font-size:13px;color:#666\\">Include historical responses:</label><input type=\\"checkbox\\" id=\\"historyToggle\\" "+(includeHistory?"checked":"")+" onchange=\\"toggleHistory(this.checked)\\" style=\\"width:18px;height:18px;cursor:pointer\\"></div>";' +
    '  h+="<div class=\\"stats-grid\\">";' +
    '  h+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+(d.verifiedResponses||d.totalResponses)+"</div><div class=\\"stat-label\\">Verified Responses</div></div>";' +
    '  h+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+d.avgSatisfaction.toFixed(1)+"</div><div class=\\"stat-label\\">Avg Satisfaction (1-10)</div></div>";' +
    '  h+="<div class=\\"stat-card blue\\"><div class=\\"stat-value\\">"+d.responseRate+"%</div><div class=\\"stat-label\\">Response Rate</div></div>";' +
    '  h+="</div>";' +
    '  if(d.includesHistory){h+="<div style=\\"background:#FEF3C7;padding:10px;border-radius:8px;margin-bottom:15px;font-size:13px;color:#92400E\\">⚠️ Showing all responses including historical (superseded) entries</div>";}' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📊 Satisfaction by Section</div><div class=\\"bar-chart\\">";' +
    '  d.sectionScores.forEach(function(s){' +
    '    var pct=Math.round(s.score*10);' +
    '    var color=s.score>=7?"#059669":s.score>=5?"#F59E0B":"#EF4444";' +
    '    h+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+s.section+"</div>";' +
    '    h+="<div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"></div></div>";' +
    '    h+="<div class=\\"bar-value\\">"+s.score.toFixed(1)+"</div></div>";' +
    '  });' +
    '  h+="</div></div>";' +
    '  document.getElementById("survey").innerHTML=h;' +
    '}' +
    'function toggleHistory(val){includeHistory=val;loadedTabs["survey"]=false;loadTab("survey");}' +
    'function renderGrievances(d){' +
    '  var h="<div class=\\"stats-grid\\">";' +
    '  h+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+d.open+"</div><div class=\\"stat-label\\">Open Grievances</div></div>";' +
    '  h+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+d.won+"</div><div class=\\"stat-label\\">Won</div></div>";' +
    '  h+="<div class=\\"stat-card blue\\"><div class=\\"stat-value\\">"+d.settled+"</div><div class=\\"stat-label\\">Settled</div></div>";' +
    '  h+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+d.avgDaysToResolve+"</div><div class=\\"stat-label\\">Avg Days to Resolve</div></div>";' +
    '  h+="</div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📊 Grievances by Type</div><div class=\\"bar-chart\\">";' +
    '  d.byType.forEach(function(t){' +
    '    var pct=d.total>0?Math.round(t.count/d.total*100):0;' +
    '    h+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+t.type+"</div>";' +
    '    h+="<div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:#5B4B9E\\"></div></div>";' +
    '    h+="<div class=\\"bar-value\\">"+t.count+"</div></div>";' +
    '  });' +
    '  h+="</div></div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📊 Grievances by Status</div><div class=\\"bar-chart\\">";' +
    '  d.byStatus.forEach(function(s){' +
    '    var pct=d.total>0?Math.round(s.count/d.total*100):0;' +
    '    var color=s.status==="Open"?"#3B82F6":s.status==="Won"?"#059669":"#6B7280";' +
    '    h+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+s.status+"</div>";' +
    '    h+="<div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"></div></div>";' +
    '    h+="<div class=\\"bar-value\\">"+s.count+"</div></div>";' +
    '  });' +
    '  h+="</div></div>";' +
    '  document.getElementById("grievances").innerHTML=h;' +
    '}' +
    'function renderStewards(d){' +
    '  var h="<div class=\\"section\\"><div class=\\"section-title\\">👥 Your Union Stewards ("+d.stewards.length+")</div>";' +
    '  h+="<p style=\\"color:#666;margin-bottom:20px\\">Contact your steward for help with workplace issues, grievances, or union questions.</p>";' +
    '  h+="<div class=\\"steward-grid\\">";' +
    '  d.stewards.forEach(function(s){' +
    '    h+="<div class=\\"steward-card\\">";' +
    '    h+="<div class=\\"steward-name\\">"+s.name+"</div>";' +
    '    h+="<div class=\\"steward-info\\">";' +
    '    h+="<div>📍 "+s.location+"</div>";' +
    '    h+="<div>📅 Office Days: "+s.officeDays+"</div>";' +
    '    h+="<div>📧 "+s.email+"</div>";' +
    '    h+="</div></div>";' +
    '  });' +
    '  h+="</div></div>";' +
    '  document.getElementById("stewards").innerHTML=h;' +
    '}' +
    'loadTab("overview");' +
    '</script></body></html>';
}

/**
 * Get public overview data (no PII)
 * @returns {Object} Overview statistics
 */
function getPublicOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    totalMembers: 0,
    totalStewards: 0,
    totalGrievances: 0,
    winRate: 0,
    locationBreakdown: []
  };

  // Count members and stewards
  if (memberSheet) {
    var memberData = memberSheet.getDataRange().getValues();
    var locationCounts = {};
    var stewardCount = 0;

    for (var i = 1; i < memberData.length; i++) {
      var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
      if (!memberId || !memberId.toString().match(/^M/i)) continue;

      result.totalMembers++;

      // Count by location
      var location = memberData[i][MEMBER_COLS.LOCATION - 1] || 'Unknown';
      locationCounts[location] = (locationCounts[location] || 0) + 1;

      // Count stewards
      var isSteward = memberData[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isSteward === true || isSteward === 'Yes' || isSteward === 'TRUE') {
        stewardCount++;
      }
    }

    result.totalStewards = stewardCount;

    // Convert location counts to array and sort
    Object.keys(locationCounts).forEach(function(loc) {
      result.locationBreakdown.push({ location: loc, count: locationCounts[loc] });
    });
    result.locationBreakdown.sort(function(a, b) { return b.count - a.count; });
    result.locationBreakdown = result.locationBreakdown.slice(0, 10); // Top 10
  }

  // Count grievances and win rate
  if (grievanceSheet) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    var won = 0, total = 0;

    for (var j = 1; j < grievanceData.length; j++) {
      var grievanceId = grievanceData[j][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      if (!grievanceId) continue;

      total++;
      var resolution = (grievanceData[j][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString().toLowerCase();
      if (resolution.includes('won') || resolution.includes('favor')) {
        won++;
      }
    }

    result.totalGrievances = total;
    result.winRate = total > 0 ? Math.round(won / total * 100) : 0;
  }

  return result;
}

/**
 * Get public survey data (anonymized)
 * Filters to only include Verified='Yes' and optionally IS_LATEST='Yes' responses
 * @param {boolean} includeHistory - If true, include superseded responses; if false, only latest per member
 * @returns {Object} Survey statistics
 */
function getPublicSurveyData(includeHistory) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = {
    totalResponses: 0,
    verifiedResponses: 0,
    avgSatisfaction: 0,
    responseRate: 0,
    sectionScores: [],
    includesHistory: includeHistory || false
  };

  if (!satSheet) return result;

  var data = satSheet.getDataRange().getValues();
  if (data.length < 2) return result;

  // Filter rows to only include verified responses
  // If includeHistory is false (default), also filter to IS_LATEST='Yes'
  var validRows = [];
  for (var i = 1; i < data.length; i++) {
    var verified = data[i][SATISFACTION_COLS.VERIFIED - 1];
    var isLatest = data[i][SATISFACTION_COLS.IS_LATEST - 1];

    // Only include Verified='Yes' responses
    if (verified !== 'Yes') continue;

    // If not including history, only include IS_LATEST='Yes'
    if (!includeHistory && isLatest !== 'Yes') continue;

    validRows.push(data[i]);
  }

  result.totalResponses = data.length - 1; // Total submissions (all)
  result.verifiedResponses = validRows.length; // Verified responses used in calculations

  // Calculate average satisfaction (Q6 - Satisfied with representation)
  var satSum = 0, satCount = 0;
  for (var j = 0; j < validRows.length; j++) {
    var sat = parseFloat(validRows[j][SATISFACTION_COLS.Q6_SATISFIED_REP - 1]);
    if (!isNaN(sat)) {
      satSum += sat;
      satCount++;
    }
  }
  result.avgSatisfaction = satCount > 0 ? satSum / satCount : 0;

  // Response rate (unique verified members / total members)
  if (memberSheet) {
    var memberCount = memberSheet.getLastRow() - 1;
    // Count unique verified member IDs
    var uniqueMembers = {};
    for (var k = 0; k < validRows.length; k++) {
      var memberId = validRows[k][SATISFACTION_COLS.MATCHED_MEMBER_ID - 1];
      if (memberId) uniqueMembers[memberId] = true;
    }
    var uniqueCount = Object.keys(uniqueMembers).length;
    result.responseRate = memberCount > 0 ? Math.round(uniqueCount / memberCount * 100) : 0;
  }

  // Section scores using only verified responses
  var sections = [
    { name: 'Overall Satisfaction', cols: [SATISFACTION_COLS.Q6_SATISFIED_REP, SATISFACTION_COLS.Q7_TRUST_UNION, SATISFACTION_COLS.Q8_FEEL_PROTECTED] },
    { name: 'Steward Ratings', cols: [SATISFACTION_COLS.Q10_TIMELY_RESPONSE, SATISFACTION_COLS.Q11_TREATED_RESPECT, SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS] },
    { name: 'Chapter Effectiveness', cols: [SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, SATISFACTION_COLS.Q22_CHAPTER_COMM, SATISFACTION_COLS.Q23_ORGANIZES] },
    { name: 'Local Leadership', cols: [SATISFACTION_COLS.Q26_DECISIONS_CLEAR, SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS, SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE] },
    { name: 'Communication', cols: [SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, SATISFACTION_COLS.Q42_ENOUGH_INFO] }
  ];

  sections.forEach(function(section) {
    var sum = 0, count = 0;
    for (var m = 0; m < validRows.length; m++) {
      section.cols.forEach(function(col) {
        if (col) {
          var val = parseFloat(validRows[m][col - 1]);
          if (!isNaN(val)) {
            sum += val;
            count++;
          }
        }
      });
    }
    result.sectionScores.push({
      section: section.name,
      score: count > 0 ? sum / count : 0
    });
  });

  result.sectionScores.sort(function(a, b) { return b.score - a.score; });

  return result;
}

/**
 * Get public grievance data (no PII)
 * @returns {Object} Grievance statistics
 */
function getPublicGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    total: 0,
    open: 0,
    won: 0,
    settled: 0,
    avgDaysToResolve: 0,
    byType: [],
    byStatus: []
  };

  if (!grievanceSheet) return result;

  var data = grievanceSheet.getDataRange().getValues();
  var typeCounts = {};
  var statusCounts = {};
  var daysToResolve = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    result.total++;

    var status = data[i][GRIEVANCE_COLS.STATUS - 1] || 'Unknown';
    var resolution = (data[i][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString();
    var gType = data[i][GRIEVANCE_COLS.GRIEVANCE_TYPE - 1] || 'Other';

    // Count by status
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    // Count by type
    typeCounts[gType] = (typeCounts[gType] || 0) + 1;

    // Track open/won/settled
    if (status === 'Open' || status === 'Pending Info') {
      result.open++;
    }
    if (resolution.toLowerCase().includes('won') || resolution.toLowerCase().includes('favor')) {
      result.won++;
    }
    if (resolution.toLowerCase().includes('settled')) {
      result.settled++;
    }

    // Calculate days to resolve for closed grievances
    if (status === 'Closed' || status === 'Resolved') {
      var dateOpened = data[i][GRIEVANCE_COLS.DATE_OPENED - 1];
      var dateClosed = data[i][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateOpened && dateClosed) {
        var days = Math.round((new Date(dateClosed) - new Date(dateOpened)) / (1000 * 60 * 60 * 24));
        if (days > 0) daysToResolve.push(days);
      }
    }
  }

  // Average days to resolve
  if (daysToResolve.length > 0) {
    result.avgDaysToResolve = Math.round(daysToResolve.reduce(function(a, b) { return a + b; }, 0) / daysToResolve.length);
  }

  // Convert to arrays
  Object.keys(typeCounts).forEach(function(t) {
    result.byType.push({ type: t, count: typeCounts[t] });
  });
  result.byType.sort(function(a, b) { return b.count - a.count; });
  result.byType = result.byType.slice(0, 8);

  Object.keys(statusCounts).forEach(function(s) {
    result.byStatus.push({ status: s, count: statusCounts[s] });
  });
  result.byStatus.sort(function(a, b) { return b.count - a.count; });

  return result;
}

/**
 * Get public steward data (contact info only)
 * @returns {Object} Steward directory
 */
function getPublicStewardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = { stewards: [] };

  if (!memberSheet) return result;

  var data = memberSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isSteward !== true && isSteward !== 'Yes' && isSteward !== 'TRUE') continue;

    var firstName = data[i][MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = data[i][MEMBER_COLS.LAST_NAME - 1] || '';

    result.stewards.push({
      name: firstName + ' ' + lastName,
      location: data[i][MEMBER_COLS.LOCATION - 1] || 'Not specified',
      officeDays: data[i][MEMBER_COLS.OFFICE_DAYS - 1] || 'Contact for availability',
      email: data[i][MEMBER_COLS.EMAIL - 1] || 'Contact union office'
    });
  }

  // Sort by name
  result.stewards.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return result;
}

/**
 * Uses hidden sheet formulas for self-healing calculations
 */
// Note: recalcAllGrievancesBatched() defined in modular file - see respective module

/**
 * Refresh Member Directory calculated columns (AB-AD: Has Open Grievance, Status, Next Deadline)
 */
function refreshMemberDirectoryFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing Member Directory...', '🔄 Refresh', 3);

  // Step 1: Refresh grievance formulas first (to get latest Next Action Due dates)
  syncGrievanceFormulasToLog();

  // Step 2: Sync grievance data to member directory (updates AB-AD columns)
  syncGrievanceToMemberDirectory();

  // Step 3: Repair member checkboxes
  repairMemberCheckboxes();

  ss.toast('Member Directory refreshed!', '✅ Success', 3);
}

function rebuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Rebuilding dashboard sheets...', '🔄 Rebuild', 3);

  // Recreate dashboard sheets with latest layout
  createDashboard(ss);
  createInteractiveDashboard(ss);

  // Refresh hidden sheet formulas and sync data
  refreshAllHiddenFormulas();

  // Reapply data validations
  setupDataValidations();

  ss.toast('Dashboard rebuilt with all 9 sections!', '✅ Success', 3);
}

/**
 * Refresh all formulas and sync all data
 */
function refreshAllFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing all formulas and syncing data...', '🔄 Refresh', 3);

  // Use the full refresh from HiddenSheets.gs
  refreshAllHiddenFormulas();
}

// ============================================================================
// VIEW CONTROLS - Timeline Simplification
// ============================================================================

/**
 * Simplify the Grievance Log timeline view
 * Hides Step II and Step III columns, keeping only essential dates
 * Shows: Incident Date, Date Filed, Date Closed, Days Open, Next Action Due, Days to Deadline
 */
function simplifyTimelineView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Simplifying timeline view...', '👁️ View', 2);

  // Hide Step I detail columns (J-K): Step I Due, Step I Rcvd
  sheet.hideColumns(GRIEVANCE_COLS.STEP1_DUE, 2);

  // Hide Step II columns (L-O): Appeal Due, Appeal Filed, Due, Rcvd
  sheet.hideColumns(GRIEVANCE_COLS.STEP2_APPEAL_DUE, 4);

  // Hide Step III columns (P-Q): Appeal Due, Appeal Filed
  sheet.hideColumns(GRIEVANCE_COLS.STEP3_APPEAL_DUE, 2);

  // Hide Filing Deadline (H) - auto-calculated, less important once filed
  sheet.hideColumns(GRIEVANCE_COLS.FILING_DEADLINE, 1);

  ss.toast('Timeline simplified! Showing only key dates: Incident, Filed, Closed, Next Due', '✅ Done', 3);
}

/**
 * Show the full timeline view
 * Unhides all date columns
 */
function showFullTimelineView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Showing full timeline...', '👁️ View', 2);

  // Show all timeline columns (H through Q)
  sheet.showColumns(GRIEVANCE_COLS.FILING_DEADLINE, 10); // H through Q

  ss.toast('Full timeline view restored!', '✅ Done', 3);
}

/**
 * Setup column groups for the timeline
 * Creates expandable/collapsible groups for Step II and Step III
 */
function setupTimelineColumnGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Setting up column groups...', '👁️ View', 2);

  // Group Step I columns (J-K)
  var step1Range = sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, 1, 2);
  sheet.getColumnGroup(GRIEVANCE_COLS.STEP1_DUE, 1);

  // Group Step II columns (L-O)
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  var step2Group = sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, 1, 4);
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  // Group Step III columns (P-Q)
  var step3Group = sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, 1, 2);

  // Create the groups
  try {
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);

    // Collapse Step II and III by default (Step I usually visible)
    sheet.collapseAllColumnGroups();

    ss.toast('Column groups created! Click +/- to expand/collapse step details', '✅ Done', 5);
  } catch (e) {
    Logger.log('Column group error: ' + e.toString());
    ss.toast('Column groups may already exist or require manual setup', '⚠️ Note', 3);
  }
}

/**
 * Apply conditional formatting to highlight the current step's dates
 * Grays out dates for steps not yet reached
 */
function applyStepHighlighting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Applying step highlighting...', '🎨 Format', 3);

  var lastRow = Math.max(sheet.getLastRow(), 2);
  var rules = sheet.getConditionalFormatRules();

  // Colors
  var grayText = SpreadsheetApp.newColor().setRgbColor('#9e9e9e').build();
  var greenBg = SpreadsheetApp.newColor().setRgbColor('#e8f5e9').build();
  var currentStepCol = GRIEVANCE_COLS.CURRENT_STEP; // Column F

  // Rule 1: Gray out Step I columns (J-K) if current step is Informal
  var step1Range = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, lastRow - 1, 2);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Informal"')
    .setFontColor('#9e9e9e')
    .setRanges([step1Range])
    .build();

  // Rule 2: Gray out Step II columns (L-O) if current step is Informal or Step I
  var step2Range = sheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, lastRow - 1, 4);
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($F2="Informal",$F2="Step I")')
    .setFontColor('#9e9e9e')
    .setRanges([step2Range])
    .build();

  // Rule 3: Gray out Step III columns (P-Q) if not at Step III or beyond
  var step3Range = sheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, lastRow - 1, 2);
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($F2="Informal",$F2="Step I",$F2="Step II")')
    .setFontColor('#9e9e9e')
    .setRanges([step3Range])
    .build();

  // -------------------------------------------------------------------------
  // DEADLINE STATUS RULES (Days to Deadline column U)
  // Order matters: more specific rules first, then broader ones
  // -------------------------------------------------------------------------

  var daysDeadlineRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);
  var nextDueRange = sheet.getRange(2, GRIEVANCE_COLS.NEXT_ACTION_DUE, lastRow - 1, 1);

  // Rule 4: 🔴 Red - Overdue (Days to Deadline shows "Overdue" or negative/0)
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($U2="Overdue",AND(ISNUMBER($U2),$U2<=0))')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 5: 🟠 Orange - Due in 1-3 days
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=1,$U2<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 6: 🟡 Yellow - Due in 4-7 days
  var rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=4,$U2<=7)')
    .setBackground('#fffde7')
    .setFontColor('#f57f17')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 7: 🟢 Green - On Track (more than 7 days remaining)
  var rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>7)')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 8: Red highlight for Next Action Due if overdue
  var rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($T2<>"",$T2<TODAY())')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([nextDueRange])
    .build();

  // Rule 9: Orange for Next Action Due within 3 days
  var rule9 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($T2<>"",($T2-TODAY())>=0,($T2-TODAY())<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([nextDueRange])
    .build();

  // -------------------------------------------------------------------------
  // OUTCOME STATUS RULES (Status column E)
  // -------------------------------------------------------------------------

  var statusRange = sheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1);

  // Rule 10: ✅ Green - Won
  var rule10 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Won')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Rule 11: ❌ Red - Denied
  var rule11 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Denied')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Rule 12: 🤝 Blue - Settled
  var rule12 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Settled')
    .setBackground('#e3f2fd')
    .setFontColor('#1565c0')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Add new rules (keep existing rules)
  rules.push(rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8, rule9, rule10, rule11, rule12);
  sheet.setConditionalFormatRules(rules);

  ss.toast('Formatting applied! Deadline colors (🟢🟡🟠🔴) and outcome status (Won/Denied/Settled)', '✅ Done', 5);
}

/**
 * Freeze key columns for easier scrolling
 * Freezes A-F (Identity & Status) so they're always visible
 */
function freezeKeyColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  // Freeze first 6 columns (A-F: ID, Member ID, Name, Status, Step)
  sheet.setFrozenColumns(6);
  // Freeze header row
  sheet.setFrozenRows(1);

  ss.toast('Frozen columns A-F and header row. Scroll right to see timeline.', '❄️ Frozen', 3);
}

/**
 * Unfreeze all columns
 */
function unfreezeAllColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  sheet.setFrozenColumns(0);
  // Keep header row frozen
  sheet.setFrozenRows(1);

  ss.toast('Columns unfrozen. Header row still frozen.', '🔓 Unfrozen', 3);
}

// ============================================================================
// TESTING FUNCTIONS
// ============================================================================

function viewTestResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TEST_RESULTS);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('No test results yet. Run tests first using 🧪 Testing menu.');
  }
}

// ============================================================================
// NAVIGATION FUNCTIONS (Menu Items)
// ============================================================================

/**
 * Refresh Custom View charts and data
 */
function refreshInteractiveCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing Custom View...', '📈 Refresh', 2);

  var sheet = ss.getSheetByName(SHEETS.INTERACTIVE);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Custom View not found. Run REPAIR DASHBOARD to create it.');
    return;
  }

  // Force recalculation by flushing
  SpreadsheetApp.flush();

  // Navigate to it
  ss.setActiveSheet(sheet);
  ss.toast('Custom View refreshed!', '✅ Done', 2);
}

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

/**
 * Create a Google Drive folder for the selected grievance
 */
// Note: setupDriveFolderForGrievance() defined in modular file - see respective module

/**
 * Get or create the root 509 Dashboard folder in Drive
 */
function getOrCreateDashboardFolder_() {
  var folderName = '509 Dashboard - Grievance Files';
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

/**
 * Show files in the selected grievance's folder
 */
function showGrievanceFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('📁 View Files', 'Please go to the Grievance Log sheet and select a grievance row first.', ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('📁 View Files', 'Please select a grievance row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var folderId = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_ID).getValue();
  var folderUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
  var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!folderId) {
    var response = ui.alert('📁 No Folder',
      'No folder exists for ' + grievanceId + '.\n\nWould you like to create one?',
      ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      setupDriveFolderForGrievance();
    }
    return;
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var fileList = [];

    while (files.hasNext()) {
      var file = files.next();
      fileList.push('• ' + file.getName());
    }

    if (fileList.length === 0) {
      var response = ui.alert('📁 ' + grievanceId + ' Files',
        'Folder is empty.\n\nWould you like to open the folder to add files?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        var html = HtmlService.createHtmlOutput(
          '<script>window.open("' + folderUrl + '", "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    } else {
      var response = ui.alert('📁 ' + grievanceId + ' Files (' + fileList.length + ')',
        fileList.join('\n') + '\n\nOpen folder in Drive?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        var html = HtmlService.createHtmlOutput(
          '<script>window.open("' + folderUrl + '", "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    }
  } catch (e) {
    ui.alert('❌ Error', 'Could not access folder: ' + e.message + '\n\nThe folder may have been deleted.', ui.ButtonSet.OK);
  }
}

/**
 * Batch create folders for all grievances without folders
 */
// Note: batchCreateGrievanceFolders() defined in modular file - see respective module

// ============================================================================
// GOOGLE CALENDAR INTEGRATION
// ============================================================================

/**
 * Sync grievance deadlines to Google Calendar with rate limit handling
 */
// Note: syncDeadlinesToCalendar() defined in modular file - see respective module

/**
 * Show upcoming deadlines from calendar with member names
 */
function showUpcomingDeadlinesFromCalendar() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var today = new Date();
    var nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

    var events = calendar.getEvents(today, nextWeek, {search: 'Grievance'});

    if (events.length === 0) {
      ui.alert('📅 Upcoming Deadlines',
        'No grievance deadlines in the next 7 days!\n\n' +
        'Use "Sync Deadlines to Calendar" to add deadline events.',
        ui.ButtonSet.OK);
      return;
    }

    // Build a lookup of grievance IDs to member names
    var memberLookup = buildGrievanceMemberLookup();

    var eventList = events.map(function(e) {
      var date = Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'MM/dd');
      var title = e.getTitle();

      // Extract grievance ID from title (format: "Grievance GR-XXXX: Step X Due")
      var match = title.match(/Grievance\s+(GR-\d+)/i);
      var memberInfo = '';

      if (match && match[1] && memberLookup[match[1]]) {
        memberInfo = ' (' + memberLookup[match[1]] + ')';
      }

      return '• ' + date + ': ' + title + memberInfo;
    });

    ui.alert('📅 Upcoming Deadlines (Next 7 Days)',
      'Events with member names:\n\n' + eventList.join('\n'),
      ui.ButtonSet.OK);

  } catch (error) {
    if (error.message.indexOf('too many') !== -1 || error.message.indexOf('rate') !== -1) {
      ui.alert('⚠️ Calendar Rate Limit',
        'Google Calendar is temporarily limiting requests.\n\n' +
        'Please wait a few minutes and try again.\n\n' +
        'Tip: Avoid running calendar operations repeatedly in quick succession.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Calendar Error', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Build a lookup map of grievance IDs to member names
 * @return {Object} Map of grievanceId -> "First Last"
 */
function buildGrievanceMemberLookup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var lookup = {};

  if (!sheet || sheet.getLastRow() <= 1) return lookup;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.LAST_NAME).getValues();

  data.forEach(function(row) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
    var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';

    if (grievanceId) {
      lookup[grievanceId] = (firstName + ' ' + lastName).trim() || 'Unknown';
    }
  });

  return lookup;
}

/**
 * Clear all 509 Dashboard calendar events
 */
// Note: clearAllCalendarEvents() defined in modular file - see respective module

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Show notification settings dialog
 */
function showNotificationSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var enabled = props.getProperty('notifications_enabled') === 'true';
  var email = props.getProperty('notification_email') || Session.getEffectiveUser().getEmail();

  var response = ui.alert('📬 Notification Settings',
    'Daily deadline notifications: ' + (enabled ? 'ENABLED ✅' : 'DISABLED ❌') + '\n' +
    'Email: ' + email + '\n\n' +
    'Notifications are sent daily at 8 AM for grievances due within 3 days.\n\n' +
    'Would you like to ' + (enabled ? 'DISABLE' : 'ENABLE') + ' notifications?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    if (enabled) {
      // Disable
      props.setProperty('notifications_enabled', 'false');
      removeDailyTrigger_();
      ui.alert('📬 Notifications Disabled', 'Daily deadline notifications have been turned off.', ui.ButtonSet.OK);
    } else {
      // Enable
      props.setProperty('notifications_enabled', 'true');
      props.setProperty('notification_email', email);
      installDailyTrigger_();
      ui.alert('📬 Notifications Enabled',
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

    if (daysToDeadline !== '' && daysToDeadline <= 3) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });
    }
  }

  if (urgent.length === 0) return;

  var subject = '⚠️ 509 Dashboard: ' + urgent.length + ' Grievance Deadline(s) Approaching';
  var body = 'The following grievances have deadlines within 3 days:\n\n';

  for (var j = 0; j < urgent.length; j++) {
    var g = urgent[j];
    body += '• ' + g.id + ' (' + g.step + ') - ' +
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

  var response = ui.alert('🧪 Test Notifications',
    'This will send a test notification email to:\n' + email + '\n\nSend test email?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    MailApp.sendEmail(email,
      '🧪 509 Dashboard Test Notification',
      'This is a test notification from your 509 Dashboard.\n\n' +
      'If you received this email, notifications are working correctly!\n\n' +
      'Dashboard: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    );
    ui.alert('✅ Test Sent', 'Test email sent to ' + email + '\n\nCheck your inbox!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Failed to send test email: ' + e.message, ui.ButtonSet.OK);
  }
}

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

    var body = '📋 509 GRIEVANCE DEADLINE ALERT\n';
    body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
    body += 'Steward: ' + stewardName + '\n';
    body += 'Date: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') + '\n\n';

    if (overdue.length > 0) {
      body += '🔴 OVERDUE (' + overdue.length + ')\n';
      body += '─────────────────────\n';
      for (var o = 0; o < overdue.length; o++) {
        body += '  ⚠️ ' + overdue[o].id + ' - ' + overdue[o].memberName + '\n';
        body += '     Step: ' + overdue[o].step + ' | Status: ' + overdue[o].status + '\n';
        body += '     OVERDUE by ' + Math.abs(overdue[o].daysRemaining) + ' day(s)\n\n';
      }
    }

    if (urgent.length > 0) {
      body += '🟠 URGENT - Due within 3 days (' + urgent.length + ')\n';
      body += '─────────────────────\n';
      for (var u = 0; u < urgent.length; u++) {
        body += '  ⏰ ' + urgent[u].id + ' - ' + urgent[u].memberName + '\n';
        body += '     Step: ' + urgent[u].step + ' | Status: ' + urgent[u].status + '\n';
        body += '     Due in ' + urgent[u].daysRemaining + ' day(s)\n\n';
      }
    }

    if (upcoming.length > 0) {
      body += '🟡 UPCOMING - Due within ' + alertDays + ' days (' + upcoming.length + ')\n';
      body += '─────────────────────\n';
      for (var up = 0; up < upcoming.length; up++) {
        body += '  📅 ' + upcoming[up].id + ' - ' + upcoming[up].memberName + '\n';
        body += '     Step: ' + upcoming[up].step + ' | Due in ' + upcoming[up].daysRemaining + ' day(s)\n\n';
      }
    }

    body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    body += '📊 Dashboard: ' + ss.getUrl() + '\n';
    body += 'Total grievances requiring attention: ' + grievances.length + '\n';

    var subject = (overdue.length > 0 ? '🔴 OVERDUE: ' : '⏰ ') +
      grievances.length + ' Grievance Deadline(s) - ' + stewardName;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: 'SEIU Local 509 Dashboard'
      });
      emailsSent++;
      Logger.log('Sent alert to ' + stewardName + ' (' + email + '): ' + grievances.length + ' grievances');
    } catch (e) {
      Logger.log('Failed to send to ' + email + ': ' + e.message);
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

  var response = ui.alert('📬 Send Steward Alerts',
    'This will send deadline alert emails to all stewards with upcoming deadlines.\n\n' +
    'Each steward will receive their own personalized digest.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var emailsSent = sendStewardDeadlineAlerts();

  ui.alert('✅ Alerts Sent',
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

  var response = ui.prompt('⚙️ Alert Settings',
    'Current settings:\n' +
    '• Alert window: ' + currentDays + ' days before deadline\n' +
    '• Per-steward alerts: ' + (stewardAlerts ? 'ENABLED' : 'DISABLED') + '\n\n' +
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

  ui.alert('✅ Settings Saved',
    'Alert window: ' + newDays + ' days\n' +
    'Per-steward alerts: ' + (stewardResponse === ui.Button.YES ? 'ENABLED' : 'DISABLED'),
    ui.ButtonSet.OK);
}

// ============================================================================
// AUDIT LOGGING - Multi-Steward Accountability
// ============================================================================

/**
 * Setup the hidden audit log sheet
 * Tracks all changes to Member Directory and Grievance Log
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

/**
 * Log an audit event
 * @param {string} sheetName - Name of the sheet where change occurred
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} fieldName - Name of the field/column
 * @param {string} oldValue - Previous value
 * @param {string} newValue - New value
 * @param {string} recordId - ID of the record (Member ID or Grievance ID)
 * @param {string} actionType - Type of action (Edit, Delete, Create)
 */
// Note: logAuditEvent() defined in modular file - see respective module

/**
 * onEdit trigger for audit logging
 * Tracks changes to Member Directory and Grievance Log
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

  logAuditEvent(sheetName, row, col, fieldName, oldValue, newValue, recordId, actionType);
}

/**
 * Install the audit trigger
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

/**
 * View the audit log sheet
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

// ============================================================================
// GRIEVANCE TOOLS - ADDITIONAL FUNCTIONS
// ============================================================================

/**
 * Setup Live Grievance Links - Syncs grievance data to Member Directory
 * Uses static values (no formulas) to avoid #REF! errors
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

/**
 * Fix existing "Overdue" text in Days to Deadline column
 * Converts text back to negative numbers for proper counting
 */
function fixOverdueTextToNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log not found.');
    return;
  }

  ss.toast('Fixing overdue data...', '🔧 Fix', 3);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var daysCol = GRIEVANCE_COLS.DAYS_TO_DEADLINE;
  var nextActionCol = GRIEVANCE_COLS.NEXT_ACTION_DUE;

  var daysData = sheet.getRange(2, daysCol, lastRow - 1, 1).getValues();
  var nextActionData = sheet.getRange(2, nextActionCol, lastRow - 1, 1).getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var updates = [];
  var fixCount = 0;

  for (var i = 0; i < daysData.length; i++) {
    var currentValue = daysData[i][0];
    var nextAction = nextActionData[i][0];

    if (currentValue === 'Overdue' && nextAction instanceof Date) {
      var days = Math.floor((nextAction - today) / (1000 * 60 * 60 * 24));
      updates.push([days]);
      fixCount++;
    } else {
      updates.push([currentValue]);
    }
  }

  if (fixCount > 0) {
    sheet.getRange(2, daysCol, updates.length, 1).setValues(updates);
    ss.toast('Fixed ' + fixCount + ' overdue entries!', '✅ Success', 3);
  } else {
    ss.toast('No "Overdue" text found to fix.', '✅ All Good', 3);
  }
}

// ============================================================================
// MEMBER SATISFACTION DASHBOARD
// ============================================================================

/**
 * Shows the Member Satisfaction Dashboard modal popup
 * Menu Location: 👤 Dashboard > 📊 Member Satisfaction
 */
function showSatisfactionDashboard() {
  var html = HtmlService.createHtmlOutput(getSatisfactionDashboardHtml())
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Member Satisfaction');
}

/**
 * Returns the HTML for the Member Satisfaction Dashboard with tabs
 */
function getSatisfactionDashboardHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header - Green theme for satisfaction
    '.header{background:linear-gradient(135deg,#059669,#047857);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(11px,2.5vw,13px);opacity:0.9}' +

    // Tab navigation
    '.tabs{display:flex;background:white;border-bottom:2px solid #e0e0e0;position:sticky;top:0;z-index:100}' +
    '.tab{flex:1;padding:clamp(12px,3vw,16px);text-align:center;font-size:clamp(12px,2.5vw,14px);font-weight:600;color:#666;' +
    'border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;min-height:44px}' +
    '.tab:hover{background:#f0fdf4;color:#059669}' +
    '.tab.active{color:#059669;border-bottom-color:#059669;background:#f0fdf4}' +
    '.tab-icon{display:block;font-size:18px;margin-bottom:4px}' +

    // Tab content
    '.tab-content{display:none;padding:15px;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0}to{opacity:1}}' +

    // Stats grid
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,5vw,32px);font-weight:bold;color:#059669}' +
    '.stat-label{font-size:clamp(10px,2vw,12px);color:#666;text-transform:uppercase;margin-top:5px}' +
    '.stat-card.green .stat-value{color:#059669}' +
    '.stat-card.red .stat-value{color:#DC2626}' +
    '.stat-card.orange .stat-value{color:#F97316}' +
    '.stat-card.blue .stat-value{color:#2563EB}' +
    '.stat-card.purple .stat-value{color:#7C3AED}' +

    // Score indicator with color gradient
    '.score-indicator{display:inline-block;padding:4px 12px;border-radius:20px;font-size:14px;font-weight:bold}' +
    '.score-high{background:#d1fae5;color:#059669}' +
    '.score-mid{background:#fef3c7;color:#d97706}' +
    '.score-low{background:#fee2e2;color:#dc2626}' +

    // Data table
    '.data-table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)}' +
    '.data-table th{background:#059669;color:white;padding:12px;text-align:left;font-size:13px}' +
    '.data-table td{padding:12px;border-bottom:1px solid #eee;font-size:13px}' +
    '.data-table tr:hover{background:#f0fdf4}' +
    '.data-table tr:last-child td{border-bottom:none}' +

    // Section cards
    '.section-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:12px}' +
    '.section-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}' +
    '.section-title{font-weight:600;color:#1f2937;font-size:14px}' +
    '.section-score{font-size:20px;font-weight:bold}' +

    // Progress bar for scores
    '.progress-bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden;margin-top:8px}' +
    '.progress-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.progress-green{background:linear-gradient(90deg,#059669,#10b981)}' +
    '.progress-yellow{background:linear-gradient(90deg,#f59e0b,#fbbf24)}' +
    '.progress-red{background:linear-gradient(90deg,#dc2626,#ef4444)}' +

    // Action buttons
    '.action-btn{display:inline-flex;align-items:center;gap:8px;padding:10px 16px;border:none;border-radius:8px;' +
    'cursor:pointer;font-size:13px;font-weight:500;transition:all 0.2s;min-height:44px}' +
    '.action-btn-primary{background:#059669;color:white}' +
    '.action-btn-primary:hover{background:#047857}' +
    '.action-btn-secondary{background:#f3f4f6;color:#374151}' +
    '.action-btn-secondary:hover{background:#e5e7eb}' +

    // List items for responses (clickable)
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:white;padding:15px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.06);cursor:pointer;transition:all 0.2s}' +
    '.list-item:hover{box-shadow:0 4px 8px rgba(0,0,0,0.1);transform:translateY(-1px)}' +
    '.list-item-header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}' +
    '.list-item-main{flex:1;min-width:200px}' +
    '.list-item-title{font-weight:600;color:#1f2937;margin-bottom:3px}' +
    '.list-item-subtitle{font-size:12px;color:#666}' +
    '.list-item-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee}' +
    '.list-item.expanded .list-item-details{display:block}' +
    '.detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}' +
    '.detail-item{font-size:12px}' +
    '.detail-item-label{color:#666;margin-bottom:2px}' +
    '.detail-item-value{font-weight:600;color:#1f2937}' +

    // Search input
    '.search-container{position:relative;margin-bottom:15px}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:2px solid #e5e7eb;border-radius:8px;font-size:14px;transition:border-color 0.2s}' +
    '.search-input:focus{outline:none;border-color:#059669}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af}' +

    // Filter buttons
    '.filter-group{display:flex;gap:8px;margin-bottom:15px;flex-wrap:wrap}' +

    // Charts section
    '.chart-container{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:15px}' +
    '.chart-title{font-weight:600;color:#1f2937;margin-bottom:15px;font-size:14px}' +
    '.bar-chart{display:flex;flex-direction:column;gap:10px}' +
    '.bar-row{display:flex;align-items:center;gap:10px}' +
    '.bar-label{width:140px;font-size:12px;color:#666;text-align:right}' +
    '.bar-container{flex:1;background:#e5e7eb;border-radius:4px;height:24px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:4px;transition:width 0.5s;display:flex;align-items:center;justify-content:flex-end;padding-right:8px}' +
    '.bar-value{width:50px;font-size:12px;font-weight:600;color:#374151}' +
    '.bar-inner-value{font-size:11px;font-weight:600;color:white}' +

    // Gauge chart
    '.gauge-container{display:flex;flex-wrap:wrap;gap:20px;justify-content:center}' +
    '.gauge{text-align:center;padding:15px}' +
    '.gauge-value{font-size:36px;font-weight:bold;margin-bottom:5px}' +
    '.gauge-label{font-size:12px;color:#666}' +
    '.gauge-ring{width:100px;height:100px;border-radius:50%;margin:0 auto 10px;position:relative;display:flex;align-items:center;justify-content:center}' +
    '.gauge-ring::before{content:"";position:absolute;inset:8px;background:white;border-radius:50%}' +
    '.gauge-ring span{position:relative;z-index:1;font-size:24px;font-weight:bold}' +

    // Trend arrows
    '.trend-up{color:#059669}' +
    '.trend-down{color:#dc2626}' +
    '.trend-neutral{color:#6b7280}' +

    // Insights card
    '.insight-card{background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-left:4px solid #059669;padding:15px;border-radius:0 8px 8px 0;margin-bottom:12px}' +
    '.insight-card.warning{background:linear-gradient(135deg,#fef3c7,#fde68a);border-left-color:#f59e0b}' +
    '.insight-card.alert{background:linear-gradient(135deg,#fee2e2,#fecaca);border-left-color:#dc2626}' +
    '.insight-title{font-weight:600;color:#1f2937;margin-bottom:5px}' +
    '.insight-text{font-size:13px;color:#374151}' +

    // Heatmap styles
    '.heatmap-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(80px,1fr));gap:8px}' +
    '.heatmap-cell{padding:12px;border-radius:8px;text-align:center;font-weight:600;font-size:14px}' +

    // Empty state
    '.empty-state{text-align:center;padding:40px;color:#9ca3af}' +
    '.empty-state-icon{font-size:48px;margin-bottom:10px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e5e7eb;border-top-color:#059669;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Responsive
    '@media (max-width:600px){' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .list-item{flex-direction:column;align-items:flex-start}' +
    '  .tab-icon{font-size:16px}' +
    '  .bar-label{width:100px}' +
    '  .gauge-container{flex-direction:column;align-items:center}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header
    '<div class="header">' +
    '<h1>📊 Member Satisfaction</h1>' +
    '<div class="subtitle">Survey results and satisfaction trends</div>' +
    '</div>' +

    // Tab Navigation
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\',this)" id="tab-overview"><span class="tab-icon">📊</span>Overview</button>' +
    '<button class="tab" onclick="switchTab(\'responses\',this)" id="tab-responses"><span class="tab-icon">📝</span>Responses</button>' +
    '<button class="tab" onclick="switchTab(\'sections\',this)" id="tab-sections"><span class="tab-icon">📈</span>By Section</button>' +
    '<button class="tab" onclick="switchTab(\'analytics\',this)" id="tab-analytics"><span class="tab-icon">🔍</span>Insights</button>' +
    '</div>' +

    // Overview Tab
    '<div class="tab-content active" id="content-overview">' +
    '<div class="stats-grid" id="overview-stats"><div class="loading"><div class="spinner"></div><p>Loading stats...</p></div></div>' +
    '<div id="overview-gauges"></div>' +
    '<div id="overview-insights" style="margin-top:15px"></div>' +
    '</div>' +

    // Responses Tab
    '<div class="tab-content" id="content-responses">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="response-search" placeholder="Search by worksite or role..." oninput="filterResponses(this.value)"></div>' +
    '<div class="filter-group">' +
    '<button class="action-btn action-btn-primary" onclick="filterResponsesBy(\'all\')">All</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'high\')">High Satisfaction</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'mid\')">Medium</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'low\')">Needs Attention</button>' +
    '</div>' +
    '<div class="list-container" id="responses-list"><div class="loading"><div class="spinner"></div><p>Loading responses...</p></div></div>' +
    '</div>' +

    // Sections Tab
    '<div class="tab-content" id="content-sections">' +
    '<div id="sections-charts"><div class="loading"><div class="spinner"></div><p>Loading section scores...</p></div></div>' +
    '</div>' +

    // Analytics Tab
    '<div class="tab-content" id="content-analytics">' +
    '<div id="analytics-content"><div class="loading"><div class="spinner"></div><p>Loading insights...</p></div></div>' +
    '</div>' +

    // JavaScript
    '<script>' +
    'var allResponses=[];var currentFilter="all";var analyticsLoaded=false;var sectionsLoaded=false;' +

    // Error handler helper
    'function handleLoadError(containerId,msg,err){' +
    '  var c=document.getElementById(containerId);' +
    '  if(c)c.innerHTML="<div class=\\"empty-state\\" style=\\"color:#dc2626\\"><div class=\\"empty-state-icon\\">⚠️</div><p>"+msg+"</p><p style=\\"font-size:12px\\">"+((err&&err.message)||"Please try again")+"</p><button onclick=\\"location.reload()\\" style=\\"margin-top:10px;padding:8px 16px;border:none;border-radius:6px;background:#3b82f6;color:white;cursor:pointer\\">Retry</button></div>";' +
    '}' +

    // Tab switching
    'function switchTab(tabName,btn){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  document.getElementById("content-"+tabName).classList.add("active");' +
    '  if(tabName==="responses"&&allResponses.length===0)loadResponses();' +
    '  if(tabName==="sections"&&!sectionsLoaded)loadSections();' +
    '  if(tabName==="analytics"&&!analyticsLoaded)loadAnalytics();' +
    '}' +

    // Score color helper
    'function getScoreClass(score){' +
    '  if(score>=7)return"high";' +
    '  if(score>=5)return"mid";' +
    '  return"low";' +
    '}' +
    'function getScoreColor(score){' +
    '  if(score>=7)return"#059669";' +
    '  if(score>=5)return"#f59e0b";' +
    '  return"#dc2626";' +
    '}' +
    'function getProgressClass(score){' +
    '  if(score>=7)return"progress-green";' +
    '  if(score>=5)return"progress-yellow";' +
    '  return"progress-red";' +
    '}' +

    // Load overview data
    'function loadOverview(){' +
    '  google.script.run.withSuccessHandler(function(data){renderOverview(data)}).withFailureHandler(function(e){handleLoadError("overview-stats","Failed to load overview",e)}).getSatisfactionOverviewData();' +
    '}' +

    // Render overview
    'function renderOverview(data){' +
    '  var html="";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.totalResponses+"</div><div class=\\"stat-label\\">Total Responses</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.avgOverall.toFixed(1)+"</div><div class=\\"stat-label\\">Avg Satisfaction</div></div>";' +
    '  html+="<div class=\\"stat-card blue\\"><div class=\\"stat-value\\">"+data.npsScore+"</div><div class=\\"stat-label\\">Loyalty Score</div></div>";' +
    '  html+="<div class=\\"stat-card purple\\"><div class=\\"stat-value\\">"+data.responseRate+"</div><div class=\\"stat-label\\">Response Rate</div></div>";' +
    '  html+="<div class=\\"stat-card "+(data.avgSteward>=7?"green":data.avgSteward>=5?"orange":"red")+"\\"><div class=\\"stat-value\\">"+data.avgSteward.toFixed(1)+"</div><div class=\\"stat-label\\">Steward Rating</div></div>";' +
    '  html+="<div class=\\"stat-card "+(data.avgLeadership>=7?"green":data.avgLeadership>=5?"orange":"red")+"\\"><div class=\\"stat-value\\">"+data.avgLeadership.toFixed(1)+"</div><div class=\\"stat-label\\">Leadership</div></div>";' +
    '  document.getElementById("overview-stats").innerHTML=html;' +
    // Gauge display
    '  var gauges="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📊 Key Metrics at a Glance</div><div class=\\"gauge-container\\">";' +
    '  gauges+=renderGauge(data.avgOverall,"Overall\\nSatisfaction");' +
    '  gauges+=renderGauge(data.avgTrust,"Trust in\\nUnion");' +
    '  gauges+=renderGauge(data.avgProtected,"Feel\\nProtected");' +
    '  gauges+=renderGauge(data.avgRecommend,"Would\\nRecommend");' +
    '  gauges+="</div></div>";' +
    '  document.getElementById("overview-gauges").innerHTML=gauges;' +
    // Insights - add Loyalty Score explanation first
    '  var insights="";' +
    '  insights+="<div class=\\"insight-card\\" style=\\"background:linear-gradient(135deg,#eff6ff,#dbeafe);border-left-color:#2563eb\\"><div class=\\"insight-title\\">ℹ️ Understanding Loyalty Score</div><div class=\\"insight-text\\">The <strong>Loyalty Score</strong> (ranging from -100 to +100) measures how likely members are to recommend the union. <strong>50+</strong> = Excellent (many advocates), <strong>0-49</strong> = Good (room for growth), <strong>Below 0</strong> = Needs work (more critics than advocates). It\'s based on the \\"Would Recommend\\" question.</div></div>";' +
    '  if(data.insights&&data.insights.length>0){' +
    '    data.insights.forEach(function(i){' +
    '      insights+="<div class=\\"insight-card "+i.type+"\\"><div class=\\"insight-title\\">"+i.icon+" "+i.title+"</div><div class=\\"insight-text\\">"+i.text+"</div></div>";' +
    '    });' +
    '  }' +
    '  document.getElementById("overview-insights").innerHTML=insights;' +
    '}' +

    // Render gauge
    'function renderGauge(value,label){' +
    '  var color=getScoreColor(value);' +
    '  var pct=value*10;' +
    '  return"<div class=\\"gauge\\"><div class=\\"gauge-ring\\" style=\\"background:conic-gradient("+color+" "+pct+"%,#e5e7eb "+pct+"%)\\"><span style=\\"color:"+color+"\\">"+value.toFixed(1)+"</span></div><div class=\\"gauge-label\\">"+label.replace("\\n","<br>")+"</div></div>";' +
    '}' +

    // Load responses
    'function loadResponses(){' +
    '  google.script.run.withSuccessHandler(function(data){allResponses=data||[];renderResponses(data)}).withFailureHandler(function(e){handleLoadError("responses-list","Failed to load responses",e)}).getSatisfactionResponseData();' +
    '}' +

    // Render responses with clickable details
    'function renderResponses(data){' +
    '  var c=document.getElementById("responses-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">📝</div><p>No responses found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(r,i){' +
    '    var scoreClass=getScoreClass(r.avgScore);' +
    '    var scoreColor=getScoreColor(r.avgScore);' +
    '    return"<div class=\\"list-item\\" onclick=\\"toggleResponse(this)\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+r.worksite+" - "+r.role+"</div><div class=\\"list-item-subtitle\\">"+r.shift+" • "+r.timeInRole+" • "+r.date+"</div></div><div><span class=\\"score-indicator score-"+scoreClass+"\\" style=\\"color:"+scoreColor+"\\">"+r.avgScore.toFixed(1)+"/10</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-grid\\">' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Satisfaction</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.satisfaction)+"\\">"+r.satisfaction+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Trust in Union</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.trust)+"\\">"+r.trust+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Feel Protected</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.protected)+"\\">"+r.protected+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Would Recommend</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.recommend)+"\\">"+r.recommend+"/10</div></div>' +
    '          "+(r.stewardContact?"<div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Steward Contact</div><div class=\\"detail-item-value\\">Yes</div></div>":"")+"' +
    '          "+(r.stewardRating>0?"<div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Steward Rating</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.stewardRating)+"\\">"+r.stewardRating.toFixed(1)+"/10</div></div>":"")+"' +
    '        </div>' +
    '      </div>' +
    '    </div>";' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" responses. Use search/filters to narrow.</p></div>";' +
    '}' +

    // Toggle response details
    'function toggleResponse(el){el.classList.toggle("expanded")}' +

    // Filter responses
    'function filterResponses(query){' +
    '  if(!query||query.length<2){applyFilters();return}' +
    '  query=query.toLowerCase();' +
    '  var filtered=allResponses.filter(function(r){return r.worksite.toLowerCase().indexOf(query)>=0||r.role.toLowerCase().indexOf(query)>=0||r.shift.toLowerCase().indexOf(query)>=0});' +
    '  if(currentFilter!=="all")filtered=applyScoreFilter(filtered,currentFilter);' +
    '  renderResponses(filtered);' +
    '}' +

    // Filter by satisfaction level
    'function filterResponsesBy(level){' +
    '  currentFilter=level;' +
    '  applyFilters();' +
    '}' +

    // Apply filters
    'function applyFilters(){' +
    '  var query=document.getElementById("response-search").value.toLowerCase();' +
    '  var filtered=allResponses;' +
    '  if(currentFilter!=="all")filtered=applyScoreFilter(filtered,currentFilter);' +
    '  if(query&&query.length>=2)filtered=filtered.filter(function(r){return r.worksite.toLowerCase().indexOf(query)>=0||r.role.toLowerCase().indexOf(query)>=0});' +
    '  renderResponses(filtered);' +
    '}' +

    // Score filter helper
    'function applyScoreFilter(data,level){' +
    '  return data.filter(function(r){' +
    '    if(level==="high")return r.avgScore>=7;' +
    '    if(level==="mid")return r.avgScore>=5&&r.avgScore<7;' +
    '    if(level==="low")return r.avgScore<5;' +
    '    return true;' +
    '  });' +
    '}' +

    // Load sections data
    'function loadSections(){' +
    '  sectionsLoaded=true;' +
    '  google.script.run.withSuccessHandler(function(data){renderSections(data)}).withFailureHandler(function(e){sectionsLoaded=false;handleLoadError("sections-charts","Failed to load sections",e)}).getSatisfactionSectionData();' +
    '}' +

    // Render sections
    'function renderSections(data){' +
    '  var c=document.getElementById("sections-charts");' +
    '  var html="";' +
    '  if(!data.sections||data.sections.length===0){c.innerHTML="<div class=\\"empty-state\\">No section data available</div>";return}' +
    // Section scores bar chart - scale to actual data range, not always 0-10
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📊 Average Score by Section (1-10 Scale)</div>";' +
    '  html+="<div style=\\"font-size:11px;color:#666;margin-bottom:12px\\">Sorted by score - areas needing attention shown first</div>";' +
    '  html+="<div class=\\"bar-chart\\">";' +
    '  var maxScore=10;' +
    '  var hasValidData=data.sections.some(function(s){return s.avg>0&&s.responseCount>0});' +
    '  if(!hasValidData){html+="<div class=\\"empty-state\\">No survey responses yet</div>";}else{' +
    '  data.sections.forEach(function(s){' +
    '    if(s.responseCount===0)return;' +  // Skip sections with no data
    '    var pct=Math.max(0,Math.min(100,(s.avg/maxScore)*100));' +  // Clamp to 0-100%
    '    var color=getScoreColor(s.avg);' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+s.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+s.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+s.responseCount+" responses</div></div>";' +
    '  });' +
    '  }' +
    '  html+="</div></div>";' +
    // Summary insights instead of redundant detail cards
    '  var lowScoring=data.sections.filter(function(s){return s.avg>0&&s.avg<6&&s.responseCount>0});' +
    '  var highScoring=data.sections.filter(function(s){return s.avg>=8&&s.responseCount>0});' +
    '  if(lowScoring.length>0||highScoring.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">💡 Section Insights</div>";' +
    '    if(lowScoring.length>0){' +
    '      html+="<div class=\\"insight-card warning\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">⚠️ Areas Needing Attention</div><div class=\\"insight-text\\">";' +
    '      lowScoring.forEach(function(s,i){html+=(i>0?", ":"")+s.name+" ("+s.avg.toFixed(1)+")"});' +
    '      html+="</div></div>";' +
    '    }' +
    '    if(highScoring.length>0){' +
    '      html+="<div class=\\"insight-card success\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">✅ Strong Performance</div><div class=\\"insight-text\\">";' +
    '      highScoring.forEach(function(s,i){html+=(i>0?", ":"")+s.name+" ("+s.avg.toFixed(1)+")"});' +
    '      html+="</div></div>";' +
    '    }' +
    '    html+="</div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Load analytics
    'function loadAnalytics(){' +
    '  analyticsLoaded=true;' +
    '  google.script.run.withSuccessHandler(function(data){renderAnalytics(data)}).withFailureHandler(function(e){analyticsLoaded=false;handleLoadError("analytics-content","Failed to load analytics",e)}).getSatisfactionAnalyticsData();' +
    '}' +

    // Render analytics/insights
    'function renderAnalytics(data){' +
    '  var c=document.getElementById("analytics-content");' +
    '  var html="";' +
    // Key insights
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">💡 Key Insights</div>";' +
    '  if(data.insights&&data.insights.length>0){' +
    '    data.insights.forEach(function(i){' +
    '      html+="<div class=\\"insight-card "+i.type+"\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">"+i.icon+" "+i.title+"</div><div class=\\"insight-text\\">"+i.text+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No insights available</div>";}' +
    '  html+="</div>";' +
    // By worksite breakdown
    '  if(data.byWorksite&&data.byWorksite.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Satisfaction by Worksite</div><div class=\\"bar-chart\\">";' +
    '    data.byWorksite.forEach(function(w){' +
    '      var pct=(w.avg/10)*100;' +
    '      var color=getScoreColor(w.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+w.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+w.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+w.count+" responses</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    // By role breakdown
    '  if(data.byRole&&data.byRole.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👤 Satisfaction by Role</div><div class=\\"bar-chart\\">";' +
    '    data.byRole.forEach(function(r){' +
    '      var pct=(r.avg/10)*100;' +
    '      var color=getScoreColor(r.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+r.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+r.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+r.count+" responses</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    // Steward contact impact
    '  if(data.stewardImpact){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🤝 Impact of Steward Contact</div>";' +
    '    html+="<div class=\\"stats-grid\\">";' +
    '    html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.stewardImpact.withContact.toFixed(1)+"</div><div class=\\"stat-label\\">With Steward Contact ("+data.stewardImpact.withContactCount+" members)</div></div>";' +
    '    html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.stewardImpact.withoutContact.toFixed(1)+"</div><div class=\\"stat-label\\">Without Contact ("+data.stewardImpact.withoutContactCount+" members)</div></div>";' +
    '    html+="</div>";' +
    '    var diff=data.stewardImpact.withContact-data.stewardImpact.withoutContact;' +
    '    if(diff>0){' +
    '      html+="<div class=\\"insight-card\\" style=\\"margin-top:10px\\"><div class=\\"insight-text\\">Members with steward contact report <strong>+"+diff.toFixed(1)+"</strong> higher satisfaction on average.</div></div>";' +
    '    }' +
    '    html+="</div>";' +
    '  }' +
    // Top priorities
    '  if(data.topPriorities&&data.topPriorities.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🎯 Top Member Priorities</div><div class=\\"bar-chart\\">";' +
    '    var maxP=Math.max.apply(null,data.topPriorities.map(function(p){return p.count}))||1;' +
    '    data.topPriorities.forEach(function(p){' +
    '      var pct=(p.count/maxP)*100;' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+p.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:#7C3AED\\"></div></div><div class=\\"bar-value\\">"+p.count+"</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Initialize
    'loadOverview();' +
    '</script>' +

    '</body></html>';
}

/**
 * Get overview data for satisfaction dashboard
 */
function getSatisfactionOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var data = {
    totalResponses: 0,
    avgOverall: 0,
    avgSteward: 0,
    avgLeadership: 0,
    avgTrust: 0,
    avgProtected: 0,
    avgRecommend: 0,
    npsScore: 0,
    responseRate: 'N/A',
    insights: [],
    distribution: { high: 0, mid: 0, low: 0 }
  };

  if (!sheet) return data;

  // Optimized: Use getLastRow() instead of iterating through A:A column
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return data;

  data.totalResponses = lastRow - 1;

  // Get satisfaction scores (Q6-Q9 are columns G-J, 1-indexed as 7-10)
  var satisfactionRange = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, data.totalResponses, 4).getValues();

  var sumOverall = 0, sumTrust = 0, sumProtected = 0, sumRecommend = 0;
  var promoters = 0, detractors = 0;
  var validCount = 0;

  satisfactionRange.forEach(function(row) {
    var satisfied = parseFloat(row[0]) || 0;
    var trust = parseFloat(row[1]) || 0;
    var protected_ = parseFloat(row[2]) || 0;
    var recommend = parseFloat(row[3]) || 0;

    if (satisfied > 0) {
      sumOverall += satisfied;
      sumTrust += trust;
      sumProtected += protected_;
      sumRecommend += recommend;
      validCount++;

      // NPS calculation (based on recommend score 1-10)
      if (recommend >= 9) promoters++;
      else if (recommend <= 6) detractors++;

      // Calculate distribution based on average score
      var avgScore = (satisfied + trust + protected_ + recommend) / 4;
      if (avgScore >= 7) data.distribution.high++;
      else if (avgScore >= 5) data.distribution.mid++;
      else data.distribution.low++;
    }
  });

  if (validCount > 0) {
    data.avgOverall = sumOverall / validCount;
    data.avgTrust = sumTrust / validCount;
    data.avgProtected = sumProtected / validCount;
    data.avgRecommend = sumRecommend / validCount;
    data.npsScore = Math.round(((promoters - detractors) / validCount) * 100);
  }

  // Get steward ratings (Q10-Q16, columns K-Q)
  var stewardRange = sheet.getRange(2, SATISFACTION_COLS.Q10_TIMELY_RESPONSE, data.totalResponses, 7).getValues();
  var sumSteward = 0, stewardCount = 0;

  stewardRange.forEach(function(row) {
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      sumSteward += rowSum / rowCount;
      stewardCount++;
    }
  });

  if (stewardCount > 0) {
    data.avgSteward = sumSteward / stewardCount;
  }

  // Get leadership ratings (Q26-Q31, columns AA-AF)
  var leadershipRange = sheet.getRange(2, SATISFACTION_COLS.Q26_DECISIONS_CLEAR, data.totalResponses, 6).getValues();
  var sumLeadership = 0, leadershipCount = 0;

  leadershipRange.forEach(function(row) {
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      sumLeadership += rowSum / rowCount;
      leadershipCount++;
    }
  });

  if (leadershipCount > 0) {
    data.avgLeadership = sumLeadership / leadershipCount;
  }

  // Calculate response rate if we have member directory
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var totalMembers = memberSheet.getLastRow() - 1;
    var rate = Math.round((data.totalResponses / totalMembers) * 100);
    data.responseRate = rate + '%';
  }

  // Generate insights
  if (data.avgOverall >= 8) {
    data.insights.push({
      type: '',
      icon: '🌟',
      title: 'High Overall Satisfaction',
      text: 'Members report strong satisfaction with union representation (avg ' + data.avgOverall.toFixed(1) + '/10).'
    });
  } else if (data.avgOverall < 5) {
    data.insights.push({
      type: 'alert',
      icon: '⚠️',
      title: 'Low Satisfaction Alert',
      text: 'Overall satisfaction is below target at ' + data.avgOverall.toFixed(1) + '/10. Consider reviewing member concerns.'
    });
  }

  if (data.npsScore >= 50) {
    data.insights.push({
      type: '',
      icon: '🎯',
      title: 'Members Highly Recommend',
      text: 'Loyalty Score of ' + data.npsScore + ' means members actively recommend the union to colleagues.'
    });
  } else if (data.npsScore >= 0) {
    data.insights.push({
      type: '',
      icon: '📊',
      title: 'Moderate Member Loyalty',
      text: 'Loyalty Score of ' + data.npsScore + ' shows members are neutral. Focus on converting neutral members to advocates.'
    });
  } else {
    data.insights.push({
      type: 'warning',
      icon: '⚠️',
      title: 'Member Loyalty Needs Attention',
      text: 'Loyalty Score of ' + data.npsScore + ' indicates more critics than advocates. Address member concerns to improve.'
    });
  }

  if (data.avgSteward >= 8) {
    data.insights.push({
      type: '',
      icon: '🤝',
      title: 'Excellent Steward Performance',
      text: 'Stewards are rated highly at ' + data.avgSteward.toFixed(1) + '/10 on average.'
    });
  } else if (data.avgSteward < 6 && stewardCount > 0) {
    data.insights.push({
      type: 'warning',
      icon: '👤',
      title: 'Steward Training Opportunity',
      text: 'Steward ratings averaging ' + data.avgSteward.toFixed(1) + '/10 suggest room for improvement.'
    });
  }

  return data;
}

/**
 * Get individual response data for satisfaction dashboard
 */
function getSatisfactionResponseData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (!sheet) return [];

  // Optimized: Use getLastRow() instead of iterating through A:A column
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  // Optimized: Batch read all needed columns in single range
  // Determine the range that covers all columns (1 to Q10 + 6)
  var lastCol = SATISFACTION_COLS.Q10_TIMELY_RESPONSE + 6;
  var allData = sheet.getRange(2, 1, numRows, lastCol).getValues();

  // Extract columns from batch data (column indices are 0-based in array)
  var timestampData = allData.map(function(row) { return [row[0]]; });
  var worksiteData = allData.map(function(row) { return [row[SATISFACTION_COLS.Q1_WORKSITE - 1]]; });
  var roleData = allData.map(function(row) { return [row[SATISFACTION_COLS.Q2_ROLE - 1]]; });
  var shiftData = allData.map(function(row) { return [row[SATISFACTION_COLS.Q3_SHIFT - 1]]; });
  var timeData = allData.map(function(row) { return [row[SATISFACTION_COLS.Q4_TIME_IN_ROLE - 1]]; });
  var stewardContactData = allData.map(function(row) { return [row[SATISFACTION_COLS.Q5_STEWARD_CONTACT - 1]]; });
  var satisfactionData = allData.map(function(row) {
    var start = SATISFACTION_COLS.Q6_SATISFIED_REP - 1;
    return row.slice(start, start + 4);
  });
  var stewardRatingsData = allData.map(function(row) {
    var start = SATISFACTION_COLS.Q10_TIMELY_RESPONSE - 1;
    return row.slice(start, start + 7);
  });

  var responses = [];
  for (var i = 0; i < numRows; i++) {
    // Get individual scores
    var satisfaction = parseFloat(satisfactionData[i][0]) || 0;
    var trust = parseFloat(satisfactionData[i][1]) || 0;
    var protected_ = parseFloat(satisfactionData[i][2]) || 0;
    var recommend = parseFloat(satisfactionData[i][3]) || 0;

    // Calculate average satisfaction score
    var sum = 0, count = 0;
    [satisfaction, trust, protected_, recommend].forEach(function(s) {
      if (s > 0) { sum += s; count++; }
    });
    var avgScore = count > 0 ? sum / count : 0;

    // Calculate steward rating average
    var stewardSum = 0, stewardCount = 0;
    stewardRatingsData[i].forEach(function(s) {
      var v = parseFloat(s);
      if (v > 0) { stewardSum += v; stewardCount++; }
    });
    var stewardRating = stewardCount > 0 ? stewardSum / stewardCount : 0;

    var ts = timestampData[i][0];
    var dateStr = ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : (ts || 'N/A');
    var stewardContact = stewardContactData[i][0];

    responses.push({
      worksite: worksiteData[i][0] || 'Unknown',
      role: roleData[i][0] || 'Unknown',
      shift: shiftData[i][0] || 'N/A',
      timeInRole: timeData[i][0] || 'N/A',
      date: dateStr,
      avgScore: avgScore,
      satisfaction: satisfaction,
      trust: trust,
      protected: protected_,
      recommend: recommend,
      stewardContact: stewardContact === 'Yes',
      stewardRating: stewardRating
    });
  }

  // Sort by date (most recent first)
  responses.sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  return responses;
}

/**
 * Get section-level data for satisfaction dashboard
 */
function getSatisfactionSectionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { sections: [] };
  if (!sheet) return result;

  // Optimized: Use getLastRow() instead of iterating through A:A column
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;

  // Define sections with their column ranges
  var sectionDefs = [
    { name: 'Overall Satisfaction', startCol: SATISFACTION_COLS.Q6_SATISFIED_REP, numCols: 4 },
    { name: 'Steward Ratings', startCol: SATISFACTION_COLS.Q10_TIMELY_RESPONSE, numCols: 7 },
    { name: 'Steward Access', startCol: SATISFACTION_COLS.Q18_KNOW_CONTACT, numCols: 3 },
    { name: 'Chapter Effectiveness', startCol: SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, numCols: 5 },
    { name: 'Local Leadership', startCol: SATISFACTION_COLS.Q26_DECISIONS_CLEAR, numCols: 6 },
    { name: 'Contract Enforcement', startCol: SATISFACTION_COLS.Q32_ENFORCES_CONTRACT, numCols: 4 },
    { name: 'Representation Process', startCol: SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS, numCols: 4 },
    { name: 'Communication Quality', startCol: SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, numCols: 5 },
    { name: 'Member Voice & Culture', startCol: SATISFACTION_COLS.Q46_VOICE_MATTERS, numCols: 5 },
    { name: 'Value & Collective Action', startCol: SATISFACTION_COLS.Q51_GOOD_VALUE, numCols: 5 },
    { name: 'Scheduling/Office Days', startCol: SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES, numCols: 7 }
  ];

  sectionDefs.forEach(function(section) {
    var data = sheet.getRange(2, section.startCol, numRows, section.numCols).getValues();
    var sum = 0, count = 0;

    data.forEach(function(row) {
      row.forEach(function(val) {
        var v = parseFloat(val);
        if (v > 0 && v <= 10) {
          sum += v;
          count++;
        }
      });
    });

    result.sections.push({
      name: section.name,
      avg: count > 0 ? sum / count : 0,
      responseCount: Math.floor(count / section.numCols),
      questions: section.numCols
    });
  });

  // Sort by score (lowest first to highlight areas needing attention)
  result.sections.sort(function(a, b) { return a.avg - b.avg; });

  return result;
}

/**
 * Get trend data for satisfaction dashboard - responses over time
 */
function getSatisfactionTrendData(period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    byMonth: [],
    satisfactionTrend: [],
    issuesTrend: [],
    totalInPeriod: 0
  };

  if (!sheet) return result;

  // Get data
  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

  // Filter by period
  var now = new Date();
  var cutoff = null;
  if (period === '30') cutoff = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  else if (period === '90') cutoff = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
  else if (period === 'year') cutoff = new Date(now.getFullYear(), 0, 1);

  // Group by month
  var monthData = {};
  for (var i = 0; i < numRows; i++) {
    var ts = timestamps[i][0];
    if (!(ts instanceof Date)) continue;
    if (cutoff && ts < cutoff) continue;

    var monthKey = Utilities.formatDate(ts, tz, 'yyyy-MM');
    var monthLabel = Utilities.formatDate(ts, tz, 'MMM yy');

    if (!monthData[monthKey]) {
      monthData[monthKey] = { label: monthLabel, count: 0, sum: 0, validCount: 0 };
    }

    monthData[monthKey].count++;
    result.totalInPeriod++;

    // Calculate avg satisfaction
    var row = satisfactionData[i];
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      monthData[monthKey].sum += rowSum / rowCount;
      monthData[monthKey].validCount++;
    }
  }

  // Convert to arrays sorted by date
  var months = Object.keys(monthData).sort();
  months.forEach(function(key) {
    var m = monthData[key];
    result.byMonth.push({ label: m.label, count: m.count });
    result.satisfactionTrend.push({
      label: m.label,
      avg: m.validCount > 0 ? m.sum / m.validCount : 0
    });
  });

  // Get common issues/priorities for trend
  try {
    var prioritiesData = sheet.getRange(2, SATISFACTION_COLS.Q64_TOP_PRIORITIES, numRows, 1).getValues();
    var issueMap = {};
    for (var i = 0; i < numRows; i++) {
      var ts = timestamps[i][0];
      if (!(ts instanceof Date)) continue;
      if (cutoff && ts < cutoff) continue;

      var priorities = String(prioritiesData[i][0] || '');
      if (priorities) {
        priorities.split(',').forEach(function(item) {
          var p = item.trim();
          if (p) issueMap[p] = (issueMap[p] || 0) + 1;
        });
      }
    }
    for (var issue in issueMap) {
      result.issuesTrend.push({ name: issue, count: issueMap[issue] });
    }
    result.issuesTrend.sort(function(a, b) { return b.count - a.count; });
  } catch(e) { /* ignore if column doesn't exist */ }

  return result;
}

/**
 * Get breakdown data for satisfaction dashboard
 */
function getSatisfactionBreakdownData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    sections: [],
    byWorksite: [],
    byRole: []
  };

  if (!sheet) return result;

  // Get sections data
  var sectionResult = getSatisfactionSectionData();
  result.sections = sectionResult.sections;

  // Get analytics data for worksite/role
  var analyticsData = getSatisfactionAnalyticsData();
  result.byWorksite = analyticsData.byWorksite;
  result.byRole = analyticsData.byRole;

  return result;
}

/**
 * Get insights data for satisfaction dashboard
 */
function getSatisfactionInsightsData() {
  var analyticsData = getSatisfactionAnalyticsData();
  var overviewData = getSatisfactionOverviewData();

  var result = {
    insights: analyticsData.insights || [],
    stewardImpact: analyticsData.stewardImpact,
    topPriorities: analyticsData.topPriorities
  };

  // Add additional insights based on overview data
  if (overviewData.avgOverall >= 8) {
    result.insights.unshift({
      type: 'success',
      icon: '🌟',
      title: 'Excellent Overall Satisfaction',
      text: 'Members report high satisfaction (' + overviewData.avgOverall.toFixed(1) + '/10). Keep up the great work!'
    });
  } else if (overviewData.avgOverall < 5) {
    result.insights.unshift({
      type: 'alert',
      icon: '⚠️',
      title: 'Satisfaction Needs Attention',
      text: 'Overall satisfaction is below target at ' + overviewData.avgOverall.toFixed(1) + '/10. Review member feedback for areas to improve.'
    });
  }

  return result;
}

/**
 * Get drill-down data for specific categories
 */
function getSatisfactionDrillData(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { items: [] };
  if (!sheet) return result;

  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  if (type === 'responses') {
    // Show recent responses
    var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
    var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
    var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

    for (var i = 0; i < numRows; i++) {
      var ts = timestamps[i][0];
      var row = satisfactionData[i];
      var sum = 0, count = 0;
      row.forEach(function(val) { var v = parseFloat(val); if (v > 0) { sum += v; count++; } });
      var avg = count > 0 ? sum / count : 0;

      result.items.push({
        label: worksiteData[i][0] || 'Unknown',
        detail: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : 'N/A',
        score: avg
      });
    }
    result.items.sort(function(a, b) { return b.score - a.score; });
    result.items = result.items.slice(0, 20);
  }

  return result;
}

/**
 * Get location-specific drill-down data
 */
function getSatisfactionLocationDrill(location) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { count: 0, avgScore: 0, responses: [] };
  if (!sheet || !location) return result;

  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
  var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
  var roleData = sheet.getRange(2, SATISFACTION_COLS.Q2_ROLE, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

  var totalScore = 0;

  for (var i = 0; i < numRows; i++) {
    if (worksiteData[i][0] !== location) continue;

    var ts = timestamps[i][0];
    var row = satisfactionData[i];
    var sum = 0, count = 0;
    row.forEach(function(val) { var v = parseFloat(val); if (v > 0) { sum += v; count++; } });
    var avg = count > 0 ? sum / count : 0;

    result.count++;
    totalScore += avg;

    result.responses.push({
      role: roleData[i][0] || 'Unknown',
      date: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : 'N/A',
      avgScore: avg
    });
  }

  result.avgScore = result.count > 0 ? totalScore / result.count : 0;
  result.responses.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

  return result;
}

/**
 * Helper function to get last row with data
 */
function getSheetLastRow(sheet) {
  // Optimized: Use built-in getLastRow() instead of iterating through A:A column
  return sheet.getLastRow();
}

/**
 * Get analytics data for satisfaction dashboard insights
 */
function getSatisfactionAnalyticsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    insights: [],
    byWorksite: [],
    byRole: [],
    stewardImpact: null,
    topPriorities: []
  };

  if (!sheet) return result;

  // Optimized: Use getLastRow() instead of iterating through A:A column
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;

  // Get all relevant data in one batch
  var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
  var roleData = sheet.getRange(2, SATISFACTION_COLS.Q2_ROLE, numRows, 1).getValues();
  var stewardContactData = sheet.getRange(2, SATISFACTION_COLS.Q5_STEWARD_CONTACT, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();
  var prioritiesData = sheet.getRange(2, SATISFACTION_COLS.Q64_TOP_PRIORITIES, numRows, 1).getValues();

  // Calculate average score for each response
  var scores = [];
  for (var i = 0; i < numRows; i++) {
    var row = satisfactionData[i];
    var sum = 0, count = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { sum += v; count++; }
    });
    scores.push(count > 0 ? sum / count : 0);
  }

  // By Worksite analysis
  var worksiteMap = {};
  for (var i = 0; i < numRows; i++) {
    var ws = worksiteData[i][0] || 'Unknown';
    if (!worksiteMap[ws]) worksiteMap[ws] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      worksiteMap[ws].sum += scores[i];
      worksiteMap[ws].count++;
    }
  }

  for (var ws in worksiteMap) {
    if (worksiteMap[ws].count > 0) {
      result.byWorksite.push({
        name: ws,
        avg: worksiteMap[ws].sum / worksiteMap[ws].count,
        count: worksiteMap[ws].count
      });
    }
  }
  result.byWorksite.sort(function(a, b) { return b.avg - a.avg; });

  // By Role analysis
  var roleMap = {};
  for (var i = 0; i < numRows; i++) {
    var role = roleData[i][0] || 'Unknown';
    if (!roleMap[role]) roleMap[role] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      roleMap[role].sum += scores[i];
      roleMap[role].count++;
    }
  }

  for (var role in roleMap) {
    if (roleMap[role].count > 0) {
      result.byRole.push({
        name: role,
        avg: roleMap[role].sum / roleMap[role].count,
        count: roleMap[role].count
      });
    }
  }
  result.byRole.sort(function(a, b) { return b.avg - a.avg; });

  // Steward contact impact
  var withContactSum = 0, withContactCount = 0;
  var withoutContactSum = 0, withoutContactCount = 0;

  for (var i = 0; i < numRows; i++) {
    var contact = String(stewardContactData[i][0]).toLowerCase();
    if (scores[i] > 0) {
      if (contact === 'yes') {
        withContactSum += scores[i];
        withContactCount++;
      } else if (contact === 'no') {
        withoutContactSum += scores[i];
        withoutContactCount++;
      }
    }
  }

  if (withContactCount > 0 || withoutContactCount > 0) {
    result.stewardImpact = {
      withContact: withContactCount > 0 ? withContactSum / withContactCount : 0,
      withContactCount: withContactCount,
      withoutContact: withoutContactCount > 0 ? withoutContactSum / withoutContactCount : 0,
      withoutContactCount: withoutContactCount
    };
  }

  // Top priorities analysis
  var priorityMap = {};
  for (var i = 0; i < numRows; i++) {
    var priorities = String(prioritiesData[i][0] || '');
    if (priorities) {
      // Split by comma and count each priority
      var items = priorities.split(',');
      items.forEach(function(item) {
        var p = item.trim();
        if (p) {
          priorityMap[p] = (priorityMap[p] || 0) + 1;
        }
      });
    }
  }

  for (var p in priorityMap) {
    result.topPriorities.push({ name: p, count: priorityMap[p] });
  }
  result.topPriorities.sort(function(a, b) { return b.count - a.count; });
  result.topPriorities = result.topPriorities.slice(0, 10); // Top 10

  // Generate insights
  // Lowest scoring worksite
  if (result.byWorksite.length > 0) {
    var lowest = result.byWorksite[result.byWorksite.length - 1];
    if (lowest.avg < 6 && lowest.count >= 3) {
      result.insights.push({
        type: 'warning',
        icon: '📍',
        title: 'Worksite Attention Needed',
        text: lowest.name + ' has the lowest satisfaction score (' + lowest.avg.toFixed(1) + '/10) with ' + lowest.count + ' responses.'
      });
    }
  }

  // Steward impact insight
  if (result.stewardImpact && result.stewardImpact.withContactCount > 0 && result.stewardImpact.withoutContactCount > 0) {
    var diff = result.stewardImpact.withContact - result.stewardImpact.withoutContact;
    if (diff > 1) {
      result.insights.push({
        type: '',
        icon: '🤝',
        title: 'Steward Contact Matters',
        text: 'Members who contacted a steward report ' + diff.toFixed(1) + ' points higher satisfaction on average.'
      });
    }
  }

  // Role insights
  if (result.byRole.length >= 2) {
    var topRole = result.byRole[0];
    var bottomRole = result.byRole[result.byRole.length - 1];
    if (topRole.avg - bottomRole.avg > 2 && bottomRole.count >= 3) {
      result.insights.push({
        type: 'warning',
        icon: '👤',
        title: 'Role Disparity',
        text: bottomRole.name + ' roles report lower satisfaction (' + bottomRole.avg.toFixed(1) + ') than ' + topRole.name + ' (' + topRole.avg.toFixed(1) + ').'
      });
    }
  }

  // Top priority insight
  if (result.topPriorities.length > 0) {
    var topP = result.topPriorities[0];
    result.insights.push({
      type: '',
      icon: '🎯',
      title: 'Top Member Priority',
      text: '"' + topP.name + '" is the most cited priority with ' + topP.count + ' mentions.'
    });
  }

  return result;
}
