/**
 * 10a_SheetCreation.gs - Main Entry Point
 *
 * Core setup functions, menu system, and sheet creation.
 *
 * ⚠️ WARNING: DO NOT DEPLOY THIS FILE DIRECTLY
 * This is a source file used to generate ConsolidatedDashboard.gs.
 * Deploy ONLY ConsolidatedDashboard.gs to avoid function conflicts.
 *
 * @version 4.6.0
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

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Create or recreate the Config sheet with dropdown values
 * Comprehensive configuration with section groupings and organization settings
 */
function createConfigSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG);
  // Only clear if sheet is new or has no meaningful data (≤2 rows = headers only)
  if (sheet.getLastRow() <= 2) {
    sheet.clear();
  } else {
    Logger.log('createConfigSheet: Config sheet has ' + sheet.getLastRow() + ' rows of data — skipping clear to preserve settings');
    return sheet;
  }

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
    '── EXTENDED CONTACT ──', '', '', '', '',           // AN-AR (5 cols)
    '── STRATEGIC COMMAND CENTER ──', '', '', '', '', '', '', // AS-AY (7 cols)
    '── MOBILE DASHBOARD ──'                            // AZ (1 col) - LAST COLUMN
  ];

  // Row 2: Column Headers (A-AZ = 52 columns total)
  var columnHeaders = [
    'Job Titles', 'Office Locations', 'Units', 'Office Days', 'Yes/No (Dropdowns)',       // A-E
    'Supervisors', 'Managers',                                                             // F-G
    'Stewards', 'Steward Committees',                                                      // H-I
    'Grievance Status', 'Grievance Step', 'Issue Category', 'Articles Violated',          // J-M
    'Communication Methods', 'Grievance Coordinators', 'Grievance Form URL', 'Contact Form URL', // N-Q
    'Admin Emails', 'Alert Days Before Deadline', 'Notification Recipients',              // R-T
    'Organization Name', 'Local Number', 'Main Office Address', 'Main Phone',             // U-X
    'Google Drive Folder ID', 'Google Calendar ID',                                       // Y-Z
    'Filing Deadline Days', 'Step I Response Days', 'Step II Appeal Days', 'Step II Response Days', // AA-AD
    'Best Times to Contact', 'Home Towns',                                                // AE-AF
    'Contract Article (Grievance)', 'Contract Article (Discipline)', 'Contract Article (Workload)', 'Contract Name', // AG-AJ
    'Union Parent', 'State/Region', 'Organization Website',                               // AK-AM
    'Office Addresses', 'Main Fax', 'Main Contact Name', 'Main Contact Email', 'Satisfaction Survey URL', // AN-AR
    'Chief Steward Email', 'Unit Codes', 'Archive Folder ID', 'Escalation Statuses', 'Escalation Steps', 'Template ID', 'PDF Folder ID', // AS-AY
    '📱 Mobile Dashboard URL'                                                              // AZ - LAST COLUMN
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
    .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
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

  // Organization defaults (placeholders — update in Config sheet)
  sheet.getRange(3, CONFIG_COLS.ORG_NAME, 1, 1).setValue('Your Union Name');
  sheet.getRange(3, CONFIG_COLS.LOCAL_NUMBER, 1, 1).setValue('000');
  sheet.getRange(3, CONFIG_COLS.MAIN_ADDRESS, 1, 1).setValue('123 Main Street, Suite 100, City, ST 00000');
  sheet.getRange(3, CONFIG_COLS.MAIN_PHONE, 1, 1).setValue('555-000-0000');

  // Deadline defaults (in days) — values from DEADLINE_DEFAULTS (01_Core.gs)
  sheet.getRange(3, CONFIG_COLS.FILING_DEADLINE_DAYS, 1, 1).setValue(DEADLINE_DEFAULTS.FILING_DAYS);
  sheet.getRange(3, CONFIG_COLS.STEP1_RESPONSE_DAYS, 1, 1).setValue(DEADLINE_DEFAULTS.STEP_1_RESPONSE);
  sheet.getRange(3, CONFIG_COLS.STEP2_APPEAL_DAYS, 1, 1).setValue(DEADLINE_DEFAULTS.STEP_2_APPEAL);
  sheet.getRange(3, CONFIG_COLS.STEP2_RESPONSE_DAYS, 1, 1).setValue(DEADLINE_DEFAULTS.STEP_2_RESPONSE);

  // Best Times to Contact (AE)
  var bestTimes = ['Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Weekends', 'Flexible'];
  sheet.getRange(3, CONFIG_COLS.BEST_TIMES, bestTimes.length, 1)
    .setValues(bestTimes.map(function(v) { return [v]; }));

  // Contract articles
  sheet.getRange(3, CONFIG_COLS.CONTRACT_GRIEVANCE, 1, 1).setValue('Article XX');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_DISCIPLINE, 1, 1).setValue('Article YY');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_WORKLOAD, 1, 1).setValue('Article ZZ');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_NAME, 1, 1).setValue('Current CBA');

  // Org identity
  sheet.getRange(3, CONFIG_COLS.UNION_PARENT, 1, 1).setValue('Your Parent Union');
  sheet.getRange(3, CONFIG_COLS.STATE_REGION, 1, 1).setValue('Your State');
  sheet.getRange(3, CONFIG_COLS.ORG_WEBSITE, 1, 1).setValue('https://www.example.org/');

  // Extended contact
  sheet.getRange(3, CONFIG_COLS.MAIN_FAX, 1, 1).setValue('555-000-0001');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_NAME, 1, 1).setValue('Contact Name');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_EMAIL, 1, 1).setValue('contact@example.org');

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
    .setBackground('#1e293b')  // Dark slate
    .setFontColor('#f8fafc')   // Light text
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);

  // ═══ ROW 2: Column Headers - Purple ═══
  sheet.getRange(2, 1, 1, lastCol)
    .setBackground('#7C3AED')  // Primary purple
    .setFontColor('#ffffff')   // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  // ═══ ROWS 3+: Data Entry Area - Light Background ═══
  if (maxRows >= 3) {
    var dataRange = sheet.getRange(3, 1, maxRows - 2, lastCol);
    dataRange
      .setBackground('#f8fafc')  // Very light gray/white
      .setFontColor('#1e293b')   // Dark text
      .setFontWeight('normal')
      .setVerticalAlignment('middle');

    // Apply alternating row colors (zebra stripes) for readability
    for (var row = 3; row <= maxRows; row++) {
      var rowRange = sheet.getRange(row, 1, 1, lastCol);
      if (row % 2 === 0) {
        rowRange.setBackground('#f1f5f9');  // Slightly darker for even rows
      } else {
        rowRange.setBackground('#ffffff');  // White for odd rows
      }
    }
  }

  // Apply section-specific column colors to headers
  applySectionColors_(sheet, lastCol);

  ss.toast('Config sheet styling applied!', 'Theme Applied', 3);
  Logger.log('Config sheet styling applied to ' + lastCol + ' columns');
}

/**
 * Applies section-specific colors to column headers in Config sheet
 * Each section gets a distinct color for easy identification
 * @param {Sheet} sheet - The Config sheet
 * @param {number} lastCol - Last column number
 * @private
 */
function applySectionColors_(sheet, lastCol) {
  // Section color definitions (15 sections, columns A-AZ)
  var SECTION_COLORS = {
    EMPLOYMENT: { bg: '#3b82f6', text: '#ffffff' },      // Blue - A-E
    SUPERVISION: { bg: '#8b5cf6', text: '#ffffff' },     // Violet - F-G
    STEWARD: { bg: '#06b6d4', text: '#ffffff' },         // Cyan - H-I
    GRIEVANCE: { bg: '#ef4444', text: '#ffffff' },       // Red - J-M
    LINKS: { bg: '#f97316', text: '#ffffff' },           // Orange - N-Q
    NOTIFICATIONS: { bg: '#eab308', text: '#1e293b' },   // Yellow - R-T
    ORGANIZATION: { bg: '#22c55e', text: '#ffffff' },    // Green - U-X
    INTEGRATION: { bg: '#14b8a6', text: '#ffffff' },     // Teal - Y-Z
    DEADLINES: { bg: '#ec4899', text: '#ffffff' },       // Pink - AA-AD
    MULTISELECT: { bg: '#a855f7', text: '#ffffff' },     // Purple - AE-AF
    CONTRACT: { bg: '#6366f1', text: '#ffffff' },        // Indigo - AG-AJ
    IDENTITY: { bg: '#0ea5e9', text: '#ffffff' },        // Sky - AK-AM
    EXTENDED: { bg: '#84cc16', text: '#1e293b' },        // Lime - AN-AR
    COMMAND: { bg: '#f43f5e', text: '#ffffff' },         // Rose - AS-AY
    MOBILE: { bg: '#10b981', text: '#ffffff' }           // Emerald - AZ (LAST COLUMN)
  };

  // Apply colors by column ranges (both row 1 section header and row 2 column header)
  // Total: 52 columns (A-AZ)
  var sections = [
    { start: 1, end: 5, color: SECTION_COLORS.EMPLOYMENT },      // A-E
    { start: 6, end: 7, color: SECTION_COLORS.SUPERVISION },     // F-G
    { start: 8, end: 9, color: SECTION_COLORS.STEWARD },         // H-I
    { start: 10, end: 13, color: SECTION_COLORS.GRIEVANCE },     // J-M
    { start: 14, end: 17, color: SECTION_COLORS.LINKS },         // N-Q
    { start: 18, end: 20, color: SECTION_COLORS.NOTIFICATIONS }, // R-T
    { start: 21, end: 24, color: SECTION_COLORS.ORGANIZATION },  // U-X
    { start: 25, end: 26, color: SECTION_COLORS.INTEGRATION },   // Y-Z
    { start: 27, end: 30, color: SECTION_COLORS.DEADLINES },     // AA-AD
    { start: 31, end: 32, color: SECTION_COLORS.MULTISELECT },   // AE-AF
    { start: 33, end: 36, color: SECTION_COLORS.CONTRACT },      // AG-AJ
    { start: 37, end: 39, color: SECTION_COLORS.IDENTITY },      // AK-AM
    { start: 40, end: 44, color: SECTION_COLORS.EXTENDED },      // AN-AR
    { start: 45, end: 51, color: SECTION_COLORS.COMMAND },       // AS-AY (Strategic Command Center)
    { start: 52, end: 52, color: SECTION_COLORS.MOBILE }         // AZ (Mobile Dashboard - LAST COLUMN)
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

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Define guide colors
  var headerBg = COLORS.STATUS_BLUE;       // Blue header
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
    sheet.getRange(row, 1).setValue(columnData[j][0]).setFontColor(COLORS.STATUS_BLUE).setFontWeight('bold').setHorizontalAlignment('center');
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

  // Set tab color
  sheet.setTabColor(COLORS.STATUS_BLUE);

  return sheet;
}

/**
 * Create or recreate the Member Directory sheet
 */
function createMemberDirectory(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEMBER_DIR);
  // getOrCreateSheet now preserves data - only set headers on empty sheets
  if (sheet.getLastRow() <= 1) {
    var headers = getMemberHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold');
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

    // Group 3: Mailing Address / PII (AK-AM, columns 37-39) - Hidden by default, PII
    sheet.getRange(1, MEMBER_COLS.STREET_ADDRESS, sheet.getMaxRows(), 3).shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    Logger.log('Member Directory column group setup skipped: ' + e.toString());
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

  // ═══════════════════════════════════════════════════════════════════════════
  // CONDITIONAL FORMATTING: Highlight members with open grievances
  // ═══════════════════════════════════════════════════════════════════════════
  var _lastRow = Math.max(sheet.getLastRow(), 2);
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

  var rules = [redRule, emptyEmailRule, emptyPhoneRule, deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule];
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

  // Set tab color
  sheet.setTabColor(COLORS.UNION_GREEN);
}

/**
 * Create or recreate the Grievance Log sheet
 * NOTE: Calculated columns (First Name, Last Name, Email, Deadlines, Days Open, etc.)
 * are managed by the hidden _Grievance_Formulas sheet for self-healing capability.
 * Users can't accidentally erase formulas because they're in the hidden sheet.
 */
function createGrievanceLog(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.GRIEVANCE_LOG);
  // getOrCreateSheet now preserves data - only set headers on empty sheets
  if (sheet.getLastRow() <= 1) {
    var headers = getGrievanceHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold');
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
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    // Group Coordinator columns AC-AF (Message Alert, Coordinator Message, Acknowledged By, Acknowledged Date)
    sheet.getRange(1, GRIEVANCE_COLS.MESSAGE_ALERT, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
    // Collapse all groups by default (including coordinator columns AC-AF)
    sheet.collapseAllColumnGroups();
    // Hide Drive Folder ID column (AG) - internal use only
    sheet.hideColumns(GRIEVANCE_COLS.DRIVE_FOLDER_ID, 1);
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
  var rules = [
    deadlineOverdueRule, deadline1to3Rule, deadline4to7Rule, deadlineOnTrackRule,
    completedRule, step3ProgressRule, step2ProgressRule, step1ProgressRule, notReachedRule
  ];
  sheet.setConditionalFormatRules(rules);

  // Set tab color
  sheet.setTabColor(COLORS.SOLIDARITY_RED);
}

/**
 * @deprecated v4.3.2 - Dashboard sheet is deprecated. Use modal dashboards instead.
 * Access dashboards via: Union Hub > Dashboards menu
 * - showInteractiveDashboardTab() - Interactive Dashboard
 * - rebuildExecutiveDashboard() - Executive Command (PII)
 * - showPublicMemberDashboard() - Member Dashboard (No PII)
 * - showStewardPerformanceModal() - Steward Performance
 *
 * This function is kept for backwards compatibility but is no longer
 * called during setup. Use removeDeprecatedTabs() to remove the sheet.
 *
 * Create or recreate the unified Dashboard sheet (Executive Dashboard theme)
 * Combines member metrics, grievance metrics, and key performance indicators
 * Uses Executive Dashboard's green QUICK STATS theme
 */
function createDashboard(ss) {
  // Show deprecation warning
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Dashboard sheet is deprecated. Use Union Hub > Dashboards menu for modal dashboards.',
    '⚠️ Deprecated',
    5
  );

  var sheet = getOrCreateSheet(ss, SHEETS.DASHBOARD);
  sheet.clear();

  // Title - Executive Dashboard style with dark header card
  sheet.getRange('A1').setValue('📊 Executive Dashboard')
    .setFontSize(24)
    .setFontWeight('bold')
    .setFontColor(COLORS.WHITE)
    .setBackground(COLORS.CARD_DARK_BG);
  sheet.getRange('A1:F1').merge();

  // Subtitle with last refresh - dark theme continuation
  sheet.getRange('A2').setValue('Real-time metrics • Auto-refreshed from live data')
    .setFontSize(11)
    .setFontStyle('italic')
    .setFontColor(COLORS.CARD_DARK_TEXT)
    .setBackground(COLORS.CARD_DARK_BG);
  sheet.getRange('A2:F2').merge();

  // Spacer row
  sheet.getRange('A3:F3').setBackground(COLORS.WHITE);
  sheet.setRowHeight(3, 8);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 1: QUICK STATS - High-Contrast Card Layout
  // ═══════════════════════════════════════════════════════════════════════════
  // Dark header for card effect
  sheet.getRange('A4').setValue('📈 QUICK STATS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A4:F4').merge();

  // Quick stats row - labels with subtle background
  var quickStatsLabels = [
    ['Total Members', 'Active Stewards', 'Active Grievances', 'Win Rate', 'Overdue Cases', 'Due This Week']
  ];
  sheet.getRange('A5:F5').setValues(quickStatsLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#E8F5E9')  // Light green tint
    .setHorizontalAlignment('center')
    .setFontColor(COLORS.TEXT_DARK);

  // Quick stats values - large vibrant numbers
  sheet.getRange('A6:F6')
    .setFontSize(24)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground(COLORS.WHITE);

  // Color-code specific stat cells
  sheet.getRange('A6').setFontColor(COLORS.CHART_BLUE);      // Total Members - blue
  sheet.getRange('B6').setFontColor(COLORS.CHART_GREEN);     // Stewards - green
  sheet.getRange('C6').setFontColor(COLORS.CHART_PURPLE);    // Active - purple
  sheet.getRange('D6').setFontColor(COLORS.UNION_GREEN);     // Win Rate - green
  sheet.getRange('E6').setFontColor(COLORS.SOLIDARITY_RED);  // Overdue - red
  sheet.getRange('F6').setFontColor(COLORS.ACCENT_ORANGE);   // Due This Week - orange

  // Card bottom border
  sheet.getRange('A7:F7').setBackground(COLORS.SECTION_STATS);
  sheet.setRowHeight(7, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 2: MEMBER METRICS - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A9').setValue('👥 MEMBER METRICS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A9:D9').merge();

  var memberMetricLabels = [['Total Members', 'Active Stewards', 'Avg Open Rate', 'YTD Vol. Hours']];
  sheet.getRange('A10:D10').setValues(memberMetricLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#E3F2FD')  // Light blue tint
    .setHorizontalAlignment('center');

  // Member metric values with color coding
  sheet.getRange('A11:D11')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground(COLORS.WHITE);
  sheet.getRange('A11').setFontColor(COLORS.CHART_BLUE);
  sheet.getRange('B11').setFontColor(COLORS.CHART_GREEN);
  sheet.getRange('C11').setFontColor(COLORS.CHART_PURPLE);
  sheet.getRange('D11').setFontColor(COLORS.CHART_YELLOW);

  // Card bottom border
  sheet.getRange('A12:D12').setBackground(COLORS.SECTION_MEMBERS);
  sheet.setRowHeight(12, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 3: GRIEVANCE METRICS - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A14').setValue('📋 GRIEVANCE METRICS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A14:F14').merge();

  var grievanceLabels = [['Open', 'Pending Info', 'Settled', 'Won', 'Denied', 'Withdrawn']];
  sheet.getRange('A15:F15').setValues(grievanceLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#FFF3E0')  // Light orange tint
    .setHorizontalAlignment('center');

  // Grievance metric values with status colors
  sheet.getRange('A16:F16')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground(COLORS.WHITE);
  sheet.getRange('A16').setFontColor(COLORS.SOLIDARITY_RED);   // Open - red
  sheet.getRange('B16').setFontColor(COLORS.ACCENT_ORANGE);    // Pending - orange
  sheet.getRange('C16').setFontColor(COLORS.STATUS_BLUE);      // Settled - blue
  sheet.getRange('D16').setFontColor(COLORS.UNION_GREEN);      // Won - green
  sheet.getRange('E16').setFontColor('#9CA3AF');               // Denied - gray
  sheet.getRange('F16').setFontColor('#6B7280');               // Withdrawn - gray

  // Card bottom border
  sheet.getRange('A17:F17').setBackground(COLORS.SECTION_GRIEVANCE);
  sheet.setRowHeight(17, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 4: TIMELINE METRICS - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A19').setValue('⏱️ TIMELINE & PERFORMANCE')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A19:D19').merge();

  var timelineLabels = [['Avg Days Open', 'Filed This Month', 'Closed This Month', 'Avg Resolution Days']];
  sheet.getRange('A20:D20').setValues(timelineLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#F3E5F5')  // Light purple tint
    .setHorizontalAlignment('center');

  // Timeline metric values
  sheet.getRange('A21:D21')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground(COLORS.WHITE);
  sheet.getRange('A21').setFontColor(COLORS.CHART_PURPLE);
  sheet.getRange('B21').setFontColor(COLORS.CHART_BLUE);
  sheet.getRange('C21').setFontColor(COLORS.CHART_GREEN);
  sheet.getRange('D21').setFontColor(COLORS.CHART_INDIGO);

  // Card bottom border
  sheet.getRange('A22:D22').setBackground(COLORS.SECTION_TIMELINE);
  sheet.setRowHeight(22, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 5: TYPE ANALYSIS (by Issue Category) - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A24').setValue('📊 ISSUE BREAKDOWN (by Category)')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A24:F24').merge();

  var typeLabels = [['Issue Category', 'Total Cases', 'Open', 'Resolved', 'Win Rate', 'Avg Days']];
  sheet.getRange('A25:F25').setValues(typeLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#EDE9FE')  // Light indigo tint
    .setHorizontalAlignment('center');

  // Type analysis values - alternate row coloring
  for (var r = 26; r <= 30; r++) {
    sheet.getRange('A' + r + ':F' + r).setHorizontalAlignment('center');
    if (r % 2 === 0) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_LIGHT);
    }
  }

  // Card bottom border
  sheet.getRange('A31:F31').setBackground(COLORS.SECTION_ANALYSIS);
  sheet.setRowHeight(31, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 6: LOCATION BREAKDOWN - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A33').setValue('📍 LOCATION BREAKDOWN')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A33:F33').merge();

  var locationLabels = [['Location', 'Members', 'Grievances', 'Open Cases', 'Win Rate', 'Satisfaction']];
  sheet.getRange('A34:F34').setValues(locationLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#E0F7FA')  // Light cyan tint
    .setHorizontalAlignment('center');

  // Location values - alternate row coloring
  for (var r = 35; r <= 39; r++) {
    sheet.getRange('A' + r + ':F' + r).setHorizontalAlignment('center');
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_BLUE);
    }
  }

  // Card bottom border
  sheet.getRange('A40:F40').setBackground(COLORS.SECTION_LOCATION);
  sheet.setRowHeight(40, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 7: MONTH-OVER-MONTH TRENDS with Sparklines - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A42').setValue('📈 MONTH-OVER-MONTH TRENDS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A42:G42').merge();

  var trendLabels = [['Metric', 'This Month', 'Last Month', 'Change', '% Change', '6-Mo Trend', 'Sparkline']];
  sheet.getRange('A43:G43').setValues(trendLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#FFEBEE')  // Light red tint
    .setHorizontalAlignment('center');

  // Trend values with sparkline column
  var trendMetrics = [['Active Grievances', '', '', '', '', '', ''], ['Total Members', '', '', '', '', '', ''], ['Cases Filed', '', '', '', '', '', '']];
  sheet.getRange('A44:G46').setValues(trendMetrics)
    .setHorizontalAlignment('center');

  // Sparkline formulas will be populated by syncDashboardValues
  // Format: =SPARKLINE({data}, {"charttype","line";"color","#DC2626"})

  // Color code the change column based on positive/negative
  sheet.getRange('D44:E46').setFontWeight('bold');

  // Card bottom border
  sheet.getRange('A47:G47').setBackground(COLORS.SECTION_TRENDS);
  sheet.setRowHeight(47, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 8: STATUS LEGEND - Compact Visual Guide
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A49').setValue('🔖 STATUS LEGEND')
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);
  sheet.getRange('A49:G49').merge();

  // Deadline status with color-coded backgrounds
  var legendDeadline = [
    ['🟢 >7 days', '🟡 4-7 days', '🟠 1-3 days', '🔴 Overdue', '✅ Won', '❌ Denied', '⏸️ Pending']
  ];
  sheet.getRange('A50:G50').setValues(legendDeadline)
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontWeight('bold');
  // Color-code each legend item background
  sheet.getRange('A50').setBackground(COLORS.GRADIENT_LOW);
  sheet.getRange('B50').setBackground(COLORS.GRADIENT_MID_LOW);
  sheet.getRange('C50').setBackground(COLORS.GRADIENT_MID);
  sheet.getRange('D50').setBackground(COLORS.GRADIENT_HIGH);
  sheet.getRange('E50').setBackground('#D1FAE5');
  sheet.getRange('F50').setBackground('#FEE2E2');
  sheet.getRange('G50').setBackground('#FEF3C7');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 9: STEWARD PERFORMANCE SUMMARY - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A52').setValue('🛡️ STEWARD PERFORMANCE SUMMARY')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A52:F52').merge();

  var stewardLabels = [['Total Stewards', 'Active (w/ Cases)', 'Avg Cases/Steward', 'Total Vol Hours', 'Contacts This Month', 'Win Rate']];
  sheet.getRange('A53:F53').setValues(stewardLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#F3E5F5')  // Light purple tint
    .setHorizontalAlignment('center');

  // Steward summary values with vibrant colors
  sheet.getRange('A54:F54')
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground(COLORS.WHITE);
  sheet.getRange('A54').setFontColor(COLORS.CHART_PURPLE);
  sheet.getRange('B54').setFontColor(COLORS.CHART_BLUE);
  sheet.getRange('C54').setFontColor(COLORS.CHART_INDIGO);
  sheet.getRange('D54').setFontColor(COLORS.CHART_YELLOW);
  sheet.getRange('E54').setFontColor(COLORS.CHART_CYAN);
  sheet.getRange('F54').setFontColor(COLORS.UNION_GREEN);

  // Card bottom border
  sheet.getRange('A55:F55').setBackground(COLORS.SECTION_PERFORMANCE);
  sheet.setRowHeight(55, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 10: TOP 30 BUSIEST STEWARDS - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A57').setValue('📊 TOP 30 BUSIEST STEWARDS (Active Caseload)')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A57:F57').merge();

  var busiestLabels = [['Rank', 'Steward Name', 'Active Cases', 'Open', 'Pending', 'Total Ever']];
  sheet.getRange('A58:F58').setValues(busiestLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#FFEBEE')  // Light red tint
    .setHorizontalAlignment('center');

  // Busiest stewards values - alternate row coloring with gradient effect
  for (var r = 59; r <= 88; r++) {
    sheet.getRange('A' + r + ':F' + r).setHorizontalAlignment('center');
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_RED);
    }
  }

  // Card bottom border
  sheet.getRange('A89:F89').setBackground('#B91C1C');
  sheet.setRowHeight(89, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 11: TOP 10 PERFORMERS BY SCORE - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A91').setValue('🏆 TOP 10 PERFORMERS BY SCORE')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A91:F91').merge();

  var topPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A92:F92').setValues(topPerfLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#D1FAE5')  // Light green tint
    .setHorizontalAlignment('center');

  // Top performers values - gradient green rows
  for (var r = 93; r <= 102; r++) {
    sheet.getRange('A' + r + ':F' + r).setHorizontalAlignment('center');
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_GREEN);
    }
  }

  // Add medal icons for top 3
  sheet.getRange('A93').setNumberFormat('"🥇 "0');
  sheet.getRange('A94').setNumberFormat('"🥈 "0');
  sheet.getRange('A95').setNumberFormat('"🥉 "0');

  // Card bottom border
  sheet.getRange('A103:F103').setBackground(COLORS.UNION_GREEN);
  sheet.setRowHeight(103, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 12: STEWARDS NEEDING SUPPORT - Card Style
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A105').setValue('⚠️ STEWARDS NEEDING SUPPORT (Lowest Scores)')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A105:F105').merge();

  var lowPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A106:F106').setValues(lowPerfLabels)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#FEE2E2')  // Light red tint
    .setHorizontalAlignment('center');

  // Stewards needing support - gradient red rows
  for (var r = 107; r <= 116; r++) {
    sheet.getRange('A' + r + ':F' + r).setHorizontalAlignment('center');
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_RED);
    }
  }

  // Card bottom border
  sheet.getRange('A117:F117').setBackground(COLORS.SOLIDARITY_RED);
  sheet.setRowHeight(117, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // FORMATTING AND CLEANUP - Enhanced Visual Styling
  // ═══════════════════════════════════════════════════════════════════════════

  // Auto-resize and format
  sheet.autoResizeColumns(1, 7);
  sheet.setFrozenRows(2);

  // Set minimum column widths for better card appearance
  for (var i = 1; i <= 7; i++) {
    if (sheet.getColumnWidth(i) < 110) {
      sheet.setColumnWidth(i, 110);
    }
  }

  // Hide column H onwards (we now use G for sparklines)
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 7) {
      sheet.deleteColumns(8, maxCols - 7);
    }
  } catch (e) {
    Logger.log('Could not delete extra columns: ' + e.toString());
  }

  // NUMBER FORMATTING: Use comma separators for all numeric values (1,000)
  var numberFormat = '#,##0';
  var decimalFormat = '#,##0.0';
  var _percentFormat = '0.0%';

  // Quick Stats row (row 6) - updated row numbers
  sheet.getRange('A6:C6').setNumberFormat(numberFormat);
  sheet.getRange('E6:F6').setNumberFormat(numberFormat);

  // Member Metrics row (row 11)
  sheet.getRange('A11:B11').setNumberFormat(numberFormat);
  sheet.getRange('D11').setNumberFormat(numberFormat);

  // Grievance Metrics row (row 16)
  sheet.getRange('A16:F16').setNumberFormat(numberFormat);

  // Timeline row (row 21)
  sheet.getRange('A21:D21').setNumberFormat(numberFormat);

  // Type Analysis rows (rows 26-30)
  sheet.getRange('B26:D30').setNumberFormat(numberFormat);

  // Location Breakdown rows (rows 35-39)
  sheet.getRange('B35:D39').setNumberFormat(numberFormat);

  // Trends rows (rows 44-46)
  sheet.getRange('B44:D46').setNumberFormat(numberFormat);

  // Steward Performance row (row 54)
  sheet.getRange('A54:E54').setNumberFormat(numberFormat);

  // Top 30 Busiest Stewards (rows 59-88)
  sheet.getRange('C59:F88').setNumberFormat(numberFormat);

  // Top 10 Performers (rows 93-102) - Score and Win Rate have decimals
  sheet.getRange('C93:D102').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E93:F102').setNumberFormat(numberFormat);   // Avg Days, Overdue

  // Stewards Needing Support (rows 107-116)
  sheet.getRange('C107:D116').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E107:F116').setNumberFormat(numberFormat);   // Avg Days, Overdue

  // ═══════════════════════════════════════════════════════════════════════════
  // ADD NATIVE CHARTS - Gauge for Win Rate, Bar for Status Distribution
  // ═══════════════════════════════════════════════════════════════════════════
  createDashboardCharts_(sheet);

  // Populate all values using JavaScript-computed metrics (no formulas in visible sheet)
  syncDashboardValues();

  // Set tab color
  sheet.setTabColor(COLORS.SECTION_TIMELINE);
}

/**
 * Creates native Google Sheets charts for the dashboard
 * Adds a chart options section where users can select which charts to display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createDashboardCharts_(sheet) {
  // Remove existing charts first
  var existingCharts = sheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    sheet.removeChart(existingCharts[i]);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CHART OPTIONS SECTION - User selects which charts to display
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A119').setValue('📊 CHART OPTIONS - Select charts to display')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A119:G119').merge();

  // Instructions
  sheet.getRange('A120').setValue('Enter chart number in cell G120 and run "Generate Selected Chart" from Dashboard menu')
    .setFontStyle('italic')
    .setFontSize(10)
    .setFontColor('#6B7280');
  sheet.getRange('A120:F120').merge();
  sheet.getRange('G120').setValue(1)
    .setBackground('#FEF3C7')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Chart options table header
  var chartHeader = [['#', 'Chart Type', 'Description', 'Data Source', 'Best For', 'Preview']];
  sheet.getRange('A121:F121').setValues(chartHeader)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground('#E0E7FF')
    .setHorizontalAlignment('center');

  // Chart options data
  var chartOptions = [
    ['1', '🎯 Gauge Chart', 'Win Rate as dial/speedometer', 'Win Rate %', 'KPI at-a-glance', '◐'],
    ['2', '📊 Bar Chart - Status', 'Horizontal bars by status', 'Open/Pending/Closed counts', 'Status distribution', '▰▰▰'],
    ['3', '🥧 Pie Chart - Issues', 'Pie slices by issue type', 'Issue Category breakdown', 'Category proportions', '◕'],
    ['4', '📈 Line Chart - Trends', 'Monthly trend lines', '6-month grievance history', 'Trend analysis', '📉'],
    ['5', '📊 Column Chart', 'Vertical bars by location', 'Location breakdown', 'Comparing locations', '▮▮▮'],
    ['6', '🎯 Scorecard', 'Big number with trend arrow', 'Any single metric', 'Executive summary', '↑42'],
    ['7', '🔥 Heatmap Table', 'Color-coded data grid', 'Steward performance matrix', 'Pattern spotting', '🟩🟨🟥'],
    ['8', '📊 Stacked Bar', 'Stacked horizontal bars', 'Status by steward', 'Composition analysis', '▰▰▱'],
    ['9', '🍩 Donut Chart', 'Pie with center hole', 'Resolution outcomes', 'Outcome breakdown', '◎'],
    ['10', '📈 Area Chart', 'Filled line chart', 'Cumulative trends', 'Volume over time', '▓▓░'],
    ['11', '📊 Combo Chart', 'Bars + line overlay', 'Volume + Rate comparison', 'Dual metrics', '▮📈'],
    ['12', '🔵 Scatter Plot', 'X-Y point distribution', 'Response time vs outcomes', 'Correlation analysis', '⋮⋮'],
    ['13', '📊 Histogram', 'Frequency distribution bars', 'Case duration ranges', 'Distribution shape', '▁▃▅▇▅▃'],
    ['14', '📋 Summary Table', 'Key metrics in tabular form', 'All KPI summaries', 'Quick reference', '☰'],
    ['15', '🎯 Steward Leaderboard', 'Ranked performance list', 'Steward metrics', 'Performance ranking', '🥇🥈🥉']
  ];

  sheet.getRange('A122:F136').setValues(chartOptions)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Alternate row coloring for chart options
  for (var r = 122; r <= 136; r++) {
    if (r % 2 === 0) {
      sheet.getRange('A' + r + ':F' + r).setBackground(COLORS.ROW_ALT_LIGHT);
    }
  }

  // Style the # column
  sheet.getRange('A122:A136')
    .setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY_PURPLE);

  // Card bottom border
  sheet.getRange('A137:G137').setBackground(COLORS.CHART_INDIGO);
  sheet.setRowHeight(132, 4);

  // ═══════════════════════════════════════════════════════════════════════════
  // CHART DISPLAY AREA - Where selected chart will appear
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A134').setValue('📈 CHART DISPLAY AREA')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(COLORS.CARD_DARK_BG)
    .setFontColor(COLORS.CARD_DARK_TEXT);
  sheet.getRange('A134:G134').merge();

  // Placeholder for chart
  sheet.getRange('A135').setValue('Select a chart option above and run "Generate Selected Chart" to display here')
    .setFontStyle('italic')
    .setFontColor('#9CA3AF')
    .setHorizontalAlignment('center');
  sheet.getRange('A135:G145').merge()
    .setBackground('#F9FAFB')
    .setVerticalAlignment('middle');

  // Border around chart area
  sheet.getRange('A135:G145').setBorder(true, true, true, true, false, false, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);
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
    'Auto-ID', 'Dropdown', 'Auto-lookup', 'MM/DD/YYYY', 'Dropdown', 'Number', 'Text', 'Text', 'Text'
  ];

  sheet.getRange(2, 1, 1, hints.length).setValues([hints])
    .setFontStyle('italic')
    .setFontSize(9)
    .setBackground('#F0F9FF')
    .setFontColor('#6B7280')
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
  var nameLookupFormula = '=IF(B3<>"", IFERROR(VLOOKUP(B3, \'Member Directory\'!A:C, 2, FALSE) & " " & VLOOKUP(B3, \'Member Directory\'!A:C, 3, FALSE), "Not Found"), "")';
  sheet.getRange('C3').setFormula(nameLookupFormula);
  sheet.getRange('C3').copyTo(sheet.getRange('C3:C1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Freeze header rows
  sheet.setFrozenRows(2);

  // Set tab color
  sheet.setTabColor('#8B5CF6');  // Purple for volunteer hours

  Logger.log('Volunteer Hours sheet created');
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
    'Auto-ID', 'MM/DD/YYYY', 'Dropdown', 'Text', 'Dropdown', 'Auto-lookup', 'Checkbox', 'Text'
  ];

  sheet.getRange(2, 1, 1, hints.length).setValues([hints])
    .setFontStyle('italic')
    .setFontSize(9)
    .setBackground('#F0F9FF')
    .setFontColor('#6B7280')
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

  // Add checkboxes for Attended column
  sheet.getRange('G3:G1000').insertCheckboxes();

  // Auto-ID formula for Entry ID (column A)
  var idFormula = '=IF(E3<>"", "MTG-" & TEXT(ROW()-2, "0000"), "")';
  sheet.getRange('A3').setFormula(idFormula);
  sheet.getRange('A3').copyTo(sheet.getRange('A3:A1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Member Name lookup formula (column F) - VLOOKUP from Member Directory
  var nameLookupFormula = '=IF(E3<>"", IFERROR(VLOOKUP(E3, \'Member Directory\'!A:C, 2, FALSE) & " " & VLOOKUP(E3, \'Member Directory\'!A:C, 3, FALSE), "Not Found"), "")';
  sheet.getRange('F3').setFormula(nameLookupFormula);
  sheet.getRange('F3').copyTo(sheet.getRange('F3:F1000'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Freeze header rows
  sheet.setFrozenRows(2);

  // Set tab color
  sheet.setTabColor('#10B981');  // Green for meeting attendance

  Logger.log('Meeting Attendance sheet created');
}

/**
 * Create the Meeting Check-In Log sheet
 * Stewards create meetings here; members check in via modal with email + PIN
 * @param {Spreadsheet} ss - The active spreadsheet
 */
function createMeetingCheckInLogSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEETING_CHECKIN_LOG);

  var headers = [
    'Meeting ID',       // A - Unique meeting identifier (steward sets)
    'Meeting Name',     // B - Meeting name/topic
    'Meeting Date',     // C - Date of meeting
    'Meeting Type',     // D - Virtual or In-Person
    'Member ID',        // E - Checked-in member
    'Member Name',      // F - First + Last name
    'Check-In Time',    // G - Timestamp when member checked in
    'Email',            // H - Member email used for check-in
    'Start Time',       // I - Meeting start time (HH:mm)
    'Duration (hrs)',   // J - Meeting duration in hours
    'Event Status',     // K - Scheduled / Active / Completed
    'Notify Stewards',  // L - Steward email(s) for attendance report
    'Calendar Event ID',// M - Google Calendar event ID
    'Notes Doc URL',    // N - Meeting Notes Google Doc URL
    'Agenda Doc URL',   // O - Meeting Agenda Google Doc URL
    'Agenda Stewards'   // P - Steward emails for early agenda sharing
  ];

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

  // Format date columns
  sheet.getRange(2, 3, 999, 1).setNumberFormat('MM/DD/YYYY');   // C - Meeting Date
  sheet.getRange(2, 7, 999, 1).setNumberFormat('MM/DD/YYYY HH:mm:ss'); // G - Check-In Time
  sheet.getRange(2, 9, 999, 1).setNumberFormat('HH:mm');        // I - Start Time

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor('#8B5CF6');  // Purple for check-in

  Logger.log('Meeting Check-In Log sheet created');
}

/**
 * Menu-callable wrapper to create the Meeting Check-In Log sheet
 */
function setupMeetingCheckInSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createMeetingCheckInLogSheet(ss);
  ss.toast('Meeting Check-In Log sheet created', '📝 Meeting Check-In', 5);
}
