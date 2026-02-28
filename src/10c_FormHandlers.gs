// ============================================================================
// FORM CONFIGURATIONS
// ============================================================================

/**
 * Grievance Form Configuration
 * Form URL reads from Config sheet (column O), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only — update via Script Properties or Config sheet.
 */
var GRIEVANCE_FORM_CONFIG = {
  // Form URL — set in Config sheet column O, read via getFormUrlFromConfig('grievance')
  FORM_URL: '',

  // Default form field entry IDs — overridden by Script Properties at runtime
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
/**
 * Contact Form Configuration
 * Form URL reads from Config sheet (column P), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only.
 */
var CONTACT_FORM_CONFIG = {
  // Form URL — set in Config sheet column P, read via getFormUrlFromConfig('contact')
  FORM_URL: '',

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
    INTEREST_LOCAL: 'entry.1902862430',        // Willing to join direct actions
    HIRE_DATE: 'entry.PLACEHOLDER_HIRE_DATE',  // Hire Date — update entry ID after form creation
    EMPLOYEE_ID: 'entry.PLACEHOLDER_EMP_ID',   // Employee ID — update entry ID after form creation
    STREET_ADDRESS: 'entry.PLACEHOLDER_STREET', // Mailing Address: Street
    CITY: 'entry.PLACEHOLDER_CITY',            // Mailing Address: City
    ZIP_CODE: 'entry.PLACEHOLDER_ZIP',          // Mailing Address: Zip Code
    STATE: 'entry.PLACEHOLDER_STATE'             // Mailing Address: State — update entry ID after form creation
  }
};

// ============================================================================
// FORM CONFIGURATION HELPERS
// ============================================================================

/**
 * Get form field IDs from Script Properties, falling back to hardcoded defaults.
 * Admins can update field IDs via Script Properties without touching code.
 *
 * Script Property key format: FORM_FIELDS_{formType} (e.g., FORM_FIELDS_GRIEVANCE)
 * Value format: JSON string of the FIELD_IDS object
 *
 * @param {string} formType - 'grievance', 'contact', or 'satisfaction'
 * @returns {Object} Field ID mapping
 */
function getFormFieldIds_(formType) {
  var defaults = {
    grievance: GRIEVANCE_FORM_CONFIG.FIELD_IDS,
    contact: CONTACT_FORM_CONFIG.FIELD_IDS,
    satisfaction: SATISFACTION_FORM_CONFIG.FIELD_IDS
  };

  try {
    var propKey = 'FORM_FIELDS_' + formType.toUpperCase();
    var stored = PropertiesService.getScriptProperties().getProperty(propKey);
    if (stored) {
      var parsed = JSON.parse(stored);
      // Merge stored over defaults so new fields still work
      var result = {};
      var defaultFields = defaults[formType] || {};
      for (var key in defaultFields) {
        result[key] = defaultFields[key];
      }
      for (key in parsed) {
        result[key] = parsed[key];
      }
      return result;
    }
  } catch (e) {
    console.log('Error reading form field IDs from Script Properties: ' + e.message);
  }

  return defaults[formType] || {};
}

/**
 * Save form field IDs to Script Properties for a given form type.
 * Call this after creating/recreating a Google Form to update the entry IDs.
 *
 * @param {string} formType - 'grievance', 'contact', or 'satisfaction'
 * @param {Object} fieldIds - Object mapping field names to entry.XXXXX IDs
 */
function saveFormFieldIds_(formType, fieldIds) {
  var propKey = 'FORM_FIELDS_' + formType.toUpperCase();
  PropertiesService.getScriptProperties().setProperty(propKey, JSON.stringify(fieldIds));
}

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

    if (email && email.toLowerCase() === currentUserEmail.toLowerCase() && isTruthyValue(isSteward)) {
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

// ============================================================================
// GRIEVANCE FORM SUBMISSION HANDLER
// ============================================================================

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
function createGrievanceFolderFromData_(grievanceId, memberId, firstName, lastName, issueCategory, dateFiled) {
  try {
    // Get or create root folder
    var rootFolder = getOrCreateDashboardFolder_();

    // Format date as YYYY-MM-DD (default to current date if not provided)
    var date = dateFiled ? new Date(dateFiled) : new Date();
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Build folder name: LastName, FirstName - YYYY-MM-DD
    // Example: "Smith, John - 2026-01-15"
    var folderName;
    var sanitizedFirst = sanitizeFolderName_(firstName || '');
    var sanitizedLast = sanitizeFolderName_(lastName || '');

    if (sanitizedFirst && sanitizedLast) {
      folderName = sanitizedLast + ', ' + sanitizedFirst + ' - ' + dateStr;
    } else {
      // Fallback if name not available
      folderName = grievanceId + ' - ' + dateStr;
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
 * Sanitize folder name by removing invalid characters
 * @param {string} name - Name to sanitize
 * @returns {string} Sanitized name
 * @private
 */
function sanitizeFolderName_(name) {
  if (!name) return '';
  // Remove characters invalid for Google Drive folder names
  return name.toString().trim()
    .replace(/[\/\\:*?"<>|]/g, '')
    .replace(/\s+/g, ' ')
    .substring(0, 50); // Limit length
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

    // Get coordinator emails from Config (column N = GRIEVANCE_COORDINATORS)
    // Config data starts at row 3 (row 1 = section headers, row 2 = column headers)
    var lastRow = configSheet.getLastRow();
    if (lastRow < 3) return; // No data rows
    var coordData = configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_COORDINATORS,
                                          lastRow - 2, 1).getValues();

    for (var i = 0; i < coordData.length; i++) {
      var email = coordData[i][0];
      if (email && email.toString().trim() !== '') {
        try {
          folder.addEditor(email.toString().trim());
        } catch (shareError) {
          // Mask email in log for privacy
          var maskedEmail = typeof maskEmail === 'function' ? maskEmail(email) : '[REDACTED]';
          Logger.log('Could not share with ' + maskedEmail + ': ' + shareError.message);
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
// MEMBER SATISFACTION SURVEY FORM HANDLER
// ============================================================================

/**
 * Satisfaction Survey Form Configuration
 * Form URL reads from Config sheet (column AP), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only.
 */
var SATISFACTION_FORM_CONFIG = {
  // Form URLs — set in Config sheet column AP, read via getFormUrlFromConfig('satisfaction')
  FORM_URL: '',
  EDIT_URL: '',

  // Form field entry IDs (from pre-filled URL)
  FIELD_IDS: {
    WORKSITE: 'entry.829990399',
    TOP_PRIORITIES: 'entry.1290096581',      // Multi-select checkboxes
    ONE_CHANGE: 'entry.1926319061',
    KEEP_DOING: 'entry.1554906279',
    ADDITIONAL_COMMENTS: 'entry.650574503'
  }
};

// ============================================================================
// REFRESH & SYNC FUNCTIONS
// ============================================================================

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

/**
 * @deprecated v4.3.2 - Dashboard sheet is deprecated. Launches Interactive Dashboard modal.
 */
function rebuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Dashboard sheets are now modal-based. Opening Interactive Dashboard...', '📊 Dashboard', 3);

  // Refresh hidden sheet formulas and sync data (still useful)
  refreshAllHiddenFormulas();

  // Launch the Interactive Dashboard modal instead of rebuilding sheet
  showInteractiveDashboardTab();
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

  // Ensure sheet has enough columns for column group operations
  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  ss.toast('Setting up column groups...', '👁️ View', 2);

  // Group Step I columns (J-K)
  var _step1Range = sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, 1, 2);
  sheet.getColumnGroup(GRIEVANCE_COLS.STEP1_DUE, 1);

  // Group Step II columns (L-O)
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  var _step2Group = sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, 1, 4);
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  // Group Step III columns (P-Q)
  var _step3Group = sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, 1, 2);

  // Create the groups
  try {
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    // Group Coordinator columns AC-AF (Message Alert, Coordinator Message, Acknowledged By, Acknowledged Date)
    sheet.getRange(1, GRIEVANCE_COLS.MESSAGE_ALERT, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);

    // Collapse all groups by default (including coordinator columns AC-AF)
    sheet.collapseAllColumnGroups();

    // Hide Drive Folder ID column (AG) - internal use only
    sheet.hideColumns(GRIEVANCE_COLS.DRIVE_FOLDER_ID, 1);

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

  // M-47: Clear existing step highlighting rules before re-applying to prevent
  // duplicate rules from accumulating on repeated calls. Filter out rules that
  // target the step/deadline/status columns we're about to re-create.
  var existingRules = sheet.getConditionalFormatRules();
  var stepHighlightCols = [
    GRIEVANCE_COLS.STEP1_DUE, GRIEVANCE_COLS.STEP2_APPEAL_DUE,
    GRIEVANCE_COLS.STEP3_APPEAL_DUE, GRIEVANCE_COLS.DAYS_TO_DEADLINE,
    GRIEVANCE_COLS.NEXT_ACTION_DUE, GRIEVANCE_COLS.STATUS
  ];
  var rules = existingRules.filter(function(rule) {
    var ranges = rule.getRanges();
    if (!ranges || ranges.length === 0) return true;
    var ruleCol = ranges[0].getColumn();
    return stepHighlightCols.indexOf(ruleCol) === -1;
  });

  // Dynamic column letters from constants
  var grColStep = getColumnLetter(GRIEVANCE_COLS.CURRENT_STEP);
  var grColDaysDeadline = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var grColNextDue = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);

  // Rule 1: Gray out Step I columns if current step is Informal
  var step1Range = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, lastRow - 1, 2);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + grColStep + '2="Informal"')
    .setFontColor('#9e9e9e')
    .setRanges([step1Range])
    .build();

  // Rule 2: Gray out Step II columns if current step is Informal or Step I
  var step2Range = sheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, lastRow - 1, 4);
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + grColStep + '2="Informal",$' + grColStep + '2="Step I")')
    .setFontColor('#9e9e9e')
    .setRanges([step2Range])
    .build();

  // Rule 3: Gray out Step III columns if not at Step III or beyond
  var step3Range = sheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, lastRow - 1, 2);
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + grColStep + '2="Informal",$' + grColStep + '2="Step I",$' + grColStep + '2="Step II")')
    .setFontColor('#9e9e9e')
    .setRanges([step3Range])
    .build();

  // -------------------------------------------------------------------------
  // DEADLINE STATUS RULES
  // Order matters: more specific rules first, then broader ones
  // -------------------------------------------------------------------------

  var daysDeadlineRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);
  var nextDueRange = sheet.getRange(2, GRIEVANCE_COLS.NEXT_ACTION_DUE, lastRow - 1, 1);

  // Rule 4: Red - Overdue (Days to Deadline shows "Overdue" or negative/0)
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + grColDaysDeadline + '2="Overdue",AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2<=0))')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 5: Orange - Due in 1-3 days
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>=1,$' + grColDaysDeadline + '2<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 6: Yellow - Due in 4-7 days
  var rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>=4,$' + grColDaysDeadline + '2<=7)')
    .setBackground('#fffde7')
    .setFontColor('#f57f17')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 7: Green - On Track (more than 7 days remaining)
  var rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + grColDaysDeadline + '2),$' + grColDaysDeadline + '2>7)')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 8: Red highlight for Next Action Due if overdue
  var rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + grColNextDue + '2<>"",$' + grColNextDue + '2<TODAY())')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([nextDueRange])
    .build();

  // Rule 9: Orange for Next Action Due within 3 days
  var rule9 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + grColNextDue + '2<>"",($' + grColNextDue + '2-TODAY())>=0,($' + grColNextDue + '2-TODAY())<=3)')
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

  // H14: Filter out existing step-highlighting rules targeting our columns before adding new ones.
  // This prevents rule accumulation when the function is called multiple times.
  stepHighlightCols = [
    GRIEVANCE_COLS.STEP1_DUE, GRIEVANCE_COLS.STEP2_APPEAL_DUE,
    GRIEVANCE_COLS.STEP3_APPEAL_DUE, GRIEVANCE_COLS.DAYS_TO_DEADLINE,
    GRIEVANCE_COLS.NEXT_ACTION_DUE, GRIEVANCE_COLS.STATUS
  ];
  var filtered = rules.filter(function(r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    var col = ranges[0].getColumn();
    return stepHighlightCols.indexOf(col) === -1;
  });
  filtered.push(rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8, rule9, rule10, rule11, rule12);
  sheet.setConditionalFormatRules(filtered);

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

  // M-20: Use GRIEVANCE_COLS.CURRENT_STEP as the freeze boundary instead of
  // hardcoding column 6. This freezes through the identity and status columns.
  var freezeUpToCol = GRIEVANCE_COLS.CURRENT_STEP || 6;
  sheet.setFrozenColumns(freezeUpToCol);
  // Freeze header row
  sheet.setFrozenRows(1);

  ss.toast('Frozen columns A-' + getColumnLetter(freezeUpToCol) + ' and header row. Scroll right to see timeline.', '❄️ Frozen', 3);
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

