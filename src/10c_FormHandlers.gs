// ============================================================================
// MENU HANDLER FUNCTIONS
// ============================================================================

/**
 * Show the desktop unified search dialog
 * Optimized for larger screens with advanced filtering
 */
// Note: showDesktopSearch() defined in modular file - see respective module

/**
 * Get locations for desktop search filter dropdown
 * @returns {Array} Array of unique locations
 */

/**
 * Get search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array} Array of search results
 */

/**
 * Navigate to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
 */

/**
 * Grievance Form Configuration
 * Maps form entry IDs to Member Directory fields for pre-filling
 */
/**
 * Grievance Form Configuration
 * Form URL reads from Config sheet (column P), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only — update via Script Properties or Config sheet.
 */
var GRIEVANCE_FORM_CONFIG = {
  // Default form URL — overridden by Config sheet column P at runtime
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLSedX8nf_xXeLe2sCL9MpjkEEmSuSPbjn3fNxMaMNaPlD0H5lA/viewform',

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
 * Form URL reads from Config sheet (column Q), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only.
 */
var CONTACT_FORM_CONFIG = {
  // Default form URL — overridden by Config sheet column Q at runtime
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
      for (var key in parsed) {
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
 * Get form URL from Config sheet, falling back to hardcoded default
 * This allows admins to update form links without touching code
 * @param {string} formType - 'grievance', 'contact', or 'satisfaction'
 * @returns {string} The form URL
 */

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

/**
 * Parse a date string from form submission
 * @private
 */

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

    // Get coordinator emails from Config (column O = GRIEVANCE_COORDINATORS)
    var coordData = configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_COORDINATORS,
                                          configSheet.getLastRow() - 1, 1).getValues();

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
// PERSONAL CONTACT INFO FORM HANDLER
// ============================================================================

/**
 * Show the Personal Contact Info form link
 * Members fill out the blank form and data is written to Member Directory on submit
 */

/**
 * Handle contact form submission
 * Writes member data to Member Directory (updates existing or creates new)
 *
 * @param {Object} e - Form submission event object
 */

/**
 * Get multiple values from form response (for checkbox questions)
 * Returns comma-separated string
 * @private
 */

/**
 * Set up the contact form submission trigger
 * Run this once to enable automatic processing of form submissions
 */

// ============================================================================
// MEMBER SATISFACTION SURVEY FORM HANDLER
// ============================================================================

/**
 * Member Satisfaction Survey Form Configuration
 */
/**
 * Satisfaction Survey Form Configuration
 * Form URL reads from Config sheet (column AR), entry IDs read from Script Properties.
 * Hardcoded values serve as defaults only.
 */
var SATISFACTION_FORM_CONFIG = {
  // Default form URLs — overridden by Config sheet column AR at runtime
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

/**
 * Save form URLs to the Config tab for easy reference and updating
 * Writes Grievance Form, Contact Form, and Satisfaction Survey URLs to Config columns P, Q, AR
 */

/**
 * Silent version - used during CREATE_509_DASHBOARD setup
 * @param {Spreadsheet} ss - The spreadsheet object
 * @private
 */

/**
 * Handle satisfaction survey form submission
 * Writes survey responses to the Member Satisfaction sheet
 *
 * @param {Object} e - Form submission event object
 */

/**
 * Set up the satisfaction survey form submission trigger
 * Run this once to enable automatic processing of survey submissions
 */

// ============================================================================
// SURVEY ENHANCEMENTS - Auto-Email, Quarterly Tracking, Member Auth
// ============================================================================

/**
 * Send satisfaction survey emails to random members
 * Allows stewards to email a configurable number of random members
 */

/**
 * Execute sending random survey emails
 * @param {Object} opts - Options {count, subject, excludeDays}
 * @returns {string} Result message
 */

/**
 * Validate that an email belongs to a member in the directory
 * @param {string} email - Email to validate
 * @returns {Object|null} Member info if valid, null otherwise
 */

/**
 * Get the current quarter string (e.g., "2026-Q1")
 * @returns {string} Quarter string
 */

/**
 * Get quarter string from a date
 * @param {Date} date - Date to get quarter from
 * @returns {string} Quarter string
 */

// ============================================================================
// FLAGGED SUBMISSIONS REVIEW - Admin interface for pending survey responses
// ============================================================================

/**
 * Show the flagged submissions review interface
 * Displays count and email addresses of Pending Review submissions
 * Protects actual survey answers - only shows metadata
 */

/**
 * Get HTML for flagged submissions review interface
 * @returns {string} HTML content
 */

/**
 * Get data for flagged submissions review
 * @returns {Object} Pending submissions data (email, date, row number - NO survey answers)
 */

/**
 * Approve a flagged submission - mark as Verified
 * @param {number} rowNum - Row number (1-indexed)
 */

/**
 * Reject a flagged submission - mark as Rejected
 * @param {number} rowNum - Row number (1-indexed)
 */

// ============================================================================
// PUBLIC MEMBER DASHBOARD - Stats without PII
// ============================================================================

// REMOVED: showPublicMemberDashboard_Code_DEPRECATED - Use showPublicMemberDashboard() in 11_SecureMemberDashboard.gs instead

/**
 * Generates HTML for the secure member dashboard
 * @param {Object} stats - Grievance statistics
 * @param {Array} stewards - Array of steward objects
 * @param {Object} satisfaction - Satisfaction statistics
 * @param {Object} coverage - Steward coverage statistics
 * @returns {string} HTML content
 */

/**
 * Get public overview data (no PII)
 * @returns {Object} Overview statistics
 */

/**
 * Get public survey data (anonymized)
 * Filters to only include Verified='Yes' and optionally IS_LATEST='Yes' responses
 * @param {boolean} includeHistory - If true, include superseded responses; if false, only latest per member
 * @returns {Object} Survey statistics
 */

/**
 * Get public grievance data (no PII)
 * @returns {Object} Grievance statistics
 */

/**
 * Get public steward data (contact info only)
 * @returns {Object} Steward directory
 */

/**
 * Aggregates survey data into chart-ready, non-PII formats.
 * Returns only aggregate metrics without exposing individual survey responses.
 * @returns {Object} Aggregate satisfaction statistics
 */

/**
 * Gets steward coverage ratio for progress tracking
 * @returns {Object} Coverage statistics
 */

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

