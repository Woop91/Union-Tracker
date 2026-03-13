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

// getFormFieldIds_ removed — dead code cleanup v4.25.11

// saveFormFieldIds_ removed — dead code cleanup v4.25.11

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

// getExistingGrievanceIds_ removed — dead code cleanup v4.25.11

// createGrievanceFolderFromData_ removed — dead code cleanup v4.25.11

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

// testGrievanceFormSubmission removed — dead code cleanup v4.25.11

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

// simplifyTimelineView removed — dead code cleanup v4.25.11

// showFullTimelineView removed — dead code cleanup v4.25.11

// setupTimelineColumnGroups removed — dead code cleanup v4.25.11

// applyStepHighlighting removed — dead code cleanup v4.25.11

// freezeKeyColumns removed — dead code cleanup v4.25.11

// unfreezeAllColumns removed — dead code cleanup v4.25.11

