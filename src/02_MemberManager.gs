/**
 * 509 Dashboard - Member Manager
 *
 * Handles all member directory operations:
 * - Member CRUD operations (Create, Read, Update, Delete)
 * - Form submissions for member contact updates
 * - Member lookup and search
 * - Member-to-grievance relationships
 * - Multi-select field handling
 *
 * @version 2.3.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// MEMBER CRUD OPERATIONS
// ============================================================================

/**
 * Shows new member dialog
 */
function showNewMemberDialog() {
  var html = HtmlService.createHtmlOutput(getNewMemberFormHtml())
    .setWidth(550)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Member');
}

/**
 * Get HTML for new member form
 * @returns {string} HTML content
 */
function getNewMemberFormHtml() {
  var departments = getDepartmentList();
  var jobTitles = getJobTitleList();
  var supervisors = getSupervisorList();
  var managers = getManagerList();
  var stewards = getStewardList();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    '.form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 15px; }' +
    '.form-group { margin-bottom: 10px; }' +
    '.form-label { display: block; font-weight: bold; margin-bottom: 5px; color: #374151; }' +
    '.form-input, .form-select { width: 100%; padding: 8px; border: 1px solid #D1D5DB; border-radius: 4px; box-sizing: border-box; }' +
    '.form-input:focus, .form-select:focus { border-color: #7C3AED; outline: none; }' +
    '.required::after { content: " *"; color: #DC2626; }' +
    '.btn { padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; }' +
    '.btn-primary { background: #7C3AED; color: white; }' +
    '.btn-secondary { background: #6B7280; color: white; margin-right: 10px; }' +
    '.btn-primary:hover { background: #6D28D9; }' +
    '.actions { margin-top: 20px; text-align: right; }' +
    '</style></head><body>' +
    '<form id="memberForm">' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label required">First Name</label>' +
    '    <input type="text" class="form-input" id="firstName" required></div>' +
    '  <div class="form-group"><label class="form-label required">Last Name</label>' +
    '    <input type="text" class="form-input" id="lastName" required></div>' +
    '</div>' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label">Employee ID</label>' +
    '    <input type="text" class="form-input" id="employeeId" placeholder="XX000000"></div>' +
    '  <div class="form-group"><label class="form-label">Job Title</label>' +
    '    <select class="form-select" id="jobTitle"><option value="">Select...</option>' +
    jobTitles.map(function(t) { return '<option value="' + t + '">' + t + '</option>'; }).join('') +
    '  </select></div>' +
    '</div>' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label">Department/Unit</label>' +
    '    <select class="form-select" id="unit"><option value="">Select...</option>' +
    departments.map(function(d) { return '<option value="' + d + '">' + d + '</option>'; }).join('') +
    '  </select></div>' +
    '  <div class="form-group"><label class="form-label">Work Location</label>' +
    '    <input type="text" class="form-input" id="workLocation"></div>' +
    '</div>' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label">Email</label>' +
    '    <input type="email" class="form-input" id="email"></div>' +
    '  <div class="form-group"><label class="form-label">Phone</label>' +
    '    <input type="tel" class="form-input" id="phone" placeholder="XXX-XXX-XXXX"></div>' +
    '</div>' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label">Supervisor</label>' +
    '    <select class="form-select" id="supervisor"><option value="">Select...</option>' +
    supervisors.map(function(s) { return '<option value="' + s + '">' + s + '</option>'; }).join('') +
    '  </select></div>' +
    '  <div class="form-group"><label class="form-label">Manager</label>' +
    '    <select class="form-select" id="manager"><option value="">Select...</option>' +
    managers.map(function(m) { return '<option value="' + m + '">' + m + '</option>'; }).join('') +
    '  </select></div>' +
    '</div>' +
    '<div class="form-row">' +
    '  <div class="form-group"><label class="form-label">Hire Date</label>' +
    '    <input type="date" class="form-input" id="hireDate"></div>' +
    '  <div class="form-group"><label class="form-label">Is Steward?</label>' +
    '    <select class="form-select" id="isSteward"><option value="No">No</option><option value="Yes">Yes</option></select></div>' +
    '</div>' +
    '</form>' +
    '<div class="actions">' +
    '  <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '  <button class="btn btn-primary" onclick="submitMember()">Add Member</button>' +
    '</div>' +
    '<script>' +
    'function submitMember() {' +
    '  var data = {' +
    '    firstName: document.getElementById("firstName").value,' +
    '    lastName: document.getElementById("lastName").value,' +
    '    employeeId: document.getElementById("employeeId").value,' +
    '    jobTitle: document.getElementById("jobTitle").value,' +
    '    unit: document.getElementById("unit").value,' +
    '    workLocation: document.getElementById("workLocation").value,' +
    '    email: document.getElementById("email").value,' +
    '    phone: document.getElementById("phone").value,' +
    '    supervisor: document.getElementById("supervisor").value,' +
    '    manager: document.getElementById("manager").value,' +
    '    hireDate: document.getElementById("hireDate").value,' +
    '    isSteward: document.getElementById("isSteward").value' +
    '  };' +
    '  if (!data.firstName || !data.lastName) { alert("First Name and Last Name are required."); return; }' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r.success) { alert("Member added: " + r.memberId); google.script.host.close(); }' +
    '    else { alert("Error: " + r.error); }' +
    '  }).addNewMember(data);' +
    '}' +
    '</script></body></html>';
}

/**
 * Add a new member to the Member Directory
 * @param {Object} memberData - Member information
 * @returns {Object} Result with success status and member ID
 */
function addNewMember(memberData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      return { success: false, error: 'Member Directory not found' };
    }

    // Generate member ID
    var existingIds = getExistingMemberIds_(sheet);
    var memberId = generateNameBasedId('M', memberData.firstName, memberData.lastName, existingIds);

    // Build row data based on MEMBER_COLS
    var rowData = new Array(Object.keys(MEMBER_COLS).length).fill('');

    rowData[MEMBER_COLS.MEMBER_ID - 1] = memberId;
    rowData[MEMBER_COLS.FIRST_NAME - 1] = memberData.firstName || '';
    rowData[MEMBER_COLS.LAST_NAME - 1] = memberData.lastName || '';
    rowData[MEMBER_COLS.EMPLOYEE_ID - 1] = memberData.employeeId || '';
    rowData[MEMBER_COLS.JOB_TITLE - 1] = memberData.jobTitle || '';
    rowData[MEMBER_COLS.UNIT - 1] = memberData.unit || '';
    rowData[MEMBER_COLS.WORK_LOCATION - 1] = memberData.workLocation || '';
    rowData[MEMBER_COLS.EMAIL - 1] = memberData.email || '';
    rowData[MEMBER_COLS.PHONE - 1] = memberData.phone || '';
    rowData[MEMBER_COLS.SUPERVISOR - 1] = memberData.supervisor || '';
    rowData[MEMBER_COLS.MANAGER - 1] = memberData.manager || '';
    rowData[MEMBER_COLS.IS_STEWARD - 1] = memberData.isSteward || 'No';
    rowData[MEMBER_COLS.MEMBER_STATUS - 1] = 'Active';

    // Parse hire date
    if (memberData.hireDate) {
      rowData[MEMBER_COLS.HIRE_DATE - 1] = new Date(memberData.hireDate);
    }

    // Append row
    sheet.appendRow(rowData);

    // Log the addition
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MEMBER_ADDED', 'Added member: ' + memberId);
    }

    return { success: true, memberId: memberId };

  } catch (error) {
    Logger.log('Error adding member: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Get existing member IDs for collision detection
 * @param {Sheet} sheet - Member Directory sheet
 * @returns {Object} Object with existing IDs as keys
 * @private
 */
function getExistingMemberIds_(sheet) {
  var ids = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return ids;

  var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      ids[data[i][0]] = true;
    }
  }

  return ids;
}

// ============================================================================
// MEMBER LOOKUP AND SEARCH
// ============================================================================

/**
 * Get member by ID
 * @param {string} memberId - Member ID to look up
 * @returns {Object|null} Member data object or null if not found
 */
function getMemberById(memberId) {
  if (!memberId) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      return mapMemberRow(data[i]);
    }
  }

  return null;
}

/**
 * Get member by row number
 * @param {number} row - Row number (1-indexed)
 * @returns {Object|null} Member data object or null
 */
function getMemberByRow(row) {
  if (row < 2) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return null;

  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return mapMemberRow(rowData);
}

/**
 * Get all members
 * @returns {Array<Object>} Array of member data objects
 */
function getMemberList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data.map(function(row) {
    return mapMemberRow(row);
  });
}

/**
 * Get all stewards
 * @returns {Array<Object>} Array of steward member objects
 */
function getStewardsList() {
  var members = getMemberList();
  return members.filter(function(m) {
    return m.isSteward === 'Yes';
  });
}

/**
 * Search members by various criteria
 * @param {Object} criteria - Search criteria
 * @returns {Array<Object>} Matching members
 */
function searchMembers(criteria) {
  var members = getMemberList();

  return members.filter(function(m) {
    if (criteria.name) {
      var fullName = (m.firstName + ' ' + m.lastName).toLowerCase();
      if (fullName.indexOf(criteria.name.toLowerCase()) === -1) return false;
    }
    if (criteria.unit && m.unit !== criteria.unit) return false;
    if (criteria.jobTitle && m.jobTitle !== criteria.jobTitle) return false;
    if (criteria.status && m.memberStatus !== criteria.status) return false;
    if (criteria.isSteward && m.isSteward !== criteria.isSteward) return false;
    return true;
  });
}

// ============================================================================
// CONFIG LIST HELPERS
// ============================================================================

/**
 * Get list of departments/units from Config
 * @returns {Array<string>} Department list
 */
function getDepartmentList() {
  return getConfigColumnValues(CONFIG_COLS.UNITS);
}

/**
 * Get list of job titles from Config
 * @returns {Array<string>} Job title list
 */
function getJobTitleList() {
  return getConfigColumnValues(CONFIG_COLS.JOB_TITLES);
}

/**
 * Get list of supervisors from Config
 * @returns {Array<string>} Supervisor list
 */
function getSupervisorList() {
  return getConfigColumnValues(CONFIG_COLS.SUPERVISORS);
}

/**
 * Get list of managers from Config
 * @returns {Array<string>} Manager list
 */
function getManagerList() {
  return getConfigColumnValues(CONFIG_COLS.MANAGERS);
}

/**
 * Get list of stewards from Config
 * @returns {Array<string>} Steward list
 */
function getStewardList() {
  return getConfigColumnValues(CONFIG_COLS.STEWARDS);
}

/**
 * Get list of work locations from Config
 * @returns {Array<string>} Location list
 */
function getLocationList() {
  return getConfigColumnValues(CONFIG_COLS.OFFICE_LOCATIONS);
}

/**
 * Get values from a Config column
 * @param {number} column - Config column number
 * @returns {Array<string>} Non-empty values from the column
 */
function getConfigColumnValues(column) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) return [];

  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return [];

  var data = configSheet.getRange(3, column, lastRow - 2, 1).getValues();

  return data
    .map(function(row) { return row[0]; })
    .filter(function(val) { return val !== '' && val !== null; });
}

// ============================================================================
// MEMBER-GRIEVANCE RELATIONSHIP
// ============================================================================

/**
 * Get all grievances for a specific member
 * @param {string} memberId - Member ID
 * @returns {Array<Object>} Array of grievance objects
 */
function getMemberGrievances(memberId) {
  if (!memberId) return [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var grievances = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
      grievances.push(mapGrievanceRow(data[i]));
    }
  }

  return grievances;
}

/**
 * Check if member has open grievances
 * @param {string} memberId - Member ID
 * @returns {boolean} True if member has open grievances
 */
function memberHasOpenGrievance(memberId) {
  var grievances = getMemberGrievances(memberId);
  var openStatuses = ['Open', 'Pending Info', 'In Arbitration', 'Appealed'];

  return grievances.some(function(g) {
    return openStatuses.indexOf(g.status) !== -1;
  });
}

/**
 * Start a grievance for a specific member (quick action)
 * @param {string} memberId - Member ID
 */
function startGrievanceForMember(memberId) {
  var member = getMemberById(memberId);

  if (!member) {
    SpreadsheetApp.getUi().alert('Error', 'Member not found: ' + memberId, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Build pre-filled form URL
  var formUrl = buildGrievanceFormUrlForMember_(member);

  // Show confirmation with link
  var html = HtmlService.createHtmlOutput(
    '<style>body { font-family: Arial; padding: 20px; }' +
    '.member-info { background: #F3F4F6; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    '.btn { display: inline-block; padding: 12px 24px; background: #7C3AED; color: white; ' +
    'text-decoration: none; border-radius: 5px; font-weight: bold; }</style>' +
    '<h2>Start Grievance for:</h2>' +
    '<div class="member-info">' +
    '<strong>' + member.firstName + ' ' + member.lastName + '</strong><br>' +
    'ID: ' + member.memberId + '<br>' +
    'Department: ' + (member.unit || 'N/A') + '<br>' +
    'Email: ' + (member.email || 'N/A') +
    '</div>' +
    '<p>Click below to open the pre-filled grievance form:</p>' +
    '<a href="' + formUrl + '" target="_blank" class="btn">Open Grievance Form</a>' +
    '<p style="margin-top: 20px; color: #6B7280; font-size: 12px;">' +
    'The form will open in a new tab with member info pre-filled.</p>'
  ).setWidth(450).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Start New Grievance');
}

/**
 * Build pre-filled grievance form URL for a member
 * @param {Object} member - Member data object
 * @returns {string} Pre-filled form URL
 * @private
 */
function buildGrievanceFormUrlForMember_(member) {
  var baseUrl = typeof getFormUrlFromConfig === 'function'
    ? getFormUrlFromConfig('grievance')
    : 'https://docs.google.com/forms/d/e/1FAIpQLSedX8nf_xXeLe2sCL9MpjkEEmSuSPbjn3fNxMaMNaPlD0H5lA/viewform';

  var params = [];

  // Add member info as URL parameters
  // Note: These entry IDs should match your grievance form
  if (member.memberId) params.push('entry.272049116=' + encodeURIComponent(member.memberId));
  if (member.firstName) params.push('entry.736822578=' + encodeURIComponent(member.firstName));
  if (member.lastName) params.push('entry.694440931=' + encodeURIComponent(member.lastName));
  if (member.jobTitle) params.push('entry.286226203=' + encodeURIComponent(member.jobTitle));
  if (member.unit) params.push('entry.2025752361=' + encodeURIComponent(member.unit));
  if (member.workLocation) params.push('entry.413952220=' + encodeURIComponent(member.workLocation));
  if (member.manager) params.push('entry.417314483=' + encodeURIComponent(member.manager));
  if (member.email) params.push('entry.710401757=' + encodeURIComponent(member.email));

  // Add current steward info
  var steward = typeof getCurrentStewardInfo_ === 'function'
    ? getCurrentStewardInfo_(SpreadsheetApp.getActiveSpreadsheet())
    : { firstName: '', lastName: '', email: '' };

  if (steward.firstName) params.push('entry.84740378=' + encodeURIComponent(steward.firstName));
  if (steward.lastName) params.push('entry.1254106933=' + encodeURIComponent(steward.lastName));
  if (steward.email) params.push('entry.732806953=' + encodeURIComponent(steward.email));

  // Add today's date as filing date
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  params.push('entry.361538394=' + encodeURIComponent(today));

  return baseUrl + '?usp=pp_url&' + params.join('&');
}

// ============================================================================
// CONTACT FORM SUBMISSION HANDLER
// ============================================================================

/**
 * Handle contact info form submission
 * Updates or creates member record from form submission
 * @param {Object} e - Form submission event
 */
function onContactFormSubmit(e) {
  try {
    var responses = e.namedValues;

    var firstName = getFormValue_(responses, 'First Name');
    var lastName = getFormValue_(responses, 'Last Name');

    if (!firstName || !lastName) {
      Logger.log('Contact form missing required fields');
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      Logger.log('Member Directory not found');
      return;
    }

    // Try to find existing member by name
    var data = sheet.getDataRange().getValues();
    var existingRow = -1;

    for (var i = 1; i < data.length; i++) {
      var rowFirst = (data[i][MEMBER_COLS.FIRST_NAME - 1] || '').toString().toLowerCase();
      var rowLast = (data[i][MEMBER_COLS.LAST_NAME - 1] || '').toString().toLowerCase();

      if (rowFirst === firstName.toLowerCase() && rowLast === lastName.toLowerCase()) {
        existingRow = i + 1;
        break;
      }
    }

    // Extract form data
    var memberData = {
      firstName: firstName,
      lastName: lastName,
      jobTitle: getFormValue_(responses, 'Job Title') || getFormValue_(responses, 'Your Job Title'),
      unit: getFormValue_(responses, 'Unit') || getFormValue_(responses, 'Department'),
      workLocation: getFormValue_(responses, 'Work Location') || getFormValue_(responses, 'Office Location'),
      email: getFormValue_(responses, 'Email') || getFormValue_(responses, 'Email Address'),
      phone: getFormValue_(responses, 'Phone') || getFormValue_(responses, 'Phone Number'),
      supervisor: getFormValue_(responses, 'Supervisor'),
      manager: getFormValue_(responses, 'Manager')
    };

    if (existingRow > 0) {
      // Update existing member
      updateMemberFromFormData_(sheet, existingRow, memberData);
      Logger.log('Updated existing member in row ' + existingRow);
    } else {
      // Add new member
      var result = addNewMember(memberData);
      Logger.log('Added new member: ' + (result.success ? result.memberId : result.error));
    }

  } catch (error) {
    Logger.log('Error in onContactFormSubmit: ' + error.message);
  }
}

/**
 * Update member row from form data
 * @param {Sheet} sheet - Member Directory sheet
 * @param {number} row - Row to update
 * @param {Object} data - Form data
 * @private
 */
function updateMemberFromFormData_(sheet, row, data) {
  // Only update non-empty fields
  if (data.jobTitle) sheet.getRange(row, MEMBER_COLS.JOB_TITLE).setValue(data.jobTitle);
  if (data.unit) sheet.getRange(row, MEMBER_COLS.UNIT).setValue(data.unit);
  if (data.workLocation) sheet.getRange(row, MEMBER_COLS.WORK_LOCATION).setValue(data.workLocation);
  if (data.email) sheet.getRange(row, MEMBER_COLS.EMAIL).setValue(data.email);
  if (data.phone) sheet.getRange(row, MEMBER_COLS.PHONE).setValue(data.phone);
  if (data.supervisor) sheet.getRange(row, MEMBER_COLS.SUPERVISOR).setValue(data.supervisor);
  if (data.manager) sheet.getRange(row, MEMBER_COLS.MANAGER).setValue(data.manager);

  // Update last contact date
  sheet.getRange(row, MEMBER_COLS.RECENT_CONTACT_DATE).setValue(new Date());
}

/**
 * Helper to get form value from named responses
 * @param {Object} responses - Named responses object
 * @param {string} fieldName - Field name to look up
 * @returns {string} Field value or empty string
 * @private
 */
function getFormValue_(responses, fieldName) {
  if (responses[fieldName] && responses[fieldName].length > 0) {
    return responses[fieldName][0].toString().trim();
  }
  return '';
}

// ============================================================================
// MULTI-SELECT FIELD HANDLING
// ============================================================================

/**
 * Show multi-select dialog for a member field
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} currentValue - Current cell value (comma-separated)
 */
function showMemberMultiSelectDialog(row, col, currentValue) {
  var config = getMultiSelectConfigForColumn(col);
  if (!config) {
    SpreadsheetApp.getUi().alert('Error', 'Multi-select not configured for this column.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var options = getConfigColumnValues(config.sourceCol);
  var selected = currentValue ? currentValue.split(',').map(function(s) { return s.trim(); }) : [];

  var html = buildMultiSelectHtml(config.title, options, selected, row, col);

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(450),
    config.title
  );
}

/**
 * Get multi-select configuration for a column
 * @param {number} col - Column number
 * @returns {Object|null} Configuration object or null
 */
function getMultiSelectConfigForColumn(col) {
  var configs = {
    [MEMBER_COLS.OFFICE_DAYS]: { title: 'Select Office Days', sourceCol: CONFIG_COLS.OFFICE_DAYS },
    [MEMBER_COLS.PREFERRED_COMM]: { title: 'Preferred Communication', sourceCol: CONFIG_COLS.PREFERRED_COMM },
    [MEMBER_COLS.BEST_TIME]: { title: 'Best Time to Contact', sourceCol: CONFIG_COLS.BEST_TIME }
  };

  return configs[col] || null;
}

/**
 * Build HTML for multi-select dialog
 * @param {string} title - Dialog title
 * @param {Array<string>} options - Available options
 * @param {Array<string>} selected - Currently selected values
 * @param {number} row - Target row
 * @param {number} col - Target column
 * @returns {string} HTML content
 */
function buildMultiSelectHtml(title, options, selected, row, col) {
  var checkboxes = options.map(function(opt) {
    var checked = selected.indexOf(opt) !== -1 ? 'checked' : '';
    return '<label style="display:block;padding:8px;cursor:pointer;">' +
      '<input type="checkbox" value="' + opt + '" ' + checked + ' style="margin-right:10px;">' +
      opt + '</label>';
  }).join('');

  return '<!DOCTYPE html><html><head>' +
    '<style>body{font-family:Arial;padding:20px;}.options{max-height:300px;overflow-y:auto;border:1px solid #ddd;border-radius:4px;}' +
    '.btn{padding:10px 20px;border:none;border-radius:5px;cursor:pointer;}.btn-primary{background:#7C3AED;color:white;}' +
    '.btn-secondary{background:#6B7280;color:white;margin-right:10px;}</style></head><body>' +
    '<h3 style="margin-bottom:15px;">' + title + '</h3>' +
    '<div class="options">' + checkboxes + '</div>' +
    '<div style="margin-top:20px;text-align:right;">' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="saveSelection()">Save</button></div>' +
    '<script>function saveSelection(){' +
    'var checked=Array.from(document.querySelectorAll("input:checked")).map(function(cb){return cb.value;});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.applyMemberMultiSelectValue(' + row + ',' + col + ',checked.join(", "));}' +
    '</script></body></html>';
}

/**
 * Apply multi-select value to member cell
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} value - Comma-separated value
 */
function applyMemberMultiSelectValue(row, col, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (sheet) {
    sheet.getRange(row, col).setValue(value);
  }
}

// ============================================================================
// MEMBER QUICK ACTIONS
// ============================================================================

/**
 * Show quick actions menu for a member
 * @param {string} memberId - Member ID
 */
function showMemberQuickActions(memberId) {
  var member = getMemberById(memberId);

  if (!member) {
    SpreadsheetApp.getUi().alert('Error', 'Member not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var hasOpenGrievance = memberHasOpenGrievance(memberId);
  var grievances = getMemberGrievances(memberId);

  var html = '<!DOCTYPE html><html><head>' +
    '<style>body{font-family:Arial;padding:20px;margin:0;}' +
    '.member-header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;margin:-20px -20px 20px -20px;border-radius:0 0 10px 10px;}' +
    '.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px;}' +
    '.stat{background:#F3F4F6;padding:10px;border-radius:8px;text-align:center;}' +
    '.stat-num{font-size:24px;font-weight:bold;color:#7C3AED;}' +
    '.action-btn{display:block;width:100%;padding:12px;margin-bottom:10px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;}' +
    '.action-btn:hover{opacity:0.9;}' +
    '.btn-grievance{background:#DC2626;color:white;}' +
    '.btn-email{background:#2563EB;color:white;}' +
    '.btn-view{background:#059669;color:white;}' +
    '</style></head><body>' +
    '<div class="member-header">' +
    '<h2 style="margin:0;">' + member.firstName + ' ' + member.lastName + '</h2>' +
    '<p style="margin:5px 0 0 0;opacity:0.9;">' + (member.jobTitle || 'No title') + ' • ' + member.memberId + '</p>' +
    '</div>' +
    '<div class="stats">' +
    '<div class="stat"><div class="stat-num">' + grievances.length + '</div>Total Grievances</div>' +
    '<div class="stat"><div class="stat-num">' + (hasOpenGrievance ? 'Yes' : 'No') + '</div>Open Grievance</div>' +
    '</div>' +
    '<button class="action-btn btn-grievance" onclick="startGrievance()">📋 Start New Grievance</button>' +
    '<button class="action-btn btn-email" onclick="sendEmail()">✉️ Send Email</button>' +
    '<button class="action-btn btn-view" onclick="viewHistory()">📊 View Grievance History</button>' +
    '<script>' +
    'function startGrievance(){google.script.run.startGrievanceForMember("' + memberId + '");google.script.host.close();}' +
    'function sendEmail(){if("' + member.email + '"){window.open("mailto:' + member.email + '");}else{alert("No email on file.");}}' +
    'function viewHistory(){google.script.run.showMemberGrievanceHistory("' + memberId + '");google.script.host.close();}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(350).setHeight(400),
    'Member Actions'
  );
}

/**
 * Show grievance history for a member
 * @param {string} memberId - Member ID
 */
function showMemberGrievanceHistory(memberId) {
  var member = getMemberById(memberId);
  var grievances = getMemberGrievances(memberId);

  if (grievances.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Grievances',
      member.firstName + ' ' + member.lastName + ' has no grievances on record.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var rows = grievances.map(function(g) {
    var statusColor = g.status === 'Won' ? '#059669' : (g.status === 'Open' ? '#DC2626' : '#6B7280');
    return '<tr>' +
      '<td style="padding:8px;">' + g.grievanceId + '</td>' +
      '<td style="padding:8px;">' + (g.issueCategory || 'N/A') + '</td>' +
      '<td style="padding:8px;"><span style="background:' + statusColor + ';color:white;padding:2px 8px;border-radius:4px;">' + g.status + '</span></td>' +
      '<td style="padding:8px;">' + (g.dateFiled ? new Date(g.dateFiled).toLocaleDateString() : 'N/A') + '</td>' +
      '</tr>';
  }).join('');

  var html = '<!DOCTYPE html><html><head>' +
    '<style>body{font-family:Arial;padding:20px;}table{width:100%;border-collapse:collapse;}' +
    'th{background:#7C3AED;color:white;padding:10px;text-align:left;}' +
    'td{border-bottom:1px solid #E5E7EB;}</style></head><body>' +
    '<h2>' + member.firstName + ' ' + member.lastName + ' - Grievance History</h2>' +
    '<table><tr><th>ID</th><th>Issue</th><th>Status</th><th>Filed</th></tr>' + rows + '</table>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(550).setHeight(400),
    'Grievance History'
  );
}
