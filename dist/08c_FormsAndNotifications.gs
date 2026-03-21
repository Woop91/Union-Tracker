/**
 * ============================================================================
 * 08c_FormsAndNotifications.gs - Form URLs, Pre-filling, and Notifications
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Form URL management, form pre-filling, notification system, and deadline
 *   alerts. getFormUrlFromConfig() reads form URLs from Config sheet.
 *   buildPreFilledGrievanceForm() generates pre-populated Google Form URLs.
 *   Notification functions send email alerts for deadlines, escalations, and
 *   overdue items.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Form URLs are stored in Config sheet (not hardcoded) so admins can swap
 *   Google Forms without code changes. Pre-filling speeds up grievance filing
 *   by auto-populating member info from the directory. Config data starts at
 *   row 3 (row 1=section headers, row 2=column headers) — a common source of
 *   off-by-one bugs documented in COLUMN_ISSUES_LOG.md.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Grievance forms can't be pre-filled — stewards must manually enter all
 *   fields. Deadline notifications stop — stewards miss filing deadlines.
 *   Email alerts for overdue items fail silently.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, CONFIG_COLS, GRIEVANCE_COLS, MEMBER_COLS),
 *   00_Security.gs (safeSendEmail_). Used by menu items, daily trigger in
 *   10_Main.gs, and form submission handlers in 10c.
 *
 * @fileoverview Form management and notification functions
 * @requires 01_Core.gs, 00_Security.gs
 */

// NOTE(F42): Form submission volume is acceptable for typical union usage (~100-5000 members).
// No throttling is needed at current scale.

// ============================================================================
// FORM URL CONFIGURATION
// ============================================================================

/**
 * Get form URL from Config sheet, with fallback to hardcoded defaults
 * @param {string} formType - Type of form: 'grievance', 'contact', or 'satisfaction'
 * @returns {string} The form URL
 */
function getFormUrlFromConfig(formType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) { Logger.log('Config sheet not found'); return ''; }

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
    // 'satisfaction' case removed v4.22.7 — Google Form deprecated, survey is now native webapp only
    default:
      Logger.log('Unknown form type: ' + formType);
      return '';
  }

  // M-15: Config data starts at row 3 (row 1 = section headers, row 2 = column headers)
  if (configSheet) {
    var url = configSheet.getRange(3, configCol).getValue();
    if (url && url !== '' && url.indexOf('http') === 0) {
      return url;
    }
  }

  // REL-02: Log a warning when falling back to default URL so admins know config is missing
  Logger.log('getFormUrlFromConfig: Config sheet missing or invalid for "' + formType + '" — using hardcoded default URL');
  return defaultUrl;
}

/**
 * Build pre-filled grievance form URL
 * @param {Object} memberData - Member information object
 * @param {Object} stewardData - Steward information object
 * @returns {string} Pre-filled form URL
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
/**
 * Silent version - used during CREATE_DASHBOARD setup
 * @param {Spreadsheet} ss - The spreadsheet object
 * @private
 */
function saveFormUrlsToConfig_silent(ss) {
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    Logger.log('Config sheet not found - cannot save form URLs');
    return;
  }

  // Config layout: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
  configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue('Grievance Form URL');
  configSheet.getRange(2, CONFIG_COLS.CONTACT_FORM_URL).setValue('Contact Form URL');
  // SATISFACTION_FORM_URL removed v4.22.7 — no longer written

  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue(GRIEVANCE_FORM_CONFIG.FORM_URL);
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setValue(CONTACT_FORM_CONFIG.FORM_URL);

  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
}

// ============================================================================
// FORM VALUE PARSING UTILITIES
// ============================================================================

/**
 * Get a value from form named responses
 * @param {Object} responses - Form named values object
 * @param {string} fieldName - Name of the field to retrieve
 * @returns {string} The field value or empty string
 * @private
 */
function getFormValue_(responses, fieldName) {
  if (responses[fieldName] && responses[fieldName].length > 0) {
    return responses[fieldName][0];
  }
  return '';
}

/**
 * Get multiple values from form response (for checkbox questions)
 * Returns comma-separated string
 * @param {Object} responses - Form named values object
 * @param {string} fieldName - Name of the field to retrieve
 * @returns {string} Comma-separated values or empty string
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
 * Parse a date string from form submission
 * @param {string} dateStr - Date string to parse
 * @returns {Date|string} Parsed Date object or original string if parsing fails
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
  } catch (_e) {
    return dateStr;
  }
}

// ============================================================================
// CONTACT FORM HANDLER
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
  var response = ui.alert('Personal Contact Info Form',
    'Share this form with members to collect their contact information.\n\n' +
    'When submitted, the data will be written to the Member Directory:\n' +
    '- Existing members (matched by name) will be updated\n' +
    '- New members will be added automatically\n\n' +
    '- Click YES to open the form\n' +
    '- Click NO to copy the link',
    ui.ButtonSet.YES_NO_CANCEL);

  if (response === ui.Button.YES) {
    // Open form in new window
    var html = HtmlService.createHtmlOutput(
      '<script>window.open(' + JSON.stringify(formUrl) + ', "_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    ui.showModalDialog(html, 'Opening form...');
  } else if (response === ui.Button.NO) {
    // Show link to copy
    var copyHtml = HtmlService.createHtmlOutput(
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + escapeHtml(formUrl) + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  var showOk = function() { document.getElementById("copied").style.display = "inline"; };' +
      '  if (navigator.clipboard) { navigator.clipboard.writeText(ta.value).then(showOk).catch(function() { ta.select(); document.execCommand("copy"); showOk(); }); }' +
      '  else { ta.select(); document.execCommand("copy"); showOk(); }' +
      '}' +
      '</script>' +
      '</div></body></html>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, 'Contact Form Link');
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
    var hireDate = getFormValue_(responses, 'Hire Date');
    var employeeId = getFormValue_(responses, 'Employee ID');
    var streetAddress = getFormValue_(responses, 'Street Address');
    var city = getFormValue_(responses, 'City');
    var zipCode = getFormValue_(responses, 'Zip Code');
    var state = getFormValue_(responses, 'State');

    // Require at least first and last name
    if (!firstName || !lastName) {
      Logger.log('Contact form submission missing name: ' + firstName + ' ' + lastName);
      return;
    }

    // Multi-Key Smart Match: Check ID, Email, then Name (hierarchical)
    var data = memberSheet.getDataRange().getValues();
    var memberId = getFormValue_(responses, 'Member ID');  // Optional field from form

    var match = findExistingMember({
      memberId: memberId,
      email: email,
      firstName: firstName,
      lastName: lastName
    }, data);

    var memberRow = match ? match.row : -1;

    if (match) {
      Logger.log('Found existing member via ' + match.matchType + ' match (confidence: ' + match.confidence + ') at row ' + match.row);
    }

    if (memberRow === -1) {
      // Member not found - create new member
      // Mask name in logs for privacy
      var maskedName = typeof maskName === 'function' ? maskName(firstName + ' ' + lastName) : '[REDACTED]';
      Logger.log('Creating new member: ' + maskedName);

      // Generate Member ID
      var existingIds = {};
      for (var k = 1; k < data.length; k++) {
        var id = data[k][MEMBER_COLS.MEMBER_ID - 1];
        if (id) existingIds[id] = true;
      }
      memberId = generateNameBasedId('M', firstName, lastName, existingIds);

      // Build new row array (escapeForFormula on all user-supplied strings to prevent formula injection)
      var newRow = [];
      newRow[MEMBER_COLS.MEMBER_ID - 1] = memberId;
      newRow[MEMBER_COLS.FIRST_NAME - 1] = escapeForFormula(firstName);
      newRow[MEMBER_COLS.LAST_NAME - 1] = escapeForFormula(lastName);
      newRow[MEMBER_COLS.JOB_TITLE - 1] = escapeForFormula(jobTitle || '');
      newRow[MEMBER_COLS.WORK_LOCATION - 1] = escapeForFormula(workLocation || '');
      newRow[MEMBER_COLS.UNIT - 1] = escapeForFormula(unit || '');
      newRow[MEMBER_COLS.OFFICE_DAYS - 1] = escapeForFormula(officeDays || '');
      newRow[MEMBER_COLS.EMAIL - 1] = escapeForFormula(email || '');
      newRow[MEMBER_COLS.PHONE - 1] = escapeForFormula(phone || '');
      newRow[MEMBER_COLS.PREFERRED_COMM - 1] = escapeForFormula(preferredComm || '');
      newRow[MEMBER_COLS.BEST_TIME - 1] = escapeForFormula(bestTime || '');
      newRow[MEMBER_COLS.SUPERVISOR - 1] = escapeForFormula(supervisor || '');
      newRow[MEMBER_COLS.MANAGER - 1] = escapeForFormula(manager || '');
      newRow[MEMBER_COLS.IS_STEWARD - 1] = 'No';
      newRow[MEMBER_COLS.INTEREST_LOCAL - 1] = escapeForFormula(interestLocal || '');
      newRow[MEMBER_COLS.INTEREST_CHAPTER - 1] = escapeForFormula(interestChapter || '');
      newRow[MEMBER_COLS.INTEREST_ALLIED - 1] = escapeForFormula(interestAllied || '');
      newRow[MEMBER_COLS.HIRE_DATE - 1] = hireDate ? parseFormDate_(hireDate) : '';
      newRow[MEMBER_COLS.EMPLOYEE_ID - 1] = escapeForFormula(employeeId || '');
      newRow[MEMBER_COLS.STREET_ADDRESS - 1] = escapeForFormula(streetAddress || '');
      newRow[MEMBER_COLS.CITY - 1] = escapeForFormula(city || '');
      newRow[MEMBER_COLS.STATE - 1] = escapeForFormula(state || '');
      newRow[MEMBER_COLS.ZIP_CODE - 1] = escapeForFormula(zipCode || '');

      // Append new member row
      memberSheet.appendRow(newRow);
      // Log with masked name for privacy
      Logger.log('Created new member ' + memberId + ': ' + maskedName);

      // Send acknowledgment email if email provided
      if (email) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + 'Welcome to WFSE Local',
            body: 'Hello ' + firstName + ',\n\n' +
              'Thank you for submitting your contact information. ' +
              'Your Member ID is: ' + memberId + '\n\n' +
              'Your information has been recorded and a union steward will be in touch.\n\n' +
              'Best regards,\nWFSE Local'
          });
        } catch (emailError) {
          Logger.log('Could not send welcome email: ' + emailError.message);
        }
      }

    } else {
      // Update existing member record with form data
      var updates = [];

      // Update all fields from form (even if they change existing values)
      // CR-18: Apply escapeForFormula() to all string values (matches new-member branch)
      if (jobTitle) updates.push({ col: MEMBER_COLS.JOB_TITLE, value: escapeForFormula(jobTitle) });
      if (unit) updates.push({ col: MEMBER_COLS.UNIT, value: escapeForFormula(unit) });
      if (workLocation) updates.push({ col: MEMBER_COLS.WORK_LOCATION, value: escapeForFormula(workLocation) });
      if (officeDays) updates.push({ col: MEMBER_COLS.OFFICE_DAYS, value: escapeForFormula(officeDays) });
      if (preferredComm) updates.push({ col: MEMBER_COLS.PREFERRED_COMM, value: escapeForFormula(preferredComm) });
      if (bestTime) updates.push({ col: MEMBER_COLS.BEST_TIME, value: escapeForFormula(bestTime) });
      if (supervisor) updates.push({ col: MEMBER_COLS.SUPERVISOR, value: escapeForFormula(supervisor) });
      if (manager) updates.push({ col: MEMBER_COLS.MANAGER, value: escapeForFormula(manager) });
      if (email) updates.push({ col: MEMBER_COLS.EMAIL, value: escapeForFormula(email) });
      if (phone) updates.push({ col: MEMBER_COLS.PHONE, value: escapeForFormula(phone) });
      if (interestLocal) updates.push({ col: MEMBER_COLS.INTEREST_LOCAL, value: escapeForFormula(interestLocal) });
      if (interestChapter) updates.push({ col: MEMBER_COLS.INTEREST_CHAPTER, value: escapeForFormula(interestChapter) });
      if (interestAllied) updates.push({ col: MEMBER_COLS.INTEREST_ALLIED, value: escapeForFormula(interestAllied) });
      if (hireDate) updates.push({ col: MEMBER_COLS.HIRE_DATE, value: parseFormDate_(hireDate) });
      if (employeeId) updates.push({ col: MEMBER_COLS.EMPLOYEE_ID, value: escapeForFormula(employeeId) });
      if (streetAddress) updates.push({ col: MEMBER_COLS.STREET_ADDRESS, value: escapeForFormula(streetAddress) });
      if (city) updates.push({ col: MEMBER_COLS.CITY, value: escapeForFormula(city) });
      if (state) updates.push({ col: MEMBER_COLS.STATE, value: escapeForFormula(state) });
      if (zipCode) updates.push({ col: MEMBER_COLS.ZIP_CODE, value: escapeForFormula(zipCode) });

      // M-PERF: Batch write — read row, apply all updates, write back in single call
      var totalCols = memberSheet.getLastColumn();
      var rowData = memberSheet.getRange(memberRow, 1, 1, totalCols).getValues()[0];
      for (var j = 0; j < updates.length; j++) {
        rowData[updates[j].col - 1] = updates[j].value;
      }
      memberSheet.getRange(memberRow, 1, 1, totalCols).setValues([rowData]);

      // Mask name in logs for privacy
      var maskedUpdateName = typeof maskName === 'function' ? maskName(firstName + ' ' + lastName) : '[REDACTED]';
      Logger.log('Updated contact info for ' + maskedUpdateName + ' (row ' + memberRow + ')');
    }

  } catch (error) {
    Logger.log('Error processing contact form submission: ' + error.message);
    throw error;
  }
}
// ============================================================================
// GRIEVANCE FORM TRIGGER SETUP
// ============================================================================
// ============================================================================
// SATISFACTION SURVEY HANDLER
// ============================================================================

/**
 * Show the Member Satisfaction Survey form link
 * Survey responses are written to the Member Satisfaction sheet
 */
function getSatisfactionSurveyLink() {
  var ui = SpreadsheetApp.getUi();
  // Get form URL from Config (allows admin to update without code changes)
  var formUrl = getFormUrlFromConfig('satisfaction');

  // Show dialog with form link options
  var response = ui.alert('Member Satisfaction Survey',
    'Share this survey with members to collect feedback.\n\n' +
    'When submitted, responses will be written to the\n' +
    'Member Satisfaction sheet.\n\n' +
    '- Click YES to open the survey\n' +
    '- Click NO to copy the link',
    ui.ButtonSet.YES_NO_CANCEL);

  if (response === ui.Button.YES) {
    // Open form in new window
    var html = HtmlService.createHtmlOutput(
      '<script>window.open(' + JSON.stringify(formUrl) + ', "_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    ui.showModalDialog(html, 'Opening survey...');
  } else if (response === ui.Button.NO) {
    // Show link to copy
    var copyHtml = HtmlService.createHtmlOutput(
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
      '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
      '<p>Copy this link and share with members:</p>' +
      '<textarea id="link" style="width: 100%; height: 80px; font-size: 12px;">' + escapeHtml(formUrl) + '</textarea>' +
      '<br><br>' +
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
      '<span id="copied" style="color: green; margin-left: 10px; display: none;">Copied!</span>' +
      '<script>' +
      'function copyLink() {' +
      '  var ta = document.getElementById("link");' +
      '  var showOk = function() { document.getElementById("copied").style.display = "inline"; };' +
      '  if (navigator.clipboard) { navigator.clipboard.writeText(ta.value).then(showOk).catch(function() { ta.select(); document.execCommand("copy"); showOk(); }); }' +
      '  else { ta.select(); document.execCommand("copy"); showOk(); }' +
      '}' +
      '</script>' +
      '</div></body></html>'
    ).setWidth(450).setHeight(180);
    ui.showModalDialog(copyHtml, 'Survey Link');
  }
}

/**
 * Handle satisfaction survey form submission.
 * Writes survey responses to the Member Satisfaction sheet, validates the
 * respondent's email against the Member Directory, and updates the survey
 * completion tracker if a member match is found.
 *
 * Triggered by: Google Forms onFormSubmit trigger
 *   (installed via setupSatisfactionFormTrigger() below)
 *
 * Survey tracking integration (added for _Survey_Tracking):
 *   After recording the response, if validateMemberEmail() finds a match,
 *   this function calls updateSurveyTrackingOnSubmit_(memberId) to mark
 *   the member as "Completed" in the tracking sheet. See the
 *   "SURVEY COMPLETION TRACKING" section below for full flow documentation.
 *
 * @param {Object} e - Form submission event object with e.namedValues
 */
/**
 * @deprecated v4.21.0 — Google Form ingest path removed.
 * In-app submitSurveyResponse() is now the only submission path.
 * v4.23.0 — Updated to use dynamic col map if somehow triggered.
 * Logs a warning and returns without writing to avoid column mismatch errors.
 */
function onSatisfactionFormSubmit(e) {
  Logger.log(
    'onSatisfactionFormSubmit called — this trigger is deprecated (v4.21.0). ' +
    'All survey responses must go through submitSurveyResponse() in the webapp. ' +
    'No data was written. Remove the Google Form trigger to stop this message.'
  );
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ Google Form survey trigger is deprecated. Use the member webapp to submit surveys.',
      'Deprecated Trigger', 8
    );
  } catch (_ui) { Logger.log('_ui: ' + (_ui.message || _ui)); }
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
    ui.alert('Trigger Exists',
      'A satisfaction survey trigger already exists.\n\n' +
      'Survey submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Satisfaction Survey Trigger',
    'This will set up automatic processing of survey submissions.\n\n' +
    'Enter the Google Form edit URL (the one ending in /edit):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var formUrl = response.getResponseText().trim();

  if (!formUrl) {
    ui.alert('No URL', 'Please provide the form edit URL.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Extract form ID from URL
    var match = formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      ui.alert('Invalid URL',
        'Could not extract form ID from URL.\n\n' +
        'Please use the form\'s edit URL. It should look like:\n' +
        'https://docs.google.com/forms/d/YOUR_FORM_ID/edit',
        ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onSatisfactionFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('Trigger Created',
      'Satisfaction survey trigger has been set up!\n\n' +
      'When a survey is submitted:\n' +
      '- Response will be added to Member Satisfaction sheet\n' +
      '- All 68 questions will be recorded\n' +
      '- Dashboard will reflect new data',
      ui.ButtonSet.OK);

    ss.toast('Survey trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

/**
 * Audits and optionally removes the deprecated onSatisfactionFormSubmit trigger.
 *
 * Background: setupSatisfactionFormTrigger() could be called at any time,
 * creating a form-submit trigger pointing at onSatisfactionFormSubmit().
 * That handler was deprecated in v4.21.0 and is now a no-op.  If the trigger
 * still exists it fires on every survey submission, wastes execution quota,
 * and logs a misleading deprecation warning.
 *
 * FIX-FORMS-01 (v4.25.9): Added to detect and remove the stale trigger.
 * Run once manually from the Admin menu or Dev Tools panel.
 *
 * @param {boolean} [autoDelete=false] If true, removes trigger without prompting.
 * @returns {boolean} true if a trigger was found and deleted, false otherwise.
 */
function auditAndRemoveSatisfactionTrigger(autoDelete) {
  var triggers = ScriptApp.getProjectTriggers();
  var found = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onSatisfactionFormSubmit') {
      found = triggers[i];
      break;
    }
  }

  if (!found) {
    if (!autoDelete) {
      SpreadsheetApp.getUi().alert(
        'Trigger Audit — Clean',
        'No stale satisfaction form trigger found. Nothing to remove.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    Logger.log('auditAndRemoveSatisfactionTrigger: no stale trigger found.');
    return false;
  }

  if (!autoDelete) {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert(
      'Stale Trigger Found',
      'A deprecated satisfaction survey trigger (onSatisfactionFormSubmit) is still active.\n\n' +
      'This trigger fires on every survey submission but does nothing (deprecated v4.21.0).\n' +
      'Removing it saves execution quota.\n\n' +
      'Remove this trigger now?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      Logger.log('auditAndRemoveSatisfactionTrigger: user declined removal.');
      return false;
    }
  }

  ScriptApp.deleteTrigger(found);
  Logger.log('auditAndRemoveSatisfactionTrigger: stale trigger removed successfully.');
  if (!autoDelete) {
    SpreadsheetApp.getUi().alert(
      'Trigger Removed',
      'The stale satisfaction form trigger has been deleted.\n\n' +
      'Survey submissions now use the in-app submitSurveyResponse() path only.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  return true;
}
/**
 * ============================================================================
 * 08h_NotificationEngine.gs - Notification and Alert System
 * ============================================================================
 *
 * This module handles all notification and alert functionality for the
 * SEIU Local Dashboard including:
 * - Deadline notification settings and triggers
 * - Steward deadline alerts
 * - Survey email distribution
 * - Member email validation
 * - Quarter utilities for notifications
 *
 * Dependencies:
 * - SHEETS constant (from 08_Code.gs)
 * - MEMBER_COLS constant (from 08_Code.gs)
 * - GRIEVANCE_COLS constant (from 08_Code.gs)
 * - CONFIG_COLS constant (from 08_Code.gs)
 * - SATISFACTION_FORM_CONFIG constant (from 08_Code.gs)
 *
 * @author SEIU Local
 * @version 1.0.0
 */

// ============================================================================
// NOTIFICATION SETTINGS AND TRIGGERS
// ============================================================================

/**
 * Show notification settings dialog
 * Allows user to enable/disable daily deadline notifications
 */
function showNotificationSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var enabled = props.getProperty('notifications_enabled') === 'true';
  var email = props.getProperty('notification_email') || Session.getActiveUser().getEmail();

  var response = ui.alert('Notification Settings',
    'Daily deadline notifications: ' + (enabled ? 'ENABLED' : 'DISABLED') + '\n' +
    'Email: ' + email + '\n\n' +
    'Notifications are sent daily at 8 AM for grievances due within 3 days.\n\n' +
    'Would you like to ' + (enabled ? 'DISABLE' : 'ENABLE') + ' notifications?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    if (enabled) {
      // Disable
      props.setProperty('notifications_enabled', 'false');
      removeDailyTrigger_();
      ui.alert('Notifications Disabled', 'Daily deadline notifications have been turned off.', ui.ButtonSet.OK);
    } else {
      // Enable
      props.setProperty('notifications_enabled', 'true');
      props.setProperty('notification_email', email);
      installDailyTrigger_();
      ui.alert('Notifications Enabled',
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
  var urgent = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    var daysToDeadline = data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var currentStep = data[i][GRIEVANCE_COLS.CURRENT_STEP - 1];

    var closedStatuses = GRIEVANCE_CLOSED_STATUSES;
    if (closedStatuses.indexOf(status) !== -1) continue;

    if (daysToDeadline === 'Overdue' || (daysToDeadline !== '' && typeof daysToDeadline === 'number' && daysToDeadline <= 3)) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });

      // NOTE: EventBus.emit removed v4.29.0 — EventBus is client-side only, never available in GAS triggers
    }
  }

  if (urgent.length === 0) return;

  var subject = COMMAND_CONFIG.SYSTEM_NAME + ': ' + urgent.length + ' Grievance Deadline(s) Approaching';
  var body = 'The following grievances have deadlines within 3 days:\n\n';

  for (var j = 0; j < urgent.length; j++) {
    var g = urgent[j];
    body += '* ' + g.id + ' (' + g.step + ') - ' +
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

  var response = ui.alert('Test Notifications',
    'This will send a test notification email to:\n' + email + '\n\nSend test email?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    MailApp.sendEmail(email,
      COMMAND_CONFIG.SYSTEM_NAME + ' Test Notification',
      'This is a test notification from your ' + COMMAND_CONFIG.SYSTEM_NAME + '.\n\n' +
      'If you received this email, notifications are working correctly!\n\n' +
      'Dashboard: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    );
    ui.alert('Test Sent', 'Test email sent to ' + email + '\n\nCheck your inbox!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to send test email: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// STEWARD DEADLINE ALERTS
// ============================================================================

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
  var closedStatuses = GRIEVANCE_CLOSED_STATUSES;

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';

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

    // NOTE: EventBus.emit removed v4.29.0 — EventBus is client-side only, never available in GAS triggers
  }

  // Get steward emails from Member Directory (stewards are members with IS_STEWARD = Yes)
  var stewardEmails = {};
  for (var s = 1; s < memberData.length; s++) {
    var isSteward = memberData[s][MEMBER_COLS.IS_STEWARD - 1];
    if (isTruthyValue(isSteward)) {
      var sFirstName = memberData[s][MEMBER_COLS.FIRST_NAME - 1] || '';
      var sLastName = memberData[s][MEMBER_COLS.LAST_NAME - 1] || '';
      var sFullName = (sFirstName + ' ' + sLastName).trim();
      var sEmail = memberData[s][MEMBER_COLS.EMAIL - 1] || '';
      if (sFullName && sEmail && sEmail.indexOf('@') !== -1) {
        stewardEmails[sFullName] = sEmail;
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

    var body = 'GRIEVANCE DEADLINE ALERT\n';
    body += '====================================\n\n';
    body += 'Steward: ' + stewardName + '\n';
    body += 'Date: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') + '\n\n';

    if (overdue.length > 0) {
      body += 'OVERDUE (' + overdue.length + ')\n';
      body += '---------------------\n';
      for (var o = 0; o < overdue.length; o++) {
        body += '  [!] ' + overdue[o].id + ' - ' + overdue[o].memberName + '\n';
        body += '     Step: ' + overdue[o].step + ' | Status: ' + overdue[o].status + '\n';
        body += '     OVERDUE by ' + Math.abs(overdue[o].daysRemaining) + ' day(s)\n\n';
      }
    }

    if (urgent.length > 0) {
      body += 'URGENT - Due within 3 days (' + urgent.length + ')\n';
      body += '---------------------\n';
      for (var u = 0; u < urgent.length; u++) {
        body += '  [*] ' + urgent[u].id + ' - ' + urgent[u].memberName + '\n';
        body += '     Step: ' + urgent[u].step + ' | Status: ' + urgent[u].status + '\n';
        body += '     Due in ' + urgent[u].daysRemaining + ' day(s)\n\n';
      }
    }

    if (upcoming.length > 0) {
      body += 'UPCOMING - Due within ' + alertDays + ' days (' + upcoming.length + ')\n';
      body += '---------------------\n';
      for (var up = 0; up < upcoming.length; up++) {
        body += '  [-] ' + upcoming[up].id + ' - ' + upcoming[up].memberName + '\n';
        body += '     Step: ' + upcoming[up].step + ' | Due in ' + upcoming[up].daysRemaining + ' day(s)\n\n';
      }
    }

    body += '====================================\n';
    body += 'Dashboard: ' + ss.getUrl() + '\n';
    body += 'Total grievances requiring attention: ' + grievances.length + '\n';

    var subject = (overdue.length > 0 ? 'OVERDUE: ' : '') +
      grievances.length + ' Grievance Deadline(s) - ' + stewardName;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: getOrgNameFromConfig_() + ' Dashboard'
      });
      emailsSent++;
      // Mask name in logs for privacy
      var maskedSteward = typeof maskName === 'function' ? maskName(stewardName) : '[REDACTED]';
      Logger.log('Sent alert to ' + maskedSteward + ': ' + grievances.length + ' grievances');
    } catch (e) {
      Logger.log('Failed to send steward alert: ' + e.message);
    }
  }

  Logger.log('Steward deadline alerts complete. Sent ' + emailsSent + ' emails.');
  return emailsSent;
}

// ============================================================================
// SURVEY EMAIL DISTRIBUTION
// ============================================================================
/**
 * Execute sending random survey emails
 * @param {Object} opts - Options {count, subject, excludeDays}
 * @returns {string} Result message
 */
function executeSendRandomSurveyEmails(opts) {
  // Validate parameters from client-side input
  opts = opts || {};
  if (typeof opts.subject !== 'string' || opts.subject.length === 0 || opts.subject.length > 200) {
    throw new Error('Invalid subject: must be a non-empty string of 200 characters or fewer.');
  }
  opts.count = parseInt(opts.count, 10);
  if (isNaN(opts.count) || opts.count < 1 || opts.count > 1000) {
    throw new Error('Invalid count: must be a positive integer up to 1000.');
  }
  opts.excludeDays = parseInt(opts.excludeDays, 10);
  if (isNaN(opts.excludeDays) || opts.excludeDays < 0) {
    throw new Error('Invalid excludeDays: must be a non-negative integer.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!memberSheet) throw new Error('Member Directory not found');
  if (!configSheet) throw new Error('Config sheet not found');

  // Get all members with valid emails
  var memberData = memberSheet.getDataRange().getValues();
  var emailCol = MEMBER_COLS.EMAIL - 1;
  var memberIdCol = MEMBER_COLS.MEMBER_ID - 1;
  var firstNameCol = MEMBER_COLS.FIRST_NAME - 1;
  var lastNameCol = MEMBER_COLS.LAST_NAME - 1;

  // Get survey email log from Config (uses dedicated columns, not PDF_FOLDER_ID)
  var surveyLogCol = CONFIG_COLS.SURVEY_LOG_IDS;
  var surveyLog = {};
  try {
    if (configSheet && configSheet.getLastRow() > 1) {
      var logData = configSheet.getRange(2, surveyLogCol, configSheet.getLastRow() - 1, 2).getValues();
      logData.forEach(function(row) {
        if (row[0]) {
          var parsed = new Date(row[1]);
          if (!isNaN(parsed.getTime())) surveyLog[row[0]] = parsed;
        }
      });
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
  // Fisher-Yates shuffle for uniform distribution
  var shuffled = eligibleMembers.slice();
  for (var si = shuffled.length - 1; si > 0; si--) {
    var sj = Math.floor(Math.random() * (si + 1));
    var temp = shuffled[si]; shuffled[si] = shuffled[sj]; shuffled[sj] = temp;
  }
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
        getOrgNameFromConfig_();

      MailApp.sendEmail({
        to: member.email,
        subject: opts.subject,
        body: body,
        name: getOrgNameFromConfig_() + ' Dashboard'
      });

      sent++;
      newLogEntries.push([member.memberId, new Date()]);
    } catch(e) {
      errors.push(member.firstName + ' ' + member.lastName + ': ' + e.message);
    }
  });

  // Update survey email log - find actual last row in log column to avoid overwriting
  if (newLogEntries.length > 0 && configSheet) {
    var nextRow = 2;
    if (configSheet.getLastRow() > 1) {
      var logValues = configSheet.getRange(2, surveyLogCol, configSheet.getLastRow() - 1, 1).getValues();
      for (var lr = logValues.length - 1; lr >= 0; lr--) {
        if (logValues[lr][0]) { nextRow = lr + 3; break; }
      }
    }
    configSheet.getRange(nextRow, surveyLogCol, newLogEntries.length, 2).setValues(newLogEntries);
  }

  var result = 'Sent ' + sent + ' survey emails';
  if (errors.length > 0) {
    result += '\n\n' + errors.length + ' errors:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) result += '\n...and ' + (errors.length - 5) + ' more';
  }

  return result;
}

// ============================================================================
// MEMBER VALIDATION UTILITIES
// ============================================================================

/**
 * Validate that an email belongs to a member in the directory.
 * Scans Member Directory column I (MEMBER_COLS.EMAIL) for a case-insensitive match.
 *
 * This is the critical link in survey completion detection:
 *   onSatisfactionFormSubmit() extracts the respondent's email from the Google Form,
 *   then calls this function. If a match is found, the returned memberId is passed
 *   to updateSurveyTrackingOnSubmit_() to mark the survey as completed in the
 *   _Survey_Tracking sheet. If no match, the response is flagged "Pending Review"
 *   and tracking status stays "Not Completed".
 *
 * @param {string} email - Email to validate (case-insensitive comparison)
 * @returns {Object|null} { memberId, firstName, lastName } if match found, null otherwise
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

// ============================================================================
// QUARTER UTILITIES FOR NOTIFICATIONS
// ============================================================================

/**
 * Get the current quarter string (e.g., "2026-Q1")
 * @returns {string} Quarter string
 */
function getCurrentQuarter() {
  var now = new Date();
  var quarter = Math.floor(now.getMonth() / 3) + 1;
  return now.getFullYear() + '-Q' + quarter;
}
// ============================================================================
// SURVEY COMPLETION TRACKING
// ============================================================================
//
// This section implements the per-member survey completion tracker.
// It monitors which members have completed the satisfaction survey in the
// current round and maintains cumulative participation history.
//
// ── HOW COMPLETION IS DETECTED ──────────────────────────────────────────────
//
//   The system detects survey completion via email matching on form submit:
//
//   1. TRIGGER: setupSatisfactionFormTrigger() (below in this file) installs a
//      Google Apps Script trigger that fires onSatisfactionFormSubmit(e) each
//      time a member submits the satisfaction Google Form.
//
//   2. EMAIL EXTRACTION: onSatisfactionFormSubmit() extracts the respondent's
//      email from the form event by trying, in order:
//        - Form field named "Email Address"
//        - Form field named "Email" or "email"
//        - Google's built-in e.response.getRespondentEmail()
//
//   3. EMAIL-TO-MEMBER MATCHING: validateMemberEmail(email) (this file, ~line
//      1518) scans the Member Directory sheet column I (MEMBER_COLS.EMAIL) for
//      a case-insensitive match. Returns { memberId, firstName, lastName } or null.
//
//   4. TRACKING UPDATE: If a match is found, onSatisfactionFormSubmit() calls
//      updateSurveyTrackingOnSubmit_(memberMatch.memberId) which finds the
//      member's row in the _Survey_Tracking sheet and writes:
//        - Current Status   = "Completed"
//        - Completed Date   = current timestamp
//        - Total Completed  = previous value + 1
//
//   5. NO MATCH: If the email doesn't match any member (typo, personal email,
//      non-member), the survey response is still recorded in the Satisfaction
//      sheet (flagged as "Pending Review"), but the tracker is NOT updated —
//      that member stays "Not Completed" for the round.
//
// ── ROUND MANAGEMENT ────────────────────────────────────────────────────────
//
//   - startNewSurveyRound(): Resets all members to "Not Completed", increments
//     Total Missed for anyone who didn't complete the previous round.
//   - sendSurveyCompletionReminders(): Emails non-respondents with a 7-day
//     cooldown between reminders. Uses survey URL from Config sheet (col AR).
//   - showSurveyTrackingDialog(): Management UI showing stats + action buttons.
//
// ── DATA FLOW ───────────────────────────────────────────────────────────────
//
//   Member Directory ──populate()──> _Survey_Tracking (hidden sheet)
//        (source)                         │
//   Google Form ─> onSatisfactionFormSubmit() ─> updateSurveyTrackingOnSubmit_()
//                                         │
//   showSurveyTrackingDialog() <──────────┘ reads stats
//
// ── RELATED CODE ────────────────────────────────────────────────────────────
//
//   Constants:    SURVEY_TRACKING_COLS    in 01_Core.gs
//   Sheet name:   SHEETS.SURVEY_TRACKING  in 01_Core.gs ('_Survey_Tracking')
//   Sheet setup:  setupSurveyTrackingSheet()  in 08d_AuditAndFormulas.gs
//   Hidden init:  setupHiddenSheets()     in 08a_SheetSetup.gs
//   Seed data:    seedSurveyTrackingData() in 07_DevTools.gs
//
// ============================================================================

/**
 * Populates the _Survey_Tracking sheet from the Member Directory.
 * Creates one row per active member with their info and initial status.
 * Safe to re-run: clears existing data rows and rebuilds from directory.
 */
function populateSurveyTrackingFromMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!trackingSheet || !memberSheet) {
    Logger.log('populateSurveyTracking: Required sheets not found');
    return;
  }

  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) {
    Logger.log('populateSurveyTracking: No members found');
    return;
  }

  // M-39: Only clear tracking data if no existing survey responses exist.
  // If any rows have a Completed Date (response timestamp), warn and skip the clear
  // to avoid destroying manually entered or response-driven data.
  if (trackingSheet.getLastRow() > 1) {
    var existingData = trackingSheet.getRange(2, 1, trackingSheet.getLastRow() - 1, SURVEY_TRACKING_HEADER_MAP_.length).getValues();
    var hasResponseData = false;
    for (var chk = 0; chk < existingData.length; chk++) {
      // Check Completed Date column for any response timestamps
      if (existingData[chk][SURVEY_TRACKING_COLS.COMPLETED_DATE - 1]) {
        hasResponseData = true;
        break;
      }
    }
    if (hasResponseData) {
      Logger.log('populateSurveyTracking: Existing survey response data found — skipping clear to avoid data loss. Use startNewSurveyRound() to reset.');
      return;
    }
    trackingSheet.getRange(2, 1, trackingSheet.getLastRow() - 1, SURVEY_TRACKING_HEADER_MAP_.length).clear();
  }

  var rows = [];
  for (var i = 1; i < memberData.length; i++) {
    var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    var firstName = memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = memberData[i][MEMBER_COLS.LAST_NAME - 1] || '';

    rows.push([
      memberId,                                                        // A - Member ID
      (firstName + ' ' + lastName).trim(),                             // B - Member Name
      memberData[i][MEMBER_COLS.EMAIL - 1] || '',                      // C - Email
      memberData[i][MEMBER_COLS.WORK_LOCATION - 1] || '',              // D - Work Location
      memberData[i][MEMBER_COLS.ASSIGNED_STEWARD - 1] || '',           // E - Assigned Steward
      'Not Completed',                                                 // F - Current Status
      '',                                                              // G - Completed Date
      0,                                                               // H - Total Missed
      0,                                                               // I - Total Completed
      ''                                                               // J - Last Reminder Sent
    ]);
  }

  if (rows.length > 0) {
    trackingSheet.getRange(2, 1, rows.length, SURVEY_TRACKING_HEADER_MAP_.length).setValues(rows);
  }

  Logger.log('populateSurveyTracking: Populated ' + rows.length + ' member rows');
}

/**
 * Updates survey tracking when a satisfaction form is submitted.
 * Marks the member as "Completed" for the current round.
 * Called from onSatisfactionFormSubmit when a member match is found.
 * @param {string} matchedMemberId - The matched member ID from form submission
 */
function updateSurveyTrackingOnSubmit_(matchedMemberId) {
  if (!matchedMemberId) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);
  if (!trackingSheet) return;

  var data = trackingSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][SURVEY_TRACKING_COLS.MEMBER_ID - 1] === matchedMemberId) {
      // M-PERF: Batch write — read row, modify 3 cells, write back in single call
      var rowNum = i + 1;
      var totalCols = trackingSheet.getLastColumn();
      var rowData = trackingSheet.getRange(rowNum, 1, 1, totalCols).getValues()[0];
      rowData[SURVEY_TRACKING_COLS.CURRENT_STATUS - 1] = 'Completed';
      rowData[SURVEY_TRACKING_COLS.COMPLETED_DATE - 1] = new Date();
      var prevCompleted = parseInt(data[i][SURVEY_TRACKING_COLS.TOTAL_COMPLETED - 1]) || 0;
      rowData[SURVEY_TRACKING_COLS.TOTAL_COMPLETED - 1] = prevCompleted + 1;
      trackingSheet.getRange(rowNum, 1, 1, totalCols).setValues([rowData]);
      Logger.log('Survey tracking updated for member: ' + matchedMemberId);
      return;
    }
  }

  Logger.log('Survey tracking: member ' + matchedMemberId + ' not found in tracking sheet');
}

/**
 * Starts a new survey round. Resets Current Status to "Not Completed"
 * for all members and clears Completed Date. Increments Total Missed
 * for members who did not complete the previous round.
 */
function startNewSurveyRound() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);

  if (!trackingSheet || trackingSheet.getLastRow() < 2) {
    Logger.log('startNewSurveyRound: No tracking data found');
    try {
      SpreadsheetApp.getUi().alert('No survey tracking data found. Run "Populate Survey Tracking" first.');
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    return;
  }

  var data = trackingSheet.getDataRange().getValues();
  var updates = [];

  for (var i = 1; i < data.length; i++) {
    var currentStatus = data[i][SURVEY_TRACKING_COLS.CURRENT_STATUS - 1];
    var totalMissed = parseInt(data[i][SURVEY_TRACKING_COLS.TOTAL_MISSED - 1]) || 0;

    // Increment missed count for members who didn't complete last round
    if (currentStatus !== 'Completed') {
      totalMissed++;
    }

    updates.push([
      'Not Completed',  // F - Reset status
      '',               // G - Clear completed date
      totalMissed       // H - Updated missed count
    ]);
  }

  if (updates.length > 0) {
    trackingSheet.getRange(2, SURVEY_TRACKING_COLS.CURRENT_STATUS, updates.length, 3)
      .setValues(updates);
  }

  Logger.log('startNewSurveyRound: Reset ' + updates.length + ' members for new round');
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'New survey round started. ' + updates.length + ' members reset to Not Completed.',
      'Survey Tracking', 5
    );
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
}

/**
 * Sends reminder emails to members who have not completed the current survey round.
 * Only sends to members with a valid email and Current Status = "Not Completed".
 * Respects a configurable cooldown period (default 7 days) between reminders.
 */
function sendSurveyCompletionReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) { Logger.log('Config sheet not found'); }

  if (!trackingSheet || trackingSheet.getLastRow() < 2) {
    Logger.log('sendSurveyCompletionReminders: No tracking data');
    return 'No survey tracking data found.';
  }

  // Survey is now native webapp — direct members to the member portal URL
  var surveyUrl = '';
  if (configSheet && configSheet.getLastRow() >= 3) {
    try {
      surveyUrl = configSheet.getRange(3, CONFIG_COLS.MOBILE_DASHBOARD_URL).getValue() || '';
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  }

  var data = trackingSheet.getDataRange().getValues();
  var now = new Date();
  var cooldownDays = 7;
  var sentCount = 0;
  var skippedCount = 0;

  var orgName = '';
  try {
    orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue() || 'Your Union';
  } catch (_e) {
    orgName = 'Your Union';
  }

  for (var i = 1; i < data.length; i++) {
    var status = data[i][SURVEY_TRACKING_COLS.CURRENT_STATUS - 1];
    var email = (data[i][SURVEY_TRACKING_COLS.EMAIL - 1] || '').toString().trim();
    var memberName = data[i][SURVEY_TRACKING_COLS.MEMBER_NAME - 1] || 'Member';
    var lastReminder = data[i][SURVEY_TRACKING_COLS.LAST_REMINDER_SENT - 1];

    if (status === 'Completed' || !email || !email.includes('@')) continue;

    // Skip if reminder was sent within cooldown period
    if (lastReminder) {
      var reminderDate = new Date(lastReminder);
      if (isNaN(reminderDate.getTime())) continue;
      var daysSinceReminder = (now - reminderDate) / (1000 * 60 * 60 * 24);
      if (daysSinceReminder < cooldownDays) {
        skippedCount++;
        continue;
      }
    }

    try {
      var body = 'Dear ' + memberName + ',\n\n' +
        'This is a friendly reminder that we have not yet received your response to the current member satisfaction survey.\n\n' +
        'Your feedback is valuable and helps us improve our representation and services.\n\n';

      if (surveyUrl) {
        body += 'Please complete the survey here: ' + surveyUrl + '\n\n';
      }

      body += 'Thank you for your participation.\n\n' +
        'In Solidarity,\n' + orgName;

      MailApp.sendEmail({
        to: email,
        subject: (typeof COMMAND_CONFIG !== 'undefined' ? COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX : '') + 'Survey Completion Reminder',
        body: body
      });

      var rowNum = i + 1;
      trackingSheet.getRange(rowNum, SURVEY_TRACKING_COLS.LAST_REMINDER_SENT).setValue(now);
      sentCount++;
    } catch (emailError) {
      // M-14: Use secureLog/maskEmail instead of logging raw email
      if (typeof secureLog === 'function') {
        secureLog('sendSurveyCompletionReminders', 'Reminder email failed', { email: email, error: emailError.message });
      } else {
        var masked = typeof maskEmail === 'function' ? maskEmail(email) : '[REDACTED]';
        Logger.log('Reminder email failed for ' + masked + ': ' + emailError.message);
      }
    }
  }

  var result = 'Reminders sent: ' + sentCount + ', Skipped (cooldown): ' + skippedCount;
  Logger.log('sendSurveyCompletionReminders: ' + result);
  return result;
}

/**
 * Shows survey completion statistics in a toast/alert.
 * Provides a quick overview of current round participation.
 */
function getSurveyCompletionStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);

  if (!trackingSheet || trackingSheet.getLastRow() < 2) {
    return { total: 0, completed: 0, notCompleted: 0, rate: '0%' };
  }

  var data = trackingSheet.getDataRange().getValues();
  var total = 0;
  var completed = 0;

  for (var i = 1; i < data.length; i++) {
    if (!data[i][SURVEY_TRACKING_COLS.MEMBER_ID - 1]) continue;
    total++;
    if (data[i][SURVEY_TRACKING_COLS.CURRENT_STATUS - 1] === 'Completed') {
      completed++;
    }
  }

  var notCompleted = total - completed;
  var rate = total > 0 ? Math.round((completed / total) * 100) + '%' : '0%';

  return {
    total: total,
    completed: completed,
    notCompleted: notCompleted,
    rate: rate
  };
}

/**
 * Shows a dialog with survey tracking management options
 */
function showSurveyTrackingDialog() {
  var stats = getSurveyCompletionStats();

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px;max-width:500px}' +
    'h2{color:#7c3aed;margin-top:0}' +
    '.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin:15px 0}' +
    '.stat{background:#f8f9fa;padding:15px;border-radius:8px;text-align:center}' +
    '.stat .num{font-size:28px;font-weight:bold;color:#7c3aed}' +
    '.stat .label{font-size:12px;color:#6b7280;margin-top:4px}' +
    '.stat.green .num{color:#10b981}' +
    '.stat.red .num{color:#ef4444}' +
    '.btn{display:block;width:100%;padding:12px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:600;margin:8px 0;box-sizing:border-box}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-success{background:#10b981;color:white}' +
    '.btn-warning{background:#f59e0b;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151}' +
    '.info{background:#eff6ff;padding:10px;border-radius:6px;font-size:13px;color:#1e40af;margin:10px 0}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>Survey Completion Tracker</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + stats.total + '</div><div class="label">Total Members</div></div>' +
    '<div class="stat green"><div class="num">' + stats.completed + '</div><div class="label">Completed</div></div>' +
    '<div class="stat red"><div class="num">' + stats.notCompleted + '</div><div class="label">Not Completed</div></div>' +
    '<div class="stat"><div class="num">' + stats.rate + '</div><div class="label">Completion Rate</div></div>' +
    '</div>' +
    '<div class="info">Survey tracking monitors member participation across survey rounds. ' +
    'Completion is auto-tracked when members submit the satisfaction survey.</div>' +
    '<button class="btn btn-primary" onclick="run(\'populateSurveyTrackingFromMembers\')">Refresh Member List</button>' +
    '<button class="btn btn-warning" onclick="confirmNewRound()">Start New Survey Round</button>' +
    '<button class="btn btn-success" onclick="run(\'sendSurveyCompletionReminders\')">Send Reminders</button>' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<script>' +
    'function run(fn){' +
    '  document.querySelectorAll(".btn").forEach(function(b){b.disabled=true;b.style.opacity="0.6"});' +
    '  google.script.run.withSuccessHandler(function(r){alert(r||"Done!");location.reload()})' +
    '  .withFailureHandler(function(e){alert("Error: "+e.message);location.reload()})[fn]();}' +
    'function confirmNewRound(){' +
    '  if(confirm("Start a new survey round?\\n\\nThis will:\\n- Reset all members to Not Completed\\n- Increment missed count for non-respondents\\n\\nContinue?")){' +
    '    run("startNewSurveyRound");}}' +
    '</script></body></html>'
  ).setWidth(520).setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, 'Survey Completion Tracker');
}

/**
 * ============================================================================
 * AUDIT LOG MODULE (08i_AuditLog.gs)
 * ============================================================================
 *
 * This module provides audit logging functionality for tracking changes
 * to the Member Directory and Grievance Log sheets.
 *
 * Features:
 * - Automatic change tracking via onEdit trigger
 * - Audit log sheet setup and management
 * - History viewing and cleanup utilities
 * - Record-specific audit trail retrieval
 *
 * Dependencies:
 * - SHEETS constant (for sheet names)
 * - COLORS constant (for formatting)
 * - logAuditEvent() function (from core module)
 *
 * @author Union Membership System
 * @version 1.0.0
 * ============================================================================
 */

// ============================================================================
// IN-APP SURVEY (v4.15.0) — Mobile-optimized questionnaire
// ============================================================================

/**
 * Returns structured survey questions derived from Member Satisfaction headers.
 * Groups questions into sections for the multi-step wizard.
 * @returns {Object} { sections: [{title, questions: [{id, text, type}]}] }
 */
/**
 * Returns the full 67-question survey definition, reading context options
 * dynamically from Config (worksites, roles, Q64 priority options).
 *
 * Each question object:
 *   { id, col, text, type, section, sectionKey, required,
 *     branchParent, branchValue,      // only if conditional
 *     options,                        // dropdown/radio/checkbox
 *     labelMin, labelMax,             // slider only (1-10)
 *     maxSelections }                 // checkbox only
 *
 * Types: 'slider-10' | 'dropdown' | 'radio' | 'radio-branch' | 'checkbox' | 'paragraph'
 * Slider labels: 1 = Strongly Disagree / 10 = Strongly Agree (universal Likert)
 *
 * @returns {Object} { period, sections, questions, sliderLabels }
 */
// ============================================================================
// SURVEY SCHEMA — v4.23.0 Dynamic (reads from 📋 Survey Questions sheet)
// ============================================================================

/**
 * Returns the active survey questions, reading from the 📋 Survey Questions sheet.
 * Caches result for 5 minutes. Call clearSurveyQuestionsCache() to force refresh.
 *
 * Falls back to creating the sheet (with seed data) if it doesn't exist.
 *
 * @returns {{ questions: Array, sections: Array, sliderLabels: Object, period: Object|null }}
 */
function getSurveyQuestions() {
  var CACHE_KEY = 'surveyQuestions_v1';
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ensure Survey Questions sheet exists
    var qSheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() < 2) {
      createSurveyQuestionsSheet(ss);
      qSheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    }

    var QC = SURVEY_QUESTIONS_COLS;

    // Dynamic Config options
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    function _configList(col) {
      if (!configSheet || !col || col < 1) return [];
      try {
        var lr = configSheet.getLastRow();
        if (lr < 3) return [];
        return configSheet.getRange(3, col, lr - 2, 1).getValues()
          .map(function(r) { return String(r[0]).trim(); })
          .filter(function(v) { return v !== ''; });
      } catch(_e) { return []; }
    }
    var worksites       = _configList(CONFIG_COLS.OFFICE_LOCATIONS);
    var roles           = _configList(CONFIG_COLS.JOB_TITLES);
    var priorityOptions = _configList(CONFIG_COLS.SURVEY_PRIORITY_OPTIONS);

    if (!worksites.length)       worksites       = ['Please add worksites to Config tab'];
    if (!roles.length)           roles           = ['Please add roles to Config tab'];
    if (!priorityOptions.length) priorityOptions = [
      'Contract Enforcement','Workload','Scheduling','Pay & Benefits',
      'Safety','Training','Equity & Inclusion','Communication',
      'Steward Support','Organizing','Other'
    ];

    // Read all rows from Survey Questions sheet
    var rawData = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, 16).getValues();

    var questions  = [];
    var seenSections = {};
    var sections   = [];

    rawData.forEach(function(row) {
      var id         = String(row[QC.QUESTION_ID    - 1] || '').trim();
      var secNum     = String(row[QC.SECTION_NUM    - 1] || '').trim();
      var secKey     = String(row[QC.SECTION_KEY    - 1] || '').trim();
      var secTitle   = String(row[QC.SECTION_TITLE  - 1] || '').trim();
      var text       = String(row[QC.QUESTION_TEXT  - 1] || '').trim();
      var type       = String(row[QC.TYPE           - 1] || '').trim().toLowerCase();
      var required   = String(row[QC.REQUIRED       - 1] || 'Y').trim().toUpperCase() !== 'N';
      var active     = String(row[QC.ACTIVE         - 1] || 'Y').trim().toUpperCase() !== 'N';
      var optRaw     = String(row[QC.OPTIONS        - 1] || '').trim();
      var branchPar  = String(row[QC.BRANCH_PARENT  - 1] || '').trim();
      var branchVal  = String(row[QC.BRANCH_VALUE   - 1] || '').trim();
      var branchTgt  = String(row[QC.BRANCH_TARGET  - 1] || '').trim();
      var maxSel     = parseInt(String(row[QC.MAX_SELECTIONS - 1] || ''), 10) || 0;
      var slMin      = String(row[QC.SLIDER_MIN     - 1] || 'Strongly Disagree').trim() || 'Strongly Disagree';
      var slMax      = String(row[QC.SLIDER_MAX     - 1] || 'Strongly Agree').trim()   || 'Strongly Agree';

      if (!id || !text) return; // skip blank rows

      // Required questions are always shown even if Active=N in the sheet
      var effectiveActive = active || required;

      // Resolve options
      var options = [];
      if (optRaw === '[Config: Office Locations]') {
        options = worksites;
      } else if (optRaw === '[Config: Job Titles]') {
        options = roles;
      } else if (optRaw === '[Config: Survey Priority Options]') {
        options = priorityOptions;
      } else if (optRaw) {
        options = optRaw.split('|').map(function(v) { return v.trim(); }).filter(Boolean);
      }

      // Build question object (matches webapp expectations)
      var q = {
        id:          id,
        text:        text,
        type:        type,
        section:     secNum,
        sectionKey:  secKey,
        sectionTitle: secTitle,
        required:    required,
        active:      effectiveActive,
        options:     options,
        labelMin:    slMin,
        labelMax:    slMax
      };
      if (branchPar) {
        q.branchParent = branchPar;
        q.branchValue  = branchVal;
        if (type === 'radio-branch' && branchTgt) {
          // For root branch questions (q5, q36), store branch target
          q.branchYes = (branchVal === 'Yes' || branchVal === '') ? branchTgt : '';
          q.branchNo  = (branchVal === 'No')                      ? branchTgt : '';
        }
      }
      if (type === 'radio-branch' && !branchPar) {
        // Root branch question: branchYes/branchNo from Notes or derived
        // Default derivation from seed: q5→3A/3B, q36→6A/7
        if (id === 'q5')  { q.branchYes = '3A'; q.branchNo = '3B'; }
        if (id === 'q36') { q.branchYes = '6A'; q.branchNo = '7';  }
      }
      if (maxSel > 0) q.maxSelections = maxSel;

      questions.push(q);

      // Build sections list (unique, in order of first appearance)
      if (secKey && !seenSections[secKey]) {
        seenSections[secKey] = true;
        sections.push({ key: secKey, number: secNum, title: secTitle });
      }
    });

    var result = {
      questions:   questions,
      sections:    sections,
      sliderLabels: { min: 'Strongly Disagree', max: 'Strongly Agree' },
      period:      getSurveyPeriod()
    };

    // Cache 5 minutes
    try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(result), 300); } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }

    return result;

  } catch(e) {
    Logger.log('getSurveyQuestions error: ' + e.message);
    return { questions: [], sections: [], sliderLabels: { min: 'Strongly Disagree', max: 'Strongly Agree' }, period: null };
  }
}

// ── Satisfaction sheet column map (runtime, cached) ─────────────────────────

/**
 * Returns a map of header → 1-indexed column number for the Satisfaction sheet.
 * Covers fixed prefix ('Timestamp', 'Period ID', 'Survey Version') + all question IDs (q1, q2…).
 * Cached for 5 minutes. Invalidated by syncSatisfactionSheetColumns_() and clearSurveyQuestionsCache().
 *
 * @returns {Object} e.g. { 'Timestamp': 1, 'Period ID': 2, 'q1': 4, 'q2': 5, … }
 */
function getSatisfactionColMap_() {
  var CACHE_KEY = 'satisfactionColMap_v1';
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }

  try {
    var satSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SATISFACTION);
    if (!satSheet || satSheet.getLastColumn() < 1) return {};

    var headers = satSheet.getRange(1, 1, 1, satSheet.getLastColumn()).getValues()[0];
    var map = {};
    headers.forEach(function(h, i) {
      var key = String(h || '').trim();
      if (key) map[key] = i + 1;
    });

    try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(map), 300); } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }
    return map;

  } catch(e) {
    Logger.log('getSatisfactionColMap_ error: ' + e.message);
    return {};
  }
}

/**
 * Ensures all active questions have a corresponding column in the Satisfaction sheet.
 * Appends any missing question ID columns to the right. Never removes or moves columns.
 * Invalidates the col map cache after adding columns.
 *
 * @param {Array} activeQuestions - Array of question objects from getSurveyQuestions()
 * @returns {Object} Updated col map
 */
function syncSatisfactionSheetColumns_(activeQuestions) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!satSheet) return {};

    // Ensure header row exists
    if (satSheet.getLastRow() < 1 || satSheet.getLastColumn() < 1) {
      // Sheet empty — rebuild via createSatisfactionSheet
      createSatisfactionSheet(ss);
      return getSatisfactionColMap_();
    }

    var colMap = getSatisfactionColMap_();

    // Find questions missing from the sheet
    var missing = (activeQuestions || []).filter(function(q) {
      return q && q.id && !colMap[q.id];
    });

    if (missing.length === 0) return colMap;

    // Append missing question ID headers
    var nextCol = satSheet.getLastColumn() + 1;
    missing.forEach(function(q) {
      satSheet.getRange(1, nextCol)
        .setValue(q.id)
        .setFontWeight('bold')
        .setBackground('#1a73e8')
        .setFontColor('#ffffff');
      satSheet.setColumnWidth(nextCol, 55);
      colMap[q.id] = nextCol;
      nextCol++;
    });

    // Invalidate cache so next call re-reads the updated headers
    try { CacheService.getScriptCache().remove('satisfactionColMap_v1'); } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }

    Logger.log('syncSatisfactionSheetColumns_: Added ' + missing.length + ' column(s): ' + missing.map(function(q){return q.id;}).join(', '));
    return colMap;

  } catch(e) {
    Logger.log('syncSatisfactionSheetColumns_ error: ' + e.message);
    return getSatisfactionColMap_();
  }
}

/**
 * Submits an in-app satisfaction survey response.
 *
 * ANONYMITY MODEL:
 *   - email resolved server-side from GAS Session — never sent from client
 *   - email hashed with hashForVault_() — raw value never persisted
 *   - Satisfaction sheet row: Timestamp + Q1-Q67 scores/text only (zero PII)
 *   - _Survey_Vault: hashed email + hashed member ID + period ID
 *   - _Survey_Tracking: completion status only (no answers)
 *
 * @param {string} callerEmail - Raw email from _resolveCallerEmail() (server-side only)
 * @param {Object} responses   - { q1: val, q2: val, ... q67: val }
 * @returns {Object} { success: bool, message: string }
 */
/**
 * Submits an in-app satisfaction survey response.
 *
 * ANONYMITY MODEL:
 *   - email resolved server-side from GAS Session — never sent from client
 *   - email hashed with hashForVault_() — raw value never persisted
 *   - Satisfaction sheet row: Timestamp + question ID columns only (zero PII)
 *   - _Survey_Vault: hashed email + hashed member ID + period ID
 *   - _Survey_Tracking: completion status only (no answers)
 *
 * v4.23.0: Row built dynamically via getSatisfactionColMap_() — no hardcoded positions.
 *
 * @param {string} callerEmail - Raw email from _resolveCallerEmail() (server-side only)
 * @param {Object} responses   - { q1: val, q2: val, ... qN: val }
 * @returns {Object} { success: bool, message: string }
 */
function submitSurveyResponse(callerEmail, responses) {
  if (!callerEmail || !responses) return { success: false, message: 'Not authenticated.' };

  return withScriptLock_(function() {
    try {
      var ss       = SpreadsheetApp.getActiveSpreadsheet();
      var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
      if (!satSheet) return { success: false, message: 'Survey sheet not found.' };

      // ── Period check ─────────────────────────────────────────────────
      var period = getSurveyPeriod();
      if (!period || period.status !== 'Active') {
        return { success: false, message: 'No survey is currently open. Check back next quarter.' };
      }

      // ── Hash email in-memory ─────────────────────────────────────────
      var emailHash = hashForVault_(callerEmail);
      var periodId  = period.periodId;

      // ── Vault duplicate check (period-scoped) ─────────────────────────
      var vaultSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT);
      if (vaultSheet && vaultSheet.getLastRow() > 1) {
        var vaultData = vaultSheet.getRange(
          2, 1, vaultSheet.getLastRow() - 1, vaultSheet.getLastColumn()
        ).getValues();
        for (var v = 0; v < vaultData.length; v++) {
          var vHash   = String(vaultData[v][SURVEY_VAULT_COLS.EMAIL   - 1] || '');
          var vPeriod = String(vaultData[v][SURVEY_VAULT_COLS.QUARTER - 1] || '');
          var vLatest = String(vaultData[v][SURVEY_VAULT_COLS.IS_LATEST - 1] || '');
          if (vHash === emailHash && vPeriod === periodId && vLatest !== 'Superseded') {
            return { success: false, message: 'You have already submitted a response for this survey period.' };
          }
        }
      }

      // ── Get active questions + sync columns ───────────────────────────
      var questionsData = getSurveyQuestions();
      var allQuestions  = questionsData.questions || [];
      var activeQs      = allQuestions.filter(function(q) { return q.active; });
      var colMap        = syncSatisfactionSheetColumns_(activeQs);

      // ── Value sanitizers ─────────────────────────────────────────────
      function _num(val) {
        var n = parseInt(val, 10);
        if (isNaN(n)) return '';
        return Math.min(10, Math.max(1, n));
      }
      function _txt(val) {
        var s = String(val || '').trim();
        if (s.charAt(0) === '=') s = "'" + s; // formula injection guard
        return s.substring(0, 3000);
      }
      function _csv(val) {
        if (Array.isArray(val)) return _txt(val.join(', '));
        return _txt(val);
      }

      // ── Build response row (sparse array, sized to last column) ────────
      var totalCols = satSheet.getLastColumn();
      var row = new Array(totalCols).fill('');

      // Fixed prefix
      var tsCol = colMap['Timestamp'];
      var piCol = colMap['Period ID'];
      var svCol = colMap['Survey Version'];
      if (tsCol) row[tsCol - 1] = new Date();
      if (piCol) row[piCol - 1] = periodId;
      if (svCol) row[svCol - 1] = periodId + '-v' + questionsData.questions.length;

      // Dynamic question columns
      activeQs.forEach(function(q) {
        var col = colMap[q.id];
        if (!col || responses[q.id] === undefined) return;
        var type = q.type;
        var val;
        if (type === 'slider-10') {
          val = _num(responses[q.id]);
        } else if (type === 'checkbox') {
          val = _csv(responses[q.id]);
        } else {
          val = _txt(responses[q.id]);
        }
        row[col - 1] = val;
      });

      satSheet.appendRow(row);
      var newResponseRow = satSheet.getLastRow();

      // ── Write vault entry (hashed PII only) ──────────────────────────
      if (vaultSheet) {
        var memberId = '';
        try {
          var memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
          if (memberSheet && memberSheet.getLastRow() > 1) {
            var memberEmails = memberSheet.getRange(
              2, MEMBER_COLS.EMAIL, memberSheet.getLastRow() - 1, 1
            ).getValues();
            var callerLc = callerEmail.toLowerCase().trim();
            for (var m = 0; m < memberEmails.length; m++) {
              if (String(memberEmails[m][0]).toLowerCase().trim() === callerLc) {
                var memberIdVal = memberSheet.getRange(m + 2, MEMBER_COLS.MEMBER_ID).getValue();
                if (memberIdVal) memberId = String(memberIdVal);
                break;
              }
            }
          }
        } catch (_ex) { Logger.log('_ex: ' + (_ex.message || _ex)); }

        var memberIdHash = memberId ? hashForVault_(memberId) : '';
        var vaultRow = new Array(8).fill('');
        vaultRow[SURVEY_VAULT_COLS.RESPONSE_ROW     - 1] = newResponseRow;
        vaultRow[SURVEY_VAULT_COLS.EMAIL            - 1] = emailHash;
        vaultRow[SURVEY_VAULT_COLS.VERIFIED         - 1] = 'Yes';
        vaultRow[SURVEY_VAULT_COLS.MATCHED_MEMBER_ID - 1] = memberIdHash;
        vaultRow[SURVEY_VAULT_COLS.QUARTER          - 1] = periodId;
        vaultRow[SURVEY_VAULT_COLS.IS_LATEST        - 1] = 'Latest';
        vaultRow[SURVEY_VAULT_COLS.SUPERSEDED_BY    - 1] = '';
        vaultRow[SURVEY_VAULT_COLS.REVIEWER_NOTES   - 1] = '';
        vaultSheet.appendRow(vaultRow);
      }

      // ── Update _Survey_Tracking completion status ──────────────────────
      var trackSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
      if (trackSheet && trackSheet.getLastRow() > 1) {
        var trackData = trackSheet.getDataRange().getValues();
        var callerNorm = callerEmail.toLowerCase().trim();
        for (var t = 1; t < trackData.length; t++) {
          var tEmail = String(trackData[t][SURVEY_TRACKING_COLS.EMAIL - 1] || '').toLowerCase().trim();
          if (tEmail === callerNorm) {
            var tLastCol = trackSheet.getLastColumn();
            var tRow = trackSheet.getRange(t + 1, 1, 1, tLastCol).getValues()[0];
            tRow[SURVEY_TRACKING_COLS.CURRENT_STATUS  - 1] = 'Completed';
            tRow[SURVEY_TRACKING_COLS.COMPLETED_DATE  - 1] = new Date();
            tRow[SURVEY_TRACKING_COLS.TOTAL_COMPLETED - 1] = (parseInt(tRow[SURVEY_TRACKING_COLS.TOTAL_COMPLETED - 1], 10) || 0) + 1;
            trackSheet.getRange(t + 1, 1, 1, tLastCol).setValues([tRow]);
            break;
          }
        }
      }

      // ── Update period response count + invalidate summary cache ────────
      try { incrementPeriodResponseCount_(periodId); } catch (_ex) { Logger.log('_ex: ' + (_ex.message || _ex)); }
      try { CacheService.getScriptCache().remove('satisfactionSummary_' + periodId); } catch (_ex) { Logger.log('_ex: ' + (_ex.message || _ex)); }

      return { success: true, message: 'Thank you — your anonymous response has been recorded.' };

    } catch(e) {
      Logger.log('submitSurveyResponse error: ' + e.message + '\n' + (e.stack || ''));
      return { success: false, message: 'Error submitting survey. Please try again.' };
    }
  }, 30);
}
// ============================================================================
// SATISFACTION COLS COMPATIBILITY SHIM — v4.23.1
// ============================================================================

/**
 * Returns a SATISFACTION_COLS-compatible object backed by the live Satisfaction
 * sheet header→col map. Any function that reads SATISFACTION_COLS positionally
 * should shadow the global at function entry:
 *
 *   var SATISFACTION_COLS = buildSatisfactionColsShim_(getSatisfactionColMap_());
 *
 * Q-keys map to their dynamic column positions (1-indexed).
 * AVG_* and SUMMARY_START return 0 — these are no longer pre-computed in the sheet.
 *   Callers that guard with `if (summaryStart > 0)` are safe; callers that write
 *   to `sheet.getRange(..., summaryStart)` must add that guard.
 *
 * @param {Object} colMap - from getSatisfactionColMap_()
 * @returns {Object} SATISFACTION_COLS-compatible map
 */
function buildSatisfactionColsShim_(colMap) {
  colMap = colMap || {};
  function c(key) { return colMap[key] || 0; }
  return {
    // Fixed prefix (v4.23.0)
    TIMESTAMP:      c('Timestamp'),
    // Question columns — keyed by question ID
    Q1_WORKSITE:               c('q1'),
    Q2_ROLE:                   c('q2'),
    Q3_SHIFT:                  c('q3'),
    Q4_TIME_IN_ROLE:           c('q4'),
    Q5_STEWARD_CONTACT:        c('q5'),
    Q6_SATISFIED_REP:          c('q6'),
    Q7_TRUST_UNION:            c('q7'),
    Q8_FEEL_PROTECTED:         c('q8'),
    Q9_RECOMMEND:              c('q9'),
    Q10_TIMELY_RESPONSE:       c('q10'),
    Q11_TREATED_RESPECT:       c('q11'),
    Q12_EXPLAINED_OPTIONS:     c('q12'),
    Q13_FOLLOWED_THROUGH:      c('q13'),
    Q14_ADVOCATED:             c('q14'),
    Q15_SAFE_CONCERNS:         c('q15'),
    Q16_CONFIDENTIALITY:       c('q16'),
    Q17_STEWARD_IMPROVE:       c('q17'),
    Q18_KNOW_CONTACT:          c('q18'),
    Q19_CONFIDENT_HELP:        c('q19'),
    Q20_EASY_FIND:             c('q20'),
    Q21_UNDERSTAND_ISSUES:     c('q21'),
    Q22_CHAPTER_COMM:          c('q22'),
    Q23_ORGANIZES:             c('q23'),
    Q24_REACH_CHAPTER:         c('q24'),
    Q25_FAIR_REP:              c('q25'),
    Q26_DECISIONS_CLEAR:       c('q26'),
    Q27_UNDERSTAND_PROCESS:    c('q27'),
    Q28_TRANSPARENT_FINANCE:   c('q28'),
    Q29_ACCOUNTABLE:           c('q29'),
    Q30_FAIR_PROCESSES:        c('q30'),
    Q31_WELCOMES_OPINIONS:     c('q31'),
    Q32_ENFORCES_CONTRACT:     c('q32'),
    Q33_REALISTIC_TIMELINES:   c('q33'),
    Q34_CLEAR_UPDATES:         c('q34'),
    Q35_FRONTLINE_PRIORITY:    c('q35'),
    Q36_FILED_GRIEVANCE:       c('q36'),
    Q37_UNDERSTOOD_STEPS:      c('q37'),
    Q38_FELT_SUPPORTED:        c('q38'),
    Q39_UPDATES_OFTEN:         c('q39'),
    Q40_OUTCOME_JUSTIFIED:     c('q40'),
    Q41_CLEAR_ACTIONABLE:      c('q41'),
    Q42_ENOUGH_INFO:           c('q42'),
    Q43_FIND_EASILY:           c('q43'),
    Q44_ALL_SHIFTS:            c('q44'),
    Q45_MEETINGS_WORTH:        c('q45'),
    Q46_VOICE_MATTERS:         c('q46'),
    Q47_SEEKS_INPUT:           c('q47'),
    Q48_DIGNITY:               c('q48'),
    Q49_NEWER_SUPPORTED:       c('q49'),
    Q50_CONFLICT_RESPECT:      c('q50'),
    Q51_GOOD_VALUE:            c('q51'),
    Q52_PRIORITIES_NEEDS:      c('q52'),
    Q53_PREPARED_MOBILIZE:     c('q53'),
    Q54_HOW_INVOLVED:          c('q54'),
    Q55_WIN_TOGETHER:          c('q55'),
    Q56_UNDERSTAND_CHANGES:    c('q56'),
    Q57_ADEQUATELY_INFORMED:   c('q57'),
    Q58_CLEAR_CRITERIA:        c('q58'),
    Q59_WORK_EXPECTATIONS:     c('q59'),
    Q60_EFFECTIVE_OUTCOMES:    c('q60'),
    Q61_SUPPORTS_WELLBEING:    c('q61'),
    Q62_CONCERNS_SERIOUS:      c('q62'),
    Q63_SCHEDULING_CHALLENGE:  c('q63'),
    Q64_TOP_PRIORITIES:        c('q64'),
    Q65_ONE_CHANGE:            c('q65'),
    Q66_KEEP_DOING:            c('q66'),
    Q67_ADDITIONAL:            c('q67'),
    // Summary/AVG columns — removed in v4.23.0 dynamic schema.
    // All return 0. Callers must guard: if (summaryStart > 0) { ... }
    // Use getSatisfactionSummary() for computed section averages.
    SUMMARY_START:         0,
    AVG_OVERALL_SAT:       0,
    AVG_STEWARD_RATING:    0,
    AVG_STEWARD_ACCESS:    0,
    AVG_CHAPTER:           0,
    AVG_LEADERSHIP:        0,
    AVG_CONTRACT:          0,
    AVG_REPRESENTATION:    0,
    AVG_COMMUNICATION:     0,
    AVG_MEMBER_VOICE:      0,
    AVG_VALUE_ACTION:      0,
    AVG_SCHEDULING:        0
  };
}
