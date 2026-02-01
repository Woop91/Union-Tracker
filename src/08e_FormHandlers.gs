/**
 * ============================================================================
 * 08e_FormHandlers.gs - Form Handlers Module
 * ============================================================================
 *
 * This module contains all form-related functions for the 509 Steward Dashboard:
 * - Form submission handlers (Contact, Satisfaction, Grievance)
 * - Form trigger setup functions
 * - Form URL handling and configuration
 * - Form value parsing utilities
 *
 * Dependencies:
 * - SHEETS, CONFIG_COLS constants from 00_Constants.gs
 * - GRIEVANCE_FORM_CONFIG, CONTACT_FORM_CONFIG, SATISFACTION_FORM_CONFIG
 * - MEMBER_COLS, SATISFACTION_COLS from column configuration
 * - Helper functions: findExistingMember, generateNameBasedId, validateMemberEmail
 * - Helper functions: getCurrentQuarter, computeSatisfactionRowAverages, syncSatisfactionValues
 *
 * Note: onGrievanceFormSubmit() is defined in 05_Integrations.gs
 *
 * @author SEIU Local 509 Development Team
 * @version 1.0.0
 */

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
 * Save form URLs to the Config tab for easy reference and updating
 * Writes Grievance Form, Contact Form, and Satisfaction Survey URLs to Config columns P, Q, AR
 */
function saveFormUrlsToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  saveFormUrlsToConfig_silent(ss);
  ss.toast('Form URLs saved to Config tab (columns P, Q, AR)', 'Saved', 3);
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
  } catch (e) {
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
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
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
    ui.alert('Trigger Exists',
      'A contact form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Contact Form Trigger',
    'This will set up automatic processing of contact info form submissions.\n\n' +
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
      ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
      return;
    }
    var formId = match[1];

    // Open the form and create trigger
    var form = FormApp.openById(formId);

    ScriptApp.newTrigger('onContactFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    ui.alert('Trigger Created',
      'Contact form trigger has been set up!\n\n' +
      'When a contact form is submitted:\n' +
      '- The member\'s record will be updated in Member Directory\n' +
      '- Contact info, preferences, and interests will be saved',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// GRIEVANCE FORM TRIGGER SETUP
// ============================================================================

/**
 * Set up the grievance form submission trigger
 * Run this once to enable automatic processing of form submissions
 *
 * Note: The actual handler onGrievanceFormSubmit() is defined in 05_Integrations.gs
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
    ui.alert('Trigger Exists',
      'A grievance form trigger already exists.\n\n' +
      'Form submissions will be automatically processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Prompt for form URL
  var response = ui.prompt('Setup Grievance Form Trigger',
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
        ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
        return;
      }
      formId = match[1];
    } else {
      // Use configured form
      var configFormUrl = GRIEVANCE_FORM_CONFIG.FORM_URL;
      var match = configFormUrl.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        ui.alert('No Form Configured',
          'No form URL provided and could not extract ID from config.\n\n' +
          'Please provide the form edit URL.',
          ui.ButtonSet.OK);
        return;
      }
      // Note: The /e/ URL is the published version, we need the actual form ID
      ui.alert('Form URL Needed',
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

    ui.alert('Trigger Created',
      'Grievance form trigger has been set up!\n\n' +
      'When a grievance form is submitted:\n' +
      '- A new row will be added to Grievance Log\n' +
      '- A Drive folder will be created automatically\n' +
      '- Deadlines will be calculated\n' +
      '- Member Directory will be updated',
      ui.ButtonSet.OK);

    ss.toast('Form trigger created successfully!', 'Success', 3);

  } catch (e) {
    ui.alert('Error',
      'Failed to create trigger: ' + e.message + '\n\n' +
      'Make sure you have edit access to the form.',
      ui.ButtonSet.OK);
  }
}

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
      '<button onclick="copyLink()" style="padding: 8px 16px; cursor: pointer;">Copy to Clipboard</button>' +
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
    ui.showModalDialog(copyHtml, 'Survey Link');
  }
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

    // ========================================================================
    // EMAIL VERIFICATION & QUARTERLY TRACKING
    // ========================================================================

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
      ui.alert('Invalid URL', 'Could not extract form ID from URL.', ui.ButtonSet.OK);
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
