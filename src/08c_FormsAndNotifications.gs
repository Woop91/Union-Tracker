
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

  // Config layout: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
  // Set column headers in row 2
  configSheet.getRange(2, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue('Grievance Form URL');
  configSheet.getRange(2, CONFIG_COLS.CONTACT_FORM_URL).setValue('Contact Form URL');
  configSheet.getRange(2, CONFIG_COLS.SATISFACTION_FORM_URL).setValue('Satisfaction Survey URL');

  // Set form URLs in row 3 (data row)
  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setValue(GRIEVANCE_FORM_CONFIG.FORM_URL);
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setValue(CONTACT_FORM_CONFIG.FORM_URL);
  configSheet.getRange(3, CONFIG_COLS.SATISFACTION_FORM_URL).setValue(SATISFACTION_FORM_CONFIG.FORM_URL);

  // Format as links
  configSheet.getRange(3, CONFIG_COLS.GRIEVANCE_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(3, CONFIG_COLS.CONTACT_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
  configSheet.getRange(3, CONFIG_COLS.SATISFACTION_FORM_URL).setFontColor('#1155cc').setFontLine('underline');
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
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
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
      // Log with masked name for privacy
      Logger.log('Created new member ' + memberId + ': ' + maskedName);

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

      // Mask name in logs for privacy
      var maskedUpdateName = typeof maskName === 'function' ? maskName(firstName + ' ' + lastName) : '[REDACTED]';
      Logger.log('Updated contact info for ' + maskedUpdateName + ' (row ' + memberRow + ')');
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
        ui.alert('Invalid URL',
        'Could not extract form ID from URL.\n\n' +
        'Please use the form\'s edit URL. It should look like:\n' +
        'https://docs.google.com/forms/d/YOUR_FORM_ID/edit',
        ui.ButtonSet.OK);
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
      '<html><head>' + getMobileOptimizedHead() + '</head><body>' +
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
      '</div></body></html>'
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
    newRow[SATISFACTION_COLS.Q9_RECOMMEND - 1] = getFormValue_(responses, 'Would recommend membership');

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
 * ============================================================================
 * 08h_NotificationEngine.gs - Notification and Alert System
 * ============================================================================
 *
 * This module handles all notification and alert functionality for the
 * SEIU Local 509 Dashboard including:
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
 * @author SEIU Local 509
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
  var email = props.getProperty('notification_email') || Session.getEffectiveUser().getEmail();

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

    if (daysToDeadline === 'Overdue' || (daysToDeadline !== '' && typeof daysToDeadline === 'number' && daysToDeadline <= 3)) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });
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

    var body = '509 GRIEVANCE DEADLINE ALERT\n';
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

/**
 * Manual trigger to send steward alerts now
 */
function sendStewardAlertsNow() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('Send Steward Alerts',
    'This will send deadline alert emails to all stewards with upcoming deadlines.\n\n' +
    'Each steward will receive their own personalized digest.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var emailsSent = sendStewardDeadlineAlerts();

  ui.alert('Alerts Sent',
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

  var response = ui.prompt('Alert Settings',
    'Current settings:\n' +
    '* Alert window: ' + currentDays + ' days before deadline\n' +
    '* Per-steward alerts: ' + (stewardAlerts ? 'ENABLED' : 'DISABLED') + '\n\n' +
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

  ui.alert('Settings Saved',
    'Alert window: ' + newDays + ' days\n' +
    'Per-steward alerts: ' + (stewardResponse === ui.Button.YES ? 'ENABLED' : 'DISABLED'),
    ui.ButtonSet.OK);
}

// ============================================================================
// SURVEY EMAIL DISTRIBUTION
// ============================================================================

/**
 * Show dialog for sending random survey emails to members
 */
function sendRandomSurveyEmails() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show configuration dialog
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
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
    '<h2>Send Survey to Random Members</h2>' +
    '<div class="info">Select how many random members to email. Each member will receive a personalized survey link.</div>' +
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
    '<button class="primary" onclick="send()">Send Surveys</button></div></div>' +
    '<script>function send(){var opts={count:parseInt(document.getElementById("count").value),' +
    'subject:document.getElementById("subject").value,excludeDays:parseInt(document.getElementById("excludeDays").value)};' +
    'google.script.run.withSuccessHandler(function(r){alert(r);google.script.host.close()})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message)}).executeSendRandomSurveyEmails(opts)}</script></body></html>'
  ).setWidth(500).setHeight(450);

  ui.showModalDialog(html, 'Send Random Survey Emails');
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

  // Update survey email log
  if (newLogEntries.length > 0) {
    var nextRow = Object.keys(surveyLog).length + 2;
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

