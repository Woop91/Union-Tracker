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

  // M-15: Config data starts at row 3 (row 1 = section headers, row 2 = column headers)
  if (configSheet) {
    var url = configSheet.getRange(3, configCol).getValue();
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
      match = configFormUrl.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
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
    // EMAIL VERIFICATION & VAULT STORAGE
    // Survey answers go to the Satisfaction sheet (anonymous).
    // Email + Member ID go to _Survey_Vault (protected, hidden).
    // No PII is ever written to the Satisfaction sheet.
    // ========================================================================

    // Get email from form (try multiple common field names)
    var email = getFormValue_(responses, 'Email Address') ||
                getFormValue_(responses, 'Email') ||
                getFormValue_(responses, 'email') ||
                (e.response ? e.response.getRespondentEmail() : '') || '';
    email = email.toString().toLowerCase().trim();

    // NOTE: email is NOT written to newRow — it goes to the vault only
    var currentQuarter = getCurrentQuarter();
    var memberMatch = validateMemberEmail(email);
    var verified = memberMatch ? 'Yes' : 'Pending Review';
    var memberId = memberMatch ? memberMatch.memberId : '';

    // ── Sanitize all user-supplied values to prevent formula injection ──
    for (var fi = 0; fi < newRow.length; fi++) {
      if (typeof newRow[fi] === 'string') {
        newRow[fi] = escapeForFormula(newRow[fi]);
      }
    }

    // ── Append ANONYMOUS response to Satisfaction sheet ──
    // Use a lock around appendRow + getLastRow to prevent race conditions
    // when two form submissions arrive simultaneously (F34d fix)
    var formLock = LockService.getScriptLock();
    formLock.waitLock(10000);
    try {
      satSheet.appendRow(newRow);
      SpreadsheetApp.flush(); // ensure row is committed before reading back
      var newRowNum = satSheet.getLastRow();
    } finally {
      formLock.releaseLock();
    }

    // ── Write hashed PII to vault (separate protected sheet) ──
    // Supersede any previous entry from same email+quarter
    if (memberMatch) {
      supersedePreviousVaultEntry_(email, currentQuarter, newRowNum);
    }
    writeVaultEntry_(newRowNum, email, verified, memberId, currentQuarter);

    // Compute section averages for the new row
    computeSatisfactionRowAverages(newRowNum);

    // Update dashboard summary values
    syncSatisfactionValues();

    // Update survey completion tracking if member was matched.
    // This marks the member as "Completed" in _Survey_Tracking.
    if (memberMatch && memberMatch.memberId) {
      try {
        updateSurveyTrackingOnSubmit_(memberMatch.memberId);
      } catch (trackErr) {
        Logger.log('Survey tracking update error: ' + trackErr.message);
      }
    }

    Logger.log('Satisfaction survey response recorded at ' + new Date() + ' | Verified: ' + verified + ' | Hashed PII stored in vault');

    // Send thank-you email if respondent email available
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + 'Thank You for Your Feedback',
          body: 'Thank you for completing the member satisfaction survey.\n\n' +
            'Your feedback helps us improve our representation and services.\n\n' +
            'Best regards,\nWFSE Local'
        });
      } catch (emailError) {
        Logger.log('Could not send survey thank-you email: ' + emailError.message);
      }
    }

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

      // Emit EventBus notification for SPA bell alerts
      if (typeof EventBus !== 'undefined' && EventBus.emit) {
        EventBus.emit('grievance:deadline:approaching', {
          grievanceId: grievanceId,
          memberName: String(data[i][GRIEVANCE_COLS.MEMBER_ID - 1] || ''),
          daysLeft: daysToDeadline === 'Overdue' ? 0 : daysToDeadline
        });
      }
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

    // Emit EventBus notification for SPA bell alerts
    if (typeof EventBus !== 'undefined' && EventBus.emit) {
      EventBus.emit('grievance:deadline:approaching', {
        grievanceId: grievanceId,
        memberName: memberInfo.name,
        daysLeft: daysRemaining
      });
    }
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
  var orgName = typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Union') : 'Union';

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
    '<input type="text" id="subject" value="' + escapeHtml(orgName) + ' - Member Satisfaction Survey"></div>' +
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
  } catch(_e) { /* No log yet */ }

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

/**
 * Get quarter string from a date
 * @param {Date} date - Date to get quarter from
 * @returns {string} Quarter string
 */
function getQuarterFromDate(date) {
  var d = new Date(date);
  if (isNaN(d.getTime())) return '';
  var quarter = Math.floor(d.getMonth() / 3) + 1;
  return d.getFullYear() + '-Q' + quarter;
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
    } catch (_e) { /* headless */ }
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
  } catch (_e) { /* headless */ }
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

  if (!trackingSheet || trackingSheet.getLastRow() < 2) {
    Logger.log('sendSurveyCompletionReminders: No tracking data');
    return 'No survey tracking data found.';
  }

  // Get survey URL from Config using CONFIG_COLS constant (data in row 3)
  var surveyUrl = '';
  if (configSheet && configSheet.getLastRow() >= 3) {
    try {
      surveyUrl = configSheet.getRange(3, CONFIG_COLS.SATISFACTION_FORM_URL).getValue() || '';
    } catch (_e) { /* no survey URL configured */ }
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
function getSurveyQuestions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);

    // ── Dynamic options from Config ──────────────────────────────────────
    function _getConfigList(colConst) {
      if (!configSheet || !colConst || colConst < 1) return [];
      try {
        var lastRow = configSheet.getLastRow();
        if (lastRow < 3) return [];
        var vals = configSheet.getRange(3, colConst, lastRow - 2, 1).getValues();
        return vals.map(function(r) { return String(r[0]).trim(); })
                   .filter(function(v) { return v !== ''; });
      } catch(e) { return []; }
    }

    var worksites       = _getConfigList(CONFIG_COLS.OFFICE_LOCATIONS);
    var roles           = _getConfigList(CONFIG_COLS.JOB_TITLES);
    var priorityOptions = _getConfigList(CONFIG_COLS.SURVEY_PRIORITY_OPTIONS);

    // Fallback defaults if Config rows not yet populated
    if (!worksites.length)       worksites       = ['Please add worksites to Config tab'];
    if (!roles.length)           roles           = ['Please add roles to Config tab'];
    if (!priorityOptions.length) priorityOptions = [
      'Contract Enforcement','Workload','Scheduling','Pay & Benefits',
      'Safety','Training','Equity & Inclusion','Communication',
      'Steward Support','Organizing','Other'
    ];

    var shiftOptions   = ['Day','Evening','Night','Rotating/Variable'];
    var tenureOptions  = ['Less than 1 year','1–3 years','4–7 years','8–15 years','15+ years'];
    var yesNo          = ['Yes','No'];
    var SL = 'slider-10';
    var PARA = 'paragraph';
    var sl  = { labelMin: 'Strongly Disagree', labelMax: 'Strongly Agree' };

    // ── Question definitions ─────────────────────────────────────────────
    // col = SATISFACTION_COLS value (1-indexed sheet column)
    var questions = [
      // ── Section 1: Work Context ────────────────────────────────────────
      { id:'q1',  col: SATISFACTION_COLS.Q1_WORKSITE,        text:'What is your worksite / program / region?',          type:'dropdown',    section:'1', sectionKey:'WORK_CONTEXT',   required:true,  options: worksites },
      { id:'q2',  col: SATISFACTION_COLS.Q2_ROLE,            text:'What is your role / job group?',                    type:'dropdown',    section:'1', sectionKey:'WORK_CONTEXT',   required:true,  options: roles },
      { id:'q3',  col: SATISFACTION_COLS.Q3_SHIFT,           text:'What shift do you work?',                           type:'radio',       section:'1', sectionKey:'WORK_CONTEXT',   required:true,  options: shiftOptions },
      { id:'q4',  col: SATISFACTION_COLS.Q4_TIME_IN_ROLE,    text:'How long have you been in your current role?',      type:'radio',       section:'1', sectionKey:'WORK_CONTEXT',   required:true,  options: tenureOptions },
      { id:'q5',  col: SATISFACTION_COLS.Q5_STEWARD_CONTACT, text:'Have you had contact with a steward in the past 12 months?', type:'radio-branch', section:'1', sectionKey:'WORK_CONTEXT', required:true,  options: yesNo,
        branchYes:'3A', branchNo:'3B' },

      // ── Section 2: Overall Satisfaction ───────────────────────────────
      { id:'q6',  col: SATISFACTION_COLS.Q6_SATISFIED_REP,   text:'I am satisfied with my union representation.',          type:SL, section:'2', sectionKey:'OVERALL_SAT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q7',  col: SATISFACTION_COLS.Q7_TRUST_UNION,     text:'I trust the union to act in members\' best interests.', type:SL, section:'2', sectionKey:'OVERALL_SAT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q8',  col: SATISFACTION_COLS.Q8_FEEL_PROTECTED,  text:'I feel more protected at work because of my union.',    type:SL, section:'2', sectionKey:'OVERALL_SAT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q9',  col: SATISFACTION_COLS.Q9_RECOMMEND,       text:'I would recommend union membership to a coworker.',     type:SL, section:'2', sectionKey:'OVERALL_SAT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 3A: Steward Ratings (Q5=Yes only) ─────────────────────
      { id:'q10', col: SATISFACTION_COLS.Q10_TIMELY_RESPONSE, text:'My steward responded to me in a timely manner.',            type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q11', col: SATISFACTION_COLS.Q11_TREATED_RESPECT, text:'My steward treated me with respect.',                       type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q12', col: SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS,'text':'My steward explained my options clearly.',               type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q13', col: SATISFACTION_COLS.Q13_FOLLOWED_THROUGH, text:'My steward followed through on their commitments.',        type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q14', col: SATISFACTION_COLS.Q14_ADVOCATED,        text:'My steward advocated effectively on my behalf.',           type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q15', col: SATISFACTION_COLS.Q15_SAFE_CONCERNS,    text:'I felt safe raising concerns with my steward.',            type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q16', col: SATISFACTION_COLS.Q16_CONFIDENTIALITY,  text:'My steward handled confidentiality appropriately.',        type:SL, section:'3A', sectionKey:'STEWARD_3A', required:true,  branchParent:'q5', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q17', col: SATISFACTION_COLS.Q17_STEWARD_IMPROVE,  text:'What should stewards do to improve? (optional)',          type:PARA, section:'3A', sectionKey:'STEWARD_3A', required:false, branchParent:'q5', branchValue:'Yes' },

      // ── Section 3B: Steward Access (Q5=No only) ────────────────────────
      { id:'q18', col: SATISFACTION_COLS.Q18_KNOW_CONTACT, text:'I know how to contact a steward or union rep.',   type:SL, section:'3B', sectionKey:'STEWARD_3B', required:true,  branchParent:'q5', branchValue:'No', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q19', col: SATISFACTION_COLS.Q19_CONFIDENT_HELP,'text':'I am confident I would get help if I needed it.', type:SL, section:'3B', sectionKey:'STEWARD_3B', required:true,  branchParent:'q5', branchValue:'No', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q20', col: SATISFACTION_COLS.Q20_EASY_FIND,    text:'It is easy to figure out who to contact.',       type:SL, section:'3B', sectionKey:'STEWARD_3B', required:true,  branchParent:'q5', branchValue:'No', labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 4: Chapter Effectiveness ──────────────────────────────
      { id:'q21', col: SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, text:'Union reps understand my workplace issues.',              type:SL, section:'4', sectionKey:'CHAPTER', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q22', col: SATISFACTION_COLS.Q22_CHAPTER_COMM,      text:'Chapter communication is regular and clear.',            type:SL, section:'4', sectionKey:'CHAPTER', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q23', col: SATISFACTION_COLS.Q23_ORGANIZES,         text:'The chapter organizes members effectively.',             type:SL, section:'4', sectionKey:'CHAPTER', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q24', col: SATISFACTION_COLS.Q24_REACH_CHAPTER,     text:'I know how to reach my chapter contact.',                type:SL, section:'4', sectionKey:'CHAPTER', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q25', col: SATISFACTION_COLS.Q25_FAIR_REP,          text:'Representation is fair across roles and shifts.',        type:SL, section:'4', sectionKey:'CHAPTER', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 5: Local Leadership ────────────────────────────────────
      { id:'q26', col: SATISFACTION_COLS.Q26_DECISIONS_CLEAR,    text:'Leadership communicates decisions clearly.',            type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q27', col: SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS, text:'I understand how decisions are made.',                 type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q28', col: SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE,'text':'The union is transparent about its finances.',        type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q29', col: SATISFACTION_COLS.Q29_ACCOUNTABLE,        text:'Leadership is accountable to member feedback.',         type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q30', col: SATISFACTION_COLS.Q30_FAIR_PROCESSES,     text:'Internal union processes feel fair.',                   type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q31', col: SATISFACTION_COLS.Q31_WELCOMES_OPINIONS,  text:'The union welcomes differing opinions.',                type:SL, section:'5', sectionKey:'LEADERSHIP', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 6: Contract Enforcement ───────────────────────────────
      { id:'q32', col: SATISFACTION_COLS.Q32_ENFORCES_CONTRACT,   text:'The union enforces our contract effectively.',         type:SL, section:'6', sectionKey:'CONTRACT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q33', col: SATISFACTION_COLS.Q33_REALISTIC_TIMELINES, text:'The union communicates realistic timelines.',          type:SL, section:'6', sectionKey:'CONTRACT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q34', col: SATISFACTION_COLS.Q34_CLEAR_UPDATES,       text:'The union provides clear updates on issues.',          type:SL, section:'6', sectionKey:'CONTRACT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q35', col: SATISFACTION_COLS.Q35_FRONTLINE_PRIORITY,  text:'The union prioritizes frontline working conditions.',  type:SL, section:'6', sectionKey:'CONTRACT', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q36', col: SATISFACTION_COLS.Q36_FILED_GRIEVANCE,  text:'Have you filed a grievance in the past 24 months?', type:'radio-branch', section:'6', sectionKey:'CONTRACT', required:true,  options: yesNo,
        branchYes:'6A', branchNo:'7' },

      // ── Section 6A: Representation Process (Q36=Yes only) ─────────────
      { id:'q37', col: SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS, text:'I understood the steps and timeline of my grievance.',    type:SL, section:'6A', sectionKey:'REPRESENTATION', required:true,  branchParent:'q36', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q38', col: SATISFACTION_COLS.Q38_FELT_SUPPORTED,   text:'I felt supported throughout the grievance process.',     type:SL, section:'6A', sectionKey:'REPRESENTATION', required:true,  branchParent:'q36', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q39', col: SATISFACTION_COLS.Q39_UPDATES_OFTEN,    text:'I received updates often enough during my case.',        type:SL, section:'6A', sectionKey:'REPRESENTATION', required:true,  branchParent:'q36', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q40', col: SATISFACTION_COLS.Q40_OUTCOME_JUSTIFIED, text:'The outcome of my grievance feels justified.',          type:SL, section:'6A', sectionKey:'REPRESENTATION', required:true,  branchParent:'q36', branchValue:'Yes', labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 7: Communication Quality ──────────────────────────────
      { id:'q41', col: SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, text:'Union communications are clear and actionable.',           type:SL, section:'7', sectionKey:'COMMUNICATION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q42', col: SATISFACTION_COLS.Q42_ENOUGH_INFO,      text:'I receive enough information from the union.',             type:SL, section:'7', sectionKey:'COMMUNICATION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q43', col: SATISFACTION_COLS.Q43_FIND_EASILY,      text:'I can find information from the union easily.',            type:SL, section:'7', sectionKey:'COMMUNICATION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q44', col: SATISFACTION_COLS.Q44_ALL_SHIFTS,       text:'Communications reach members on all shifts and locations.',type:SL, section:'7', sectionKey:'COMMUNICATION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q45', col: SATISFACTION_COLS.Q45_MEETINGS_WORTH,   text:'Union meetings are worth attending.',                      type:SL, section:'7', sectionKey:'COMMUNICATION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 8: Member Voice & Culture ─────────────────────────────
      { id:'q46', col: SATISFACTION_COLS.Q46_VOICE_MATTERS, text:'My voice matters in this union.',                         type:SL, section:'8', sectionKey:'MEMBER_VOICE', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q47', col: SATISFACTION_COLS.Q47_SEEKS_INPUT,   text:'The union actively seeks member input.',                  type:SL, section:'8', sectionKey:'MEMBER_VOICE', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q48', col: SATISFACTION_COLS.Q48_DIGNITY,       text:'Members are treated with dignity by union leadership.',   type:SL, section:'8', sectionKey:'MEMBER_VOICE', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q49', col: SATISFACTION_COLS.Q49_NEWER_SUPPORTED,'text':'Newer members are well-supported.',                    type:SL, section:'8', sectionKey:'MEMBER_VOICE', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q50', col: SATISFACTION_COLS.Q50_CONFLICT_RESPECT,'text':'Internal conflicts are handled respectfully.',        type:SL, section:'8', sectionKey:'MEMBER_VOICE', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 9: Value & Collective Action ──────────────────────────
      { id:'q51', col: SATISFACTION_COLS.Q51_GOOD_VALUE,       text:'Union membership provides good value for my dues.',  type:SL, section:'9', sectionKey:'VALUE_ACTION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q52', col: SATISFACTION_COLS.Q52_PRIORITIES_NEEDS, text:'The union\'s priorities reflect member needs.',       type:SL, section:'9', sectionKey:'VALUE_ACTION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q53', col: SATISFACTION_COLS.Q53_PREPARED_MOBILIZE,'text':'The union is prepared to mobilize when needed.',    type:SL, section:'9', sectionKey:'VALUE_ACTION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q54', col: SATISFACTION_COLS.Q54_HOW_INVOLVED,     text:'I understand how to get more involved.',             type:SL, section:'9', sectionKey:'VALUE_ACTION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q55', col: SATISFACTION_COLS.Q55_WIN_TOGETHER,     text:'Acting together, we can win real improvements.',     type:SL, section:'9', sectionKey:'VALUE_ACTION', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },

      // ── Section 10: Scheduling / Office Days ──────────────────────────
      { id:'q56', col: SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES, text:'I understand proposed scheduling/office day changes.', type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q57', col: SATISFACTION_COLS.Q57_ADEQUATELY_INFORMED,'text':'I am adequately informed about scheduling decisions.', type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q58', col: SATISFACTION_COLS.Q58_CLEAR_CRITERIA,     text:'Scheduling decisions use clear and fair criteria.',   type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q59', col: SATISFACTION_COLS.Q59_WORK_EXPECTATIONS,  text:'My work can reasonably be done under current expectations.', type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q60', col: SATISFACTION_COLS.Q60_EFFECTIVE_OUTCOMES, text:'The current scheduling approach supports effective outcomes.', type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q61', col: SATISFACTION_COLS.Q61_SUPPORTS_WELLBEING, text:'The scheduling approach supports my wellbeing.',       type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q62', col: SATISFACTION_COLS.Q62_CONCERNS_SERIOUS,   text:'My scheduling concerns would be taken seriously.',    type:SL, section:'10', sectionKey:'SCHEDULING', required:true, labelMin:sl.labelMin, labelMax:sl.labelMax },
      { id:'q63', col: SATISFACTION_COLS.Q63_SCHEDULING_CHALLENGE,'text':'What is your biggest scheduling challenge? (optional)', type:PARA, section:'10', sectionKey:'SCHEDULING', required:false },

      // ── Section 11: Priorities & Close ────────────────────────────────
      { id:'q64', col: SATISFACTION_COLS.Q64_TOP_PRIORITIES, text:'Select your top 3 union priorities for the next 6–12 months.', type:'checkbox', section:'11', sectionKey:'PRIORITIES', required:true,  options: priorityOptions, maxSelections: 3 },
      { id:'q65', col: SATISFACTION_COLS.Q65_ONE_CHANGE,     text:'The #1 change you want the union to make.',              type:PARA, section:'11', sectionKey:'PRIORITIES', required:true  },
      { id:'q66', col: SATISFACTION_COLS.Q66_KEEP_DOING,     text:'One thing the union should keep doing.',                 type:PARA, section:'11', sectionKey:'PRIORITIES', required:true  },
      { id:'q67', col: SATISFACTION_COLS.Q67_ADDITIONAL,     text:'Additional comments — please do not include names.',     type:PARA, section:'11', sectionKey:'PRIORITIES', required:false }
    ];

    var sections = [
      { key:'WORK_CONTEXT',   number:'1',  title:'Work Context' },
      { key:'OVERALL_SAT',    number:'2',  title:'Overall Satisfaction' },
      { key:'STEWARD_3A',     number:'3A', title:'Steward Experience (if you had contact)' },
      { key:'STEWARD_3B',     number:'3B', title:'Steward Access (if you had no contact)' },
      { key:'CHAPTER',        number:'4',  title:'Chapter Effectiveness' },
      { key:'LEADERSHIP',     number:'5',  title:'Local Leadership' },
      { key:'CONTRACT',       number:'6',  title:'Contract Enforcement' },
      { key:'REPRESENTATION', number:'6A', title:'Representation Process (if you filed a grievance)' },
      { key:'COMMUNICATION',  number:'7',  title:'Communication Quality' },
      { key:'MEMBER_VOICE',   number:'8',  title:'Member Voice & Culture' },
      { key:'VALUE_ACTION',   number:'9',  title:'Value & Collective Action' },
      { key:'SCHEDULING',     number:'10', title:'Scheduling & Office Days' },
      { key:'PRIORITIES',     number:'11', title:'Priorities & Closing Thoughts' }
    ];

    return {
      questions: questions,
      sections: sections,
      sliderLabels: { min: 'Strongly Disagree', max: 'Strongly Agree' },
      period: getSurveyPeriod()
    };

  } catch (e) {
    Logger.log('getSurveyQuestions error: ' + e.message);
    return { questions: [], sections: [], sliderLabels: { min: 'Strongly Disagree', max: 'Strongly Agree' }, period: null };
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
function submitSurveyResponse(callerEmail, responses) {
  if (!callerEmail || !responses) return { success: false, message: 'Not authenticated.' };

  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
      if (!satSheet) return { success: false, message: 'Survey sheet not found.' };

      // ── Period check ─────────────────────────────────────────────────
      var period = getSurveyPeriod();
      if (!period || period.status !== 'Active') {
        return { success: false, message: 'No survey is currently open. Check back next quarter.' };
      }

      // ── Hash email in-memory — raw value used only for tracking update below ──
      var emailHash   = hashForVault_(callerEmail);
      var periodId    = period.periodId;

      // ── Vault duplicate check (period-scoped) ─────────────────────────
      var vaultSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT);
      if (vaultSheet && vaultSheet.getLastRow() > 1) {
        var vaultData = vaultSheet.getRange(
          2, 1, vaultSheet.getLastRow() - 1, vaultSheet.getLastColumn()
        ).getValues();
        for (var v = 0; v < vaultData.length; v++) {
          var vHash    = String(vaultData[v][SURVEY_VAULT_COLS.EMAIL - 1]           || '');
          var vPeriod  = String(vaultData[v][SURVEY_VAULT_COLS.QUARTER - 1]         || '');
          var vLatest  = String(vaultData[v][SURVEY_VAULT_COLS.IS_LATEST - 1]       || '');
          if (vHash === emailHash && vPeriod === periodId && vLatest !== 'Superseded') {
            return { success: false, message: 'You have already submitted a response for this survey period.' };
          }
        }
      }

      // ── Build response row (Timestamp + 67 question columns) ─────────
      // Column order matches SATISFACTION_COLS exactly.
      // Unanswered conditional questions (branch not taken) → empty string.
      // Numeric slider values clamped to 1-10.
      function _num(val) {
        var n = parseInt(val, 10);
        if (isNaN(n)) return '';
        return Math.min(10, Math.max(1, n));
      }
      function _txt(val) {
        // Prevent formula injection per CR-FORMULA
        var s = String(val || '').trim();
        if (s.charAt(0) === '=') s = "'" + s;
        return s.substring(0, 3000); // cap free-text at 3k chars
      }
      function _csv(val) {
        // Checkbox: accept array or comma string
        if (Array.isArray(val)) return _txt(val.join(', '));
        return _txt(val);
      }

      var r = responses;
      var row = [
        new Date(),                                    // A - Timestamp
        _txt(r.q1),                                    // B - Q1 Worksite
        _txt(r.q2),                                    // C - Q2 Role
        _txt(r.q3),                                    // D - Q3 Shift
        _txt(r.q4),                                    // E - Q4 Tenure
        _txt(r.q5),                                    // F - Q5 Steward contact (branch)
        _num(r.q6),                                    // G - Q6
        _num(r.q7),                                    // H - Q7
        _num(r.q8),                                    // I - Q8
        _num(r.q9),                                    // J - Q9
        _num(r.q10),                                   // K - Q10 (3A)
        _num(r.q11),                                   // L
        _num(r.q12),                                   // M
        _num(r.q13),                                   // N
        _num(r.q14),                                   // O
        _num(r.q15),                                   // P
        _num(r.q16),                                   // Q
        _txt(r.q17),                                   // R - Q17 paragraph
        _num(r.q18),                                   // S - Q18 (3B)
        _num(r.q19),                                   // T
        _num(r.q20),                                   // U
        _num(r.q21),                                   // V - Q21 (Chapter)
        _num(r.q22),                                   // W
        _num(r.q23),                                   // X
        _num(r.q24),                                   // Y
        _num(r.q25),                                   // Z
        _num(r.q26),                                   // AA - Q26 (Leadership)
        _num(r.q27),                                   // AB
        _num(r.q28),                                   // AC
        _num(r.q29),                                   // AD
        _num(r.q30),                                   // AE
        _num(r.q31),                                   // AF
        _num(r.q32),                                   // AG - Q32 (Contract)
        _num(r.q33),                                   // AH
        _num(r.q34),                                   // AI
        _num(r.q35),                                   // AJ
        _txt(r.q36),                                   // AK - Q36 grievance branch
        _num(r.q37),                                   // AL - Q37 (6A)
        _num(r.q38),                                   // AM
        _num(r.q39),                                   // AN
        _num(r.q40),                                   // AO
        _num(r.q41),                                   // AP - Q41 (Comm)
        _num(r.q42),                                   // AQ
        _num(r.q43),                                   // AR
        _num(r.q44),                                   // AS
        _num(r.q45),                                   // AT
        _num(r.q46),                                   // AU - Q46 (Voice)
        _num(r.q47),                                   // AV
        _num(r.q48),                                   // AW
        _num(r.q49),                                   // AX
        _num(r.q50),                                   // AY
        _num(r.q51),                                   // AZ - Q51 (Value)
        _num(r.q52),                                   // BA
        _num(r.q53),                                   // BB
        _num(r.q54),                                   // BC
        _num(r.q55),                                   // BD
        _num(r.q56),                                   // BE - Q56 (Scheduling)
        _num(r.q57),                                   // BF
        _num(r.q58),                                   // BG
        _num(r.q59),                                   // BH
        _num(r.q60),                                   // BI
        _num(r.q61),                                   // BJ
        _num(r.q62),                                   // BK
        _txt(r.q63),                                   // BL - Q63 paragraph
        _csv(r.q64),                                   // BM - Q64 checkboxes
        _txt(r.q65),                                   // BN - Q65 paragraph
        _txt(r.q66),                                   // BO - Q66 paragraph
        _txt(r.q67)                                    // BP - Q67 paragraph
      ];

      satSheet.appendRow(row);
      var newResponseRow = satSheet.getLastRow();

      // ── Write vault entry (hashed PII only) ──────────────────────────
      if (vaultSheet) {
        // Attempt member ID lookup (in-memory only — not stored in Satisfaction sheet)
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
        } catch(ex) { /* vault entry written without member ID hash */ }

        var memberIdHash = memberId ? hashForVault_(memberId) : '';
        var vaultRow = new Array(8).fill('');
        vaultRow[SURVEY_VAULT_COLS.RESPONSE_ROW - 1]     = newResponseRow;
        vaultRow[SURVEY_VAULT_COLS.EMAIL - 1]             = emailHash;
        vaultRow[SURVEY_VAULT_COLS.VERIFIED - 1]          = 'Yes';
        vaultRow[SURVEY_VAULT_COLS.MATCHED_MEMBER_ID - 1] = memberIdHash;
        vaultRow[SURVEY_VAULT_COLS.QUARTER - 1]           = periodId;
        vaultRow[SURVEY_VAULT_COLS.IS_LATEST - 1]         = 'Latest';
        vaultRow[SURVEY_VAULT_COLS.SUPERSEDED_BY - 1]     = '';
        vaultRow[SURVEY_VAULT_COLS.REVIEWER_NOTES - 1]    = '';
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
            tRow[SURVEY_TRACKING_COLS.CURRENT_STATUS   - 1] = 'Completed';
            tRow[SURVEY_TRACKING_COLS.COMPLETED_DATE   - 1] = new Date();
            tRow[SURVEY_TRACKING_COLS.TOTAL_COMPLETED  - 1] = (parseInt(tRow[SURVEY_TRACKING_COLS.TOTAL_COMPLETED - 1], 10) || 0) + 1;
            trackSheet.getRange(t + 1, 1, 1, tLastCol).setValues([tRow]);
            break;
          }
        }
      }

      // ── Update period response count ───────────────────────────────────
      try { incrementPeriodResponseCount_(periodId); } catch(ex) {}

      // ── Invalidate satisfaction summary cache ─────────────────────────
      try {
        CacheService.getScriptCache().remove('satisfactionSummary_' + periodId);
      } catch(ex) {}

      return { success: true, message: 'Thank you — your anonymous response has been recorded.' };

    } catch (e) {
      Logger.log('submitSurveyResponse error: ' + e.message + '\n' + (e.stack || ''));
      return { success: false, message: 'Error submitting survey. Please try again.' };
    }
  }, 30);
}

