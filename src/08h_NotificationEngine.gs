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

    if (daysToDeadline !== '' && daysToDeadline <= 3) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });
    }
  }

  if (urgent.length === 0) return;

  var subject = '509 Dashboard: ' + urgent.length + ' Grievance Deadline(s) Approaching';
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
      '509 Dashboard Test Notification',
      'This is a test notification from your 509 Dashboard.\n\n' +
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
