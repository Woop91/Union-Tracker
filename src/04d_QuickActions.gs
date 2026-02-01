/**
 * ============================================================================
 * QuickActions.gs - Quick Actions and Email Functions
 * ============================================================================
 *
 * This module handles all quick action and email-related functions including:
 * - Quick Actions menu for Member Directory and Grievance Log
 * - Member email composition and sending
 * - Survey, contact form, and dashboard link emails
 * - Grievance status email updates
 * - Member grievance history display
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Quick actions menu and member email functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// QUICK ACTIONS MENU - Main Entry Points
// ============================================================================

/**
 * Shows the Quick Actions menu based on current sheet and selection
 * Routes to appropriate quick actions dialog for Member Directory or Grievance Log
 */
function showQuickActionsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  var name = sheet.getName();
  var selection = sheet.getActiveRange();

  if (!selection) {
    ui.alert('⚡ Quick Actions - How to Use',
      'Quick Actions provides contextual shortcuts for the selected row.\n\n' +
      'TO USE:\n' +
      '1. Go to Member Directory or Grievance Log\n' +
      '2. Click on a data row (not the header)\n' +
      '3. Run this menu item again\n\n' +
      'MEMBER DIRECTORY ACTIONS:\n' +
      '• Start new grievance for member\n' +
      '• Send email to member\n' +
      '• View grievance history\n' +
      '• Copy Member ID\n\n' +
      'GRIEVANCE LOG ACTIONS:\n' +
      '• Sync deadlines to calendar\n' +
      '• Setup Drive folder\n' +
      '• Quick status update\n' +
      '• Copy Grievance ID',
      ui.ButtonSet.OK);
    return;
  }

  var row = selection.getRow();
  if (row < 2) {
    ui.alert('Quick Actions',
      'Please select a data row, not the header row.\n\n' +
      'Click on row 2 or below to use Quick Actions.',
      ui.ButtonSet.OK);
    return;
  }

  if (name === SHEETS.MEMBER_DIR) {
    showMemberQuickActions(row);
  } else if (name === SHEETS.GRIEVANCE_LOG) {
    showGrievanceQuickActions(row);
  } else {
    ui.alert('⚡ Quick Actions',
      'Quick Actions is available for:\n\n' +
      '• Member Directory - actions for members\n' +
      '• Grievance Log - actions for grievances\n\n' +
      'Current sheet: ' + name + '\n\n' +
      'Navigate to one of the supported sheets and select a row.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// MEMBER QUICK ACTIONS
// ============================================================================

/**
 * Shows quick actions dialog for a member row in Member Directory
 * @param {number} row - The selected row number
 */
function showMemberQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var memberId = data[MEMBER_COLS.MEMBER_ID - 1];
  var name = data[MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[MEMBER_COLS.LAST_NAME - 1];
  var email = data[MEMBER_COLS.EMAIL - 1];
  var hasOpen = data[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1];

  // Build email section buttons (only if email exists)
  var emailButtons = '';
  if (email) {
    emailButtons =
      '<div class="section-header">📨 Email Options</div>' +
      '<button class="action-btn" onclick="google.script.run.composeEmailForMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Send Custom Email</div><div class="desc">Compose email to ' + email + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailDashboardLinkToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">🔗</span><span><div class="title">Send Dashboard Link</div><div class="desc">Share dashboard access with member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#1a73e8;display:flex;align-items:center;gap:10px}' +
    '.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}' +
    '.name{font-size:18px;font-weight:bold}' +
    '.id{color:#666;font-size:14px}' +
    '.status{margin-top:10px}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}' +
    '.open{background:#ffebee;color:#c62828}' +
    '.none{background:#e8f5e9;color:#2e7d32}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}' +
    '.action-btn:hover{background:#e8f4fd}' +
    '.icon{font-size:24px}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:12px;color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#1a73e8;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Quick Actions</h2>' +
    '<div class="info">' +
    '<div class="name">' + name + '</div>' +
    '<div class="id">' + memberId + ' | ' + (email || 'No email') + '</div>' +
    '<div class="status">' + (hasOpen === 'Yes' ? '<span class="badge open">🔴 Has Open Grievance</span>' : '<span class="badge none">🟢 No Open Grievances</span>') + '</div>' +
    '</div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Member Actions</div>' +
    '<button class="action-btn" onclick="google.script.run.openGrievanceFormForMember(' + row + ');google.script.host.close()"><span class="icon">📋</span><span><div class="title">Start New Grievance</div><div class="desc">Create a grievance for this member</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.showMemberGrievanceHistory(\'' + memberId + '\');google.script.host.close()"><span class="icon">📁</span><span><div class="title">View Grievance History</div><div class="desc">See all grievances for this member</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(\'' + memberId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Member ID</div><div class="desc">' + memberId + '</div></span></button>' +
    emailButtons +
    '</div>' +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(email ? 650 : 400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Quick Actions');
}

// ============================================================================
// GRIEVANCE QUICK ACTIONS
// ============================================================================

/**
 * Shows quick actions dialog for a grievance row in Grievance Log
 * @param {number} row - The selected row number
 */
function showGrievanceQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var grievanceId = data[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
  var memberId = data[GRIEVANCE_COLS.MEMBER_ID - 1];
  var name = data[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + data[GRIEVANCE_COLS.LAST_NAME - 1];
  var status = data[GRIEVANCE_COLS.STATUS - 1];
  var step = data[GRIEVANCE_COLS.CURRENT_STEP - 1];
  var daysTo = data[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
  var memberEmail = data[GRIEVANCE_COLS.MEMBER_EMAIL - 1];
  var isOpen = status === 'Open' || status === 'Pending Info' || status === 'In Arbitration' || status === 'Appealed';

  // Build email button (only if member has email)
  var emailStatusBtn = '';
  if (memberEmail) {
    emailStatusBtn =
      '<div class="section-header">📨 Communication</div>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailGrievanceStatusToMember(\'' + grievanceId + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Email Status to Member</div><div class="desc">Send grievance status update to ' + memberEmail + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#DC2626}' +
    '.info{background:#fff5f5;padding:15px;border-radius:8px;margin-bottom:20px;border-left:4px solid #DC2626}' +
    '.gid{font-size:18px;font-weight:bold}' +
    '.gmem{color:#666;font-size:14px}' +
    '.gstatus{margin-top:10px;display:flex;gap:10px;flex-wrap:wrap}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}' +
    '.action-btn:hover{background:#fff5f5}' +
    '.icon{font-size:24px}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:12px;color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#DC2626;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0}' +
    '.divider{height:1px;background:#e0e0e0;margin:10px 0}' +
    '.status-section{margin-top:15px;padding:15px;background:#f8f9fa;border-radius:8px}' +
    '.status-section h4{margin:0 0 10px}' +
    'select{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Grievance Actions</h2>' +
    '<div class="info">' +
    '<div class="gid">' + grievanceId + '</div>' +
    '<div class="gmem">' + name + ' (' + memberId + ')' + (memberEmail ? ' - ' + memberEmail : '') + '</div>' +
    '<div class="gstatus">' +
    '<span class="badge">' + status + '</span>' +
    '<span class="badge">' + step + '</span>' +
    (daysTo !== null && daysTo !== '' ? '<span class="badge" style="background:' + (daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0) ? '#ffebee;color:#c62828' : '#e3f2fd;color:#1565c0') + '">' + (daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0) ? '⚠️ Overdue' : '📅 ' + daysTo + ' days') + '</span>' : '') +
    '</div></div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Case Management</div>' +
    '<button class="action-btn" onclick="google.script.run.syncSingleGrievanceToCalendar(\'' + grievanceId + '\');google.script.host.close()"><span class="icon">📅</span><span><div class="title">Sync to Calendar</div><div class="desc">Add deadlines to Google Calendar</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.setupDriveFolderForGrievance();google.script.host.close()"><span class="icon">📁</span><span><div class="title">Setup Drive Folder</div><div class="desc">Create document folder</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(\'' + grievanceId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Grievance ID</div><div class="desc">' + grievanceId + '</div></span></button>' +
    emailStatusBtn +
    '</div>' +
    (isOpen ? '<div class="status-section"><h4>Quick Status Update</h4><select id="statusSelect"><option value="">-- Select --</option><option value="Open">Open</option><option value="Pending Info">Pending Info</option><option value="Settled">Settled</option><option value="Withdrawn">Withdrawn</option><option value="Won">Won</option><option value="Denied">Denied</option><option value="Closed">Closed</option></select><button class="action-btn" style="margin-top:10px" onclick="var s=document.getElementById(\'statusSelect\').value;if(!s){alert(\'Select status\');return}google.script.run.withSuccessHandler(function(){alert(\'Updated!\');google.script.host.close()}).quickUpdateGrievanceStatus(' + row + ',s)"><span class="icon">✓</span><span><div class="title">Update Status</div></span></button></div>' : '') +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(memberEmail ? 750 : 550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance Quick Actions');
}

/**
 * Quick update grievance status from the quick actions dialog
 * @param {number} row - The grievance row number
 * @param {string} newStatus - The new status value
 */
function quickUpdateGrievanceStatus(row, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  sheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newStatus);
  if (['Closed', 'Settled', 'Withdrawn'].indexOf(newStatus) >= 0) {
    var closeCol = GRIEVANCE_COLS.DATE_CLOSED;
    if (!sheet.getRange(row, closeCol).getValue()) sheet.getRange(row, closeCol).setValue(new Date());
  }
  ss.toast('Grievance status updated to: ' + newStatus, 'Status Updated', 3);
}

// ============================================================================
// EMAIL COMPOSITION AND SENDING
// ============================================================================

/**
 * Opens email composition dialog for a member
 * @param {string} memberId - The member ID to compose email for
 */
function composeEmailForMember(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      var email = data[i][MEMBER_COLS.EMAIL - 1];
      var name = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      if (!email) { SpreadsheetApp.getUi().alert('No email on file.'); return; }
      var html = HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}.form-group{margin:15px 0}label{display:block;font-weight:bold;margin-bottom:5px}input,textarea{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box}textarea{min-height:200px}input:focus,textarea:focus{outline:none;border-color:#1a73e8}.buttons{display:flex;gap:10px;margin-top:20px}button{padding:12px 24px;font-size:14px;border:none;border-radius:4px;cursor:pointer;flex:1}.primary{background:#1a73e8;color:white}.secondary{background:#6c757d;color:white}</style></head><body><div class="container"><h2>📧 Email to Member</h2><div class="info"><strong>' + name + '</strong> (' + memberId + ')<br>' + email + '</div><div class="form-group"><label>Subject:</label><input type="text" id="subject" placeholder="Email subject"></div><div class="form-group"><label>Message:</label><textarea id="message" placeholder="Type your message..."></textarea></div><div class="buttons"><button class="primary" onclick="send()">📤 Send</button><button class="secondary" onclick="google.script.host.close()">Cancel</button></div></div><script>function send(){var s=document.getElementById("subject").value.trim();var m=document.getElementById("message").value.trim();if(!s||!m){alert("Fill in subject and message");return}google.script.run.withSuccessHandler(function(){alert("Email sent!");google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).sendQuickEmail("' + email + '",s,m,"' + memberId + '")}</script></body></html>'
      ).setWidth(600).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(html, '📧 Compose Email');
      return;
    }
  }
}

/**
 * Sends a quick email to a member
 * @param {string} to - Recipient email address
 * @param {string} subject - Email subject
 * @param {string} body - Email body text
 * @param {string} memberId - Member ID for logging purposes
 * @returns {Object} Success status object
 */
function sendQuickEmail(to, subject, body, memberId) {
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'SEIU Local 509 Dashboard' });
    return { success: true };
  } catch (e) { throw new Error('Failed to send: ' + e.message); }
}

// ============================================================================
// QUICK ACTION EMAIL FUNCTIONS - Send Forms, Surveys, and Status Updates
// ============================================================================

/**
 * Email the satisfaction survey link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailSurveyToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var surveyUrl = getFormUrlFromConfig('satisfaction');
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Member Satisfaction Survey';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'We value your feedback! Please take a few minutes to complete our Member Satisfaction Survey.\n\n' +
    'Survey Link: ' + surveyUrl + '\n\n' +
    'Your responses help us improve our representation and services.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Survey sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send survey: ' + e.message);
  }
}

/**
 * Email the contact info update form link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailContactFormToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var formUrl = getFormUrlFromConfig('contact');
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Update Your Contact Information';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'Please take a moment to verify and update your contact information. ' +
    'Keeping your information current helps us serve you better.\n\n' +
    'Update Form: ' + formUrl + '\n\n' +
    'This only takes a minute and helps ensure you receive important updates.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Contact form sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send contact form: ' + e.message);
  }
}

/**
 * Email the member dashboard/portal link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailDashboardLinkToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  // Use the deployed web app URL with personalized member ID parameter
  var webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert('Web app is not deployed. Please deploy the web app first via Extensions > Apps Script > Deploy.');
    return;
  }
  var portalUrl = webAppUrl + '?id=' + memberId;
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Member Dashboard Access';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'You can access your personalized union member dashboard at:\n\n' +
    'Dashboard Link: ' + portalUrl + '\n\n' +
    'From the dashboard you can:\n' +
    '- View your member profile\n' +
    '- Track grievance status (if applicable)\n' +
    '- Stay updated on union activities\n\n' +
    'Keep this link private - it is personalized for you.\n\n' +
    'If you have any questions, please contact your steward.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard link sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send dashboard link: ' + e.message);
  }
}

/**
 * Email grievance status update to the member
 * @param {string} grievanceId - Grievance ID to look up details
 */
function emailGrievanceStatusToMember(grievanceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('Grievance Log not found.');
    return;
  }

  // Find the grievance
  var data = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var grievance = null;

  for (var i = 0; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      grievance = {
        id: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        memberId: data[i][GRIEVANCE_COLS.MEMBER_ID - 1],
        firstName: data[i][GRIEVANCE_COLS.FIRST_NAME - 1],
        lastName: data[i][GRIEVANCE_COLS.LAST_NAME - 1],
        status: data[i][GRIEVANCE_COLS.STATUS - 1],
        step: data[i][GRIEVANCE_COLS.CURRENT_STEP - 1],
        dateFiled: data[i][GRIEVANCE_COLS.DATE_FILED - 1],
        nextAction: data[i][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1],
        daysToDeadline: data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
        issueCategory: data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
        steward: data[i][GRIEVANCE_COLS.STEWARD - 1],
        email: data[i][GRIEVANCE_COLS.MEMBER_EMAIL - 1]
      };
      break;
    }
  }

  if (!grievance) {
    SpreadsheetApp.getUi().alert('Grievance not found: ' + grievanceId);
    return;
  }

  if (!grievance.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Grievance Status Update (' + grievance.id + ')';
  var body = 'Dear ' + grievance.firstName + ',\n\n' +
    'Here is the current status of your grievance:\n\n' +
    '================================\n' +
    'GRIEVANCE STATUS UPDATE\n' +
    '================================\n\n' +
    'Grievance ID: ' + grievance.id + '\n' +
    'Issue Category: ' + (grievance.issueCategory || 'Not specified') + '\n' +
    'Current Status: ' + grievance.status + '\n' +
    'Current Step: ' + grievance.step + '\n' +
    'Date Filed: ' + (grievance.dateFiled ? new Date(grievance.dateFiled).toLocaleDateString() : 'N/A') + '\n';

  if (grievance.daysToDeadline !== null && grievance.daysToDeadline !== '') {
    if (grievance.daysToDeadline === 'Overdue' || (typeof grievance.daysToDeadline === 'number' && grievance.daysToDeadline < 0)) {
      body += 'Next Deadline: OVERDUE\n';
    } else {
      body += 'Days Until Next Deadline: ' + grievance.daysToDeadline + '\n';
    }
  }

  if (grievance.steward) {
    body += 'Assigned Steward: ' + grievance.steward + '\n';
  }

  body += '\n================================\n\n' +
    'If you have any questions about your grievance, please contact your steward.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: grievance.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Status update sent to ' + grievance.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send status update: ' + e.message);
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Helper: Get member data by Member ID
 * @private
 * @param {string} memberId - The member ID to look up
 * @returns {Object|null} Member data object or null if not found
 */
function getMemberDataById_(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      return {
        memberId: memberId,
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1],
        lastName: data[i][MEMBER_COLS.LAST_NAME - 1],
        email: data[i][MEMBER_COLS.EMAIL - 1]
      };
    }
  }
  return null;
}

/**
 * Helper: Get organization name from Config
 * @private
 * @returns {string} Organization name or default value
 */
function getOrgNameFromConfig_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (configSheet) {
    var orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue();
    if (orgName) return orgName;
  }
  return 'SEIU Local 509';
}

// ============================================================================
// MEMBER GRIEVANCE HISTORY AND ACTIONS
// ============================================================================

/**
 * Shows grievance history dialog for a specific member
 * @param {string} memberId - The member ID to show history for
 */
function showMemberGrievanceHistory(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) { SpreadsheetApp.getUi().alert('No grievances found.'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.DATE_CLOSED).getValues();
  var mine = [];
  data.forEach(function(row) {
    if (row[GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
      mine.push({ id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1], status: row[GRIEVANCE_COLS.STATUS - 1], step: row[GRIEVANCE_COLS.CURRENT_STEP - 1], issue: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1], filed: row[GRIEVANCE_COLS.DATE_FILED - 1], closed: row[GRIEVANCE_COLS.DATE_CLOSED - 1] });
    }
  });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances for this member.'); return; }
  var list = mine.map(function(g) {
    return '<div style="background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;border-left:4px solid ' + (g.status === 'Open' ? '#f44336' : '#4caf50') + '"><strong>' + g.id + '</strong><br><span style="color:#666">Status: ' + g.status + ' | Step: ' + g.step + '</span><br><span style="color:#888;font-size:12px">' + g.issue + ' | Filed: ' + (g.filed ? new Date(g.filed).toLocaleDateString() : 'N/A') + '</span></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}</style></head><body><h2>📁 Grievance History</h2><div class="summary"><strong>Member ID:</strong> ' + memberId + '<br><strong>Total:</strong> ' + mine.length + '<br><strong>Open:</strong> ' + mine.filter(function(g) { return g.status === 'Open'; }).length + '<br><strong>Closed:</strong> ' + mine.filter(function(g) { return g.status !== 'Open'; }).length + '</div>' + list + '</body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance History - ' + memberId);
}

/**
 * Opens grievance form for a member (placeholder - directs to Grievance Log)
 * @param {number} row - The member row number
 */
function openGrievanceFormForMember(row) {
  SpreadsheetApp.getUi().alert('ℹ️ New Grievance', 'To start a new grievance for this member, navigate to the Grievance Log sheet and add a new row with their Member ID.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Syncs a single grievance to the calendar
 * @param {string} grievanceId - The grievance ID to sync
 */
function syncSingleGrievanceToCalendar(grievanceId) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing ' + grievanceId + '...', 'Calendar', 3);
  if (typeof syncDeadlinesToCalendar === 'function') syncDeadlinesToCalendar();
}

// ============================================================================
// DASHBOARD LINK AND STEWARD TOOLKIT EMAILS
// ============================================================================

/**
 * Prompts for email and sends member dashboard link
 * Opens a prompt dialog to enter email address and sends the dashboard URL
 */
function sendMemberDashboardLink() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl();

  var response = ui.prompt('Send Report', 'Enter Member Email:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    var email = response.getResponseText();

    if (!email || !email.includes('@')) {
      ui.alert('Please enter a valid email address.');
      return;
    }

    var body = "Hello,\n\n" +
      "Access your Union Member Dashboard:\n" + url + "\n\n" +
      "HOW TO ACCESS:\n" +
      "1. Open the spreadsheet using the link above\n" +
      "2. Go to: 509 Command > Command Center > Member Dashboard (No PII)\n" +
      "3. The dashboard shows aggregate union metrics (no personal info visible)\n\n" +
      "WHAT YOU'LL SEE:\n" +
      "- Morale & Trust Scores\n" +
      "- Leadership Pipeline\n" +
      "- Grievance Statistics\n" +
      "- Steward Contact Search\n" +
      "- Emergency Weingarten Rights\n\n" +
      "In Solidarity,\n" +
      "509 Strategic Command Center";

    try {
      MailApp.sendEmail(email, COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + " Your Union Dashboard Access", body);
      ui.alert('Dashboard access link sent to ' + email);
    } catch (e) {
      ui.alert('Error sending email: ' + e.message);
    }
  }
}

/**
 * Sends the Member Dashboard URL to the selected member from Member Directory.
 * Uses the currently selected row to get member email and name.
 * This is a PII-protected view link.
 * NOTE: Duplicate exists in 11_SecureMemberDashboard.gs - this version kept for compatibility
 * @deprecated Use emailDashboardLink() in 11_SecureMemberDashboard.gs
 */
function emailDashboardLink_UIService_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveRange().getRow();

  // Validate we're on Member Directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR || row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a member row in the Member Directory first.');
    return;
  }

  // Get member email and name from the selected row
  var email = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();

  if (!email || !email.toString().includes('@')) {
    SpreadsheetApp.getUi().alert('No valid email found for this member.');
    return;
  }

  // Get organization name from config if available
  var orgName = '509';
  try {
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var configOrgName = configSheet.getRange(2, CONFIG_COLS.ORG_NAME).getValue();
      if (configOrgName) orgName = configOrgName;
    }
  } catch (e) {
    // Use default org name
  }

  // Build email body
  var dashboardUrl = ss.getUrl();
  var body = 'Hi ' + firstName + ',\n\n' +
    'You can view current union stats and representation here:\n' +
    dashboardUrl + '\n\n' +
    'This is a PII-protected view showing only aggregate union statistics.\n\n' +
    'From the dashboard you can:\n' +
    '- View active grievance counts and outcomes\n' +
    '- See member satisfaction trends\n' +
    '- Find your steward contact information\n' +
    '- Track union coverage and goals\n\n' +
    'If you have questions about your specific case or concerns, ' +
    'please contact your assigned steward directly.\n\n' +
    'In Solidarity,\n' +
    orgName + ' Union Leadership';

  try {
    MailApp.sendEmail({
      to: email,
      subject: orgName + ' - Your Member Dashboard Access',
      body: body,
      name: orgName + ' Union'
    });
    ss.toast('Sent dashboard link to ' + email, 'Success', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + e.message);
  }
}

/**
 * Sends the steward toolkit welcome email to a newly promoted steward
 * @private
 * @param {string} email - The steward's email address
 * @param {string} name - The steward's name
 */
function sendStewardToolkit_(email, name) {
  try {
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Welcome to the Steward Team!';
    var body = 'Congratulations ' + name + '!\n\n' +
      'You have been promoted to Union Steward. Welcome to the leadership team!\n\n' +
      'STEWARD TOOLKIT:\n' +
      '1. Access the Grievance Log to track and manage cases\n' +
      '2. Use the Executive Command dashboard for your analytics\n' +
      '3. Review the Member Directory for your assigned members\n' +
      '4. Contact your Chief Steward for mentorship and guidance\n\n' +
      'KEY RESPONSIBILITIES:\n' +
      '- Represent members in grievance proceedings\n' +
      '- Document workplace issues and contract violations\n' +
      '- Maintain confidentiality of member information\n' +
      '- Attend steward training and meetings\n\n' +
      'We are stronger together!' +
      COMMAND_CONFIG.EMAIL.FOOTER;

    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('Error sending steward toolkit: ' + e.message);
  }
}
