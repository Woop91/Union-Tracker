/**
 * ============================================================================
 * 09b_TabModals.gs — Tab-Specific Modal Dialogs
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Provides modal dialog UIs for data entry and viewing across multiple tabs.
 *   Each modal replaces error-prone direct-cell editing with clean forms.
 *   Includes: Take Attendance, Submit Feedback, Add Event, Add Minutes,
 *   Log Volunteer Hours, Add Contact, Case Progress, Survey Response Viewer,
 *   Welcome Wizard, Searchable Help, Question Editor, Weekly Question Creator.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Direct cell editing in data-dense sheets is error-prone and slow. Modals
 *   provide validation, auto-population (timestamps, IDs, member names), and
 *   a form UX that prevents malformed data. Uses inline HTML (not separate
 *   .html files) to stay within the GAS 60-file limit.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Modal dialogs fail to open. Users can still enter data directly in sheets.
 *   No data loss — modals are convenience wrappers over sheet writes.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, SHEET_COLORS, DIALOG_SIZES),
 *               23_PortalSheets.gs (portal sheet helpers),
 *               00_Security.gs (escapeHtml)
 *   Used by:    03_UIComponents.gs (menu items)
 *
 * @fileoverview Tab-specific modal dialogs for data entry
 * @requires 01_Core.gs, 00_Security.gs
 */

// ============================================================================
// SHARED MODAL STYLES
// ============================================================================

/**
 * Returns shared CSS for all tab modals.
 * @returns {string} CSS block
 * @private
 */
function getModalStyles_() {
  return '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; color: #1F2937; background: #F9FAFB; }' +
    'h2 { color: #1A2A4A; margin-bottom: 16px; font-size: 18px; }' +
    '.form-group { margin-bottom: 14px; }' +
    'label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 13px; color: #374151; }' +
    'input, select, textarea { width: 100%; padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; font-family: inherit; }' +
    'input:focus, select:focus, textarea:focus { outline: none; border-color: #2C5282; box-shadow: 0 0 0 3px rgba(44,82,130,0.1); }' +
    'textarea { resize: vertical; min-height: 80px; }' +
    '.btn { display: inline-block; padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; margin-right: 8px; }' +
    '.btn-primary { background: #1A2A4A; color: white; }' +
    '.btn-primary:hover { background: #2C5282; }' +
    '.btn-secondary { background: #E5E7EB; color: #374151; }' +
    '.btn-secondary:hover { background: #D1D5DB; }' +
    '.btn-success { background: #276749; color: white; }' +
    '.btn-success:hover { background: #059669; }' +
    '.actions { margin-top: 20px; text-align: right; }' +
    '.msg { padding: 10px 14px; border-radius: 6px; margin-bottom: 12px; font-size: 13px; }' +
    '.msg-success { background: #D1FAE5; color: #065F46; }' +
    '.msg-error { background: #FEE2E2; color: #991B1B; }' +
    '.msg-info { background: #DBEAFE; color: #1E40AF; }' +
    '.row { display: flex; gap: 12px; }' +
    '.row > .form-group { flex: 1; }' +
    '.checklist-item { display: flex; align-items: center; padding: 8px 12px; border-bottom: 1px solid #E5E7EB; }' +
    '.checklist-item label { margin: 0; font-weight: normal; cursor: pointer; flex: 1; }' +
    '.checklist-item input[type=checkbox] { width: auto; margin-right: 10px; transform: scale(1.2); }' +
    '.progress-bar { height: 8px; background: #E5E7EB; border-radius: 4px; overflow: hidden; margin: 8px 0; }' +
    '.progress-fill { height: 100%; background: #276749; border-radius: 4px; transition: width 0.3s; }' +
    '.card { background: white; border-radius: 8px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    '.search-box { position: relative; margin-bottom: 16px; }' +
    '.search-box input { padding-left: 36px; }' +
    '.search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #9CA3AF; }' +
    '.tabs { display: flex; border-bottom: 2px solid #E5E7EB; margin-bottom: 16px; }' +
    '.tab { padding: 8px 16px; cursor: pointer; font-weight: 600; color: #6B7280; border-bottom: 2px solid transparent; margin-bottom: -2px; }' +
    '.tab.active { color: #1A2A4A; border-bottom-color: #D4A017; }' +
    '.tab-content { display: none; }' +
    '.tab-content.active { display: block; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }' +
    '.spinner { display: none; }' +
    '.spinner.show { display: inline-block; width: 16px; height: 16px; border: 2px solid #E5E7EB; border-top-color: #1A2A4A; border-radius: 50%; animation: spin 0.6s linear infinite; margin-right: 8px; }' +
    '@keyframes spin { to { transform: rotate(360deg); } }' +
    '</style>';
}

/**
 * Returns the base target and meta tags for modal HTML.
 * @returns {string}
 * @private
 */
function getModalHead_() {
  return '<base target="_top"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">';
}

// ============================================================================
// MODAL 1: SUBMIT FEEDBACK
// ============================================================================

/**
 * Shows the Submit Feedback modal dialog.
 * Menu: Tools > Feedback > Submit New Idea/Bug Report
 */
function showSubmitFeedbackModal() {
  var html = HtmlService.createHtmlOutput(getSubmitFeedbackHtml_())
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.MEDIUM.height);
  SpreadsheetApp.getUi().showModalDialog(html, '💡 Submit Feedback');
}

function getSubmitFeedbackHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Submit Feedback or Feature Request</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Category</label>' +
    '<select id="category"><option value="">Select...</option><option>Bug Report</option><option>Feature Request</option><option>Improvement</option><option>Question</option><option>Other</option></select></div>' +
    '<div class="row"><div class="form-group"><label>Priority</label>' +
    '<select id="priority"><option value="">Select...</option><option>Low</option><option>Medium</option><option>High</option><option>Critical</option></select></div>' +
    '<div class="form-group"><label>Status</label><input id="status" value="New" readonly></div></div>' +
    '<div class="form-group"><label>Title</label><input id="title" placeholder="Brief summary of the feedback"></div>' +
    '<div class="form-group"><label>Description</label><textarea id="description" rows="4" placeholder="Detailed description..."></textarea></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitFeedback()">Submit</button></div>' +
    '<script>' +
    'function submitFeedback() {' +
    '  var cat = document.getElementById("category").value;' +
    '  var pri = document.getElementById("priority").value;' +
    '  var title = document.getElementById("title").value;' +
    '  var desc = document.getElementById("description").value;' +
    '  if (!cat || !pri || !title) { showMsg("Please fill in Category, Priority, and Title.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  document.getElementById("submitBtn").textContent = "Submitting...";' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Feedback submitted! ID: " + r.id, "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed to submit", "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Submit"; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Submit"; })' +
    '  .modalSubmitFeedback(cat, pri, title, desc);' +
    '}' +
    'function showMsg(text, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = text; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Server-side handler for feedback submission.
 * @param {string} category
 * @param {string} priority
 * @param {string} title
 * @param {string} description
 * @returns {{ success: boolean, id?: string, error?: string }}
 */
function modalSubmitFeedback(category, priority, title, description) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.FEEDBACK);
    if (!sheet) return { success: false, error: 'Feedback sheet not found' };

    var timestamp = new Date();
    var user = Session.getActiveUser().getEmail() || 'Unknown';
    var id = 'FB-' + Utilities.formatDate(timestamp, 'America/New_York', 'yyyyMMdd-HHmmss');

    sheet.appendRow([
      escapeForFormula(timestamp),
      escapeForFormula(user),
      escapeForFormula(category),
      escapeForFormula(priority),
      escapeForFormula(title),
      escapeForFormula(description),
      'New',
      ''
    ]);

    return { success: true, id: id };
  } catch (e) {
    Logger.log('modalSubmitFeedback error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 2: ADD EVENT
// ============================================================================

/**
 * Shows the Add Event modal dialog.
 * Menu: Tools > Calendar & Meetings > Add New Event
 */
function showAddEventModal() {
  var html = HtmlService.createHtmlOutput(getAddEventHtml_())
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Add New Event');
}

function getAddEventHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Create a New Event</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Title</label><input id="title" placeholder="Event title"></div>' +
    '<div class="row"><div class="form-group"><label>Type</label>' +
    '<select id="type"><option>Meeting</option><option>Negotiation</option><option>Training</option><option>Social</option><option>Community</option></select></div>' +
    '<div class="form-group"><label>Date & Time</label><input type="datetime-local" id="dateTime"></div></div>' +
    '<div class="row"><div class="form-group"><label>End Time (optional)</label><input type="datetime-local" id="endTime"></div>' +
    '<div class="form-group"><label>Location</label><input id="location" placeholder="Room, address, or virtual"></div></div>' +
    '<div class="form-group"><label>Description</label><textarea id="description" rows="3" placeholder="Event details..."></textarea></div>' +
    '<div class="form-group"><label>Zoom/Video Link (optional)</label><input id="zoomLink" placeholder="https://..."></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitEvent()">Create Event</button></div>' +
    '<script>' +
    'function submitEvent() {' +
    '  var title = document.getElementById("title").value;' +
    '  var dt = document.getElementById("dateTime").value;' +
    '  if (!title || !dt) { showMsg("Title and Date/Time are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  document.getElementById("submitBtn").textContent = "Creating...";' +
    '  var data = { title: title, type: document.getElementById("type").value, dateTime: dt,' +
    '    endTime: document.getElementById("endTime").value, location: document.getElementById("location").value,' +
    '    description: document.getElementById("description").value, zoomLink: document.getElementById("zoomLink").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Event created! ID: " + r.id, "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Create Event"; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Create Event"; })' +
    '  .modalAddEvent(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Server-side handler for event creation.
 * @param {Object} data - Event data
 * @returns {{ success: boolean, id?: string, error?: string }}
 */
function modalAddEvent(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.PORTAL_EVENTS) || ss.getSheetByName('Events');
    if (!sheet) return { success: false, error: 'Events sheet not found' };

    var id = 'EVT-' + Date.now();
    var user = Session.getActiveUser().getEmail() || 'Unknown';

    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(data.title),
      escapeForFormula(data.type),
      escapeForFormula(data.dateTime),
      escapeForFormula(data.endTime || ''),
      escapeForFormula(data.location || ''),
      escapeForFormula(data.description || ''),
      escapeForFormula(data.zoomLink || ''),
      escapeForFormula(user),
      new Date()
    ]);

    return { success: true, id: id };
  } catch (e) {
    Logger.log('modalAddEvent error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 3: ADD MEETING MINUTES
// ============================================================================

function showAddMinutesModal() {
  var html = HtmlService.createHtmlOutput(getAddMinutesHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Add Meeting Minutes');
}

function getAddMinutesHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Add Meeting Minutes</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>Meeting Date</label><input type="date" id="meetingDate"></div>' +
    '<div class="form-group"><label>Title</label><input id="title" placeholder="Monthly General Meeting"></div></div>' +
    '<div class="form-group"><label>Bullet Points (one per line)</label><textarea id="bullets" rows="5" placeholder="Motion to approve budget passed unanimously\nCommittee reports presented\nNext meeting scheduled for..."></textarea></div>' +
    '<div class="form-group"><label>Full Minutes</label><textarea id="fullMinutes" rows="6" placeholder="Detailed meeting minutes..."></textarea></div>' +
    '<div class="form-group"><label>Google Drive Doc URL (optional)</label><input id="driveUrl" placeholder="https://docs.google.com/..."></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitMinutes()">Save Minutes</button></div>' +
    '<script>' +
    'function submitMinutes() {' +
    '  var date = document.getElementById("meetingDate").value;' +
    '  var title = document.getElementById("title").value;' +
    '  if (!date || !title) { showMsg("Date and Title are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { meetingDate: date, title: title,' +
    '    bullets: document.getElementById("bullets").value, fullMinutes: document.getElementById("fullMinutes").value,' +
    '    driveUrl: document.getElementById("driveUrl").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Minutes saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalAddMinutes(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalAddMinutes(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.PORTAL_MINUTES) || ss.getSheetByName('MeetingMinutes');
    if (!sheet) return { success: false, error: 'MeetingMinutes sheet not found' };

    var id = 'MIN-' + Date.now();
    var user = Session.getActiveUser().getEmail() || 'Unknown';

    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(data.meetingDate),
      escapeForFormula(data.title),
      escapeForFormula(data.bullets || ''),
      escapeForFormula(data.fullMinutes || ''),
      escapeForFormula(user),
      new Date(),
      escapeForFormula(data.driveUrl || '')
    ]);

    return { success: true, id: id };
  } catch (e) {
    Logger.log('modalAddMinutes error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 4: LOG VOLUNTEER HOURS
// ============================================================================

function showLogVolunteerHoursModal() {
  var html = HtmlService.createHtmlOutput(getLogVolunteerHoursHtml_())
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.MEDIUM.height);
  SpreadsheetApp.getUi().showModalDialog(html, '🤝 Log Volunteer Hours');
}

function getLogVolunteerHoursHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Log Volunteer Hours</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>Member ID</label><input id="memberId" placeholder="MBR-001"></div>' +
    '<div class="form-group"><label>Member Name</label><input id="memberName" readonly placeholder="Auto-populated"></div></div>' +
    '<div class="row"><div class="form-group"><label>Activity Date</label><input type="date" id="activityDate"></div>' +
    '<div class="form-group"><label>Activity Type</label>' +
    '<select id="activityType"><option>Meeting</option><option>Outreach</option><option>Event</option><option>Training</option><option>Admin</option><option>Other</option></select></div></div>' +
    '<div class="row"><div class="form-group"><label>Hours</label><input type="number" id="hours" min="0.25" step="0.25" placeholder="2"></div>' +
    '<div class="form-group"><label>Verified By (optional)</label><input id="verifiedBy" placeholder="Steward name"></div></div>' +
    '<div class="form-group"><label>Description</label><textarea id="description" rows="3" placeholder="What was done..."></textarea></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitHours()">Log Hours</button></div>' +
    '<script>' +
    'document.getElementById("memberId").addEventListener("blur", function() {' +
    '  var mid = this.value.trim();' +
    '  if (mid) { google.script.run.withSuccessHandler(function(name) {' +
    '    document.getElementById("memberName").value = name || "Not found";' +
    '  }).modalLookupMemberName(mid); }' +
    '});' +
    'function submitHours() {' +
    '  var mid = document.getElementById("memberId").value;' +
    '  var date = document.getElementById("activityDate").value;' +
    '  var hours = document.getElementById("hours").value;' +
    '  if (!mid || !date || !hours || parseFloat(hours) <= 0) { showMsg("Member ID, Date, and Hours are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { memberId: mid, memberName: document.getElementById("memberName").value,' +
    '    activityDate: date, activityType: document.getElementById("activityType").value,' +
    '    hours: parseFloat(hours), description: document.getElementById("description").value,' +
    '    verifiedBy: document.getElementById("verifiedBy").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Hours logged!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalLogVolunteerHours(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Looks up a member name by ID from the Member Directory.
 * @param {string} memberId
 * @returns {string} Member name or empty string
 */
function modalLookupMemberName(memberId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return '';
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === String(memberId).trim()) {
        // Column B = First Name, Column C = Last Name (typical layout)
        return String(data[r][1] || '') + ' ' + String(data[r][2] || '');
      }
    }
    return '';
  } catch (_e) {
    return '';
  }
}

function modalLogVolunteerHours(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.VOLUNTEER_HOURS);
    if (!sheet) return { success: false, error: 'Volunteer Hours sheet not found' };

    var id = 'VH-' + Date.now();
    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(data.memberId),
      escapeForFormula(data.memberName || ''),
      escapeForFormula(data.activityDate),
      escapeForFormula(data.activityType),
      data.hours,
      escapeForFormula(data.description || ''),
      escapeForFormula(data.verifiedBy || ''),
      ''
    ]);

    return { success: true, id: id };
  } catch (e) {
    Logger.log('modalLogVolunteerHours error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 5: ADD NON-MEMBER CONTACT
// ============================================================================

function showAddContactModal() {
  var html = HtmlService.createHtmlOutput(getAddContactHtml_())
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.MEDIUM.height);
  SpreadsheetApp.getUi().showModalDialog(html, '👥 Add Non-Member Contact');
}

function getAddContactHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Add Non-Member Contact</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>First Name</label><input id="firstName"></div>' +
    '<div class="form-group"><label>Last Name</label><input id="lastName"></div></div>' +
    '<div class="row"><div class="form-group"><label>Job Title</label><input id="jobTitle" placeholder="Supervisor, Manager..."></div>' +
    '<div class="form-group"><label>Work Location</label><input id="workLocation"></div></div>' +
    '<div class="row"><div class="form-group"><label>Unit</label><input id="unit"></div>' +
    '<div class="form-group"><label>Cubicle/Office</label><input id="cubicle"></div></div>' +
    '<div class="form-group"><label>Office Days</label><input id="officeDays" placeholder="Mon, Tue, Wed..."></div>' +
    '<div class="row"><div class="form-group"><label>Email</label><input type="email" id="email"></div>' +
    '<div class="form-group"><label>Phone</label><input id="phone" placeholder="555-123-4567"></div></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitContact()">Add Contact</button></div>' +
    '<script>' +
    'function submitContact() {' +
    '  var fn = document.getElementById("firstName").value;' +
    '  var ln = document.getElementById("lastName").value;' +
    '  if (!fn || !ln) { showMsg("First and Last Name are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { firstName: fn, lastName: ln, jobTitle: document.getElementById("jobTitle").value,' +
    '    workLocation: document.getElementById("workLocation").value, unit: document.getElementById("unit").value,' +
    '    cubicle: document.getElementById("cubicle").value, officeDays: document.getElementById("officeDays").value,' +
    '    email: document.getElementById("email").value, phone: document.getElementById("phone").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Contact added!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalAddContact(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalAddContact(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS) || ss.getSheetByName('Non member contacts');
    if (!sheet) return { success: false, error: 'Non-Member Contacts sheet not found' };

    // Duplicate check: same first+last name at same location
    var existing = sheet.getDataRange().getValues();
    for (var r = 1; r < existing.length; r++) {
      if (String(existing[r][0]).trim().toLowerCase() === data.firstName.trim().toLowerCase() &&
          String(existing[r][1]).trim().toLowerCase() === data.lastName.trim().toLowerCase() &&
          String(existing[r][3]).trim().toLowerCase() === data.workLocation.trim().toLowerCase()) {
        return { success: false, error: 'Duplicate: ' + data.firstName + ' ' + data.lastName + ' at ' + data.workLocation + ' already exists.' };
      }
    }

    sheet.appendRow([
      escapeForFormula(data.firstName),
      escapeForFormula(data.lastName),
      escapeForFormula(data.jobTitle || ''),
      escapeForFormula(data.workLocation || ''),
      escapeForFormula(data.unit || ''),
      escapeForFormula(data.cubicle || ''),
      escapeForFormula(data.officeDays || ''),
      escapeForFormula(data.email || ''),
      escapeForFormula(data.phone || '')
    ]);

    return { success: true };
  } catch (e) {
    Logger.log('modalAddContact error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 6: TAKE ATTENDANCE
// ============================================================================

function showTakeAttendanceModal() {
  var html = HtmlService.createHtmlOutput(getTakeAttendanceHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Take Meeting Attendance');
}

function getTakeAttendanceHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.member-list { max-height: 340px; overflow-y: auto; border: 1px solid #E5E7EB; border-radius: 6px; } .count { font-size: 14px; color: #6B7280; margin-top: 8px; }</style>' +
    '</head><body>' +
    '<h2>Take Meeting Attendance</h2>' +
    '<div id="msg"></div>' +
    '<div class="card">' +
    '<div class="row"><div class="form-group"><label>Meeting Date</label><input type="date" id="meetingDate"></div>' +
    '<div class="form-group"><label>Meeting Type</label>' +
    '<select id="meetingType"><option>Regular</option><option>Special</option><option>Committee</option><option>Emergency</option></select></div></div>' +
    '<div class="form-group"><label>Meeting Name</label><input id="meetingName" placeholder="March General Meeting"></div>' +
    '</div>' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin:12px 0"><h3>Members</h3>' +
    '<label style="font-weight:normal"><input type="checkbox" id="selectAll" onchange="toggleAll(this.checked)"> Select All</label></div>' +
    '<div class="member-list" id="memberList"><div class="msg msg-info">Loading members...</div></div>' +
    '<div class="count" id="count"></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-success" id="submitBtn" onclick="saveAttendance()">Save Attendance</button></div>' +
    '<script>' +
    'var members = [];' +
    'google.script.run.withSuccessHandler(function(list) {' +
    '  members = list || [];' +
    '  var html = "";' +
    '  for (var i = 0; i < members.length; i++) {' +
    '    html += "<div class=\\"checklist-item\\"><input type=\\"checkbox\\" id=\\"m" + i + "\\" onchange=\\"updateCount()\\"><label for=\\"m" + i + "\\">" + members[i].name + " <span style=\\"color:#9CA3AF;font-size:12px\\">(" + members[i].id + ")</span></label></div>";' +
    '  }' +
    '  document.getElementById("memberList").innerHTML = html || "<div class=\\"msg msg-info\\">No members found in directory.</div>";' +
    '  updateCount();' +
    '}).modalGetMemberList();' +
    'function toggleAll(checked) { for (var i = 0; i < members.length; i++) { document.getElementById("m"+i).checked = checked; } updateCount(); }' +
    'function updateCount() { var c = 0; for (var i = 0; i < members.length; i++) { if (document.getElementById("m"+i).checked) c++; }' +
    '  document.getElementById("count").textContent = c + " / " + members.length + " members present"; }' +
    'function saveAttendance() {' +
    '  var date = document.getElementById("meetingDate").value;' +
    '  var name = document.getElementById("meetingName").value;' +
    '  if (!date || !name) { showMsg("Date and Meeting Name required.", "error"); return; }' +
    '  var attended = [];' +
    '  for (var i = 0; i < members.length; i++) { attended.push(document.getElementById("m"+i).checked); }' +
    '  document.getElementById("submitBtn").disabled = true; document.getElementById("submitBtn").textContent = "Saving...";' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg(r.count + " attendance records saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Save Attendance"; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Save Attendance"; })' +
    '  .modalSaveAttendance({ date: date, name: name, type: document.getElementById("meetingType").value, attended: attended, members: members });' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Returns a list of active members for the attendance checklist.
 * @returns {Array<{id: string, name: string, email: string}>}
 */
function modalGetMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var members = [];
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][0] || '').trim();
      var name = String(data[r][1] || '').trim();
      if (data[r].length > 2) name += ' ' + String(data[r][2] || '').trim();
      if (id && name.trim()) {
        members.push({ id: escapeHtml(id), name: escapeHtml(name.trim()), email: String(data[r][3] || '') });
      }
    }
    return members;
  } catch (e) {
    Logger.log('modalGetMemberList error: ' + e.message);
    return [];
  }
}

/**
 * Saves attendance records — one row per member.
 * @param {Object} data - { date, name, type, attended: boolean[], members: [{id,name}] }
 * @returns {{ success: boolean, count?: number, error?: string }}
 */
function modalSaveAttendance(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_ATTENDANCE);
    if (!sheet) return { success: false, error: 'Meeting Attendance sheet not found' };

    var rows = [];
    var count = 0;

    for (var i = 0; i < data.members.length; i++) {
      var entryId = 'ATT-' + Date.now() + '-' + i;
      rows.push([
        escapeForFormula(entryId),
        escapeForFormula(data.date),
        escapeForFormula(data.type),
        escapeForFormula(data.name),
        escapeForFormula(data.members[i].id),
        escapeForFormula(data.members[i].name),
        data.attended[i] ? true : false,
        ''
      ]);
      if (data.attended[i]) count++;
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: count };
  } catch (e) {
    Logger.log('modalSaveAttendance error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 7: CASE PROGRESS VIEWER
// ============================================================================

function showCaseProgressModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var caseId = '';

  // Try to get Case ID from selected cell
  if (activeSheet.getName() === SHEETS.CASE_CHECKLIST || activeSheet.getName() === SHEETS.GRIEVANCE_LOG) {
    caseId = String(activeSheet.getActiveCell().getValue() || '');
  }

  var html = HtmlService.createHtmlOutput(getCaseProgressHtml_(caseId))
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '✅ Case Progress');
}

function getCaseProgressHtml_(caseId) {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.item-list { max-height: 380px; overflow-y: auto; } .item-done { text-decoration: line-through; color: #9CA3AF; } .item-required { color: #9B2335; font-weight: 600; }</style>' +
    '</head><body>' +
    '<h2>Case Checklist Progress</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Case ID</label><div class="row"><input id="caseId" value="' + escapeHtml(caseId) + '" placeholder="Enter Case ID">' +
    '<button class="btn btn-primary" style="white-space:nowrap" onclick="loadCase()">Load</button></div></div>' +
    '<div id="progressArea" style="display:none"><div class="card"><strong id="progressText"></strong>' +
    '<div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div></div>' +
    '<div class="item-list" id="items"></div>' +
    '<div class="actions"><button class="btn btn-success" onclick="saveChecks()">Save Changes</button></div></div>' +
    '<script>' +
    'var items = [];' +
    'function loadCase() {' +
    '  var cid = document.getElementById("caseId").value.trim();' +
    '  if (!cid) { showMsg("Enter a Case ID", "error"); return; }' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    items = result || [];' +
    '    if (items.length === 0) { showMsg("No checklist items found for " + cid, "info"); return; }' +
    '    renderItems();' +
    '    document.getElementById("progressArea").style.display = "block";' +
    '  }).modalGetCaseChecklist(cid);' +
    '}' +
    'function renderItems() {' +
    '  var done = 0, html = "";' +
    '  for (var i = 0; i < items.length; i++) {' +
    '    var it = items[i]; if (it.completed) done++;' +
    '    var cls = it.completed ? "item-done" : (it.required ? "item-required" : "");' +
    '    html += "<div class=\\"checklist-item\\"><input type=\\"checkbox\\" id=\\"c"+i+"\\" "+(it.completed?"checked":"")+' +
    '    " onchange=\\"updateProgress()\\"><label for=\\"c"+i+"\\" class=\\""+cls+"\\">"+it.text+' +
    '    (it.required && !it.completed ? " ⚠️" : "")+(it.actionType ? " <span class=\\"badge\\" style=\\"background:#DBEAFE;color:#1E40AF\\">"+it.actionType+"</span>" : "")+' +
    '    "</label></div>";' +
    '  }' +
    '  document.getElementById("items").innerHTML = html;' +
    '  updateProgress();' +
    '}' +
    'function updateProgress() {' +
    '  var done = 0;' +
    '  for (var i = 0; i < items.length; i++) { if (document.getElementById("c"+i).checked) done++; }' +
    '  var pct = Math.round(done / items.length * 100);' +
    '  document.getElementById("progressText").textContent = done + "/" + items.length + " complete (" + pct + "%)";' +
    '  document.getElementById("progressFill").style.width = pct + "%";' +
    '}' +
    'function saveChecks() {' +
    '  var updates = [];' +
    '  for (var i = 0; i < items.length; i++) { updates.push({ row: items[i].row, completed: document.getElementById("c"+i).checked }); }' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) showMsg("Saved!", "success"); else showMsg(r && r.error || "Failed", "error");' +
    '  }).modalSaveCaseChecklist(updates);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    (caseId ? 'loadCase();' : '') +
    '</script></body></html>';
}

function modalGetCaseChecklist(caseId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var items = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === String(caseId).trim()) {
        items.push({
          row: r + 1,
          checklistId: data[r][0],
          caseId: data[r][1],
          actionType: escapeHtml(String(data[r][2] || '')),
          text: escapeHtml(String(data[r][3] || '')),
          category: String(data[r][4] || ''),
          required: String(data[r][5]).toUpperCase() === 'Y',
          completed: data[r][6] === true,
          completedBy: data[r][7] || '',
          completedDate: data[r][8] || ''
        });
      }
    }
    return items;
  } catch (e) {
    Logger.log('modalGetCaseChecklist error: ' + e.message);
    return [];
  }
}

function modalSaveCaseChecklist(updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
    if (!sheet) return { success: false, error: 'Case Checklist sheet not found' };

    var user = Session.getActiveUser().getEmail() || 'Unknown';
    var now = new Date();

    for (var i = 0; i < updates.length; i++) {
      var u = updates[i];
      sheet.getRange(u.row, 7).setValue(u.completed);  // Completed column
      if (u.completed) {
        sheet.getRange(u.row, 8).setValue(escapeForFormula(user));  // Completed By
        sheet.getRange(u.row, 9).setValue(now);                     // Completed Date
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log('modalSaveCaseChecklist error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 8: SURVEY RESPONSE VIEWER
// ============================================================================

function showSurveyResponseViewer() {
  var html = HtmlService.createHtmlOutput(getSurveyResponseViewerHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Survey Response Viewer');
}

function getSurveyResponseViewerHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.response-card { padding: 16px; } .q-row { display:flex; align-items:center; margin-bottom:8px; } ' +
    '.q-label { flex:1; font-size:13px; color:#374151; } .q-value { font-weight:600; font-size:14px; min-width:40px; text-align:right; }' +
    '.bar-bg { flex:2; height:12px; background:#E5E7EB; border-radius:6px; margin:0 12px; overflow:hidden; }' +
    '.bar-fill { height:100%; border-radius:6px; transition:width 0.3s; }' +
    '.nav-btns { display:flex; justify-content:space-between; margin-top:16px; }' +
    '.demo-info { display:flex; gap:16px; flex-wrap:wrap; margin-bottom:16px; padding:12px; background:#EBF4FF; border-radius:8px; }' +
    '.demo-item { font-size:13px; } .demo-item strong { color:#1A2A4A; }</style>' +
    '</head><body>' +
    '<h2>Survey Response Viewer</h2>' +
    '<div id="msg"></div>' +
    '<div id="loading" class="msg msg-info">Loading responses...</div>' +
    '<div id="viewer" style="display:none">' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">' +
    '<span id="counter" style="font-size:13px;color:#6B7280"></span>' +
    '</div>' +
    '<div class="card response-card" id="responseCard"></div>' +
    '<div class="nav-btns"><button class="btn btn-secondary" onclick="prev()">← Previous</button>' +
    '<button class="btn btn-secondary" onclick="next()">Next →</button></div></div>' +
    '<script>' +
    'var responses = [], headers = [], idx = 0;' +
    'google.script.run.withSuccessHandler(function(data) {' +
    '  document.getElementById("loading").style.display = "none";' +
    '  if (!data || data.responses.length === 0) { showMsg("No survey responses found.", "info"); return; }' +
    '  responses = data.responses; headers = data.headers;' +
    '  document.getElementById("viewer").style.display = "block";' +
    '  renderResponse();' +
    '}).modalGetSurveyResponses();' +
    'function renderResponse() {' +
    '  var r = responses[idx]; document.getElementById("counter").textContent = "Response " + (idx+1) + " of " + responses.length;' +
    '  var html = "<div class=\\"demo-info\\">";' +
    '  for (var i = 0; i < Math.min(5, r.length); i++) { html += "<div class=\\"demo-item\\"><strong>" + headers[i] + ":</strong> " + (r[i] || "—") + "</div>"; }' +
    '  html += "</div>";' +
    '  for (var j = 5; j < r.length; j++) {' +
    '    var val = parseFloat(r[j]); var label = headers[j] || "Q" + (j+1);' +
    '    if (!isNaN(val) && val >= 1 && val <= 10) {' +
    '      var pct = val * 10; var color = val >= 7 ? "#6EE7B7" : (val >= 4 ? "#FDE68A" : "#FCA5A5");' +
    '      html += "<div class=\\"q-row\\"><span class=\\"q-label\\">" + label + "</span><div class=\\"bar-bg\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"></div></div><span class=\\"q-value\\">"+val+"</span></div>";' +
    '    } else if (r[j]) {' +
    '      html += "<div style=\\"margin-bottom:8px\\"><strong style=\\"font-size:13px\\">" + label + ":</strong> <span style=\\"font-size:13px\\">" + r[j] + "</span></div>";' +
    '    }' +
    '  }' +
    '  document.getElementById("responseCard").innerHTML = html;' +
    '}' +
    'function next() { if (idx < responses.length - 1) { idx++; renderResponse(); } }' +
    'function prev() { if (idx > 0) { idx--; renderResponse(); } }' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalGetSurveyResponses() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!sheet) return { headers: [], responses: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], responses: [] };

    var headers = data[0].map(function(h) {
      // Shorten long question headers to first 40 chars, then escape for safe HTML rendering
      var s = String(h).trim();
      s = s.length > 40 ? s.substring(0, 37) + '...' : s;
      return escapeHtml(s);
    });

    // Escape all response cell values for safe HTML rendering
    var responses = data.slice(1).map(function(row) {
      return row.map(function(cell) { return escapeHtml(String(cell || '')); });
    });

    return { headers: headers, responses: responses };
  } catch (e) {
    Logger.log('modalGetSurveyResponses error: ' + e.message);
    return { headers: [], responses: [] };
  }
}

// ============================================================================
// MODAL 9: WELCOME / SETUP WIZARD
// ============================================================================

function showWelcomeWizardModal() {
  var html = HtmlService.createHtmlOutput(getWelcomeWizardHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '🏗️ Welcome — Setup Wizard');
}

function getWelcomeWizardHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>' +
    '.wizard-steps { display:flex; gap:4px; margin-bottom:20px; } .step-dot { flex:1; height:6px; border-radius:3px; background:#E5E7EB; } .step-dot.active { background:#D4A017; } .step-dot.done { background:#276749; }' +
    '.step-content { min-height: 250px; } .step-title { font-size: 16px; font-weight: 700; color: #1A2A4A; margin-bottom: 12px; }' +
    '.step-desc { font-size: 14px; color: #4B5563; line-height: 1.6; margin-bottom: 16px; }' +
    '.check-item { display:flex; align-items:center; padding:10px; background:white; border-radius:6px; margin-bottom:8px; border:1px solid #E5E7EB; }' +
    '.check-item input { width:auto; margin-right:12px; transform:scale(1.3); }' +
    '</style></head><body>' +
    '<div class="wizard-steps" id="stepDots"></div>' +
    '<div class="step-content" id="stepContent"></div>' +
    '<div class="actions">' +
    '<button class="btn btn-secondary" id="prevBtn" onclick="prevStep()" style="display:none">← Back</button>' +
    '<button class="btn btn-secondary" id="skipBtn" onclick="google.script.host.close()">Skip for now</button>' +
    '<button class="btn btn-primary" id="nextBtn" onclick="nextStep()">Next →</button>' +
    '</div>' +
    '<script>' +
    'var step = 0, totalSteps = 4;' +
    'var steps = [' +
    '  { title: "Welcome to Your Union Dashboard!", desc: "This wizard will guide you through the essential setup steps. You can always come back to this wizard from the Admin menu.<br><br><strong>What you will configure:</strong><br>1. Organization details<br>2. Steward setup<br>3. Key features<br>4. Final checks" },' +
    '  { title: "Step 1: Organization Setup", desc: "Open the <strong>Config</strong> tab and fill in:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Organization Name</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Local Number</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Time Zone</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Contact Email</div>" },' +
    '  { title: "Step 2: Add Your First Members", desc: "Open <strong>Member Directory</strong> and add at least one member:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Add yourself as the first steward</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Import or manually add members</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Assign steward roles</div>" },' +
    '  { title: "Step 3: Explore Key Features", desc: "Try these essential features:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Open Union Hub menu — explore Search & Members</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Create a test grievance case</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Check the Getting Started tab for more guidance</div>" },' +
    '  { title: "Setup Complete! 🎉", desc: "You are ready to start using your Union Dashboard!<br><br><strong>Quick links:</strong><br>• 📊 Union Hub — Daily operations<br>• 🔧 Tools — Calendar, drive, notifications<br>• 🛠️ Admin — System management<br>• ❓ FAQ — Common questions<br><br>You can re-run this wizard anytime from Admin menu." }' +
    '];' +
    'totalSteps = steps.length;' +
    'function render() {' +
    '  var dots = "";' +
    '  for (var i = 0; i < totalSteps; i++) { dots += "<div class=\\"step-dot " + (i < step ? "done" : (i === step ? "active" : "")) + "\\"></div>"; }' +
    '  document.getElementById("stepDots").innerHTML = dots;' +
    '  document.getElementById("stepContent").innerHTML = "<div class=\\"step-title\\">" + steps[step].title + "</div><div class=\\"step-desc\\">" + steps[step].desc + "</div>";' +
    '  document.getElementById("prevBtn").style.display = step > 0 ? "inline-block" : "none";' +
    '  document.getElementById("nextBtn").textContent = step === totalSteps - 1 ? "Finish ✅" : "Next →";' +
    '  document.getElementById("skipBtn").style.display = step === totalSteps - 1 ? "none" : "inline-block";' +
    '}' +
    'function nextStep() { if (step < totalSteps - 1) { step++; render(); } else { google.script.host.close(); } }' +
    'function prevStep() { if (step > 0) { step--; render(); } }' +
    'render();' +
    '</script></body></html>';
}

// ============================================================================
// MODAL 10: SEARCHABLE HELP GUIDE
// ============================================================================

function showSearchableHelpModal() {
  var html = HtmlService.createHtmlOutput(getSearchableHelpHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '❓ Help & Documentation');
}

function getSearchableHelpHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.result { padding:12px; border-bottom:1px solid #E5E7EB; cursor:pointer; } .result:hover { background:#EBF4FF; } ' +
    '.result-source { font-size:11px; color:#6B7280; } .result-title { font-weight:600; } .result-text { font-size:13px; color:#4B5563; margin-top:4px; } ' +
    '.results-container { max-height:450px; overflow-y:auto; border:1px solid #E5E7EB; border-radius:6px; }</style>' +
    '</head><body>' +
    '<div class="search-box"><span class="search-icon">🔍</span><input id="searchInput" placeholder="Search help articles, FAQ, features..." oninput="doSearch(this.value)"></div>' +
    '<div class="tabs" id="tabs">' +
    '<div class="tab active" onclick="switchTab(\'all\')">All</div>' +
    '<div class="tab" onclick="switchTab(\'faq\')">FAQ</div>' +
    '<div class="tab" onclick="switchTab(\'features\')">Features</div>' +
    '<div class="tab" onclick="switchTab(\'tips\')">Quick Tips</div></div>' +
    '<div class="results-container" id="results"><div class="msg msg-info">Loading help content...</div></div>' +
    '<script>' +
    'var allItems = [], activeTab = "all";' +
    'google.script.run.withSuccessHandler(function(data) {' +
    '  allItems = data || [];' +
    '  doSearch("");' +
    '}).modalGetHelpContent();' +
    'function switchTab(tab) { activeTab = tab;' +
    '  var tabs = document.querySelectorAll(".tab"); tabs.forEach(function(t) { t.className = "tab"; });' +
    '  event.target.className = "tab active"; doSearch(document.getElementById("searchInput").value); }' +
    'function doSearch(query) {' +
    '  var q = query.toLowerCase().trim();' +
    '  var filtered = allItems.filter(function(item) {' +
    '    if (activeTab !== "all" && item.source !== activeTab) return false;' +
    '    if (!q) return true;' +
    '    return (item.title + " " + item.text).toLowerCase().indexOf(q) !== -1;' +
    '  });' +
    '  var html = "";' +
    '  if (filtered.length === 0) { html = "<div class=\\"msg msg-info\\">No results found.</div>"; }' +
    '  for (var i = 0; i < Math.min(filtered.length, 50); i++) {' +
    '    var it = filtered[i];' +
    '    var badge = it.source === "faq" ? "❓ FAQ" : (it.source === "features" ? "📋 Feature" : "💡 Tip");' +
    '    html += "<div class=\\"result\\"><span class=\\"result-source\\">" + badge + "</span><div class=\\"result-title\\">" + it.title + "</div><div class=\\"result-text\\">" + (it.text.length > 150 ? it.text.substring(0, 147) + "..." : it.text) + "</div></div>";' +
    '  }' +
    '  document.getElementById("results").innerHTML = html;' +
    '}' +
    '</script></body></html>';
}

function modalGetHelpContent() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var items = [];

    // Pull from FAQ sheet
    var faqSheet = ss.getSheetByName(SHEETS.FAQ);
    if (faqSheet) {
      var faqData = faqSheet.getDataRange().getValues();
      for (var r = 1; r < faqData.length; r++) {
        var text = String(faqData[r][0] || '').trim();
        if (text && text.indexOf('?') !== -1) {
          var answer = (r + 1 < faqData.length) ? String(faqData[r + 1][0] || '').trim() : '';
          items.push({ source: 'faq', title: escapeHtml(text), text: escapeHtml(answer) });
        }
      }
    }

    // Pull from Features Reference sheet
    var featSheet = ss.getSheetByName(SHEETS.FEATURES_REFERENCE);
    if (featSheet) {
      var featData = featSheet.getDataRange().getValues();
      for (var f = 1; f < featData.length; f++) {
        var name = String(featData[f][1] || '').trim();
        var desc = String(featData[f][2] || '').trim();
        if (name && desc) {
          items.push({ source: 'features', title: escapeHtml(name), text: escapeHtml(desc) });
        }
      }
    }

    // Add quick tips
    var tips = [
      { title: 'Keyboard Shortcut: Quick Search', text: 'Press Ctrl+F (Cmd+F on Mac) to search any sheet.' },
      { title: 'Traffic Light Colors', text: 'Red = overdue, Orange = due soon (1-3 days), Green = on track.' },
      { title: 'Bulk Actions', text: 'Select multiple grievance rows, then use Union Hub > Grievances > Bulk Update.' },
      { title: 'Mobile Access', text: 'The web dashboard works on phones and tablets — share the URL with members.' },
      { title: 'Data Backup', text: 'The system automatically creates audit logs. Check Admin > Security > Audit Log.' }
    ];
    for (var t = 0; t < tips.length; t++) {
      items.push({ source: 'tips', title: tips[t].title, text: tips[t].text });
    }

    return items;
  } catch (e) {
    Logger.log('modalGetHelpContent error: ' + e.message);
    return [];
  }
}

// ============================================================================
// MODAL 11: QUESTION EDITOR (Survey Questions)
// ============================================================================

function showQuestionEditorModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  var questionRow = row;

  // Only works from Survey Questions sheet
  if (sheet.getName() !== SHEETS.SURVEY_QUESTIONS) {
    SpreadsheetApp.getUi().alert('Please select a question row in the Survey Questions sheet first.');
    return;
  }

  var html = HtmlService.createHtmlOutput(getQuestionEditorHtml_(questionRow))
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Edit Survey Question');
}

function getQuestionEditorHtml_(row) {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Edit Survey Question</h2>' +
    '<div id="msg"></div>' +
    '<div id="loading" class="msg msg-info">Loading question...</div>' +
    '<div id="form" style="display:none">' +
    '<div class="row"><div class="form-group"><label>Question ID</label><input id="qId" readonly></div>' +
    '<div class="form-group"><label>Section</label><input id="section"></div></div>' +
    '<div class="form-group"><label>Section Title</label><input id="sectionTitle"></div>' +
    '<div class="form-group"><label>Question Text</label><textarea id="questionText" rows="3"></textarea></div>' +
    '<div class="row"><div class="form-group"><label>Type</label>' +
    '<select id="qType"><option>dropdown</option><option>slider-10</option><option>radio</option><option>paragraph</option></select></div>' +
    '<div class="form-group"><label>Required</label><select id="required"><option>Y</option><option>N</option></select></div>' +
    '<div class="form-group"><label>Active</label><select id="active"><option>Y</option><option>N</option></select></div></div>' +
    '<div class="form-group"><label>Options (pipe-separated: Option A|Option B|Option C)</label><textarea id="options" rows="2"></textarea></div>' +
    '<div class="form-group"><label>Branch Parent (optional)</label><input id="branchParent"></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="saveQuestion()">Save Changes</button></div></div>' +
    '<script>' +
    'var ROW = ' + row + ';' +
    'google.script.run.withSuccessHandler(function(q) {' +
    '  document.getElementById("loading").style.display = "none";' +
    '  if (!q) { showMsg("No question data found at row " + ROW, "error"); return; }' +
    '  document.getElementById("form").style.display = "block";' +
    '  document.getElementById("qId").value = q.id || "";' +
    '  document.getElementById("section").value = q.section || "";' +
    '  document.getElementById("sectionTitle").value = q.sectionTitle || "";' +
    '  document.getElementById("questionText").value = q.questionText || "";' +
    '  document.getElementById("qType").value = q.type || "dropdown";' +
    '  document.getElementById("required").value = q.required || "Y";' +
    '  document.getElementById("active").value = q.active || "Y";' +
    '  document.getElementById("options").value = q.options || "";' +
    '  document.getElementById("branchParent").value = q.branchParent || "";' +
    '}).modalGetQuestion(ROW);' +
    'function saveQuestion() {' +
    '  var data = { row: ROW, section: document.getElementById("section").value,' +
    '    sectionTitle: document.getElementById("sectionTitle").value,' +
    '    questionText: document.getElementById("questionText").value,' +
    '    type: document.getElementById("qType").value,' +
    '    required: document.getElementById("required").value,' +
    '    active: document.getElementById("active").value,' +
    '    options: document.getElementById("options").value,' +
    '    branchParent: document.getElementById("branchParent").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Question saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else showMsg(r && r.error || "Failed", "error");' +
    '  }).modalSaveQuestion(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalGetQuestion(row) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    if (!sheet || row < 2) return null;

    var data = sheet.getRange(row, 1, 1, 10).getValues()[0];
    return {
      id: data[0], section: data[1], sectionKey: data[2], sectionTitle: data[3],
      questionText: data[4], type: data[5], required: data[6], active: data[7],
      options: data[8], branchParent: data[9]
    };
  } catch (_e) {
    return null;
  }
}

function modalSaveQuestion(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    if (!sheet) return { success: false, error: 'Survey Questions sheet not found' };

    // Write back edited fields (skip ID/col 1 and sectionKey/col 3 — auto-generated)
    sheet.getRange(data.row, 2).setValue(escapeForFormula(data.section));
    sheet.getRange(data.row, 4).setValue(escapeForFormula(data.sectionTitle));
    sheet.getRange(data.row, 5).setValue(escapeForFormula(data.questionText));
    sheet.getRange(data.row, 6).setValue(escapeForFormula(data.type));
    sheet.getRange(data.row, 7).setValue(escapeForFormula(data.required));
    sheet.getRange(data.row, 8).setValue(escapeForFormula(data.active));
    sheet.getRange(data.row, 9).setValue(escapeForFormula(data.options));
    sheet.getRange(data.row, 10).setValue(escapeForFormula(data.branchParent));

    return { success: true };
  } catch (e) {
    Logger.log('modalSaveQuestion error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 12: CREATE WEEKLY QUESTION (replaces Flash Polls)
// ============================================================================

function showCreateWeeklyQuestionModal() {
  var html = HtmlService.createHtmlOutput(getCreateWeeklyQuestionHtml_())
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.MEDIUM.height);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Create Weekly Question');
}

function getCreateWeeklyQuestionHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.option-row { display:flex; gap:8px; margin-bottom:8px; align-items:center; } .option-row input { flex:1; } .remove-btn { background:#FEE2E2; color:#991B1B; border:none; border-radius:50%; width:28px; height:28px; cursor:pointer; font-size:16px; }</style>' +
    '</head><body>' +
    '<h2>Create Weekly Question</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Question</label><textarea id="question" rows="3" placeholder="What is your biggest workplace concern this week?"></textarea></div>' +
    '<div class="form-group"><label>Answer Options</label><div id="options"></div>' +
    '<button class="btn btn-secondary" style="margin-top:8px;font-size:13px" onclick="addOption()">+ Add Option</button></div>' +
    '<div class="row"><div class="form-group"><label>Active</label><select id="active"><option value="Y">Yes — show to members</option><option value="N">No — save as draft</option></select></div></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitQuestion()">Create Question</button></div>' +
    '<script>' +
    'var optCount = 0;' +
    'function addOption() {' +
    '  optCount++; var div = document.createElement("div"); div.className = "option-row"; div.id = "opt" + optCount;' +
    '  div.innerHTML = "<input placeholder=\\"Option " + optCount + "\\" id=\\"optVal" + optCount + "\\"><button class=\\"remove-btn\\" onclick=\\"this.parentElement.remove()\\">×</button>";' +
    '  document.getElementById("options").appendChild(div);' +
    '}' +
    'addOption(); addOption(); addOption();' +
    'function submitQuestion() {' +
    '  var q = document.getElementById("question").value.trim();' +
    '  if (!q) { showMsg("Question text is required.", "error"); return; }' +
    '  var opts = []; var optEls = document.querySelectorAll("[id^=optVal]");' +
    '  for (var i = 0; i < optEls.length; i++) { var v = optEls[i].value.trim(); if (v) opts.push(v); }' +
    '  if (opts.length < 2) { showMsg("At least 2 options are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Weekly question created!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalCreateWeeklyQuestion(q, opts, document.getElementById("active").value);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalCreateWeeklyQuestion(question, options, active) {
  try {
    // Delegate to WeeklyQuestions module if available
    if (typeof WeeklyQuestions !== 'undefined' && typeof WeeklyQuestions.addStewardQuestion === 'function') {
      WeeklyQuestions.addStewardQuestion(question, options.join('|'));
      return { success: true };
    }

    // Fallback: write directly to _Weekly_Questions sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.WEEKLY_QUESTIONS);
    if (!sheet) return { success: false, error: 'Weekly Questions sheet not found. Run Admin > Initial Setup first.' };

    var id = 'WQ-' + Date.now();
    var user = Session.getActiveUser().getEmail() || 'Unknown';

    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(question),
      escapeForFormula(options.join('|')),
      escapeForFormula(active),
      escapeForFormula(user),
      new Date()
    ]);

    return { success: true };
  } catch (e) {
    Logger.log('modalCreateWeeklyQuestion error: ' + e.message);
    return { success: false, error: e.message };
  }
}
