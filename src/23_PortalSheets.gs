/**
 * 23_PortalSheets.gs
 * Sheet setup infrastructure for the Member & Steward Portal.
 * Creates the 8 portal-specific sheets and defines their column constants.
 *
 * Sheets created:
 *   PortalMemberDirectory (hidden) — portal profile extensions (office days, badges, etc.)
 *   Events               — upcoming union events and meetings
 *   MeetingMinutes       — union meeting notes with bullet summaries
 *   FlashPolls           — active quick-vote polls
 *   PollResponses (hidden) — individual poll vote records
 *   PortalGrievances (hidden) — portal grievance cases (separate from Grievance Log)
 *   StewardLog (hidden)  — steward-member interaction audit trail
 *   MegaSurvey (hidden)  — 56-question survey progress
 *
 * Note: Sheet names PortalMemberDirectory and PortalGrievances are distinct from
 * DDS's existing "Member Directory" and "Grievance Log" sheets to avoid schema collision.
 */

// ═══════════════════════════════════════════════════════
// COLUMN CONSTANTS (PORTAL_ prefix to avoid global conflicts)
// ═══════════════════════════════════════════════════════

var PORTAL_MEMBER_DIR_COLS = {
  EMAIL: 0, NAME: 1, PERSONAL_EMAIL: 2, PHONE: 3, ROLE: 4,
  OFFICE: 5, UNIT: 6, JOIN_DATE: 7, OFFICE_DAYS: 8, BADGES: 9,
  GRIEVANCE_MODE: 10, LAST_ACTIVE: 11
};

var PORTAL_EVENT_COLS = {
  ID: 0, TITLE: 1, TYPE: 2, DATE_TIME: 3, END_TIME: 4,
  LOCATION: 5, DESCRIPTION: 6, ZOOM_LINK: 7, CREATED_BY: 8, CREATED_DATE: 9
};

var PORTAL_MINUTES_COLS = {
  ID: 0, MEETING_DATE: 1, TITLE: 2, BULLETS: 3, FULL_MINUTES: 4,
  CREATED_BY: 5, CREATED_DATE: 6
};

var PORTAL_POLL_COLS = {
  ID: 0, QUESTION: 1, OPTIONS: 2, ACTIVE: 3, UNIT: 4,
  CREATED_BY: 5, CREATED_DATE: 6
};

var PORTAL_POLL_RESPONSE_COLS = {
  POLL_ID: 0, EMAIL: 1, RESPONSE: 2, TIMESTAMP: 3
};

var PORTAL_GRIEVANCE_COLS = {
  ID: 0, MEMBER_EMAIL: 1, STEWARD_EMAIL: 2, STEP: 3, STATUS: 4,
  DESCRIPTION: 5, GDRIVE_LINK: 6, ZOOM_LINK: 7, DEADLINES: 8,
  NOTES: 9, CREATED_DATE: 10, UPDATED_DATE: 11
};

var PORTAL_STEWARD_LOG_COLS = {
  ID: 0, STEWARD_EMAIL: 1, MEMBER_EMAIL: 2, TYPE: 3,
  NOTES: 4, OFFICE: 5, TIMESTAMP: 6
};

var PORTAL_MEGA_SURVEY_COLS = {
  EMAIL: 0, RESPONSES: 1, PROGRESS: 2, LAST_UPDATED: 3, COMPLETED: 4
};

// ═══════════════════════════════════════════════════════
// SHEET SETUP FUNCTIONS
// ═══════════════════════════════════════════════════════

/**
 * Internal helper to get or create a sheet with headers.
 * Named portalGetOrCreateSheet_ to avoid conflict with
 * 08a_SheetSetup.gs's getOrCreateSheet(ss, name).
 */
function portalGetOrCreateSheet_(name, headers, hidden) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    if (hidden) sheet.hideSheet();
  }
  return sheet;
}

function getOrCreatePortalMemberDirectory() {
  return portalGetOrCreateSheet_('PortalMemberDirectory', [
    'Email', 'Name', 'Personal Email', 'Phone', 'Role',
    'Office', 'Unit', 'Join Date', 'Office Days', 'Badges',
    'Grievance Mode', 'Last Active'
  ], true);
}

function getOrCreateEventsSheet() {
  return portalGetOrCreateSheet_('Events', [
    'ID', 'Title', 'Type', 'DateTime', 'EndTime',
    'Location', 'Description', 'ZoomLink', 'CreatedBy', 'CreatedDate'
  ], false);
}

function getOrCreateMinutesSheet() {
  return portalGetOrCreateSheet_('MeetingMinutes', [
    'ID', 'MeetingDate', 'Title', 'Bullets', 'FullMinutes',
    'CreatedBy', 'CreatedDate'
  ], false);
}

function getOrCreatePollsSheet() {
  return portalGetOrCreateSheet_('FlashPolls', [
    'ID', 'Question', 'Options', 'Active', 'Unit',
    'CreatedBy', 'CreatedDate'
  ], false);
}

function getOrCreatePollResponsesSheet() {
  return portalGetOrCreateSheet_('PollResponses', [
    'PollID', 'Email', 'Response', 'Timestamp'
  ], true);
}

function getOrCreatePortalGrievanceSheet() {
  return portalGetOrCreateSheet_('PortalGrievances', [
    'ID', 'MemberEmail', 'StewardEmail', 'Step', 'Status',
    'Description', 'GDriveLink', 'ZoomLink', 'Deadlines',
    'Notes', 'CreatedDate', 'UpdatedDate'
  ], true);
}

function getOrCreateStewardLogSheet() {
  return portalGetOrCreateSheet_('StewardLog', [
    'ID', 'StewardEmail', 'MemberEmail', 'Type',
    'Notes', 'Office', 'Timestamp'
  ], true);
}

function getOrCreateMegaSurveySheet() {
  return portalGetOrCreateSheet_('MegaSurvey', [
    'Email', 'Responses', 'Progress', 'LastUpdated', 'Completed'
  ], true);
}

/**
 * Creates all 8 portal sheets. Called from CREATE_DASHBOARD()
 * in 08a_SheetSetup.gs, and directly via initPortalSheetsManual().
 */
function initPortalSheets() {
  getOrCreatePortalMemberDirectory();
  getOrCreateEventsSheet();
  getOrCreateMinutesSheet();
  getOrCreatePollsSheet();
  getOrCreatePollResponsesSheet();
  getOrCreatePortalGrievanceSheet();
  getOrCreateStewardLogSheet();
  getOrCreateMegaSurveySheet();
  Logger.log('Portal sheets initialized.');
}

/**
 * Standalone entry point for existing deployments.
 * Run from the GAS editor when CREATE_DASHBOARD has already been executed.
 */
function initPortalSheetsManual() {
  initPortalSheets();
  SpreadsheetApp.getUi().alert('Portal sheets created successfully.\n\nNew sheets: PortalMemberDirectory, Events, MeetingMinutes, FlashPolls, PollResponses, PortalGrievances, StewardLog, MegaSurvey');
}
