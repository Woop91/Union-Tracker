/**
 * 23_PortalSheets.gs — Portal sheet infrastructure with 0-indexed column constants
 *
 * WHAT THIS FILE DOES:
 *   Sheet setup infrastructure for the Member & Steward Portal. Creates 6
 *   portal-specific sheets (PortalMemberDirectory, Events, MeetingMinutes,
 *   PortalGrievances, StewardLog, MegaSurvey) and defines their column
 *   constants (PORTAL_*_COLS). Sheets with PII are hidden via
 *   setSheetVeryHidden_().
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Portal sheets are SEPARATE from the core sheets (Member Directory,
 *   Grievance Log) to avoid schema collision. Portal columns use 0-INDEXED
 *   constants (unlike the rest of the codebase which uses 1-indexed) because
 *   portal modules work exclusively with getValues() arrays (0-indexed)
 *   rather than getRange() calls (1-indexed). This is a critical difference
 *   documented at the top of the file — using PORTAL_*_COLS with getRange()
 *   will cause off-by-one errors.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Portal sheets can't be created. Events, meeting minutes, and steward
 *   interaction logs stop working. If the column constants are wrong, data
 *   appears in incorrect columns. The 0-indexed vs 1-indexed confusion is
 *   the #1 source of bugs in portal code (see COLUMN_ISSUES_LOG.md).
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS).
 *   Used by: 24_WeeklyQuestions.gs, 25_WorkloadService.gs, 26_QAForum.gs,
 *            27_TimelineService.gs, and SPA views.
 *
 * Note: FlashPolls/PollResponses removed v4.24.0 — replaced by 24_WeeklyQuestions.gs.
 */

// ═══════════════════════════════════════════════════════
// COLUMN CONSTANTS (PORTAL_ prefix to avoid global conflicts)
// ═══════════════════════════════════════════════════════
//
// *** IMPORTANT: 0-INDEXED CONSTANTS ***
// Unlike the rest of the codebase (GRIEVANCE_COLS, MEMBER_COLS, CONFIG_COLS)
// which use 1-indexed values for getRange(), the PORTAL_*_COLS below are
// 0-indexed for direct array access on getValues() data.
//
// Usage:  row[PORTAL_EVENT_COLS.TITLE]          — correct (array access)
// Do NOT: sheet.getRange(r, PORTAL_EVENT_COLS.TITLE) — wrong (off by one)
//
// This convention exists because portal modules exclusively work with
// getValues() arrays rather than individual getRange() calls.
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
  CREATED_BY: 5, CREATED_DATE: 6,
  DRIVE_DOC_URL: 7  // v4.20.18 — Google Doc URL saved to Minutes/ Drive folder
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
// PORTAL SHEET NAME HELPERS
// Use SHEETS.* constants from 01_Core.gs when available;
// fall back to hardcoded names if SHEETS is not yet loaded.
// ═══════════════════════════════════════════════════════

var PORTAL_SHEET_NAMES_ = {
  MEMBER_DIR:      (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_MEMBER_DIR)     ? SHEETS.PORTAL_MEMBER_DIR     : 'PortalMemberDirectory',
  EVENTS:          (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_EVENTS)         ? SHEETS.PORTAL_EVENTS         : 'Events',
  MINUTES:         (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_MINUTES)        ? SHEETS.PORTAL_MINUTES        : 'MeetingMinutes',
  // POLLS / POLL_RESPONSES removed v4.24.0 — replaced by _Weekly_Questions/_Weekly_Responses in 24_WeeklyQuestions.gs
  GRIEVANCES:      (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_GRIEVANCES)     ? SHEETS.PORTAL_GRIEVANCES     : 'PortalGrievances',
  STEWARD_LOG:     (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_STEWARD_LOG)    ? SHEETS.PORTAL_STEWARD_LOG    : 'StewardLog',
  MEGA_SURVEY:     (typeof SHEETS !== 'undefined' && SHEETS.PORTAL_MEGA_SURVEY)    ? SHEETS.PORTAL_MEGA_SURVEY    : 'MegaSurvey'
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
  if (!ss) {
    throw new Error('portalGetOrCreateSheet_: getActiveSpreadsheet() returned null for sheet "' + name + '"');
  }
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground(SHEET_COLORS.HEADER_NAVY)
      .setFontColor(SHEET_COLORS.BG_WHITE);
    sheet.setFrozenRows(1);
    if (hidden) sheet.hideSheet();
  }
  return sheet;
}

function getOrCreatePortalMemberDirectory() {
  return portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.MEMBER_DIR, [
    'Email', 'Name', 'Personal Email', 'Phone', 'Role',
    'Office', 'Unit', 'Join Date', 'Office Days', 'Badges',
    'Grievance Mode', 'Last Active'
  ], true);
}

function getOrCreateEventsSheet() {
  return portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.EVENTS, [
    'ID', 'Title', 'Type', 'DateTime', 'EndTime',
    'Location', 'Description', 'ZoomLink', 'CreatedBy', 'CreatedDate'
  ], false);
}

function getOrCreateMinutesSheet() {
  var sheet = portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.MINUTES, [
    'ID', 'MeetingDate', 'Title', 'Bullets', 'FullMinutes',
    'CreatedBy', 'CreatedDate', 'DriveDocUrl'  // col 8 — v4.20.18 fix: matches PORTAL_MINUTES_COLS.DRIVE_DOC_URL=7
  ], false);

  // Migration (v4.20.18): existing sheets created before v4.20.18 have only 7 headers.
  // Add DriveDocUrl header to col H if missing — safe to re-run.
  try {
    var headerRange = sheet.getRange(1, PORTAL_MINUTES_COLS.DRIVE_DOC_URL + 1); // col 8
    if (!headerRange.getValue()) {
      headerRange.setValue('DriveDocUrl')
        .setFontWeight('bold')
        .setBackground(SHEET_COLORS.HEADER_NAVY)
        .setFontColor(SHEET_COLORS.BG_WHITE);
      Logger.log('getOrCreateMinutesSheet: migrated — added DriveDocUrl header to col 8');
    }
  } catch (migErr) {
    Logger.log('getOrCreateMinutesSheet: migration check failed (non-fatal): ' + migErr.message);
  }

  return sheet;
}

// getOrCreatePollsSheet + getOrCreatePollResponsesSheet removed v4.24.0 — FlashPolls replaced by 24_WeeklyQuestions.gs

function getOrCreatePortalGrievanceSheet() {
  return portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.GRIEVANCES, [
    'ID', 'MemberEmail', 'StewardEmail', 'Step', 'Status',
    'Description', 'GDriveLink', 'ZoomLink', 'Deadlines',
    'Notes', 'CreatedDate', 'UpdatedDate'
  ], true);
}

function getOrCreateStewardLogSheet() {
  return portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.STEWARD_LOG, [
    'ID', 'StewardEmail', 'MemberEmail', 'Type',
    'Notes', 'Office', 'Timestamp'
  ], true);
}

function getOrCreateMegaSurveySheet() {
  return portalGetOrCreateSheet_(PORTAL_SHEET_NAMES_.MEGA_SURVEY, [
    'Email', 'Responses', 'Progress', 'LastUpdated', 'Completed'
  ], true);
}

/**
 * Creates all 6 portal sheets (FlashPolls/PollResponses removed v4.24.0).
 * Called from CREATE_DASHBOARD() in 08a_SheetSetup.gs, and directly via initPortalSheetsManual().
 */
function initPortalSheets() {
  getOrCreatePortalMemberDirectory();
  getOrCreateEventsSheet();
  getOrCreateMinutesSheet();
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
  SpreadsheetApp.getUi().alert('Portal sheets created successfully.\n\nSheets: PortalMemberDirectory, Events, MeetingMinutes, PortalGrievances, StewardLog, MegaSurvey\n\nNote: FlashPolls/PollResponses removed — polls now in _Weekly_Questions/_Weekly_Responses.');
}
