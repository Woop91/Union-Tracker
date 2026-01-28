/**
 * ============================================================================
 * LOOKER STUDIO INTEGRATION
 * ============================================================================
 * Provides a read-only data layer for Google Looker Studio (Data Studio).
 * Creates optimized hidden sheets that Looker can connect to as data sources.
 *
 * RESTRICTED DATA SOURCES:
 * This integration ONLY exports data from:
 * - Member Directory → _Looker_Members
 * - Grievance Log → _Looker_Grievances
 * - Member Satisfaction → _Looker_Satisfaction
 *
 * ARCHITECTURE:
 * - Does NOT modify existing sheets or modal dashboards
 * - Creates separate hidden "_Looker_*" sheets for data export
 * - Looker connects to these sheets as native Google Sheets data sources
 * - Data is refreshed on-demand or via scheduled trigger
 *
 * USAGE:
 * 1. Run setupLookerIntegration() once to create sheets
 * 2. Run refreshLookerData() to update data (or install trigger)
 * 3. Connect Looker Studio to the spreadsheet, select _Looker_* sheets
 */

const LOOKER_CONFIG = {
  // Sheet names for Looker data sources (restricted to 3 source sheets)
  SHEETS: {
    GRIEVANCES: '_Looker_Grievances',
    MEMBERS: '_Looker_Members',
    SATISFACTION: '_Looker_Satisfaction'
  },
  // PII-FREE sheets for external/compliance use (no personally identifiable information)
  SHEETS_ANON: {
    GRIEVANCES: '_Looker_Anon_Grievances',
    MEMBERS: '_Looker_Anon_Members',
    SATISFACTION: '_Looker_Anon_Satisfaction'
  },
  // Source sheets (only these are allowed)
  ALLOWED_SOURCES: ['Member Directory', 'Grievance Log', 'Member Satisfaction'],
  // Refresh settings
  AUTO_REFRESH_HOUR: 6 // 6 AM daily refresh
};

// ============================================================================
// SETUP & MANAGEMENT
// ============================================================================

/**
 * Sets up Looker Studio integration by creating hidden data sheets.
 * Only creates sheets for allowed data sources.
 * @returns {Object} Result with created sheet names
 */
function setupLookerIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Looker Studio Integration',
    'This will create hidden data sheets for Looker Studio:\n\n' +
    '• _Looker_Members - From Member Directory\n' +
    '• _Looker_Grievances - From Grievance Log\n' +
    '• _Looker_Satisfaction - From Member Satisfaction Survey\n\n' +
    'Only these 3 source sheets will be accessible to Looker.\n' +
    'Existing sheets will NOT be modified.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  const created = [];

  // Create each Looker sheet
  for (const key in LOOKER_CONFIG.SHEETS) {
    const sheetName = LOOKER_CONFIG.SHEETS[key];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.hideSheet();
      created.push(sheetName);
    }
  }

  // Initialize headers for each sheet
  initializeLookerGrievancesSheet_();
  initializeLookerMembersSheet_();
  initializeLookerSatisfactionSheet_();

  // Do initial data population
  refreshLookerData();

  const summary = created.length > 0
    ? 'Created sheets: ' + created.join(', ')
    : 'All Looker sheets already exist';

  ui.alert(
    'Setup Complete',
    summary + '\n\nData has been populated from:\n' +
    '• Member Directory\n• Grievance Log\n• Member Satisfaction\n\n' +
    'Connect Looker Studio to this spreadsheet and select the _Looker_* sheets as data sources.',
    ui.ButtonSet.OK
  );

  // Log setup
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_INITIALIZED, {
      action: 'LOOKER_SETUP',
      sheetsCreated: created,
      allowedSources: LOOKER_CONFIG.ALLOWED_SOURCES,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, created: created };
}

/**
 * Initializes the Looker Grievances sheet with headers.
 * @private
 */
function initializeLookerGrievancesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.GRIEVANCES);
  if (!sheet) return;

  const headers = [
    'Grievance ID', 'Member ID', 'Member Name', 'Status', 'Current Step',
    'Incident Date', 'Date Filed', 'Date Closed', 'Days Open',
    'Days to Deadline', 'Is Overdue', 'Issue Category', 'Articles',
    'Unit', 'Location', 'Assigned Steward', 'Resolution',
    'Action Type', 'Checklist Progress',
    'Has Reminder 1', 'Reminder 1 Date', 'Has Reminder 2', 'Reminder 2 Date',
    'Filed Month', 'Filed Quarter', 'Filed Year',
    'Closed Month', 'Closed Quarter', 'Closed Year',
    'Outcome Category', 'Time to Resolution Days',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the Looker Members sheet with headers.
 * @private
 */
function initializeLookerMembersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);
  if (!sheet) return;

  const headers = [
    'Member ID', 'Full Name', 'First Name', 'Last Name',
    'Job Title', 'Unit', 'Location',
    'Is Steward', 'Is Member Leader', 'Assigned Steward',
    'Email', 'Phone', 'Preferred Communication',
    'Has Open Grievance', 'Total Grievances', 'Grievances Won', 'Grievances Lost',
    'Last Contact Date', 'Days Since Contact',
    'Volunteer Hours', 'Last Virtual Meeting', 'Last In-Person Meeting',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the Looker Satisfaction sheet with headers.
 * @private
 */
function initializeLookerSatisfactionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SATISFACTION);
  if (!sheet) return;

  const headers = [
    'Response ID', 'Timestamp', 'Response Month', 'Response Quarter', 'Response Year',
    'Worksite', 'Role', 'Shift', 'Time in Role', 'Has Steward Contact',
    'Overall Satisfaction Avg', 'Steward Rating Avg', 'Steward Access Avg',
    'Chapter Effectiveness Avg', 'Leadership Avg', 'Contract Enforcement Avg',
    'Communication Avg', 'Member Voice Avg', 'Value Action Avg', 'Scheduling Avg',
    'Satisfied with Rep', 'Trust Union', 'Feel Protected', 'Would Recommend',
    'Filed Grievance', 'Representation Avg',
    'Verification Status', 'Matched Member ID', 'Quarter Period',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

// ============================================================================
// DATA REFRESH FUNCTIONS
// ============================================================================

/**
 * Refreshes all Looker data sheets.
 * Only pulls from allowed source sheets.
 * @returns {Object} Result with refresh counts
 */
function refreshLookerData() {
  const startTime = new Date();

  const grievanceCount = refreshLookerGrievances_();
  const memberCount = refreshLookerMembers_();
  const satisfactionCount = refreshLookerSatisfaction_();

  const duration = (new Date() - startTime) / 1000;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Members: ${memberCount}, Grievances: ${grievanceCount}, Survey: ${satisfactionCount}`,
    'Looker Data Refreshed',
    5
  );

  return {
    success: true,
    sources: LOOKER_CONFIG.ALLOWED_SOURCES,
    grievances: grievanceCount,
    members: memberCount,
    satisfaction: satisfactionCount,
    durationSeconds: duration
  };
}

/**
 * Refreshes the Looker Grievances sheet from Grievance Log.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerGrievances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.GRIEVANCES);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = GRIEVANCE_COLS;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const grievanceId = row[cols.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    const firstName = row[cols.FIRST_NAME - 1] || '';
    const lastName = row[cols.LAST_NAME - 1] || '';
    const status = row[cols.STATUS - 1] || '';
    const dateFiled = row[cols.DATE_FILED - 1];
    const dateClosed = row[cols.DATE_CLOSED - 1];
    const daysToDeadline = row[cols.DAYS_TO_DEADLINE - 1];
    const r1Date = row[cols.REMINDER_1_DATE - 1];
    const r2Date = row[cols.REMINDER_2_DATE - 1];

    // Compute derived fields
    const isOverdue = typeof daysToDeadline === 'number' && daysToDeadline < 0;
    const outcomeCategory = getOutcomeCategory_(status);
    const timeToResolution = (dateFiled instanceof Date && dateClosed instanceof Date)
      ? Math.ceil((dateClosed - dateFiled) / (1000 * 60 * 60 * 24))
      : '';

    // Date dimensions
    const filedMonth = dateFiled instanceof Date ? Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const filedQuarter = dateFiled instanceof Date ? getQuarter_(dateFiled) : '';
    const filedYear = dateFiled instanceof Date ? dateFiled.getFullYear() : '';
    const closedMonth = dateClosed instanceof Date ? Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const closedQuarter = dateClosed instanceof Date ? getQuarter_(dateClosed) : '';
    const closedYear = dateClosed instanceof Date ? dateClosed.getFullYear() : '';

    exportData.push([
      grievanceId,
      row[cols.MEMBER_ID - 1] || '',
      (firstName + ' ' + lastName).trim(),
      status,
      row[cols.CURRENT_STEP - 1] || '',
      row[cols.INCIDENT_DATE - 1] || '',
      dateFiled || '',
      dateClosed || '',
      row[cols.DAYS_OPEN - 1] || '',
      daysToDeadline || '',
      isOverdue ? 'Yes' : 'No',
      row[cols.ISSUE_CATEGORY - 1] || '',
      row[cols.ARTICLES - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.LOCATION - 1] || '',
      row[cols.STEWARD - 1] || '',
      row[cols.RESOLUTION - 1] || '',
      row[cols.ACTION_TYPE - 1] || 'Grievance',
      row[cols.CHECKLIST_PROGRESS - 1] || '',
      r1Date ? 'Yes' : 'No',
      r1Date || '',
      r2Date ? 'Yes' : 'No',
      r2Date || '',
      filedMonth,
      filedQuarter,
      filedYear,
      closedMonth,
      closedQuarter,
      closedYear,
      outcomeCategory,
      timeToResolution,
      now
    ]);
  }

  // Clear old data and write new
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes the Looker Members sheet from Member Directory.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerMembers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.MEMBER_DIR || 'Member Directory');
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);

  if (!sourceSheet || !targetSheet) return 0;

  const memberData = sourceSheet.getDataRange().getValues();
  if (memberData.length < 2) return 0;

  // Build grievance stats by member
  const grievanceStats = {};
  if (grievanceSheet) {
    const gData = grievanceSheet.getDataRange().getValues();
    for (let i = 1; i < gData.length; i++) {
      const memberId = gData[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      const status = gData[i][GRIEVANCE_COLS.STATUS - 1];
      if (!memberId) continue;

      if (!grievanceStats[memberId]) {
        grievanceStats[memberId] = { total: 0, won: 0, lost: 0, open: 0 };
      }
      grievanceStats[memberId].total++;
      if (status === 'Won') grievanceStats[memberId].won++;
      if (status === 'Denied') grievanceStats[memberId].lost++;
      if (status === 'Open' || status === 'Pending Info' || status === 'Appealed' || status === 'In Arbitration') {
        grievanceStats[memberId].open++;
      }
    }
  }

  const now = new Date();
  const exportData = [];
  const cols = MEMBER_COLS;

  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const memberId = row[cols.MEMBER_ID - 1];
    if (!memberId) continue;

    const firstName = row[cols.FIRST_NAME - 1] || '';
    const lastName = row[cols.LAST_NAME - 1] || '';
    const isSteward = row[cols.IS_STEWARD - 1];
    const lastContact = row[cols.RECENT_CONTACT_DATE - 1];
    const stats = grievanceStats[memberId] || { total: 0, won: 0, lost: 0, open: 0 };

    // Days since last contact
    const daysSinceContact = lastContact instanceof Date
      ? Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24))
      : '';

    exportData.push([
      memberId,
      (firstName + ' ' + lastName).trim(),
      firstName,
      lastName,
      row[cols.JOB_TITLE - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.WORK_LOCATION - 1] || '',
      isSteward === 'Yes' ? 'Yes' : 'No',
      isSteward === 'Member Leader' ? 'Yes' : 'No',
      row[cols.ASSIGNED_STEWARD - 1] || '',
      row[cols.EMAIL - 1] || '',
      row[cols.PHONE - 1] || '',
      row[cols.PREFERRED_COMM - 1] || '',
      stats.open > 0 ? 'Yes' : 'No',
      stats.total,
      stats.won,
      stats.lost,
      lastContact || '',
      daysSinceContact,
      parseFloat(row[cols.VOLUNTEER_HOURS - 1]) || 0,
      row[cols.LAST_VIRTUAL_MTG - 1] || '',
      row[cols.LAST_INPERSON_MTG - 1] || '',
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes the Looker Satisfaction sheet from Member Satisfaction survey.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerSatisfaction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.SATISFACTION || '📊 Member Satisfaction');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SATISFACTION);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = typeof SATISFACTION_COLS !== 'undefined' ? SATISFACTION_COLS : {};

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const timestamp = row[0]; // First column is always timestamp
    if (!timestamp) continue;

    // Generate response ID
    const responseId = 'SR' + String(i).padStart(5, '0');

    // Date dimensions
    const respMonth = timestamp instanceof Date ? Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const respQuarter = timestamp instanceof Date ? getQuarter_(timestamp) : '';
    const respYear = timestamp instanceof Date ? timestamp.getFullYear() : '';

    // Calculate section averages using SATISFACTION_COLS if available
    const overallSatAvg = calculateSectionAvg_(row, [6, 7, 8, 9]); // Q6-Q9
    const stewardRatingAvg = calculateSectionAvg_(row, [10, 11, 12, 13, 14, 15, 16]); // Q10-Q16
    const stewardAccessAvg = calculateSectionAvg_(row, [18, 19, 20]); // Q18-Q20
    const chapterAvg = calculateSectionAvg_(row, [21, 22, 23, 24, 25]); // Q21-Q25
    const leadershipAvg = calculateSectionAvg_(row, [26, 27, 28, 29, 30, 31]); // Q26-Q31
    const contractAvg = calculateSectionAvg_(row, [32, 33, 34, 35]); // Q32-Q35
    const commAvg = calculateSectionAvg_(row, [41, 42, 43, 44, 45]); // Q41-Q45
    const voiceAvg = calculateSectionAvg_(row, [46, 47, 48, 49, 50]); // Q46-Q50
    const valueAvg = calculateSectionAvg_(row, [51, 52, 53, 54, 55]); // Q51-Q55
    const schedAvg = calculateSectionAvg_(row, [56, 57, 58, 59, 60, 61, 62]); // Q56-Q62
    const repAvg = calculateSectionAvg_(row, [37, 38, 39, 40]); // Q37-Q40

    // Get individual key questions (0-indexed)
    const worksite = row[1] || '';
    const role = row[2] || '';
    const shift = row[3] || '';
    const timeInRole = row[4] || '';
    const hasStewardContact = row[5] || '';
    const satisfiedRep = row[6] || '';
    const trustUnion = row[7] || '';
    const feelProtected = row[8] || '';
    const wouldRecommend = row[9] || '';
    const filedGrievance = row[36] || '';

    // Verification columns (if they exist)
    const verificationStatus = cols.VERIFIED ? (row[cols.VERIFIED - 1] || '') : '';
    const matchedMemberId = cols.MATCHED_MEMBER_ID ? (row[cols.MATCHED_MEMBER_ID - 1] || '') : '';
    const quarterPeriod = cols.QUARTER ? (row[cols.QUARTER - 1] || '') : respQuarter;

    exportData.push([
      responseId,
      timestamp,
      respMonth,
      respQuarter,
      respYear,
      worksite,
      role,
      shift,
      timeInRole,
      hasStewardContact,
      overallSatAvg,
      stewardRatingAvg,
      stewardAccessAvg,
      chapterAvg,
      leadershipAvg,
      contractAvg,
      commAvg,
      voiceAvg,
      valueAvg,
      schedAvg,
      satisfiedRep,
      trustUnion,
      feelProtected,
      wouldRecommend,
      filedGrievance,
      repAvg,
      verificationStatus,
      matchedMemberId,
      quarterPeriod,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

// ============================================================================
// PII-FREE LOOKER INTEGRATION (Anonymized Data)
// ============================================================================
// For external stakeholders, compliance, or public-facing dashboards.
// Excludes: Names, Emails, Phones, Addresses, Member IDs

/**
 * Sets up PII-free Looker sheets for external/compliance use.
 * @returns {Object} Result with created sheet names
 */
function setupLookerAnonIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup PII-Free Looker Integration',
    'This will create ANONYMIZED data sheets for external Looker reports:\n\n' +
    '• _Looker_Anon_Members - Aggregated member data (no names/contact)\n' +
    '• _Looker_Anon_Grievances - Case data (no member info)\n' +
    '• _Looker_Anon_Satisfaction - Survey data (already anonymous)\n\n' +
    'These sheets contain NO personally identifiable information (PII).\n' +
    'Safe for external dashboards and compliance reporting.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  const created = [];

  // Create each anonymous Looker sheet
  for (const key in LOOKER_CONFIG.SHEETS_ANON) {
    const sheetName = LOOKER_CONFIG.SHEETS_ANON[key];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.hideSheet();
      created.push(sheetName);
    }
  }

  // Initialize headers for each sheet
  initializeLookerAnonGrievancesSheet_();
  initializeLookerAnonMembersSheet_();
  initializeLookerAnonSatisfactionSheet_();

  // Do initial data population
  refreshLookerAnonData();

  const summary = created.length > 0
    ? 'Created sheets: ' + created.join(', ')
    : 'All PII-free Looker sheets already exist';

  ui.alert(
    'PII-Free Setup Complete',
    summary + '\n\nThese sheets contain NO personally identifiable information:\n' +
    '• No names, emails, or phone numbers\n' +
    '• No member IDs (uses anonymous hashes)\n' +
    '• Safe for external/public dashboards',
    ui.ButtonSet.OK
  );

  // Log setup
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_INITIALIZED, {
      action: 'LOOKER_ANON_SETUP',
      sheetsCreated: created,
      piiExcluded: true,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, created: created, piiExcluded: true };
}

/**
 * Initializes the anonymized Grievances sheet headers.
 * @private
 */
function initializeLookerAnonGrievancesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.GRIEVANCES);
  if (!sheet) return;

  // No Member Name, No Steward Name - only aggregation-safe fields
  const headers = [
    'Case Hash', 'Status', 'Current Step',
    'Incident Month', 'Incident Quarter', 'Incident Year',
    'Filed Month', 'Filed Quarter', 'Filed Year',
    'Closed Month', 'Closed Quarter', 'Closed Year',
    'Days Open', 'Days to Deadline', 'Is Overdue',
    'Issue Category', 'Articles Violated',
    'Unit', 'Location', 'Action Type',
    'Outcome Category', 'Time to Resolution Days',
    'Has Reminder', 'Checklist Progress',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the anonymized Members sheet headers.
 * @private
 */
function initializeLookerAnonMembersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.MEMBERS);
  if (!sheet) return;

  // Aggregation-focused: No names, no contact info, no member IDs
  const headers = [
    'Member Hash', 'Job Title', 'Unit', 'Location',
    'Is Steward', 'Is Member Leader',
    'Has Open Grievance', 'Total Grievances', 'Grievances Won', 'Grievances Lost',
    'Days Since Contact Bucket', 'Contact Frequency Category',
    'Volunteer Hours Bucket', 'Engagement Level',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the anonymized Satisfaction sheet headers.
 * @private
 */
function initializeLookerAnonSatisfactionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.SATISFACTION);
  if (!sheet) return;

  // Survey data is already mostly anonymous, just remove any potential PII linkages
  const headers = [
    'Response Hash', 'Response Month', 'Response Quarter', 'Response Year',
    'Worksite', 'Role Category', 'Shift', 'Tenure Bucket',
    'Has Steward Contact',
    'Overall Satisfaction Avg', 'Steward Rating Avg', 'Steward Access Avg',
    'Chapter Effectiveness Avg', 'Leadership Avg', 'Contract Enforcement Avg',
    'Communication Avg', 'Member Voice Avg', 'Value Action Avg', 'Scheduling Avg',
    'Representation Avg',
    'Satisfied Bucket', 'Trust Bucket', 'Protected Bucket', 'Recommend Bucket',
    'Filed Grievance',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Refreshes all PII-free Looker data sheets.
 * @returns {Object} Result with refresh counts
 */
function refreshLookerAnonData() {
  const startTime = new Date();

  const grievanceCount = refreshLookerAnonGrievances_();
  const memberCount = refreshLookerAnonMembers_();
  const satisfactionCount = refreshLookerAnonSatisfaction_();

  const duration = (new Date() - startTime) / 1000;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Anon Members: ${memberCount}, Grievances: ${grievanceCount}, Survey: ${satisfactionCount}`,
    'PII-Free Looker Data Refreshed',
    5
  );

  return {
    success: true,
    piiExcluded: true,
    grievances: grievanceCount,
    members: memberCount,
    satisfaction: satisfactionCount,
    durationSeconds: duration
  };
}

/**
 * Refreshes anonymized Grievances sheet - NO member names or IDs.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonGrievances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.GRIEVANCES);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = GRIEVANCE_COLS;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const grievanceId = row[cols.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    const status = row[cols.STATUS - 1] || '';
    const incidentDate = row[cols.INCIDENT_DATE - 1];
    const dateFiled = row[cols.DATE_FILED - 1];
    const dateClosed = row[cols.DATE_CLOSED - 1];
    const daysToDeadline = row[cols.DAYS_TO_DEADLINE - 1];
    const r1Date = row[cols.REMINDER_1_DATE - 1];
    const r2Date = row[cols.REMINDER_2_DATE - 1];

    // Generate anonymous hash (not reversible)
    const caseHash = generateAnonHash_(grievanceId);

    // Compute derived fields
    const isOverdue = typeof daysToDeadline === 'number' && daysToDeadline < 0;
    const outcomeCategory = getOutcomeCategory_(status);
    const timeToResolution = (dateFiled instanceof Date && dateClosed instanceof Date)
      ? Math.ceil((dateClosed - dateFiled) / (1000 * 60 * 60 * 24))
      : '';

    // Date dimensions (incident)
    const incMonth = incidentDate instanceof Date ? Utilities.formatDate(incidentDate, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const incQuarter = incidentDate instanceof Date ? getQuarter_(incidentDate) : '';
    const incYear = incidentDate instanceof Date ? incidentDate.getFullYear() : '';

    // Date dimensions (filed)
    const filedMonth = dateFiled instanceof Date ? Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const filedQuarter = dateFiled instanceof Date ? getQuarter_(dateFiled) : '';
    const filedYear = dateFiled instanceof Date ? dateFiled.getFullYear() : '';

    // Date dimensions (closed)
    const closedMonth = dateClosed instanceof Date ? Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const closedQuarter = dateClosed instanceof Date ? getQuarter_(dateClosed) : '';
    const closedYear = dateClosed instanceof Date ? dateClosed.getFullYear() : '';

    exportData.push([
      caseHash,
      status,
      row[cols.CURRENT_STEP - 1] || '',
      incMonth,
      incQuarter,
      incYear,
      filedMonth,
      filedQuarter,
      filedYear,
      closedMonth,
      closedQuarter,
      closedYear,
      row[cols.DAYS_OPEN - 1] || '',
      daysToDeadline || '',
      isOverdue ? 'Yes' : 'No',
      row[cols.ISSUE_CATEGORY - 1] || '',
      row[cols.ARTICLES - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.LOCATION - 1] || '',
      row[cols.ACTION_TYPE - 1] || 'Grievance',
      outcomeCategory,
      timeToResolution,
      (r1Date || r2Date) ? 'Yes' : 'No',
      row[cols.CHECKLIST_PROGRESS - 1] || '',
      now
    ]);
  }

  // Clear old data and write new
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes anonymized Members sheet - NO names, emails, phones, or IDs.
 * Uses bucketed/categorized values for privacy.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonMembers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.MEMBER_DIR || 'Member Directory');
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.MEMBERS);

  if (!sourceSheet || !targetSheet) return 0;

  const memberData = sourceSheet.getDataRange().getValues();
  if (memberData.length < 2) return 0;

  // Build grievance stats by member
  const grievanceStats = {};
  if (grievanceSheet) {
    const gData = grievanceSheet.getDataRange().getValues();
    for (let i = 1; i < gData.length; i++) {
      const memberId = gData[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      const status = gData[i][GRIEVANCE_COLS.STATUS - 1];
      if (!memberId) continue;

      if (!grievanceStats[memberId]) {
        grievanceStats[memberId] = { total: 0, won: 0, lost: 0, open: 0 };
      }
      grievanceStats[memberId].total++;
      if (status === 'Won') grievanceStats[memberId].won++;
      if (status === 'Denied') grievanceStats[memberId].lost++;
      if (status === 'Open' || status === 'Pending Info' || status === 'Appealed' || status === 'In Arbitration') {
        grievanceStats[memberId].open++;
      }
    }
  }

  const now = new Date();
  const exportData = [];
  const cols = MEMBER_COLS;

  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const memberId = row[cols.MEMBER_ID - 1];
    if (!memberId) continue;

    const isSteward = row[cols.IS_STEWARD - 1];
    const lastContact = row[cols.RECENT_CONTACT_DATE - 1];
    const stats = grievanceStats[memberId] || { total: 0, won: 0, lost: 0, open: 0 };
    const volunteerHours = parseFloat(row[cols.VOLUNTEER_HOURS - 1]) || 0;

    // Generate anonymous hash
    const memberHash = generateAnonHash_(memberId);

    // Days since last contact - BUCKETED for privacy
    let daysSinceContactBucket = 'Unknown';
    if (lastContact instanceof Date) {
      const days = Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24));
      daysSinceContactBucket = getDaysBucket_(days);
    }

    // Contact frequency category
    const contactFreq = getContactFrequencyCategory_(lastContact, now);

    // Volunteer hours bucket
    const volunteerBucket = getVolunteerHoursBucket_(volunteerHours);

    // Engagement level
    const engagementLevel = getEngagementLevel_(volunteerHours, lastContact, isSteward, stats.total);

    exportData.push([
      memberHash,
      row[cols.JOB_TITLE - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.WORK_LOCATION - 1] || '',
      isSteward === 'Yes' ? 'Yes' : 'No',
      isSteward === 'Member Leader' ? 'Yes' : 'No',
      stats.open > 0 ? 'Yes' : 'No',
      stats.total,
      stats.won,
      stats.lost,
      daysSinceContactBucket,
      contactFreq,
      volunteerBucket,
      engagementLevel,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes anonymized Satisfaction sheet - surveys are already mostly anonymous.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonSatisfaction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.SATISFACTION || '📊 Member Satisfaction');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.SATISFACTION);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const timestamp = row[0];
    if (!timestamp) continue;

    // Generate anonymous response hash
    const responseHash = generateAnonHash_('SR' + i + (timestamp instanceof Date ? timestamp.getTime() : ''));

    // Date dimensions
    const respMonth = timestamp instanceof Date ? Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const respQuarter = timestamp instanceof Date ? getQuarter_(timestamp) : '';
    const respYear = timestamp instanceof Date ? timestamp.getFullYear() : '';

    // Calculate section averages
    const overallSatAvg = calculateSectionAvg_(row, [6, 7, 8, 9]);
    const stewardRatingAvg = calculateSectionAvg_(row, [10, 11, 12, 13, 14, 15, 16]);
    const stewardAccessAvg = calculateSectionAvg_(row, [18, 19, 20]);
    const chapterAvg = calculateSectionAvg_(row, [21, 22, 23, 24, 25]);
    const leadershipAvg = calculateSectionAvg_(row, [26, 27, 28, 29, 30, 31]);
    const contractAvg = calculateSectionAvg_(row, [32, 33, 34, 35]);
    const commAvg = calculateSectionAvg_(row, [41, 42, 43, 44, 45]);
    const voiceAvg = calculateSectionAvg_(row, [46, 47, 48, 49, 50]);
    const valueAvg = calculateSectionAvg_(row, [51, 52, 53, 54, 55]);
    const schedAvg = calculateSectionAvg_(row, [56, 57, 58, 59, 60, 61, 62]);
    const repAvg = calculateSectionAvg_(row, [37, 38, 39, 40]);

    // Get individual questions - BUCKETED for aggregation
    const worksite = row[1] || '';
    const role = categorizeRole_(row[2] || '');
    const shift = row[3] || '';
    const timeInRole = categorizeTenure_(row[4] || '');
    const hasStewardContact = row[5] || '';
    const filedGrievance = row[36] || '';

    // Bucket satisfaction scores (1-10 → Low/Medium/High)
    const satisfiedBucket = getScoreBucket_(row[6]);
    const trustBucket = getScoreBucket_(row[7]);
    const protectedBucket = getScoreBucket_(row[8]);
    const recommendBucket = getScoreBucket_(row[9]);

    exportData.push([
      responseHash,
      respMonth,
      respQuarter,
      respYear,
      worksite,
      role,
      shift,
      timeInRole,
      hasStewardContact,
      overallSatAvg,
      stewardRatingAvg,
      stewardAccessAvg,
      chapterAvg,
      leadershipAvg,
      contractAvg,
      commAvg,
      voiceAvg,
      valueAvg,
      schedAvg,
      repAvg,
      satisfiedBucket,
      trustBucket,
      protectedBucket,
      recommendBucket,
      filedGrievance,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Generates an anonymous, non-reversible hash from an identifier.
 * @private
 */
function generateAnonHash_(id) {
  // Use a simple hash that can't be reversed to original ID
  const salt = 'anon509data';
  const combined = salt + String(id);
  let hash = 0;
  for (let i = 0; i < combined.length; i++) {
    const char = combined.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return 'A' + Math.abs(hash).toString(36).toUpperCase().substring(0, 8);
}

/**
 * Categorizes days into buckets for privacy.
 * @private
 */
function getDaysBucket_(days) {
  if (days <= 7) return 'Within Week';
  if (days <= 30) return 'Within Month';
  if (days <= 90) return '1-3 Months';
  if (days <= 180) return '3-6 Months';
  if (days <= 365) return '6-12 Months';
  return 'Over 1 Year';
}

/**
 * Categorizes contact frequency.
 * @private
 */
function getContactFrequencyCategory_(lastContact, now) {
  if (!(lastContact instanceof Date)) return 'No Contact';
  const days = Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24));
  if (days <= 30) return 'Active';
  if (days <= 90) return 'Regular';
  if (days <= 180) return 'Occasional';
  return 'Inactive';
}

/**
 * Categorizes volunteer hours into buckets.
 * @private
 */
function getVolunteerHoursBucket_(hours) {
  if (hours === 0) return 'None';
  if (hours <= 5) return '1-5 Hours';
  if (hours <= 20) return '6-20 Hours';
  if (hours <= 50) return '21-50 Hours';
  return '50+ Hours';
}

/**
 * Calculates engagement level from multiple factors.
 * @private
 */
function getEngagementLevel_(volunteerHours, lastContact, isSteward, grievanceCount) {
  let score = 0;

  // Volunteer hours contribution
  if (volunteerHours > 50) score += 3;
  else if (volunteerHours > 20) score += 2;
  else if (volunteerHours > 5) score += 1;

  // Recent contact contribution
  if (lastContact instanceof Date) {
    const days = Math.ceil((new Date() - lastContact) / (1000 * 60 * 60 * 24));
    if (days <= 30) score += 2;
    else if (days <= 90) score += 1;
  }

  // Leadership role
  if (isSteward === 'Yes' || isSteward === 'Member Leader') score += 2;

  // Grievance involvement
  if (grievanceCount > 0) score += 1;

  // Map to engagement level
  if (score >= 6) return 'Highly Engaged';
  if (score >= 4) return 'Engaged';
  if (score >= 2) return 'Somewhat Engaged';
  if (score >= 1) return 'Low Engagement';
  return 'Not Engaged';
}

/**
 * Categorizes role into broader categories.
 * @private
 */
function categorizeRole_(role) {
  const roleLower = String(role).toLowerCase();
  if (roleLower.includes('steward') || roleLower.includes('leader')) return 'Leadership';
  if (roleLower.includes('nurse') || roleLower.includes('rn') || roleLower.includes('lpn')) return 'Nursing';
  if (roleLower.includes('tech') || roleLower.includes('aide')) return 'Technical/Support';
  if (roleLower.includes('admin') || roleLower.includes('clerk')) return 'Administrative';
  return 'Other';
}

/**
 * Categorizes tenure into buckets.
 * @private
 */
function categorizeTenure_(tenure) {
  const tenureLower = String(tenure).toLowerCase();
  if (tenureLower.includes('less than') || tenureLower.includes('< 1')) return 'New (< 1 year)';
  if (tenureLower.includes('1-3') || tenureLower.includes('1 to 3')) return '1-3 Years';
  if (tenureLower.includes('3-5') || tenureLower.includes('3 to 5')) return '3-5 Years';
  if (tenureLower.includes('5-10') || tenureLower.includes('5 to 10')) return '5-10 Years';
  return '10+ Years';
}

/**
 * Converts numeric score (1-10) to bucket.
 * @private
 */
function getScoreBucket_(score) {
  const num = parseFloat(score);
  if (isNaN(num)) return 'No Response';
  if (num >= 8) return 'High (8-10)';
  if (num >= 5) return 'Medium (5-7)';
  return 'Low (1-4)';
}

/**
 * Shows help dialog for PII-free Looker connection.
 */
function showLookerAnonConnectionHelp() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const html = `<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Roboto, sans-serif; padding: 20px; color: #1F2937; }
    h2 { color: #059669; margin-bottom: 16px; }
    .pii-badge { display: inline-block; background: #059669; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; margin-bottom: 16px; }
    .step { background: #F3F4F6; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; }
    .step-num { display: inline-block; width: 24px; height: 24px; background: #059669; color: white; border-radius: 50%; text-align: center; line-height: 24px; margin-right: 8px; font-size: 12px; }
    code { background: #E5E7EB; padding: 2px 6px; border-radius: 4px; font-size: 12px; }
    .sheets { margin: 16px 0; }
    .sheet-name { display: inline-block; background: #334155; color: #F8FAFC; padding: 4px 12px; border-radius: 4px; margin: 4px; font-size: 13px; }
    .excluded { margin: 16px 0; padding: 12px; background: #ECFDF5; border: 1px solid #059669; border-radius: 8px; }
    .excluded h4 { color: #059669; margin-bottom: 8px; font-size: 13px; }
    .excluded ul { margin: 0; padding-left: 20px; font-size: 12px; color: #065F46; }
    .btn { display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #059669, #047857); color: white; text-decoration: none; border-radius: 6px; margin-top: 12px; }
  </style>
</head>
<body>
  <h2>PII-Free Looker Connection</h2>
  <span class="pii-badge">✓ NO PERSONAL DATA</span>

  <div class="excluded">
    <h4>Excluded from these sheets:</h4>
    <ul>
      <li>Member names (first, last, full)</li>
      <li>Email addresses</li>
      <li>Phone numbers</li>
      <li>Member IDs (uses anonymous hashes)</li>
      <li>Steward names</li>
    </ul>
  </div>

  <div class="step">
    <span class="step-num">1</span>
    Open <a href="https://lookerstudio.google.com" target="_blank">Looker Studio</a> and create a new report
  </div>

  <div class="step">
    <span class="step-num">2</span>
    Add a data source → Select <strong>Google Sheets</strong>
  </div>

  <div class="step">
    <span class="step-num">3</span>
    Select these ANONYMIZED sheets:
    <div class="sheets">
      <span class="sheet-name">_Looker_Anon_Members</span>
      <span class="sheet-name">_Looker_Anon_Grievances</span>
      <span class="sheet-name">_Looker_Anon_Satisfaction</span>
    </div>
  </div>

  <a class="btn" href="https://lookerstudio.google.com/reporting/create" target="_blank">
    Open Looker Studio →
  </a>

  <p style="margin-top: 20px; font-size: 11px; color: #6B7280;">
    Safe for: External stakeholders, public dashboards, compliance reporting
  </p>
</body>
</html>`;

  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(480)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(dialog, 'PII-Free Looker Connection');
}

/**
 * Gets status of PII-free Looker integration.
 * @returns {Object} Status info
 */
function getLookerAnonStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = {
    isSetup: true,
    piiExcluded: true,
    sheets: {},
    recordCounts: {}
  };

  for (const key in LOOKER_CONFIG.SHEETS_ANON) {
    const sheetName = LOOKER_CONFIG.SHEETS_ANON[key];
    const sheet = ss.getSheetByName(sheetName);

    status.sheets[key] = {
      exists: !!sheet,
      name: sheetName,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
    };

    if (!sheet) status.isSetup = false;
    status.recordCounts[key] = status.sheets[key].rows;
  }

  return status;
}

/**
 * Installs trigger to refresh both standard and PII-free Looker data.
 */
function installLookerAllRefreshTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'refreshLookerData' || fn === 'refreshLookerAnonData' || fn === 'refreshAllLookerData_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install combined refresh trigger
  ScriptApp.newTrigger('refreshAllLookerData_')
    .timeBased()
    .everyDays(1)
    .atHour(LOOKER_CONFIG.AUTO_REFRESH_HOUR)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily refresh installed for both standard and PII-free data',
    'Looker Triggers Installed',
    5
  );

  return { success: true };
}

/**
 * Refreshes both standard and PII-free Looker data.
 * @private
 */
function refreshAllLookerData_() {
  refreshLookerData();
  refreshLookerAnonData();
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Calculates average of numeric values in specified columns.
 * @private
 */
function calculateSectionAvg_(row, colIndices) {
  let sum = 0;
  let count = 0;

  for (let i = 0; i < colIndices.length; i++) {
    const val = row[colIndices[i]];
    const num = parseFloat(val);
    if (!isNaN(num) && num >= 1 && num <= 10) {
      sum += num;
      count++;
    }
  }

  return count > 0 ? Math.round((sum / count) * 10) / 10 : '';
}

/**
 * Gets quarter string from date.
 * @private
 */
function getQuarter_(date) {
  const month = date.getMonth();
  const quarter = Math.floor(month / 3) + 1;
  return date.getFullYear() + '-Q' + quarter;
}

/**
 * Gets outcome category from status.
 * @private
 */
function getOutcomeCategory_(status) {
  if (status === 'Won') return 'Win';
  if (status === 'Denied') return 'Loss';
  if (status === 'Settled') return 'Settlement';
  if (status === 'Withdrawn') return 'Withdrawn';
  if (status === 'Closed') return 'Closed';
  return 'Active';
}

// ============================================================================
// TRIGGERS & AUTOMATION
// ============================================================================

/**
 * Installs daily auto-refresh trigger for Looker data.
 */
function installLookerRefreshTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshLookerData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install new trigger
  ScriptApp.newTrigger('refreshLookerData')
    .timeBased()
    .everyDays(1)
    .atHour(LOOKER_CONFIG.AUTO_REFRESH_HOUR)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily refresh installed (' + LOOKER_CONFIG.AUTO_REFRESH_HOUR + ':00 AM)',
    'Looker Trigger Installed',
    5
  );

  return { success: true };
}

/**
 * Gets the Looker Studio connection URL for this spreadsheet.
 * @returns {string} URL to create new Looker Studio report
 */
function getLookerConnectionUrl() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return 'https://lookerstudio.google.com/reporting/create?c.reportId=&r.reportName=509%20Dashboard%20Report&ds.connector=googleSheets&ds.spreadsheetId=' + ssId;
}

/**
 * Shows dialog with Looker Studio connection instructions.
 * Lists only the 3 allowed data sources.
 */
function showLookerConnectionHelp() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const html = `<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Roboto, sans-serif; padding: 20px; color: #1F2937; }
    h2 { color: #7C3AED; margin-bottom: 16px; }
    .step { background: #F3F4F6; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; }
    .step-num { display: inline-block; width: 24px; height: 24px; background: #7C3AED; color: white; border-radius: 50%; text-align: center; line-height: 24px; margin-right: 8px; font-size: 12px; }
    code { background: #E5E7EB; padding: 2px 6px; border-radius: 4px; font-size: 12px; }
    .sheets { margin: 16px 0; }
    .sheet-name { display: inline-block; background: #1E293B; color: #F8FAFC; padding: 4px 12px; border-radius: 4px; margin: 4px; font-size: 13px; }
    .source-info { font-size: 11px; color: #6B7280; display: block; margin-top: 2px; }
    .btn { display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; text-decoration: none; border-radius: 6px; margin-top: 12px; }
    .btn:hover { opacity: 0.9; }
    .restricted { background: #FEF3C7; border: 1px solid #F59E0B; padding: 10px; border-radius: 6px; margin: 16px 0; font-size: 13px; }
  </style>
</head>
<body>
  <h2>Connect to Looker Studio</h2>

  <div class="restricted">
    <strong>Restricted Data Access:</strong> Only Member Directory, Grievance Log, and Member Satisfaction data is available to Looker Studio.
  </div>

  <div class="step">
    <span class="step-num">1</span>
    Open <a href="https://lookerstudio.google.com" target="_blank">Looker Studio</a> and create a new report
  </div>

  <div class="step">
    <span class="step-num">2</span>
    Add a data source → Select <strong>Google Sheets</strong>
  </div>

  <div class="step">
    <span class="step-num">3</span>
    Find this spreadsheet or paste the URL:<br>
    <code style="word-break: break-all; display: block; margin-top: 8px;">${ssUrl}</code>
  </div>

  <div class="step">
    <span class="step-num">4</span>
    Select one of these sheets as your data source:
    <div class="sheets">
      <span class="sheet-name">_Looker_Members<span class="source-info">← Member Directory</span></span>
      <span class="sheet-name">_Looker_Grievances<span class="source-info">← Grievance Log</span></span>
      <span class="sheet-name">_Looker_Satisfaction<span class="source-info">← Survey Data</span></span>
    </div>
  </div>

  <div class="step">
    <span class="step-num">5</span>
    Add multiple data sources to create relationships between members and grievances
  </div>

  <a class="btn" href="https://lookerstudio.google.com/reporting/create" target="_blank">
    Open Looker Studio →
  </a>

  <p style="margin-top: 20px; font-size: 12px; color: #6B7280;">
    Spreadsheet ID: <code>${ssId}</code>
  </p>
</body>
</html>`;

  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(520)
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(dialog, 'Looker Studio Connection');
}

/**
 * Gets Looker integration status.
 * @returns {Object} Status of Looker sheets and data
 */
function getLookerStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = {
    isSetup: true,
    allowedSources: LOOKER_CONFIG.ALLOWED_SOURCES,
    sheets: {},
    lastRefresh: null,
    recordCounts: {}
  };

  for (const key in LOOKER_CONFIG.SHEETS) {
    const sheetName = LOOKER_CONFIG.SHEETS[key];
    const sheet = ss.getSheetByName(sheetName);

    status.sheets[key] = {
      exists: !!sheet,
      name: sheetName,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
    };

    if (!sheet) status.isSetup = false;
    status.recordCounts[key] = status.sheets[key].rows;
  }

  // Get last refresh time from first data sheet
  const memberSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    const lastCol = memberSheet.getLastColumn();
    const lastUpdate = memberSheet.getRange(2, lastCol).getValue();
    if (lastUpdate instanceof Date) {
      status.lastRefresh = lastUpdate;
    }
  }

  return status;
}
