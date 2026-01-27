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
