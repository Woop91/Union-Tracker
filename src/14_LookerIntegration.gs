/**
 * ============================================================================
 * LOOKER STUDIO INTEGRATION
 * ============================================================================
 * Provides a read-only data layer for Google Looker Studio (Data Studio).
 * Creates optimized hidden sheets that Looker can connect to as data sources.
 *
 * ARCHITECTURE:
 * - Does NOT modify existing sheets or modal dashboards
 * - Creates separate hidden "_Looker_*" sheets for data export
 * - Looker connects to these sheets as native Google Sheets data sources
 * - Data is refreshed on-demand or via scheduled trigger
 *
 * SHEETS CREATED:
 * - _Looker_Grievances: Flattened grievance data with all fields
 * - _Looker_Members: Member directory data with computed fields
 * - _Looker_Metrics: Time-series metrics for trend analysis
 * - _Looker_Summary: Snapshot metrics for KPI cards
 *
 * USAGE:
 * 1. Run setupLookerIntegration() once to create sheets
 * 2. Run refreshLookerData() to update data (or install trigger)
 * 3. Connect Looker Studio to the spreadsheet, select _Looker_* sheets
 */

const LOOKER_CONFIG = {
  // Sheet names for Looker data sources
  SHEETS: {
    GRIEVANCES: '_Looker_Grievances',
    MEMBERS: '_Looker_Members',
    METRICS: '_Looker_Metrics',
    SUMMARY: '_Looker_Summary'
  },
  // Refresh settings
  AUTO_REFRESH_HOUR: 6, // 6 AM daily refresh
  // Data retention for metrics history
  METRICS_HISTORY_DAYS: 365
};

// ============================================================================
// SETUP & MANAGEMENT
// ============================================================================

/**
 * Sets up Looker Studio integration by creating hidden data sheets.
 * Run this once to initialize the integration.
 * @returns {Object} Result with created sheet names
 */
function setupLookerIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Looker Studio Integration',
    'This will create hidden data sheets for Looker Studio:\n\n' +
    '• _Looker_Grievances - Case data\n' +
    '• _Looker_Members - Member data\n' +
    '• _Looker_Metrics - Historical metrics\n' +
    '• _Looker_Summary - KPI snapshots\n\n' +
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

  // Initialize headers
  initializeLookerGrievancesSheet_();
  initializeLookerMembersSheet_();
  initializeLookerMetricsSheet_();
  initializeLookerSummarySheet_();

  // Do initial data population
  refreshLookerData();

  const summary = created.length > 0
    ? 'Created sheets: ' + created.join(', ')
    : 'All Looker sheets already exist';

  ui.alert('Setup Complete', summary + '\n\nData has been populated. Connect Looker Studio to this spreadsheet and select the _Looker_* sheets as data sources.', ui.ButtonSet.OK);

  // Log setup
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_INITIALIZED, {
      action: 'LOOKER_SETUP',
      sheetsCreated: created,
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
    'Member ID', 'Full Name', 'Job Title', 'Unit', 'Location',
    'Is Steward', 'Is Member Leader', 'Assigned Steward',
    'Has Open Grievance', 'Total Grievances', 'Grievances Won', 'Grievances Lost',
    'Last Contact Date', 'Days Since Contact',
    'Engagement Score', 'Volunteer Hours',
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
 * Initializes the Looker Metrics sheet with headers.
 * @private
 */
function initializeLookerMetricsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.METRICS);
  if (!sheet) return;

  const headers = [
    'Snapshot Date', 'Snapshot Month', 'Snapshot Quarter', 'Snapshot Year',
    'Total Members', 'Total Stewards', 'Total Member Leaders',
    'Open Grievances', 'Pending Grievances', 'Closed This Month',
    'Won This Month', 'Lost This Month', 'Settled This Month',
    'Avg Days to Resolution', 'Overdue Count',
    'Win Rate Percent', 'Active Case Load Per Steward',
    'Members With Open Cases', 'Reminders Due This Week'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the Looker Summary sheet with headers.
 * @private
 */
function initializeLookerSummarySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SUMMARY);
  if (!sheet) return;

  const headers = ['Metric Name', 'Current Value', 'Previous Value', 'Change', 'Change Percent', 'Last Updated'];

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
 * Call manually or via scheduled trigger.
 * @returns {Object} Result with refresh counts
 */
function refreshLookerData() {
  const startTime = new Date();

  const grievanceCount = refreshLookerGrievances_();
  const memberCount = refreshLookerMembers_();
  const metricsAdded = refreshLookerMetrics_();
  const summaryCount = refreshLookerSummary_();

  const duration = (new Date() - startTime) / 1000;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Grievances: ${grievanceCount}, Members: ${memberCount}, Metrics: ${metricsAdded ? 'Added' : 'Updated'}`,
    'Looker Data Refreshed',
    5
  );

  return {
    success: true,
    grievances: grievanceCount,
    members: memberCount,
    metricsAdded: metricsAdded,
    summaryMetrics: summaryCount,
    durationSeconds: duration
  };
}

/**
 * Refreshes the Looker Grievances sheet.
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

  // Pre-compute column indices
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

    // Date dimensions for filtering
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

  // Clear old data (keep header) and write new data
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes the Looker Members sheet.
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

    // Simple engagement score (customize as needed)
    const volunteerHours = parseFloat(row[cols.VOLUNTEER_HOURS - 1]) || 0;
    const engagementScore = calculateEngagementScore_(row, stats);

    exportData.push([
      memberId,
      (firstName + ' ' + lastName).trim(),
      row[cols.JOB_TITLE - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.WORK_LOCATION - 1] || '',
      isSteward === 'Yes' ? 'Yes' : 'No',
      isSteward === 'Member Leader' ? 'Yes' : 'No',
      row[cols.ASSIGNED_STEWARD - 1] || '',
      stats.open > 0 ? 'Yes' : 'No',
      stats.total,
      stats.won,
      stats.lost,
      lastContact || '',
      daysSinceContact,
      engagementScore,
      volunteerHours,
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
 * Refreshes the Looker Metrics sheet with today's snapshot.
 * Appends a new row if today's date doesn't exist; updates if it does.
 * @private
 * @returns {boolean} True if new row added, false if updated
 */
function refreshLookerMetrics_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.METRICS);
  if (!targetSheet) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Calculate current metrics
  const metrics = calculateCurrentMetrics_();

  const newRow = [
    today,
    Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM'),
    getQuarter_(today),
    today.getFullYear(),
    metrics.totalMembers,
    metrics.totalStewards,
    metrics.totalMemberLeaders,
    metrics.openGrievances,
    metrics.pendingGrievances,
    metrics.closedThisMonth,
    metrics.wonThisMonth,
    metrics.lostThisMonth,
    metrics.settledThisMonth,
    metrics.avgDaysToResolution,
    metrics.overdueCount,
    metrics.winRatePercent,
    metrics.activeCaseLoadPerSteward,
    metrics.membersWithOpenCases,
    metrics.remindersDueThisWeek
  ];

  // Check if today's row exists
  const existingData = targetSheet.getDataRange().getValues();
  let todayRowIndex = -1;

  for (let i = 1; i < existingData.length; i++) {
    const rowDate = existingData[i][0];
    if (rowDate instanceof Date) {
      const rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowDateStr === todayStr) {
        todayRowIndex = i + 1;
        break;
      }
    }
  }

  if (todayRowIndex > 0) {
    // Update existing row
    targetSheet.getRange(todayRowIndex, 1, 1, newRow.length).setValues([newRow]);
    return false;
  } else {
    // Append new row
    targetSheet.appendRow(newRow);
    return true;
  }
}

/**
 * Refreshes the Looker Summary sheet with KPI data.
 * @private
 * @returns {number} Number of metrics exported
 */
function refreshLookerSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SUMMARY);
  const metricsSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.METRICS);
  if (!targetSheet) return 0;

  const current = calculateCurrentMetrics_();
  const now = new Date();

  // Get previous month's metrics for comparison
  let previous = {};
  if (metricsSheet && metricsSheet.getLastRow() > 1) {
    const metricsData = metricsSheet.getDataRange().getValues();
    const lastMonth = new Date(now);
    lastMonth.setMonth(lastMonth.getMonth() - 1);
    const lastMonthStr = Utilities.formatDate(lastMonth, Session.getScriptTimeZone(), 'yyyy-MM');

    // Find last month's data
    for (let i = metricsData.length - 1; i >= 1; i--) {
      const rowDate = metricsData[i][0];
      if (rowDate instanceof Date) {
        const rowMonthStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM');
        if (rowMonthStr === lastMonthStr) {
          previous = {
            totalMembers: metricsData[i][4],
            openGrievances: metricsData[i][7],
            winRatePercent: metricsData[i][15]
          };
          break;
        }
      }
    }
  }

  // Build summary rows
  const summaryData = [
    buildSummaryRow_('Total Members', current.totalMembers, previous.totalMembers, now),
    buildSummaryRow_('Total Stewards', current.totalStewards, null, now),
    buildSummaryRow_('Member Leaders', current.totalMemberLeaders, null, now),
    buildSummaryRow_('Open Grievances', current.openGrievances, previous.openGrievances, now),
    buildSummaryRow_('Pending Info', current.pendingGrievances, null, now),
    buildSummaryRow_('Overdue Cases', current.overdueCount, null, now),
    buildSummaryRow_('Win Rate %', current.winRatePercent, previous.winRatePercent, now),
    buildSummaryRow_('Avg Days to Resolution', current.avgDaysToResolution, null, now),
    buildSummaryRow_('Cases Per Steward', current.activeCaseLoadPerSteward, null, now),
    buildSummaryRow_('Won This Month', current.wonThisMonth, null, now),
    buildSummaryRow_('Lost This Month', current.lostThisMonth, null, now),
    buildSummaryRow_('Settled This Month', current.settledThisMonth, null, now),
    buildSummaryRow_('Reminders This Week', current.remindersDueThisWeek, null, now)
  ];

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  targetSheet.getRange(2, 1, summaryData.length, summaryData[0].length).setValues(summaryData);

  return summaryData.length;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Calculates current metrics from source sheets.
 * @private
 * @returns {Object} Current metrics
 */
function calculateCurrentMetrics_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR || 'Member Directory');
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  const metrics = {
    totalMembers: 0,
    totalStewards: 0,
    totalMemberLeaders: 0,
    openGrievances: 0,
    pendingGrievances: 0,
    closedThisMonth: 0,
    wonThisMonth: 0,
    lostThisMonth: 0,
    settledThisMonth: 0,
    avgDaysToResolution: 0,
    overdueCount: 0,
    winRatePercent: 0,
    activeCaseLoadPerSteward: 0,
    membersWithOpenCases: 0,
    remindersDueThisWeek: 0
  };

  const now = new Date();
  const thisMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  // Member metrics
  if (memberSheet && memberSheet.getLastRow() > 1) {
    const memberData = memberSheet.getDataRange().getValues();
    for (let i = 1; i < memberData.length; i++) {
      const memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
      if (!memberId) continue;

      metrics.totalMembers++;
      const isSteward = memberData[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isSteward === 'Yes') metrics.totalStewards++;
      if (isSteward === 'Member Leader') metrics.totalMemberLeaders++;
    }
  }

  // Grievance metrics
  const membersWithOpen = new Set();
  let totalResolutionDays = 0;
  let resolvedCount = 0;
  let totalClosed = 0;

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    const gData = grievanceSheet.getDataRange().getValues();
    const weekAhead = new Date(now);
    weekAhead.setDate(weekAhead.getDate() + 7);

    for (let i = 1; i < gData.length; i++) {
      const gId = gData[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      if (!gId) continue;

      const status = gData[i][GRIEVANCE_COLS.STATUS - 1];
      const memberId = gData[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      const dateFiled = gData[i][GRIEVANCE_COLS.DATE_FILED - 1];
      const dateClosed = gData[i][GRIEVANCE_COLS.DATE_CLOSED - 1];
      const daysToDeadline = gData[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
      const r1Date = gData[i][GRIEVANCE_COLS.REMINDER_1_DATE - 1];
      const r2Date = gData[i][GRIEVANCE_COLS.REMINDER_2_DATE - 1];

      // Count open/pending
      if (status === 'Open') {
        metrics.openGrievances++;
        if (memberId) membersWithOpen.add(memberId);
      } else if (status === 'Pending Info') {
        metrics.pendingGrievances++;
        if (memberId) membersWithOpen.add(memberId);
      } else if (status === 'Appealed' || status === 'In Arbitration') {
        metrics.openGrievances++;
        if (memberId) membersWithOpen.add(memberId);
      }

      // Overdue
      if (typeof daysToDeadline === 'number' && daysToDeadline < 0 &&
          status !== 'Closed' && status !== 'Won' && status !== 'Denied' &&
          status !== 'Settled' && status !== 'Withdrawn') {
        metrics.overdueCount++;
      }

      // This month's closures
      if (dateClosed instanceof Date) {
        const closedMonth = Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM');
        if (closedMonth === thisMonth) {
          metrics.closedThisMonth++;
          if (status === 'Won') metrics.wonThisMonth++;
          if (status === 'Denied') metrics.lostThisMonth++;
          if (status === 'Settled') metrics.settledThisMonth++;
        }

        // Resolution time
        if (dateFiled instanceof Date) {
          const days = Math.ceil((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
          if (days > 0) {
            totalResolutionDays += days;
            resolvedCount++;
          }
        }

        totalClosed++;
      }

      // Reminders due this week
      if (r1Date instanceof Date && r1Date >= now && r1Date <= weekAhead) {
        metrics.remindersDueThisWeek++;
      }
      if (r2Date instanceof Date && r2Date >= now && r2Date <= weekAhead) {
        metrics.remindersDueThisWeek++;
      }
    }
  }

  metrics.membersWithOpenCases = membersWithOpen.size;
  metrics.avgDaysToResolution = resolvedCount > 0 ? Math.round(totalResolutionDays / resolvedCount) : 0;

  // Win rate
  const wonLost = metrics.wonThisMonth + metrics.lostThisMonth;
  if (wonLost > 0) {
    metrics.winRatePercent = Math.round((metrics.wonThisMonth / wonLost) * 100);
  } else if (totalClosed > 0) {
    // Use all-time if no this-month data
    const allWon = countByStatus_('Won');
    const allLost = countByStatus_('Denied');
    if (allWon + allLost > 0) {
      metrics.winRatePercent = Math.round((allWon / (allWon + allLost)) * 100);
    }
  }

  // Case load per steward
  const activeCount = metrics.openGrievances + metrics.pendingGrievances;
  if (metrics.totalStewards > 0) {
    metrics.activeCaseLoadPerSteward = Math.round((activeCount / metrics.totalStewards) * 10) / 10;
  }

  return metrics;
}

/**
 * Counts grievances by status.
 * @private
 */
function countByStatus_(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.STATUS - 1] === status) count++;
  }
  return count;
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

/**
 * Calculates member engagement score.
 * @private
 */
function calculateEngagementScore_(memberRow, grievanceStats) {
  let score = 0;
  const cols = MEMBER_COLS;

  // Volunteer hours (0-30 points)
  const volunteerHours = parseFloat(memberRow[cols.VOLUNTEER_HOURS - 1]) || 0;
  score += Math.min(30, volunteerHours);

  // Recent contact (0-25 points)
  const lastContact = memberRow[cols.RECENT_CONTACT_DATE - 1];
  if (lastContact instanceof Date) {
    const daysSince = Math.ceil((new Date() - lastContact) / (1000 * 60 * 60 * 24));
    if (daysSince < 30) score += 25;
    else if (daysSince < 60) score += 15;
    else if (daysSince < 90) score += 10;
    else if (daysSince < 180) score += 5;
  }

  // Meeting attendance (0-20 points)
  const lastVirtual = memberRow[cols.LAST_VIRTUAL_MTG - 1];
  const lastInPerson = memberRow[cols.LAST_INPERSON_MTG - 1];
  if (lastVirtual instanceof Date || lastInPerson instanceof Date) {
    score += 20;
  }

  // Grievance involvement (0-25 points)
  if (grievanceStats.total > 0) score += 15;
  if (grievanceStats.won > 0) score += 10;

  return Math.min(100, score);
}

/**
 * Builds a summary row with change calculation.
 * @private
 */
function buildSummaryRow_(name, current, previous, timestamp) {
  const change = (previous !== null && previous !== undefined)
    ? current - previous
    : '';
  const changePercent = (previous !== null && previous !== undefined && previous !== 0)
    ? Math.round(((current - previous) / previous) * 100)
    : '';

  return [name, current, previous || '', change, changePercent, timestamp];
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
    .btn { display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; text-decoration: none; border-radius: 6px; margin-top: 12px; }
    .btn:hover { opacity: 0.9; }
  </style>
</head>
<body>
  <h2>Connect to Looker Studio</h2>

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
      <span class="sheet-name">_Looker_Grievances</span>
      <span class="sheet-name">_Looker_Members</span>
      <span class="sheet-name">_Looker_Metrics</span>
      <span class="sheet-name">_Looker_Summary</span>
    </div>
  </div>

  <div class="step">
    <span class="step-num">5</span>
    Add multiple data sources to combine member and grievance data
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
    .setWidth(500)
    .setHeight(480);

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

  // Get last refresh time from Summary sheet
  const summarySheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SUMMARY);
  if (summarySheet && summarySheet.getLastRow() > 1) {
    const lastUpdate = summarySheet.getRange(2, 6).getValue();
    if (lastUpdate instanceof Date) {
      status.lastRefresh = lastUpdate;
    }
  }

  return status;
}