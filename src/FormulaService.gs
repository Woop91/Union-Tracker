/**
 * ============================================================================
 * FormulaService.gs - Hidden Sheet & Formula Logic
 * ============================================================================
 *
 * This module manages the six hidden calculation sheets that power the
 * dashboard's "self-healing" formula system. These sheets contain complex
 * formulas that aggregate, calculate, and cross-reference data.
 *
 * SEPARATION OF CONCERNS: This logic is highly specialized and most users
 * will never need to touch this file. Isolating it reduces the risk of
 * breaking cross-sheet data syncs.
 *
 * Hidden Sheets:
 * - _CalcMembers: Member statistics and lookups
 * - _CalcGrievances: Grievance aggregations
 * - _CalcDeadlines: Deadline calculations and alerts
 * - _CalcStats: Dashboard statistics
 * - _CalcSync: Cross-sheet synchronization
 * - _CalcFormulas: Named formula references
 *
 * @fileoverview Hidden sheet and formula management
 * @version 2.0.0
 * @requires Constants.gs
 */

// ============================================================================
// HIDDEN SHEET SETUP
// ============================================================================

/**
 * Sets up all hidden calculation sheets
 * Creates missing sheets and initializes their formulas
 * @return {Object} Result with counts of created/repaired sheets
 */
function setupAllHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = {
    created: 0,
    repaired: 0,
    errors: []
  };

  // Setup each hidden sheet
  const setupFunctions = {
    [HIDDEN_SHEETS.CALC_MEMBERS]: setupCalcMembersSheet,
    [HIDDEN_SHEETS.CALC_GRIEVANCES]: setupCalcGrievancesSheet,
    [HIDDEN_SHEETS.CALC_DEADLINES]: setupCalcDeadlinesSheet,
    [HIDDEN_SHEETS.CALC_STATS]: setupCalcStatsSheet,
    [HIDDEN_SHEETS.CALC_SYNC]: setupCalcSyncSheet,
    [HIDDEN_SHEETS.CALC_FORMULAS]: setupCalcFormulasSheet
  };

  for (const [sheetName, setupFn] of Object.entries(setupFunctions)) {
    try {
      let sheet = ss.getSheetByName(sheetName);
      const wasCreated = !sheet;

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        results.created++;
      }

      // Run setup function
      setupFn(sheet);

      // Hide the sheet
      sheet.hideSheet();

      if (!wasCreated) {
        results.repaired++;
      }

    } catch (error) {
      console.error(`Error setting up ${sheetName}:`, error);
      results.errors.push(`${sheetName}: ${error.message}`);
    }
  }

  return results;
}

/**
 * Repairs all hidden sheets by regenerating their formulas
 * @return {Object} Result with repair count
 */
function repairAllHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let repaired = 0;

  for (const sheetName of Object.values(HIDDEN_SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Clear and regenerate
      sheet.clear();

      // Re-run setup based on sheet name
      switch (sheetName) {
        case HIDDEN_SHEETS.CALC_MEMBERS:
          setupCalcMembersSheet(sheet);
          break;
        case HIDDEN_SHEETS.CALC_GRIEVANCES:
          setupCalcGrievancesSheet(sheet);
          break;
        case HIDDEN_SHEETS.CALC_DEADLINES:
          setupCalcDeadlinesSheet(sheet);
          break;
        case HIDDEN_SHEETS.CALC_STATS:
          setupCalcStatsSheet(sheet);
          break;
        case HIDDEN_SHEETS.CALC_SYNC:
          setupCalcSyncSheet(sheet);
          break;
        case HIDDEN_SHEETS.CALC_FORMULAS:
          setupCalcFormulasSheet(sheet);
          break;
      }

      repaired++;
    }
  }

  return { repaired: repaired };
}

// ============================================================================
// INDIVIDUAL SHEET SETUP FUNCTIONS
// ============================================================================

/**
 * Sets up the _CalcMembers hidden sheet
 * Contains member statistics and lookup tables
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcMembersSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;

  // Header row
  sheet.getRange('A1').setValue('Member Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Total Members
  sheet.getRange('A2').setValue('Total Members');
  sheet.getRange('B2').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  // Active Members
  sheet.getRange('A3').setValue('Active Members');
  sheet.getRange('B3').setFormula(
    `=COUNTIF('${memberSheetName}'!K:K,"Active")`
  );

  // Members by Department (dynamic list)
  sheet.getRange('A5').setValue('Department');
  sheet.getRange('B5').setValue('Count');
  sheet.getRange('A5:B5').setFontWeight('bold');

  sheet.getRange('A6').setFormula(
    `=UNIQUE(FILTER('${memberSheetName}'!E:E,'${memberSheetName}'!E:E<>"Department",'${memberSheetName}'!E:E<>""))`
  );

  sheet.getRange('B6').setFormula(
    `=ARRAYFORMULA(IF(A6:A<>"",COUNTIF('${memberSheetName}'!E:E,A6:A),""))`
  );

  // Union Status breakdown
  sheet.getRange('D2').setValue('Union Status');
  sheet.getRange('E2').setValue('Count');
  sheet.getRange('D2:E2').setFontWeight('bold');

  sheet.getRange('D3').setValue('Full Member');
  sheet.getRange('E3').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Full Member")`
  );

  sheet.getRange('D4').setValue('Agency Fee');
  sheet.getRange('E4').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Agency Fee")`
  );

  sheet.getRange('D5').setValue('Non-Member');
  sheet.getRange('E5').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Non-Member")`
  );

  // Lookup helper for member names
  sheet.getRange('G1').setValue('ID->Name Lookup');
  sheet.getRange('G1').setFontWeight('bold');
  sheet.getRange('G2').setFormula(
    `=ARRAYFORMULA(IF('${memberSheetName}'!A2:A<>"",` +
    `'${memberSheetName}'!A2:A&"|"&'${memberSheetName}'!B2:B&" "&'${memberSheetName}'!C2:C,""))`
  );
}

/**
 * Sets up the _CalcGrievances hidden sheet
 * Contains grievance aggregations and summaries
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcGrievancesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Grievance Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Status counts
  sheet.getRange('A3').setValue('Status');
  sheet.getRange('B3').setValue('Count');
  sheet.getRange('A3:B3').setFontWeight('bold');

  const statuses = Object.values(GRIEVANCE_STATUS);
  statuses.forEach((status, index) => {
    sheet.getRange(4 + index, 1).setValue(status);
    sheet.getRange(4 + index, 2).setFormula(
      `=COUNTIF('${grievanceSheetName}'!W:W,"${status}")`
    );
  });

  // Grievances by Type
  sheet.getRange('D3').setValue('Type');
  sheet.getRange('E3').setValue('Count');
  sheet.getRange('D3:E3').setFontWeight('bold');

  sheet.getRange('D4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!E:E,` +
    `'${grievanceSheetName}'!E:E<>"Grievance Type",'${grievanceSheetName}'!E:E<>""))`
  );

  sheet.getRange('E4').setFormula(
    `=ARRAYFORMULA(IF(D4:D<>"",COUNTIF('${grievanceSheetName}'!E:E,D4:D),""))`
  );

  // Grievances by Current Step
  sheet.getRange('G3').setValue('Step');
  sheet.getRange('H3').setValue('Count');
  sheet.getRange('G3:H3').setFontWeight('bold');

  for (let step = 1; step <= 4; step++) {
    sheet.getRange(3 + step, 7).setValue(`Step ${step}`);
    sheet.getRange(3 + step, 8).setFormula(
      `=COUNTIF('${grievanceSheetName}'!H:H,${step})`
    );
  }

  // Monthly filing trend (last 12 months)
  sheet.getRange('A15').setValue('Monthly Filings');
  sheet.getRange('A15').setFontWeight('bold');

  sheet.getRange('A16').setValue('Month');
  sheet.getRange('B16').setValue('Filings');
  sheet.getRange('A16:B16').setFontWeight('bold');

  // Generate last 12 months
  for (let i = 0; i < 12; i++) {
    sheet.getRange(17 + i, 1).setFormula(
      `=EOMONTH(TODAY(),-${i})`
    );
    sheet.getRange(17 + i, 2).setFormula(
      `=SUMPRODUCT((MONTH('${grievanceSheetName}'!D:D)=MONTH(A${17 + i}))*` +
      `(YEAR('${grievanceSheetName}'!D:D)=YEAR(A${17 + i})))`
    );
  }
}

/**
 * Sets up the _CalcDeadlines hidden sheet
 * Contains deadline calculations and alert logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcDeadlinesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Deadline Calculations');
  sheet.getRange('A1').setFontWeight('bold');

  // Configuration reference
  sheet.getRange('A3').setValue('Deadline Rules (Days)');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setValue('Step 1 Response');
  sheet.getRange('B4').setValue(DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);

  sheet.getRange('A5').setValue('Step 2 Appeal');
  sheet.getRange('B5').setValue(DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL);

  sheet.getRange('A6').setValue('Step 2 Response');
  sheet.getRange('B6').setValue(DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE);

  sheet.getRange('A7').setValue('Step 3 Appeal');
  sheet.getRange('B7').setValue(DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL);

  sheet.getRange('A8').setValue('Step 3 Response');
  sheet.getRange('B8').setValue(DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE);

  sheet.getRange('A9').setValue('Arbitration Demand');
  sheet.getRange('B9').setValue(DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND);

  // Upcoming deadlines calculation
  sheet.getRange('D1').setValue('Upcoming Deadlines (Next 14 Days)');
  sheet.getRange('D1').setFontWeight('bold');

  sheet.getRange('D2').setValue('Grievance ID');
  sheet.getRange('E2').setValue('Step');
  sheet.getRange('F2').setValue('Due Date');
  sheet.getRange('G2').setValue('Days Left');
  sheet.getRange('D2:G2').setFontWeight('bold');

  // Complex formula to extract upcoming deadlines
  // This uses FILTER to get open grievances and calculate their current deadline
  sheet.getRange('D3').setFormula(
    `=IFERROR(FILTER('${grievanceSheetName}'!A:A,` +
    `('${grievanceSheetName}'!W:W="Open")+('${grievanceSheetName}'!W:W="Pending Response")+` +
    `('${grievanceSheetName}'!W:W="Appealed")),"")`
  );

  // Overdue grievances
  sheet.getRange('I1').setValue('Overdue Grievances');
  sheet.getRange('I1').setFontWeight('bold');
  sheet.getRange('I2').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"<>Resolved",` +
    `'${grievanceSheetName}'!W:W,"<>Closed",` +
    `'${grievanceSheetName}'!W:W,"<>Withdrawn",` +
    `'${grievanceSheetName}'!J:J,"<"&TODAY())`
  );

  // Alert thresholds
  sheet.getRange('I4').setValue('Alert Thresholds');
  sheet.getRange('I4').setFontWeight('bold');

  sheet.getRange('I5').setValue('Critical (<=3 days)');
  sheet.getRange('J5').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">="&TODAY(),` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+3)`
  );

  sheet.getRange('I6').setValue('Warning (4-7 days)');
  sheet.getRange('J6').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">"&TODAY()+3,` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+7)`
  );
}

/**
 * Sets up the _CalcStats hidden sheet
 * Contains dashboard-wide statistics
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcStatsSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Dashboard Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Quick stats for sidebar
  sheet.getRange('A3').setValue('Sidebar Stats');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setValue('open_grievances');
  sheet.getRange('B4').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Open")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Appealed")`
  );

  sheet.getRange('A5').setValue('pending_response');
  sheet.getRange('B5').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")`
  );

  sheet.getRange('A6').setValue('total_members');
  sheet.getRange('B6').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  sheet.getRange('A7').setValue('resolved_ytd');
  sheet.getRange('B7').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Resolved",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))+` +
    `COUNTIFS('${grievanceSheetName}'!W:W,"Closed",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))`
  );

  // Win rate calculation
  sheet.getRange('A9').setValue('Performance Metrics');
  sheet.getRange('A9').setFontWeight('bold');

  sheet.getRange('A10').setValue('total_resolved');
  sheet.getRange('B10').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Resolved")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Closed")`
  );

  sheet.getRange('A11').setValue('sustained_count');
  sheet.getRange('B11').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Sustained")`
  );

  sheet.getRange('A12').setValue('settled_count');
  sheet.getRange('B12').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Settled")`
  );

  sheet.getRange('A13').setValue('denied_count');
  sheet.getRange('B13').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Denied")`
  );

  sheet.getRange('A14').setValue('win_rate');
  sheet.getRange('B14').setFormula(
    `=IFERROR((B11+B12)/B10*100,0)`
  );

  // Average time to resolution
  sheet.getRange('A16').setValue('avg_days_to_resolve');
  sheet.getRange('B16').setFormula(
    `=IFERROR(AVERAGEIFS('${grievanceSheetName}'!X:X-'${grievanceSheetName}'!D:D,` +
    `'${grievanceSheetName}'!W:W,"Resolved"),0)`
  );
}

/**
 * Sets up the _CalcSync hidden sheet
 * Contains cross-sheet synchronization logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcSyncSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Cross-Sheet Sync');
  sheet.getRange('A1').setFontWeight('bold');

  // Member ID validation list
  sheet.getRange('A3').setValue('Valid Member IDs');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setFormula(
    `=FILTER('${memberSheetName}'!A:A,'${memberSheetName}'!A:A<>"ID",'${memberSheetName}'!A:A<>"")`
  );

  // Grievances per member count
  sheet.getRange('C3').setValue('Member ID');
  sheet.getRange('D3').setValue('Grievance Count');
  sheet.getRange('C3:D3').setFontWeight('bold');

  sheet.getRange('C4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!B:B,'${grievanceSheetName}'!B:B<>"Member ID",'${grievanceSheetName}'!B:B<>""))`
  );

  sheet.getRange('D4').setFormula(
    `=ARRAYFORMULA(IF(C4:C<>"",COUNTIF('${grievanceSheetName}'!B:B,C4:C),""))`
  );

  // Data consistency checks
  sheet.getRange('F3').setValue('Data Consistency');
  sheet.getRange('F3').setFontWeight('bold');

  sheet.getRange('F4').setValue('Orphaned Grievances');
  sheet.getRange('G4').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!B:B,"<>",'${grievanceSheetName}'!B:B,"<>Member ID")-` +
    `SUMPRODUCT(COUNTIF(A4:A,'${grievanceSheetName}'!B2:B))`
  );

  sheet.getRange('F5').setValue('Members with Grievances');
  sheet.getRange('G5').setFormula(
    `=COUNTA(C4:C)`
  );

  // Last sync timestamp
  sheet.getRange('F7').setValue('Last Formula Update');
  sheet.getRange('G7').setFormula('=NOW()');
}

/**
 * Sets up the _CalcFormulas hidden sheet
 * Contains named formula references for use in other sheets
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcFormulasSheet(sheet) {
  // Header
  sheet.getRange('A1').setValue('Named Formula References');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A2').setValue('Use these formulas via indirect references');

  // Department list formula
  sheet.getRange('A4').setValue('DEPARTMENT_LIST');
  sheet.getRange('B4').setFormula(
    `=SORT(UNIQUE(FILTER('${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E,` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"Department",` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"")))`
  );

  // Status list formula
  sheet.getRange('A6').setValue('STATUS_LIST');
  sheet.getRange('B6').setValue(Object.values(GRIEVANCE_STATUS).join(','));

  // Outcome list formula
  sheet.getRange('A8').setValue('OUTCOME_LIST');
  sheet.getRange('B8').setValue(Object.values(GRIEVANCE_OUTCOMES).join(','));

  // Grievance type list
  sheet.getRange('A10').setValue('GRIEVANCE_TYPES');
  sheet.getRange('B10').setValue('Contract Violation,Discipline,Discharge,Working Conditions,Safety,Other');

  // Date formatting formula
  sheet.getRange('A12').setValue('TODAY_FORMATTED');
  sheet.getRange('B12').setFormula('=TEXT(TODAY(),"MMMM D, YYYY")');

  // Year calculation
  sheet.getRange('A14').setValue('CURRENT_YEAR');
  sheet.getRange('B14').setFormula('=YEAR(TODAY())');

  // Next grievance ID prefix
  sheet.getRange('A16').setValue('GRIEVANCE_ID_PREFIX');
  sheet.getRange('B16').setFormula('="GRV-"&YEAR(TODAY())&"-"');
}

// ============================================================================
// FORMULA HELPERS
// ============================================================================

/**
 * Gets a value from a hidden calculation sheet
 * @param {string} sheetName - The hidden sheet name
 * @param {string} cellRef - The cell reference (e.g., 'B4')
 * @return {*} The cell value
 */
function getCalcValue(sheetName, cellRef) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.error(`Hidden sheet ${sheetName} not found`);
    return null;
  }

  return sheet.getRange(cellRef).getValue();
}

/**
 * Gets dashboard statistics from the calc sheet
 * Used by the sidebar
 * @return {Object} Statistics object
 */
function getDashboardStats() {
  const statsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HIDDEN_SHEETS.CALC_STATS);

  if (!statsSheet) {
    // Fallback to direct calculation
    return {
      open: getOpenGrievances().length,
      pending: 0,
      members: 0,
      resolved: 0
    };
  }

  // Read from pre-calculated values
  const data = statsSheet.getRange('A4:B7').getValues();
  const stats = {};

  data.forEach(row => {
    if (row[0] === 'open_grievances') stats.open = row[1];
    if (row[0] === 'pending_response') stats.pending = row[1];
    if (row[0] === 'total_members') stats.members = row[1];
    if (row[0] === 'resolved_ytd') stats.resolved = row[1];
  });

  return stats;
}

/**
 * Gets department list from calc sheet
 * @return {string[]} Array of department names
 */
function getDepartmentList() {
  const formulaSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HIDDEN_SHEETS.CALC_FORMULAS);

  if (!formulaSheet) {
    // Fallback to direct query
    const memberSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

    if (!memberSheet) return [];

    const data = memberSheet.getRange(2, MEMBER_COLUMNS.DEPARTMENT + 1,
      memberSheet.getLastRow() - 1, 1).getValues();

    const depts = new Set();
    data.forEach(row => {
      if (row[0]) depts.add(row[0]);
    });

    return Array.from(depts).sort();
  }

  // Read from pre-calculated list
  const deptData = formulaSheet.getRange('B4:B').getValues();
  return deptData.filter(row => row[0]).map(row => row[0]);
}

/**
 * Gets member list for dropdowns
 * @return {Array} Array of member objects with id, name, department
 */
function getMemberList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const members = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLUMNS.ID]) {
      members.push({
        id: data[i][MEMBER_COLUMNS.ID],
        name: `${data[i][MEMBER_COLUMNS.FIRST_NAME]} ${data[i][MEMBER_COLUMNS.LAST_NAME]}`,
        department: data[i][MEMBER_COLUMNS.DEPARTMENT]
      });
    }
  }

  return members;
}

/**
 * Gets member by ID
 * @param {string} memberId - The member ID
 * @return {Object|null} Member object or null
 */
function getMemberById(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLUMNS.ID] === memberId) {
      const member = {};
      headers.forEach((header, index) => {
        member[header] = data[i][index];
      });
      return member;
    }
  }

  return null;
}

// ============================================================================
// SEARCH FUNCTIONS (Used by UIService)
// ============================================================================

/**
 * Searches the dashboard for matching records
 * @param {string} query - Search query
 * @param {string} searchType - 'all', 'members', or 'grievances'
 * @param {Object} filters - Additional filters
 * @return {Array} Search results
 */
function searchDashboard(query, searchType, filters) {
  const results = [];
  const queryLower = query.toLowerCase();

  // Search members
  if (searchType === 'all' || searchType === 'members') {
    const members = getMemberList();
    members.forEach(m => {
      if (m.name.toLowerCase().includes(queryLower) ||
          m.id.toLowerCase().includes(queryLower) ||
          m.department.toLowerCase().includes(queryLower)) {
        results.push({
          id: m.id,
          type: 'member',
          title: m.name,
          subtitle: `${m.department} - ID: ${m.id}`
        });
      }
    });
  }

  // Search grievances
  if (searchType === 'all' || searchType === 'grievances') {
    const grievances = getOpenGrievances();
    grievances.forEach(g => {
      const grievanceId = g['Grievance ID'] || '';
      const memberName = g['Member Name'] || '';
      const description = g['Description'] || '';
      const status = g['Status'] || '';

      if (grievanceId.toLowerCase().includes(queryLower) ||
          memberName.toLowerCase().includes(queryLower) ||
          description.toLowerCase().includes(queryLower)) {

        // Apply status filter
        if (filters.status && status.toLowerCase() !== filters.status.toLowerCase()) {
          return;
        }

        results.push({
          id: grievanceId,
          type: 'grievance',
          title: `${grievanceId} - ${memberName}`,
          subtitle: `${status} - Step ${g['Current Step'] || 1}`
        });
      }
    });
  }

  return results.slice(0, 50); // Limit results
}

/**
 * Quick search for instant results
 * @param {string} query - Search query
 * @return {Array} Quick search results
 */
function quickSearchDashboard(query) {
  if (!query || query.length < 2) return [];

  const results = searchDashboard(query, 'all', {});
  return results.slice(0, 10);
}

/**
 * Advanced search with complex filters
 * @param {Object} filters - Search filters
 * @return {Array} Search results
 */
function advancedSearch(filters) {
  const results = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search members if included
  if (filters.includeMembers) {
    const memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
    if (memberSheet) {
      const data = memberSheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let matches = true;

        // Keyword filter
        if (filters.keywords) {
          const keywords = filters.keywords.toLowerCase();
          const rowText = row.join(' ').toLowerCase();
          if (!rowText.includes(keywords)) matches = false;
        }

        // Department filter
        if (filters.department && row[MEMBER_COLUMNS.DEPARTMENT] !== filters.department) {
          matches = false;
        }

        if (matches && row[MEMBER_COLUMNS.ID]) {
          results.push({
            id: row[MEMBER_COLUMNS.ID],
            type: 'Member',
            name: `${row[MEMBER_COLUMNS.FIRST_NAME]} ${row[MEMBER_COLUMNS.LAST_NAME]}`,
            details: row[MEMBER_COLUMNS.DEPARTMENT],
            status: row[MEMBER_COLUMNS.STATUS] || 'Active',
            date: row[MEMBER_COLUMNS.LAST_UPDATED] ?
                  Utilities.formatDate(new Date(row[MEMBER_COLUMNS.LAST_UPDATED]),
                    Session.getScriptTimeZone(), 'MM/dd/yyyy') : ''
          });
        }
      }
    }
  }

  // Search grievances if included
  if (filters.includeGrievances) {
    const grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    if (grievanceSheet) {
      const data = grievanceSheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let matches = true;

        // Keyword filter
        if (filters.keywords) {
          const keywords = filters.keywords.toLowerCase();
          const rowText = row.join(' ').toLowerCase();
          if (!rowText.includes(keywords)) matches = false;
        }

        // Status filter
        if (filters.statuses && filters.statuses.length > 0) {
          if (!filters.statuses.includes(row[GRIEVANCE_COLUMNS.STATUS])) {
            matches = false;
          }
        }

        // Date range filter
        if (filters.dateFrom) {
          const filingDate = new Date(row[GRIEVANCE_COLUMNS.FILING_DATE]);
          const fromDate = new Date(filters.dateFrom);
          if (filingDate < fromDate) matches = false;
        }

        if (filters.dateTo) {
          const filingDate = new Date(row[GRIEVANCE_COLUMNS.FILING_DATE]);
          const toDate = new Date(filters.dateTo);
          if (filingDate > toDate) matches = false;
        }

        if (matches && row[GRIEVANCE_COLUMNS.GRIEVANCE_ID]) {
          results.push({
            id: row[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
            type: 'Grievance',
            name: `${row[GRIEVANCE_COLUMNS.GRIEVANCE_ID]} - ${row[GRIEVANCE_COLUMNS.MEMBER_NAME]}`,
            details: row[GRIEVANCE_COLUMNS.GRIEVANCE_TYPE],
            status: row[GRIEVANCE_COLUMNS.STATUS],
            date: row[GRIEVANCE_COLUMNS.FILING_DATE] ?
                  Utilities.formatDate(new Date(row[GRIEVANCE_COLUMNS.FILING_DATE]),
                    Session.getScriptTimeZone(), 'MM/dd/yyyy') : ''
          });
        }
      }
    }
  }

  return results;
}
