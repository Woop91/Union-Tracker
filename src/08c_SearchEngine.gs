/**
 * ============================================================================
 * SearchEngine.gs - Search and Navigation Functions
 * ============================================================================
 *
 * This module handles all search-related functions including:
 * - Desktop search
 * - Quick search
 * - Advanced search with filters
 * - Navigation to search results
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Search and navigation functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// SEARCH FUNCTIONS
// ============================================================================

/**
 * Gets locations for desktop search filter dropdown
 * @returns {Array<string>} Array of unique locations sorted alphabetically
 */
function getDesktopSearchLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = [];

  var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (mSheet && mSheet.getLastRow() > 1) {
    var mData = mSheet.getRange(2, MEMBER_COLS.WORK_LOCATION, mSheet.getLastRow() - 1, 1).getValues();
    mData.forEach(function(row) {
      var loc = row[0];
      if (loc && locations.indexOf(loc) === -1) {
        locations.push(loc);
      }
    });
  }

  return locations.sort();
}

/**
 * Gets search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array<Object>} Array of search results
 */
function getDesktopSearchData(query, tab, filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = (query || '').toLowerCase();
  filters = filters || {};

  // Search Members
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var lastCol = Math.max(MEMBER_COLS.IS_STEWARD, MEMBER_COLS.WORK_LOCATION, MEMBER_COLS.JOB_TITLE, MEMBER_COLS.EMAIL);
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, lastCol).getValues();

      mData.forEach(function(row, index) {
        var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var firstName = row[MEMBER_COLS.FIRST_NAME - 1] || '';
        var lastName = row[MEMBER_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
        var jobTitle = row[MEMBER_COLS.JOB_TITLE - 1] || '';
        var location = row[MEMBER_COLS.WORK_LOCATION - 1] || '';
        var isSteward = row[MEMBER_COLS.IS_STEWARD - 1] || '';

        // Apply filters
        if (filters.location && location !== filters.location) return;
        if (filters.isSteward && isSteward !== filters.isSteward) return;

        // Search across fields
        var searchable = (memberId + ' ' + fullName + ' ' + email + ' ' + jobTitle + ' ' + location).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.location && !filters.isSteward) return;

        results.push({
          type: 'member',
          id: memberId,
          title: fullName.trim() || 'Unnamed Member',
          email: email,
          jobTitle: jobTitle,
          location: location,
          isSteward: isSteward,
          row: index + 2
        });
      });
    }
  }

  // Search Grievances
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var lastGCol = Math.max(GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.ISSUE_CATEGORY, GRIEVANCE_COLS.LOCATION, GRIEVANCE_COLS.STEWARD, GRIEVANCE_COLS.DATE_FILED);
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, lastGCol).getValues();

      gData.forEach(function(row, index) {
        var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
        var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
        var issueType = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '';
        var location = row[GRIEVANCE_COLS.LOCATION - 1] || '';
        var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
        var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1] || '';

        // Apply filters
        if (filters.status && status !== filters.status) return;
        if (filters.location && location !== filters.location) return;

        // Search across fields
        var searchable = (grievanceId + ' ' + fullName + ' ' + status + ' ' + issueType + ' ' + location + ' ' + steward).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.status && !filters.location) return;

        // Format date
        var filedDateStr = '';
        if (dateFiled) {
          try {
            filedDateStr = Utilities.formatDate(new Date(dateFiled), Session.getScriptTimeZone(), 'MM/dd/yyyy');
          } catch(e) {
            filedDateStr = dateFiled.toString();
          }
        }

        results.push({
          type: 'grievance',
          id: grievanceId,
          title: fullName.trim() || 'Unknown Member',
          status: status,
          issueType: issueType,
          location: location,
          steward: steward,
          filedDate: filedDateStr,
          row: index + 2
        });
      });
    }
  }

  return results.slice(0, 50);
}

/**
 * Navigates to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
 * @returns {void}
 */
function navigateToSearchResult(type, id, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = type === 'member' ? SHEETS.MEMBER_DIR : SHEETS.GRIEVANCE_LOG;
  var sheet = ss.getSheetByName(sheetName);

  if (sheet && row) {
    ss.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(row, 1));
    SpreadsheetApp.flush();
  }
}

/**
 * Navigates to active grievances view
 * @returns {void}
 */
function viewActiveGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * Searches the dashboard for matching records
 * @param {string} query - Search query
 * @param {string} searchType - 'all', 'members', or 'grievances'
 * @param {Object} filters - Additional filters
 * @returns {Array<Object>} Search results
 */
function searchDashboard(query, searchType, filters) {
  var results = [];
  var queryLower = query.toLowerCase();

  // Search members
  if (searchType === 'all' || searchType === 'members') {
    var members = getMemberList();
    members.forEach(function(m) {
      if (m.name.toLowerCase().indexOf(queryLower) !== -1 ||
          m.id.toLowerCase().indexOf(queryLower) !== -1 ||
          m.department.toLowerCase().indexOf(queryLower) !== -1) {
        results.push({
          id: m.id,
          type: 'member',
          title: m.name,
          subtitle: m.department + ' - ID: ' + m.id
        });
      }
    });
  }

  // Search grievances
  if (searchType === 'all' || searchType === 'grievances') {
    var grievances = getOpenGrievances();
    grievances.forEach(function(g) {
      var grievanceId = g['Grievance ID'] || '';
      var memberName = g['Member Name'] || '';
      var description = g['Description'] || '';
      var status = g['Status'] || '';

      if (grievanceId.toLowerCase().indexOf(queryLower) !== -1 ||
          memberName.toLowerCase().indexOf(queryLower) !== -1 ||
          description.toLowerCase().indexOf(queryLower) !== -1) {

        // Apply status filter
        if (filters.status && status.toLowerCase() !== filters.status.toLowerCase()) {
          return;
        }

        results.push({
          id: grievanceId,
          type: 'grievance',
          title: grievanceId + ' - ' + memberName,
          subtitle: status + ' - Step ' + (g['Current Step'] || 1)
        });
      }
    });
  }

  return results.slice(0, 50);
}

/**
 * Quick search for instant results
 * @param {string} query - Search query
 * @returns {Array<Object>} Quick search results (max 10)
 */
function quickSearchDashboard(query) {
  if (!query || query.length < 2) return [];

  var results = searchDashboard(query, 'all', {});
  return results.slice(0, 10);
}

/**
 * Advanced search with complex filters
 * @param {Object} filters - Search filters
 * @returns {Array<Object>} Search results
 */
function advancedSearch(filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search members if included
  if (filters.includeMembers) {
    var memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
    if (memberSheet) {
      var data = memberSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var matches = true;

        // Apply department filter
        if (filters.department && row[MEMBER_COLUMNS.DEPARTMENT] !== filters.department) {
          matches = false;
        }

        // Apply name filter
        if (filters.name && matches) {
          var fullName = (row[MEMBER_COLUMNS.FIRST_NAME] + ' ' + row[MEMBER_COLUMNS.LAST_NAME]).toLowerCase();
          if (fullName.indexOf(filters.name.toLowerCase()) === -1) {
            matches = false;
          }
        }

        // Apply steward filter
        if (filters.stewardOnly && matches) {
          if (row[MEMBER_COLUMNS.IS_STEWARD] !== 'Yes') {
            matches = false;
          }
        }

        if (matches && row[MEMBER_COLUMNS.ID]) {
          results.push({
            id: row[MEMBER_COLUMNS.ID],
            type: 'member',
            title: row[MEMBER_COLUMNS.FIRST_NAME] + ' ' + row[MEMBER_COLUMNS.LAST_NAME],
            subtitle: row[MEMBER_COLUMNS.DEPARTMENT],
            row: i + 1
          });
        }
      }
    }
  }

  // Search grievances if included
  if (filters.includeGrievances) {
    var grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    if (grievanceSheet) {
      var gData = grievanceSheet.getDataRange().getValues();

      for (var j = 1; j < gData.length; j++) {
        var gRow = gData[j];
        var gMatches = true;

        // Apply status filter
        if (filters.status && gRow[GRIEVANCE_COLUMNS.STATUS] !== filters.status) {
          gMatches = false;
        }

        // Apply date range filter
        if (filters.startDate && gMatches) {
          var filedDate = gRow[GRIEVANCE_COLUMNS.FILING_DATE];
          if (filedDate && new Date(filedDate) < new Date(filters.startDate)) {
            gMatches = false;
          }
        }

        if (filters.endDate && gMatches) {
          var endFiledDate = gRow[GRIEVANCE_COLUMNS.FILING_DATE];
          if (endFiledDate && new Date(endFiledDate) > new Date(filters.endDate)) {
            gMatches = false;
          }
        }

        if (gMatches && gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID]) {
          results.push({
            id: gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
            type: 'grievance',
            title: gRow[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
            subtitle: gRow[GRIEVANCE_COLUMNS.STATUS] + ' - ' + gRow[GRIEVANCE_COLUMNS.TYPE],
            row: j + 1
          });
        }
      }
    }
  }

  return results.slice(0, 100);
}

/**
 * Gets department list from calc sheet or direct query
 * @returns {Array<string>} Array of department names
 */
function getDepartmentList() {
  var formulaSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HIDDEN_SHEETS.CALC_FORMULAS);

  if (!formulaSheet) {
    var memberSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

    if (!memberSheet) return [];

    var data = memberSheet.getRange(2, MEMBER_COLUMNS.DEPARTMENT + 1,
      memberSheet.getLastRow() - 1, 1).getValues();

    var depts = {};
    data.forEach(function(row) {
      if (row[0]) depts[row[0]] = true;
    });

    return Object.keys(depts).sort();
  }

  var deptData = formulaSheet.getRange('B4:B').getValues();
  return deptData.filter(function(row) { return row[0]; }).map(function(row) { return row[0]; });
}

/**
 * Gets member list for dropdowns
 * @returns {Array<Object>} Array of member objects with id, name, department
 */
function getMemberList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLUMNS.ID]) {
      members.push({
        id: data[i][MEMBER_COLUMNS.ID],
        name: data[i][MEMBER_COLUMNS.FIRST_NAME] + ' ' + data[i][MEMBER_COLUMNS.LAST_NAME],
        department: data[i][MEMBER_COLUMNS.DEPARTMENT]
      });
    }
  }

  return members;
}

/**
 * Gets a value from a hidden calculation sheet
 * @param {string} sheetName - The hidden sheet name
 * @param {string} cellRef - The cell reference (e.g., 'B4')
 * @returns {*} The cell value or null if not found
 */
function getCalcValue(sheetName, cellRef) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Hidden sheet ' + sheetName + ' not found');
    return null;
  }

  return sheet.getRange(cellRef).getValue();
}
