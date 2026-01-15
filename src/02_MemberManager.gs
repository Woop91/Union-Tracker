/**
 * ============================================================================
 * 02_MemberManager.gs - Member Directory Operations
 * ============================================================================
 *
 * This module handles all member-related operations including:
 * - Member directory management
 * - Steward promotion/demotion
 * - Member ID generation and validation
 * - Member data sync
 *
 * @fileoverview Member directory operations and steward management
 * @version 3.6.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MEMBER DIRECTORY OPERATIONS
// ============================================================================

/**
 * Adds a new member to the Member Directory
 * @param {Object} memberData - Member information object
 * @returns {string} The generated Member ID
 */
function addMember(memberData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  // Generate Member ID if not provided
  var memberId = memberData.memberId || generateMemberID_(memberData.firstName, memberData.lastName);

  // Find next empty row
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;

  // Set member data
  sheet.getRange(newRow, MEMBER_COLS.MEMBER_ID).setValue(memberId);
  sheet.getRange(newRow, MEMBER_COLS.FIRST_NAME).setValue(memberData.firstName || '');
  sheet.getRange(newRow, MEMBER_COLS.LAST_NAME).setValue(memberData.lastName || '');
  sheet.getRange(newRow, MEMBER_COLS.EMAIL).setValue(memberData.email || '');
  sheet.getRange(newRow, MEMBER_COLS.PHONE).setValue(memberData.phone || '');
  sheet.getRange(newRow, MEMBER_COLS.JOB_TITLE).setValue(memberData.jobTitle || '');
  sheet.getRange(newRow, MEMBER_COLS.WORK_LOCATION).setValue(memberData.workLocation || '');
  sheet.getRange(newRow, MEMBER_COLS.UNIT).setValue(memberData.unit || '');

  return memberId;
}

/**
 * Updates an existing member's information
 * @param {string} memberId - The member ID to update
 * @param {Object} updateData - Fields to update
 */
function updateMember(memberId, updateData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  // Find the member row
  var data = sheet.getDataRange().getValues();
  var memberRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      memberRow = i + 1;
      break;
    }
  }

  if (memberRow === -1) {
    throw new Error('Member not found: ' + memberId);
  }

  // Update fields
  if (updateData.firstName) sheet.getRange(memberRow, MEMBER_COLS.FIRST_NAME).setValue(updateData.firstName);
  if (updateData.lastName) sheet.getRange(memberRow, MEMBER_COLS.LAST_NAME).setValue(updateData.lastName);
  if (updateData.email) sheet.getRange(memberRow, MEMBER_COLS.EMAIL).setValue(updateData.email);
  if (updateData.phone) sheet.getRange(memberRow, MEMBER_COLS.PHONE).setValue(updateData.phone);
  if (updateData.jobTitle) sheet.getRange(memberRow, MEMBER_COLS.JOB_TITLE).setValue(updateData.jobTitle);
  if (updateData.workLocation) sheet.getRange(memberRow, MEMBER_COLS.WORK_LOCATION).setValue(updateData.workLocation);
  if (updateData.unit) sheet.getRange(memberRow, MEMBER_COLS.UNIT).setValue(updateData.unit);
}

/**
 * Gets a member by their ID
 * @param {string} memberId - The member ID to find
 * @returns {Object|null} Member data object or null if not found
 */
function getMemberById(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      var member = {};
      for (var j = 0; j < headers.length; j++) {
        member[headers[j]] = data[i][j];
      }
      return member;
    }
  }

  return null;
}

/**
 * Searches members by name, email, or other fields
 * @param {string} query - Search query
 * @returns {Array} Array of matching member objects
 */
function searchMembers(query) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var results = [];
  var queryLower = query.toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var firstName = String(row[MEMBER_COLS.FIRST_NAME - 1] || '').toLowerCase();
    var lastName = String(row[MEMBER_COLS.LAST_NAME - 1] || '').toLowerCase();
    var email = String(row[MEMBER_COLS.EMAIL - 1] || '').toLowerCase();
    var memberId = String(row[MEMBER_COLS.MEMBER_ID - 1] || '').toLowerCase();

    if (firstName.indexOf(queryLower) !== -1 ||
        lastName.indexOf(queryLower) !== -1 ||
        email.indexOf(queryLower) !== -1 ||
        memberId.indexOf(queryLower) !== -1) {
      var member = {};
      for (var j = 0; j < headers.length; j++) {
        member[headers[j]] = row[j];
      }
      results.push(member);
    }
  }

  return results;
}

// ============================================================================
// MEMBER ID GENERATION
// ============================================================================

/**
 * Generates a unique Member ID based on name
 * Format: M + First 2 letters of first name + First 2 letters of last name + 3 digits
 * Example: MJOSM123
 * @param {string} firstName - Member's first name
 * @param {string} lastName - Member's last name
 * @returns {string} Generated Member ID
 * @private
 */
function generateMemberID_(firstName, lastName) {
  var prefix = 'M';
  var firstPart = (firstName || 'XX').substring(0, 2).toUpperCase();
  var lastPart = (lastName || 'XX').substring(0, 2).toUpperCase();
  var namePrefix = prefix + firstPart + lastPart;

  // Get existing IDs to avoid duplicates
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var existingIds = {};

  if (sheet && sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function(row) {
      if (row[0]) existingIds[row[0]] = true;
    });
  }

  // Find next available number
  for (var num = 100; num < 1000; num++) {
    var newId = namePrefix + num;
    if (!existingIds[newId]) {
      return newId;
    }
  }

  // Fallback with timestamp
  return namePrefix + String(Date.now()).slice(-3);
}

// ============================================================================
// STEWARD MANAGEMENT
// ============================================================================

/**
 * Gets all active stewards from the Member Directory
 * @returns {Array} Array of steward objects
 */
function getAllStewards() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var stewards = [];

  for (var i = 1; i < data.length; i++) {
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isSteward === 'Yes' || isSteward === true) {
      var steward = {};
      for (var j = 0; j < headers.length; j++) {
        steward[headers[j]] = data[i][j];
      }
      steward.rowNumber = i + 1;
      stewards.push(steward);
    }
  }

  return stewards;
}

/**
 * Gets steward workload statistics
 * @returns {Array} Array of steward workload objects
 */
function getStewardWorkload() {
  var stewards = getAllStewards();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return stewards;

  var grievances = grievanceSheet.getDataRange().getValues();

  stewards.forEach(function(steward) {
    var fullName = steward[Object.keys(steward).find(function(k) { return k.toLowerCase().indexOf('first') !== -1; })] + ' ' +
                   steward[Object.keys(steward).find(function(k) { return k.toLowerCase().indexOf('last') !== -1; })];

    var activeCases = 0;
    var totalCases = 0;
    var wonCases = 0;

    for (var i = 1; i < grievances.length; i++) {
      var assignedSteward = grievances[i][GRIEVANCE_COLS.STEWARD - 1];
      if (assignedSteward === fullName) {
        totalCases++;
        var status = grievances[i][GRIEVANCE_COLS.STATUS - 1];
        if (status === 'Open' || status === 'Pending Info') {
          activeCases++;
        }
        if (status === 'Won') {
          wonCases++;
        }
      }
    }

    steward.activeCases = activeCases;
    steward.totalCases = totalCases;
    steward.wonCases = wonCases;
    steward.winRate = totalCases > 0 ? Math.round((wonCases / totalCases) * 100) : 0;
  });

  return stewards;
}

// ============================================================================
// MEMBER DATA SYNC
// ============================================================================

/**
 * Syncs member data with grievance records
 * Updates grievance counts and status in Member Directory
 */
function syncMemberGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    Logger.log('Required sheets not found for sync');
    return;
  }

  var members = memberSheet.getDataRange().getValues();
  var grievances = grievanceSheet.getDataRange().getValues();

  // Build grievance count map
  var grievanceCounts = {};
  for (var i = 1; i < grievances.length; i++) {
    var memberId = grievances[i][GRIEVANCE_COLS.MEMBER_ID - 1];
    if (memberId) {
      if (!grievanceCounts[memberId]) {
        grievanceCounts[memberId] = { total: 0, active: 0 };
      }
      grievanceCounts[memberId].total++;
      var status = grievances[i][GRIEVANCE_COLS.STATUS - 1];
      if (status === 'Open' || status === 'Pending Info') {
        grievanceCounts[memberId].active++;
      }
    }
  }

  // Update member rows (if grievance count columns exist)
  if (MEMBER_COLS.TOTAL_GRIEVANCES && MEMBER_COLS.ACTIVE_GRIEVANCES) {
    for (var j = 1; j < members.length; j++) {
      var memberId = members[j][MEMBER_COLS.MEMBER_ID - 1];
      var counts = grievanceCounts[memberId] || { total: 0, active: 0 };
      memberSheet.getRange(j + 1, MEMBER_COLS.TOTAL_GRIEVANCES).setValue(counts.total);
      memberSheet.getRange(j + 1, MEMBER_COLS.ACTIVE_GRIEVANCES).setValue(counts.active);
    }
  }

  Logger.log('Member grievance data synced');
}
