/**
 * ============================================================================
 * 02_DataManagers.gs - Member & Grievance Data Operations
 * ============================================================================
 *
 * This module handles all member and grievance data operations including:
 * - Member directory management
 * - Steward promotion/demotion
 * - Member ID generation and validation
 * - Member data sync
 * - Grievance tracking and management
 *
 * @fileoverview Member directory and grievance data operations
 * @version 4.6.0
 * @requires 01_Core.gs
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
    if (isTruthyValue(isSteward)) {
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
 * Gets detailed steward workload statistics including win rates
 * NOTE: A simpler getStewardWorkload() exists in 11_SecureMemberDashboard.gs for dashboard use
 * This version provides detailed metrics: activeCases, totalCases, wonCases, winRate
 * @returns {Array} Array of steward workload objects with detailed metrics
 */
function getStewardWorkloadDetailed() {
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

// ============================================================================
// UNIT-BASED ID GENERATION (Strategic Command Center)
// ============================================================================

/**
 * Generates missing Member IDs for all members without one
 * Uses unit-based prefixes from COMMAND_CONFIG or Config sheet
 * Format: UNIT_CODE-SEQUENCE-H (e.g., MS-101-H)
 */
function generateMissingMemberIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory sheet not found');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var unitCodes = getUnitCodes_();
  var countAdded = 0;

  for (var i = 1; i < data.length; i++) {
    var currentId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    var unit = data[i][MEMBER_COLS.UNIT - 1];

    // If ID is blank but member has data
    if (!currentId && (unit || data[i][MEMBER_COLS.FIRST_NAME - 1])) {
      var prefix = unitCodes[unit] || 'GEN';
      var nextNum = getNextSequence_(prefix, sheet);
      var newId = prefix + '-' + nextNum + '-H';

      sheet.getRange(i + 1, MEMBER_COLS.MEMBER_ID).setValue(newId);
      countAdded++;
    }
  }

  ss.toast('Generated ' + countAdded + ' new Member IDs', COMMAND_CONFIG.SYSTEM_NAME, 5);
  return countAdded;
}

/**
 * Gets the next available sequence number for a given prefix
 * Scans existing IDs to find the highest number and increments
 * @param {string} prefix - The unit code prefix (e.g., "MS")
 * @param {Sheet} sheet - The Member Directory sheet
 * @returns {number} Next available sequence number
 * @private
 */
function getNextSequence_(prefix, sheet) {
  var ids = sheet.getRange(1, MEMBER_COLS.MEMBER_ID, sheet.getLastRow(), 1).getValues().flat();
  var max = 100;

  ids.forEach(function(id) {
    if (typeof id === 'string' && id.startsWith(prefix + '-')) {
      var parts = id.split('-');
      if (parts.length >= 2) {
        var n = parseInt(parts[1], 10);
        if (!isNaN(n) && n > max) {
          max = n;
        }
      }
    }
  });

  return max + 1;
}

/**
 * Checks for duplicate Member IDs and reports them
 */
function checkDuplicateMemberIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory sheet not found');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var idCounts = {};
  var duplicates = [];

  for (var i = 1; i < data.length; i++) {
    var id = data[i][MEMBER_COLS.MEMBER_ID - 1];
    if (id) {
      if (idCounts[id]) {
        idCounts[id].count++;
        idCounts[id].rows.push(i + 1);
      } else {
        idCounts[id] = { count: 1, rows: [i + 1] };
      }
    }
  }

  for (var memberId in idCounts) {
    if (idCounts[memberId].count > 1) {
      duplicates.push({
        id: memberId,
        count: idCounts[memberId].count,
        rows: idCounts[memberId].rows
      });
    }
  }

  if (duplicates.length === 0) {
    SpreadsheetApp.getUi().alert('No duplicate Member IDs found.');
  } else {
    var message = 'Found ' + duplicates.length + ' duplicate ID(s):\n\n';
    duplicates.forEach(function(dup) {
      message += 'ID: ' + dup.id + ' (rows: ' + dup.rows.join(', ') + ')\n';
    });
    SpreadsheetApp.getUi().alert(message);
  }

  return duplicates;
}

/**
 * Multi-Key Smart Match (v4.1)
 * Hierarchical matching across Member ID, Email, and Name.
 * Used during form submissions and bulk imports to prevent duplicate records.
 *
 * Match Priority:
 *   1. Exact Member ID match (highest confidence)
 *   2. Exact Email match (high confidence)
 *   3. Exact First + Last Name match (fallback)
 *
 * @param {Object} searchParams - Search parameters object
 * @param {string} [searchParams.memberId] - Member ID to search for
 * @param {string} [searchParams.email] - Email address to search for
 * @param {string} [searchParams.firstName] - First name to search for
 * @param {string} [searchParams.lastName] - Last name to search for
 * @param {Array<Array>} dataArray - 2D array of Member Directory data (batch processed)
 * @returns {Object|null} Match result with row number and match type, or null if no match
 *
 * @example
 * var data = sheet.getDataRange().getValues();
 * var match = findExistingMember({
 *   memberId: 'MS-101-H',
 *   email: 'john.doe@email.com',
 *   firstName: 'John',
 *   lastName: 'Doe'
 * }, data);
 *
 * if (match) {
 *   // Update existing record at match.row
 *   Logger.log('Found via ' + match.matchType + ' at row ' + match.row);
 * } else {
 *   // Create new record
 * }
 */
function findExistingMember(searchParams, dataArray) {
  var searchId = searchParams.memberId || '';
  var searchEmail = searchParams.email || '';
  var searchFirstName = searchParams.firstName || '';
  var searchLastName = searchParams.lastName || '';

  // Column indices (0-based for array access)
  var COL_ID = MEMBER_COLS.MEMBER_ID - 1;        // 0
  var COL_FIRST = MEMBER_COLS.FIRST_NAME - 1;    // 1
  var COL_LAST = MEMBER_COLS.LAST_NAME - 1;      // 2
  var COL_EMAIL = MEMBER_COLS.EMAIL - 1;         // 7

  // Normalize search values
  var normId = searchId.toString().trim();
  var normEmail = searchEmail.toString().trim().toLowerCase();
  var normFirst = searchFirstName.toString().trim().toLowerCase();
  var normLast = searchLastName.toString().trim().toLowerCase();

  // Track potential name match (lower priority)
  var nameMatch = null;

  for (var i = 1; i < dataArray.length; i++) {
    var row = dataArray[i];

    // Extract and normalize row values
    var rowId = row[COL_ID] ? row[COL_ID].toString().trim() : '';
    var rowEmail = row[COL_EMAIL] ? row[COL_EMAIL].toString().trim().toLowerCase() : '';
    var rowFirst = row[COL_FIRST] ? row[COL_FIRST].toString().trim().toLowerCase() : '';
    var rowLast = row[COL_LAST] ? row[COL_LAST].toString().trim().toLowerCase() : '';

    // Priority 1: Exact Member ID match (immediate return)
    if (normId && rowId && rowId === normId) {
      return {
        row: i + 1,  // Convert to 1-indexed sheet row
        matchType: 'MEMBER_ID',
        confidence: 'HIGH'
      };
    }

    // Priority 2: Exact Email match (immediate return)
    if (normEmail && rowEmail && rowEmail === normEmail) {
      return {
        row: i + 1,
        matchType: 'EMAIL',
        confidence: 'HIGH'
      };
    }

    // Priority 3: Name match (store but continue searching for higher-priority match)
    if (!nameMatch && normFirst && normLast && rowFirst && rowLast) {
      if (rowFirst === normFirst && rowLast === normLast) {
        nameMatch = {
          row: i + 1,
          matchType: 'NAME',
          confidence: 'MEDIUM'
        };
        // Don't return yet - keep searching for ID/Email match
      }
    }
  }

  // Return name match if found (no higher-priority match existed)
  if (nameMatch) {
    return nameMatch;
  }

  // No match found - safe to create new record
  return null;
}

// ============================================================================
// STEWARD PROMOTION/DEMOTION (Strategic Command Center)
// ============================================================================

/**
 * Promotes the currently selected member to Steward status
 * Updates IS_STEWARD column and adds to Config steward list
 * Requires two confirmation dialogs for safety
 */
function promoteSelectedMemberToSteward() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Please select a member row in the Member Directory sheet');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a member row (not the header)');
    return;
  }

  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
  var fullName = firstName + ' ' + lastName;
  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();

  if (isTruthyValue(currentStatus)) {
    ui.alert(fullName + ' is already a Steward');
    return;
  }

  // WARNING 1: Initial confirmation
  var response1 = ui.alert(
    '⬆️ Promote to Steward - Step 1 of 2',
    'You are about to promote ' + fullName + ' to Steward status.\n\n' +
    'This will:\n' +
    '• Set "Is Steward" to Yes\n' +
    '• Add to the Steward dropdown list\n' +
    '• Grant access to steward-level functions\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('Promotion cancelled.');
    return;
  }

  // WARNING 2: Final confirmation
  var response2 = ui.alert(
    '⚠️ Final Confirmation - Step 2 of 2',
    'PLEASE CONFIRM: You are promoting ' + fullName + ' to Steward.\n\n' +
    'This action grants significant responsibilities including:\n' +
    '• Representing members in grievances\n' +
    '• Access to sensitive member information\n' +
    '• Authority to act on behalf of the union\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    ui.alert('Promotion cancelled.');
    return;
  }

  // Update member record
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('Yes');

  // Add to Config steward list
  addToConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);

  ss.toast('✅ ' + fullName + ' has been promoted to Steward', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Demotes the currently selected steward back to regular member
 * Requires two confirmation dialogs for safety
 */
function demoteSelectedSteward() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Please select a steward row in the Member Directory sheet');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a member row (not the header)');
    return;
  }

  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
  var fullName = firstName + ' ' + lastName;
  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();

  if (!isTruthyValue(currentStatus)) {
    ui.alert(fullName + ' is not currently a Steward');
    return;
  }

  // WARNING 1: Initial confirmation
  var response1 = ui.alert(
    '⬇️ Demote Steward - Step 1 of 2',
    'You are about to remove Steward status from ' + fullName + '.\n\n' +
    'This will:\n' +
    '• Set "Is Steward" to No\n' +
    '• Remove from the Steward dropdown list\n' +
    '• Remove steward-level access\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('Demotion cancelled.');
    return;
  }

  // WARNING 2: Final confirmation
  var response2 = ui.alert(
    '⚠️ Final Confirmation - Step 2 of 2',
    'PLEASE CONFIRM: You are removing Steward status from ' + fullName + '.\n\n' +
    'This is a significant action that:\n' +
    '• Removes their authority to represent members\n' +
    '• Should be documented appropriately\n' +
    '• May require notification to the member\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    ui.alert('Demotion cancelled.');
    return;
  }

  // Update member record
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('No');

  // Remove from Config steward list
  removeFromConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);

  ss.toast('✅ ' + fullName + ' has been demoted from Steward', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Adds a value to a Config column dropdown list
 * @param {number} configCol - The Config column number
 * @param {string} value - The value to add
 * @private
 */
function addToConfigDropdown_(configCol, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) return;

  // Find first empty row in the column
  var colData = configSheet.getRange(3, configCol, configSheet.getLastRow() - 2, 1).getValues();
  var emptyRow = 3;

  for (var i = 0; i < colData.length; i++) {
    if (!colData[i][0]) {
      emptyRow = i + 3;
      break;
    }
    emptyRow = i + 4;
  }

  configSheet.getRange(emptyRow, configCol).setValue(value);
}

/**
 * Removes a value from a Config column dropdown list
 * @param {number} configCol - The Config column number
 * @param {string} value - The value to remove
 * @private
 */
function removeFromConfigDropdown_(configCol, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) return;

  var colData = configSheet.getRange(3, configCol, configSheet.getLastRow() - 2, 1).getValues();

  for (var i = 0; i < colData.length; i++) {
    if (colData[i][0] === value) {
      configSheet.getRange(i + 3, configCol).clearContent();
      break;
    }
  }
}

// ============================================================================
// BATCH PROCESSING (Performance Optimization)
// ============================================================================

/**
 * Updates member data in batch mode for better performance
 * Reads all data once, modifies in memory, writes back in one operation
 * @param {string} memberId - The member ID to update
 * @param {Object} newValuesObj - Object with field values to update
 */
function updateMemberDataBatch(memberId, newValuesObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  var range = sheet.getDataRange();
  var data = range.getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      // Modify array in memory
      if (newValuesObj.email !== undefined) {
        data[i][MEMBER_COLS.EMAIL - 1] = newValuesObj.email;
      }
      if (newValuesObj.phone !== undefined) {
        data[i][MEMBER_COLS.PHONE - 1] = newValuesObj.phone;
      }
      if (newValuesObj.firstName !== undefined) {
        data[i][MEMBER_COLS.FIRST_NAME - 1] = newValuesObj.firstName;
      }
      if (newValuesObj.lastName !== undefined) {
        data[i][MEMBER_COLS.LAST_NAME - 1] = newValuesObj.lastName;
      }
      if (newValuesObj.unit !== undefined) {
        data[i][MEMBER_COLS.UNIT - 1] = newValuesObj.unit;
      }
      if (newValuesObj.workLocation !== undefined) {
        data[i][MEMBER_COLS.WORK_LOCATION - 1] = newValuesObj.workLocation;
      }
      if (newValuesObj.isSteward !== undefined) {
        data[i][MEMBER_COLS.IS_STEWARD - 1] = newValuesObj.isSteward;
      }

      // Write the specific row back in one shot
      sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Member ' + memberId + ' updated via batch process',
        COMMAND_CONFIG.SYSTEM_NAME,
        3
      );
      return true;
    }
  }

  return false;
}

// ============================================================================
// IMPORT/EXPORT DIALOGS
// ============================================================================

/**
 * Shows the import members dialog
 * Allows importing members from CSV data with column mapping
 */
function showImportMembersDialog() {
  var html = HtmlService.createHtmlOutput(getImportMembersHtml_())
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Import Members');
}

/**
 * Generates HTML for the import members dialog
 * @returns {string} HTML content
 * @private
 */
function getImportMembersHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    getMobileOptimizedHead() +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", Roboto, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; padding: 20px; color: #F8FAFC; }' +
    'h3 { margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }' +
    '.step { margin-bottom: 20px; padding: 16px; background: rgba(255,255,255,0.05); border-radius: 8px; border: 1px solid #334155; }' +
    '.step-title { font-weight: 600; margin-bottom: 12px; color: #A78BFA; }' +
    'textarea { width: 100%; height: 120px; padding: 12px; border: 2px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-family: monospace; font-size: 12px; resize: vertical; }' +
    'textarea:focus { border-color: #7C3AED; outline: none; }' +
    '.mapping-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; max-height: 200px; overflow-y: auto; }' +
    '.mapping-row { display: flex; align-items: center; gap: 8px; padding: 6px; background: rgba(255,255,255,0.03); border-radius: 4px; }' +
    '.mapping-row label { flex: 1; font-size: 13px; color: #94A3B8; }' +
    'select { padding: 6px 10px; border: 1px solid #334155; border-radius: 4px; background: #1E293B; color: #F8FAFC; font-size: 12px; min-width: 140px; }' +
    '.btn-row { display: flex; gap: 10px; margin-top: 16px; }' +
    'button { padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; }' +
    '.btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; flex: 1; }' +
    '.btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }' +
    '.btn-primary:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }' +
    '.btn-secondary { background: #334155; color: #F8FAFC; }' +
    '.preview { margin-top: 12px; font-size: 12px; color: #64748B; }' +
    '.preview-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 11px; }' +
    '.preview-table th, .preview-table td { padding: 6px 8px; border: 1px solid #334155; text-align: left; }' +
    '.preview-table th { background: #334155; color: #F8FAFC; }' +
    '.preview-table td { background: rgba(255,255,255,0.02); }' +
    '.status { padding: 12px; border-radius: 6px; margin-top: 12px; font-size: 13px; }' +
    '.status.success { background: rgba(16,185,129,0.2); border: 1px solid #10B981; }' +
    '.status.error { background: rgba(239,68,68,0.2); border: 1px solid #EF4444; }' +
    '.help { font-size: 12px; color: #64748B; margin-top: 8px; }' +
    '</style>' +
    '</head><body>' +
    '<h3>📥 Import Members from CSV</h3>' +
    '' +
    '<div class="step">' +
    '  <div class="step-title">Step 1: Paste CSV Data</div>' +
    '  <textarea id="csvData" placeholder="Paste your CSV data here...&#10;&#10;Example:&#10;First Name,Last Name,Email,Phone,Job Title,Unit&#10;John,Doe,john@example.com,555-1234,Analyst,Main Station&#10;Jane,Smith,jane@example.com,555-5678,Manager,Field Ops"></textarea>' +
    '  <div class="help">Paste data from Excel, Google Sheets, or a CSV file. First row should be headers.</div>' +
    '</div>' +
    '' +
    '<div class="step" id="mappingStep" style="display:none;">' +
    '  <div class="step-title">Step 2: Map Columns</div>' +
    '  <div class="mapping-grid" id="mappingGrid"></div>' +
    '  <div class="preview" id="preview"></div>' +
    '</div>' +
    '' +
    '<div class="btn-row">' +
    '  <button class="btn-secondary" id="parseBtn" onclick="parseCSV()">Parse CSV</button>' +
    '  <button class="btn-primary" id="importBtn" onclick="importMembers()" disabled>Import Members</button>' +
    '  <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '</div>' +
    '' +
    '<div id="statusArea"></div>' +
    '' +
    '<script>' +
    'var parsedData = [];' +
    'var csvHeaders = [];' +
    'var targetFields = [' +
    '  {key: "firstName", label: "First Name", required: true},' +
    '  {key: "lastName", label: "Last Name", required: true},' +
    '  {key: "email", label: "Email", required: false},' +
    '  {key: "phone", label: "Phone", required: false},' +
    '  {key: "jobTitle", label: "Job Title", required: false},' +
    '  {key: "workLocation", label: "Work Location", required: false},' +
    '  {key: "unit", label: "Unit", required: false},' +
    '  {key: "supervisor", label: "Supervisor", required: false},' +
    '  {key: "manager", label: "Manager", required: false}' +
    '];' +
    '' +
    'function parseCSV() {' +
    '  var csvText = document.getElementById("csvData").value.trim();' +
    '  if (!csvText) { showStatus("Please paste CSV data first", true); return; }' +
    '  ' +
    '  var lines = csvText.split(/\\r?\\n/);' +
    '  if (lines.length < 2) { showStatus("CSV must have header row and at least one data row", true); return; }' +
    '  ' +
    '  csvHeaders = parseCSVLine(lines[0]);' +
    '  parsedData = [];' +
    '  for (var i = 1; i < lines.length; i++) {' +
    '    if (lines[i].trim()) parsedData.push(parseCSVLine(lines[i]));' +
    '  }' +
    '  ' +
    '  buildMappingUI();' +
    '  showPreview();' +
    '  document.getElementById("mappingStep").style.display = "block";' +
    '  document.getElementById("importBtn").disabled = false;' +
    '  showStatus("Parsed " + parsedData.length + " rows. Map columns and click Import.", false);' +
    '}' +
    '' +
    'function parseCSVLine(line) {' +
    '  var result = [];' +
    '  var current = "";' +
    '  var inQuotes = false;' +
    '  for (var i = 0; i < line.length; i++) {' +
    '    var c = line[i];' +
    '    if (c === \'"\' && (i === 0 || line[i-1] !== \'\\\\\')) { inQuotes = !inQuotes; }' +
    '    else if ((c === "," || c === "\\t") && !inQuotes) { result.push(current.trim()); current = ""; }' +
    '    else { current += c; }' +
    '  }' +
    '  result.push(current.trim());' +
    '  return result;' +
    '}' +
    '' +
    'function buildMappingUI() {' +
    '  var grid = document.getElementById("mappingGrid");' +
    '  grid.innerHTML = "";' +
    '  targetFields.forEach(function(field) {' +
    '    var row = document.createElement("div");' +
    '    row.className = "mapping-row";' +
    '    var label = document.createElement("label");' +
    '    label.textContent = field.label + (field.required ? " *" : "");' +
    '    var select = document.createElement("select");' +
    '    select.id = "map_" + field.key;' +
    '    select.innerHTML = "<option value=\\"\\">-- Skip --</option>";' +
    '    csvHeaders.forEach(function(h, idx) {' +
    '      var opt = document.createElement("option");' +
    '      opt.value = idx;' +
    '      opt.textContent = h;' +
    '      if (h.toLowerCase().replace(/[^a-z]/g, "").indexOf(field.key.toLowerCase()) !== -1 ||' +
    '          field.key.toLowerCase().indexOf(h.toLowerCase().replace(/[^a-z]/g, "")) !== -1) {' +
    '        opt.selected = true;' +
    '      }' +
    '      select.appendChild(opt);' +
    '    });' +
    '    row.appendChild(label);' +
    '    row.appendChild(select);' +
    '    grid.appendChild(row);' +
    '  });' +
    '}' +
    '' +
    'function showPreview() {' +
    '  var preview = document.getElementById("preview");' +
    '  var rows = parsedData.slice(0, 3);' +
    '  if (rows.length === 0) { preview.innerHTML = ""; return; }' +
    '  var html = "<strong>Preview (first " + rows.length + " rows):</strong><table class=\\"preview-table\\"><tr>";' +
    '  csvHeaders.forEach(function(h) { html += "<th>" + h + "</th>"; });' +
    '  html += "</tr>";' +
    '  rows.forEach(function(row) {' +
    '    html += "<tr>";' +
    '    row.forEach(function(cell) { html += "<td>" + (cell || "-") + "</td>"; });' +
    '    html += "</tr>";' +
    '  });' +
    '  html += "</table>";' +
    '  preview.innerHTML = html;' +
    '}' +
    '' +
    'function importMembers() {' +
    '  var mapping = {};' +
    '  targetFields.forEach(function(field) {' +
    '    var sel = document.getElementById("map_" + field.key);' +
    '    if (sel.value !== "") mapping[field.key] = parseInt(sel.value);' +
    '  });' +
    '  ' +
    '  if (mapping.firstName === undefined || mapping.lastName === undefined) {' +
    '    showStatus("First Name and Last Name are required mappings", true);' +
    '    return;' +
    '  }' +
    '  ' +
    '  document.getElementById("importBtn").disabled = true;' +
    '  document.getElementById("importBtn").textContent = "Importing...";' +
    '  ' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.success) {' +
    '        showStatus("Successfully imported " + result.imported + " members! " + (result.skipped > 0 ? "(" + result.skipped + " skipped as duplicates)" : ""), false);' +
    '        document.getElementById("importBtn").textContent = "Done!";' +
    '      } else {' +
    '        showStatus("Import failed: " + result.message, true);' +
    '        document.getElementById("importBtn").disabled = false;' +
    '        document.getElementById("importBtn").textContent = "Import Members";' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      showStatus("Error: " + e.message, true);' +
    '      document.getElementById("importBtn").disabled = false;' +
    '      document.getElementById("importBtn").textContent = "Import Members";' +
    '    })' +
    '    .importMembersFromData(parsedData, mapping);' +
    '}' +
    '' +
    'function escapeHtml(t){if(t==null)return"";return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#x27;").replace(/\\//g,"&#x2F;");}' +
    'function showStatus(msg, isError) {' +
    '  var area = document.getElementById("statusArea");' +
    '  area.innerHTML = "<div class=\\"status " + (isError ? "error" : "success") + "\\">" + escapeHtml(msg) + "</div>";' +
    '}' +
    '</script>' +
    '</body></html>';
}

/**
 * Imports members from parsed CSV data
 * @param {Array<Array>} data - 2D array of CSV data (without headers)
 * @param {Object} mapping - Column mapping object {fieldName: columnIndex}
 * @returns {Object} Result with imported count and any errors
 */
function importMembersFromData(data, mapping) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      return errorResponse('Member Directory sheet not found', 'bulkImportMembers');
    }

    // Get existing data for duplicate checking
    var existingData = sheet.getDataRange().getValues();
    var existingEmails = {};
    var existingNames = {};

    for (var i = 1; i < existingData.length; i++) {
      var email = (existingData[i][MEMBER_COLS.EMAIL - 1] || '').toString().toLowerCase().trim();
      var name = ((existingData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (existingData[i][MEMBER_COLS.LAST_NAME - 1] || '')).toLowerCase().trim();
      if (email) existingEmails[email] = true;
      if (name) existingNames[name] = true;
    }

    var imported = 0;
    var skipped = 0;
    var newRows = [];

    for (var j = 0; j < data.length; j++) {
      var row = data[j];

      var firstName = mapping.firstName !== undefined ? (row[mapping.firstName] || '').trim() : '';
      var lastName = mapping.lastName !== undefined ? (row[mapping.lastName] || '').trim() : '';
      var email = mapping.email !== undefined ? (row[mapping.email] || '').trim() : '';

      // Skip if no name
      if (!firstName && !lastName) {
        skipped++;
        continue;
      }

      // Check for duplicates
      var emailLower = email.toLowerCase();
      var nameLower = (firstName + ' ' + lastName).toLowerCase().trim();

      if ((emailLower && existingEmails[emailLower]) || existingNames[nameLower]) {
        skipped++;
        continue;
      }

      // Mark as existing to prevent duplicates within import batch
      if (emailLower) existingEmails[emailLower] = true;
      existingNames[nameLower] = true;

      // Generate Member ID
      var memberId = generateMemberID_(firstName, lastName);

      // Build new row with empty values for all columns (up to last column: STATE)
      var newRow = new Array(MEMBER_COLS.STATE).fill('');
      newRow[MEMBER_COLS.MEMBER_ID - 1] = memberId;
      newRow[MEMBER_COLS.FIRST_NAME - 1] = firstName;
      newRow[MEMBER_COLS.LAST_NAME - 1] = lastName;

      if (mapping.email !== undefined) newRow[MEMBER_COLS.EMAIL - 1] = row[mapping.email] || '';
      if (mapping.phone !== undefined) newRow[MEMBER_COLS.PHONE - 1] = row[mapping.phone] || '';
      if (mapping.jobTitle !== undefined) newRow[MEMBER_COLS.JOB_TITLE - 1] = row[mapping.jobTitle] || '';
      if (mapping.workLocation !== undefined) newRow[MEMBER_COLS.WORK_LOCATION - 1] = row[mapping.workLocation] || '';
      if (mapping.unit !== undefined) newRow[MEMBER_COLS.UNIT - 1] = row[mapping.unit] || '';
      if (mapping.supervisor !== undefined) newRow[MEMBER_COLS.SUPERVISOR - 1] = row[mapping.supervisor] || '';
      if (mapping.manager !== undefined) newRow[MEMBER_COLS.MANAGER - 1] = row[mapping.manager] || '';

      // Default Is Steward to No
      newRow[MEMBER_COLS.IS_STEWARD - 1] = 'No';

      newRows.push(newRow);
      imported++;
    }

    // Batch write all new rows
    if (newRows.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    // Log the import
    logAuditEvent(AUDIT_EVENTS.MEMBER_ADDED, {
      action: 'BULK_IMPORT',
      importedCount: imported,
      skippedCount: skipped,
      importedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      imported: imported,
      skipped: skipped,
      message: 'Import completed'
    };

  } catch (e) {
    console.error('Import error: ' + e.message);
    return errorResponse(e.message, 'bulkImportMembers');
  }
}

/**
 * Shows the export members dialog
 * Allows exporting members to CSV or Google Sheets
 */
function showExportMembersDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    ui.alert('Error', 'Member Directory not found.', ui.ButtonSet.OK);
    return;
  }

  var lastRow = sheet.getLastRow();
  var memberCount = lastRow > 1 ? lastRow - 1 : 0;

  var response = ui.alert('📤 Export Members',
    'Export ' + memberCount + ' members?\n\n' +
    'Options:\n' +
    '• Download as CSV: File > Download > Comma-separated values (.csv)\n' +
    '• Download as Excel: File > Download > Microsoft Excel (.xlsx)\n' +
    '• Copy to another sheet: Right-click the Member Directory tab > Copy to\n\n' +
    'Would you like to navigate to the Member Directory sheet now?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    ss.setActiveSheet(sheet);
    sheet.getRange('A1').activate();
  }
}



/**
 * ============================================================================
 * GrievanceManager.gs - Grievance Lifecycle Management
 * ============================================================================
 *
 * This module handles all grievance-related operations including:
 * - Creating new grievances
 * - Advancing grievance steps
 * - Deadline calculations based on Article 23A
 * - Status updates and lifecycle management
 * - Batch recalculation of deadlines
 *
 * SEPARATION OF CONCERNS: This file contains ONLY grievance business logic.
 * UI components are in 04a_UIMenus.gs, integrations in 05_Integrations.gs.
 *
 * @fileoverview Grievance lifecycle management
 * @version 4.6.0
 * @requires 01_Core.gs
 */

// ============================================================================
// GRIEVANCE CREATION
// ============================================================================

/**
 * Starts a new grievance for a member
 * @param {Object} grievanceData - Initial grievance data
 * @return {Object} Result with grievance ID or error
 */
function startNewGrievance(grievanceData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);

    if (!grievanceSheet) {
      throw new Error('Grievance Tracker sheet not found');
    }

    // Validate required fields
    const validation = validateGrievanceData(grievanceData);
    if (!validation.valid) {
      return errorResponse(validation.error, 'createGrievance');
    }

    // Generate new grievance ID
    const grievanceId = getNextGrievanceId(grievanceSheet);

    // Calculate initial deadlines
    const filingDate = grievanceData.filingDate ? new Date(grievanceData.filingDate) : new Date();
    const deadlines = calculateInitialDeadlines(filingDate);

    // Prepare row data
    const rowData = [
      grievanceId,                                    // Grievance ID
      grievanceData.memberId || '',                   // Member ID
      grievanceData.memberName || '',                 // Member Name
      filingDate,                                     // Filing Date
      grievanceData.grievanceType || '',              // Grievance Type
      grievanceData.articleViolated || '',            // Article Violated
      grievanceData.description || '',                // Description
      1,                                              // Current Step (starts at 1)
      filingDate,                                     // Step 1 Date
      deadlines.step1Due,                             // Step 1 Due
      'Pending',                                      // Step 1 Status
      '',                                             // Step 2 Date
      '',                                             // Step 2 Due
      '',                                             // Step 2 Status
      '',                                             // Step 3 Date
      '',                                             // Step 3 Due
      '',                                             // Step 3 Status
      '',                                             // Arbitration Date
      '',                                             // Resolution
      GRIEVANCE_OUTCOMES.PENDING,                     // Outcome
      '',                                             // Drive Folder
      grievanceData.notes || '',                      // Notes
      GRIEVANCE_STATUS.OPEN,                          // Status
      new Date()                                      // Last Updated
    ];

    // Append to sheet
    grievanceSheet.appendRow(rowData);

    // Log the creation
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_CREATED, {
      grievanceId: grievanceId,
      memberId: grievanceData.memberId,
      createdBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      message: `Grievance ${grievanceId} created successfully`
    };

  } catch (error) {
    console.error('Error creating grievance:', error);
    return errorResponse(error.message, 'createGrievance');
  }
}

/**
 * Handles grievance dialog form submission from UI modal
 * NOTE: Renamed from onGrievanceFormSubmit to avoid conflict with Google Form trigger handler
 * @param {Object} formData - Form data from UI dialog
 * @return {Object} Result object
 */
function handleGrievanceDialogSubmit(formData) {
  // Map form data to grievance data structure
  const grievanceData = {
    memberId: formData.memberId,
    memberName: formData.memberName,
    filingDate: formData.filingDate,
    grievanceType: formData.grievanceType,
    articleViolated: formData.articleViolated,
    description: formData.description,
    notes: formData.notes,
    actionType: formData.actionType || 'Grievance'
  };

  const result = startNewGrievance(grievanceData);

  if (result.success) {
    // Set the Action Type on the new row
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
      if (grievanceSheet && GRIEVANCE_COLS.ACTION_TYPE) {
        // Find the row with this grievance ID
        const lastRow = grievanceSheet.getLastRow();
        const grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, lastRow - 1, 1).getValues();
        for (let i = 0; i < grievanceIds.length; i++) {
          if (grievanceIds[i][0] === result.grievanceId) {
            grievanceSheet.getRange(i + 2, GRIEVANCE_COLS.ACTION_TYPE).setValue(grievanceData.actionType);
            break;
          }
        }
      }
    } catch (e) {
      console.log('Note: Could not set action type: ' + e.message);
    }

    // Create checklist from template
    if (typeof createChecklistFromTemplate === 'function' && formData.createChecklist !== false) {
      try {
        const checklistResult = createChecklistFromTemplate(
          result.grievanceId,
          grievanceData.actionType,
          grievanceData.grievanceType
        );
        if (checklistResult.success) {
          console.log('Checklist created with ' + checklistResult.count + ' items');
        }
      } catch (e) {
        console.log('Note: Checklist creation skipped: ' + e.message);
      }
    }

    // Optionally create Drive folder
    if (formData.createFolder) {
      setupDriveFolderForGrievance(result.grievanceId);
    }

    // Optionally sync to calendar
    if (formData.syncCalendar) {
      syncSingleGrievanceToCalendar(result.grievanceId);
    }

    showToast(`Grievance ${result.grievanceId} created!`, 'Success');
  }

  return result;
}

/**
 * Validates grievance data before creation
 * @param {Object} data - Grievance data to validate
 * @return {Object} Validation result
 */
function validateGrievanceData(data) {
  if (!data.memberName && !data.memberId) {
    return { valid: false, error: 'Member name or ID is required' };
  }

  if (!data.grievanceType) {
    return { valid: false, error: 'Grievance type is required' };
  }

  if (!data.description || data.description.trim().length < 10) {
    return { valid: false, error: 'Description must be at least 10 characters' };
  }

  return { valid: true };
}

/**
 * Generates a grievance ID from a sequence number
 * @param {number} sequence - The sequence number
 * @return {string} Formatted grievance ID (e.g. GRV-2026-0001)
 */
function generateGrievanceId(sequence) {
  var year = new Date().getFullYear();
  return 'GRV-' + year + '-' + String(sequence).padStart(4, '0');
}

/**
 * Gets the next available grievance ID
 * @param {Sheet} sheet - Grievance tracker sheet
 * @return {string} Next grievance ID
 */
function getNextGrievanceId(sheet) {
  const data = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  let maxSequence = 0;

  // Find highest sequence number for current year
  for (let i = 1; i < data.length; i++) {
    const id = data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID];
    if (id && typeof id === 'string') {
      const match = id.match(/^GRV-(\d{4})-(\d{4})$/);
      if (match && parseInt(match[1]) === currentYear) {
        maxSequence = Math.max(maxSequence, parseInt(match[2]));
      }
    }
  }

  return generateGrievanceId(maxSequence + 1);
}

// ============================================================================
// DEADLINE CALCULATIONS - Based on Article 23A
// ============================================================================

/**
 * Calculates initial deadlines for a new grievance
 * @param {Date} filingDate - The grievance filing date
 * @return {Object} Calculated deadline dates
 */
function calculateInitialDeadlines(filingDate) {
  const step1Due = addBusinessDays(filingDate, DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);

  return {
    step1Due: step1Due
  };
}

/**
 * Calculates deadline for advancing to next step
 * @param {number} currentStep - Current grievance step (1-3)
 * @param {Date} currentStepDate - Date of current step
 * @return {Date} Deadline for next step
 */
function calculateNextStepDeadline(currentStep, currentStepDate) {
  let daysToAdd;

  switch (currentStep) {
    case 1:
      daysToAdd = DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL;
      break;
    case 2:
      daysToAdd = DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL;
      break;
    case 3:
      daysToAdd = DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND;
      break;
    default:
      return null;
  }

  return addBusinessDays(currentStepDate, daysToAdd);
}

/**
 * Calculates response deadline for a given step
 * @param {number} step - Grievance step
 * @param {Date} stepDate - Date step was initiated
 * @return {Date} Response deadline
 */
function calculateResponseDeadline(step, stepDate) {
  let daysToAdd;

  switch (step) {
    case 1:
      daysToAdd = DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE;
      break;
    case 2:
      daysToAdd = DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE;
      break;
    case 3:
      daysToAdd = DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE;
      break;
    default:
      return null;
  }

  return addBusinessDays(stepDate, daysToAdd);
}

/**
 * Adds business days to a date (excludes weekends)
 * @param {Date} startDate - Starting date
 * @param {number} days - Number of business days to add
 * @return {Date} Resulting date
 */
function addBusinessDays(startDate, days) {
  const result = new Date(startDate);
  let addedDays = 0;

  while (addedDays < days) {
    result.setDate(result.getDate() + 1);
    const dayOfWeek = result.getDay();
    // Skip weekends (0 = Sunday, 6 = Saturday)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      addedDays++;
    }
  }

  return result;
}

/**
 * Calculates days remaining until a deadline
 * @param {Date} deadline - Deadline date
 * @return {number} Days remaining (negative if overdue)
 */
function getDaysUntilDeadline(deadline) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const deadlineDate = new Date(deadline);
  deadlineDate.setHours(0, 0, 0, 0);

  const diffTime = deadlineDate - today;
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

// ============================================================================
// STEP ADVANCEMENT
// ============================================================================

/**
 * Advances a grievance to the next step
 * @param {string} grievanceId - The grievance ID to advance
 * @param {Object} options - Additional options (response, notes, etc.)
 * @return {Object} Result object
 */
function advanceGrievanceStep(grievanceId, options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    const data = sheet.getDataRange().getValues();

    // Find the grievance row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
        rowIndex = i + 1; // 1-indexed for sheet operations
        break;
      }
    }

    if (rowIndex === -1) {
      return errorResponse('Grievance not found', 'advanceGrievanceStep');
    }

    const currentStep = Number(data[rowIndex - 1][GRIEVANCE_COLUMNS.CURRENT_STEP]);
    if (isNaN(currentStep) || currentStep < 1) {
      return errorResponse('Invalid current step value for this grievance', 'advanceGrievanceStep');
    }
    const nextStep = currentStep + 1;

    if (nextStep > 4) {
      return errorResponse('Grievance is already at arbitration level', 'advanceGrievanceStep');
    }

    const today = new Date();
    const responseDue = calculateResponseDeadline(nextStep, today);

    // Batch all updates into a single row write where possible
    const currentStepStatusCol = getStepStatusColumn(currentStep);
    sheet.getRange(rowIndex, currentStepStatusCol).setValue(options.currentStepOutcome || 'Appealed');

    // Collect column updates to batch write
    var updates = [];
    updates.push({ col: GRIEVANCE_COLUMNS.CURRENT_STEP + 1, val: nextStep });
    updates.push({ col: GRIEVANCE_COLUMNS.STATUS + 1, val: nextStep === 4 ? GRIEVANCE_STATUS.AT_ARBITRATION : GRIEVANCE_STATUS.APPEALED });
    updates.push({ col: GRIEVANCE_COLUMNS.LAST_UPDATED + 1, val: today });

    if (nextStep <= 3) {
      const nextStepDateCol = getStepDateColumn(nextStep);
      updates.push({ col: nextStepDateCol, val: today });
      updates.push({ col: nextStepDateCol + 1, val: responseDue });
      updates.push({ col: nextStepDateCol + 2, val: 'Pending' });
    } else {
      updates.push({ col: GRIEVANCE_COLUMNS.ARBITRATION_DATE + 1, val: today });
    }

    // Add notes if provided
    if (options.notes) {
      const existingNotes = data[rowIndex - 1][GRIEVANCE_COLUMNS.NOTES] || '';
      const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                       `[${timestamp}] Step ${currentStep} -> ${nextStep}: ${options.notes}`;
      updates.push({ col: GRIEVANCE_COLUMNS.NOTES + 1, val: newNotes });
    }

    // Write all collected updates
    for (var u = 0; u < updates.length; u++) {
      sheet.getRange(rowIndex, updates[u].col).setValue(updates[u].val);
    }

    // Log the advancement
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_STEP_ADVANCED, {
      grievanceId: grievanceId,
      fromStep: currentStep,
      toStep: nextStep,
      advancedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      newStep: nextStep,
      message: `Grievance advanced to Step ${nextStep}`
    };

  } catch (error) {
    console.error('Error advancing grievance:', error);
    return errorResponse(error.message, 'advanceGrievanceStep');
  }
}

/**
 * Gets column index for step date
 * @param {number} step - Step number
 * @return {number} 1-indexed column number
 */
function getStepDateColumn(step) {
  switch (step) {
    case 1: return GRIEVANCE_COLUMNS.STEP_1_DATE + 1;
    case 2: return GRIEVANCE_COLUMNS.STEP_2_DATE + 1;
    case 3: return GRIEVANCE_COLUMNS.STEP_3_DATE + 1;
    default: return null;
  }
}

/**
 * Gets column index for step status
 * @param {number} step - Step number
 * @return {number} 1-indexed column number
 */
function getStepStatusColumn(step) {
  switch (step) {
    case 1: return GRIEVANCE_COLUMNS.STEP_1_STATUS + 1;
    case 2: return GRIEVANCE_COLUMNS.STEP_2_STATUS + 1;
    case 3: return GRIEVANCE_COLUMNS.STEP_3_STATUS + 1;
    default: return null;
  }
}

// ============================================================================
// BATCH OPERATIONS
// ============================================================================

/**
 * Recalculates all grievance deadlines in batches
 * Used when deadline rules change or for data repair
 * @return {Object} Result with count of updated grievances
 */
function recalcAllGrievancesBatched() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  let updatedCount = 0;
  let batchCount = 0;
  const startTime = new Date().getTime();

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    // Check execution time limit
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
      console.log('Approaching time limit, stopping batch');
      break;
    }

    // Batch pause
    if (batchCount >= BATCH_LIMITS.MAX_ROWS_PER_BATCH) {
      Utilities.sleep(BATCH_LIMITS.PAUSE_BETWEEN_BATCHES_MS);
      batchCount = 0;
    }

    const row = data[i];
    const status = row[GRIEVANCE_COLUMNS.STATUS];

    // Only recalculate open/pending grievances
    if (status === GRIEVANCE_STATUS.OPEN ||
        status === GRIEVANCE_STATUS.PENDING ||
        status === GRIEVANCE_STATUS.APPEALED) {

      const currentStep = row[GRIEVANCE_COLUMNS.CURRENT_STEP];
      const stepDate = row[getStepDateColumn(currentStep) - 1]; // 0-indexed for data array

      if (stepDate instanceof Date) {
        const newDue = calculateResponseDeadline(currentStep, stepDate);
        const dueColumn = getStepDateColumn(currentStep) + 1; // Due is next column after date

        sheet.getRange(i + 1, dueColumn).setValue(newDue);
        updatedCount++;
      }
    }

    batchCount++;
  }

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'RECALC_DEADLINES',
    grievancesUpdated: updatedCount,
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    updatedCount: updatedCount,
    message: `Recalculated deadlines for ${updatedCount} grievances`
  };
}

/**
 * Updates status for multiple grievances at once
 * @param {string[]} grievanceIds - Array of grievance IDs
 * @param {string} newStatus - New status to set
 * @param {string} notes - Optional notes to add
 * @return {Object} Result object
 */
function bulkUpdateGrievanceStatus(grievanceIds, newStatus, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  let updatedCount = 0;
  const today = new Date();
  const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

  for (let i = 1; i < data.length; i++) {
    const grievanceId = data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID];

    if (grievanceIds.includes(grievanceId)) {
      const rowIndex = i + 1;

      // Update status
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(newStatus);
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(today);

      // Add notes if provided
      if (notes) {
        const existingNotes = data[i][GRIEVANCE_COLUMNS.NOTES] || '';
        const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                         `[${timestamp}] Bulk status update to "${newStatus}": ${notes}`;
        sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(newNotes);
      }

      updatedCount++;
    }
  }

  return {
    success: true,
    updatedCount: updatedCount,
    message: `Updated status for ${updatedCount} grievances`
  };
}

// ============================================================================
// GRIEVANCE QUERIES
// ============================================================================

/**
 * Gets grievance data by ID
 * @param {string} grievanceId - The grievance ID
 * @return {Object|null} Grievance data or null if not found
 */
function getGrievanceById(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
      const grievance = {};
      headers.forEach((header, index) => {
        grievance[header] = data[i][index];
      });
      return grievance;
    }
  }

  return null;
}

/**
 * Gets all open grievances
 * @return {Array} Array of open grievance objects
 */
function getOpenGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const openStatuses = [
    GRIEVANCE_STATUS.OPEN,
    GRIEVANCE_STATUS.PENDING,
    GRIEVANCE_STATUS.APPEALED,
    GRIEVANCE_STATUS.AT_ARBITRATION
  ];

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][GRIEVANCE_COLUMNS.STATUS];
    if (openStatuses.includes(status)) {
      const grievance = {};
      headers.forEach((header, index) => {
        grievance[header] = data[i][index];
      });
      results.push(grievance);
    }
  }

  return results;
}

/**
 * Gets grievances with upcoming deadlines
 * @param {number} daysAhead - Number of days to look ahead
 * @return {Array} Array of grievances with deadline info
 */
function getUpcomingDeadlines(daysAhead) {
  const openGrievances = getOpenGrievances();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() + daysAhead);

  const upcoming = [];

  openGrievances.forEach(g => {
    const currentStep = g['Current Step'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.CURRENT_STEP]];
    let deadline;

    // Get the due date for current step
    switch (currentStep) {
      case 1:
        deadline = g['Step 1 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_1_DUE]];
        break;
      case 2:
        deadline = g['Step 2 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_2_DUE]];
        break;
      case 3:
        deadline = g['Step 3 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_3_DUE]];
        break;
    }

    if (deadline instanceof Date) {
      const deadlineDate = new Date(deadline);
      deadlineDate.setHours(0, 0, 0, 0);

      if (deadlineDate <= cutoffDate) {
        const daysLeft = Math.ceil((deadlineDate - today) / (1000 * 60 * 60 * 24));
        upcoming.push({
          grievanceId: g['Grievance ID'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.GRIEVANCE_ID]],
          memberName: g['Member Name'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.MEMBER_NAME]],
          step: `Step ${currentStep}`,
          deadline: deadline,
          date: Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          daysLeft: daysLeft
        });
      }
    }
  });

  // Sort by deadline (soonest first)
  upcoming.sort((a, b) => a.deadline - b.deadline);

  return upcoming;
}

/**
 * Gets grievance statistics for dashboard
 * Includes category breakdown for charts
 * @return {Object} Statistics object with categoryData for Google Charts
 */
function getGrievanceStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      open: 0,
      pending: 0,
      resolved: 0,
      closedThisYear: 0,
      total: 0,
      categoryData: [['Category', 'Count'], ['No Data', 1]]
    };
  }

  const data = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  let open = 0;
  let pending = 0;
  let resolved = 0;
  let closedThisYear = 0;
  let won = 0;

  // Track categories for chart data
  const categoryCounts = {};

  for (let i = 1; i < data.length; i++) {
    const status = data[i][GRIEVANCE_COLUMNS.STATUS];
    const lastUpdated = data[i][GRIEVANCE_COLUMNS.LAST_UPDATED];
    const category = data[i][GRIEVANCE_COLUMNS.ISSUE_CATEGORY] || 'Other';

    // Count by category
    categoryCounts[category] = (categoryCounts[category] || 0) + 1;

    switch (status) {
      case GRIEVANCE_STATUS.OPEN:
        open++;
        break;
      case GRIEVANCE_STATUS.PENDING:
      case GRIEVANCE_STATUS.APPEALED:
        pending++;
        break;
      case GRIEVANCE_STATUS.WON:
        won++;
        resolved++;
        if (lastUpdated instanceof Date && lastUpdated.getFullYear() === currentYear) {
          closedThisYear++;
        }
        break;
      case GRIEVANCE_STATUS.RESOLVED:
      case GRIEVANCE_STATUS.SETTLED:
      case GRIEVANCE_STATUS.CLOSED:
        if (lastUpdated instanceof Date && lastUpdated.getFullYear() === currentYear) {
          closedThisYear++;
        }
        resolved++;
        break;
    }
  }

  // Build categoryData array for Google Charts
  // Format: [['Category', 'Count'], ['Discipline', 5], ['Scheduling', 3], ...]
  const categoryData = [['Category', 'Count']];
  const sortedCategories = Object.keys(categoryCounts)
    .sort((a, b) => categoryCounts[b] - categoryCounts[a])
    .slice(0, 6); // Top 6 categories for cleaner charts

  sortedCategories.forEach(cat => {
    categoryData.push([cat, categoryCounts[cat]]);
  });

  // Ensure we have at least one data point
  if (categoryData.length === 1) {
    categoryData.push(['No Data', 0]);
  }

  return {
    open: open,
    pending: pending,
    resolved: resolved,
    won: won,
    closedThisYear: closedThisYear,
    total: data.length - 1,
    categoryData: categoryData
  };
}

// ============================================================================
// GRIEVANCE RESOLUTION
// ============================================================================

/**
 * Resolves a grievance with outcome
 * @param {string} grievanceId - The grievance ID
 * @param {string} outcome - Resolution outcome
 * @param {string} resolution - Resolution description
 * @param {string} notes - Additional notes
 * @return {Object} Result object
 */
function resolveGrievance(grievanceId, outcome, resolution, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return errorResponse('Grievance not found');
    }

    const today = new Date();
    const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

    // Update resolution fields
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.RESOLUTION + 1).setValue(resolution);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.OUTCOME + 1).setValue(outcome);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(GRIEVANCE_STATUS.RESOLVED);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(today);

    // Mark current step as completed
    const currentStep = data[rowIndex - 1][GRIEVANCE_COLUMNS.CURRENT_STEP];
    const stepStatusCol = getStepStatusColumn(currentStep);
    if (stepStatusCol) {
      sheet.getRange(rowIndex, stepStatusCol).setValue(outcome);
    }

    // Add resolution notes
    if (notes) {
      const existingNotes = data[rowIndex - 1][GRIEVANCE_COLUMNS.NOTES] || '';
      const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                       `[${timestamp}] RESOLVED - ${outcome}: ${notes}`;
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(newNotes);
    }

    // Log the resolution
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      action: 'RESOLVED',
      outcome: outcome,
      resolvedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      message: `Grievance ${grievanceId} resolved as "${outcome}"`
    };

  } catch (error) {
    console.error('Error resolving grievance:', error);
    return errorResponse(error.message);
  }
}

// ============================================================================
// UI DIALOG TRIGGERS
// ============================================================================

/**
 * Shows new grievance dialog
 */
function showNewGrievanceDialog() {
  const html = HtmlService.createHtmlOutput(getNewGrievanceFormHtml())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);

  SpreadsheetApp.getUi().showModalDialog(html, 'New Grievance');
}

/**
 * Shows edit grievance dialog for selected row
 */
function showEditGrievanceDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.GRIEVANCE_TRACKER) {
    showAlert('Please select a grievance in the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    showAlert('Please select a grievance row', 'No Selection');
    return;
  }

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const grievanceId = data[GRIEVANCE_COLUMNS.GRIEVANCE_ID];

  const html = HtmlService.createHtmlOutput(getEditGrievanceFormHtml(grievanceId))
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);

  SpreadsheetApp.getUi().showModalDialog(html, `Edit Grievance: ${grievanceId}`);
}

/**
 * Shows bulk status update dialog
 */
function showBulkStatusUpdate() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.GRIEVANCE_TRACKER) {
    showAlert('Please use this from the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const openGrievances = getOpenGrievances();
  const items = openGrievances.map(g => ({
    id: g['Grievance ID'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.GRIEVANCE_ID]],
    label: `${g['Grievance ID']} - ${g['Member Name']}`,
    selected: false
  }));

  showMultiSelectDialog('Select Grievances to Update', items, 'handleBulkStatusSelection');
}

/**
 * Generates HTML for new grievance form
 * @return {string} HTML content
 */
function getNewGrievanceFormHtml() {
  // Get member list for dropdown
  const members = getMemberList();
  const memberOptions = members.map(m =>
    `<option value="${m.id}">${m.name} (${m.department})</option>`
  ).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .form-actions { margin-top: 20px; display: flex; justify-content: flex-end; gap: 10px; }
        .checkbox-row { display: flex; gap: 20px; margin-top: 15px; }
        .checkbox-label { display: flex; align-items: center; gap: 8px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <form id="grievanceForm">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Member *</label>
              <select class="form-select" id="memberId" required>
                <option value="">Select a member...</option>
                ${memberOptions}
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Filing Date *</label>
              <input type="date" class="form-input" id="filingDate" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Action Type *</label>
              <select class="form-select" id="actionType" required onchange="updateFormForActionType()">
                <option value="Grievance">Grievance</option>
                <option value="Records Request">Records Request (Art. 21)</option>
                <option value="Information Request">Union Information Request</option>
                <option value="Weingarten">Weingarten Documentation</option>
                <option value="ULP Filing">Unfair Labor Practice (DLR)</option>
                <option value="EEOC/MCAD">EEOC/MCAD Complaint</option>
                <option value="Accommodation">ADA/Reasonable Accommodation</option>
                <option value="Other Admin">Other Administrative Action</option>
              </select>
            </div>
            <div class="form-group" id="grievanceTypeGroup">
              <label class="form-label">Issue Category *</label>
              <select class="form-select" id="grievanceType" required>
                <option value="">Select category...</option>
                <option value="Contract Violation">Contract Violation</option>
                <option value="Discipline">Discipline</option>
                <option value="Workload">Workload</option>
                <option value="Scheduling">Scheduling</option>
                <option value="Pay">Pay</option>
                <option value="Benefits">Benefits</option>
                <option value="Safety">Safety</option>
                <option value="Harassment">Harassment</option>
                <option value="Discrimination">Discrimination</option>
                <option value="Other">Other</option>
              </select>
            </div>
          </div>

          <div class="form-row" id="articleGroup">
            <div class="form-group">
              <label class="form-label">Article Violated</label>
              <input type="text" class="form-input" id="articleViolated"
                     placeholder="e.g., Article 23A, Section 5">
            </div>
            <div class="form-group"></div>
          </div>

          <div class="form-group">
            <label class="form-label">Description *</label>
            <textarea class="form-textarea" id="description" required
                      placeholder="Describe the grievance in detail..."></textarea>
          </div>

          <div class="form-group">
            <label class="form-label">Initial Notes</label>
            <textarea class="form-textarea" id="notes" rows="2"
                      placeholder="Any additional notes..."></textarea>
          </div>

          <div class="checkbox-row">
            <label class="checkbox-label">
              <input type="checkbox" id="createChecklist" checked>
              Create checklist from template
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="createFolder" checked>
              Create Drive folder
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="syncCalendar" checked>
              Sync to calendar
            </label>
          </div>

          <div class="form-actions">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="submit" class="btn btn-primary">Create Grievance</button>
          </div>
        </form>
      </div>

      <script>
        // Set default date to today
        document.getElementById('filingDate').valueAsDate = new Date();

        // Update form visibility based on action type
        function updateFormForActionType() {
          const actionType = document.getElementById('actionType').value;
          const isGrievance = actionType === 'Grievance';

          // Show/hide Issue Category (only for Grievances)
          document.getElementById('grievanceTypeGroup').style.display = isGrievance ? 'block' : 'none';
          document.getElementById('grievanceType').required = isGrievance;

          // Show/hide Articles Violated (only for Grievances)
          document.getElementById('articleGroup').style.display = isGrievance ? 'flex' : 'none';

          // Update submit button text
          const btn = document.querySelector('button[type="submit"]');
          btn.textContent = isGrievance ? 'Create Grievance' : 'Create ' + actionType;

          // Update description placeholder based on action type
          const desc = document.getElementById('description');
          if (isGrievance) {
            desc.placeholder = 'Describe the grievance in detail...';
          } else if (actionType === 'Records Request') {
            desc.placeholder = 'Describe what records are being requested...';
          } else if (actionType === 'Weingarten') {
            desc.placeholder = 'Describe the meeting and representation provided...';
          } else if (actionType === 'ULP Filing') {
            desc.placeholder = 'Describe the unfair labor practice...';
          } else {
            desc.placeholder = 'Describe the action or request in detail...';
          }
        }

        // Initialize form state
        updateFormForActionType();

        document.getElementById('grievanceForm').addEventListener('submit', function(e) {
          e.preventDefault();

          const memberSelect = document.getElementById('memberId');
          const actionType = document.getElementById('actionType').value;
          const isGrievance = actionType === 'Grievance';

          const formData = {
            memberId: memberSelect.value,
            memberName: memberSelect.options[memberSelect.selectedIndex].text.split(' (')[0],
            filingDate: document.getElementById('filingDate').value,
            actionType: actionType,
            grievanceType: isGrievance ? document.getElementById('grievanceType').value : actionType,
            articleViolated: document.getElementById('articleViolated').value,
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value,
            createChecklist: document.getElementById('createChecklist').checked,
            createFolder: document.getElementById('createFolder').checked,
            syncCalendar: document.getElementById('syncCalendar').checked
          };

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                google.script.host.close();
              } else {
                alert('Error: ' + result.error);
              }
            })
            .withFailureHandler(function(e) {
              alert('Error: ' + e.message);
            })
            .handleGrievanceDialogSubmit(formData);
        });
      </script>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for edit grievance form
 * @param {string} grievanceId - The grievance ID to edit
 * @return {string} HTML content
 */
function getEditGrievanceFormHtml(grievanceId) {
  const grievance = getGrievanceById(grievanceId);

  if (!grievance) {
    return '<p>Grievance not found</p>';
  }

  // Similar to new form but pre-populated with existing data
  // Abbreviated for space - would include full edit form
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-actions { margin-top: 20px; display: flex; justify-content: flex-end; gap: 10px; }
        .status-badge { padding: 4px 12px; border-radius: 12px; font-size: 12px; }
        .status-open { background: #e8f0fe; color: #1967d2; }
        .status-pending { background: #fef7e0; color: #ea8600; }
        .info-box { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <div class="info-box">
          <strong>Grievance ID:</strong> ${grievanceId}<br>
          <strong>Current Step:</strong> ${grievance['Current Step'] || 1}<br>
          <strong>Status:</strong> <span class="status-badge status-open">
            ${grievance['Status'] || 'Open'}</span>
        </div>

        <form id="editForm">
          <div class="form-group">
            <label class="form-label">Description</label>
            <textarea class="form-textarea" id="description">${grievance['Description'] || ''}</textarea>
          </div>

          <div class="form-group">
            <label class="form-label">Notes</label>
            <textarea class="form-textarea" id="notes">${grievance['Notes'] || ''}</textarea>
          </div>

          <div class="form-actions">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="button" class="btn btn-danger" onclick="advanceStep()">
              Advance Step
            </button>
            <button type="submit" class="btn btn-primary">Save Changes</button>
          </div>
        </form>
      </div>

      <script>
        const grievanceId = '${grievanceId}';

        document.getElementById('editForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const updates = {
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value
          };
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .updateGrievance(grievanceId, updates);
        });

        function advanceStep() {
          if (confirm('Advance this grievance to the next step?')) {
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  alert(result.message);
                  google.script.host.close();
                } else {
                  alert('Error: ' + result.error);
                }
              })
              .advanceGrievanceStep(grievanceId, {});
          }
        }
      </script>
    </body>
    </html>
  `;
}

// ============================================================================
// TRAFFIC LIGHT INDICATORS (Strategic Command Center)
// ============================================================================

/**
 * Applies visual traffic light indicators to grievance rows
 * Colors the ID column based on days to deadline:
 * - Red: Overdue (< 0 days)
 * - Orange/Yellow: Urgent (1-3 days)
 * - Green: On track (> 3 days)
 */
function applyTrafficLightIndicators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Grievance Log sheet not found');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No grievances to process', COMMAND_CONFIG.SYSTEM_NAME, 3);
    return;
  }

  // Read all days to deadline values
  var daysRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);
  var daysValues = daysRange.getValues();

  // Build color array
  var colors = [];
  var statusCounts = { overdue: 0, urgent: 0, onTrack: 0 };

  for (var i = 0; i < daysValues.length; i++) {
    var days = daysValues[i][0];

    if (typeof days !== 'number' || isNaN(days)) {
      colors.push([null]);  // No color for empty/invalid cells
    } else if (days < 0) {
      colors.push([COLORS.STATUS_RED]);  // Red - Overdue
      statusCounts.overdue++;
    } else if (days <= 3) {
      colors.push([COLORS.STATUS_ORANGE]);  // Orange - Urgent
      statusCounts.urgent++;
    } else {
      colors.push([COLORS.STATUS_GREEN]);  // Green - On track
      statusCounts.onTrack++;
    }
  }

  // Batch apply background colors to the ID column (column A)
  sheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, colors.length, 1).setBackgrounds(colors);

  // Show summary toast
  var message = 'Traffic lights applied: ' +
                statusCounts.overdue + ' overdue, ' +
                statusCounts.urgent + ' urgent, ' +
                statusCounts.onTrack + ' on track';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, COMMAND_CONFIG.SYSTEM_NAME, 5);

  return statusCounts;
}

/**
 * Removes traffic light indicators from grievance rows
 */
function clearTrafficLightIndicators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Clear background colors from ID column
  sheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, lastRow - 1, 1).setBackground(null);

  SpreadsheetApp.getActiveSpreadsheet().toast('Traffic light indicators cleared', COMMAND_CONFIG.SYSTEM_NAME, 3);
}

/**
 * Applies deadline-based row highlighting
 * Highlights entire rows based on urgency level
 */
function highlightUrgentGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  var daysValues = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1).getValues();

  // Apply row backgrounds based on urgency
  for (var i = 0; i < daysValues.length; i++) {
    var days = daysValues[i][0];
    var rowRange = sheet.getRange(i + 2, 1, 1, lastCol);

    if (typeof days !== 'number' || isNaN(days)) {
      continue;
    } else if (days < 0) {
      rowRange.setBackground(COLORS.ROW_ALT_RED);  // Light red tint
    } else if (days <= 3) {
      rowRange.setBackground('#fefce8');  // Light yellow tint
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Urgent grievances highlighted', COMMAND_CONFIG.SYSTEM_NAME, 3);
}
