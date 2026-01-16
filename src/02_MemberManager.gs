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

  if (currentStatus === 'Yes' || currentStatus === true) {
    ui.alert(fullName + ' is already a Steward');
    return;
  }

  var response = ui.alert(
    'Promote to Steward',
    'Promote ' + fullName + ' to Steward?\n\nThis will:\n' +
    '- Set "Is Steward" to Yes\n' +
    '- Add to the Steward dropdown list',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Update member record
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('Yes');

  // Add to Config steward list
  addToConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);

  ss.toast(fullName + ' has been promoted to Steward', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Demotes the currently selected steward back to regular member
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

  if (currentStatus !== 'Yes' && currentStatus !== true) {
    ui.alert(fullName + ' is not currently a Steward');
    return;
  }

  var response = ui.alert(
    'Demote Steward',
    'Demote ' + fullName + ' from Steward?\n\nThis will:\n' +
    '- Set "Is Steward" to No\n' +
    '- Remove from the Steward dropdown list',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Update member record
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('No');

  // Remove from Config steward list
  removeFromConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);

  ss.toast(fullName + ' has been demoted from Steward', COMMAND_CONFIG.SYSTEM_NAME, 5);
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
