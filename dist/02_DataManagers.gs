/**
 * ============================================================================
 * 02_DataManagers.gs - Member & Grievance Data Operations
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Contains all data mutation operations for the Member Directory and
 *   Grievance Log sheets. Includes:
 *   - addMember() / updateMember() — Member Directory CRUD.
 *   - startNewGrievance() / advanceGrievanceStep() / resolveGrievance() —
 *     Grievance Log lifecycle (create, advance workflow, resolve).
 *   - Steward promotion and demotion workflows.
 *   - Member ID generation and validation.
 *   Every write operation in this file uses withScriptLock_() for concurrency
 *   safety and escapeForFormula() on all string inputs before writing to
 *   prevent formula injection attacks.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Centralizes ALL write operations into a single module so that security
 *   controls are applied consistently and cannot be bypassed:
 *   - withScriptLock_() (from 00_DataAccess.gs) serializes concurrent edits,
 *     preventing race conditions when multiple users submit data at the same
 *     time.
 *   - escapeForFormula() (from 00_Security.gs) strips leading =, +, -, @
 *     characters from every input string, blocking formula injection.
 *   - 1-indexed MEMBER_COLS / GRIEVANCE_COLS (from 01_Core.gs) are used for
 *     all column positioning so that column references stay in sync with
 *     sheet structure automatically.
 *   No other module should write directly to Member Directory or Grievance
 *   Log; all writes must flow through this file.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   No members can be added or updated, no grievances can be filed or
 *   modified. The entire data-entry pipeline stops. If escapeForFormula()
 *   calls are lost, formula injection becomes possible. If withScriptLock_()
 *   calls are lost, concurrent edits can corrupt sheet data.
 *
 * DEPENDENCIES:
 *   - Depends on: 00_DataAccess.gs (withScriptLock_), 00_Security.gs
 *     (escapeForFormula), 01_Core.gs (SHEETS, MEMBER_COLS, GRIEVANCE_COLS,
 *     successResponse, errorResponse).
 *   - Used by: 03_UIComponents.gs, 08c_FormsAndNotifications.gs,
 *     10c_FormHandlers.gs, the SPA data service (21_WebDashDataService.gs),
 *     and form submission handlers.
 *
 * @fileoverview Member directory and grievance data operations
 * @version 4.31.0
 * @requires 00_DataAccess.gs
 * @requires 00_Security.gs
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
  return withScriptLock_(function() {
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

    // Build row array and write with single setValues call (F1: batch write)
    var totalCols = sheet.getLastColumn() || Object.keys(MEMBER_COLS).length;
    var rowData = new Array(totalCols).fill('');
    rowData[MEMBER_COLS.MEMBER_ID - 1] = escapeForFormula(memberId);
    rowData[MEMBER_COLS.FIRST_NAME - 1] = escapeForFormula(memberData.firstName || '');
    rowData[MEMBER_COLS.LAST_NAME - 1] = escapeForFormula(memberData.lastName || '');
    rowData[MEMBER_COLS.EMAIL - 1] = escapeForFormula(memberData.email || '');
    rowData[MEMBER_COLS.PHONE - 1] = escapeForFormula(memberData.phone || '');
    rowData[MEMBER_COLS.JOB_TITLE - 1] = escapeForFormula(memberData.jobTitle || '');
    rowData[MEMBER_COLS.WORK_LOCATION - 1] = escapeForFormula(memberData.workLocation || '');
    rowData[MEMBER_COLS.UNIT - 1] = escapeForFormula(memberData.unit || '');
    // REL-03: Wrap sheet write in try/catch with contextual error message
    try {
      sheet.getRange(newRow, 1, 1, totalCols).setValues([rowData]);
    } catch (writeErr) {
      throw new Error('addMember: sheet write failed at row ' + newRow + ': ' + writeErr.message);
    }

    return memberId;
  });
}

/**
 * Updates an existing member's information
 * @param {string} memberId - The member ID to update
 * @param {Object} updateData - Fields to update
 */
function updateMember(memberId, updateData) {
  return withScriptLock_(function() {
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
        memberRow = i + 1; // 1-indexed for getRange
        break;
      }
    }
    if (memberRow === -1) {
      throw new Error('Member not found: ' + memberId);
    }

    // PERF-02: Use the already-loaded data array directly instead of re-reading the row.
    var currentRow = data[memberRow - 1].slice(); // copy to avoid mutating cached array
    var totalCols = currentRow.length;

    // F11-12: Use !== undefined instead of truthiness to allow clearing fields with empty strings
    if (updateData.firstName !== undefined) currentRow[MEMBER_COLS.FIRST_NAME - 1] = escapeForFormula(updateData.firstName);
    if (updateData.lastName !== undefined) currentRow[MEMBER_COLS.LAST_NAME - 1] = escapeForFormula(updateData.lastName);
    if (updateData.email !== undefined) currentRow[MEMBER_COLS.EMAIL - 1] = escapeForFormula(updateData.email);
    if (updateData.phone !== undefined) currentRow[MEMBER_COLS.PHONE - 1] = escapeForFormula(updateData.phone);
    if (updateData.jobTitle !== undefined) currentRow[MEMBER_COLS.JOB_TITLE - 1] = escapeForFormula(updateData.jobTitle);
    if (updateData.workLocation !== undefined) currentRow[MEMBER_COLS.WORK_LOCATION - 1] = escapeForFormula(updateData.workLocation);
    if (updateData.unit !== undefined) currentRow[MEMBER_COLS.UNIT - 1] = escapeForFormula(updateData.unit);

    // REL-03: Wrap sheet write in try/catch with contextual error message
    try {
      sheet.getRange(memberRow, 1, 1, totalCols).setValues([currentRow]);
    } catch (writeErr) {
      throw new Error('updateMember: sheet write failed for ' + memberId + ' at row ' + memberRow + ': ' + writeErr.message);
    }
  });
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
 * @param {number} [maxResults=100] - Maximum number of results to return
 * @returns {Array} Array of matching member objects
 */
var MAX_SEARCH_RESULTS = 200;

function searchMembers(query) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var results = [];
  var queryLower = query.toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (results.length >= MAX_SEARCH_RESULTS) break;

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

  // Fallback with timestamp (last 6 digits) + random component to reduce collision risk
  return namePrefix + String(Date.now()).slice(-6) + String(Math.floor(Math.random() * 100)).padStart(2, '0');
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
    var fullName = (steward['First Name'] || '') + ' ' + (steward['Last Name'] || '');

    var activeCases = 0;
    var totalCases = 0;
    var wonCases = 0;
    var resolvedCases = 0; // H-20: track resolved cases separately for accurate win rate

    for (var i = 1; i < grievances.length; i++) {
      var assignedSteward = grievances[i][GRIEVANCE_COLS.STEWARD - 1];
      if (assignedSteward === fullName) {
        totalCases++;
        var status = grievances[i][GRIEVANCE_COLS.STATUS - 1];
        if (status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING_INFO) {
          activeCases++;
        }
        if (status === GRIEVANCE_STATUS.WON) {
          wonCases++;
        }
        // H-20: denominator matches getDashboardStats — Won+Denied+Settled+Withdrawn only
        // (excludes 'Closed' which lacks a clear outcome, and excludes active cases)
        if (status === GRIEVANCE_STATUS.WON || status === GRIEVANCE_STATUS.DENIED ||
            status === GRIEVANCE_STATUS.SETTLED || status === GRIEVANCE_STATUS.WITHDRAWN) {
          resolvedCases++;
        }
      }
    }

    steward.activeCases = activeCases;
    steward.totalCases = totalCases;
    steward.wonCases = wonCases;
    steward.winRate = resolvedCases > 0 ? Math.round((wonCases / resolvedCases) * 100) : 0;
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
  return withScriptLock_(function() {
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
      if (status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING_INFO) {
        grievanceCounts[memberId].active++;
      }
    }
  }

  // Update member rows (if grievance count columns exist)
  // M-PERF: Batch write — build column arrays and write with 2 setValues instead of 2n setValue
  if (MEMBER_COLS.TOTAL_GRIEVANCES && MEMBER_COLS.ACTIVE_GRIEVANCES) {
    var totalGrievCol = [];
    var activeGrievCol = [];
    for (var j = 1; j < members.length; j++) {
      memberId = members[j][MEMBER_COLS.MEMBER_ID - 1];
      var counts = grievanceCounts[memberId] || { total: 0, active: 0 };
      totalGrievCol.push([counts.total]);
      activeGrievCol.push([counts.active]);
    }
    if (totalGrievCol.length > 0) {
      memberSheet.getRange(2, MEMBER_COLS.TOTAL_GRIEVANCES, totalGrievCol.length, 1).setValues(totalGrievCol);
      memberSheet.getRange(2, MEMBER_COLS.ACTIVE_GRIEVANCES, activeGrievCol.length, 1).setValues(activeGrievCol);
    }
  }

  Logger.log('Member grievance data synced');
  });
}

// ============================================================================
// MEMBER ID GENERATION
// ============================================================================

/**
 * Generates missing Member IDs for all members without one
 * Uses name-based format: M + first 2 letters of first name + first 2 letters of last name + 3 random digits
 * Example: Jane Smith → MJASM472
 */
function generateMissingMemberIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory sheet not found');
    return;
  }

  var data = sheet.getDataRange().getValues();

  // Collect all existing IDs into a set for uniqueness checks
  var existingIds = {};
  for (var j = 1; j < data.length; j++) {
    var id = data[j][MEMBER_COLS.MEMBER_ID - 1];
    if (id) existingIds[id] = true;
  }

  var countAdded = 0;

  // M-PERF: Batch write — modify column array in-memory, write once
  var memberIdCol = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, data.length - 1, 1).getValues();
  for (var i = 1; i < data.length; i++) {
    var currentId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    var firstName = data[i][MEMBER_COLS.FIRST_NAME - 1];
    var lastName = data[i][MEMBER_COLS.LAST_NAME - 1];

    // If ID is blank but member has data
    if (!currentId && (firstName || lastName)) {
      var newId = generateNameBasedId('M', firstName, lastName, existingIds);

      existingIds[newId] = true;
      memberIdCol[i - 1][0] = newId;
      countAdded++;
    }
  }
  if (countAdded > 0) {
    sheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberIdCol.length, 1).setValues(memberIdCol);
  }

  ss.toast('Generated ' + countAdded + ' new Member IDs', COMMAND_CONFIG.SYSTEM_NAME, 5);
  return countAdded;
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
  try {
    withScriptLock_(function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var configSheet = ss.getSheetByName(SHEETS.CONFIG);

      if (!configSheet) return;

      var lastRow = configSheet.getLastRow();
      var rowCount = lastRow >= 3 ? lastRow - 2 : 0;

      if (rowCount > 0) {
        var colData = configSheet.getRange(3, configCol, rowCount, 1).getValues();
        var emptyRow = -1;

        for (var i = 0; i < colData.length; i++) {
          var cellVal = (colData[i][0] || '').toString().trim();
          // Skip if value already exists (prevent duplicates)
          if (cellVal === value.toString().trim()) return;
          if (!cellVal && emptyRow === -1) {
            emptyRow = i + 3;
          }
        }

        if (emptyRow === -1) emptyRow = lastRow + 1;
        configSheet.getRange(emptyRow, configCol).setValue(escapeForFormula(value));
      } else {
        configSheet.getRange(3, configCol).setValue(escapeForFormula(value));
      }
    });
  } catch (lockErr) {
    // Lock timeout is non-critical for config dropdown updates — log and continue
    Logger.log('addToConfigDropdown_: lock unavailable for col ' + configCol + ': ' + lockErr.message);
  }
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

  if (configSheet.getLastRow() < 3) return;

  var colData = configSheet.getRange(3, configCol, configSheet.getLastRow() - 2, 1).getValues();

  for (var i = 0; i < colData.length; i++) {
    if (String(colData[i][0]).trim() === String(value).trim()) {
      configSheet.getRange(i + 3, configCol).clearContent();
      break;
    }
  }
}

// ============================================================================
// STEWARD STATUS SYNC (Bidirectional)
// ============================================================================

/**
 * Bidirectional sync between Member Directory IS_STEWARD column and Config STEWARDS list.
 * - Members marked "Yes" in IS_STEWARD are added to Config STEWARDS if missing.
 * - Names in Config STEWARDS that match a member get IS_STEWARD set to "Yes" if not already.
 * - Members marked "No"/blank in IS_STEWARD are removed from Config STEWARDS.
 * - Names in Config STEWARDS that don't match any member are left untouched (manual entries).
 * Can be called manually from the menu or automatically via triggers.
 */
function syncStewardStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!memberSheet || !configSheet) {
    Logger.log('syncStewardStatus: Required sheets not found');
    return;
  }

  // --- Read Member Directory ---
  var memberData = memberSheet.getDataRange().getValues();
  // Build a map: "First Last" -> { row, isSteward }
  var memberMap = {};        // fullName -> { row: number, isSteward: boolean }
  var stewardsInDirectory = []; // fullNames of members with IS_STEWARD = Yes

  for (var i = 1; i < memberData.length; i++) {
    var firstName = (memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '').toString().trim();
    var lastName = (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '').toString().trim();
    if (!firstName && !lastName) continue; // skip blank rows

    var fullName = firstName + ' ' + lastName;
    var isSteward = isTruthyValue(memberData[i][MEMBER_COLS.IS_STEWARD - 1]);

    memberMap[fullName] = { row: i + 1, isSteward: isSteward };
    if (isSteward) {
      stewardsInDirectory.push(fullName);
    }
  }

  // --- Read Config STEWARDS column ---
  var lastRow = configSheet.getLastRow();
  var configRows = lastRow >= 3 ? lastRow - 2 : 0;
  var configStewardNames = []; // { name: string, row: number }

  if (configRows > 0) {
    var configData = configSheet.getRange(3, CONFIG_COLS.STEWARDS, configRows, 1).getValues();
    for (var c = 0; c < configData.length; c++) {
      var name = (configData[c][0] || '').toString().trim();
      if (name) {
        configStewardNames.push({ name: name, row: c + 3 });
      }
    }
  }

  var configNameSet = {};
  for (var s = 0; s < configStewardNames.length; s++) {
    configNameSet[configStewardNames[s].name] = configStewardNames[s].row;
  }

  var changes = 0;

  // --- Direction 1: Member Directory -> Config ---
  // If IS_STEWARD = Yes but name not in Config, add it
  for (var d = 0; d < stewardsInDirectory.length; d++) {
    var stewardName = stewardsInDirectory[d];
    if (!configNameSet[stewardName]) {
      addToConfigDropdown_(CONFIG_COLS.STEWARDS, stewardName);
      changes++;
    }
  }

  // If IS_STEWARD != Yes but name IS in Config, remove from Config
  for (name in configNameSet) {
    if (memberMap[name] && !memberMap[name].isSteward) {
      configSheet.getRange(configNameSet[name], CONFIG_COLS.STEWARDS).clearContent();
      changes++;
    }
  }

  // --- Direction 2: Config -> Member Directory ---
  // If name is in Config STEWARDS but member IS_STEWARD != Yes, set to Yes
  for (var cn = 0; cn < configStewardNames.length; cn++) {
    var configName = configStewardNames[cn].name;
    if (memberMap[configName] && !memberMap[configName].isSteward) {
      memberSheet.getRange(memberMap[configName].row, MEMBER_COLS.IS_STEWARD).setValue('Yes');
      changes++;
    }
  }

  // --- Compact Config column (remove blank gaps) ---
  if (changes > 0) {
    compactConfigColumn_(CONFIG_COLS.STEWARDS);
  }

  Logger.log('syncStewardStatus: ' + changes + ' change(s) applied');
  ss.toast(changes > 0
    ? '✅ Steward status synced (' + changes + ' change' + (changes > 1 ? 's' : '') + ')'
    : '✅ Steward status already in sync',
    COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Compacts a Config column by removing blank gaps between values.
 * Shifts all non-empty values to the top starting at row 3.
 * @param {number} configCol - The Config column number to compact
 * @private
 */
function compactConfigColumn_(configCol) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return;

  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return;

  var range = configSheet.getRange(3, configCol, lastRow - 2, 1);
  var values = range.getValues();

  // Collect non-empty values
  var compacted = [];
  for (var i = 0; i < values.length; i++) {
    var val = (values[i][0] || '').toString().trim();
    if (val) compacted.push([val]);
  }

  // Pad with empty strings to fill the original range
  while (compacted.length < values.length) {
    compacted.push(['']);
  }

  range.setValues(compacted);
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
        data[i][MEMBER_COLS.EMAIL - 1] = typeof newValuesObj.email === 'string' ? escapeForFormula(newValuesObj.email) : newValuesObj.email;
      }
      if (newValuesObj.phone !== undefined) {
        data[i][MEMBER_COLS.PHONE - 1] = typeof newValuesObj.phone === 'string' ? escapeForFormula(newValuesObj.phone) : newValuesObj.phone;
      }
      if (newValuesObj.firstName !== undefined) {
        data[i][MEMBER_COLS.FIRST_NAME - 1] = typeof newValuesObj.firstName === 'string' ? escapeForFormula(newValuesObj.firstName) : newValuesObj.firstName;
      }
      if (newValuesObj.lastName !== undefined) {
        data[i][MEMBER_COLS.LAST_NAME - 1] = typeof newValuesObj.lastName === 'string' ? escapeForFormula(newValuesObj.lastName) : newValuesObj.lastName;
      }
      if (newValuesObj.unit !== undefined) {
        data[i][MEMBER_COLS.UNIT - 1] = typeof newValuesObj.unit === 'string' ? escapeForFormula(newValuesObj.unit) : newValuesObj.unit;
      }
      if (newValuesObj.workLocation !== undefined) {
        data[i][MEMBER_COLS.WORK_LOCATION - 1] = typeof newValuesObj.workLocation === 'string' ? escapeForFormula(newValuesObj.workLocation) : newValuesObj.workLocation;
      }
      if (newValuesObj.isSteward !== undefined) {
        data[i][MEMBER_COLS.IS_STEWARD - 1] = typeof newValuesObj.isSteward === 'string' ? escapeForFormula(newValuesObj.isSteward) : newValuesObj.isSteward;
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
    '.progress-section { margin-top: 16px; }' +
    '.progress-container { width: 100%; background: #0F172A; border-radius: 8px; border: 1px solid #334155; overflow: hidden; height: 32px; }' +
    '.progress-bar { height: 100%; background: linear-gradient(90deg, #7C3AED, #A78BFA); border-radius: 8px; transition: width 0.3s ease; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 600; color: white; min-width: 0; }' +
    '.progress-text { text-align: center; font-size: 12px; color: #94A3B8; margin-top: 8px; }' +
    '.progress-stats { display: flex; justify-content: space-between; font-size: 11px; color: #64748B; margin-top: 4px; }' +
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
    '<div id="progressSection" class="progress-section" style="display:none;">' +
    '  <div class="progress-container">' +
    '    <div class="progress-bar" id="progressBar" style="width:0%">0%</div>' +
    '  </div>' +
    '  <div class="progress-text" id="progressText">Preparing import...</div>' +
    '  <div class="progress-stats">' +
    '    <span id="progressImported">Imported: 0</span>' +
    '    <span id="progressSkipped">Skipped: 0</span>' +
    '  </div>' +
    '</div>' +
    '' +
    '<script>' +
    getClientSideEscapeHtml() +
    'var parsedData = [];' +
    'var csvHeaders = [];' +
    'var BATCH_SIZE = 25;' +
    'var currentMapping = {};' +
    'var totalImported = 0;' +
    'var totalSkipped = 0;' +
    'var importStartTime = 0;' +
    'var filteredData = [];' +
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
    '  csvHeaders.forEach(function(h) { html += "<th>" + escapeHtml(h) + "</th>"; });' +
    '  html += "</tr>";' +
    '  rows.forEach(function(row) {' +
    '    html += "<tr>";' +
    '    row.forEach(function(cell) { html += "<td>" + escapeHtml(cell || "-") + "</td>"; });' +
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
    '  if (mapping.firstName === undefined || mapping.lastName === undefined) {' +
    '    showStatus("First Name and Last Name are required mappings", true);' +
    '    return;' +
    '  }' +
    '  currentMapping = mapping;' +
    '  totalImported = 0;' +
    '  totalSkipped = 0;' +
    '  importStartTime = Date.now();' +
    '  document.getElementById("importBtn").disabled = true;' +
    '  document.getElementById("importBtn").textContent = "Importing...";' +
    '  document.getElementById("parseBtn").disabled = true;' +
    '  document.getElementById("statusArea").innerHTML = "";' +
    '  document.getElementById("progressSection").style.display = "block";' +
    '  updateProgress(0, parsedData.length, 0, 0);' +
    '  document.getElementById("progressText").textContent = "Checking for duplicates...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(existingKeys) {' +
    '      var uniqueRows = [];' +
    '      var seenEmails = existingKeys.emails || {};' +
    '      var seenNames = existingKeys.names || {};' +
    '      for (var i = 0; i < parsedData.length; i++) {' +
    '        var row = parsedData[i];' +
    '        var fn = currentMapping.firstName !== undefined ? (row[currentMapping.firstName] || "").trim() : "";' +
    '        var ln = currentMapping.lastName !== undefined ? (row[currentMapping.lastName] || "").trim() : "";' +
    '        var em = currentMapping.email !== undefined ? (row[currentMapping.email] || "").trim() : "";' +
    '        if (!fn && !ln) { totalSkipped++; continue; }' +
    '        var emLow = em.toLowerCase();' +
    '        var nmLow = (fn + " " + ln).toLowerCase().trim();' +
    '        if ((emLow && seenEmails[emLow]) || seenNames[nmLow]) { totalSkipped++; continue; }' +
    '        if (emLow) seenEmails[emLow] = true;' +
    '        seenNames[nmLow] = true;' +
    '        uniqueRows.push(row);' +
    '      }' +
    '      filteredData = uniqueRows;' +
    '      if (filteredData.length === 0) {' +
    '        updateProgress(parsedData.length, parsedData.length, 0, totalSkipped);' +
    '        document.getElementById("progressText").textContent = "Complete!";' +
    '        document.getElementById("importBtn").textContent = "Done!";' +
    '        showStatus("No new members to import. " + totalSkipped + " skipped as duplicates.", false);' +
    '        return;' +
    '      }' +
    '      updateProgress(totalSkipped, parsedData.length, 0, totalSkipped);' +
    '      processBatch(0);' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      showStatus("Error: " + e.message, true);' +
    '      showPartialResults();' +
    '    })' +
    '    .getExistingMemberKeys();' +
    '}' +
    '' +
    'function processBatch(startIndex) {' +
    '  var batch = filteredData.slice(startIndex, startIndex + BATCH_SIZE);' +
    '  if (batch.length === 0) { finishImport(); return; }' +
    '  var batchEnd = Math.min(startIndex + batch.length, filteredData.length);' +
    '  document.getElementById("progressText").textContent = "Importing " + (startIndex + 1) + " \\u2013 " + batchEnd + " of " + filteredData.length + " unique members...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.success) {' +
    '        totalImported += result.imported;' +
    '        totalSkipped += result.skipped;' +
    '        var nextStart = startIndex + BATCH_SIZE;' +
    '        updateProgress(totalImported + totalSkipped, parsedData.length, totalImported, totalSkipped);' +
    '        processBatch(nextStart);' +
    '      } else {' +
    '        showStatus("Import failed at row " + (startIndex + 1) + ": " + (result.error || result.message), true);' +
    '        showPartialResults();' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      showStatus("Error at row " + (startIndex + 1) + ": " + e.message, true);' +
    '      showPartialResults();' +
    '    })' +
    '    .importMembersBatch(batch, currentMapping);' +
    '}' +
    '' +
    'function updateProgress(current, total, imported, skipped) {' +
    '  var pct = total > 0 ? Math.round((current / total) * 100) : 0;' +
    '  var bar = document.getElementById("progressBar");' +
    '  bar.style.width = pct + "%";' +
    '  bar.textContent = pct + "%";' +
    '  document.getElementById("progressImported").textContent = "Imported: " + imported;' +
    '  document.getElementById("progressSkipped").textContent = "Skipped: " + skipped;' +
    '}' +
    '' +
    'function finishImport() {' +
    '  updateProgress(parsedData.length, parsedData.length, totalImported, totalSkipped);' +
    '  var elapsed = ((Date.now() - importStartTime) / 1000).toFixed(1);' +
    '  document.getElementById("progressText").textContent = "Complete! Processed " + parsedData.length + " rows in " + elapsed + "s";' +
    '  document.getElementById("importBtn").textContent = "Done!";' +
    '  var msg = "Successfully imported " + totalImported + " members!";' +
    '  if (totalSkipped > 0) msg += " (" + totalSkipped + " skipped as duplicates)";' +
    '  showStatus(msg, false);' +
    '}' +
    '' +
    'function showPartialResults() {' +
    '  document.getElementById("importBtn").disabled = false;' +
    '  document.getElementById("importBtn").textContent = "Retry Import";' +
    '  document.getElementById("parseBtn").disabled = false;' +
    '  if (totalImported > 0) {' +
    '    showStatus("Partially imported " + totalImported + " members before error. " + totalSkipped + " skipped.", true);' +
    '  }' +
    '}' +
    '' +
    'function showStatus(msg, isError) {' +
    '  var area = document.getElementById("statusArea");' +
    '  area.innerHTML = "<div class=\\"status " + (isError ? "error" : "success") + "\\">" + escapeHtml(msg) + "</div>";' +
    '}' +
    '</script>' +
    '</body></html>';
}
/**
 * Returns existing member emails and names for client-side duplicate checking.
 * Called once before import begins to avoid redundant sheet reads per batch.
 * @returns {Object} {emails: {email: true}, names: {name: true}}
 */
function getExistingMemberKeys() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var emails = {};
  var names = {};

  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var email = (data[i][MEMBER_COLS.EMAIL - 1] || '').toString().toLowerCase().trim();
      var name = ((data[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (data[i][MEMBER_COLS.LAST_NAME - 1] || '')).toLowerCase().trim();
      if (email) emails[email] = true;
      if (name) names[name] = true;
    }
  }

  return { emails: emails, names: names };
}

/**
 * Imports a batch of members from parsed CSV data (called by progress bar UI)
 * Data has already been deduplicated client-side against existing members.
 * @param {Array<Array>} batchData - 2D array of CSV data for this batch
 * @param {Object} mapping - Column mapping object {fieldName: columnIndex}
 * @returns {Object} Result with imported/skipped counts for this batch
 */
function importMembersBatch(batchData, mapping) {
  var caller = Session.getEffectiveUser().getEmail();
  if (!caller) {
    Logger.log('Unauthorized batch import attempt');
    return { success: false, message: 'Not authorized' };
  }
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

      if (!sheet) {
        return errorResponse('Member Directory sheet not found', 'importMembersBatch');
      }

      // Read existing Member IDs once for this batch (for ID generation)
      var existingIds = {};
      if (sheet.getLastRow() > 1) {
        var ids = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
        ids.forEach(function(r) { if (r[0]) existingIds[r[0]] = true; });
      }

      var imported = 0;
      var skipped = 0;
      var newRows = [];

      // Dedup already done client-side; just build rows and generate IDs
      for (var j = 0; j < batchData.length; j++) {
        var row = batchData[j];

        var firstName = mapping.firstName !== undefined ? (row[mapping.firstName] || '').trim() : '';
        var lastName = mapping.lastName !== undefined ? (row[mapping.lastName] || '').trim() : '';

        if (!firstName && !lastName) {
          skipped++;
          continue;
        }

        // Generate Member ID inline using pre-read existingIds
        var prefix = 'M';
        var firstPart = (firstName || 'XX').substring(0, 2).toUpperCase();
        var lastPart = (lastName || 'XX').substring(0, 2).toUpperCase();
        var namePrefix = prefix + firstPart + lastPart;
        var memberId = '';
        for (var num = 100; num < 1000; num++) {
          var newId = namePrefix + num;
          if (!existingIds[newId]) {
            memberId = newId;
            existingIds[newId] = true;
            break;
          }
        }
        if (!memberId) memberId = namePrefix + String(Date.now()).slice(-3);

        // Build new row
        var newRow = new Array(MEMBER_HEADER_MAP_.length).fill('');
        newRow[MEMBER_COLS.MEMBER_ID - 1] = escapeForFormula(memberId);
        newRow[MEMBER_COLS.FIRST_NAME - 1] = escapeForFormula(firstName);
        newRow[MEMBER_COLS.LAST_NAME - 1] = escapeForFormula(lastName);

        if (mapping.email !== undefined) newRow[MEMBER_COLS.EMAIL - 1] = escapeForFormula(row[mapping.email] || '');
        if (mapping.phone !== undefined) newRow[MEMBER_COLS.PHONE - 1] = escapeForFormula(row[mapping.phone] || '');
        if (mapping.jobTitle !== undefined) newRow[MEMBER_COLS.JOB_TITLE - 1] = escapeForFormula(row[mapping.jobTitle] || '');
        if (mapping.workLocation !== undefined) newRow[MEMBER_COLS.WORK_LOCATION - 1] = escapeForFormula(row[mapping.workLocation] || '');
        if (mapping.unit !== undefined) newRow[MEMBER_COLS.UNIT - 1] = escapeForFormula(row[mapping.unit] || '');
        if (mapping.supervisor !== undefined) newRow[MEMBER_COLS.SUPERVISOR - 1] = escapeForFormula(row[mapping.supervisor] || '');
        if (mapping.manager !== undefined) newRow[MEMBER_COLS.MANAGER - 1] = escapeForFormula(row[mapping.manager] || '');

        newRow[MEMBER_COLS.IS_STEWARD - 1] = 'No';

        newRows.push(newRow);
        imported++;
      }

      // Batch write new rows
      if (newRows.length > 0) {
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }

      // Log the batch import
      if (imported > 0) {
        logAuditEvent(AUDIT_EVENTS.MEMBER_ADDED, {
          action: 'BATCH_IMPORT',
          importedCount: imported,
          skippedCount: skipped,
          importedBy: caller
        });
      }

      return {
        success: true,
        imported: imported,
        skipped: skipped,
        message: 'Batch import completed'
      };

    } catch (e) {
      Logger.log('Batch import error: ' + e.message);
      return errorResponse(e.message, 'importMembersBatch');
    }
  });
}

/**
 * Shows the export members dialog
 * Allows exporting members to CSV or Google Sheets
 */
function showExportMembersDialog() {
  // Permission check: only editors can export member data (viewers lack menu access anyway)
  var ui = SpreadsheetApp.getUi();
  try {
    var authResult = checkWebAppAuthorization('steward');
    if (!authResult.isAuthorized) {
      ui.alert('Access Denied', 'Steward or admin access is required to export member data.', ui.ButtonSet.OK);
      return;
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
 * @version 4.7.0
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
    return withScriptLock_(function() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

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

      // Prepare row data using GRIEVANCE_COLS constants (1-indexed; subtract 1 for array)
      const totalCols = getGrievanceHeaders().length;
      const rowData = new Array(totalCols).fill('');

      // H-46: Split memberName into first/last if both fields not provided separately
      var firstName = grievanceData.firstName || '';
      var lastName = grievanceData.lastName || '';
      if (!firstName && !lastName && grievanceData.memberName) {
        var nameParts = grievanceData.memberName.trim().split(/\s+/);
        firstName = nameParts[0] || '';
        lastName = nameParts.length > 1 ? nameParts.slice(1).join(' ') : '';
      }

      rowData[GRIEVANCE_COLS.GRIEVANCE_ID - 1]   = grievanceId;
      rowData[GRIEVANCE_COLS.MEMBER_ID - 1]       = escapeForFormula(grievanceData.memberId || '');
      rowData[GRIEVANCE_COLS.FIRST_NAME - 1]      = escapeForFormula(firstName);
      if (GRIEVANCE_COLS.LAST_NAME) {
        rowData[GRIEVANCE_COLS.LAST_NAME - 1]     = escapeForFormula(lastName);
      }
      rowData[GRIEVANCE_COLS.STATUS - 1]          = GRIEVANCE_STATUS.OPEN;
      rowData[GRIEVANCE_COLS.CURRENT_STEP - 1]    = 1;
      rowData[GRIEVANCE_COLS.DATE_FILED - 1]      = filingDate;
      rowData[GRIEVANCE_COLS.STEP1_DUE - 1]       = deadlines.step1Due;
      rowData[GRIEVANCE_COLS.ARTICLES - 1]        = escapeForFormula(grievanceData.articleViolated || '');
      rowData[GRIEVANCE_COLS.ISSUE_CATEGORY - 1]  = escapeForFormula(grievanceData.grievanceType || '');
      rowData[GRIEVANCE_COLS.RESOLUTION - 1]      = escapeForFormula(grievanceData.notes || '');
      rowData[GRIEVANCE_COLS.LAST_UPDATED - 1]    = new Date();

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
    });
  } catch (error) {
    Logger.log('Error creating grievance:', error);
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
      const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (grievanceSheet && GRIEVANCE_COLS.ACTION_TYPE) {
        ensureMinimumColumns(grievanceSheet, getGrievanceHeaders().length);
        // Find the row with this grievance ID
        const lastRow = grievanceSheet.getLastRow();
        const grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, lastRow - 1, 1).getValues();
        for (let i = 0; i < grievanceIds.length; i++) {
          if (grievanceIds[i][0] === result.grievanceId) {
            grievanceSheet.getRange(i + 2, GRIEVANCE_COLS.ACTION_TYPE).setValue(escapeForFormula(grievanceData.actionType));
            break;
          }
        }
      }
    } catch (e) {
      Logger.log('Note: Could not set action type: ' + e.message);
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
          Logger.log('Checklist created with ' + checklistResult.count + ' items');
        }
      } catch (e) {
        Logger.log('Note: Checklist creation skipped: ' + e.message);
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
    const id = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
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
  var rules = getDeadlineRules();
  const step1Due = addBusinessDays(filingDate, rules.STEP_1.DAYS_FOR_RESPONSE);

  return {
    step1Due: step1Due
  };
}
/**
 * Calculates response deadline for a given step
 * @param {number} step - Grievance step
 * @param {Date} stepDate - Date step was initiated
 * @return {Date} Response deadline
 */
function calculateResponseDeadline(step, stepDate) {
  var rules = getDeadlineRules();
  let daysToAdd;

  switch (step) {
    case 1:
      daysToAdd = rules.STEP_1.DAYS_FOR_RESPONSE;
      break;
    case 2:
      daysToAdd = rules.STEP_2.DAYS_FOR_RESPONSE;
      break;
    case 3:
      daysToAdd = rules.STEP_3.DAYS_FOR_RESPONSE;
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
    return withScriptLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    const data = sheet.getDataRange().getValues();

    // Find the grievance row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
        rowIndex = i + 1; // 1-indexed for sheet operations
        break;
      }
    }

    if (rowIndex === -1) {
      return errorResponse('Grievance not found', 'advanceGrievanceStep');
    }

    const currentStep = Number(data[rowIndex - 1][GRIEVANCE_COLS.CURRENT_STEP - 1]);
    if (isNaN(currentStep) || currentStep < 1) {
      return errorResponse('Invalid current step value for this grievance', 'advanceGrievanceStep');
    }
    const nextStep = currentStep + 1;

    if (nextStep > 4) {
      return errorResponse('Grievance is already at arbitration level', 'advanceGrievanceStep');
    }

    const today = new Date();
    const responseDue = calculateResponseDeadline(nextStep, today);

    // Collect column updates to batch write (use GRIEVANCE_COLS, 1-indexed)
    var updates = [];
    updates.push({ col: GRIEVANCE_COLS.CURRENT_STEP, val: nextStep });
    updates.push({ col: GRIEVANCE_COLS.STATUS, val: nextStep === 4 ? GRIEVANCE_STATUS.AT_ARBITRATION : GRIEVANCE_STATUS.APPEALED });
    updates.push({ col: GRIEVANCE_COLS.LAST_UPDATED, val: today });

    if (nextStep <= 3) {
      // F137: Use explicit column map instead of +1/+2 arithmetic.
      // The old code wrote responseDue to DATE_CLOSED and 'Pending' to DAYS_OPEN for Step 3.
      // Verified: column map matches GRIEVANCE_HEADER_MAP_ in 01_Core.gs
      var ADVANCE_STEP_COLS_ = {
        2: { date: GRIEVANCE_COLS.STEP2_APPEAL_FILED, due: GRIEVANCE_COLS.STEP2_DUE },
        3: { date: GRIEVANCE_COLS.STEP3_APPEAL_FILED, due: GRIEVANCE_COLS.STEP3_APPEAL_DUE }
      };
      var stepCols = ADVANCE_STEP_COLS_[nextStep];
      if (stepCols) {
        updates.push({ col: stepCols.date, val: today });
        updates.push({ col: stepCols.due, val: responseDue });
      }
      // Removed erroneous 'Pending' write — STATUS is already set above
    } else {
      updates.push({ col: GRIEVANCE_COLS.DATE_CLOSED, val: today });
    }

    // Add notes if provided — escape once at the final write point only
    if (options.notes) {
      const existingResolution = data[rowIndex - 1][GRIEVANCE_COLS.RESOLUTION - 1] || '';
      const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      const newResolution = existingResolution + (existingResolution ? '\n' : '') +
                       `[${timestamp}] Step ${currentStep} -> ${nextStep}: ${options.notes}`;
      updates.push({ col: GRIEVANCE_COLS.RESOLUTION, val: escapeForFormula(newResolution) });
    }

    // M-PERF: Batch write — read row, apply all updates, write back in single call
    var totalCols = sheet.getLastColumn();
    var rowData = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
    for (var u = 0; u < updates.length; u++) {
      rowData[updates[u].col - 1] = updates[u].val;
    }
    sheet.getRange(rowIndex, 1, 1, totalCols).setValues([rowData]);

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
    });
  } catch (error) {
    Logger.log('Error advancing grievance:', error);
    return errorResponse(error.message, 'advanceGrievanceStep');
  }
}
/**
 * Returns explicit column references for a step's filed/due/rcvd columns.
 * Avoids hardcoded +1/+2 offsets that break when header order changes.
 * @param {number} step - Step number (1-3)
 * @return {Object|null} { filed, due, rcvd } column numbers (1-indexed), or null
 */
function getStepColumnSet(step) {
  switch (step) {
    case 1: return {
      filed: GRIEVANCE_COLS.STEP1_RCVD,
      due:   GRIEVANCE_COLS.STEP1_DUE,
      rcvd:  null  // Step 1 has no separate received column
    };
    case 2: return {
      filed: GRIEVANCE_COLS.STEP2_APPEAL_FILED,
      due:   GRIEVANCE_COLS.STEP2_DUE,
      rcvd:  GRIEVANCE_COLS.STEP2_RCVD
    };
    case 3: return {
      filed: GRIEVANCE_COLS.STEP3_APPEAL_FILED,
      due:   GRIEVANCE_COLS.STEP3_APPEAL_DUE,
      rcvd:  null  // Step 3 has no separate received column
    };
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
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const data = sheet.getDataRange().getValues();

  let updatedCount = 0;
  const startTime = new Date().getTime();

  // Collect updates in memory, then write in a single batch
  var pendingUpdates = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    // Check execution time limit
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
      Logger.log('Approaching time limit, stopping batch');
      break;
    }

    const row = data[i];
    const status = row[GRIEVANCE_COLS.STATUS - 1];

    // Only recalculate open/pending grievances
    if (status === GRIEVANCE_STATUS.OPEN ||
        status === GRIEVANCE_STATUS.PENDING ||
        status === GRIEVANCE_STATUS.APPEALED) {

      const currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];
      const stepCols = getStepColumnSet(currentStep);
      const stepDate = stepCols ? row[stepCols.filed - 1] : null; // 0-indexed for data array

      if (stepDate instanceof Date && stepCols && stepCols.due) {
        const newDue = calculateResponseDeadline(currentStep, stepDate);
        pendingUpdates.push({ row: i + 1, col: stepCols.due, value: newDue });
        updatedCount++;
      }
    }

  }

  // Batch write all updates
  for (let u = 0; u < pendingUpdates.length; u++) {
    var update = pendingUpdates[u];
    data[update.row - 1][update.col - 1] = update.value;
  }
  if (pendingUpdates.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
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
  // Authorization check — only stewards and admins may bulk-update grievances
  var authResult = checkWebAppAuthorization('steward');
  if (!authResult.isAuthorized) {
    return errorResponse(authResult.message || 'Unauthorized: steward access required', 'bulkUpdateGrievanceStatus');
  }

  try {
    return withScriptLock_(function() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      ensureMinimumColumns(sheet, getGrievanceHeaders().length);
      const data = sheet.getDataRange().getValues();
      const numRows = data.length - 1;

      let updatedCount = 0;
      const today = new Date();
      const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

      // M-PERF: Batch write — read columns, modify in-memory, write back
      const statusCol = sheet.getRange(2, GRIEVANCE_COLS.STATUS, numRows, 1).getValues();
      const lastUpdatedCol = sheet.getRange(2, GRIEVANCE_COLS.LAST_UPDATED, numRows, 1).getValues();
      const resolutionCol = notes ? sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, numRows, 1).getValues() : null;

      for (let i = 1; i < data.length; i++) {
        const grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];

        if (grievanceIds.includes(grievanceId)) {
          statusCol[i - 1][0] = newStatus;
          lastUpdatedCol[i - 1][0] = today;

          if (notes) {
            const existingResolution = resolutionCol[i - 1][0] || '';
            const newResolution = existingResolution + (existingResolution ? '\n' : '') +
                             `[${timestamp}] Bulk status update to "${newStatus}": ${notes}`;
            resolutionCol[i - 1][0] = escapeForFormula(newResolution);
          }

          updatedCount++;
        }
      }

      // Write all columns back (2-3 calls total instead of 2-3 per matched row)
      sheet.getRange(2, GRIEVANCE_COLS.STATUS, numRows, 1).setValues(statusCol);
      sheet.getRange(2, GRIEVANCE_COLS.LAST_UPDATED, numRows, 1).setValues(lastUpdatedCol);
      if (notes && resolutionCol) {
        sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, numRows, 1).setValues(resolutionCol);
      }

      return {
        success: true,
        updatedCount: updatedCount,
        message: `Updated status for ${updatedCount} grievances`
      };
    });
  } catch (error) {
    Logger.log('Error in bulk status update:', error);
    return errorResponse(error.message, 'bulkUpdateGrievanceStatus');
  }
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
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
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
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
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
    const status = data[i][GRIEVANCE_COLS.STATUS - 1];
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

  // Use header strings matching GRIEVANCE_HEADER_MAP_ (Roman numerals)
  openGrievances.forEach(g => {
    const currentStep = g['Current Step'];
    let deadline;

    // Get the due date for current step using actual header names
    var stepStr = String(currentStep).trim();
    switch (stepStr) {
      case '1': case 'Step I': case 'Informal':
        deadline = g['Step I Due'];
        break;
      case '2': case 'Step II':
        deadline = g['Step II Due'];
        break;
      case '3': case 'Step III':
        deadline = g['Step III Appeal Due'];
        break;
    }

    if (deadline instanceof Date) {
      const deadlineDate = new Date(deadline);
      deadlineDate.setHours(0, 0, 0, 0);

      if (deadlineDate <= cutoffDate) {
        const daysLeft = Math.ceil((deadlineDate - today) / (1000 * 60 * 60 * 24));
        const memberName = (g['First Name'] || '') + (g['Last Name'] ? ' ' + g['Last Name'] : '');
        upcoming.push({
          grievanceId: g['Grievance ID'],
          memberName: memberName,
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
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

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
    const status = data[i][GRIEVANCE_COLS.STATUS - 1];
    const lastUpdated = data[i][GRIEVANCE_COLS.LAST_UPDATED - 1];
    const category = data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';

    // Count by category
    categoryCounts[category] = (categoryCounts[category] || 0) + 1;

    switch (status) {
      case GRIEVANCE_STATUS.OPEN:
        open++;
        break;
      case GRIEVANCE_STATUS.PENDING:
      case GRIEVANCE_STATUS.APPEALED:
      case GRIEVANCE_STATUS.IN_ARBITRATION:
        pending++;
        break;
      case GRIEVANCE_STATUS.WON:
        won++;
        resolved++;
        if (lastUpdated instanceof Date && lastUpdated.getFullYear() === currentYear) {
          closedThisYear++;
        }
        break;
      case GRIEVANCE_STATUS.DENIED:
      case GRIEVANCE_STATUS.WITHDRAWN:
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
    return withScriptLock_(function() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      ensureMinimumColumns(sheet, getGrievanceHeaders().length);
      const data = sheet.getDataRange().getValues();

      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex === -1) {
        return errorResponse('Grievance not found');
      }

      const today = new Date();
      const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

      // Build combined resolution text
      var resolutionText = outcome || '';
      if (resolution) {
        resolutionText += (resolutionText ? ': ' : '') + resolution;
      }
      if (notes) {
        resolutionText += (resolutionText ? '\n' : '') +
                          '[' + timestamp + '] ' + notes;
      }
      // H-29: Use outcome parameter to determine status instead of always 'Settled'
      var validOutcomes = {
        'Won': GRIEVANCE_STATUS.WON,
        'Denied': GRIEVANCE_STATUS.DENIED,
        'Settled': GRIEVANCE_STATUS.SETTLED,
        'Withdrawn': GRIEVANCE_STATUS.WITHDRAWN,
        'Closed': GRIEVANCE_STATUS.CLOSED
      };
      var resolvedStatus = (outcome && validOutcomes[outcome]) ? validOutcomes[outcome] : GRIEVANCE_STATUS.SETTLED;

      // M-PERF: Batch write — read row, apply 4 updates, write back in single call
      var totalCols = sheet.getLastColumn();
      var rowData = sheet.getRange(rowIndex, 1, 1, totalCols).getValues()[0];
      rowData[GRIEVANCE_COLS.RESOLUTION - 1] = escapeForFormula(resolutionText);
      rowData[GRIEVANCE_COLS.STATUS - 1] = resolvedStatus;
      rowData[GRIEVANCE_COLS.DATE_CLOSED - 1] = today;
      rowData[GRIEVANCE_COLS.LAST_UPDATED - 1] = today;
      sheet.getRange(rowIndex, 1, 1, totalCols).setValues([rowData]);

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
    });
  } catch (error) {
    Logger.log('Error resolving grievance:', error);
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
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    showAlert('Please select a grievance in the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    showAlert('Please select a grievance row', 'No Selection');
    return;
  }

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const grievanceId = data[GRIEVANCE_COLS.GRIEVANCE_ID - 1];

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
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    showAlert('Please use this from the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const openGrievances = getOpenGrievances();
  const items = openGrievances.map(g => {
    var memberName = ((g['First Name'] || '') + ' ' + (g['Last Name'] || '')).trim() || 'Unknown';
    return {
      id: g['Grievance ID'],
      label: `${g['Grievance ID']} - ${memberName}`,
      selected: false
    };
  });

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
    `<option value="${escapeHtml(String(m.id))}">${escapeHtml(String(m.name))} (${escapeHtml(String(m.department))})</option>`
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
          <strong>Grievance ID:</strong> ${escapeHtml(String(grievanceId))}<br>
          <strong>Current Step:</strong> ${escapeHtml(String(grievance['Current Step'] || 1))}<br>
          <strong>Status:</strong> <span class="status-badge status-open">
            ${escapeHtml(String(grievance['Status'] || 'Open'))}</span>
        </div>

        <form id="editForm">
          <div class="form-group">
            <label class="form-label">Issue Category</label>
            <textarea class="form-textarea" id="description">${escapeHtml(String(grievance['Issue Category'] || ''))}</textarea>
          </div>

          <div class="form-group">
            <label class="form-label">Resolution / Notes</label>
            <textarea class="form-textarea" id="notes">${escapeHtml(String(grievance['Resolution'] || ''))}</textarea>
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
        const grievanceId = ${JSON.stringify(grievanceId)};

        document.getElementById('editForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const updates = {
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value
          };
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(err) { alert('Error saving: ' + err.message); })
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
              .withFailureHandler(function(err) { alert('Error advancing step: ' + err.message); })
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
// ============================================================================
// BULK ACTIONS — Grievance Log (PHASE2 Feature 4)
// ============================================================================

/**
 * Reads the QUICK_ACTIONS checkbox column from the Grievance Log and returns
 * an array of 1-indexed row numbers where the checkbox is TRUE.
 * @returns {number[]} Row numbers of selected grievances
 */
function getSelectedGrievanceRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var checkboxData = sheet.getRange(2, GRIEVANCE_COLS.QUICK_ACTIONS, lastRow - 1, 1).getValues();
  var selected = [];
  for (var i = 0; i < checkboxData.length; i++) {
    if (checkboxData[i][0] === true) {
      selected.push(i + 2); // Convert to 1-indexed row number
    }
  }
  return selected;
}

/**
 * Sets MESSAGE_ALERT to TRUE for the given row numbers in the Grievance Log.
 * Uses batch write for performance.
 * @param {number[]} rowNumbers - 1-indexed row numbers to flag
 * @returns {{success: boolean, count: number, error: string}}
 */
function bulkFlagGrievances(rowNumbers) {
  var caller = Session.getEffectiveUser().getEmail();
  if (!caller) {
    Logger.log('Unauthorized bulk flag attempt');
    return { success: false, message: 'Not authorized' };
  }
  if (!rowNumbers || !rowNumbers.length) {
    return errorResponse('No rows selected');
  }
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet) return errorResponse('Grievance Log sheet not found');

      // M-PERF: Batch write — read column, set matching rows, write back
      var lastRow = sheet.getLastRow();
      var alertCol = lastRow > 1 ? sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).getValues() : [];
      var count = 0;
      for (var i = 0; i < rowNumbers.length; i++) {
        var row = rowNumbers[i];
        if (typeof row !== 'number' || row < 2 || row > lastRow) continue;
        alertCol[row - 2][0] = true;
        count++;
      }
      if (count > 0) {
        sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, alertCol.length, 1).setValues(alertCol);
      }

      logAuditEvent('BULK_FLAG_GRIEVANCES', 'Flagged ' + count + ' grievances for message alert. Rows: ' + rowNumbers.join(', '));

      return successResponse({ count: count }, 'Flagged ' + count + ' grievance(s) for message alert');
    } catch (e) {
      handleError(e, 'bulkFlagGrievances');
      return errorResponse(e.message, 'bulkFlagGrievances');
    }
  });
}

/**
 * Sends a bulk email to members associated with the selected grievance rows.
 * @param {number[]} rowNumbers - 1-indexed row numbers
 * @param {string} subject - Email subject line
 * @param {string} body - Email body (plain text)
 * @returns {{success: boolean, data: {sent: number, failed: number, skipped: number}, error: string}}
 */
function bulkEmailGrievanceMembers(rowNumbers, subject, body) {
  var caller = Session.getEffectiveUser().getEmail();
  if (!caller) {
    Logger.log('Unauthorized bulk email attempt');
    return { success: false, message: 'Not authorized' };
  }
  if (!rowNumbers || !rowNumbers.length) {
    return errorResponse('No rows selected');
  }
  if (!subject || !body) {
    return errorResponse('Subject and body are required');
  }
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet) return errorResponse('Grievance Log sheet not found');

      var sent = 0;
      var failed = 0;
      var skipped = 0;
      var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

      for (var i = 0; i < rowNumbers.length; i++) {
        var row = rowNumbers[i];
        if (typeof row !== 'number' || row < 2) { skipped++; continue; }

        var email = sheet.getRange(row, GRIEVANCE_COLS.MEMBER_EMAIL).getValue();
        if (!email || !emailRegex.test(String(email).trim())) {
          skipped++;
          continue;
        }

        try {
          var emailSent = safeSendEmail({
            to: String(email).trim(),
            subject: subject,
            body: body
          });
          if (emailSent) { sent++; } else { failed++; }
        } catch (emailErr) {
          Logger.log('bulkEmailGrievanceMembers: failed for row ' + row + ': ' + emailErr.message);
          failed++;
        }
      }

      logAuditEvent('BULK_EMAIL_GRIEVANCE_MEMBERS', 'Sent ' + sent + ', failed ' + failed + ', skipped ' + skipped + ' of ' + rowNumbers.length + ' selected rows');

      return successResponse(
        { sent: sent, failed: failed, skipped: skipped },
        'Sent ' + sent + ' email(s), ' + failed + ' failed, ' + skipped + ' skipped (no valid email)'
      );
    } catch (e) {
      handleError(e, 'bulkEmailGrievanceMembers');
      return errorResponse(e.message, 'bulkEmailGrievanceMembers');
    }
  });
}

/**
 * Exports selected grievance rows as CSV and emails it to the current user.
 * @param {number[]} rowNumbers - 1-indexed row numbers to export
 * @returns {{success: boolean, message: string, error: string}}
 */
function bulkExportGrievancesToCsv(rowNumbers) {
  var caller = Session.getEffectiveUser().getEmail();
  if (!caller) {
    Logger.log('Unauthorized bulk export attempt');
    return { success: false, message: 'Not authorized' };
  }
  if (!rowNumbers || !rowNumbers.length) {
    return errorResponse('No rows selected');
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return errorResponse('Grievance Log sheet not found');

    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return errorResponse('Could not determine current user email');

    // Get header row
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Build CSV rows
    var csvRows = [];
    // Header row — escape for CSV
    csvRows.push(headers.map(function(h) {
      return '"' + String(h).replace(/"/g, '""') + '"';
    }).join(','));

    // Data rows
    for (var i = 0; i < rowNumbers.length; i++) {
      var row = rowNumbers[i];
      if (typeof row !== 'number' || row < 2) continue;
      var rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
      csvRows.push(rowData.map(function(cell) {
        var val = (cell instanceof Date) ? cell.toISOString() : String(cell);
        return '"' + val.replace(/"/g, '""') + '"';
      }).join(','));
    }

    var csvContent = csvRows.join('\r\n');
    var blob = Utilities.newBlob(csvContent, 'text/csv', 'grievance_export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv');

    safeSendEmail({
      to: userEmail,
      subject: 'Grievance Export — ' + rowNumbers.length + ' records',
      body: 'Attached is a CSV export of ' + rowNumbers.length + ' selected grievance record(s).\n\nExported on: ' + new Date().toLocaleString(),
      attachments: [blob]
    });

    logAuditEvent('BULK_EXPORT_GRIEVANCES_CSV', 'Exported ' + rowNumbers.length + ' grievances as CSV to ' + userEmail);

    return successResponse(null, 'CSV with ' + rowNumbers.length + ' record(s) sent to ' + userEmail);
  } catch (e) {
    handleError(e, 'bulkExportGrievancesToCsv');
    return errorResponse(e.message, 'bulkExportGrievancesToCsv');
  }
}

/**
 * Selects (checks) QUICK_ACTIONS for all open/pending grievance rows.
 * @returns {{success: boolean, count: number}}
 */
function selectAllOpenCases() {
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet) return errorResponse('Grievance Log sheet not found');

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return successResponse({ count: 0 }, 'No data rows found');

      var statusData = sheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1).getValues();
      var checkboxValues = [];
      var count = 0;

      var openLower = GRIEVANCE_STATUS.OPEN.toLowerCase();
      var pendingLower = GRIEVANCE_STATUS.PENDING_INFO.toLowerCase();
      for (var i = 0; i < statusData.length; i++) {
        var status = String(statusData[i][0]).trim().toLowerCase();
        var isOpen = (status === openLower || status === pendingLower);
        if (isOpen) count++;
        checkboxValues.push([isOpen]);
      }

      sheet.getRange(2, GRIEVANCE_COLS.QUICK_ACTIONS, lastRow - 1, 1).setValues(checkboxValues);

      SpreadsheetApp.getActiveSpreadsheet().toast('Selected ' + count + ' open/pending cases', COMMAND_CONFIG.SYSTEM_NAME, 3);
      return successResponse({ count: count }, 'Selected ' + count + ' open/pending case(s)');
    } catch (e) {
      handleError(e, 'selectAllOpenCases');
      return errorResponse(e.message, 'selectAllOpenCases');
    }
  });
}

/**
 * Clears (unchecks) all QUICK_ACTIONS checkboxes in the Grievance Log.
 * @returns {{success: boolean, message: string}}
 */
function clearAllSelections() {
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet) return errorResponse('Grievance Log sheet not found');

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return successResponse(null, 'No data rows to clear');

      var clearValues = [];
      for (var i = 0; i < lastRow - 1; i++) {
        clearValues.push([false]);
      }

      sheet.getRange(2, GRIEVANCE_COLS.QUICK_ACTIONS, lastRow - 1, 1).setValues(clearValues);

      SpreadsheetApp.getActiveSpreadsheet().toast('Selection cleared', COMMAND_CONFIG.SYSTEM_NAME, 3);
      return successResponse(null, 'All selections cleared');
    } catch (e) {
      handleError(e, 'clearAllSelections');
      return errorResponse(e.message, 'clearAllSelections');
    }
  });
}
