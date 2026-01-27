/**
 * ============================================================================
 * FEATURE EXTENSION: DYNAMIC ENGINE & MEMBER LEADERS
 * ============================================================================
 * This script handles dynamic column detection and organizational leadership.
 *
 * OPTIMIZATIONS:
 * - CacheService for header maps (5-minute TTL)
 * - Single sheet read for multiple operations
 * - Batch writes using setValues()
 * - Pre-computed column indices
 * - Minimized API calls via shared spreadsheet reference
 *
 * FEATURES:
 * ---------
 * 1. Member Leaders (Organizational Layer)
 *    - Creates a new hierarchy level for organizational "nodes"
 *    - Member Leaders connect members without case-management duties
 *    - Filtered out of Grievance dropdowns via IS_STEWARD column
 *    - Uses "Member Leader" value in IS_STEWARD column (column N)
 *
 * 2. Column Expansion (Dynamic Engine)
 *    - "No-Code" feature for adding custom columns
 *    - Header Scanner performs census of top row on dashboard open
 *    - Auto-generates form fields with matching CSS styling
 *    - Detection Zone: columns beyond Core data (Column 32+)
 *
 * 3. Self-Healing Hidden Architecture
 *    - Script-to-Sheet Bridge for formula management
 *    - Uses dedicated section in _Dashboard_Calc (row 50+)
 *    - Repair Trigger re-injects formulas if broken
 *    - Dynamic Referencing uses Col1, Col2 logic (not cell refs)
 *
 * INTEGRATION NOTES:
 * - Uses MEMBER_COLS from 01_Constants.gs for column positions
 * - Uses SHEETS constant for sheet names
 * - Integrates with getAllStewards() via getStewardsForGrievance()
 */

const EXTENSION_CONFIG = {
  HIDDEN_CALC_SHEET: "_Dashboard_Calc",
  DYNAMIC_FORMULA_ROW: 50,
  MEMBER_SHEET: "Member Directory",
  GRIEVANCE_SHEET: "Grievance Log",
  LEADER_ROLE_NAME: "Member Leader",
  CORE_COLUMN_COUNT: 32,
  CACHE_TTL_SECONDS: 300 // 5-minute cache for header maps
};

// Pre-computed column indices (0-based for array access)
const COL_IDX = {
  MEMBER_ID: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.MEMBER_ID : 1) - 1,
  FIRST_NAME: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.FIRST_NAME : 2) - 1,
  LAST_NAME: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.LAST_NAME : 3) - 1,
  EMAIL: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.EMAIL : 8) - 1,
  UNIT: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.UNIT : 5) - 1,
  WORK_LOCATION: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.WORK_LOCATION : 6) - 1,
  IS_STEWARD: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.IS_STEWARD : 14) - 1
};

// ============================================================================
// CACHING LAYER
// ============================================================================

/**
 * Gets cached header map or builds fresh one.
 * Uses CacheService for 5-minute TTL to reduce sheet reads.
 * @param {string} sheetName - Sheet to scan
 * @param {boolean} forceRefresh - Force cache refresh
 * @returns {Object} Header name -> column index map
 */
function getHeaderMap(sheetName, forceRefresh) {
  const cacheKey = 'headerMap_' + sheetName.replace(/\s/g, '_');
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // Cache corrupted, rebuild
      }
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {};

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = {};

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header && header.toString().trim() !== "") {
      headerMap[header.toString().trim()] = i + 1;
    }
  }

  // Cache for 5 minutes
  try {
    cache.put(cacheKey, JSON.stringify(headerMap), EXTENSION_CONFIG.CACHE_TTL_SECONDS);
  } catch (e) {
    // Cache write failed, continue without caching
  }

  return headerMap;
}

/**
 * Invalidates header map cache for a sheet.
 * Call after structural changes (adding/removing columns).
 * @param {string} sheetName - Sheet to invalidate
 */
function invalidateHeaderCache(sheetName) {
  const cacheKey = 'headerMap_' + sheetName.replace(/\s/g, '_');
  CacheService.getScriptCache().remove(cacheKey);
}

// ============================================================================
// UNIFIED DATA LOADER
// ============================================================================

/**
 * Loads member directory data once and returns parsed results.
 * Optimized single-read function for multiple queries.
 * @param {Object} options - Query options
 * @returns {Object} Parsed member data with leaders, stewards, and full dataset
 */
function loadMemberData_(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { leaders: [], stewards: [], data: [], headers: [], found: null };
  }

  const range = sheet.getDataRange();
  const data = range.getValues();

  if (data.length < 2) {
    return { leaders: [], stewards: [], data: data, headers: data[0] || [], found: null };
  }

  const headers = data[0];
  const leaders = [];
  const stewards = [];
  let foundMember = null;

  // Single pass through data
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const roleValue = row[COL_IDX.IS_STEWARD];
    const roleStr = roleValue ? roleValue.toString().trim() : '';
    const firstName = row[COL_IDX.FIRST_NAME] || '';
    const lastName = row[COL_IDX.LAST_NAME] || '';
    const fullName = (firstName + ' ' + lastName).trim();
    const memberId = row[COL_IDX.MEMBER_ID] || '';

    // Check for specific member lookup
    if (options.findMemberId && memberId === options.findMemberId) {
      foundMember = { row: i + 1, data: row, roleValue: roleStr };
    }

    // Categorize by role
    if (roleStr === EXTENSION_CONFIG.LEADER_ROLE_NAME) {
      if (fullName) {
        leaders.push({
          name: fullName,
          firstName: firstName,
          lastName: lastName,
          memberId: memberId,
          email: row[COL_IDX.EMAIL] || '',
          unit: row[COL_IDX.UNIT] || '',
          location: row[COL_IDX.WORK_LOCATION] || '',
          rowNumber: i + 1
        });
      }
    } else if (roleStr === 'Yes' || roleValue === true) {
      const steward = { rowNumber: i + 1, fullName: fullName };
      for (let j = 0; j < headers.length; j++) {
        steward[headers[j]] = row[j];
      }
      stewards.push(steward);
    }
  }

  return {
    leaders: leaders,
    stewards: stewards,
    data: data,
    headers: headers,
    found: foundMember
  };
}

// ============================================================================
// PUBLIC API - MEMBER LEADERS
// ============================================================================

/**
 * Fetches all members tagged as 'Member Leader'.
 * These individuals are organizational and do not appear in Grievance lists.
 * @returns {Array} Array of leader objects
 */
function getMemberLeaders() {
  return loadMemberData_().leaders;
}

/**
 * Gets stewards for grievance dropdowns, EXCLUDING Member Leaders.
 * Use this instead of getAllStewards() when populating grievance assignment dropdowns.
 * @returns {Array} Array of steward objects (excludes Member Leaders)
 */
function getStewardsForGrievance() {
  return loadMemberData_().stewards;
}

/**
 * Gets both leaders and stewards in a single optimized call.
 * Use when you need both datasets to avoid duplicate sheet reads.
 * @returns {Object} { leaders: Array, stewards: Array }
 */
function getLeadersAndStewards() {
  const result = loadMemberData_();
  return { leaders: result.leaders, stewards: result.stewards };
}

/**
 * Checks if a member is a Member Leader (not a Steward).
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if member is a Member Leader
 */
function isMemberLeader(memberId) {
  if (!memberId) return false;
  const result = loadMemberData_({ findMemberId: memberId });
  return result.found && result.found.roleValue === EXTENSION_CONFIG.LEADER_ROLE_NAME;
}

// ============================================================================
// SELF-HEALING DYNAMIC FORMULAS
// ============================================================================

/**
 * Re-injects logic into a dedicated section of the hidden sheet.
 * Uses batch writes for performance.
 * @returns {Object} Result with success status
 */
function repairDynamicFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);

  if (!calcSheet) {
    calcSheet = ss.insertSheet(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
    calcSheet.hideSheet();
  }

  const startRow = EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW;
  const isStewardCol = (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.IS_STEWARD : 14);
  const colLetter = typeof getColumnLetter === 'function' ? getColumnLetter(isStewardCol) : 'N';

  // Build all values and formulas for batch write
  const batchData = [
    ['=== DYNAMIC ENGINE FORMULAS ===', ''],
    ['Role Distribution:', ''],
    [`=QUERY('${EXTENSION_CONFIG.MEMBER_SHEET}'!A:ZZ, "SELECT Col${isStewardCol}, count(Col1) WHERE Col1 IS NOT NULL GROUP BY Col${isStewardCol} LABEL count(Col1) 'Total'", 1)`, ''],
    ['', ''],
    ['', ''],
    ['', ''],
    ['Member Leaders:', `=COUNTIF('${EXTENSION_CONFIG.MEMBER_SHEET}'!${colLetter}:${colLetter},"${EXTENSION_CONFIG.LEADER_ROLE_NAME}")`],
    ['Active Stewards:', `=COUNTIF('${EXTENSION_CONFIG.MEMBER_SHEET}'!${colLetter}:${colLetter},"Yes")`]
  ];

  // Single batch write
  calcSheet.getRange(startRow, 1, batchData.length, 2).setValues(batchData);

  ss.toast('Dynamic formulas injected at row ' + startRow, 'Self-Healing Complete', 5);

  return { success: true, startRow: startRow };
}

// ============================================================================
// COLUMN EXPANSION ENGINE
// ============================================================================

/**
 * Gets expansion column data with optional member data.
 * Optimized to minimize sheet reads.
 * @param {string} memberId - Optional member ID to fetch data for
 * @returns {Object} Object with extraHeaders and memberData
 */
function getExpansionColumnData(memberId) {
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
  const headerMap = getHeaderMap(sheetName);
  const coreCount = EXTENSION_CONFIG.CORE_COLUMN_COUNT;

  // Identify extra columns (beyond core)
  const extraHeaders = [];
  const headerEntries = Object.entries(headerMap);

  for (let i = 0; i < headerEntries.length; i++) {
    const [name, col] = headerEntries[i];
    if (col > coreCount) {
      extraHeaders.push({
        name: name,
        column: col,
        letter: typeof getColumnLetter === 'function' ? getColumnLetter(col) : ''
      });
    }
  }

  // Sort by column position
  extraHeaders.sort((a, b) => a.column - b.column);

  // Fetch member data if requested and there are extra columns
  const memberData = {};
  if (memberId && extraHeaders.length > 0) {
    const result = loadMemberData_({ findMemberId: memberId });
    if (result.found) {
      for (let i = 0; i < extraHeaders.length; i++) {
        const extra = extraHeaders[i];
        memberData[extra.name] = result.found.data[extra.column - 1] || '';
      }
    }
  }

  return {
    extraHeaders: extraHeaders,
    memberData: memberData,
    coreColumnCount: coreCount
  };
}

/**
 * Generates HTML form fields for custom columns.
 * Uses array join for efficient string building.
 * @param {string} memberId - Optional member ID to pre-fill values
 * @returns {string} HTML string with form fields
 */
function generateExpansionFieldsHtml(memberId) {
  const expansionData = getExpansionColumnData(memberId);

  if (expansionData.extraHeaders.length === 0) {
    return '<!-- No custom columns detected -->';
  }

  const parts = [
    '<div class="expansion-fields" style="margin-top:16px;padding-top:16px;border-top:1px solid #E5E7EB;">',
    '<div class="expansion-header" style="font-weight:600;color:#7C3AED;margin-bottom:12px;">Custom Fields</div>'
  ];

  const fields = expansionData.extraHeaders;
  for (let i = 0; i < fields.length; i++) {
    const field = fields[i];
    const value = (expansionData.memberData[field.name] || '').toString().replace(/"/g, '&quot;');
    const fieldId = 'custom_' + field.name.replace(/[^a-zA-Z0-9]/g, '_');

    parts.push(
      '<div class="form-group" style="margin-bottom:12px;">',
      '<label class="form-label" style="display:block;font-size:13px;color:#374151;margin-bottom:4px;">',
      field.name,
      '</label>',
      '<input type="text" class="form-input" id="', fieldId, '" name="', fieldId,
      '" data-column="', field.column, '" value="', value,
      '" style="width:100%;padding:8px 12px;border:1px solid #D1D5DB;border-radius:6px;font-family:Roboto,sans-serif;">',
      '</div>'
    );
  }

  parts.push('</div>');
  return parts.join('');
}

/**
 * Saves custom column data for a member using batch write.
 * @param {string} memberId - The member ID
 * @param {Object} customData - Object with field names as keys
 * @returns {Object} Result with success status
 */
function saveExpansionData(memberId, customData) {
  if (!memberId || !customData) {
    return { success: false, error: 'Member ID and data required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, error: 'Sheet not found' };
  }

  const headerMap = getHeaderMap(sheetName);
  const result = loadMemberData_({ findMemberId: memberId });

  if (!result.found) {
    return { success: false, error: 'Member not found' };
  }

  const memberRow = result.found.row;
  const coreCount = EXTENSION_CONFIG.CORE_COLUMN_COUNT;

  // Collect updates for batch write
  const updates = [];
  const fieldNames = Object.keys(customData);

  for (let i = 0; i < fieldNames.length; i++) {
    const fieldName = fieldNames[i];
    const col = headerMap[fieldName];
    if (col && col > coreCount) {
      updates.push({ col: col, value: customData[fieldName] });
    }
  }

  if (updates.length === 0) {
    return { success: true, updated: 0 };
  }

  // Sort by column for potential range optimization
  updates.sort((a, b) => a.col - b.col);

  // Check if columns are contiguous for single batch write
  const minCol = updates[0].col;
  const maxCol = updates[updates.length - 1].col;

  if (maxCol - minCol + 1 === updates.length) {
    // Contiguous columns - single batch write
    const values = updates.map(u => u.value);
    sheet.getRange(memberRow, minCol, 1, values.length).setValues([values]);
  } else {
    // Non-contiguous - use individual writes (still fewer than per-field)
    for (let i = 0; i < updates.length; i++) {
      sheet.getRange(memberRow, updates[i].col).setValue(updates[i].value);
    }
  }

  return { success: true, updated: updates.length };
}

// ============================================================================
// SETUP & CONFIGURATION
// ============================================================================

/**
 * Adds "Member Leader" option to the IS_STEWARD validation dropdown.
 * @returns {Object} Result with success status
 */
function setupMemberLeaderRole() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheetName = (typeof SHEETS !== 'undefined' && SHEETS.CONFIG) || 'Config';
  const configSheet = ss.getSheetByName(configSheetName);

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found');
    return { success: false, error: 'Config sheet not found' };
  }

  const yesNoCol = (typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.YES_NO) || 5;
  const existingValues = configSheet.getRange(3, yesNoCol, 10, 1).getValues().flat().filter(Boolean);

  if (existingValues.indexOf(EXTENSION_CONFIG.LEADER_ROLE_NAME) !== -1) {
    ss.toast('Member Leader role already configured', 'Setup Complete', 3);
    return { success: true, alreadyExists: true };
  }

  const nextRow = existingValues.length + 3;
  configSheet.getRange(nextRow, yesNoCol).setValue(EXTENSION_CONFIG.LEADER_ROLE_NAME);

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
      action: 'MEMBER_LEADER_SETUP',
      addedValue: EXTENSION_CONFIG.LEADER_ROLE_NAME,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  ss.toast('Member Leader role added to Config.', 'Setup Complete', 5);
  return { success: true, addedAt: nextRow };
}

/**
 * Runs all Dynamic Engine setup tasks.
 * @returns {Object} Result with setup details
 */
function setupDynamicEngine() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Dynamic Engine',
    'This will:\n\n' +
    '1. Add "Member Leader" to IS_STEWARD dropdown options\n' +
    '2. Inject self-healing formulas into _Dashboard_Calc\n' +
    '3. Scan for custom columns beyond Column AF\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  // Invalidate cache before setup
  invalidateHeaderCache(EXTENSION_CONFIG.MEMBER_SHEET);

  const roleResult = setupMemberLeaderRole();
  const formulaResult = repairDynamicFormulas();
  const expansionData = getExpansionColumnData();

  const summary =
    'Dynamic Engine Setup Complete:\n\n' +
    '• Member Leader Role: ' + (roleResult.success ? 'OK' : 'Failed') + '\n' +
    '• Self-Healing Formulas: ' + (formulaResult.success ? 'Injected at row ' + formulaResult.startRow : 'Failed') + '\n' +
    '• Custom Columns Found: ' + expansionData.extraHeaders.length;

  ui.alert('Setup Complete', summary, ui.ButtonSet.OK);

  return {
    success: true,
    roleSetup: roleResult,
    formulas: formulaResult,
    expansionColumns: expansionData.extraHeaders.length
  };
}

/**
 * Gets Dynamic Engine status for display in admin panels.
 * Optimized single-pass data collection.
 * @returns {Object} Status object with feature states
 */
function getDynamicEngineStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Single load for leaders
  const memberData = loadMemberData_();

  // Get expansion columns (uses cached header map)
  const expansionData = getExpansionColumnData();

  // Check self-healing status
  const calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
  let formulasInjected = false;

  if (calcSheet) {
    const checkValue = calcSheet.getRange(EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW, 1).getValue();
    formulasInjected = checkValue && checkValue.toString().indexOf('DYNAMIC ENGINE') !== -1;
  }

  return {
    memberLeaders: {
      count: memberData.leaders.length,
      names: memberData.leaders.map(function(l) { return l.name; })
    },
    stewards: {
      count: memberData.stewards.length
    },
    expansionColumns: {
      count: expansionData.extraHeaders.length,
      headers: expansionData.extraHeaders.map(function(h) { return h.name; })
    },
    selfHealing: {
      active: formulasInjected,
      sheet: EXTENSION_CONFIG.HIDDEN_CALC_SHEET,
      row: EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW
    }
  };
}

// ============================================================================
// GRIEVANCE REMINDERS FEATURE
// ============================================================================
// Allows users to set two reminder dates with notes for meetings/follow-ups

/**
 * Sets a reminder for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @param {number} reminderNum - Which reminder (1 or 2)
 * @param {Date|string} reminderDate - The reminder date
 * @param {string} reminderNote - Brief note about the reminder
 * @returns {Object} Result with success status
 */
function setGrievanceReminder(grievanceId, reminderNum, reminderDate, reminderNote) {
  if (!grievanceId || (reminderNum !== 1 && reminderNum !== 2)) {
    return { success: false, error: 'Invalid grievance ID or reminder number' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  if (!sheet) {
    return { success: false, error: 'Grievance Log sheet not found' };
  }

  // Find the grievance row
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, error: 'Grievance not found: ' + grievanceId };
  }

  // Determine which columns to update
  const dateCol = reminderNum === 1 ? GRIEVANCE_COLS.REMINDER_1_DATE : GRIEVANCE_COLS.REMINDER_2_DATE;
  const noteCol = reminderNum === 1 ? GRIEVANCE_COLS.REMINDER_1_NOTE : GRIEVANCE_COLS.REMINDER_2_NOTE;

  // Parse date if string
  let parsedDate = reminderDate;
  if (typeof reminderDate === 'string' && reminderDate) {
    parsedDate = new Date(reminderDate);
    if (isNaN(parsedDate.getTime())) {
      return { success: false, error: 'Invalid date format' };
    }
  }

  // Batch update both columns
  const updates = [[parsedDate || '', reminderNote || '']];
  sheet.getRange(rowIndex, dateCol, 1, 2).setValues(updates);

  // Log the action
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      action: 'REMINDER_SET',
      reminderNum: reminderNum,
      reminderDate: parsedDate ? parsedDate.toISOString() : null,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, grievanceId: grievanceId, reminderNum: reminderNum };
}

/**
 * Gets reminders for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @returns {Object} Reminder data or null if not found
 */
function getGrievanceReminders(grievanceId) {
  if (!grievanceId) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      return {
        grievanceId: grievanceId,
        memberName: (data[i][GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' +
                    (data[i][GRIEVANCE_COLS.LAST_NAME - 1] || ''),
        status: data[i][GRIEVANCE_COLS.STATUS - 1] || '',
        reminder1: {
          date: data[i][GRIEVANCE_COLS.REMINDER_1_DATE - 1] || null,
          note: data[i][GRIEVANCE_COLS.REMINDER_1_NOTE - 1] || ''
        },
        reminder2: {
          date: data[i][GRIEVANCE_COLS.REMINDER_2_DATE - 1] || null,
          note: data[i][GRIEVANCE_COLS.REMINDER_2_NOTE - 1] || ''
        }
      };
    }
  }

  return null;
}

/**
 * Clears a reminder for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @param {number} reminderNum - Which reminder to clear (1 or 2)
 * @returns {Object} Result with success status
 */
function clearGrievanceReminder(grievanceId, reminderNum) {
  return setGrievanceReminder(grievanceId, reminderNum, '', '');
}

/**
 * Gets all grievances with reminders due within specified days.
 * Optimized single-pass through grievance data.
 * @param {number} daysAhead - Number of days to look ahead (default 3)
 * @returns {Array} Array of grievances with due reminders
 */
function getDueReminders(daysAhead) {
  daysAhead = daysAhead || 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() + daysAhead);

  const dueReminders = [];

  // Pre-compute column indices
  const colGrievanceId = GRIEVANCE_COLS.GRIEVANCE_ID - 1;
  const colFirstName = GRIEVANCE_COLS.FIRST_NAME - 1;
  const colLastName = GRIEVANCE_COLS.LAST_NAME - 1;
  const colStatus = GRIEVANCE_COLS.STATUS - 1;
  const colSteward = GRIEVANCE_COLS.STEWARD - 1;
  const colR1Date = GRIEVANCE_COLS.REMINDER_1_DATE - 1;
  const colR1Note = GRIEVANCE_COLS.REMINDER_1_NOTE - 1;
  const colR2Date = GRIEVANCE_COLS.REMINDER_2_DATE - 1;
  const colR2Note = GRIEVANCE_COLS.REMINDER_2_NOTE - 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const grievanceId = row[colGrievanceId];
    const status = row[colStatus];

    // Skip closed/resolved grievances
    if (!grievanceId || status === 'Closed' || status === 'Won' ||
        status === 'Denied' || status === 'Withdrawn' || status === 'Settled') {
      continue;
    }

    const memberName = ((row[colFirstName] || '') + ' ' + (row[colLastName] || '')).trim();
    const steward = row[colSteward] || '';

    // Check reminder 1
    const r1Date = row[colR1Date];
    if (r1Date instanceof Date) {
      const r1DateNorm = new Date(r1Date);
      r1DateNorm.setHours(0, 0, 0, 0);

      if (r1DateNorm >= today && r1DateNorm <= cutoffDate) {
        const daysUntil = Math.ceil((r1DateNorm - today) / (1000 * 60 * 60 * 24));
        dueReminders.push({
          grievanceId: grievanceId,
          memberName: memberName,
          status: status,
          steward: steward,
          reminderNum: 1,
          date: r1Date,
          dateFormatted: Utilities.formatDate(r1Date, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          note: row[colR1Note] || '',
          daysUntil: daysUntil,
          isToday: daysUntil === 0,
          isOverdue: daysUntil < 0,
          rowNumber: i + 1
        });
      }
    }

    // Check reminder 2
    const r2Date = row[colR2Date];
    if (r2Date instanceof Date) {
      const r2DateNorm = new Date(r2Date);
      r2DateNorm.setHours(0, 0, 0, 0);

      if (r2DateNorm >= today && r2DateNorm <= cutoffDate) {
        const daysUntil = Math.ceil((r2DateNorm - today) / (1000 * 60 * 60 * 24));
        dueReminders.push({
          grievanceId: grievanceId,
          memberName: memberName,
          status: status,
          steward: steward,
          reminderNum: 2,
          date: r2Date,
          dateFormatted: Utilities.formatDate(r2Date, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          note: row[colR2Note] || '',
          daysUntil: daysUntil,
          isToday: daysUntil === 0,
          isOverdue: daysUntil < 0,
          rowNumber: i + 1
        });
      }
    }
  }

  // Sort by date (soonest first)
  dueReminders.sort(function(a, b) { return a.date - b.date; });

  return dueReminders;
}

/**
 * Shows the reminder management dialog for a grievance.
 * Can be called from Quick Actions or menu.
 * @param {string} grievanceId - Optional grievance ID (uses selected row if not provided)
 */
function showReminderDialog(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // If no grievanceId provided, try to get from selected row
  if (!grievanceId) {
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== (SHEETS.GRIEVANCE_LOG || 'Grievance Log')) {
      ui.alert('Please select a grievance in the Grievance Log sheet, or provide a Grievance ID.');
      return;
    }

    const row = sheet.getActiveRange().getRow();
    if (row <= 1) {
      ui.alert('Please select a grievance row (not the header).');
      return;
    }

    grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  }

  if (!grievanceId) {
    ui.alert('No grievance selected.');
    return;
  }

  // Get current reminders
  const reminders = getGrievanceReminders(grievanceId);
  if (!reminders) {
    ui.alert('Grievance not found: ' + grievanceId);
    return;
  }

  // Build the HTML dialog
  const html = buildReminderDialogHtml_(grievanceId, reminders);
  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  ui.showModalDialog(dialog, '⏰ Reminders: ' + grievanceId);
}

/**
 * Builds the HTML for the reminder dialog.
 * @private
 */
function buildReminderDialogHtml_(grievanceId, reminders) {
  const r1Date = reminders.reminder1.date instanceof Date
    ? Utilities.formatDate(reminders.reminder1.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';
  const r2Date = reminders.reminder2.date instanceof Date
    ? Utilities.formatDate(reminders.reminder2.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';

  return `<!DOCTYPE html>
<html>
<head>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Google Sans', Roboto, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; padding: 20px; color: #F8FAFC; }
    .header { margin-bottom: 20px; }
    .header h2 { font-size: 18px; color: #A78BFA; margin-bottom: 4px; }
    .header .member { font-size: 14px; color: #94A3B8; }
    .reminder-card { background: rgba(255,255,255,0.05); border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 16px; }
    .reminder-card h3 { font-size: 14px; color: #7C3AED; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
    .reminder-card h3::before { content: '⏰'; }
    .form-group { margin-bottom: 12px; }
    .form-label { display: block; font-size: 12px; color: #94A3B8; margin-bottom: 4px; }
    .form-input { width: 100%; padding: 10px 12px; border: 1px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-size: 14px; }
    .form-input:focus { border-color: #7C3AED; outline: none; }
    .form-input::placeholder { color: #64748B; }
    .btn-row { display: flex; gap: 8px; margin-top: 8px; }
    .btn { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; }
    .btn-clear { background: #334155; color: #F8FAFC; }
    .btn-clear:hover { background: #475569; }
    .actions { display: flex; gap: 12px; margin-top: 20px; justify-content: flex-end; }
    .btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; padding: 12px 24px; }
    .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }
    .btn-secondary { background: #334155; color: #F8FAFC; padding: 12px 24px; }
    .status { padding: 8px 12px; border-radius: 6px; margin-top: 12px; font-size: 13px; display: none; }
    .status.success { display: block; background: rgba(16,185,129,0.2); border: 1px solid #10B981; }
    .status.error { display: block; background: rgba(239,68,68,0.2); border: 1px solid #EF4444; }
  </style>
</head>
<body>
  <div class="header">
    <h2>Grievance ${grievanceId}</h2>
    <div class="member">${reminders.memberName} • ${reminders.status}</div>
  </div>

  <div class="reminder-card">
    <h3>Reminder 1</h3>
    <div class="form-group">
      <label class="form-label">Date</label>
      <input type="date" class="form-input" id="r1Date" value="${r1Date}">
    </div>
    <div class="form-group">
      <label class="form-label">Note (e.g., "Schedule Step II meeting")</label>
      <input type="text" class="form-input" id="r1Note" value="${reminders.reminder1.note.replace(/"/g, '&quot;')}" placeholder="Brief description...">
    </div>
    <div class="btn-row">
      <button class="btn btn-clear" onclick="clearReminder(1)">Clear</button>
    </div>
  </div>

  <div class="reminder-card">
    <h3>Reminder 2</h3>
    <div class="form-group">
      <label class="form-label">Date</label>
      <input type="date" class="form-input" id="r2Date" value="${r2Date}">
    </div>
    <div class="form-group">
      <label class="form-label">Note</label>
      <input type="text" class="form-input" id="r2Note" value="${reminders.reminder2.note.replace(/"/g, '&quot;')}" placeholder="Brief description...">
    </div>
    <div class="btn-row">
      <button class="btn btn-clear" onclick="clearReminder(2)">Clear</button>
    </div>
  </div>

  <div id="statusArea"></div>

  <div class="actions">
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn btn-primary" onclick="saveReminders()">Save Reminders</button>
  </div>

  <script>
    var grievanceId = '${grievanceId}';

    function saveReminders() {
      var r1Date = document.getElementById('r1Date').value;
      var r1Note = document.getElementById('r1Note').value;
      var r2Date = document.getElementById('r2Date').value;
      var r2Note = document.getElementById('r2Note').value;

      showStatus('Saving...', false);

      // Save reminder 1
      google.script.run
        .withSuccessHandler(function() {
          // Save reminder 2
          google.script.run
            .withSuccessHandler(function() {
              showStatus('Reminders saved!', false);
              setTimeout(function() { google.script.host.close(); }, 1000);
            })
            .withFailureHandler(function(e) { showStatus('Error: ' + e.message, true); })
            .setGrievanceReminder(grievanceId, 2, r2Date, r2Note);
        })
        .withFailureHandler(function(e) { showStatus('Error: ' + e.message, true); })
        .setGrievanceReminder(grievanceId, 1, r1Date, r1Note);
    }

    function clearReminder(num) {
      if (num === 1) {
        document.getElementById('r1Date').value = '';
        document.getElementById('r1Note').value = '';
      } else {
        document.getElementById('r2Date').value = '';
        document.getElementById('r2Note').value = '';
      }
    }

    function showStatus(msg, isError) {
      var area = document.getElementById('statusArea');
      area.className = 'status ' + (isError ? 'error' : 'success');
      area.textContent = msg;
    }
  </script>
</body>
</html>`;
}

/**
 * Checks for due reminders and sends notifications.
 * Call this from a daily time-driven trigger.
 * @param {number} daysAhead - Days to look ahead (default 1 = today only)
 * @returns {Object} Result with notification count
 */
function checkAndNotifyReminders(daysAhead) {
  daysAhead = daysAhead || 1;

  const dueReminders = getDueReminders(daysAhead);

  if (dueReminders.length === 0) {
    return { success: true, notified: 0, message: 'No reminders due' };
  }

  // Group reminders by steward for consolidated notifications
  const bysteward = {};
  for (let i = 0; i < dueReminders.length; i++) {
    const reminder = dueReminders[i];
    const steward = reminder.steward || 'Unassigned';
    if (!bysteward[steward]) {
      bysteward[steward] = [];
    }
    bysteward[steward].push(reminder);
  }

  // Show toast summary
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayCount = dueReminders.filter(function(r) { return r.isToday; }).length;
  const upcomingCount = dueReminders.length - todayCount;

  let message = '';
  if (todayCount > 0) {
    message += todayCount + ' reminder(s) due TODAY';
  }
  if (upcomingCount > 0) {
    message += (message ? ', ' : '') + upcomingCount + ' upcoming';
  }

  ss.toast(message, '⏰ Grievance Reminders', 10);

  return {
    success: true,
    notified: dueReminders.length,
    todayCount: todayCount,
    upcomingCount: upcomingCount,
    reminders: dueReminders
  };
}

/**
 * Installs the daily reminder check trigger.
 * Call once during setup to enable automatic reminder notifications.
 */
function installReminderTrigger() {
  // Remove existing reminder triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndNotifyReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install new daily trigger at 8 AM
  ScriptApp.newTrigger('checkAndNotifyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily reminder check installed (8 AM)',
    'Trigger Installed',
    5
  );

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.TRIGGER_INSTALLED, {
      triggerName: 'checkAndNotifyReminders',
      schedule: 'Daily at 8 AM',
      installedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true };
}

/**
 * Gets reminder summary for dashboard display.
 * @returns {Object} Summary with counts and upcoming reminders
 */
function getReminderSummary() {
  const todayReminders = getDueReminders(0); // Due today
  const weekReminders = getDueReminders(7);  // Due within a week

  return {
    dueToday: todayReminders.length,
    dueThisWeek: weekReminders.length,
    upcoming: weekReminders.slice(0, 5) // Top 5 for dashboard
  };
}
