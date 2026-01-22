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
