/**
 * ============================================================================
 * FEATURE EXTENSION: DYNAMIC ENGINE & MEMBER LEADERS
 * ============================================================================
 * This script handles dynamic column detection and organizational leadership.
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
  DYNAMIC_FORMULA_ROW: 50, // Start row for dynamic formulas (avoids conflict with existing)
  MEMBER_SHEET: "Member Directory",
  GRIEVANCE_SHEET: "Grievance Log",
  LEADER_ROLE_NAME: "Member Leader",
  CORE_COLUMN_COUNT: 32 // Standard columns in Member Directory (A-AF)
};

/**
 * 1. DYNAMIC HEADER SCANNER
 * Scans any sheet and returns a map of Header Name -> Column Index (1-based).
 * This ensures the script never breaks if you add or move columns.
 */
function getHeaderMap(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {};

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((acc, header, index) => {
    if (header && header.toString().trim() !== "") {
      acc[header.toString().trim()] = index + 1;
    }
    return acc;
  }, {});
}

/**
 * 2. MEMBER LEADER REGISTRY
 * Fetches all members tagged as 'Member Leader' for use in dropdowns.
 * These individuals are organizational and do not appear in Grievance lists.
 *
 * SCHEMA NOTE: Uses IS_STEWARD column (N) which can now accept:
 * - "Yes" = Steward (appears in grievance dropdowns)
 * - "No" = Regular member
 * - "Member Leader" = Organizational leader (filtered OUT of grievance dropdowns)
 */
function getMemberLeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR || EXTENSION_CONFIG.MEMBER_SHEET);

  if (!sheet) {
    Logger.log('getMemberLeaders: Member Directory sheet not found');
    return [];
  }

  const data = sheet.getDataRange().getValues();

  // Use existing column constants from 01_Constants.gs
  const roleIdx = MEMBER_COLS.IS_STEWARD - 1; // Column N (0-indexed)
  const firstNameIdx = MEMBER_COLS.FIRST_NAME - 1; // Column B
  const lastNameIdx = MEMBER_COLS.LAST_NAME - 1; // Column C

  const leaders = [];

  for (let i = 1; i < data.length; i++) { // Skip header row
    const row = data[i];
    const roleValue = row[roleIdx] ? row[roleIdx].toString().trim() : '';

    if (roleValue === EXTENSION_CONFIG.LEADER_ROLE_NAME) {
      const firstName = row[firstNameIdx] || '';
      const lastName = row[lastNameIdx] || '';
      const fullName = (firstName + ' ' + lastName).trim();

      if (fullName) {
        leaders.push({
          name: fullName,
          firstName: firstName,
          lastName: lastName,
          memberId: row[MEMBER_COLS.MEMBER_ID - 1] || '',
          email: row[MEMBER_COLS.EMAIL - 1] || '',
          unit: row[MEMBER_COLS.UNIT - 1] || '',
          location: row[MEMBER_COLS.WORK_LOCATION - 1] || ''
        });
      }
    }
  }

  return leaders;
}

/**
 * Gets stewards for grievance dropdowns, EXCLUDING Member Leaders.
 * This is the integration point with the existing getAllStewards() function.
 * Use this instead of getAllStewards() when populating grievance assignment dropdowns.
 *
 * @returns {Array} Array of steward objects (excludes Member Leaders)
 */
function getStewardsForGrievance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR || EXTENSION_CONFIG.MEMBER_SHEET);

  if (!sheet) {
    Logger.log('getStewardsForGrievance: Member Directory sheet not found');
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const stewards = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isStewardValue = row[MEMBER_COLS.IS_STEWARD - 1];

    // Only include "Yes" - exclude "Member Leader" and "No"
    if (isStewardValue === 'Yes' || isStewardValue === true) {
      const steward = {};
      for (let j = 0; j < headers.length; j++) {
        steward[headers[j]] = row[j];
      }
      steward.rowNumber = i + 1;

      // Add computed full name for convenience
      steward.fullName = ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                          (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim();

      stewards.push(steward);
    }
  }

  return stewards;
}

/**
 * Checks if a member is a Member Leader (not a Steward)
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if member is a Member Leader
 */
function isMemberLeader(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR || EXTENSION_CONFIG.MEMBER_SHEET);

  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      const roleValue = data[i][MEMBER_COLS.IS_STEWARD - 1];
      return roleValue === EXTENSION_CONFIG.LEADER_ROLE_NAME;
    }
  }

  return false;
}

/**
 * 3. SELF-HEALING DYNAMIC FORMULAS
 * Re-injects logic into a dedicated section of the hidden sheet.
 * Uses rows 50+ to avoid conflicts with existing _Dashboard_Calc formulas.
 *
 * This uses a "Column-Agnostic" approach via dynamic header scanning.
 */
function repairDynamicFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);

  // Create if missing (normally created by setupAllHiddenSheets)
  if (!calcSheet) {
    calcSheet = ss.insertSheet(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
    calcSheet.hideSheet();
  }

  const startRow = EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW; // Row 50

  // Use existing column constant for IS_STEWARD (column N = 14)
  const isStewardCol = MEMBER_COLS.IS_STEWARD;

  // Label for this section
  calcSheet.getRange(startRow, 1).setValue('=== DYNAMIC ENGINE FORMULAS ===');

  // Formula 1: Count of Member Leaders vs Stewards vs Regular Members
  // This "heals" itself by using the column position from constants
  const roleCountFormula = `=QUERY('${EXTENSION_CONFIG.MEMBER_SHEET}'!A:ZZ, "SELECT Col${isStewardCol}, count(Col1) WHERE Col1 IS NOT NULL GROUP BY Col${isStewardCol} LABEL count(Col1) 'Total'", 1)`;
  calcSheet.getRange(startRow + 1, 1).setValue('Role Distribution:');
  calcSheet.getRange(startRow + 2, 1).setFormula(roleCountFormula);

  // Formula 2: Member Leader count
  const leaderCountFormula = `=COUNTIF('${EXTENSION_CONFIG.MEMBER_SHEET}'!${getColumnLetter(isStewardCol)}:${getColumnLetter(isStewardCol)},"${EXTENSION_CONFIG.LEADER_ROLE_NAME}")`;
  calcSheet.getRange(startRow + 6, 1).setValue('Member Leaders:');
  calcSheet.getRange(startRow + 6, 2).setFormula(leaderCountFormula);

  // Formula 3: Active Stewards count (excludes Member Leaders)
  const stewardCountFormula = `=COUNTIF('${EXTENSION_CONFIG.MEMBER_SHEET}'!${getColumnLetter(isStewardCol)}:${getColumnLetter(isStewardCol)},"Yes")`;
  calcSheet.getRange(startRow + 7, 1).setValue('Active Stewards:');
  calcSheet.getRange(startRow + 7, 2).setFormula(stewardCountFormula);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Dynamic formulas injected at row ' + startRow,
    'Self-Healing Complete',
    5
  );

  return { success: true, startRow: startRow };
}

/**
 * 4. DYNAMIC DATA COLLECTION (The "Hook")
 * This function is designed to be called by your Modal.
 * It gathers all data from "Extra" columns added by the user.
 *
 * @param {string} memberId - The member ID to fetch data for
 * @returns {Object} Object with extraHeaders array and memberData object
 */
function getExpansionColumnData(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR || EXTENSION_CONFIG.MEMBER_SHEET);

  if (!sheet) {
    return { extraHeaders: [], memberData: {} };
  }

  const headerMap = getHeaderMap(EXTENSION_CONFIG.MEMBER_SHEET);
  const coreCount = EXTENSION_CONFIG.CORE_COLUMN_COUNT; // 32 standard columns (A-AF)

  // Get all headers and identify extra/custom ones
  const allHeaders = Object.keys(headerMap);
  const extraHeaders = [];

  for (const header of allHeaders) {
    if (headerMap[header] > coreCount) {
      extraHeaders.push({
        name: header,
        column: headerMap[header],
        letter: getColumnLetter(headerMap[header])
      });
    }
  }

  // If memberId provided, fetch that member's data for extra columns
  const memberData = {};
  if (memberId) {
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
        // Extract data for each extra column
        for (const extra of extraHeaders) {
          memberData[extra.name] = data[i][extra.column - 1] || '';
        }
        break;
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
 * 5. EXPANSION COLUMN UI GENERATOR
 * Generates HTML form fields for custom columns.
 * Called by modal dialogs to dynamically add fields.
 *
 * @param {string} memberId - Optional member ID to pre-fill values
 * @returns {string} HTML string with form fields for extra columns
 */
function generateExpansionFieldsHtml(memberId) {
  const expansionData = getExpansionColumnData(memberId);

  if (expansionData.extraHeaders.length === 0) {
    return '<!-- No custom columns detected -->';
  }

  let html = '<div class="expansion-fields" style="margin-top: 16px; padding-top: 16px; border-top: 1px solid #E5E7EB;">';
  html += '<div class="expansion-header" style="font-weight: 600; color: #7C3AED; margin-bottom: 12px;">Custom Fields</div>';

  for (const field of expansionData.extraHeaders) {
    const value = expansionData.memberData[field.name] || '';
    const fieldId = 'custom_' + field.name.replace(/[^a-zA-Z0-9]/g, '_');

    html += `
      <div class="form-group" style="margin-bottom: 12px;">
        <label class="form-label" style="display: block; font-size: 13px; color: #374151; margin-bottom: 4px;">
          ${field.name}
        </label>
        <input type="text"
               class="form-input"
               id="${fieldId}"
               name="${fieldId}"
               data-column="${field.column}"
               value="${value.toString().replace(/"/g, '&quot;')}"
               style="width: 100%; padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-family: Roboto, sans-serif;">
      </div>
    `;
  }

  html += '</div>';
  return html;
}

/**
 * 6. SAVE EXPANSION DATA
 * Saves custom column data for a member.
 *
 * @param {string} memberId - The member ID
 * @param {Object} customData - Object with field names as keys and values
 * @returns {Object} Result with success status
 */
function saveExpansionData(memberId, customData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR || EXTENSION_CONFIG.MEMBER_SHEET);

  if (!sheet || !memberId) {
    return { success: false, error: 'Sheet or member ID not found' };
  }

  const headerMap = getHeaderMap(EXTENSION_CONFIG.MEMBER_SHEET);
  const data = sheet.getDataRange().getValues();

  // Find member row
  let memberRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      memberRow = i + 1;
      break;
    }
  }

  if (memberRow === -1) {
    return { success: false, error: 'Member not found' };
  }

  // Update each custom field
  let updated = 0;
  for (const fieldName in customData) {
    if (headerMap[fieldName] && headerMap[fieldName] > EXTENSION_CONFIG.CORE_COLUMN_COUNT) {
      sheet.getRange(memberRow, headerMap[fieldName]).setValue(customData[fieldName]);
      updated++;
    }
  }

  return { success: true, updated: updated };
}

// ============================================================================
// NOTE: columnToLetter helper removed - use getColumnLetter() from 01_Constants.gs
// The existing getColumnLetter(columnNumber) function performs the same conversion.
// ============================================================================

// ============================================================================
// SETUP & CONFIGURATION
// ============================================================================

/**
 * Adds "Member Leader" option to the IS_STEWARD validation dropdown.
 * Run this once after deploying the Dynamic Engine feature.
 *
 * This updates the Config sheet's YES_NO column (E) to include "Member Leader"
 * as a valid option for the IS_STEWARD column.
 */
function setupMemberLeaderRole() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG || 'Config');

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found');
    return { success: false, error: 'Config sheet not found' };
  }

  // The YES_NO column in Config (column E) provides dropdown values for IS_STEWARD
  const yesNoCol = CONFIG_COLS.YES_NO || 5;

  // Get existing values
  const existingValues = configSheet.getRange(3, yesNoCol, 10, 1).getValues().flat().filter(v => v);

  // Check if "Member Leader" already exists
  if (existingValues.includes(EXTENSION_CONFIG.LEADER_ROLE_NAME)) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Member Leader role already configured',
      'Setup Complete',
      3
    );
    return { success: true, alreadyExists: true };
  }

  // Add "Member Leader" to the list
  const nextRow = existingValues.length + 3;
  configSheet.getRange(nextRow, yesNoCol).setValue(EXTENSION_CONFIG.LEADER_ROLE_NAME);

  // Log the setup
  if (typeof logAuditEvent === 'function') {
    logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
      action: 'MEMBER_LEADER_SETUP',
      addedValue: EXTENSION_CONFIG.LEADER_ROLE_NAME,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Member Leader role added to Config. Update IS_STEWARD validations if needed.',
    'Setup Complete',
    5
  );

  return { success: true, addedAt: nextRow };
}

/**
 * Runs all Dynamic Engine setup tasks.
 * Call this from the Admin menu after deployment.
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

  // Step 1: Setup Member Leader role
  const roleResult = setupMemberLeaderRole();

  // Step 2: Inject self-healing formulas
  const formulaResult = repairDynamicFormulas();

  // Step 3: Scan for expansion columns
  const expansionData = getExpansionColumnData();

  // Report results
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
 * @returns {Object} Status object with feature states
 */
function getDynamicEngineStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check for Member Leaders
  const leaders = getMemberLeaders();

  // Check for expansion columns
  const expansionData = getExpansionColumnData();

  // Check for self-healing formulas
  const calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
  let formulasInjected = false;
  if (calcSheet) {
    const checkValue = calcSheet.getRange(EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW, 1).getValue();
    formulasInjected = checkValue && checkValue.toString().includes('DYNAMIC ENGINE');
  }

  return {
    memberLeaders: {
      count: leaders.length,
      names: leaders.map(l => l.name)
    },
    expansionColumns: {
      count: expansionData.extraHeaders.length,
      headers: expansionData.extraHeaders.map(h => h.name)
    },
    selfHealing: {
      active: formulasInjected,
      sheet: EXTENSION_CONFIG.HIDDEN_CALC_SHEET,
      row: EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW
    }
  };
}
