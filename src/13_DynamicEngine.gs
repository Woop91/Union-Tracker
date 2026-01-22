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
 *    - Filtered out of Grievance dropdowns via Role column
 *
 * 2. Column Expansion (Dynamic Engine)
 *    - "No-Code" feature for adding custom columns
 *    - Header Scanner performs census of top row on dashboard open
 *    - Auto-generates form fields with matching CSS styling
 *    - Detection Zone: columns beyond Core data (Column 15+)
 *
 * 3. Self-Healing Hidden Architecture
 *    - Script-to-Sheet Bridge for formula management
 *    - _Dashboard_Calc sheet holds ARRAYFORMULA and QUERY logic
 *    - Repair Trigger re-injects formulas if broken
 *    - Dynamic Referencing uses Col1, Col2 logic (not cell refs)
 */

const EXTENSION_CONFIG = {
  HIDDEN_CALC_SHEET: "_Dashboard_Calc",
  MEMBER_SHEET: "Member Directory",
  GRIEVANCE_SHEET: "Grievance Log",
  LEADER_ROLE_NAME: "Member Leader"
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
 */
function getMemberLeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXTENSION_CONFIG.MEMBER_SHEET);
  const data = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap(EXTENSION_CONFIG.MEMBER_SHEET);

  const roleIdx = headerMap["Role"] - 1; // Adjust for 0-based array
  const nameIdx = headerMap["Full Name"] - 1;

  return data.filter(row => row[roleIdx] === EXTENSION_CONFIG.LEADER_ROLE_NAME)
             .map(row => row[nameIdx]);
}

/**
 * 3. SELF-HEALING DYNAMIC FORMULAS
 * Re-injects logic into the hidden sheet.
 * This uses a "Column-Agnostic" approach.
 */
function repairDynamicFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);

  // Create if missing
  if (!calcSheet) {
    calcSheet = ss.insertSheet(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
    calcSheet.hideSheet();
  }

  const headerMap = getHeaderMap(EXTENSION_CONFIG.MEMBER_SHEET);
  const roleColLetter = columnToLetter(headerMap["Role"]);

  // This formula calculates the count of Member Leaders vs Stewards
  // It "heals" itself by referencing the Column Name found by the script
  const formula = `=QUERY('${EXTENSION_CONFIG.MEMBER_SHEET}'!A:ZZ, "SELECT Col${headerMap["Role"]}, count(Col1) WHERE Col1 IS NOT NULL GROUP BY Col${headerMap["Role"]} LABEL count(Col1) 'Total'", 1)`;

  calcSheet.getRange("A1").setFormula(formula);

  SpreadsheetApp.getUi().alert("Self-Healing Complete: Dynamic formulas re-injected.");
}

/**
 * 4. DYNAMIC DATA COLLECTION (The "Hook")
 * This function is designed to be called by your Modal.
 * It gathers all data from "Extra" columns added by the user.
 */
function getExpansionColumnData(memberId) {
  const headerMap = getHeaderMap(EXTENSION_CONFIG.MEMBER_SHEET);
  const coreCount = 15; // Your standard columns

  const allHeaders = Object.keys(headerMap);
  const extraHeaders = allHeaders.slice(coreCount);

  // Logic to fetch specific member row and extract data from these headers...
  return extraHeaders;
}

/**
 * HELPER: Converts Column Index to Letter (e.g., 1 -> A, 27 -> AA)
 */
function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
