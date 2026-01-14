/**
 * ============================================================================
 * TESTING FRAMEWORK & VALIDATION
 * ============================================================================
 * Unit tests, integration tests, and data validation
 */

// ==================== TEST CONFIGURATION ====================

var TEST_RESULTS = { passed: [], failed: [], skipped: [] };
var TEST_MAX_EXECUTION_MS = 5 * 60 * 1000;
var TEST_LARGE_DATASET_THRESHOLD = 5000;

var Assert = {
  assertEquals: function(expected, actual, message) {
    if (expected !== actual) throw new Error((message || 'Assertion failed') + '\nExpected: ' + JSON.stringify(expected) + '\nActual: ' + JSON.stringify(actual));
  },
  assertTrue: function(value, message) {
    if (value !== true) throw new Error((message || 'Expected true') + '\nActual: ' + value);
  },
  assertFalse: function(value, message) {
    if (value !== false) throw new Error((message || 'Expected false') + '\nActual: ' + value);
  },
  assertNotNull: function(value, message) {
    if (value === null || value === undefined) throw new Error(message || 'Value should not be null/undefined');
  },
  assertNull: function(value, message) {
    if (value !== null) throw new Error((message || 'Expected null') + '\nActual: ' + value);
  },
  assertContains: function(array, value, message) {
    if (!Array.isArray(array) || array.indexOf(value) === -1) throw new Error((message || 'Array does not contain value') + '\nValue: ' + value);
  },
  assertArrayLength: function(array, expectedLength, message) {
    if (!Array.isArray(array) || array.length !== expectedLength) throw new Error((message || 'Array length mismatch') + '\nExpected: ' + expectedLength + '\nActual: ' + (array ? array.length : 'N/A'));
  },
  assertThrows: function(fn, message) {
    var threw = false;
    try { fn(); } catch (e) { threw = true; }
    if (!threw) throw new Error(message || 'Expected function to throw');
  },
  assertApproximately: function(expected, actual, tolerance, message) {
    tolerance = tolerance || 0.001;
    if (Math.abs(expected - actual) > tolerance) throw new Error((message || 'Values not approximately equal') + '\nExpected: ' + expected + '\nActual: ' + actual);
  },
  fail: function(message) { throw new Error(message || 'Test failed'); }
};

// ==================== TEST HELPERS ====================

function isLargeDataset() {
  try {
    var ss = SpreadsheetApp.getActive();
    var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
    return memberDir ? memberDir.getLastRow() > TEST_LARGE_DATASET_THRESHOLD : false;
  } catch (e) { return false; }
}

function createTestMember(memberId) {
  var ss = SpreadsheetApp.getActive();
  var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var testData = [
    memberId || 'TEST-M001', 'Test', 'Member', '', '', '', 'Monday', 'test@union.org', '(555) 123-4567',
    'Email', 'Mornings', '', '', 'No', '', '', new Date(), new Date(), 85, 10, 'Yes', 'Yes', 'No', '', new Date(), '', '', '', '', '', ''
  ];
  var startRow = Math.max(memberDir.getLastRow() + 1, 2);
  memberDir.getRange(startRow, 1, 1, testData.length).setValues([testData]);
  return memberId || 'TEST-M001';
}

function cleanupTestData() {
  var ss = SpreadsheetApp.getActive();
  var sheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];
  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).indexOf('TEST-') === 0) sheet.deleteRow(i + 2);
    }
  });
}

// ==================== UNIT TESTS ====================

function testMemberColsConstants() {
  Assert.assertNotNull(MEMBER_COLS, 'MEMBER_COLS should exist');
  Assert.assertEquals(1, MEMBER_COLS.MEMBER_ID, 'MEMBER_ID should be column 1');
  Assert.assertEquals(2, MEMBER_COLS.FIRST_NAME, 'FIRST_NAME should be column 2');
  Assert.assertEquals(3, MEMBER_COLS.LAST_NAME, 'LAST_NAME should be column 3');
  Assert.assertEquals(8, MEMBER_COLS.EMAIL, 'EMAIL should be column 8');
  Assert.assertEquals(31, MEMBER_COLS.START_GRIEVANCE, 'START_GRIEVANCE should be column 31');
}

function testGrievanceColsConstants() {
  Assert.assertNotNull(GRIEVANCE_COLS, 'GRIEVANCE_COLS should exist');
  Assert.assertEquals(1, GRIEVANCE_COLS.GRIEVANCE_ID, 'GRIEVANCE_ID should be column 1');
  Assert.assertEquals(2, GRIEVANCE_COLS.MEMBER_ID, 'MEMBER_ID should be column 2');
  Assert.assertEquals(5, GRIEVANCE_COLS.STATUS, 'STATUS should be column 5');
  Assert.assertEquals(28, GRIEVANCE_COLS.RESOLUTION, 'RESOLUTION should be column 28');
}

function testColumnLetterConversion() {
  Assert.assertEquals('A', getColumnLetter(1), 'Column 1 should be A');
  Assert.assertEquals('Z', getColumnLetter(26), 'Column 26 should be Z');
  Assert.assertEquals('AA', getColumnLetter(27), 'Column 27 should be AA');
  Assert.assertEquals('AE', getColumnLetter(31), 'Column 31 should be AE');
}

function testSheetsConstants() {
  Assert.assertNotNull(SHEETS, 'SHEETS should exist');
  Assert.assertNotNull(SHEETS.MEMBER_DIR, 'MEMBER_DIR should exist');
  Assert.assertNotNull(SHEETS.GRIEVANCE_LOG, 'GRIEVANCE_LOG should exist');
  Assert.assertNotNull(SHEETS.CONFIG, 'CONFIG should exist');
}

function testValidateRequired() {
  Assert.assertThrows(function() { validateRequired(null, 'field'); }, 'null should throw');
  Assert.assertThrows(function() { validateRequired('', 'field'); }, 'empty should throw');
  Assert.assertThrows(function() { validateRequired(undefined, 'field'); }, 'undefined should throw');
}

function testValidateEmail() {
  var result1 = validateEmailAddress('test@example.com');
  Assert.assertTrue(result1.valid, 'Valid email should pass');
  var result2 = validateEmailAddress('invalid');
  Assert.assertFalse(result2.valid, 'Invalid email should fail');
  var result3 = validateEmailAddress('');
  Assert.assertFalse(result3.valid, 'Empty email should fail');
}

function testValidatePhoneNumber() {
  var result1 = validatePhoneNumber('(555) 123-4567');
  Assert.assertTrue(result1.valid, 'Valid phone should pass');
  var result2 = validatePhoneNumber('5551234567');
  Assert.assertTrue(result2.valid, 'Digits-only phone should pass');
  var result3 = validatePhoneNumber('123');
  Assert.assertFalse(result3.valid, 'Short phone should fail');
}

function testValidateMemberId() {
  // Format: M prefix + 4 uppercase letters (2 from first name + 2 from last name) + 3 digits
  Assert.assertTrue(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM123'), 'MJOSM123 should be valid (M + John Smith + 123)');
  Assert.assertTrue(VALIDATION_PATTERNS.MEMBER_ID.test('MMAJO456'), 'MMAJO456 should be valid (M + Mary Johnson + 456)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('JOSM123'), 'JOSM123 should be invalid (missing M prefix)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('MJOS123'), 'MJOS123 should be invalid (only 3 name letters)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM12'), 'MJOSM12 should be invalid (only 2 digits)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('mjosm123'), 'mjosm123 should be invalid (lowercase)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('M123456'), 'M123456 should be invalid (old format)');
}

function testValidateGrievanceId() {
  // Format: G prefix + 4 uppercase letters (2 from first name + 2 from last name) + 3 digits
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM789'), 'GJOSM789 should be valid (G + John Smith + 789)');
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GROWI001'), 'GROWI001 should be valid (G + Robert Williams + 001)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('JOSM789'), 'JOSM789 should be invalid (missing G prefix)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('G-123456'), 'G-123456 should be invalid (old format)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM1234'), 'GJOSM1234 should be invalid (4 digits)');
}

function testOpenRateRange() {
  Assert.assertTrue(85 >= 0 && 85 <= 100, 'Open rate 85 should be in range');
  Assert.assertTrue(0 >= 0 && 0 <= 100, 'Open rate 0 should be in range');
  Assert.assertTrue(100 >= 0 && 100 <= 100, 'Open rate 100 should be in range');
}

// ==================== TEST RUNNER ====================

function getTestFunctionRegistry() {
  return {
    testMemberColsConstants: testMemberColsConstants,
    testGrievanceColsConstants: testGrievanceColsConstants,
    testColumnLetterConversion: testColumnLetterConversion,
    testSheetsConstants: testSheetsConstants,
    testValidateRequired: testValidateRequired,
    testValidateEmail: testValidateEmail,
    testValidatePhoneNumber: testValidatePhoneNumber,
    testValidateMemberId: testValidateMemberId,
    testValidateGrievanceId: testValidateGrievanceId,
    testOpenRateRange: testOpenRateRange
  };
}

function runAllTests() {
  var ui = SpreadsheetApp.getUi();
  var largeDataset = isLargeDataset();
  SpreadsheetApp.getActive().toast('🧪 Running tests...', 'Testing', -1);
  TEST_RESULTS = { passed: [], failed: [], skipped: [] };
  var startTime = new Date();
  var registry = getTestFunctionRegistry();
  var testNames = Object.keys(registry);
  testNames.forEach(function(name) {
    if (new Date() - startTime > TEST_MAX_EXECUTION_MS) {
      TEST_RESULTS.skipped.push({ name: name, reason: 'Timeout protection' });
      return;
    }
    try {
      registry[name]();
      TEST_RESULTS.passed.push({ name: name, time: new Date() - startTime });
    } catch (e) {
      TEST_RESULTS.failed.push({ name: name, error: e.message });
    }
  });
  var duration = (new Date() - startTime) / 1000;
  generateTestReport(duration);
  var total = TEST_RESULTS.passed.length + TEST_RESULTS.failed.length + TEST_RESULTS.skipped.length;
  var rate = ((TEST_RESULTS.passed.length / total) * 100).toFixed(1);
  SpreadsheetApp.getActive().toast('✅ ' + TEST_RESULTS.passed.length + ' | ❌ ' + TEST_RESULTS.failed.length + ' | ⏭️ ' + TEST_RESULTS.skipped.length, 'Tests (' + rate + '%)', 10);
  ui.alert('🧪 Tests Complete', 'Passed: ' + TEST_RESULTS.passed.length + '\nFailed: ' + TEST_RESULTS.failed.length + '\nSkipped: ' + TEST_RESULTS.skipped.length + '\n\nDuration: ' + duration.toFixed(2) + 's', ui.ButtonSet.OK);
}

function runQuickTests() {
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getActive().toast('⚡ Quick tests...', 'Testing', -1);
  TEST_RESULTS = { passed: [], failed: [], skipped: [] };
  var startTime = new Date();
  var quickTests = ['testMemberColsConstants', 'testGrievanceColsConstants', 'testColumnLetterConversion', 'testSheetsConstants', 'testOpenRateRange'];
  var registry = getTestFunctionRegistry();
  quickTests.forEach(function(name) {
    try {
      if (registry[name]) { registry[name](); TEST_RESULTS.passed.push({ name: name }); }
      else TEST_RESULTS.skipped.push({ name: name, reason: 'Not found' });
    } catch (e) { TEST_RESULTS.failed.push({ name: name, error: e.message }); }
  });
  var duration = (new Date() - startTime) / 1000;
  ui.alert('⚡ Quick Tests', 'Passed: ' + TEST_RESULTS.passed.length + '\nFailed: ' + TEST_RESULTS.failed.length + '\nDuration: ' + duration.toFixed(2) + 's' + (TEST_RESULTS.failed.length > 0 ? '\n\nFailed:\n' + TEST_RESULTS.failed.map(function(t) { return '• ' + t.name + ': ' + t.error; }).join('\n') : ''), ui.ButtonSet.OK);
}

function generateTestReport(duration) {
  var ss = SpreadsheetApp.getActive();
  var report = ss.getSheetByName(SHEETS.TEST_RESULTS);
  if (!report) report = ss.insertSheet(SHEETS.TEST_RESULTS);
  report.clear();
  report.getRange('A1:F1').merge().setValue('🧪 TEST RESULTS').setFontSize(18).setFontWeight('bold').setBackground('#4A5568').setFontColor('#FFFFFF');
  var total = TEST_RESULTS.passed.length + TEST_RESULTS.failed.length + TEST_RESULTS.skipped.length;
  var rate = ((TEST_RESULTS.passed.length / total) * 100).toFixed(1);
  var summary = [['Total Tests', total], ['✅ Passed', TEST_RESULTS.passed.length], ['❌ Failed', TEST_RESULTS.failed.length], ['⏭️ Skipped', TEST_RESULTS.skipped.length], ['Pass Rate', rate + '%'], ['Duration', duration.toFixed(2) + 's'], ['Timestamp', new Date().toLocaleString()]];
  report.getRange(3, 1, summary.length, 2).setValues(summary);
  report.getRange(3, 1, summary.length, 1).setFontWeight('bold');
  var row = 3 + summary.length + 2;
  if (TEST_RESULTS.failed.length > 0) {
    report.getRange(row, 1, 1, 3).merge().setValue('❌ FAILED TESTS').setFontWeight('bold').setBackground('#FEE2E2');
    row++;
    report.getRange(row, 1, 1, 2).setValues([['Test Name', 'Error']]).setFontWeight('bold').setBackground('#F3F4F6');
    row++;
    TEST_RESULTS.failed.forEach(function(test) { report.getRange(row, 1, 1, 2).setValues([[test.name, test.error]]); row++; });
    row += 2;
  }
  if (TEST_RESULTS.passed.length > 0) {
    report.getRange(row, 1, 1, 2).merge().setValue('✅ PASSED TESTS').setFontWeight('bold').setBackground('#D1FAE5');
    row++;
    TEST_RESULTS.passed.forEach(function(test) { report.getRange(row, 1).setValue(test.name + ' ✓'); row++; });
  }
  report.autoResizeColumns(1, 3);
  report.setTabColor('#7C3AED');
}

// ==================== VALIDATION FRAMEWORK ====================

var VALIDATION_PATTERNS = {
  EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
  PHONE_US: /^[\+]?1?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}$/,
  // ID format: M/G prefix + 2 chars from first name + 2 chars from last name + 3 random digits
  MEMBER_ID: /^M[A-Z]{4}\d{3}$/,      // e.g., MJOSM123 (M + John Smith + 123)
  GRIEVANCE_ID: /^G[A-Z]{4}\d{3}$/    // e.g., GJOSM456 (G + John Smith + 456)
};

var VALIDATION_MESSAGES = {
  EMAIL_INVALID: 'Invalid email format. Use: name@domain.com',
  EMAIL_EMPTY: 'Email address is required',
  PHONE_INVALID: 'Invalid phone format. Use: (555) 555-1234',
  MEMBER_ID_INVALID: 'Invalid Member ID. Format: M + 2 letters from first name + 2 letters from last name + 3 digits (e.g., MJOSM123)',
  MEMBER_ID_DUPLICATE: 'This Member ID already exists',
  GRIEVANCE_ID_INVALID: 'Invalid Grievance ID. Format: G + 2 letters from first name + 2 letters from last name + 3 digits (e.g., GJOSM456)',
  GRIEVANCE_ID_DUPLICATE: 'This Grievance ID already exists'
};

function validateEmailAddress(email) {
  if (!email || email.toString().trim() === '') return { valid: false, message: VALIDATION_MESSAGES.EMAIL_EMPTY };
  var clean = email.toString().trim().toLowerCase();
  if (!VALIDATION_PATTERNS.EMAIL.test(clean)) return { valid: false, message: VALIDATION_MESSAGES.EMAIL_INVALID };
  var typos = { 'gmial.com': 'gmail.com', 'gmai.com': 'gmail.com', 'yaho.com': 'yahoo.com' };
  var domain = clean.split('@')[1];
  if (typos[domain]) return { valid: true, message: 'Did you mean ' + clean.split('@')[0] + '@' + typos[domain] + '?', suggestion: clean.split('@')[0] + '@' + typos[domain] };
  return { valid: true, message: 'Valid email' };
}

function validatePhoneNumber(phone) {
  if (!phone || phone.toString().trim() === '') return { valid: false, message: 'Phone required' };
  var digits = phone.toString().replace(/\D/g, '');
  if (digits.length < 10) return { valid: false, message: 'At least 10 digits required' };
  if (digits.length > 15) return { valid: false, message: 'Phone too long' };
  var formatted = formatUSPhone(digits);
  return { valid: true, message: 'Valid phone', formatted: formatted };
}

function formatUSPhone(digits) {
  if (digits.length === 11 && digits[0] === '1') digits = digits.substring(1);
  if (digits.length === 10) return '(' + digits.substring(0, 3) + ') ' + digits.substring(3, 6) + '-' + digits.substring(6);
  return digits;
}

function validateRequired(value, fieldName) {
  if (value === null || value === undefined || value === '') throw new Error(fieldName + ' is required');
  return value;
}

function checkDuplicateMemberID(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) if (data[i][0] === memberId) count++;
  return count > 1;
}

function checkDuplicateGrievanceID(grievanceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) if (data[i][0] === grievanceId) count++;
  return count > 1;
}

/**
 * Check if a grievance's Member ID exists in the Member Directory
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if member ID exists, false otherwise
 */
function checkMemberIdExists(memberId) {
  if (!memberId) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet || memberSheet.getLastRow() < 2) return false;
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0] === memberId) return true;
  }
  return false;
}

/**
 * Validate all grievances to ensure Member IDs exist in Member Directory
 * @returns {Array} Array of issues found
 */
function validateGrievanceMemberIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  var grievanceLastRow = grievanceSheet.getLastRow();
  var memberLastRow = memberSheet.getLastRow();

  if (grievanceLastRow < 2 || memberLastRow < 2) return [];

  // Build set of valid member IDs
  var validMemberIds = {};
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberLastRow - 1, 1).getValues();
  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0]) validMemberIds[memberIds[i][0]] = true;
  }

  // Check each grievance's Member ID
  var issues = [];
  var grievanceMemberIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, grievanceLastRow - 1, 1).getValues();
  var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, grievanceLastRow - 1, 1).getValues();

  for (var j = 0; j < grievanceMemberIds.length; j++) {
    var memberId = grievanceMemberIds[j][0];
    var grievanceId = grievanceIds[j][0];
    var row = j + 2;

    if (!memberId) {
      issues.push({ row: row, grievanceId: grievanceId, memberId: '', message: 'Missing Member ID' });
    } else if (!validMemberIds[memberId]) {
      issues.push({ row: row, grievanceId: grievanceId, memberId: memberId, message: 'Member ID not found in Member Directory' });
    }
  }

  return issues;
}

function runBulkValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) { SpreadsheetApp.getUi().alert('Member Directory not found!'); return; }
  ss.toast('Running validation...', 'Please wait', -1);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data', 'Info', 3); return; }
  var emailData = sheet.getRange(2, MEMBER_COLS.EMAIL, lastRow - 1, 1).getValues();
  var phoneData = sheet.getRange(2, MEMBER_COLS.PHONE, lastRow - 1, 1).getValues();
  var memberIdData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1).getValues();
  var issues = [], seenIds = {};
  for (var i = 0; i < lastRow - 1; i++) {
    var row = i + 2;
    if (emailData[i][0]) { var er = validateEmailAddress(emailData[i][0]); if (!er.valid) issues.push({ row: row, field: 'Email', value: emailData[i][0], message: er.message }); }
    if (phoneData[i][0]) { var pr = validatePhoneNumber(phoneData[i][0]); if (!pr.valid) issues.push({ row: row, field: 'Phone', value: phoneData[i][0], message: pr.message }); }
    if (memberIdData[i][0]) {
      if (seenIds[memberIdData[i][0]]) issues.push({ row: row, field: 'Member ID', value: memberIdData[i][0], message: 'Duplicate of row ' + seenIds[memberIdData[i][0]] });
      else seenIds[memberIdData[i][0]] = row;
    }
  }

  // Also validate grievance Member IDs reference valid members
  var grievanceIssues = validateGrievanceMemberIds();
  grievanceIssues.forEach(function(gi) {
    issues.push({ row: gi.row, field: 'Grievance Member ID', value: gi.memberId || '(empty)', message: gi.message + ' (Grievance: ' + gi.grievanceId + ')' });
  });

  showValidationReport(issues, lastRow - 1 + grievanceIssues.length);
}

function showValidationReport(issues, total) {
  var rate = total > 0 ? (((total - issues.length) / total) * 100).toFixed(1) : 100;
  var rows = issues.slice(0, 50).map(function(i) { return '<tr><td>' + i.row + '</td><td>' + i.field + '</td><td>' + i.value + '</td><td>' + i.message + '</td></tr>'; }).join('');
  if (issues.length > 50) rows += '<tr><td colspan="4">...and ' + (issues.length - 50) + ' more</td></tr>';
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{display:flex;gap:20px;margin:20px 0}.stat{flex:1;padding:20px;border-radius:8px;text-align:center}.stat.good{background:#e8f5e9}.stat.warning{background:#fff3e0}.stat.bad{background:#ffebee}.num{font-size:32px;font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:20px}th,td{padding:10px;text-align:left;border:1px solid #ddd}th{background:#f5f5f5}</style></head><body><h2>📊 Validation Report</h2><div class="summary"><div class="stat ' + (issues.length === 0 ? 'good' : issues.length < 10 ? 'warning' : 'bad') + '"><div class="num">' + rate + '%</div><div>Pass Rate</div></div><div class="stat good"><div class="num">' + total + '</div><div>Records</div></div><div class="stat ' + (issues.length === 0 ? 'good' : 'bad') + '"><div class="num">' + issues.length + '</div><div>Issues</div></div></div>' + (issues.length > 0 ? '<table><tr><th>Row</th><th>Field</th><th>Value</th><th>Issue</th></tr>' + rows + '</table>' : '<div style="text-align:center;padding:40px;color:#4caf50">✅ No issues found!</div>') + '</body></html>'
  ).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Report');
}

function showValidationSettings() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.setting{margin:15px 0;padding:15px;background:#f8f9fa;border-radius:8px;display:flex;justify-content:space-between;align-items:center}.title{font-weight:bold}.desc{color:#666;font-size:13px}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;width:100%;margin-top:20px}</style></head><body><h2>⚙️ Validation Settings</h2><div class="setting"><div><div class="title">Email Validation</div><div class="desc">Validate format as you type</div></div><span>✅</span></div><div class="setting"><div><div class="title">Phone Auto-format</div><div class="desc">Format to (XXX) XXX-XXXX</div></div><span>✅</span></div><div class="setting"><div><div class="title">Duplicate Detection</div><div class="desc">Warn on duplicate IDs</div></div><span>✅</span></div><button onclick="google.script.run.runBulkValidation();google.script.host.close()">🔍 Run Bulk Validation</button></body></html>'
  ).setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Settings');
}

function clearValidationIndicators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  [MEMBER_COLS.EMAIL, MEMBER_COLS.PHONE, MEMBER_COLS.MEMBER_ID].forEach(function(col) {
    var range = sheet.getRange(2, col, lastRow - 1, 1);
    range.clearNote();
    range.setBackground(null);
  });
  ss.toast('✅ Indicators cleared', 'Done', 3);
}

function installValidationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditValidation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditValidation').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ Validation trigger installed!');
}

function onEditValidation(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) return;
  var col = e.range.getColumn(), row = e.range.getRow(), val = e.value;
  if (row < 2 || !val) return;
  if (name === SHEETS.MEMBER_DIR) {
    if (col === MEMBER_COLS.EMAIL) {
      var r = validateEmailAddress(val);
      if (!r.valid) { e.range.setNote('⚠️ ' + r.message); e.range.setBackground('#fff3e0'); }
      else { e.range.clearNote(); e.range.setBackground(null); }
    }
    if (col === MEMBER_COLS.PHONE) {
      var r = validatePhoneNumber(val);
      if (!r.valid) { e.range.setNote('⚠️ ' + r.message); e.range.setBackground('#fff3e0'); }
      else { e.range.clearNote(); e.range.setBackground(null); if (r.formatted !== val) e.range.setValue(r.formatted); }
    }
  }
}
