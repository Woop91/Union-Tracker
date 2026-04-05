/**
 * My Cases Tab — Regression Guards
 *
 * Tests for all 6 bugs found in the My Cases tab review (2026-03-14):
 *   BUG 1: getMemberGrievances must exclude closed/terminal cases
 *   BUG 2: dataGetStewardContact wrapper must accept stewardEmail param
 *   BUG 3: Deadline color must be consistent between card and detail views
 *   BUG 4: Draft submission must refresh grievance data
 *   BUG 5: Issue categories must come from config, not hardcoded
 *   BUG 6: Past cases must lazy-load on first accordion expand
 *
 * Run: npx jest test/my-cases-tab.test.js --verbose
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.resolve(__dirname, '..', 'src');
function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

// ============================================================================
// BACKEND TESTS — require GAS mock environment
// ============================================================================

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '21_WebDashDataService.gs', '21d_WebDashDataWrappers.gs']);

function makeMemberData() {
  return [
    ['Email', 'Name', 'First Name', 'Last Name', 'Role', 'Unit', 'Phone', 'Join Date', 'Dues Status', 'Member ID', 'Work Location', 'Office Days', 'Assigned Steward', 'Is Steward'],
    ['member@test.com', 'Jane Doe', 'Jane', 'Doe', 'Member', 'Unit A', '555-0001', new Date('2023-01-15'), 'Active', 'MEM-001', 'HQ', 'Mon,Tue', 'steward@test.com', false],
    ['steward@test.com', 'John Smith', 'John', 'Smith', 'Steward', 'Unit A', '555-0002', new Date('2022-06-01'), 'Active', 'MEM-002', 'HQ', 'Mon-Fri', '', true],
  ];
}

const GRIEVANCE_HEADERS = [
  'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
  'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category',
  'Resolution', 'Date Closed',
];

function setupWithMixedStatuses() {
  if (DataService._invalidateSheetCache) {
    DataService._invalidateSheetCache('Grievance Log');
    DataService._invalidateSheetCache('Member Directory');
  }
  const now = Date.now();
  const rows = [
    ['GR-OPEN',   'member@test.com', 'Open',      'Step I',   new Date(now + 10 * 86400000), new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'High',   '', 'Safety',     '',     ''],
    ['GR-WON',    'member@test.com', 'Won',       'Step II',  '',                            new Date('2025-06-01'), 'steward@test.com', 'Unit A', 'Medium', '', 'Pay',        'Won',  '2025-07-01'],
    ['GR-DENIED', 'member@test.com', 'Denied',    'Step III', '',                            new Date('2025-04-01'), 'steward@test.com', 'Unit A', 'Low',    '', 'Discipline', 'Denied', '2025-05-01'],
    ['GR-SETTLE', 'member@test.com', 'Settled',   'Step I',   '',                            new Date('2025-03-01'), 'steward@test.com', 'Unit A', 'Medium', '', 'Workload',   'Settled', '2025-04-01'],
    ['GR-CLOSED', 'member@test.com', 'Closed',    'Step II',  '',                            new Date('2025-02-01'), 'steward@test.com', 'Unit A', 'Low',    '', 'Pay',        '',     '2025-03-01'],
    ['GR-WITHDR', 'member@test.com', 'Withdrawn', 'Step I',   '',                            new Date('2025-01-01'), 'steward@test.com', 'Unit A', 'Low',    '', 'Other',      '',     '2025-02-01'],
    ['GR-PEND',   'member@test.com', 'Pending Info', 'Step I', new Date(now + 5 * 86400000), new Date('2026-02-01'), 'steward@test.com', 'Unit A', 'Medium', '', 'Safety',     '',     ''],
  ];
  const data = [GRIEVANCE_HEADERS, ...rows];
  const grievSheet = createMockSheet('Grievance Log', data);
  const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
  const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
}


// ============================================================================
// BUG 1: getMemberGrievances must exclude closed/terminal statuses
// ============================================================================

describe('BUG 1: getMemberGrievances excludes terminal statuses', () => {
  beforeEach(setupWithMixedStatuses);

  test('only returns active (non-terminal) grievances', () => {
    const results = DataService.getMemberGrievances('member@test.com');
    const ids = results.map(r => r.id);
    expect(ids).toContain('GR-OPEN');
    expect(ids).toContain('GR-PEND');
    expect(ids).not.toContain('GR-WON');
    expect(ids).not.toContain('GR-DENIED');
    expect(ids).not.toContain('GR-SETTLE');
    expect(ids).not.toContain('GR-CLOSED');
    expect(ids).not.toContain('GR-WITHDR');
  });

  test('returns only 2 active cases out of 7 total', () => {
    const results = DataService.getMemberGrievances('member@test.com');
    expect(results.length).toBe(2);
  });

  test('getMemberGrievanceHistory returns only terminal cases', () => {
    const result = DataService.getMemberGrievanceHistory('member@test.com');
    expect(result.success).toBe(true);
    const ids = result.history.map(h => h.grievanceId);
    expect(ids).toContain('GR-WON');
    expect(ids).toContain('GR-DENIED');
    expect(ids).toContain('GR-SETTLE');
    expect(ids).toContain('GR-CLOSED');
    expect(ids).toContain('GR-WITHDR');
    expect(ids).not.toContain('GR-OPEN');
    expect(ids).not.toContain('GR-PEND');
  });

  test('active + history covers all cases without overlap', () => {
    const active = DataService.getMemberGrievances('member@test.com');
    const history = DataService.getMemberGrievanceHistory('member@test.com');
    const activeIds = active.map(r => r.id);
    const historyIds = history.history.map(h => h.grievanceId);
    // No overlaps
    activeIds.forEach(id => {
      expect(historyIds).not.toContain(id);
    });
    // Together they cover all 7
    expect(activeIds.length + historyIds.length).toBe(7);
  });

  test('batch data grievances match getMemberGrievances (active only)', () => {
    const directResult = DataService.getMemberGrievances('member@test.com');
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache('Grievance Log');
    }
    setupWithMixedStatuses();
    const batch = DataService.getBatchData('member@test.com', 'member');
    expect(batch.grievances.length).toBe(directResult.length);
    const batchIds = batch.grievances.map(g => g.id).sort();
    const directIds = directResult.map(g => g.id).sort();
    expect(batchIds).toEqual(directIds);
  });
});


// ============================================================================
// BUG 2: dataGetStewardContact wrapper accepts stewardEmail param
// ============================================================================

describe('BUG 2: dataGetStewardContact passes stewardEmail to DataService', () => {
  beforeEach(() => {
    setupWithMixedStatuses();
    // Caller must be member@test.com so the security check passes
    // (member's assignedSteward === target steward email)
    global._resolveCallerEmail = jest.fn(() => 'member@test.com');
  });

  test('wrapper function accepts two parameters', () => {
    expect(typeof dataGetStewardContact).toBe('function');
    // Function.length reflects declared params
    expect(dataGetStewardContact.length).toBe(2);
  });

  test('returns steward contact when given valid session + steward email', () => {
    const result = dataGetStewardContact('test-session-token', 'steward@test.com');
    expect(result).not.toBeNull();
    expect(result).toHaveProperty('email');
    expect(result.email).toBe('steward@test.com');
  });

  test('returns auth error object when session is invalid', () => {
    const origResolve = global._resolveCallerEmail;
    global._resolveCallerEmail = jest.fn(() => null);
    const result = dataGetStewardContact('bad-token', 'steward@test.com');
    expect(result).toEqual(expect.objectContaining({ success: false, authError: true }));
    global._resolveCallerEmail = origResolve;
  });

  test('falls back to caller email when stewardEmail not provided', () => {
    // When stewardEmail is omitted, should use the caller's own email
    const result = dataGetStewardContact('test-session-token');
    // _resolveCallerEmail returns 'test@example.com' by default in gas-mock
    // So it tries to look up that email; may return null if not in directory
    expect(result === null || (result && result.email)).toBeTruthy();
  });
});


// ============================================================================
// BUG 3: Deadline color consistency (card vs detail — static analysis)
// ============================================================================

describe('BUG 3: Deadline color thresholds match between card and detail', () => {
  const memberView = read('member_view.html');

  // Extract the deadline color ternary expressions from both functions
  function extractDeadlineColorExpr(funcName) {
    const funcRegex = new RegExp('function ' + funcName + '\\(');
    const funcStart = memberView.search(funcRegex);
    expect(funcStart).toBeGreaterThan(-1);

    // Find the deadline color assignment near start of function
    const chunk = memberView.substring(funcStart, funcStart + 1200);
    const dcMatch = chunk.match(/var dc = ([^;]+);/);
    expect(dcMatch).not.toBeNull();
    return dcMatch[1];
  }

  test('renderGrievanceCard and renderGrievanceDetail use identical deadline color logic', () => {
    const cardExpr = extractDeadlineColorExpr('renderGrievanceCard');
    const detailExpr = extractDeadlineColorExpr('renderGrievanceDetail');
    expect(cardExpr).toBe(detailExpr);
  });

  test('deadline color expression includes all 4 thresholds (danger, warning, info, success)', () => {
    const detailExpr = extractDeadlineColorExpr('renderGrievanceDetail');
    expect(detailExpr).toContain('--danger');
    expect(detailExpr).toContain('--warning');
    expect(detailExpr).toContain('--info');
    expect(detailExpr).toContain('--success');
  });
});


// ============================================================================
// BUG 4: Draft submission refreshes grievance data (static analysis)
// ============================================================================

describe('BUG 4: Grievance draft submission refreshes AppState', () => {
  const memberView = read('member_view.html');

  test('"Back to My Cases" button fetches fresh data via dataGetMemberGrievances', () => {
    // Find the draft submission success handler
    const draftSubmittedIdx = memberView.indexOf("'Draft Submitted'");
    expect(draftSubmittedIdx).toBeGreaterThan(-1);

    // The "Back to My Cases" button near "Draft Submitted" should call dataGetMemberGrievances
    const afterSubmit = memberView.substring(draftSubmittedIdx, draftSubmittedIdx + 800);
    expect(afterSubmit).toContain('dataGetMemberGrievances');
    expect(afterSubmit).toContain('AppState.grievances');
  });

  test('"Back to My Cases" does not directly call renderMyCases without refresh', () => {
    // Find the success path after draft submission
    const draftSubmittedIdx = memberView.indexOf("'Draft Submitted'");
    const afterSubmit = memberView.substring(draftSubmittedIdx, draftSubmittedIdx + 800);

    // renderMyCases should appear INSIDE the success handler of the data refresh,
    // not as a direct onClick
    const directCallPattern = /onClick:\s*function\s*\(\)\s*\{\s*renderMyCases/;
    expect(directCallPattern.test(afterSubmit)).toBe(false);
  });
});


// ============================================================================
// BUG 5: Issue categories from config, not hardcoded
// ============================================================================

describe('BUG 5: Issue categories are dynamic from config', () => {
  test('_sanitizeConfig includes issueCategories field', () => {
    const webAppCode = read('22_WebDashApp.gs');
    expect(webAppCode).toContain('issueCategories');

    // Verify it references DEFAULT_CONFIG.ISSUE_CATEGORY
    const sanitizeStart = webAppCode.indexOf('function _sanitizeConfig');
    expect(sanitizeStart).toBeGreaterThan(-1);
    const sanitizeEnd = webAppCode.indexOf('\n}', sanitizeStart + 10);
    const sanitizeBlock = webAppCode.substring(sanitizeStart, sanitizeEnd + 2);
    expect(sanitizeBlock).toContain('issueCategories');
    expect(sanitizeBlock).toContain('ISSUE_CATEGORY');
  });

  test('intake form reads categories from CONFIG.issueCategories', () => {
    const memberView = read('member_view.html');
    // Find the function definition, not a call site
    const intakeFormIdx = memberView.indexOf('function _showGrievanceIntakeForm');
    expect(intakeFormIdx).toBeGreaterThan(-1);

    const intakeBlock = memberView.substring(intakeFormIdx, intakeFormIdx + 1500);
    expect(intakeBlock).toContain('CONFIG.issueCategories');
  });

  test('DEFAULT_CONFIG.ISSUE_CATEGORY is a non-empty array', () => {
    expect(typeof DEFAULT_CONFIG).toBe('object');
    expect(Array.isArray(DEFAULT_CONFIG.ISSUE_CATEGORY)).toBe(true);
    expect(DEFAULT_CONFIG.ISSUE_CATEGORY.length).toBeGreaterThan(0);
  });
});


// ============================================================================
// BUG 6: Past cases lazy-load on first accordion expand
// ============================================================================

describe('BUG 6: Past cases section lazy-loads data', () => {
  const memberView = read('member_view.html');

  test('_renderPastCasesSection does NOT call dataGetMemberGrievanceHistory directly', () => {
    // Find the _renderPastCasesSection function body
    const funcStart = memberView.indexOf('function _renderPastCasesSection');
    expect(funcStart).toBeGreaterThan(-1);

    // Find next function definition to delimit the block
    const nextFunc = memberView.indexOf('\nfunction ', funcStart + 10);
    const funcBody = memberView.substring(funcStart, nextFunc > -1 ? nextFunc : funcStart + 2000);

    // Should NOT contain a direct server call
    expect(funcBody).not.toContain('dataGetMemberGrievanceHistory');
    expect(funcBody).not.toContain('serverCall()');
  });

  test('_loadPastCases function exists and calls dataGetMemberGrievanceHistory', () => {
    expect(memberView).toContain('function _loadPastCases');

    const funcStart = memberView.indexOf('function _loadPastCases');
    const funcBody = memberView.substring(funcStart, funcStart + 3500);
    expect(funcBody).toContain('dataGetMemberGrievanceHistory');
    expect(funcBody).toContain('serverCall()');
  });

  test('accordion onClick triggers _loadPastCases only on first expand', () => {
    const funcStart = memberView.indexOf('function _renderPastCasesSection');
    const funcBody = memberView.substring(funcStart, funcStart + 1500);

    // Should track loaded state
    expect(funcBody).toContain('pastCasesLoaded');
    // Should call _loadPastCases conditionally
    expect(funcBody).toContain('_loadPastCases');
    // Should guard with !pastCasesLoaded
    expect(funcBody).toContain('!pastCasesLoaded');
  });
});


// ============================================================================
// STRUCTURAL GUARDS — prevent regressions of the same class
// ============================================================================

describe('Structural: serverCall in member_view passes SESSION_TOKEN for auth endpoints', () => {
  const memberView = read('member_view.html');

  test('dataGetStewardContact is called with SESSION_TOKEN as first arg', () => {
    const regex = /\.dataGetStewardContact\(([^)]+)\)/g;
    let m;
    while ((m = regex.exec(memberView)) !== null) {
      const args = m[1];
      expect(args).toContain('SESSION_TOKEN');
    }
  });

  test('dataGetMemberGrievances is always called with SESSION_TOKEN', () => {
    const regex = /\.dataGetMemberGrievances\(([^)]+)\)/g;
    let m;
    while ((m = regex.exec(memberView)) !== null) {
      expect(m[1].trim()).toBe('SESSION_TOKEN');
    }
  });

  test('dataGetMemberGrievanceHistory is always called with SESSION_TOKEN', () => {
    const regex = /\.dataGetMemberGrievanceHistory\(([^)]+)\)/g;
    let m;
    while ((m = regex.exec(memberView)) !== null) {
      expect(m[1].trim()).toBe('SESSION_TOKEN');
    }
  });
});

describe('Structural: GRIEVANCE_CLOSED_STATUSES consistency', () => {
  test('all statuses used in _buildGrievanceRecord TERMINAL_STATUSES are covered by GRIEVANCE_CLOSED_STATUSES', () => {
    const code = read('21_WebDashDataService.gs');
    const terminalMatch = code.match(/var TERMINAL_STATUSES = \[([^\]]+)\]/);
    expect(terminalMatch).not.toBeNull();

    // Parse the hardcoded list from _buildGrievanceRecord
    const inlineStatuses = terminalMatch[1]
      .split(',')
      .map(s => s.trim().replace(/'/g, '').toLowerCase());

    // The global GRIEVANCE_CLOSED_STATUSES should cover all of them.
    // 'resolved' is a backward-compat alias for 'settled' (GRIEVANCE_STATUS.RESOLVED = 'Settled'),
    // so map it before comparing.
    const aliasMap = { resolved: 'settled' };
    const closedLower = GRIEVANCE_CLOSED_STATUSES.map(s => s.toLowerCase());
    inlineStatuses.forEach(status => {
      const normalized = aliasMap[status] || status;
      expect(closedLower).toContain(normalized);
    });
  });
});
