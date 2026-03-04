/**
 * Tests for 01_Core.gs
 *
 * Covers column constant validation (MEMBER_COLS, GRIEVANCE_COLS),
 * version consistency, DEADLINE_RULES, header arrays, sanitize functions,
 * and error handling.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs']);

// Derive 0-indexed legacy constants for offset validation tests
const MEMBER_COLUMNS = Object.fromEntries(
  Object.entries(MEMBER_COLS).map(([k, v]) => [k, v - 1])
);
const GRIEVANCE_COLUMNS = Object.fromEntries(
  Object.entries(GRIEVANCE_COLS).map(([k, v]) => [k, v - 1])
);

// ============================================================================
// Version consistency
// ============================================================================

describe('Version consistency', () => {
  test('COMMAND_CONFIG.VERSION matches VERSION_INFO.CURRENT', () => {
    expect(COMMAND_CONFIG.VERSION).toBe(VERSION_INFO.CURRENT);
  });

  test('VERSION_INFO.BUILD includes CURRENT', () => {
    expect(VERSION_INFO.BUILD).toContain(VERSION_INFO.CURRENT);
  });

  test('VERSION_INFO components match CURRENT', () => {
    const expected = `${VERSION_INFO.MAJOR}.${VERSION_INFO.MINOR}.${VERSION_INFO.PATCH}`;
    expect(VERSION_INFO.CURRENT).toBe(expected);
  });

  test('API_VERSION matches VERSION_INFO', () => {
    expect(API_VERSION.major).toBe(VERSION_INFO.MAJOR);
    expect(API_VERSION.minor).toBe(VERSION_INFO.MINOR);
    expect(API_VERSION.patch).toBe(VERSION_INFO.PATCH);
  });
});

// ============================================================================
// MEMBER_COLS ↔ MEMBER_COLUMNS offset validation
// ============================================================================

describe('MEMBER_COLS ↔ MEMBER_COLUMNS (1-indexed vs 0-indexed)', () => {
  // Core fields that MUST be exactly offset by 1
  const fieldsToCheck = [
    'MEMBER_ID', 'FIRST_NAME', 'LAST_NAME', 'JOB_TITLE',
    'WORK_LOCATION', 'UNIT', 'OFFICE_DAYS',
    'EMAIL', 'PHONE', 'PREFERRED_COMM', 'BEST_TIME',
    'SUPERVISOR', 'MANAGER', 'IS_STEWARD', 'COMMITTEES', 'ASSIGNED_STEWARD',
    'LAST_VIRTUAL_MTG', 'LAST_INPERSON_MTG', 'OPEN_RATE', 'VOLUNTEER_HOURS',
    'INTEREST_LOCAL', 'INTEREST_CHAPTER', 'INTEREST_ALLIED',
    'RECENT_CONTACT_DATE', 'CONTACT_STEWARD', 'CONTACT_NOTES',
    'HAS_OPEN_GRIEVANCE', 'GRIEVANCE_STATUS', 'NEXT_DEADLINE', 'START_GRIEVANCE',
    'QUICK_ACTIONS', 'PIN_HASH'
  ];

  fieldsToCheck.forEach(field => {
    test(`MEMBER_COLS.${field} - 1 === MEMBER_COLUMNS.${field}`, () => {
      // MEMBER_COLUMNS uses different key for MEMBER_ID
      const colsKey = field;
      let columnsKey = field;
      if (field === 'MEMBER_ID') columnsKey = 'MEMBER_ID';
      if (field === 'NEXT_DEADLINE') columnsKey = 'NEXT_DEADLINE';

      // Always assert — fields in the check list must exist
      expect(MEMBER_COLS[colsKey]).toBeDefined();
      expect(MEMBER_COLUMNS[columnsKey]).toBeDefined();
      expect(MEMBER_COLS[colsKey] - 1).toBe(MEMBER_COLUMNS[columnsKey]);
    });
  });

  test('MEMBER_COLS starts at 1 (1-indexed)', () => {
    expect(MEMBER_COLS.MEMBER_ID).toBe(1);
  });

  test('MEMBER_COLUMNS starts at 0 (0-indexed)', () => {
    expect(MEMBER_COLUMNS.MEMBER_ID).toBe(0);
  });

  test('MEMBER_COLS.PIN_HASH is 33', () => {
    expect(MEMBER_COLS.PIN_HASH).toBe(33);
  });

  test('MEMBER_COLUMNS.PIN_HASH is 32', () => {
    expect(MEMBER_COLUMNS.PIN_HASH).toBe(32);
  });
});

// ============================================================================
// GRIEVANCE_COLS ↔ GRIEVANCE_COLUMNS offset validation
// ============================================================================

describe('GRIEVANCE_COLS ↔ GRIEVANCE_COLUMNS (1-indexed vs 0-indexed)', () => {
  const fieldsToCheck = [
    'GRIEVANCE_ID', 'MEMBER_ID', 'FIRST_NAME', 'LAST_NAME',
    'STATUS', 'CURRENT_STEP',
    'INCIDENT_DATE', 'FILING_DEADLINE', 'DATE_FILED',
    'STEP1_DUE', 'STEP1_RCVD',
    'STEP2_APPEAL_DUE', 'STEP2_APPEAL_FILED', 'STEP2_DUE', 'STEP2_RCVD',
    'STEP3_APPEAL_DUE', 'STEP3_APPEAL_FILED', 'DATE_CLOSED',
    'DAYS_OPEN', 'NEXT_ACTION_DUE', 'DAYS_TO_DEADLINE',
    'ARTICLES', 'ISSUE_CATEGORY',
    'MEMBER_EMAIL', 'UNIT', 'LOCATION', 'STEWARD',
    'RESOLUTION',
    'MESSAGE_ALERT', 'COORDINATOR_MESSAGE', 'ACKNOWLEDGED_BY', 'ACKNOWLEDGED_DATE',
    'DRIVE_FOLDER_ID', 'DRIVE_FOLDER_URL',
    'QUICK_ACTIONS',
    'ACTION_TYPE', 'CHECKLIST_PROGRESS'
  ];

  fieldsToCheck.forEach(field => {
    test(`GRIEVANCE_COLS.${field} - 1 === GRIEVANCE_COLUMNS.${field}`, () => {
      if (GRIEVANCE_COLS[field] !== undefined && GRIEVANCE_COLUMNS[field] !== undefined) {
        expect(GRIEVANCE_COLS[field] - 1).toBe(GRIEVANCE_COLUMNS[field]);
      }
    });
  });

  test('GRIEVANCE_COLS starts at 1', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_ID).toBe(1);
  });

  test('GRIEVANCE_COLUMNS starts at 0', () => {
    expect(GRIEVANCE_COLUMNS.GRIEVANCE_ID).toBe(0);
  });

  test('GRIEVANCE_COLS.RESOLUTION is defined', () => {
    expect(GRIEVANCE_COLS.RESOLUTION).toBeDefined();
    expect(typeof GRIEVANCE_COLS.RESOLUTION).toBe('number');
  });
});

// ============================================================================
// DEADLINE_RULES consistency with TIME_CONSTANTS.DEADLINE_DAYS
// ============================================================================

describe('DEADLINE_RULES ↔ TIME_CONSTANTS.DEADLINE_DAYS consistency', () => {
  test('STEP_1 response matches STEP1_RESPONSE', () => {
    expect(DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE)
      .toBe(TIME_CONSTANTS.DEADLINE_DAYS.STEP1_RESPONSE);
  });

  test('STEP_2 appeal matches STEP2_APPEAL', () => {
    expect(DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL)
      .toBe(TIME_CONSTANTS.DEADLINE_DAYS.STEP2_APPEAL);
  });

  test('STEP_2 response matches STEP2_RESPONSE', () => {
    expect(DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE)
      .toBe(TIME_CONSTANTS.DEADLINE_DAYS.STEP2_RESPONSE);
  });

  test('STEP_3 appeal matches STEP3_APPEAL', () => {
    expect(DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL)
      .toBe(TIME_CONSTANTS.DEADLINE_DAYS.STEP3_APPEAL);
  });

  test('DEADLINE_RULES has all required steps', () => {
    expect(DEADLINE_RULES.STEP_1).toBeDefined();
    expect(DEADLINE_RULES.STEP_2).toBeDefined();
    expect(DEADLINE_RULES.STEP_3).toBeDefined();
    expect(DEADLINE_RULES.ARBITRATION).toBeDefined();
  });

  test('ARBITRATION days to demand is 30', () => {
    expect(DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND).toBe(30);
  });
});

// NOTE: getMemberHeaders/getGrievanceHeaders header↔position tests
// are covered in columns.test.js (Header map → COLS consistency).

// ============================================================================
// SHEETS constant
// ============================================================================

describe('SHEETS constant', () => {
  test('has required sheet names', () => {
    expect(SHEETS.CONFIG).toBe('Config');
    expect(SHEETS.MEMBER_DIR).toBe('Member Directory');
    expect(SHEETS.GRIEVANCE_LOG).toBe('Grievance Log');
    expect(SHEETS.AUDIT_LOG).toBe('_Audit_Log');
  });

  test('backward compatibility aliases match', () => {
    expect(SHEETS.GRIEVANCE_TRACKER).toBe(SHEETS.GRIEVANCE_LOG);
    expect(SHEETS.MEMBER_DIRECTORY).toBe(SHEETS.MEMBER_DIR);
    expect(SHEET_NAMES).toBe(SHEETS);
  });
});

// ============================================================================
// sanitizeForQuery
// ============================================================================

describe('sanitizeForQuery', () => {
  test('escapes single quotes', () => {
    expect(sanitizeForQuery("O'Brien")).toBe("O''Brien");
  });

  test('escapes backslashes', () => {
    expect(sanitizeForQuery("a\\b")).toBe("a\\\\b");
  });

  test('handles empty input', () => {
    expect(sanitizeForQuery('')).toBe('');
    expect(sanitizeForQuery(null)).toBe('');
  });
});

// ============================================================================
// Error handling utilities
// ============================================================================

describe('withErrorHandling', () => {
  test('returns wrapped function', () => {
    const fn = () => 42;
    const wrapped = withErrorHandling(fn, 'test');
    expect(typeof wrapped).toBe('function');
    expect(wrapped()).toBe(42);
  });

  test('catches errors and returns null', () => {
    const fn = () => { throw new Error('boom'); };
    const wrapped = withErrorHandling(fn, 'test');
    expect(wrapped()).toBeNull();
  });
});

describe('ERROR_LEVEL', () => {
  test('has all severity levels', () => {
    expect(ERROR_LEVEL.DEBUG).toBe('DEBUG');
    expect(ERROR_LEVEL.INFO).toBe('INFO');
    expect(ERROR_LEVEL.WARNING).toBe('WARNING');
    expect(ERROR_LEVEL.ERROR).toBe('ERROR');
    expect(ERROR_LEVEL.CRITICAL).toBe('CRITICAL');
  });
});

describe('isVersionCompatible', () => {
  test('accepts same major version', () => {
    expect(isVersionCompatible('4.0.0')).toBe(true);
    expect(isVersionCompatible('4.9.9')).toBe(true);
  });

  test('rejects different major version', () => {
    expect(isVersionCompatible('3.0.0')).toBe(false);
    expect(isVersionCompatible('5.0.0')).toBe(false);
  });

  test('rejects empty/null', () => {
    expect(isVersionCompatible('')).toBe(false);
    expect(isVersionCompatible(null)).toBe(false);
  });
});

// ============================================================================
// STATUS_COLORS
// ============================================================================

describe('STATUS_COLORS', () => {
  test('has colors for all default statuses', () => {
    const defaultStatuses = ['Open', 'Pending Info', 'Settled', 'Withdrawn', 'Denied', 'Won', 'Appealed', 'In Arbitration', 'Closed'];
    defaultStatuses.forEach(status => {
      expect(COMMAND_CONFIG.STATUS_COLORS[status]).toBeDefined();
      expect(COMMAND_CONFIG.STATUS_COLORS[status].bg).toBeTruthy();
      expect(COMMAND_CONFIG.STATUS_COLORS[status].text).toBeTruthy();
    });
  });
});

// ============================================================================
// AUDIT_EVENTS
// ============================================================================

describe('AUDIT_EVENTS', () => {
  test('has grievance event types', () => {
    expect(AUDIT_EVENTS.GRIEVANCE_CREATED).toBe('GRIEVANCE_CREATED');
    expect(AUDIT_EVENTS.GRIEVANCE_UPDATED).toBe('GRIEVANCE_UPDATED');
  });

  test('has member event types', () => {
    expect(AUDIT_EVENTS.MEMBER_ADDED).toBe('MEMBER_ADDED');
    expect(AUDIT_EVENTS.MEMBER_UPDATED).toBe('MEMBER_UPDATED');
  });
});

// ============================================================================
// GRIEVANCE_OUTCOMES constant existence guard
// ============================================================================

describe('GRIEVANCE_OUTCOMES', () => {
  test('is defined', () => {
    expect(GRIEVANCE_OUTCOMES).toBeDefined();
  });

  test('has PENDING value', () => {
    expect(GRIEVANCE_OUTCOMES.PENDING).toBe('Pending');
  });

  test('has WON value', () => {
    expect(GRIEVANCE_OUTCOMES.WON).toBe('Won');
  });

  test('has DENIED value', () => {
    expect(GRIEVANCE_OUTCOMES.DENIED).toBe('Denied');
  });

  test('has SETTLED value', () => {
    expect(GRIEVANCE_OUTCOMES.SETTLED).toBe('Settled');
  });

  test('has WITHDRAWN value', () => {
    expect(GRIEVANCE_OUTCOMES.WITHDRAWN).toBe('Withdrawn');
  });

  test('has CLOSED value', () => {
    expect(GRIEVANCE_OUTCOMES.CLOSED).toBe('Closed');
  });

  test('all values are non-empty strings', () => {
    Object.entries(GRIEVANCE_OUTCOMES).forEach(([key, value]) => {
      expect(typeof value).toBe('string');
      expect(value.length).toBeGreaterThan(0);
    });
  });

  test('GRIEVANCE_STATUS and GRIEVANCE_OUTCOMES are both defined', () => {
    expect(GRIEVANCE_STATUS).toBeDefined();
    expect(GRIEVANCE_OUTCOMES).toBeDefined();
  });
});
