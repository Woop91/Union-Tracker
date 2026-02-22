/**
 * Tests for the Dynamic Column System
 *
 * Validates the auto-discovery column infrastructure:
 * - buildColsFromMap_() correctly generates 1-indexed column constants
 * - getHeadersFromMap_() correctly extracts header text arrays
 * - buildLegacyCols_() correctly generates 0-indexed compat objects
 * - getColumnLetter() correctly converts column numbers to letters
 * - All header maps produce valid, contiguous column constants
 * - SATISFACTION_COLS and SATISFACTION_SECTIONS are cross-consistent
 * - No hardcoded column indices remain in source files (regression guard)
 */

const fs = require('fs');
const path = require('path');

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load core constants and utilities
loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs']);

// ============================================================================
// buildColsFromMap_
// ============================================================================

describe('buildColsFromMap_', () => {
  test('generates 1-indexed column constants from header map', () => {
    const map = [
      { key: 'A', header: 'Alpha' },
      { key: 'B', header: 'Bravo' },
      { key: 'C', header: 'Charlie' }
    ];
    const cols = buildColsFromMap_(map);
    expect(cols.A).toBe(1);
    expect(cols.B).toBe(2);
    expect(cols.C).toBe(3);
  });

  test('respects explicit pos override', () => {
    const map = [
      { key: 'A', header: 'Alpha' },
      { key: 'B', header: 'Bravo', pos: 10 },
      { key: 'C', header: 'Charlie' }
    ];
    const cols = buildColsFromMap_(map);
    expect(cols.A).toBe(1);
    expect(cols.B).toBe(10);
    expect(cols.C).toBe(3);
  });

  test('applies aliases correctly', () => {
    const map = [
      { key: 'FIRST_NAME', header: 'First Name' },
      { key: 'LAST_NAME', header: 'Last Name' }
    ];
    const cols = buildColsFromMap_(map, { NAME: 'FIRST_NAME' });
    expect(cols.NAME).toBe(cols.FIRST_NAME);
    expect(cols.NAME).toBe(1);
  });

  test('returns empty object for empty map', () => {
    const cols = buildColsFromMap_([]);
    expect(Object.keys(cols)).toHaveLength(0);
  });

  test('handles map without aliases param', () => {
    const map = [{ key: 'X', header: 'X Label' }];
    const cols = buildColsFromMap_(map);
    expect(cols.X).toBe(1);
  });
});

// ============================================================================
// getHeadersFromMap_
// ============================================================================

describe('getHeadersFromMap_', () => {
  test('extracts header text in order', () => {
    const map = [
      { key: 'A', header: 'Alpha' },
      { key: 'B', header: 'Bravo' },
      { key: 'C', header: 'Charlie' }
    ];
    expect(getHeadersFromMap_(map)).toEqual(['Alpha', 'Bravo', 'Charlie']);
  });

  test('returns empty array for empty map', () => {
    expect(getHeadersFromMap_([])).toEqual([]);
  });

  test('preserves exact header strings including special characters', () => {
    const map = [
      { key: 'A', header: 'Yes/No (Dropdowns)' },
      { key: 'B', header: 'Step 1 Due Date' }
    ];
    expect(getHeadersFromMap_(map)).toEqual(['Yes/No (Dropdowns)', 'Step 1 Due Date']);
  });
});

// ============================================================================
// buildLegacyCols_
// ============================================================================

describe('buildLegacyCols_', () => {
  test('subtracts 1 from all values to create 0-indexed map', () => {
    const primary = { A: 1, B: 2, C: 3 };
    const legacy = buildLegacyCols_(primary);
    expect(legacy.A).toBe(0);
    expect(legacy.B).toBe(1);
    expect(legacy.C).toBe(2);
  });

  test('applies extra aliases pointing to primary keys', () => {
    const primary = { FIRST_NAME: 1, DATE_FILED: 5 };
    const legacy = buildLegacyCols_(primary, { MEMBER_NAME: 'FIRST_NAME', FILING_DATE: 'DATE_FILED' });
    expect(legacy.MEMBER_NAME).toBe(0); // FIRST_NAME(1) - 1
    expect(legacy.FILING_DATE).toBe(4); // DATE_FILED(5) - 1
  });

  test('includes all original keys plus aliases', () => {
    const primary = { A: 1, B: 2 };
    const legacy = buildLegacyCols_(primary, { C: 'A' });
    expect(Object.keys(legacy)).toContain('A');
    expect(Object.keys(legacy)).toContain('B');
    expect(Object.keys(legacy)).toContain('C');
  });

  test('handles no extra aliases', () => {
    const primary = { X: 10 };
    const legacy = buildLegacyCols_(primary);
    expect(legacy.X).toBe(9);
  });
});

// ============================================================================
// getColumnLetter
// ============================================================================

describe('getColumnLetter', () => {
  test('converts single-letter columns correctly', () => {
    expect(getColumnLetter(1)).toBe('A');
    expect(getColumnLetter(26)).toBe('Z');
  });

  test('converts double-letter columns correctly', () => {
    expect(getColumnLetter(27)).toBe('AA');
    expect(getColumnLetter(28)).toBe('AB');
    expect(getColumnLetter(52)).toBe('AZ');
    expect(getColumnLetter(53)).toBe('BA');
  });

  test('converts triple-letter columns correctly', () => {
    expect(getColumnLetter(703)).toBe('AAA');
  });

  test('handles column numbers used in satisfaction survey', () => {
    // BT = column 72 (SATISFACTION_COLS.SUMMARY_START)
    expect(getColumnLetter(72)).toBe('BT');
    // BK = column 63 (Q62_CONCERNS_SERIOUS)
    expect(getColumnLetter(63)).toBe('BK');
  });
});

// ============================================================================
// Header map → COLS consistency for ALL header maps
// ============================================================================

describe('Header map to COLS consistency', () => {
  const headerMapTests = [
    { name: 'MEMBER', map: 'MEMBER_HEADER_MAP_', cols: 'MEMBER_COLS' },
    { name: 'GRIEVANCE', map: 'GRIEVANCE_HEADER_MAP_', cols: 'GRIEVANCE_COLS' },
    { name: 'CONFIG', map: 'CONFIG_HEADER_MAP_', cols: 'CONFIG_COLS' },
    { name: 'MEETING_CHECKIN', map: 'MEETING_CHECKIN_HEADER_MAP_', cols: 'MEETING_CHECKIN_COLS' },
    { name: 'STEWARD_PERF', map: 'STEWARD_PERF_HEADER_MAP_', cols: 'STEWARD_PERF_COLS' },
    { name: 'AUDIT_LOG', map: 'AUDIT_LOG_HEADER_MAP_', cols: 'AUDIT_LOG_COLS' },
    { name: 'EVENT_AUDIT', map: 'EVENT_AUDIT_HEADER_MAP_', cols: 'EVENT_AUDIT_COLS' },
    { name: 'SURVEY_VAULT', map: 'SURVEY_VAULT_HEADER_MAP_', cols: 'SURVEY_VAULT_COLS' },
    { name: 'SURVEY_TRACKING', map: 'SURVEY_TRACKING_HEADER_MAP_', cols: 'SURVEY_TRACKING_COLS' },
    { name: 'FEEDBACK', map: 'FEEDBACK_HEADER_MAP_', cols: 'FEEDBACK_COLS' },
    { name: 'CHECKLIST', map: 'CHECKLIST_HEADER_MAP_', cols: 'CHECKLIST_COLS' }
  ];

  headerMapTests.forEach(({ name, map, cols }) => {
    describe(`${name}`, () => {
      test('header map is a non-empty array', () => {
        expect(Array.isArray(global[map])).toBe(true);
        expect(global[map].length).toBeGreaterThan(0);
      });

      test('every entry has key and header properties', () => {
        global[map].forEach((entry, i) => {
          expect(entry.key).toBeTruthy();
          expect(typeof entry.key).toBe('string');
          expect(entry.header).toBeTruthy();
          expect(typeof entry.header).toBe('string');
        });
      });

      test('all keys are unique', () => {
        const keys = global[map].map(e => e.key);
        expect(new Set(keys).size).toBe(keys.length);
      });

      test('COLS object has a key for every header map entry', () => {
        const colsObj = global[cols];
        global[map].forEach((entry) => {
          expect(colsObj[entry.key]).toBeDefined();
          expect(typeof colsObj[entry.key]).toBe('number');
        });
      });

      test('COLS values match header map positions (1-indexed)', () => {
        const colsObj = global[cols];
        global[map].forEach((entry, i) => {
          const expected = entry.pos !== undefined ? entry.pos : (i + 1);
          expect(colsObj[entry.key]).toBe(expected);
        });
      });

      test('getHeadersFromMap_ output length matches header map length', () => {
        const headers = getHeadersFromMap_(global[map]);
        expect(headers.length).toBe(global[map].length);
      });

      test('getHeadersFromMap_ output matches header map headers', () => {
        const headers = getHeadersFromMap_(global[map]);
        global[map].forEach((entry, i) => {
          expect(headers[i]).toBe(entry.header);
        });
      });
    });
  });
});

// ============================================================================
// Legacy compat objects (MEMBER_COLUMNS, GRIEVANCE_COLUMNS)
// ============================================================================

describe('Legacy compat: MEMBER_COLUMNS', () => {
  test('every key in MEMBER_COLS exists in MEMBER_COLUMNS with offset -1', () => {
    Object.keys(MEMBER_COLS).forEach(key => {
      expect(MEMBER_COLUMNS[key]).toBeDefined();
      expect(MEMBER_COLUMNS[key]).toBe(MEMBER_COLS[key] - 1);
    });
  });

  test('legacy aliases exist and map correctly', () => {
    expect(MEMBER_COLUMNS.ID).toBe(MEMBER_COLS.MEMBER_ID - 1);
    expect(MEMBER_COLUMNS.JOB_DEPT).toBe(MEMBER_COLS.JOB_TITLE - 1);
    expect(MEMBER_COLUMNS.STATUS).toBe(MEMBER_COLS.GRIEVANCE_STATUS - 1);
    expect(MEMBER_COLUMNS.LAST_UPDATED).toBe(MEMBER_COLS.RECENT_CONTACT_DATE - 1);
  });
});

describe('Legacy compat: GRIEVANCE_COLUMNS', () => {
  test('every key in GRIEVANCE_COLS exists in GRIEVANCE_COLUMNS with offset -1', () => {
    Object.keys(GRIEVANCE_COLS).forEach(key => {
      expect(GRIEVANCE_COLUMNS[key]).toBeDefined();
      expect(GRIEVANCE_COLUMNS[key]).toBe(GRIEVANCE_COLS[key] - 1);
    });
  });

  test('legacy aliases exist and map correctly', () => {
    expect(GRIEVANCE_COLUMNS.MEMBER_NAME).toBe(GRIEVANCE_COLS.FIRST_NAME - 1);
    expect(GRIEVANCE_COLUMNS.FILING_DATE).toBe(GRIEVANCE_COLS.DATE_FILED - 1);
    expect(GRIEVANCE_COLUMNS.ARBITRATION_DATE).toBe(GRIEVANCE_COLS.DATE_CLOSED - 1);
    expect(GRIEVANCE_COLUMNS.GRIEVANCE_TYPE).toBe(GRIEVANCE_COLS.ISSUE_CATEGORY - 1);
  });
});

// ============================================================================
// SATISFACTION_COLS ↔ SATISFACTION_SECTIONS cross-consistency
// ============================================================================

describe('SATISFACTION_COLS ↔ SATISFACTION_SECTIONS cross-consistency', () => {
  test('SATISFACTION_SECTIONS question values are all valid SATISFACTION_COLS column numbers', () => {
    const allColValues = new Set(Object.values(SATISFACTION_COLS));
    Object.entries(SATISFACTION_SECTIONS).forEach(([section, config]) => {
      config.questions.forEach(q => {
        expect(allColValues.has(q)).toBe(true);
      });
    });
  });

  test('OVERALL_SAT references Q6-Q9', () => {
    expect(SATISFACTION_SECTIONS.OVERALL_SAT.questions).toContain(SATISFACTION_COLS.Q6_SATISFIED_REP);
    expect(SATISFACTION_SECTIONS.OVERALL_SAT.questions).toContain(SATISFACTION_COLS.Q7_TRUST_UNION);
    expect(SATISFACTION_SECTIONS.OVERALL_SAT.questions).toContain(SATISFACTION_COLS.Q8_FEEL_PROTECTED);
    expect(SATISFACTION_SECTIONS.OVERALL_SAT.questions).toContain(SATISFACTION_COLS.Q9_RECOMMEND);
  });

  test('STEWARD_3A references Q10-Q16', () => {
    const expected = [
      SATISFACTION_COLS.Q10_TIMELY_RESPONSE, SATISFACTION_COLS.Q11_TREATED_RESPECT,
      SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS, SATISFACTION_COLS.Q13_FOLLOWED_THROUGH,
      SATISFACTION_COLS.Q14_ADVOCATED, SATISFACTION_COLS.Q15_SAFE_CONCERNS,
      SATISFACTION_COLS.Q16_CONFIDENTIALITY
    ];
    expected.forEach(q => {
      expect(SATISFACTION_SECTIONS.STEWARD_3A.questions).toContain(q);
    });
  });

  test('STEWARD_3B references Q18-Q20', () => {
    expect(SATISFACTION_SECTIONS.STEWARD_3B.questions).toContain(SATISFACTION_COLS.Q18_KNOW_CONTACT);
    expect(SATISFACTION_SECTIONS.STEWARD_3B.questions).toContain(SATISFACTION_COLS.Q19_CONFIDENT_HELP);
    expect(SATISFACTION_SECTIONS.STEWARD_3B.questions).toContain(SATISFACTION_COLS.Q20_EASY_FIND);
  });

  test('SCHEDULING references Q56-Q62', () => {
    const expected = [
      SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES, SATISFACTION_COLS.Q57_ADEQUATELY_INFORMED,
      SATISFACTION_COLS.Q58_CLEAR_CRITERIA, SATISFACTION_COLS.Q59_WORK_EXPECTATIONS,
      SATISFACTION_COLS.Q60_EFFECTIVE_OUTCOMES, SATISFACTION_COLS.Q61_SUPPORTS_WELLBEING,
      SATISFACTION_COLS.Q62_CONCERNS_SERIOUS
    ];
    expected.forEach(q => {
      expect(SATISFACTION_SECTIONS.SCHEDULING.questions).toContain(q);
    });
  });

  test('SUMMARY_START column is after all question columns', () => {
    const maxQuestion = Math.max(
      ...Object.values(SATISFACTION_COLS).filter(v => v > 0 && v < SATISFACTION_COLS.SUMMARY_START)
    );
    expect(SATISFACTION_COLS.SUMMARY_START).toBeGreaterThan(maxQuestion);
  });

  test('AVG columns are contiguous from SUMMARY_START', () => {
    const avgKeys = Object.keys(SATISFACTION_COLS).filter(k => k.startsWith('AVG_'));
    const avgValues = avgKeys.map(k => SATISFACTION_COLS[k]).sort((a, b) => a - b);
    expect(avgValues[0]).toBe(SATISFACTION_COLS.SUMMARY_START);
    for (let i = 1; i < avgValues.length; i++) {
      expect(avgValues[i]).toBe(avgValues[i - 1] + 1);
    }
  });

  test('all 11 section averages are present', () => {
    const expectedAvgs = [
      'AVG_OVERALL_SAT', 'AVG_STEWARD_RATING', 'AVG_STEWARD_ACCESS',
      'AVG_CHAPTER', 'AVG_LEADERSHIP', 'AVG_CONTRACT', 'AVG_REPRESENTATION',
      'AVG_COMMUNICATION', 'AVG_MEMBER_VOICE', 'AVG_VALUE_ACTION', 'AVG_SCHEDULING'
    ];
    expectedAvgs.forEach(key => {
      expect(SATISFACTION_COLS[key]).toBeDefined();
      expect(typeof SATISFACTION_COLS[key]).toBe('number');
    });
  });
});

// ============================================================================
// PII_MEMBER_COLS derived from MEMBER_COLS
// ============================================================================

describe('PII_MEMBER_COLS', () => {
  test('all PII columns are valid MEMBER_COLS values', () => {
    const memberColValues = new Set(Object.values(MEMBER_COLS));
    PII_MEMBER_COLS.forEach(col => {
      expect(memberColValues.has(col)).toBe(true);
    });
  });

  test('includes address fields', () => {
    expect(PII_MEMBER_COLS).toContain(MEMBER_COLS.STREET_ADDRESS);
    expect(PII_MEMBER_COLS).toContain(MEMBER_COLS.CITY);
    expect(PII_MEMBER_COLS).toContain(MEMBER_COLS.STATE);
    expect(PII_MEMBER_COLS).toContain(MEMBER_COLS.ZIP_CODE);
  });
});

// ============================================================================
// syncColumnMaps (mocked sheet scenario)
// ============================================================================

describe('syncColumnMaps', () => {
  test('returns result object with warnings and synced arrays', () => {
    // With mock spreadsheet (no actual sheets), should return empty result
    const result = syncColumnMaps();
    expect(result).toHaveProperty('warnings');
    expect(result).toHaveProperty('synced');
    expect(Array.isArray(result.warnings)).toBe(true);
    expect(Array.isArray(result.synced)).toBe(true);
  });
});

// ============================================================================
// resolveColumnsFromSheet_ (mocked)
// ============================================================================

describe('resolveColumnsFromSheet_', () => {
  test('returns null when sheet does not exist', () => {
    const result = resolveColumnsFromSheet_('NonExistentSheet', MEMBER_HEADER_MAP_);
    expect(result).toBeNull();
  });
});

// ============================================================================
// Regression guard: no raw numeric column indices in data access patterns
// ============================================================================

describe('No hardcoded column indices (regression guard)', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');

  // Files that should use *_COLS constants for satisfaction survey access
  const satisfactionFiles = ['09_Dashboards.gs', '12_Features.gs'];

  satisfactionFiles.forEach(file => {
    test(`${file}: calculateSectionAvg_ calls use SATISFACTION_COLS, not raw numbers`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      // Find calculateSectionAvg_ calls with raw numeric arrays like [6, 7, 8, 9]
      const rawArrayCalls = content.match(/calculateSectionAvg_\(\s*row\s*,\s*\[\s*\d+/g);
      expect(rawArrayCalls).toBeNull();
    });
  });

  // Files that should NOT have raw row[N] access to satisfaction survey data
  satisfactionFiles.forEach(file => {
    test(`${file}: satisfaction row access uses SATISFACTION_COLS, not row[N]`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      // Look for row[digit] that isn't row[0] (timestamp is acceptable) and isn't row[i]
      // This catches patterns like: row[1], row[36], row[6]
      // But we need to be careful — some row[0] for timestamp is fine, and loop indices like row[i] are fine
      const lines = content.split('\n');
      const violations = [];
      lines.forEach((line, idx) => {
        // Skip comments
        if (line.trim().startsWith('//') || line.trim().startsWith('*')) return;
        // Match row[literal_number] where number > 0
        const matches = line.match(/\brow\[([1-9]\d*)\]/g);
        if (matches) {
          // Only flag if this is in a satisfaction context (near calculateSectionAvg_ or satisfaction-related functions)
          // We can check if SATISFACTION_COLS is referenced nearby
          const nearbyContext = lines.slice(Math.max(0, idx - 30), idx + 1).join('\n');
          if (nearbyContext.includes('satisfaction') || nearbyContext.includes('Satisfaction') ||
              nearbyContext.includes('satRow') || nearbyContext.includes('satAvg') ||
              nearbyContext.includes('Anon') || nearbyContext.includes('Looker')) {
            violations.push(`Line ${idx + 1}: ${line.trim()}`);
          }
        }
      });
      expect(violations).toEqual([]);
    });
  });

  // Audit log access should use EVENT_AUDIT_COLS or AUDIT_LOG_COLS
  test('06_Maintenance.gs: audit event access uses EVENT_AUDIT_COLS, not row[N]', () => {
    const content = fs.readFileSync(path.join(srcDir, '06_Maintenance.gs'), 'utf8');
    const lines = content.split('\n');
    const violations = [];
    lines.forEach((line, idx) => {
      if (line.trim().startsWith('//') || line.trim().startsWith('*')) return;
      const matches = line.match(/\brow\[([0-9]+)\]/g);
      if (matches) {
        const nearbyContext = lines.slice(Math.max(0, idx - 20), idx + 1).join('\n');
        if (nearbyContext.includes('audit') || nearbyContext.includes('Audit')) {
          violations.push(`Line ${idx + 1}: ${line.trim()}`);
        }
      }
    });
    expect(violations).toEqual([]);
  });

  test('08d_AuditAndFormulas.gs: audit/vault access uses *_COLS, not data[i][N]', () => {
    const content = fs.readFileSync(path.join(srcDir, '08d_AuditAndFormulas.gs'), 'utf8');
    const lines = content.split('\n');
    const violations = [];
    lines.forEach((line, idx) => {
      if (line.trim().startsWith('//') || line.trim().startsWith('*')) return;
      // Match data[i][literal_number] or data[j][literal_number]
      const matches = line.match(/\bdata\[\w+\]\[(\d+)\]/g);
      if (matches) {
        // Only flag if it looks like column access (not array building/headers)
        const nearbyContext = lines.slice(Math.max(0, idx - 10), idx + 1).join('\n');
        if (nearbyContext.includes('audit') || nearbyContext.includes('Audit') ||
            nearbyContext.includes('vault') || nearbyContext.includes('Vault') ||
            nearbyContext.includes('verify') || nearbyContext.includes('hash')) {
          violations.push(`Line ${idx + 1}: ${line.trim()}`);
        }
      }
    });
    expect(violations).toEqual([]);
  });

  // Formula column letters should use getColumnLetter(), not hardcoded letters
  const formulaFiles = ['10a_SheetCreation.gs', '10c_FormHandlers.gs', '12_Features.gs'];
  formulaFiles.forEach(file => {
    test(`${file}: conditional format formulas use getColumnLetter(), not hardcoded $X2 letters`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      // Look for hardcoded column letters in formula strings like '$B2', '$G2', '$AD2'
      // These should be replaced with getColumnLetter() calls
      // Match string literals containing $LETTER2 patterns (column refs in formulas)
      const lines = content.split('\n');
      const violations = [];
      lines.forEach((line, idx) => {
        if (line.trim().startsWith('//') || line.trim().startsWith('*')) return;
        // Look for formula patterns with hardcoded column letters like '$B2', '$AD2'
        // But only in string contexts (inside quotes)
        const formulaMatch = line.match(/['"`]\$[A-Z]{1,3}\d+/g);
        if (formulaMatch) {
          // Exclude lines that use getColumnLetter (they build the formula dynamically)
          if (!line.includes('getColumnLetter')) {
            violations.push(`Line ${idx + 1}: ${line.trim()}`);
          }
        }
      });
      expect(violations).toEqual([]);
    });
  });

  // Column counts should use HEADER_MAP_.length, not hardcoded numbers
  test('08c_FormsAndNotifications.gs: column counts use .length, not hardcoded numbers', () => {
    const content = fs.readFileSync(path.join(srcDir, '08c_FormsAndNotifications.gs'), 'utf8');
    // Look for getRange calls with a literal number that could be a column count for survey tracking
    const lines = content.split('\n');
    const violations = [];
    lines.forEach((line, idx) => {
      if (line.trim().startsWith('//') || line.trim().startsWith('*')) return;
      // Check for getRange with a trailing argument that looks like a hardcoded column count
      const nearbyContext = lines.slice(Math.max(0, idx - 5), idx + 1).join('\n');
      if (nearbyContext.includes('racking') || nearbyContext.includes('tracking')) {
        // Look for patterns like: getRange(x, y, z, 10) where 10 is a literal column count
        const rangeCall = line.match(/getRange\([^)]*,\s*(\d+)\s*\)/);
        if (rangeCall && parseInt(rangeCall[1]) >= 8 && parseInt(rangeCall[1]) <= 15) {
          if (!line.includes('.length') && !line.includes('_COLS')) {
            violations.push(`Line ${idx + 1}: ${line.trim()}`);
          }
        }
      }
    });
    expect(violations).toEqual([]);
  });
});

// ============================================================================
// Column number to letter round-trip consistency
// ============================================================================

describe('getColumnLetter for all *_COLS values', () => {
  // Verify that key column numbers produce the expected letters
  // (as documented in the SATISFACTION_COLS comments)
  const knownMappings = {
    1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G',
    26: 'Z', 27: 'AA', 28: 'AB', 33: 'AG', 37: 'AK', 42: 'AP',
    52: 'AZ', 56: 'BD', 57: 'BE', 63: 'BK', 64: 'BL', 69: 'BQ',
    72: 'BT', 82: 'CD'
  };

  Object.entries(knownMappings).forEach(([num, letter]) => {
    test(`column ${num} → ${letter}`, () => {
      expect(getColumnLetter(parseInt(num))).toBe(letter);
    });
  });
});

// ============================================================================
// Header map lengths match actual column coverage
// ============================================================================

describe('Header map column coverage completeness', () => {
  test('MEMBER_HEADER_MAP_ length matches MEMBER_COLS.ZIP_CODE (last column)', () => {
    expect(MEMBER_HEADER_MAP_.length).toBe(MEMBER_COLS.ZIP_CODE);
  });

  test('GRIEVANCE_HEADER_MAP_ length matches max GRIEVANCE_COLS value', () => {
    const maxCol = Math.max(...Object.values(GRIEVANCE_COLS));
    expect(GRIEVANCE_HEADER_MAP_.length).toBe(maxCol);
  });

  test('AUDIT_LOG_HEADER_MAP_ length matches max AUDIT_LOG_COLS value', () => {
    const maxCol = Math.max(...Object.values(AUDIT_LOG_COLS));
    expect(AUDIT_LOG_HEADER_MAP_.length).toBe(maxCol);
  });

  test('EVENT_AUDIT_HEADER_MAP_ length matches max EVENT_AUDIT_COLS value', () => {
    const maxCol = Math.max(...Object.values(EVENT_AUDIT_COLS));
    expect(EVENT_AUDIT_HEADER_MAP_.length).toBe(maxCol);
  });

  test('CHECKLIST_HEADER_MAP_ length matches max CHECKLIST_COLS value', () => {
    const maxCol = Math.max(...Object.values(CHECKLIST_COLS));
    expect(CHECKLIST_HEADER_MAP_.length).toBe(maxCol);
  });

  test('SURVEY_TRACKING_HEADER_MAP_ length matches max SURVEY_TRACKING_COLS value', () => {
    const maxCol = Math.max(...Object.values(SURVEY_TRACKING_COLS));
    expect(SURVEY_TRACKING_HEADER_MAP_.length).toBe(maxCol);
  });

  test('FEEDBACK_HEADER_MAP_ length matches max FEEDBACK_COLS value', () => {
    const maxCol = Math.max(...Object.values(FEEDBACK_COLS));
    expect(FEEDBACK_HEADER_MAP_.length).toBe(maxCol);
  });
});

// ============================================================================
// Two audit schemas don't conflict
// ============================================================================

describe('Audit log schemas', () => {
  test('AUDIT_LOG_COLS and EVENT_AUDIT_COLS have different column counts', () => {
    // Edit-level audit has more columns than event-level audit
    expect(AUDIT_LOG_HEADER_MAP_.length).toBeGreaterThan(EVENT_AUDIT_HEADER_MAP_.length);
  });

  test('both audit schemas have TIMESTAMP as first column', () => {
    expect(AUDIT_LOG_COLS.TIMESTAMP).toBe(1);
    expect(EVENT_AUDIT_COLS.TIMESTAMP).toBe(1);
  });

  test('EVENT_AUDIT_COLS has integrity hash column', () => {
    expect(EVENT_AUDIT_COLS.INTEGRITY_HASH).toBeDefined();
    expect(typeof EVENT_AUDIT_COLS.INTEGRITY_HASH).toBe('number');
  });

  test('AUDIT_LOG_COLS has record tracking columns', () => {
    expect(AUDIT_LOG_COLS.RECORD_ID).toBeDefined();
    expect(AUDIT_LOG_COLS.ACTION_TYPE).toBeDefined();
    expect(AUDIT_LOG_COLS.OLD_VALUE).toBeDefined();
    expect(AUDIT_LOG_COLS.NEW_VALUE).toBeDefined();
  });
});

// ============================================================================
// Regression: no undefined keys in legacy alias usage
// ============================================================================

describe('Legacy alias coverage (no undefined keys)', () => {
  test('GRIEVANCE_COLUMNS has no UNIT key (grievance sheet has LOCATION instead)', () => {
    expect(GRIEVANCE_COLUMNS.UNIT).toBeUndefined();
    expect(GRIEVANCE_COLUMNS.LOCATION).toBeDefined();
    expect(typeof GRIEVANCE_COLUMNS.LOCATION).toBe('number');
  });

  test('GRIEVANCE_COLUMNS has no TYPE key (use GRIEVANCE_TYPE or ISSUE_CATEGORY)', () => {
    expect(GRIEVANCE_COLUMNS.TYPE).toBeUndefined();
    expect(GRIEVANCE_COLUMNS.GRIEVANCE_TYPE).toBeDefined();
    expect(GRIEVANCE_COLUMNS.ISSUE_CATEGORY).toBeDefined();
  });

  test('GRIEVANCE_COLUMNS.NOTES, OUTCOME, and RESOLUTION all resolve to same column', () => {
    // These three legacy aliases all point to the RESOLUTION column.
    // Code must not write them separately or later writes overwrite earlier ones.
    expect(GRIEVANCE_COLUMNS.NOTES).toBe(GRIEVANCE_COLUMNS.RESOLUTION);
    expect(GRIEVANCE_COLUMNS.OUTCOME).toBe(GRIEVANCE_COLUMNS.RESOLUTION);
  });

  test('GRIEVANCE_COLUMNS.STEP_X_STATUS aliases all resolve to STATUS', () => {
    // All step-status aliases point to the main STATUS column.
    expect(GRIEVANCE_COLUMNS.STEP_1_STATUS).toBe(GRIEVANCE_COLUMNS.STATUS);
    expect(GRIEVANCE_COLUMNS.STEP_2_STATUS).toBe(GRIEVANCE_COLUMNS.STATUS);
    expect(GRIEVANCE_COLUMNS.STEP_3_STATUS).toBe(GRIEVANCE_COLUMNS.STATUS);
  });
});
