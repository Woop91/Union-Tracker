/**
 * Tests for 10b_SurveyDocSheets.gs
 * Covers survey question sheet, satisfaction sheet, feedback sheet,
 * FAQ, features reference, resources, and notification sheet creation.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '05_Integrations.gs', '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs', '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs'
]);

describe('10b function existence', () => {
  const required = [
    'createSurveyQuestionsSheet', 'clearSurveyQuestionsCache',
    'createSatisfactionSheet',
    'createGettingStartedSheet', 'createFAQSheet',
    'createFeaturesReferenceSheet', 'createResourcesSheet',
    'createNotificationsSheet', 'createResourceConfigSheet'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('Notification sheet uses dynamic email', () => {
  test('createNotificationsSheet does not hardcode massability.org', () => {
    const fs = require('fs');
    const path = require('path');
    const code = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10b_SurveyDocSheets.gs'), 'utf8'
    );
    // Should use systemEmail variable, not hardcoded
    expect(code).not.toContain("'system@massability.org'");
  });
});

describe('Sheet name constants', () => {
  test('SHEETS.NOTIFICATIONS is defined', () => {
    expect(SHEETS.NOTIFICATIONS).toBeDefined();
    expect(typeof SHEETS.NOTIFICATIONS).toBe('string');
  });

  test('SHEETS.SATISFACTION is defined', () => {
    expect(SHEETS.SATISFACTION).toBeDefined();
  });

  test('SHEETS.SURVEY_QUESTIONS is defined', () => {
    expect(SHEETS.SURVEY_QUESTIONS).toBeDefined();
    expect(typeof SHEETS.SURVEY_QUESTIONS).toBe('string');
  });
});

// ============================================================================
// createSurveyQuestionsSheet behavioral tests
// ============================================================================

describe('createSurveyQuestionsSheet behavioral', () => {
  test('creates sheet with exactly 16 headers', () => {
    const ss = createMockSpreadsheet([]);
    let capturedHeaders = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 1 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            capturedHeaders = vals[0];
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    expect(() => createSurveyQuestionsSheet(ss)).not.toThrow();
    expect(capturedHeaders).not.toBeNull();
    expect(capturedHeaders.length).toBe(16);
  });

  test('header row contains expected column names', () => {
    const expectedHeaders = [
      'Question ID', 'Section', 'Section Key', 'Section Title',
      'Question Text', 'Type', 'Required', 'Active', 'Options',
      'Branch Parent', 'Branch Value', 'Branch Target', 'Max Selections',
      'Slider Min Label', 'Slider Max Label', 'Notes'
    ];

    const ss = createMockSpreadsheet([]);
    let capturedHeaders = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 1 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            capturedHeaders = vals[0];
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    expect(capturedHeaders).toEqual(expectedHeaders);
  });

  test('seeds survey questions data on new sheet', () => {
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        // Seed data is written starting from row 2 with 16 columns
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    expect(seedData).not.toBeNull();
    expect(seedData.length).toBeGreaterThan(60); // 76+ seed questions
  });

  test('seed data starts with question q1', () => {
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    expect(seedData[0][0]).toBe('q1');
  });

  test('each seed row has exactly 16 columns', () => {
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    seedData.forEach((row, i) => {
      expect(row.length).toBe(16);
    });
  });

  test('seed data covers all expected sections', () => {
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    const sectionKeys = new Set(seedData.map(row => row[2])); // Section Key is col 3 (index 2)
    expect(sectionKeys.has('WORK_CONTEXT')).toBe(true);
    expect(sectionKeys.has('OVERALL_SAT')).toBe(true);
    expect(sectionKeys.has('STEWARD_3A')).toBe(true);
    expect(sectionKeys.has('STEWARD_3B')).toBe(true);
    expect(sectionKeys.has('CHAPTER')).toBe(true);
    expect(sectionKeys.has('LEADERSHIP')).toBe(true);
    expect(sectionKeys.has('CONTRACT')).toBe(true);
    expect(sectionKeys.has('COMMUNICATION')).toBe(true);
    expect(sectionKeys.has('MEMBER_VOICE')).toBe(true);
    expect(sectionKeys.has('VALUE_ACTION')).toBe(true);
    expect(sectionKeys.has('SCHEDULING')).toBe(true);
    expect(sectionKeys.has('PRIORITIES')).toBe(true);
  });

  test('all question IDs are unique', () => {
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    const ids = seedData.map(row => row[0]);
    const uniqueIds = new Set(ids);
    expect(uniqueIds.size).toBe(ids.length);
  });

  test('all questions have a valid type', () => {
    const validTypes = ['dropdown', 'radio', 'radio-branch', 'slider-10', 'paragraph', 'checkbox'];
    const ss = createMockSpreadsheet([]);
    let seedData = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 2 && col === 1 && numCols === 16) {
          range.setValues = jest.fn(function(vals) {
            seedData = vals;
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    seedData.forEach((row, i) => {
      const type = row[5]; // Type is col 6 (index 5)
      expect(validTypes).toContain(type);
    });
  });

  test('freezes header row', () => {
    const ss = createMockSpreadsheet([]);
    let frozenCalled = false;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      sheet.setFrozenRows = jest.fn(() => { frozenCalled = true; });
      return sheet;
    });

    createSurveyQuestionsSheet(ss);
    expect(frozenCalled).toBe(true);
  });
});

// ============================================================================
// SURVEY_QUESTIONS_COLS positional integrity
// ============================================================================

describe('SURVEY_QUESTIONS_COLS positional integrity', () => {
  test('QUESTION_ID is column 1', () => {
    expect(SURVEY_QUESTIONS_COLS.QUESTION_ID).toBe(1);
  });

  test('SECTION_NUM is column 2', () => {
    expect(SURVEY_QUESTIONS_COLS.SECTION_NUM).toBe(2);
  });

  test('TYPE is column 6', () => {
    expect(SURVEY_QUESTIONS_COLS.TYPE).toBe(6);
  });

  test('ACTIVE is column 8', () => {
    expect(SURVEY_QUESTIONS_COLS.ACTIVE).toBe(8);
  });

  test('all 16 column positions are defined', () => {
    const cols = SURVEY_QUESTIONS_COLS;
    const positions = Object.values(cols).filter(v => typeof v === 'number');
    expect(positions.length).toBeGreaterThanOrEqual(16);
    // Max position should be 16
    expect(Math.max(...positions)).toBe(16);
  });
});
