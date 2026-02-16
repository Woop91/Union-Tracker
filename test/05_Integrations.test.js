/**
 * Tests for 05_Integrations.gs
 *
 * Covers CALENDAR_CONFIG existence, step string comparison logic,
 * and DRIVE_CONFIG structure.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// We need to mock some functions that Integrations.gs calls
global.logAuditEvent = jest.fn();
global.getGrievanceById = jest.fn();
global.updateGrievanceFolderLink = jest.fn();
global.sanitizeFolderName = jest.fn(name => name.replace(/[^a-zA-Z0-9\s\-_.,]/g, ''));
global.BATCH_LIMITS = {
  MAX_EXECUTION_TIME_MS: 300000,
  MAX_API_CALLS_PER_BATCH: 50,
  PAUSE_BETWEEN_BATCHES_MS: 1000
};
global.AUDIT_EVENTS = {
  FOLDER_CREATED: 'FOLDER_CREATED',
  CALENDAR_SYNCED: 'CALENDAR_SYNCED'
};

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '05_Integrations.gs']);

// ============================================================================
// CALENDAR_CONFIG
// ============================================================================

describe('CALENDAR_CONFIG', () => {
  test('is defined (bug fix verification — was missing before)', () => {
    expect(CALENDAR_CONFIG).toBeDefined();
  });

  test('has CALENDAR_NAME', () => {
    expect(CALENDAR_CONFIG.CALENDAR_NAME).toBe('Grievance Deadlines');
  });

  test('has REMINDER_DAYS array', () => {
    expect(Array.isArray(CALENDAR_CONFIG.REMINDER_DAYS)).toBe(true);
    expect(CALENDAR_CONFIG.REMINDER_DAYS).toEqual([7, 3, 1]);
  });

  test('REMINDER_DAYS are in descending order', () => {
    const days = CALENDAR_CONFIG.REMINDER_DAYS;
    for (let i = 1; i < days.length; i++) {
      expect(days[i]).toBeLessThan(days[i - 1]);
    }
  });
});

// ============================================================================
// DRIVE_CONFIG
// ============================================================================

describe('DRIVE_CONFIG', () => {
  test('has ROOT_FOLDER_NAME', () => {
    expect(DRIVE_CONFIG.ROOT_FOLDER_NAME).toBeTruthy();
  });

  test('SUBFOLDER_TEMPLATE contains placeholders', () => {
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{lastName}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{firstName}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{date}');
  });

  test('SUBFOLDER_TEMPLATE_SIMPLE contains fallback placeholders', () => {
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE_SIMPLE).toContain('{grievanceId}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE_SIMPLE).toContain('{date}');
  });
});

// ============================================================================
// Calendar step comparison (string-based, not number)
// ============================================================================

describe('Calendar step comparison', () => {
  // The bug was: calendar sync compared step values as numbers (1, 2, 3)
  // when Grievance Log stores them as strings ('Step I', 'Step II', 'Step III')
  test('COMMAND_CONFIG.ESCALATION_STEPS uses string step values', () => {
    const steps = COMMAND_CONFIG.ESCALATION_STEPS;
    expect(steps).toContain('Step II');
    expect(steps).toContain('Step III');
    expect(steps).toContain('Arbitration');
    // Should NOT contain numeric values
    expect(steps).not.toContain(1);
    expect(steps).not.toContain(2);
    expect(steps).not.toContain(3);
  });
});

// ============================================================================
// Email HTML escaping
// ============================================================================

describe('Email HTML escaping', () => {
  test('escapeHtml prevents XSS in email content', () => {
    const memberName = '<script>alert(1)</script>';
    const escaped = escapeHtml(memberName);
    expect(escaped).not.toContain('<script>');
    expect(escaped).toContain('&lt;script&gt;');
  });
});

// ============================================================================
// sanitizeFolderName
// ============================================================================

describe('sanitizeFolderName', () => {
  test('returns Unknown for empty input', () => {
    expect(sanitizeFolderName('')).toBe('Unknown');
    expect(sanitizeFolderName(null)).toBe('Unknown');
    expect(sanitizeFolderName(undefined)).toBe('Unknown');
  });

  test('removes illegal characters', () => {
    expect(sanitizeFolderName('file<>:"/\\|?*name')).toBe('filename');
  });

  test('replaces spaces with underscores', () => {
    expect(sanitizeFolderName('John Doe Case')).toBe('John_Doe_Case');
  });

  test('collapses multiple spaces', () => {
    expect(sanitizeFolderName('John   Doe')).toBe('John_Doe');
  });

  test('truncates to 50 characters', () => {
    const longName = 'A'.repeat(60);
    expect(sanitizeFolderName(longName).length).toBe(50);
  });

  test('handles normal names without changes', () => {
    expect(sanitizeFolderName('GRV-001_Smith_2026')).toBe('GRV-001_Smith_2026');
  });
});

// ============================================================================
// BATCH_LIMITS configuration
// ============================================================================

describe('BATCH_LIMITS', () => {
  test('has reasonable execution time limit', () => {
    expect(BATCH_LIMITS.MAX_EXECUTION_TIME_MS).toBe(300000);
  });

  test('has max API calls per batch', () => {
    expect(BATCH_LIMITS.MAX_API_CALLS_PER_BATCH).toBe(50);
  });
});

// ============================================================================
// CALENDAR_CONFIG structure
// ============================================================================

describe('CALENDAR_CONFIG structure', () => {
  test('has MEETING_CALENDAR_NAME', () => {
    expect(CALENDAR_CONFIG.MEETING_CALENDAR_NAME).toBeTruthy();
  });

  test('has MEETING_DEACTIVATE_HOURS', () => {
    expect(typeof CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS).toBe('number');
    expect(CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS).toBeGreaterThan(0);
  });
});
