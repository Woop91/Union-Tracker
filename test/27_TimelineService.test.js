/**
 * Tests for 27_TimelineService.gs
 *
 * Covers the TimelineService IIFE: initTimelineSheet, getTimelineEvents,
 * addTimelineEvent, updateTimelineEvent, deleteTimelineEvent,
 * importCalendarEvents, attachDriveFiles, and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '27_TimelineService.gs']);

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

/** Standard timeline header row */
var TIMELINE_HEADERS = [
  'ID', 'Title', 'Event Date', 'Description', 'Category',
  'Calendar Event ID', 'Drive File IDs', 'Drive File Names',
  'Meeting Minutes ID', 'Created By', 'Created Date', 'Updated Date'
];

/** Build a data set with headers + rows for a mock Timeline sheet */
function buildTimelineData(rows) {
  return [TIMELINE_HEADERS].concat(rows || []);
}

/** Create a mock timeline sheet with custom data rows */
function buildTimelineSheet(rows) {
  var data = buildTimelineData(rows);
  var sheet = createMockSheet(SHEETS.TIMELINE_EVENTS, data);

  // Wire up getRange to support per-cell setValue and bulk getValues
  sheet.getRange = jest.fn(function (row, col, numRows, numCols) {
    return {
      getValue: jest.fn(function () {
        if (data[row - 1] && data[row - 1][col - 1] !== undefined) return data[row - 1][col - 1];
        return '';
      }),
      getValues: jest.fn(function () {
        var result = [];
        for (var r = row - 1; r < row - 1 + (numRows || 1); r++) {
          var rowArr = [];
          for (var c = col - 1; c < col - 1 + (numCols || 1); c++) {
            rowArr.push(data[r] && data[r][c] !== undefined ? data[r][c] : '');
          }
          result.push(rowArr);
        }
        return result;
      }),
      setValue: jest.fn(),
      setValues: jest.fn()
    };
  });

  return sheet;
}

/** Install a spreadsheet mock with the given timeline sheet (or null) */
function installSS(timelineSheet) {
  var sheets = timelineSheet ? [timelineSheet] : [];
  var ss = createMockSpreadsheet(sheets);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

// ---------------------------------------------------------------------------
// Global mocks
// ---------------------------------------------------------------------------

beforeEach(() => {
  jest.clearAllMocks();
  global.logAuditEvent = jest.fn();
  global.logIntegrityEvent = jest.fn();
});

// ============================================================================
// TimelineService.initTimelineSheet
// ============================================================================

describe('TimelineService.initTimelineSheet', () => {
  test('creates sheet if missing', () => {
    var ss = installSS(null);
    var result = TimelineService.initTimelineSheet();

    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.TIMELINE_EVENTS);
    expect(result).toBeTruthy();
  });

  test('does not recreate if sheet exists', () => {
    var existingSheet = buildTimelineSheet([]);
    var ss = installSS(existingSheet);
    var result = TimelineService.initTimelineSheet();

    expect(ss.insertSheet).not.toHaveBeenCalled();
    expect(result).toBe(existingSheet);
  });

  test('sets headers on creation', () => {
    var ss = installSS(null);
    // insertSheet returns a mock sheet — capture the call
    var created = TimelineService.initTimelineSheet();
    var rangeCall = created.getRange;
    expect(rangeCall).toHaveBeenCalledWith(1, 1, 1, 12);
  });
});

// ============================================================================
// TimelineService.getTimelineEvents
// ============================================================================

describe('TimelineService.getTimelineEvents', () => {
  test('returns empty array when no data rows', () => {
    var sheet = buildTimelineSheet([]);
    // Override getLastRow to return 1 (header only)
    sheet.getLastRow.mockReturnValue(1);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents();
    expect(result.events).toEqual([]);
    expect(result.total).toBe(0);
  });

  test('returns events with all fields', () => {
    var jan15 = new Date('2026-01-15T00:00:00');
    var created = new Date('2026-01-10T00:00:00');
    var sheet = buildTimelineSheet([
      ['TL_1', 'Meeting', jan15, 'Monthly meeting', 'meeting', 'cal123', 'f1,f2', 'File1.pdf,File2.doc', 'min1', 'admin@test.com', created, created]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents();
    expect(result.events.length).toBe(1);
    var evt = result.events[0];
    expect(evt.id).toBe('TL_1');
    expect(evt.title).toBe('Meeting');
    expect(evt.description).toBe('Monthly meeting');
    expect(evt.category).toBe('meeting');
    expect(evt.calendarEventId).toBe('cal123');
    expect(evt.createdBy).toBe('admin@test.com');
    expect(evt.driveFiles.length).toBe(2);
    expect(evt.driveFiles[0]).toEqual({ id: 'f1', name: 'File1.pdf' });
    expect(evt.driveFiles[1]).toEqual({ id: 'f2', name: 'File2.doc' });
  });

  test('supports year filter', () => {
    var jan2025 = new Date('2025-06-15T00:00:00');
    var jan2026 = new Date('2026-03-01T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Old Event', jan2025, 'Desc', 'meeting', '', '', '', '', 'a@b.com', now, now],
      ['TL_2', 'New Event', jan2026, 'Desc', 'action', '', '', '', '', 'a@b.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents(1, 25, 2026);
    expect(result.events.length).toBe(1);
    expect(result.events[0].id).toBe('TL_2');
  });

  test('supports category filter', () => {
    var d = new Date('2026-01-10T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'A Meeting', d, '', 'meeting', '', '', '', '', 'a@b.com', now, now],
      ['TL_2', 'An Action', d, '', 'action', '', '', '', '', 'a@b.com', now, now],
      ['TL_3', 'Another Meeting', d, '', 'meeting', '', '', '', '', 'a@b.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents(1, 25, null, 'action');
    expect(result.events.length).toBe(1);
    expect(result.events[0].id).toBe('TL_2');
  });

  test('sorts by date descending', () => {
    var jan = new Date('2026-01-01T00:00:00');
    var feb = new Date('2026-02-15T00:00:00');
    var mar = new Date('2026-03-10T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'January', jan, '', 'other', '', '', '', '', 'a@b.com', now, now],
      ['TL_2', 'March', mar, '', 'other', '', '', '', '', 'a@b.com', now, now],
      ['TL_3', 'February', feb, '', 'other', '', '', '', '', 'a@b.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents();
    expect(result.events[0].id).toBe('TL_2');
    expect(result.events[1].id).toBe('TL_3');
    expect(result.events[2].id).toBe('TL_1');
  });

  test('pagination (page, pageSize)', () => {
    var now = new Date();
    var rows = [];
    for (var i = 0; i < 10; i++) {
      var d = new Date('2026-01-' + String(10 + i).padStart(2, '0') + 'T00:00:00');
      rows.push(['TL_' + i, 'Event ' + i, d, '', 'other', '', '', '', '', 'a@b.com', now, now]);
    }
    var sheet = buildTimelineSheet(rows);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents(2, 3);
    expect(result.page).toBe(2);
    expect(result.pageSize).toBe(3);
    expect(result.events.length).toBe(3);
    expect(result.total).toBe(10);
  });

  test('returns total count', () => {
    var now = new Date();
    var d = new Date('2026-01-15T00:00:00');
    var sheet = buildTimelineSheet([
      ['TL_1', 'E1', d, '', 'other', '', '', '', '', 'a@b.com', now, now],
      ['TL_2', 'E2', d, '', 'other', '', '', '', '', 'a@b.com', now, now],
      ['TL_3', 'E3', d, '', 'other', '', '', '', '', 'a@b.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents(1, 2);
    expect(result.total).toBe(3);
    expect(result.events.length).toBe(2);
  });

  test('parses Drive files from comma-separated IDs', () => {
    var d = new Date('2026-02-01T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'E1', d, '', 'other', '', 'abc,def,ghi', 'Doc.pdf,Sheet.xlsx,Slides.pptx', '', 'a@b.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.getTimelineEvents();
    expect(result.events[0].driveFiles.length).toBe(3);
    expect(result.events[0].driveFiles[2]).toEqual({ id: 'ghi', name: 'Slides.pptx' });
  });
});

// ============================================================================
// TimelineService.addTimelineEvent
// ============================================================================

describe('TimelineService.addTimelineEvent', () => {
  test('rejects missing stewardEmail', () => {
    var result = TimelineService.addTimelineEvent(null, { title: 'Test' });
    expect(result.success).toBe(false);
  });

  test('rejects missing title', () => {
    var result = TimelineService.addTimelineEvent('user@test.com', { description: 'No title' });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Title');
  });

  test('sanitizes title with escapeForFormula', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    var spy = jest.spyOn(global, 'escapeForFormula');
    TimelineService.addTimelineEvent('user@test.com', { title: '=DANGEROUS()' });
    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });

  test('truncates title to 200 chars', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    var longTitle = 'A'.repeat(300);
    TimelineService.addTimelineEvent('user@test.com', { title: longTitle });

    var appendCall = sheet.appendRow.mock.calls[0][0];
    // The title (index 1) should be at most 200 chars (after sanitization)
    expect(appendCall[1].length).toBeLessThanOrEqual(200);
  });

  test('validates category against allowed list', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    TimelineService.addTimelineEvent('user@test.com', { title: 'Test', category: 'meeting' });
    var appendCall = sheet.appendRow.mock.calls[0][0];
    expect(appendCall[4]).toBe('meeting');
  });

  test('defaults invalid category to other', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    TimelineService.addTimelineEvent('user@test.com', { title: 'Test', category: 'INVALID_CAT' });
    var appendCall = sheet.appendRow.mock.calls[0][0];
    expect(appendCall[4]).toBe('other');
  });

  test('uses ScriptLock', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    TimelineService.addTimelineEvent('user@test.com', { title: 'Test' });
    expect(LockService.getScriptLock).toHaveBeenCalled();
  });

  test('returns {success:true, eventId}', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    var result = TimelineService.addTimelineEvent('user@test.com', { title: 'My Event' });
    expect(result.success).toBe(true);
    expect(result.eventId).toBeTruthy();
    expect(result.eventId).toMatch(/^TL_/);
  });
});

// ============================================================================
// TimelineService.updateTimelineEvent
// ============================================================================

describe('TimelineService.updateTimelineEvent', () => {
  var sheet;

  beforeEach(() => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    sheet = buildTimelineSheet([
      ['TL_1', 'Original Title', d, 'Original desc', 'meeting', '', '', '', '', 'admin@test.com', now, now]
    ]);
    installSS(sheet);
  });

  test('updates title', () => {
    var result = TimelineService.updateTimelineEvent('admin@test.com', 'TL_1', { title: 'Updated Title' });
    expect(result.success).toBe(true);
    // setValue called for title (col 2) and updatedDate (col 12)
    expect(sheet.getRange).toHaveBeenCalledWith(2, 2);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 12);
  });

  test('updates eventDate', () => {
    var result = TimelineService.updateTimelineEvent('admin@test.com', 'TL_1', { eventDate: '2026-06-01' });
    expect(result.success).toBe(true);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 3);
  });

  test('updates description', () => {
    var result = TimelineService.updateTimelineEvent('admin@test.com', 'TL_1', { description: 'New desc' });
    expect(result.success).toBe(true);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 4);
  });

  test('validates category on update', () => {
    var result = TimelineService.updateTimelineEvent('admin@test.com', 'TL_1', { category: 'decision' });
    expect(result.success).toBe(true);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 5);
  });

  test('returns not found for missing event', () => {
    var result = TimelineService.updateTimelineEvent('admin@test.com', 'TL_NONEXISTENT', { title: 'X' });
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });
});

// ============================================================================
// TimelineService.deleteTimelineEvent
// ============================================================================

describe('TimelineService.deleteTimelineEvent', () => {
  test('deletes row for valid event', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event To Delete', d, 'Desc', 'meeting', '', '', '', '', 'admin@test.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.deleteTimelineEvent('admin@test.com', 'TL_1');
    expect(result.success).toBe(true);
    expect(sheet.deleteRow).toHaveBeenCalledWith(2);
  });

  test('returns not found for missing event', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, 'Desc', 'meeting', '', '', '', '', 'admin@test.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.deleteTimelineEvent('admin@test.com', 'TL_NONEXISTENT');
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });

  test('rejects missing params', () => {
    var result = TimelineService.deleteTimelineEvent(null, null);
    expect(result.success).toBe(false);
    expect(result.message).toContain('Missing');
  });

  test('logs audit event on delete', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, 'Desc', 'meeting', '', '', '', '', 'admin@test.com', now, now]
    ]);
    installSS(sheet);

    TimelineService.deleteTimelineEvent('admin@test.com', 'TL_1');
    expect(global.logAuditEvent).toHaveBeenCalledWith(
      'TIMELINE_EVENT_DELETED',
      expect.stringContaining('TL_1')
    );
  });
});

// ============================================================================
// TimelineService.importCalendarEvents
// ============================================================================

describe('TimelineService.importCalendarEvents', () => {
  test('rejects missing params', () => {
    var result = TimelineService.importCalendarEvents(null, null, null);
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('handles calendar import errors gracefully', () => {
    CalendarApp.getAllCalendars.mockImplementation(() => {
      throw new Error('Calendar access denied');
    });
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    var result = TimelineService.importCalendarEvents('user@test.com', '2026-01-01', '2026-01-31');
    expect(result.success).toBe(false);
    expect(result.message).toContain('Calendar import failed');
  });

  test('avoids importing duplicates', () => {
    // Existing event with calendarEventId 'CAL_EXISTING'
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Existing', d, '', 'meeting', 'CAL_EXISTING', '', '', '', 'user@test.com', now, now]
    ]);
    installSS(sheet);

    // Mock calendar returning one event with the same ID
    var mockCalEvent = {
      getId: jest.fn(() => 'CAL_EXISTING'),
      getTitle: jest.fn(() => 'Duplicate'),
      getStartTime: jest.fn(() => new Date()),
      getDescription: jest.fn(() => '')
    };
    var mockCalendar = {
      getEvents: jest.fn(() => [mockCalEvent])
    };
    CalendarApp.getAllCalendars.mockReturnValue([mockCalendar]);

    var result = TimelineService.importCalendarEvents('user@test.com', '2026-01-01', '2026-01-31');
    expect(result.success).toBe(true);
    expect(result.imported).toBe(0);
    expect(sheet.appendRow).not.toHaveBeenCalled();
  });

  test('returns imported count', () => {
    var sheet = buildTimelineSheet([]);
    sheet.getLastRow.mockReturnValue(1);
    installSS(sheet);

    var mockCalEvent1 = {
      getId: jest.fn(() => 'CAL_NEW_1'),
      getTitle: jest.fn(() => 'New Event 1'),
      getStartTime: jest.fn(() => new Date('2026-01-10')),
      getDescription: jest.fn(() => 'Desc 1')
    };
    var mockCalEvent2 = {
      getId: jest.fn(() => 'CAL_NEW_2'),
      getTitle: jest.fn(() => 'New Event 2'),
      getStartTime: jest.fn(() => new Date('2026-01-20')),
      getDescription: jest.fn(() => 'Desc 2')
    };
    var mockCalendar = {
      getEvents: jest.fn(() => [mockCalEvent1, mockCalEvent2])
    };
    CalendarApp.getAllCalendars.mockReturnValue([mockCalendar]);

    var result = TimelineService.importCalendarEvents('user@test.com', '2026-01-01', '2026-01-31');
    expect(result.success).toBe(true);
    expect(result.imported).toBe(2);
    expect(sheet.appendRow).toHaveBeenCalledTimes(2);
  });
});

// ============================================================================
// TimelineService.attachDriveFiles
// ============================================================================

describe('TimelineService.attachDriveFiles', () => {
  test('attaches file IDs and names', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, '', 'meeting', '', '', '', '', 'user@test.com', now, now]
    ]);
    installSS(sheet);

    DriveApp.getFileById.mockReturnValue({ getName: jest.fn(() => 'Report.pdf') });

    var result = TimelineService.attachDriveFiles('user@test.com', 'TL_1', 'file123');
    expect(result.success).toBe(true);
    expect(result.filesAttached).toBe(1);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 7);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 8);
  });

  test('returns not found for missing event', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, '', 'meeting', '', '', '', '', 'user@test.com', now, now]
    ]);
    installSS(sheet);

    var result = TimelineService.attachDriveFiles('user@test.com', 'TL_NONEXISTENT', 'file123');
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });

  test('handles DriveApp errors gracefully', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, '', 'meeting', '', '', '', '', 'user@test.com', now, now]
    ]);
    installSS(sheet);

    DriveApp.getFileById.mockImplementation(() => { throw new Error('File not found'); });

    var result = TimelineService.attachDriveFiles('user@test.com', 'TL_1', 'bad_id');
    expect(result.success).toBe(true);
    // Should still succeed but with 'Unknown' name
    expect(result.filesAttached).toBe(1);
  });

  test('appends to existing files', () => {
    var d = new Date('2026-01-15T00:00:00');
    var now = new Date();
    var sheet = buildTimelineSheet([
      ['TL_1', 'Event', d, '', 'meeting', '', 'existingId', 'ExistingFile.pdf', '', 'user@test.com', now, now]
    ]);
    installSS(sheet);

    DriveApp.getFileById.mockReturnValue({ getName: jest.fn(() => 'NewFile.doc') });

    var result = TimelineService.attachDriveFiles('user@test.com', 'TL_1', 'newId');
    expect(result.success).toBe(true);
    expect(result.filesAttached).toBe(1);

    // getRange(2,7).setValue should have been called with combined IDs
    var setValueCalls = sheet.getRange.mock.results;
    // Verify the range for col 7 was accessed (Drive File IDs column)
    expect(sheet.getRange).toHaveBeenCalledWith(2, 7);
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('tlGetTimelineEvents delegates to TimelineService', () => {
    var sheet = buildTimelineSheet([]);
    sheet.getLastRow.mockReturnValue(1);
    installSS(sheet);

    var result = tlGetTimelineEvents(1, 10, null, null);
    expect(result).toHaveProperty('events');
    expect(result).toHaveProperty('total');
  });

  test('tlAddTimelineEvent delegates to TimelineService', () => {
    var sheet = buildTimelineSheet([]);
    installSS(sheet);

    var result = tlAddTimelineEvent('user@test.com', { title: 'From Wrapper' });
    expect(result.success).toBe(true);
    expect(result.eventId).toMatch(/^TL_/);
  });

  test('tlDeleteTimelineEvent delegates to TimelineService', () => {
    var result = tlDeleteTimelineEvent(null, null);
    expect(result.success).toBe(false);
  });

  test('tlInitSheets delegates to TimelineService', () => {
    var ss = installSS(null);
    tlInitSheets();
    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.TIMELINE_EVENTS);
  });
});
