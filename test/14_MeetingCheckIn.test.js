/**
 * Tests for 14_MeetingCheckIn.gs
 *
 * Covers meeting creation, check-in processing, attendance retrieval,
 * status lifecycle management, cleanup, and HTML template generation.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load dependencies then meeting check-in module
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '13_MemberSelfService.gs',
  '14_MeetingCheckIn.gs'
]);

// ============================================================================
// createMeeting
// ============================================================================

describe('createMeeting', () => {
  test('rejects null input', () => {
    var result = createMeeting(null);
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('rejects missing name', () => {
    var result = createMeeting({ date: '2026-03-15' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('name');
  });

  test('rejects missing date', () => {
    var result = createMeeting({ name: 'Team Meeting' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('date');
  });

  test('creates meeting with valid data', () => {
    var meetingSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name', 'Date', 'Type', 'Member ID', 'Member Name', 'Check-In Time', 'Email', 'Time', 'Duration', 'Status', 'Notify', 'CalendarEventId', 'NotesUrl', 'AgendaUrl', 'AgendaStewards']
    ]);
    var ss = createMockSpreadsheet([meetingSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = createMeeting({
      name: 'Monthly Meeting',
      date: '2026-03-15',
      time: '14:00',
      duration: '1.5',
      type: 'In-Person'
    });

    expect(result.success).toBe(true);
    expect(result.meetingId).toMatch(/^MTG-/);
    expect(meetingSheet.appendRow).toHaveBeenCalled();
  });

  test('attempts to auto-create sheet if missing', () => {
    var newSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([]);
    // First call returns null (sheet missing), second call returns the new sheet
    ss.getSheetByName
      .mockReturnValueOnce(null)
      .mockReturnValue(newSheet);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    global.createMeetingCheckInLogSheet = jest.fn();

    var result = createMeeting({
      name: 'Monthly Meeting',
      date: '2026-03-15'
    });

    expect(global.createMeetingCheckInLogSheet).toHaveBeenCalled();
    expect(result.success).toBe(true);
    delete global.createMeetingCheckInLogSheet;
  });

  test('meeting ID contains date string', () => {
    var meetingSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([meetingSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = createMeeting({
      name: 'Test',
      date: '2026-05-20'
    });

    if (result.success) {
      expect(result.meetingId).toContain('20260520');
    }
  });

  test('defaults type to In-Person when not provided', () => {
    var meetingSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([meetingSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = createMeeting({
      name: 'Test Meeting',
      date: '2026-06-01'
    });

    expect(result.success).toBe(true);
    var appendedRow = meetingSheet.appendRow.mock.calls[0][0];
    expect(appendedRow[3]).toBe('In-Person');
  });

  test('defaults duration to 1 hour when not provided', () => {
    var meetingSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([meetingSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    createMeeting({
      name: 'Test',
      date: '2026-06-01'
    });

    var appendedRow = meetingSheet.appendRow.mock.calls[0][0];
    expect(appendedRow[9]).toBe(1);
  });

  test('truncates long meeting names to 200 characters', () => {
    var meetingSheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([meetingSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var longName = 'A'.repeat(300);
    createMeeting({
      name: longName,
      date: '2026-06-01'
    });

    var appendedRow = meetingSheet.appendRow.mock.calls[0][0];
    expect(appendedRow[1].length).toBe(200);
  });
});

// ============================================================================
// processMeetingCheckIn
// ============================================================================

describe('processMeetingCheckIn', () => {
  test('rejects missing meetingId', () => {
    var result = processMeetingCheckIn('', 'test@example.com', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('rejects missing email', () => {
    var result = processMeetingCheckIn('MTG-001', '', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('rejects missing PIN', () => {
    var result = processMeetingCheckIn('MTG-001', 'test@example.com', '');
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('rejects all empty inputs', () => {
    var result = processMeetingCheckIn('', '', '');
    expect(result.success).toBe(false);
  });

  test('rejects invalid email format', () => {
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [
      ['header']
    ]);
    var ss = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = processMeetingCheckIn('MTG-001', 'not-an-email', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toContain('valid email');
  });

  test('returns error when member directory sheet is missing', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = processMeetingCheckIn('MTG-001', 'test@example.com', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toContain('Member directory');
  });

  test('returns error when member not found by email', () => {
    var memberData = [
      new Array(35).fill('header'),
      new Array(35).fill('')
    ];
    memberData[1][MEMBER_COLS.EMAIL - 1] = 'other@example.com';
    memberData[1][MEMBER_COLS.MEMBER_ID - 1] = 'MEM-001';

    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var ss = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = processMeetingCheckIn('MTG-001', 'unknown@example.com', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toContain('No member found');
  });
});

// ============================================================================
// getMeetingAttendees
// ============================================================================

describe('getMeetingAttendees', () => {
  test('rejects missing meetingId', () => {
    var result = getMeetingAttendees('');
    expect(result.success).toBe(false);
    expect(result.error).toContain('Meeting ID');
  });

  test('returns empty list when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getMeetingAttendees('MTG-001');
    expect(result.success).toBe(true);
    expect(result.attendees).toEqual([]);
    expect(result.count).toBe(0);
  });

  test('returns empty list when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID', 'Meeting Name']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getMeetingAttendees('MTG-001');
    expect(result.success).toBe(true);
    expect(result.attendees).toEqual([]);
    expect(result.count).toBe(0);
  });

  test('returns attendees for matching meeting ID', () => {
    var data = [
      new Array(16).fill('header'),
      new Array(16).fill(''),
      new Array(16).fill('')
    ];
    // Row with matching meeting ID and member
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-20260315-001';
    data[1][MEETING_CHECKIN_COLS.MEMBER_ID - 1] = 'MEM-001';
    data[1][MEETING_CHECKIN_COLS.MEMBER_NAME - 1] = 'Jane Doe';
    data[1][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1] = new Date('2026-03-15T14:05:00');
    // Row with different meeting ID
    data[2][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-20260316-001';
    data[2][MEETING_CHECKIN_COLS.MEMBER_ID - 1] = 'MEM-002';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(3);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getMeetingAttendees('MTG-20260315-001');
    expect(result.success).toBe(true);
    expect(result.count).toBe(1);
    expect(result.attendees[0].memberId).toBe('MEM-001');
    expect(result.attendees[0].name).toBe('Jane Doe');
  });

  test('skips rows without member ID (header rows)', () => {
    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-20260315-001';
    data[1][MEETING_CHECKIN_COLS.MEMBER_ID - 1] = '';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getMeetingAttendees('MTG-20260315-001');
    expect(result.success).toBe(true);
    expect(result.count).toBe(0);
  });
});

// ============================================================================
// getActiveMeetings
// ============================================================================

describe('getActiveMeetings', () => {
  test('returns empty list when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getActiveMeetings();
    expect(result.success).toBe(true);
    expect(result.meetings).toEqual([]);
  });

  test('returns empty list when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getActiveMeetings();
    expect(result.success).toBe(true);
    expect(result.meetings).toEqual([]);
  });

  test('excludes completed meetings', () => {
    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = new Date('2099-12-31');
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.COMPLETED;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getActiveMeetings();
    expect(result.meetings.length).toBe(0);
  });

  test('includes future scheduled meetings', () => {
    var futureDate = new Date();
    futureDate.setDate(futureDate.getDate() + 30);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_NAME - 1] = 'Future Meeting';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = futureDate;
    data[1][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] = 'Virtual';
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getActiveMeetings();
    expect(result.meetings.length).toBe(1);
    expect(result.meetings[0].name).toBe('Future Meeting');
    expect(result.meetings[0].status).toBe(MEETING_STATUS.SCHEDULED);
  });

  test('deduplicates meetings by ID', () => {
    var futureDate = new Date();
    futureDate.setDate(futureDate.getDate() + 5);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill(''),
      new Array(16).fill('')
    ];
    // Two rows with same meeting ID (header + check-in)
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = futureDate;
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;
    data[1][MEETING_CHECKIN_COLS.MEETING_NAME - 1] = 'Same Meeting';
    data[2][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[2][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = futureDate;
    data[2][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;
    data[2][MEETING_CHECKIN_COLS.MEETING_NAME - 1] = 'Same Meeting';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(3);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getActiveMeetings();
    expect(result.meetings.length).toBe(1);
  });
});

// ============================================================================
// getCheckInEligibleMeetings
// ============================================================================

describe('getCheckInEligibleMeetings', () => {
  test('returns empty list when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getCheckInEligibleMeetings();
    expect(result.success).toBe(true);
    expect(result.meetings).toEqual([]);
  });

  test('returns empty list when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getCheckInEligibleMeetings();
    expect(result.success).toBe(true);
    expect(result.meetings).toEqual([]);
  });

  test('excludes completed meetings', () => {
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = today;
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.COMPLETED;
    data[1][MEETING_CHECKIN_COLS.MEETING_TIME - 1] = '09:00';
    data[1][MEETING_CHECKIN_COLS.MEETING_DURATION - 1] = 1;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getCheckInEligibleMeetings();
    expect(result.meetings.length).toBe(0);
  });

  test('excludes meetings not dated today', () => {
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = tomorrow;
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;
    data[1][MEETING_CHECKIN_COLS.MEETING_TIME - 1] = '09:00';
    data[1][MEETING_CHECKIN_COLS.MEETING_DURATION - 1] = 1;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getCheckInEligibleMeetings();
    expect(result.meetings.length).toBe(0);
  });
});

// ============================================================================
// updateMeetingStatuses
// ============================================================================

describe('updateMeetingStatuses', () => {
  test('returns zeros when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = updateMeetingStatuses();
    expect(result.activated).toBe(0);
    expect(result.deactivated).toBe(0);
  });

  test('returns zeros when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = updateMeetingStatuses();
    expect(result.activated).toBe(0);
    expect(result.deactivated).toBe(0);
  });

  test('activates today scheduled meetings', () => {
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = today;
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;
    data[1][MEETING_CHECKIN_COLS.MEETING_TIME - 1] = '23:59';
    data[1][MEETING_CHECKIN_COLS.MEETING_DURATION - 1] = 1;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = updateMeetingStatuses();
    expect(result.activated).toBe(1);
  });

  test('deactivates past-date scheduled meetings', () => {
    var pastDate = new Date();
    pastDate.setDate(pastDate.getDate() - 7);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-001';
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = pastDate;
    data[1][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = MEETING_STATUS.SCHEDULED;
    data[1][MEETING_CHECKIN_COLS.MEETING_TIME - 1] = '09:00';
    data[1][MEETING_CHECKIN_COLS.MEETING_DURATION - 1] = 1;

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = updateMeetingStatuses();
    expect(result.deactivated).toBeGreaterThanOrEqual(1);
  });

  test('skips rows without meeting ID', () => {
    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = '';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = updateMeetingStatuses();
    expect(result.activated).toBe(0);
    expect(result.deactivated).toBe(0);
  });
});

// ============================================================================
// cleanupExpiredMeetings
// ============================================================================

describe('cleanupExpiredMeetings', () => {
  test('returns 0 when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = cleanupExpiredMeetings();
    expect(result).toBe(0);
  });

  test('returns 0 when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, [
      ['Meeting ID']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = cleanupExpiredMeetings();
    expect(result).toBe(0);
  });

  test('deletes meetings older than 90 days', () => {
    var oldDate = new Date();
    oldDate.setDate(oldDate.getDate() - 100);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = oldDate;
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-OLD';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = cleanupExpiredMeetings();
    expect(result).toBe(1);
    expect(sheet.deleteRow).toHaveBeenCalledWith(2);
  });

  test('does not delete meetings within 90 days', () => {
    var recentDate = new Date();
    recentDate.setDate(recentDate.getDate() - 30);

    var data = [
      new Array(16).fill('header'),
      new Array(16).fill('')
    ];
    data[1][MEETING_CHECKIN_COLS.MEETING_DATE - 1] = recentDate;
    data[1][MEETING_CHECKIN_COLS.MEETING_ID - 1] = 'MTG-RECENT';

    var sheet = createMockSheet(SHEETS.MEETING_CHECKIN_LOG, data);
    sheet.getLastRow.mockReturnValue(2);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = cleanupExpiredMeetings();
    expect(result).toBe(0);
    expect(sheet.deleteRow).not.toHaveBeenCalled();
  });
});

// ============================================================================
// getStewardEmailsForMeetingSetup
// ============================================================================

describe('getStewardEmailsForMeetingSetup', () => {
  test('returns empty array when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getStewardEmailsForMeetingSetup();
    expect(result).toEqual([]);
  });

  test('returns empty array when sheet has only headers', () => {
    var sheet = createMockSheet(SHEETS.MEMBER_DIR, [
      ['header']
    ]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getStewardEmailsForMeetingSetup();
    expect(result).toEqual([]);
  });
});

// ============================================================================
// HTML Templates
// ============================================================================

describe('showSetupMeetingDialog', () => {
  test('creates HTML dialog via HtmlService', () => {
    var mockUi = SpreadsheetApp.getUi();
    mockUi.showModalDialog = jest.fn();
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showSetupMeetingDialog();
    expect(HtmlService.createHtmlOutput).toHaveBeenCalled();
    expect(mockUi.showModalDialog).toHaveBeenCalled();
  });
});

describe('showMeetingCheckInDialog', () => {
  test('creates HTML dialog via HtmlService', () => {
    var mockUi = SpreadsheetApp.getUi();
    mockUi.showModalDialog = jest.fn();
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showMeetingCheckInDialog();
    expect(HtmlService.createHtmlOutput).toHaveBeenCalled();
    expect(mockUi.showModalDialog).toHaveBeenCalled();
  });
});

describe('getSetupMeetingHtml_', () => {
  test('returns a string', () => {
    var html = getSetupMeetingHtml_();
    expect(typeof html).toBe('string');
  });

  test('contains DOCTYPE declaration', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('<!DOCTYPE html>');
  });

  test('contains meeting name input field', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('meetingName');
  });

  test('contains meeting date input field', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('meetingDate');
  });

  test('contains duration selector', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('meetingDuration');
  });

  test('contains meeting type selector', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('meetingType');
  });

  test('contains notification email field', () => {
    var html = getSetupMeetingHtml_();
    expect(html).toContain('notifyEmails');
  });
});

describe('getMeetingCheckInHtml_', () => {
  test('returns a string', () => {
    var html = getMeetingCheckInHtml_();
    expect(typeof html).toBe('string');
  });

  test('contains DOCTYPE declaration', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('<!DOCTYPE html>');
  });

  test('contains email input field', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('type="email"');
  });

  test('contains PIN input field', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('type="password"');
    expect(html).toContain('pin');
  });

  test('contains meeting selector', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('meetingSelect');
  });

  test('contains check-in button', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('checkinBtn');
  });

  test('contains client-side escapeHtml function', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('escapeHtml');
  });

  test('calls getCheckInEligibleMeetings on load', () => {
    var html = getMeetingCheckInHtml_();
    expect(html).toContain('getCheckInEligibleMeetings');
  });
});
