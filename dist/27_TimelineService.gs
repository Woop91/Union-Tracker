/**
 * ============================================================================
 * 27_TimelineService.gs - Timeline of Events Module
 * ============================================================================
 *
 * Chronological event timeline with Google Calendar and Drive integration.
 *
 * Sheet:
 *   _Timeline_Events — event records (12 columns)
 *
 * @fileoverview Timeline Service IIFE module
 * @version 4.17.0
 * @requires 01_Core.gs, 06_Maintenance.gs
 */

var TimelineService = (function () {

  var CATEGORIES = ['meeting', 'announcement', 'milestone', 'action', 'decision', 'other'];

  // ═══════════════════════════════════════
  // Sheet Setup
  // ═══════════════════════════════════════

  function initTimelineSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken — getActiveSpreadsheet() returned null.');
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.TIMELINE_EVENTS);
      sheet.getRange(1, 1, 1, 12).setValues([[
        'ID', 'Title', 'Event Date', 'Description', 'Category',
        'Calendar Event ID', 'Drive File IDs', 'Drive File Names',
        'Meeting Minutes ID', 'Created By', 'Created Date', 'Updated Date'
      ]]);
      sheet.hideSheet();
    }
    return sheet;
  }

  // ═══════════════════════════════════════
  // CRUD Operations
  // ═══════════════════════════════════════

  function getTimelineEvents(page, pageSize, year, category) {
    page = page || 1;
    pageSize = pageSize || 25;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { events: [], total: 0, page: page, pageSize: pageSize };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet) { sheet = initTimelineSheet(); }
    if (!sheet || sheet.getLastRow() <= 1) return { events: [], total: 0, page: page, pageSize: pageSize };

    var data = sheet.getDataRange().getValues();
    var events = [];
    for (var i = 1; i < data.length; i++) {
      var eventDate = data[i][2];
      var eventYear = eventDate instanceof Date ? eventDate.getFullYear() : null;

      // Year filter
      if (year && eventYear && eventYear !== parseInt(year, 10)) continue;

      // Category filter
      var cat = String(data[i][4] || 'other').toLowerCase().trim();
      if (category && cat !== category.toLowerCase().trim()) continue;

      var driveIds = String(data[i][6] || '');
      var driveNames = String(data[i][7] || '');

      events.push({
        id: data[i][0],
        title: String(data[i][1] || ''),
        eventDate: eventDate instanceof Date ? _fmtDate(eventDate) : String(eventDate || ''),
        eventDateRaw: eventDate instanceof Date ? eventDate.getTime() : 0,
        description: String(data[i][3] || ''),
        category: cat,
        calendarEventId: String(data[i][5] || ''),
        driveFiles: driveIds ? _zipDriveFiles(driveIds, driveNames) : [],
        meetingMinutesId: String(data[i][8] || ''),
        createdBy: String(data[i][9] || ''),
        createdDate: data[i][10] instanceof Date ? _fmtDate(data[i][10]) : ''
      });
    }

    // Sort by event date descending
    events.sort(function (a, b) { return b.eventDateRaw - a.eventDateRaw; });

    var total = events.length;
    var start = (page - 1) * pageSize;
    var paged = events.slice(start, start + pageSize);
    return { events: paged, total: total, page: page, pageSize: pageSize };
  }

  function addTimelineEvent(stewardEmail, data) {
    if (!stewardEmail || !data || !data.title) return { success: false, message: 'Title is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet) { initTimelineSheet(); sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS); }

    var cat = String(data.category || 'other').toLowerCase().trim();
    if (CATEGORIES.indexOf(cat) === -1) cat = 'other';

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'TL_' + Date.now().toString(36);
      var now = new Date();
      var eventDate = data.eventDate ? new Date(data.eventDate) : now;

      sheet.appendRow([
        id, _sanitize(data.title.substring(0, 200)), eventDate,
        _sanitize((data.description || '').substring(0, 2000)), cat,
        data.calendarEventId || '', data.driveFileIds || '', data.driveFileNames || '',
        data.meetingMinutesId || '', stewardEmail.toLowerCase().trim(), now, now
      ]);

      logAuditEvent('TIMELINE_EVENT_ADDED', 'Event ' + id + ' by ' + stewardEmail);
      return { success: true, eventId: id };
    } finally {
      lock.releaseLock();
    }
  }

  function updateTimelineEvent(stewardEmail, eventId, updates) {
    if (!stewardEmail || !eventId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Event not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        if (updates.title) sheet.getRange(i + 1, 2).setValue(_sanitize(updates.title.substring(0, 200)));
        if (updates.eventDate) sheet.getRange(i + 1, 3).setValue(new Date(updates.eventDate));
        if (updates.description) sheet.getRange(i + 1, 4).setValue(_sanitize(updates.description.substring(0, 2000)));
        if (updates.category) {
          var cat = updates.category.toLowerCase().trim();
          if (CATEGORIES.indexOf(cat) !== -1) sheet.getRange(i + 1, 5).setValue(cat);
        }
        sheet.getRange(i + 1, 12).setValue(new Date());
        logAuditEvent('TIMELINE_EVENT_UPDATED', 'Event ' + eventId + ' updated by ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Event not found.' };
  }

  function deleteTimelineEvent(stewardEmail, eventId) {
    if (!stewardEmail || !eventId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Event not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        sheet.deleteRow(i + 1);
        logAuditEvent('TIMELINE_EVENT_DELETED', 'Event ' + eventId + ' deleted by ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Event not found.' };
  }

  // ═══════════════════════════════════════
  // Calendar Import
  // ═══════════════════════════════════════

  function importCalendarEvents(stewardEmail, startDate, endDate) {
    if (!stewardEmail || !startDate || !endDate) return { success: false, message: 'Date range required.' };

    try {
      var start = new Date(startDate);
      var end = new Date(endDate);
      var calendars = CalendarApp.getAllCalendars();
      var imported = 0;

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
      if (!sheet) { initTimelineSheet(); sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS); }

      // Collect existing calendar event IDs to avoid duplicates
      var existingCalIds = {};
      if (sheet.getLastRow() > 1) {
        var existing = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
        for (var e = 0; e < existing.length; e++) {
          if (existing[e][0]) existingCalIds[String(existing[e][0])] = true;
        }
      }

      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        for (var c = 0; c < calendars.length; c++) {
          var calEvents = calendars[c].getEvents(start, end);
          for (var j = 0; j < calEvents.length; j++) {
            var calEvent = calEvents[j];
            var calId = calEvent.getId();
            if (existingCalIds[calId]) continue;

            var id = 'TL_' + Date.now().toString(36) + '_' + imported;
            var now = new Date();
            sheet.appendRow([
              id, calEvent.getTitle().substring(0, 200), calEvent.getStartTime(),
              (calEvent.getDescription() || '').substring(0, 2000), 'meeting',
              calId, '', '', '', stewardEmail.toLowerCase().trim(), now, now
            ]);
            existingCalIds[calId] = true;
            imported++;
          }
        }
      } finally {
        lock.releaseLock();
      }

      logAuditEvent('TIMELINE_CALENDAR_IMPORTED', imported + ' events imported by ' + stewardEmail);
      return { success: true, imported: imported };
    } catch (err) {
      Logger.log('Calendar import error: ' + err.message);
      return { success: false, message: 'Calendar import failed: ' + err.message };
    }
  }

  // ═══════════════════════════════════════
  // Drive Integration
  // ═══════════════════════════════════════

  function attachDriveFiles(stewardEmail, eventId, fileIds) {
    if (!stewardEmail || !eventId || !fileIds) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Event not found.' };

    var idsArr = String(fileIds).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    var namesArr = [];
    for (var f = 0; f < idsArr.length; f++) {
      try {
        var file = DriveApp.getFileById(idsArr[f]);
        namesArr.push(file.getName());
      } catch (_e) {
        namesArr.push('Unknown');
      }
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        // Append to existing
        var existingIds = String(data[i][6] || '');
        var existingNames = String(data[i][7] || '');
        var newIds = existingIds ? existingIds + ',' + idsArr.join(',') : idsArr.join(',');
        var newNames = existingNames ? existingNames + ',' + namesArr.join(',') : namesArr.join(',');
        sheet.getRange(i + 1, 7).setValue(newIds);
        sheet.getRange(i + 1, 8).setValue(newNames);
        sheet.getRange(i + 1, 12).setValue(new Date());
        logAuditEvent('TIMELINE_FILES_ATTACHED', idsArr.length + ' files attached to event ' + eventId);
        return { success: true, filesAttached: idsArr.length };
      }
    }
    return { success: false, message: 'Event not found.' };
  }

  // ═══════════════════════════════════════
  // Helpers
  // ═══════════════════════════════════════

  // Delegate to shared helper in 01_Core.gs (eliminates duplicate definition)
  function _fmtDate(date) { return fmtDateShort_(date); }

  function _zipDriveFiles(idsStr, namesStr) {
    var ids = idsStr.split(',');
    var names = namesStr.split(',');
    var files = [];
    for (var i = 0; i < ids.length; i++) {
      if (ids[i].trim()) {
        files.push({ id: ids[i].trim(), name: (names[i] || 'File').trim() });
      }
    }
    return files;
  }

  function _sanitize(text) {
    if (typeof escapeForFormula === 'function') text = escapeForFormula(text);
    return text;
  }

  // ═══════════════════════════════════════
  // Public API
  // ═══════════════════════════════════════

  return {
    initTimelineSheet: initTimelineSheet,
    getTimelineEvents: getTimelineEvents,
    addTimelineEvent: addTimelineEvent,
    updateTimelineEvent: updateTimelineEvent,
    deleteTimelineEvent: deleteTimelineEvent,
    importCalendarEvents: importCalendarEvents,
    attachDriveFiles: attachDriveFiles
  };

})();


// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════

function tlGetTimelineEvents(sessionToken, page, pageSize, year, category) { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getTimelineEvents(page, pageSize, year, category); }
function tlAddTimelineEvent(sessionToken, data) { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.addTimelineEvent(e, data); }
function tlUpdateTimelineEvent(sessionToken, eventId, data) { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.updateTimelineEvent(e, eventId, data); }
function tlDeleteTimelineEvent(sessionToken, eventId) { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.deleteTimelineEvent(e, eventId); }
function tlImportCalendarEvents(sessionToken, startDate, endDate) { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.importCalendarEvents(e, startDate, endDate); }
function tlAttachDriveFiles(sessionToken, eventId, fileIds) { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.attachDriveFiles(e, eventId, fileIds); }
function tlInitSheets() { return TimelineService.initTimelineSheet(); }
