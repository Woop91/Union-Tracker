/**
 * ============================================================================
 * 27_TimelineService.gs - Timeline of Events Module
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Chronological event timeline with Google Calendar and Drive integration.
 *   Stewards create timeline events (meetings, announcements, milestones,
 *   actions, decisions) that appear in the SPA activity feed. Events can be
 *   linked to Calendar events and Drive documents. Supports inline editing,
 *   meeting minutes linking, pagination, and dynamic categories. Two hidden
 *   sheets:
 *     _Timeline_Events     — event records (12 columns)
 *     _Timeline_Categories — steward-managed category list
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Categories are fully dynamic — driven by _Timeline_Categories sheet
 *   rather than hardcoded, so each union can customize their event types.
 *   In-memory cache (2-minute TTL) prevents redundant sheet reads for
 *   rapidly accessed timeline data. Default categories (meeting,
 *   announcement, milestone, action, decision, other) are seeded on first
 *   init but can be modified.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The SPA timeline tab shows no events. New events can't be created.
 *   Calendar links and Drive document links are broken. If the cache fails,
 *   every timeline view triggers a sheet read (slow but functional).
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS), CalendarApp, DriveApp (GAS built-ins).
 *   Used by SPA timeline views and event management features.
 *
 * @version 4.31.0
 */

var TimelineService = (function () {

  var DEFAULT_CATEGORIES = ['meeting', 'announcement', 'milestone', 'action', 'decision', 'other'];

  // ── In-memory cache (2-min TTL) ──────────────────────────────────────
  var _tlCache = {};
  var _TL_CACHE_TTL = 2 * 60 * 1000; // 2 minutes

  function _cacheGet(key) {
    var entry = _tlCache[key];
    if (entry && (Date.now() - entry.ts) < _TL_CACHE_TTL) return entry.val;
    delete _tlCache[key];
    return undefined;
  }
  function _cacheSet(key, val) { _tlCache[key] = { val: val, ts: Date.now() }; }
  function _cacheInvalidate() { _tlCache = {}; }
  // ─────────────────────────────────────────────────────────────────────

  function initTimelineSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
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
    _initCategorySheet(ss);
    return sheet;
  }

  function _initCategorySheet(ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.TIMELINE_CATEGORIES);
      sheet.getRange(1, 1).setValue('Category');
      DEFAULT_CATEGORIES.forEach(function (cat, i) {
        sheet.getRange(i + 2, 1).setValue(cat);
      });
      sheet.hideSheet();
    }
    return sheet;
  }

  function getCategories() {
    var cached = _cacheGet('categories');
    if (cached) return cached;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return DEFAULT_CATEGORIES.slice();
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES);
    if (!sheet) { _initCategorySheet(ss); sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES); }
    if (!sheet || sheet.getLastRow() <= 1) return DEFAULT_CATEGORIES.slice();
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var cats = [];
    data.forEach(function (row) {
      var c = String(row[0] || '').trim().toLowerCase();
      if (c) cats.push(c);
    });
    var result = cats.length ? cats : DEFAULT_CATEGORIES.slice();
    _cacheSet('categories', result);
    return result;
  }

  function addCategory(stewardEmail, category) {
    if (!stewardEmail || !category) return { success: false, message: 'Category name required.' };
    var cat = String(category).trim().toLowerCase().replace(/[^a-z0-9 _-]/g, '').substring(0, 40);
    if (!cat) return { success: false, message: 'Invalid category name.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES);
    if (!sheet) { _initCategorySheet(ss); sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES); }
    var existing = getCategories();
    if (existing.indexOf(cat) !== -1) return { success: false, message: 'Category already exists.' };
    sheet.appendRow([cat]);
    _cacheInvalidate();
    logAuditEvent('TIMELINE_CATEGORY_ADDED', 'Category "' + cat + '" added by ' + stewardEmail);
    return { success: true, category: cat };
  }

  function deleteCategory(stewardEmail, category) {
    if (!stewardEmail || !category) return { success: false, message: 'Missing data.' };
    var cat = String(category).trim().toLowerCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_CATEGORIES);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Category not found.' };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === cat) {
        sheet.deleteRow(i + 2);
        _cacheInvalidate();
        logAuditEvent('TIMELINE_CATEGORY_DELETED', 'Category "' + cat + '" deleted by ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Category not found.' };
  }

  function getTimelineYears() {
    var cached = _cacheGet('years');
    if (cached) return cached;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    var yearsSet = {};
    data.forEach(function (row) {
      if (row[0] instanceof Date) yearsSet[row[0].getFullYear()] = true;
    });
    var result = Object.keys(yearsSet).map(Number).sort(function (a, b) { return b - a; });
    _cacheSet('years', result);
    return result;
  }

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
      if (year && eventYear && eventYear !== parseInt(year, 10)) continue;
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
    var validCats = getCategories();
    var cat = String(data.category || validCats[0] || 'other').toLowerCase().trim();
    if (validCats.indexOf(cat) === -1) cat = validCats[0] || 'other';
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'System busy. Please try again in a moment.' };
    try {
      var id = 'TL_' + Date.now().toString(36);
      var now = new Date();
      var eventDate = data.eventDate ? new Date(data.eventDate) : now;
      var driveIds   = String(data.driveFileIds   || '').trim();
      var driveNames = String(data.driveFileNames || '').trim();
      if (driveIds && !driveNames) {
        var idsArr = driveIds.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
        var namesArr = [];
        for (var f = 0; f < idsArr.length; f++) {
          try { namesArr.push(DriveApp.getFileById(idsArr[f]).getName()); }
          catch (_e) { namesArr.push('Unknown'); }
        }
        driveNames = namesArr.join(',');
      }
      sheet.appendRow([
        id, _sanitize(data.title.substring(0, 200)), eventDate,
        _sanitize((data.description || '').substring(0, 2000)), cat,
        data.calendarEventId || '', driveIds, driveNames,
        data.meetingMinutesId || '', stewardEmail.toLowerCase().trim(), now, now
      ]);
      _cacheInvalidate();
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
    var validCats = getCategories();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        if (updates.title) sheet.getRange(i + 1, 2).setValue(_sanitize(updates.title.substring(0, 200)));
        if (updates.eventDate) sheet.getRange(i + 1, 3).setValue(new Date(updates.eventDate));
        if (updates.description !== undefined) sheet.getRange(i + 1, 4).setValue(_sanitize(updates.description.substring(0, 2000)));
        if (updates.category) {
          var cat = updates.category.toLowerCase().trim();
          if (validCats.indexOf(cat) !== -1) sheet.getRange(i + 1, 5).setValue(cat);
        }
        if (updates.meetingMinutesId !== undefined) sheet.getRange(i + 1, 9).setValue(String(updates.meetingMinutesId || ''));
        sheet.getRange(i + 1, 12).setValue(new Date());
        _cacheInvalidate();
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
        _cacheInvalidate();
        logAuditEvent('TIMELINE_EVENT_DELETED', 'Event ' + eventId + ' deleted by ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Event not found.' };
  }

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
      var existingCalIds = {};
      if (sheet.getLastRow() > 1) {
        var existing = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
        for (var e = 0; e < existing.length; e++) {
          if (existing[e][0]) existingCalIds[String(existing[e][0])] = true;
        }
      }
      var lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) return { success: false, message: 'System busy. Please try again in a moment.', imported: 0 };
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

  function attachDriveFiles(stewardEmail, eventId, fileIds) {
    if (!stewardEmail || !eventId || !fileIds) return { success: false, message: 'Missing data.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Event not found.' };
    var idsArr = String(fileIds).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    var namesArr = [];
    for (var f = 0; f < idsArr.length; f++) {
      try { namesArr.push(DriveApp.getFileById(idsArr[f]).getName()); }
      catch (_e) { namesArr.push('Unknown'); }
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        var existingIds = String(data[i][6] || '');
        var existingNames = String(data[i][7] || '');
        sheet.getRange(i + 1, 7).setValue(existingIds ? existingIds + ',' + idsArr.join(',') : idsArr.join(','));
        sheet.getRange(i + 1, 8).setValue(existingNames ? existingNames + ',' + namesArr.join(',') : namesArr.join(','));
        sheet.getRange(i + 1, 12).setValue(new Date());
        logAuditEvent('TIMELINE_FILES_ATTACHED', idsArr.length + ' files attached to event ' + eventId);
        return { success: true, filesAttached: idsArr.length };
      }
    }
    return { success: false, message: 'Event not found.' };
  }

  function _fmtDate(date) { return fmtDateShort_(date); }

  function _zipDriveFiles(idsStr, namesStr) {
    var ids = idsStr.split(',');
    var names = namesStr.split(',');
    var files = [];
    for (var i = 0; i < ids.length; i++) {
      if (ids[i].trim()) files.push({ id: ids[i].trim(), name: (names[i] || 'File').trim() });
    }
    return files;
  }

  function _sanitize(text) {
    if (typeof escapeForFormula === 'function') text = escapeForFormula(text);
    return text;
  }

  return {
    initTimelineSheet:    initTimelineSheet,
    getCategories:        getCategories,
    addCategory:          addCategory,
    deleteCategory:       deleteCategory,
    getTimelineYears:     getTimelineYears,
    getTimelineEvents:    getTimelineEvents,
    addTimelineEvent:     addTimelineEvent,
    updateTimelineEvent:  updateTimelineEvent,
    deleteTimelineEvent:  deleteTimelineEvent,
    importCalendarEvents: importCalendarEvents,
    attachDriveFiles:     attachDriveFiles
  };

})();
// ═══════════════════════════════════════
// GLOBAL WRAPPERS
// ═══════════════════════════════════════

function tlGetCategories(sessionToken)                                     { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getCategories(); }
function tlGetTimelineYears(sessionToken)                                   { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getTimelineYears(); }
function tlGetTimelineEvents(sessionToken, page, pageSize, year, category) { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getTimelineEvents(page, pageSize, year, category); }
function tlAddCategory(sessionToken, category)                             { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.addCategory(e, category); }
function tlDeleteCategory(sessionToken, category)                          { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.deleteCategory(e, category); }
function tlAddTimelineEvent(sessionToken, data)                            { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.addTimelineEvent(e, data); }
function tlUpdateTimelineEvent(sessionToken, eventId, data)                { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.updateTimelineEvent(e, eventId, data); }
function tlDeleteTimelineEvent(sessionToken, eventId)                      { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.deleteTimelineEvent(e, eventId); }
function tlImportCalendarEvents(sessionToken, startDate, endDate)          { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.importCalendarEvents(e, startDate, endDate); }
function tlAttachDriveFiles(sessionToken, eventId, fileIds)                { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.attachDriveFiles(e, eventId, fileIds); }
function tlInitSheets()                                                    { return TimelineService.initTimelineSheet(); }
