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
 * @version 4.33.0
 */

var TimelineService = (function () {

  var DEFAULT_CATEGORIES = ['meeting', 'announcement', 'milestone', 'action', 'decision', 'other'];

  // ── In-memory cache (2-min TTL) ──────────────────────────────────────
  var _tlCache = {};
  var _TL_CACHE_TTL = 2 * 60 * 1000; // 2 minutes

  /**
   * Retrieves a value from the in-memory cache if it has not expired.
   * @param {string} key - Cache key to look up.
   * @returns {*|undefined} Cached value or undefined if missing/expired.
   */
  function _cacheGet(key) {
    var entry = _tlCache[key];
    if (entry && (Date.now() - entry.ts) < _TL_CACHE_TTL) return entry.val;
    delete _tlCache[key];
    return undefined;
  }
  /**
   * Stores a value in the in-memory cache with a timestamp.
   * @param {string} key - Cache key.
   * @param {*} val - Value to cache.
   */
  function _cacheSet(key, val) { _tlCache[key] = { val: val, ts: Date.now() }; }
  /** Clears all entries from the in-memory cache. */
  function _cacheInvalidate() { _tlCache = {}; }
  // ─────────────────────────────────────────────────────────────────────

  /**
   * Creates the _Timeline_Events sheet with headers if it does not exist, and ensures the category sheet is initialized.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The timeline events sheet.
   */
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

  /**
   * Creates the _Timeline_Categories sheet with default categories if it does not exist.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet instance; defaults to active spreadsheet.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The timeline categories sheet.
   */
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

  /**
   * Returns the list of timeline categories from the sheet, falling back to defaults.
   * @returns {string[]} Lowercase category names.
   */
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

  /**
   * Adds a new timeline category to the categories sheet.
   * @param {string} stewardEmail - Email of the steward performing the action.
   * @param {string} category - Category name to add.
   * @returns {{success: boolean, message?: string, category?: string}} Result object.
   */
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

  /**
   * Removes a category from the categories sheet by name.
   * @param {string} stewardEmail - Email of the steward performing the action.
   * @param {string} category - Category name to delete.
   * @returns {{success: boolean, message?: string}} Result object.
   */
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

  /**
   * Returns a descending list of distinct years that have timeline events.
   * @returns {number[]} Array of years sorted newest-first.
   */
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

  /**
   * Returns paginated timeline events, optionally filtered by year and category.
   * @param {number} [page=1] - Page number (1-based).
   * @param {number} [pageSize=25] - Number of events per page.
   * @param {number|string} [year] - Filter to events in this year.
   * @param {string} [category] - Filter to events matching this category.
   * @returns {{events: Object[], total: number, page: number, pageSize: number}} Paginated result.
   */
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

  /**
   * Creates a new timeline event row with optional calendar/Drive links.
   * @param {string} stewardEmail - Email of the steward creating the event.
   * @param {Object} data - Event data with title, eventDate, description, category, calendarEventId, driveFileIds, driveFileNames, meetingMinutesId.
   * @returns {{success: boolean, message?: string, eventId?: string}} Result object.
   */
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

  /**
   * Updates fields of an existing timeline event by ID.
   * @param {string} stewardEmail - Email of the steward performing the update.
   * @param {string} eventId - The timeline event ID to update.
   * @param {Object} updates - Fields to update (title, eventDate, description, category, meetingMinutesId).
   * @returns {{success: boolean, message?: string}} Result object.
   */
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

  /**
   * Deletes a timeline event row by ID.
   * @param {string} stewardEmail - Email of the steward performing the deletion.
   * @param {string} eventId - The timeline event ID to delete.
   * @returns {{success: boolean, message?: string}} Result object.
   */
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

  /**
   * Imports Google Calendar events within a date range as timeline events, skipping duplicates.
   * @param {string} stewardEmail - Email of the steward performing the import.
   * @param {string|Date} startDate - Start of the date range.
   * @param {string|Date} endDate - End of the date range.
   * @returns {{success: boolean, message?: string, imported?: number}} Result object.
   */
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

  /**
   * Attaches Google Drive files to an existing timeline event by appending file IDs and resolved names.
   * @param {string} stewardEmail - Email of the steward performing the action.
   * @param {string} eventId - The timeline event ID.
   * @param {string} fileIds - Comma-separated Drive file IDs.
   * @returns {{success: boolean, message?: string, filesAttached?: number}} Result object.
   */
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

  /**
   * Formats a Date using the shared short-date formatter.
   * @param {Date} date - Date to format.
   * @returns {string} Formatted date string.
   */
  function _fmtDate(date) { return fmtDateShort_(date); }

  /**
   * Zips parallel comma-separated ID and name strings into an array of {id, name} objects.
   * @param {string} idsStr - Comma-separated Drive file IDs.
   * @param {string} namesStr - Comma-separated Drive file names.
   * @returns {{id: string, name: string}[]} Array of file descriptor objects.
   */
  function _zipDriveFiles(idsStr, namesStr) {
    var ids = idsStr.split(',');
    var names = namesStr.split(',');
    var files = [];
    for (var i = 0; i < ids.length; i++) {
      if (ids[i].trim()) files.push({ id: ids[i].trim(), name: (names[i] || 'File').trim() });
    }
    return files;
  }

  /**
   * Sanitizes text by escaping formula injection characters if the helper is available.
   * @param {string} text - Raw text to sanitize.
   * @returns {string} Sanitized text.
   */
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

/**
 * Global wrapper: returns the list of timeline categories for an authenticated user.
 * @param {string} sessionToken - Session token for authentication.
 * @returns {string[]} Category names, or empty array if unauthenticated.
 */
function tlGetCategories(sessionToken)                                     { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getCategories(); }

/**
 * Global wrapper: returns distinct years that have timeline events.
 * @param {string} sessionToken - Session token for authentication.
 * @returns {number[]} Array of years, or empty array if unauthenticated.
 */
function tlGetTimelineYears(sessionToken)                                   { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getTimelineYears(); }

/**
 * Global wrapper: returns paginated timeline events with optional year/category filters.
 * @param {string} sessionToken - Session token for authentication.
 * @param {number} [page] - Page number.
 * @param {number} [pageSize] - Events per page.
 * @param {number|string} [year] - Year filter.
 * @param {string} [category] - Category filter.
 * @returns {Object} Paginated event result, or empty array if unauthenticated.
 */
function tlGetTimelineEvents(sessionToken, page, pageSize, year, category) { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return TimelineService.getTimelineEvents(page, pageSize, year, category); }

/**
 * Global wrapper: adds a timeline category (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string} category - Category name to add.
 * @returns {Object|null} Result object, or null if unauthorized.
 */
function tlAddCategory(sessionToken, category)                             { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.addCategory(e, category); }

/**
 * Global wrapper: deletes a timeline category (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string} category - Category name to delete.
 * @returns {Object|null} Result object, or null if unauthorized.
 */
function tlDeleteCategory(sessionToken, category)                          { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.deleteCategory(e, category); }

/**
 * Global wrapper: creates a new timeline event (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {Object} data - Event data object.
 * @returns {Object|null} Result object with eventId, or null if unauthorized.
 */
function tlAddTimelineEvent(sessionToken, data)                            { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.addTimelineEvent(e, data); }

/**
 * Global wrapper: updates an existing timeline event (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string} eventId - Event ID to update.
 * @param {Object} data - Fields to update.
 * @returns {Object|null} Result object, or null if unauthorized.
 */
function tlUpdateTimelineEvent(sessionToken, eventId, data)                { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.updateTimelineEvent(e, eventId, data); }

/**
 * Global wrapper: deletes a timeline event (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string} eventId - Event ID to delete.
 * @returns {Object|null} Result object, or null if unauthorized.
 */
function tlDeleteTimelineEvent(sessionToken, eventId)                      { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.deleteTimelineEvent(e, eventId); }

/**
 * Global wrapper: imports Google Calendar events into the timeline (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string|Date} startDate - Start of the date range.
 * @param {string|Date} endDate - End of the date range.
 * @returns {Object|null} Result object with imported count, or null if unauthorized.
 */
function tlImportCalendarEvents(sessionToken, startDate, endDate)          { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.importCalendarEvents(e, startDate, endDate); }

/**
 * Global wrapper: attaches Drive files to a timeline event (steward-only).
 * @param {string} sessionToken - Session token for steward authentication.
 * @param {string} eventId - Event ID to attach files to.
 * @param {string} fileIds - Comma-separated Drive file IDs.
 * @returns {Object|null} Result object, or null if unauthorized.
 */
function tlAttachDriveFiles(sessionToken, eventId, fileIds)                { var e = _requireStewardAuth(sessionToken); if (!e) return null; return TimelineService.attachDriveFiles(e, eventId, fileIds); }

/**
 * Global wrapper: initializes the timeline sheets (events + categories).
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The timeline events sheet.
 */
function tlInitSheets()                                                    { return TimelineService.initTimelineSheet(); }
