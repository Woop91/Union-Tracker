/**
 * ============================================================================
 * 00_DataAccess.gs - Data Access Layer (DAL)
 * ============================================================================
 *
 * This module provides a centralized data access layer for all spreadsheet
 * operations. It implements:
 * - Cached spreadsheet/sheet references
 * - Batch read/write operations for performance
 * - Consistent error handling
 * - Single point of access for all data operations
 *
 * Benefits:
 * - Reduces API calls by caching references
 * - Provides consistent error handling
 * - Makes testing easier (can mock this layer)
 * - Centralizes all data access patterns
 *
 * @fileoverview Data Access Layer for the dashboard
 * @version 1.0.0
 */

// ============================================================================
// SPREADSHEET CACHE
// ============================================================================

/**
 * Cached spreadsheet and sheet references
 * @private
 */
var _spreadsheetCache = {
  ss: null,
  sheets: {},
  lastAccess: null,
  cacheTimeout: 300000  // 5 minutes
};

/**
 * Data Access Layer - singleton pattern
 * @namespace
 */
var DataAccess = {

  // ==========================================================================
  // SPREADSHEET ACCESS
  // ==========================================================================

  /**
   * Gets the active spreadsheet (cached)
   * @returns {Spreadsheet} The active spreadsheet
   */
  getSpreadsheet: function() {
    var now = new Date().getTime();

    // Check if cache is valid
    if (_spreadsheetCache.ss &&
        _spreadsheetCache.lastAccess &&
        (now - _spreadsheetCache.lastAccess) < _spreadsheetCache.cacheTimeout) {
      return _spreadsheetCache.ss;
    }

    // Refresh cache
    _spreadsheetCache.ss = SpreadsheetApp.getActiveSpreadsheet();
    _spreadsheetCache.lastAccess = now;
    _spreadsheetCache.sheets = {};  // Clear sheet cache too

    return _spreadsheetCache.ss;
  },

  /**
   * Gets a sheet by name (cached)
   * @param {string} sheetName - Name of the sheet
   * @returns {Sheet|null} The sheet or null if not found
   */
  getSheet: function(sheetName) {
    if (!sheetName) return null;

    // Check sheet cache - verify cached sheet is still valid
    if (_spreadsheetCache.sheets[sheetName]) {
      try {
        // Verify the cached sheet reference is still valid
        _spreadsheetCache.sheets[sheetName].getName();
        return _spreadsheetCache.sheets[sheetName];
      } catch (_e) {
        // Cached reference is stale (sheet may have been deleted/renamed)
        delete _spreadsheetCache.sheets[sheetName];
      }
    }

    // Get from spreadsheet
    var ss = this.getSpreadsheet();
    try {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        _spreadsheetCache.sheets[sheetName] = sheet;
      }
      return sheet;
    } catch (e) {
      Logger.log('Error accessing sheet "' + sheetName + '": ' + e.message);
      return null;
    }
  },

  /**
   * Gets or creates a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Object} [options] - Options for sheet creation
   * @param {Array} [options.headers] - Header row values
   * @param {boolean} [options.hidden] - Whether to hide the sheet
   * @returns {Sheet} The existing or new sheet
   */
  getOrCreateSheet: function(sheetName, options) {
    var sheet = this.getSheet(sheetName);

    if (!sheet) {
      var ss = this.getSpreadsheet();
      sheet = ss.insertSheet(sheetName);
      _spreadsheetCache.sheets[sheetName] = sheet;

      if (options) {
        if (options.headers && Array.isArray(options.headers)) {
          sheet.getRange(1, 1, 1, options.headers.length).setValues([options.headers]);
          sheet.getRange(1, 1, 1, options.headers.length).setFontWeight('bold');
        }
        if (options.hidden) {
          setSheetVeryHidden_(sheet);
        }
      }
    }

    return sheet;
  },

  /**
   * Invalidates the cache (call after structural changes)
   */
  invalidateCache: function() {
    _spreadsheetCache.ss = null;
    _spreadsheetCache.sheets = {};
    _spreadsheetCache.lastAccess = null;
  },

  // ==========================================================================
  // DATA READING - BATCH OPERATIONS
  // ==========================================================================

  /**
   * Gets all data from a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Object} [options] - Options
   * @param {boolean} [options.includeHeaders=true] - Include header row
   * @returns {Array<Array>} 2D array of values
   */
  getAllData: function(sheetName, options) {
    options = options || {};
    var includeHeaders = options.includeHeaders !== false;

    var sheet = this.getSheet(sheetName);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();

    if (!includeHeaders && data.length > 0) {
      return data.slice(1);
    }

    return data;
  },

  /**
   * Gets a single row by row number
   * @param {string} sheetName - Name of the sheet
   * @param {number} rowNumber - Row number (1-indexed)
   * @returns {Array} Row values
   */
  getRow: function(sheetName, rowNumber) {
    var sheet = this.getSheet(sheetName);
    if (!sheet) return [];

    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) return [];

    return sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  },

  /**
   * Gets multiple rows efficiently (batch operation)
   * @param {string} sheetName - Name of the sheet
   * @param {Array<number>} rowNumbers - Array of row numbers (1-indexed)
   * @returns {Object} Map of rowNumber -> row values
   */
  getRows: function(sheetName, rowNumbers) {
    var result = {};
    if (!rowNumbers || rowNumbers.length === 0) return result;

    var sheet = this.getSheet(sheetName);
    if (!sheet) return result;

    // For efficiency, get all data and filter
    var allData = sheet.getDataRange().getValues();

    for (var i = 0; i < rowNumbers.length; i++) {
      var rowNum = rowNumbers[i];
      if (rowNum > 0 && rowNum <= allData.length) {
        result[rowNum] = allData[rowNum - 1];
      }
    }

    return result;
  },

  /**
   * Finds a row by a column value
   * @param {string} sheetName - Name of the sheet
   * @param {number} columnIndex - Column index (0-indexed)
   * @param {*} searchValue - Value to search for
   * @returns {Object|null} Object with {rowNumber, data} or null
   */
  findRow: function(sheetName, columnIndex, searchValue) {
    var data = this.getAllData(sheetName, { includeHeaders: true });

    for (var i = 1; i < data.length; i++) {  // Skip header
      if (String(data[i][columnIndex]) === String(searchValue)) {  // String comparison for consistent matching
        return {
          rowNumber: i + 1,  // 1-indexed
          data: data[i]
        };
      }
    }

    return null;
  },

  /**
   * Finds all rows matching a condition
   * @param {string} sheetName - Name of the sheet
   * @param {Function} predicate - Function(row, rowIndex) returning boolean
   * @returns {Array<Object>} Array of {rowNumber, data} objects
   */
  findAllRows: function(sheetName, predicate) {
    var results = [];
    var data = this.getAllData(sheetName, { includeHeaders: true });

    for (var i = 1; i < data.length; i++) {  // Skip header
      if (predicate(data[i], i)) {
        results.push({
          rowNumber: i + 1,  // 1-indexed
          data: data[i]
        });
      }
    }

    return results;
  },

  // ==========================================================================
  // DATA WRITING - BATCH OPERATIONS
  // ==========================================================================

  /**
   * Sets a single cell value
   * @param {string} sheetName - Name of the sheet
   * @param {number} row - Row number (1-indexed)
   * @param {number} col - Column number (1-indexed)
   * @param {*} value - Value to set
   */
  setCell: function(sheetName, row, col, value) {
    var sheet = this.getSheet(sheetName);
    if (!sheet) return;

    sheet.getRange(row, col).setValue(value);
  },

  /**
   * Sets multiple cells efficiently (batch operation)
   * @param {string} sheetName - Name of the sheet
   * @param {Array<Object>} cells - Array of {row, col, value}
   */
  setCells: function(sheetName, cells) {
    if (!cells || cells.length === 0) return;

    var sheet = this.getSheet(sheetName);
    if (!sheet) return;

    // Compute bounding box for a single setValues() call
    var minRow = cells[0].row, maxRow = cells[0].row;
    var minCol = cells[0].col, maxCol = cells[0].col;
    for (var i = 1; i < cells.length; i++) {
      if (cells[i].row < minRow) minRow = cells[i].row;
      if (cells[i].row > maxRow) maxRow = cells[i].row;
      if (cells[i].col < minCol) minCol = cells[i].col;
      if (cells[i].col > maxCol) maxCol = cells[i].col;
    }

    var numRows = maxRow - minRow + 1;
    var numCols = maxCol - minCol + 1;

    // Read current values for the bounding box region
    var range = sheet.getRange(minRow, minCol, numRows, numCols);
    var values = range.getValues();

    // Apply updates into the 2D array
    for (var j = 0; j < cells.length; j++) {
      values[cells[j].row - minRow][cells[j].col - minCol] = cells[j].value;
    }

    // Write back in a single call
    range.setValues(values);
  },

  /**
   * Updates an entire row
   * @param {string} sheetName - Name of the sheet
   * @param {number} rowNumber - Row number (1-indexed)
   * @param {Array} values - Values for the row
   * @param {number} [startCol=1] - Starting column (1-indexed)
   */
  setRow: function(sheetName, rowNumber, values, startCol) {
    var sheet = this.getSheet(sheetName);
    if (!sheet || !values || values.length === 0) return;

    startCol = startCol || 1;
    sheet.getRange(rowNumber, startCol, 1, values.length).setValues([values]);
  },

  /**
   * Appends a new row to the sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Array} values - Values for the new row
   * @returns {number} The row number of the new row
   */
  appendRow: function(sheetName, values) {
    var sheet = this.getSheet(sheetName);
    if (!sheet || !values || values.length === 0) return -1;

    sheet.appendRow(values);
    return sheet.getLastRow();
  },

  /**
   * Deletes a row
   * @param {string} sheetName - Name of the sheet
   * @param {number} rowNumber - Row number (1-indexed)
   */
  deleteRow: function(sheetName, rowNumber) {
    var sheet = this.getSheet(sheetName);
    if (!sheet) return;

    sheet.deleteRow(rowNumber);
  },

  // ==========================================================================
  // MEMBER-SPECIFIC OPERATIONS
  // ==========================================================================

  /**
   * Gets a member by ID
   * @param {string} memberId - The member ID
   * @returns {Object|null} Member data object or null
   */
  getMemberById: function(memberId) {
    if (!memberId) return null;

    var sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) ?
                    SHEETS.MEMBER_DIR : 'Member Directory';

    var result = this.findRow(sheetName, MEMBER_COLS.MEMBER_ID - 1, memberId);

    if (!result) return null;

    var row = result.data;
    return {
      rowNumber: result.rowNumber,
      memberId: row[MEMBER_COLS.MEMBER_ID - 1],
      firstName: row[MEMBER_COLS.FIRST_NAME - 1],
      lastName: row[MEMBER_COLS.LAST_NAME - 1],
      jobTitle: row[MEMBER_COLS.JOB_TITLE - 1],
      workLocation: row[MEMBER_COLS.WORK_LOCATION - 1],
      unit: row[MEMBER_COLS.UNIT - 1],
      email: row[MEMBER_COLS.EMAIL - 1],
      phone: row[MEMBER_COLS.PHONE - 1],
      isSteward: row[MEMBER_COLS.IS_STEWARD - 1],
      assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1],
      hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1],
      grievanceStatus: row[MEMBER_COLS.GRIEVANCE_STATUS - 1]
    };
  },

  /**
   * Gets all members
   * @param {Object} [options] - Filter options
   * @param {string} [options.unit] - Filter by unit
   * @param {string} [options.location] - Filter by location
   * @param {boolean} [options.stewardsOnly] - Only return stewards
   * @returns {Array<Object>} Array of member objects
   */
  getAllMembers: function(options) {
    options = options || {};

    var sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) ?
                    SHEETS.MEMBER_DIR : 'Member Directory';

    var data = this.getAllData(sheetName, { includeHeaders: false });
    var members = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // Skip empty rows
      if (row[MEMBER_COLS.MEMBER_ID - 1] === '' || row[MEMBER_COLS.MEMBER_ID - 1] == null) continue;

      // Apply filters (use String() coercion for type-safe comparison)
      if (options.unit && String(row[MEMBER_COLS.UNIT - 1]) !== String(options.unit)) continue;
      if (options.location && String(row[MEMBER_COLS.WORK_LOCATION - 1]) !== String(options.location)) continue;
      if (options.stewardsOnly) {
        var isSteward = row[MEMBER_COLS.IS_STEWARD - 1];
        if (!isTruthyValue(isSteward)) continue;
      }

      members.push({
        rowNumber: i + 2,  // Account for header and 0-index
        memberId: row[MEMBER_COLS.MEMBER_ID - 1],
        firstName: row[MEMBER_COLS.FIRST_NAME - 1],
        lastName: row[MEMBER_COLS.LAST_NAME - 1],
        jobTitle: row[MEMBER_COLS.JOB_TITLE - 1],
        workLocation: row[MEMBER_COLS.WORK_LOCATION - 1],
        unit: row[MEMBER_COLS.UNIT - 1],
        email: row[MEMBER_COLS.EMAIL - 1],
        phone: row[MEMBER_COLS.PHONE - 1],
        isSteward: row[MEMBER_COLS.IS_STEWARD - 1],
        assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1]
      });
    }

    return members;
  },

  // ==========================================================================
  // GRIEVANCE-SPECIFIC OPERATIONS
  // ==========================================================================

  /**
   * Gets a grievance by ID
   * @param {string} grievanceId - The grievance ID
   * @returns {Object|null} Grievance data object or null
   */
  getGrievanceById: function(grievanceId) {
    if (!grievanceId) return null;

    var sheetName = (typeof SHEETS !== 'undefined' && SHEETS.GRIEVANCE_LOG) ?
                    SHEETS.GRIEVANCE_LOG : 'Grievance Log';

    var result = this.findRow(sheetName, GRIEVANCE_COLS.GRIEVANCE_ID - 1, grievanceId);

    if (!result) return null;

    var row = result.data;
    return {
      rowNumber: result.rowNumber,
      grievanceId: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
      memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1],
      firstName: row[GRIEVANCE_COLS.FIRST_NAME - 1],
      lastName: row[GRIEVANCE_COLS.LAST_NAME - 1],
      status: row[GRIEVANCE_COLS.STATUS - 1],
      currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1],
      incidentDate: row[GRIEVANCE_COLS.INCIDENT_DATE - 1],
      filingDeadline: row[GRIEVANCE_COLS.FILING_DEADLINE - 1],
      dateFiled: row[GRIEVANCE_COLS.DATE_FILED - 1],
      daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1],
      nextActionDue: row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1],
      daysToDeadline: row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
      articles: row[GRIEVANCE_COLS.ARTICLES - 1],
      issueCategory: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
      location: row[GRIEVANCE_COLS.LOCATION - 1],
      steward: row[GRIEVANCE_COLS.STEWARD - 1],
      resolution: row[GRIEVANCE_COLS.RESOLUTION - 1]
    };
  },

  /**
   * Gets all grievances with optional filters
   * @param {Object} [options] - Filter options
   * @param {string} [options.status] - Filter by status
   * @param {string} [options.steward] - Filter by steward name
   * @param {boolean} [options.overdueOnly] - Only return overdue grievances
   * @returns {Array<Object>} Array of grievance objects
   */
  getAllGrievances: function(options) {
    options = options || {};

    var sheetName = (typeof SHEETS !== 'undefined' && SHEETS.GRIEVANCE_LOG) ?
                    SHEETS.GRIEVANCE_LOG : 'Grievance Log';

    var data = this.getAllData(sheetName, { includeHeaders: false });
    var grievances = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // Skip empty rows
      if (row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] === '' || row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] == null) continue;

      // Apply filters (use String() coercion for type-safe comparison)
      if (options.status && String(row[GRIEVANCE_COLS.STATUS - 1]) !== String(options.status)) continue;
      if (options.steward && String(row[GRIEVANCE_COLS.STEWARD - 1]) !== String(options.steward)) continue;
      if (options.overdueOnly) {
        var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
        if (typeof daysToDeadline !== 'number' || daysToDeadline >= 0) continue;
      }

      grievances.push({
        rowNumber: i + 2,
        grievanceId: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1],
        firstName: row[GRIEVANCE_COLS.FIRST_NAME - 1],
        lastName: row[GRIEVANCE_COLS.LAST_NAME - 1],
        status: row[GRIEVANCE_COLS.STATUS - 1],
        currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1],
        incidentDate: row[GRIEVANCE_COLS.INCIDENT_DATE - 1],
        daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1],
        daysToDeadline: row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
        issueCategory: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
        steward: row[GRIEVANCE_COLS.STEWARD - 1]
      });
    }

    return grievances;
  }
};

// ============================================================================
// TIME CONSTANTS
// ============================================================================

/**
 * Time-related constants for deadline calculations
 * @const {Object}
 */
var TIME_CONSTANTS = {
  /** Milliseconds per second */
  MS_PER_SECOND: 1000,

  /** Milliseconds per minute */
  MS_PER_MINUTE: 60 * 1000,

  /** Milliseconds per hour */
  MS_PER_HOUR: 60 * 60 * 1000,

  /** Milliseconds per day */
  MS_PER_DAY: 24 * 60 * 60 * 1000,

  /** Milliseconds per week */
  MS_PER_WEEK: 7 * 24 * 60 * 60 * 1000,

  /** Deadline days configuration — references DEADLINE_DEFAULTS (01_Core.gs) as single source of truth */
  DEADLINE_DAYS: {
    /** Days to file grievance after incident */
    get FILING() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.FILING_DAYS : 21; },

    /** Days for Step I response */
    get STEP1_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_1_RESPONSE : 7; },

    /** Days to appeal to Step II */
    get STEP2_APPEAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_APPEAL : 7; },

    /** Days for Step II response */
    get STEP2_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_RESPONSE : 14; },

    /** Days to appeal to Step III */
    get STEP3_APPEAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_3_APPEAL : 10; },

    /** Days before deadline to show warning */
    get WARNING_THRESHOLD() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.WARNING_THRESHOLD : 5; },

    /** Days before deadline to show critical alert */
    get CRITICAL_THRESHOLD() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.CRITICAL_THRESHOLD : 2; }
  },

  /** Reminder intervals — references DEADLINE_DEFAULTS (01_Core.gs) as single source of truth */
  REMINDER_DAYS: {
    /** First reminder before deadline */
    get FIRST() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_FIRST : 7; },

    /** Second reminder before deadline */
    get SECOND() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_SECOND : 3; },

    /** Final reminder before deadline */
    get FINAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_FINAL : 1; }
  }
};

/**
 * Calculates a deadline date from a start date
 * @param {Date} startDate - The start date
 * @param {number} days - Number of days to add
 * @returns {Date} The deadline date
 */
function calculateDeadline(startDate, days) {
  if (!startDate || !(startDate instanceof Date)) {
    startDate = new Date();
  }
  return new Date(startDate.getTime() + (days * TIME_CONSTANTS.MS_PER_DAY));
}

/**
 * Calculates days between two dates
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {number} Number of days (can be negative)
 */
function daysBetween(startDate, endDate) {
  if (!startDate || !endDate) return 0;

  var start = startDate instanceof Date ? startDate : new Date(startDate);
  var end = endDate instanceof Date ? endDate : new Date(endDate);

  return Math.floor((end.getTime() - start.getTime()) / TIME_CONSTANTS.MS_PER_DAY);
}

/**
 * Gets the urgency level based on days to deadline
 * @param {number} daysToDeadline - Days until deadline
 * @returns {string} Urgency level ('critical', 'warning', 'normal', 'overdue')
 */
function getDeadlineUrgency(daysToDeadline) {
  if (daysToDeadline < 0) return 'overdue';
  if (daysToDeadline <= TIME_CONSTANTS.DEADLINE_DAYS.CRITICAL_THRESHOLD) return 'critical';
  if (daysToDeadline <= TIME_CONSTANTS.DEADLINE_DAYS.WARNING_THRESHOLD) return 'warning';
  return 'normal';
}
