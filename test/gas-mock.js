/**
 * Google Apps Script environment mocks for testing.
 *
 * Provides mock implementations of GAS global objects so that
 * source files can be loaded and tested in a Node.js environment.
 */

// --- Logger ---
global.Logger = {
  log: jest.fn()
};

// --- Utilities ---
global.Utilities = {
  getUuid: jest.fn(() => 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'),
  formatDate: jest.fn(function(date, tz, format) {
    if (!date) return '';
    var d = new Date(date);
    // Basic format support for the most common patterns used in tests
    var yyyy = d.getFullYear();
    var MM = String(d.getMonth() + 1).padStart(2, '0');
    var dd = String(d.getDate()).padStart(2, '0');
    if (format === 'yyyy-MM-dd') return yyyy + '-' + MM + '-' + dd;
    if (format === 'MM/dd/yyyy') return MM + '/' + dd + '/' + yyyy;
    if (format === 'M/d/yyyy') return (d.getMonth()+1) + '/' + d.getDate() + '/' + yyyy;
    return d.toISOString().slice(0, 10);
  }),
  computeDigest: jest.fn((algo, data, charset) => {
    // Mock hash with full avalanche: every input byte affects all output bytes
    const hash = new Array(32).fill(0);
    for (let i = 0; i < data.length; i++) {
      const c = data.charCodeAt(i);
      for (let j = 0; j < 32; j++) {
        hash[j] = (hash[j] * 31 + c + i + j) & 0xFF;
      }
    }
    return hash;
  }),
  base64EncodeWebSafe: jest.fn((input) => {
    if (Array.isArray(input)) {
      return Buffer.from(input).toString('base64url');
    }
    return Buffer.from(String(input)).toString('base64url');
  }),
  newBlob: jest.fn((data, type, name) => ({ data, type, name })),
  base64Decode: jest.fn((input) => [1, 2, 3]),
  base64Encode: jest.fn((input) => {
    // Simple mock base64 encoding for byte arrays
    if (Array.isArray(input)) {
      return Buffer.from(input).toString('base64');
    }
    return Buffer.from(String(input)).toString('base64');
  }),
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  Charset: { UTF_8: 'UTF_8' },
  sleep: jest.fn()
};

// --- Session ---
global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getEffectiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getScriptTimeZone: jest.fn(() => 'America/Los_Angeles'),
  getTemporaryActiveUserKey: jest.fn(() => 'mock-session-key')
};

// --- PropertiesService ---
let _scriptProperties = {};
let _userProperties = {};
var _scriptPropertiesMock = {
  getProperty: jest.fn(function(key) { return _scriptProperties[key] || null; }),
  setProperty: jest.fn(function(key, value) { _scriptProperties[key] = value; }),
  deleteProperty: jest.fn(function(key) { delete _scriptProperties[key]; }),
  getProperties: jest.fn(function() { return Object.assign({}, _scriptProperties); }),
  setProperties: jest.fn(function(obj) { Object.assign(_scriptProperties, obj); }),
  deleteAllProperties: jest.fn(function() { for (var k in _scriptProperties) delete _scriptProperties[k]; })
};
var _userPropertiesMock = {
  getProperty: jest.fn(function(key) { return _userProperties[key] || null; }),
  setProperty: jest.fn(function(key, value) { _userProperties[key] = value; }),
  deleteProperty: jest.fn(function(key) { delete _userProperties[key]; }),
  getProperties: jest.fn(function() { return Object.assign({}, _userProperties); }),
  setProperties: jest.fn(function(obj) { Object.assign(_userProperties, obj); }),
  deleteAllProperties: jest.fn(function() { for (var k in _userProperties) delete _userProperties[k]; })
};
global.PropertiesService = {
  getScriptProperties: jest.fn(function() { return _scriptPropertiesMock; }),
  getUserProperties: jest.fn(function() { return _userPropertiesMock; }),
  getDocumentProperties: jest.fn(function() { return _scriptPropertiesMock; })
};

// --- CacheService ---
// TEST-03: TTL enforcement — expired entries return null, matching real GAS behavior
let _cache = {};
let _cacheTTLs = {};
var _scriptCacheMock = {
  get: jest.fn(function(key) {
    if (_cacheTTLs[key] && Date.now() > _cacheTTLs[key]) {
      delete _cache[key];
      delete _cacheTTLs[key];
      return null;
    }
    return _cache[key] || null;
  }),
  put: jest.fn(function(key, val, ttl) {
    _cache[key] = val;
    if (ttl) _cacheTTLs[key] = Date.now() + (ttl * 1000);
  }),
  remove: jest.fn(function(key) { delete _cache[key]; delete _cacheTTLs[key]; })
};
global.CacheService = {
  getScriptCache: jest.fn(function() { return _scriptCacheMock; }),
  getUserCache: jest.fn(function() { return _scriptCacheMock; }),
  getDocumentCache: jest.fn(function() { return _scriptCacheMock; })
};

// Reset stores between tests to prevent singleton leaks
afterEach(() => {
  for (const key of Object.keys(_scriptProperties)) delete _scriptProperties[key];
  for (const key of Object.keys(_userProperties)) delete _userProperties[key];
  for (const key of Object.keys(_cache)) delete _cache[key];
  for (const key of Object.keys(_cacheTTLs)) delete _cacheTTLs[key];
  _lockShouldFail = false;
  // Clear mock call history on stable singleton mocks
  Object.values(_scriptPropertiesMock).forEach(fn => { if (fn.mockClear) fn.mockClear(); });
  Object.values(_userPropertiesMock).forEach(fn => { if (fn.mockClear) fn.mockClear(); });
  Object.values(_scriptCacheMock).forEach(fn => { if (fn.mockClear) fn.mockClear(); });
});

// --- Mock Sheet / Range / Spreadsheet ---
function createMockRange(values) {
  return {
    getRow: jest.fn(() => 2),
    getColumn: jest.fn(() => 1),
    getSheet: jest.fn(() => createMockSheet('Mock')),
    getNumColumns: jest.fn(() => 1),
    getNumRows: jest.fn(() => 1),
    getA1Notation: jest.fn(() => 'A2'),
    getValue: jest.fn(() => (values && values[0] && values[0][0]) || ''),
    getValues: jest.fn(() => values || [['']]),
    setValue: jest.fn(),
    setValues: jest.fn(),
    setFontWeight: jest.fn(function() { return this; }),
    setBackground: jest.fn(function() { return this; }),
    setFontColor: jest.fn(function() { return this; }),
    setFontFamily: jest.fn(function() { return this; }),
    setFontSize: jest.fn(function() { return this; }),
    setVerticalAlignment: jest.fn(function() { return this; }),
    setHorizontalAlignment: jest.fn(function() { return this; }),
    setDataValidation: jest.fn(),
    insertCheckboxes: jest.fn()
  };
}

function createMockProtection() {
  return {
    setDescription: jest.fn(function() { return this; }),
    setWarningOnly: jest.fn(function() { return this; }),
    addEditor: jest.fn(function() { return this; }),
    removeEditor: jest.fn(function() { return this; }),
    getEditors: jest.fn(() => []),
    getDescription: jest.fn(() => '')
  };
}

function createMockSheet(name, data) {
  return {
    getName: jest.fn(() => name),
    getDataRange: jest.fn(() => createMockRange(data || [['header']])),
    // TEST-03: Bounds-checked getRange — validates row/col >= 1 like real GAS
    getRange: jest.fn(function(row, col, numRows, numCols) {
      if (typeof row === 'number' && row < 1) throw new Error('Those values are not valid for the row property (minimum 1).');
      if (typeof col === 'number' && col < 1) throw new Error('Those values are not valid for the column property (minimum 1).');
      // Return data slice when data is available and row/col are specified
      if (data && typeof row === 'number' && typeof col === 'number') {
        var nr = numRows || 1;
        var nc = numCols || 1;
        var sliced = [];
        for (var r = row - 1; r < row - 1 + nr; r++) {
          var rowData = [];
          for (var c = col - 1; c < col - 1 + nc; c++) {
            rowData.push(data[r] && data[r][c] !== undefined ? data[r][c] : '');
          }
          sliced.push(rowData);
        }
        return createMockRange(sliced.length > 0 ? sliced : [['']]);
      }
      return createMockRange();
    }),
    getLastRow: jest.fn(() => (data ? data.length : 1)),
    getLastColumn: jest.fn(() => (data && data[0] ? data[0].length : 1)),
    getMaxColumns: jest.fn(() => (data && data[0] ? data[0].length : 10)),
    appendRow: jest.fn(),
    deleteRow: jest.fn(),
    deleteRows: jest.fn(),
    deleteColumns: jest.fn(),
    insertColumnsAfter: jest.fn(),
    hideSheet: jest.fn(),
    showSheet: jest.fn(),
    setFrozenRows: jest.fn(),
    setColumnWidth: jest.fn(),
    setTabColor: jest.fn(),
    clear: jest.fn(),
    protect: jest.fn(() => createMockProtection()),
    getProtections: jest.fn(() => []),
    setName: jest.fn(),
    copyTo: jest.fn(),
    getParent: jest.fn(function() { return createMockSpreadsheet([]); }),
    getIndex: jest.fn(function() { return 1; }),
    insertRowsAfter: jest.fn(),
    getFilter: jest.fn(function() { return null; })
  };
}

function createMockSpreadsheet(sheets) {
  const sheetMap = {};
  (sheets || []).forEach(s => { sheetMap[s.getName()] = s; });

  return {
    getSheetByName: jest.fn(name => {
      if (sheetMap[name]) return sheetMap[name];
      // Fallback: check SHEET_LEGACY_NAMES_ for old→new mapping
      if (typeof SHEET_LEGACY_NAMES_ !== 'undefined') {
        var keys = Object.keys(SHEET_LEGACY_NAMES_);
        for (var i = 0; i < keys.length; i++) {
          if (SHEET_LEGACY_NAMES_[keys[i]] === name && sheetMap[keys[i]]) {
            return sheetMap[keys[i]];
          }
        }
      }
      return null;
    }),
    getSheets: jest.fn(() => sheets || []),
    toast: jest.fn(),
    insertSheet: jest.fn(name => {
      const s = createMockSheet(name);
      sheetMap[name] = s;
      return s;
    }),
    getActiveSheet: jest.fn(() => (sheets && sheets[0]) || createMockSheet('Sheet1')),
    setActiveSheet: jest.fn(),
    moveActiveSheet: jest.fn(),
    getOwner: jest.fn(() => ({ getEmail: jest.fn(() => 'owner@example.com') })),
    getSpreadsheetTimeZone: jest.fn(() => 'America/New_York')
  };
}

global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => createMockSpreadsheet()),
  getActiveSheet: jest.fn(() => createMockSheet('Test')),
  getUi: jest.fn(() => ({
    alert: jest.fn(),
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    Button: { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL' },
    createMenu: jest.fn(() => ({
      addItem: jest.fn(function() { return this; }),
      addSubMenu: jest.fn(function() { return this; }),
      addSeparator: jest.fn(function() { return this; }),
      addToUi: jest.fn()
    }))
  })),
  newDataValidation: jest.fn(() => ({
    requireValueInList: jest.fn(function() { return this; }),
    setAllowInvalid: jest.fn(function() { return this; }),
    setHelpText: jest.fn(function() { return this; }),
    build: jest.fn(() => ({}))
  })),
  ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' }
};

// --- HtmlService ---
global.HtmlService = {
  createHtmlOutput: jest.fn(html => ({
    setWidth: jest.fn(function() { return this; }),
    setHeight: jest.fn(function() { return this; }),
    setTitle: jest.fn(function() { return this; }),
    setXFrameOptionsMode: jest.fn(function() { return this; }),
    getContent: jest.fn(() => html)
  })),
  createTemplateFromFile: jest.fn(name => ({
    view: '',
    pageData: '',
    evaluate: jest.fn(function() {
      return {
        setTitle: jest.fn(function() { return this; }),
        setXFrameOptionsMode: jest.fn(function() { return this; }),
        addMetaTag: jest.fn(function() { return this; }),
        getContent: jest.fn(() => '<html></html>')
      };
    })
  })),
  createHtmlOutputFromFile: jest.fn(name => ({
    getContent: jest.fn(() => '<html>' + name + '</html>')
  })),
  XFrameOptionsMode: { DEFAULT: 'DEFAULT', ALLOWALL: 'ALLOWALL' },
  SandboxMode: { IFRAME: 'IFRAME' }
};

// --- MailApp ---
global.MailApp = {
  sendEmail: jest.fn(),
  getRemainingDailyQuota: jest.fn(() => 100)
};

// --- GmailApp ---
global.GmailApp = {
  sendEmail: jest.fn(),
  search: jest.fn(() => []),
  createDraft: jest.fn()
};

// --- DriveApp ---
global.DriveApp = {
  createFolder: jest.fn(() => ({
    getId: jest.fn(() => 'folder-id'),
    getUrl: jest.fn(() => 'https://drive.google.com/folder-id'),
    setDescription: jest.fn(),
    setSharing: jest.fn(),
    getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
    createFolder: jest.fn(),
    getFiles: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
    createFile: jest.fn()
  })),
  getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
  getRootFolder: jest.fn(() => ({ getId: jest.fn(() => 'root') })),
  getFileById: jest.fn(() => ({ getName: jest.fn(() => 'TestFile.pdf') })),
  // SEC-07: Access and Permission enums for folder sharing control
  Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', DOMAIN: 'DOMAIN', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', PRIVATE: 'PRIVATE' },
  Permission: { VIEW: 'VIEW', EDIT: 'EDIT', COMMENT: 'COMMENT', NONE: 'NONE' }
};

// --- MimeType ---
global.MimeType = { CSV: 'text/csv', PDF: 'application/pdf' };

// --- DocumentApp ---
global.DocumentApp = {
  create: jest.fn(() => ({
    getBody: jest.fn(() => ({
      appendParagraph: jest.fn(),
      appendTable: jest.fn()
    })),
    getUrl: jest.fn(() => 'https://docs.google.com/mock-doc'),
    getId: jest.fn(() => 'mock-doc-id')
  }))
};

// --- CalendarApp ---
global.CalendarApp = {
  getAllCalendars: jest.fn(() => []),
  getCalendarsByName: jest.fn(() => []),
  getDefaultCalendar: jest.fn(() => ({
    getName: jest.fn(() => 'Default Calendar'),
    createEvent: jest.fn(),
    getEvents: jest.fn(() => [])
  })),
  createCalendar: jest.fn(() => ({
    getName: jest.fn(() => 'Test Calendar'),
    createEvent: jest.fn()
  }))
};

// --- ScriptApp ---
global.ScriptApp = {
  getProjectTriggers: jest.fn(() => []),
  deleteTrigger: jest.fn(),
  newTrigger: jest.fn(() => ({
    timeBased: jest.fn(function() { return this; }),
    atHour: jest.fn(function() { return this; }),
    everyDays: jest.fn(function() { return this; }),
    onWeekDay: jest.fn(function() { return this; }),
    create: jest.fn()
  })),
  getService: jest.fn(() => ({
    getUrl: jest.fn(() => 'https://script.google.com/macros/s/test/exec')
  })),
  WeekDay: { SUNDAY: 1, MONDAY: 2, TUESDAY: 3, WEDNESDAY: 4, THURSDAY: 5, FRIDAY: 6, SATURDAY: 7 }
};

// --- LockService ---
let _lockShouldFail = false;
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => !_lockShouldFail),
    waitLock: jest.fn(() => { if (_lockShouldFail) throw new Error('Could not obtain lock'); }),
    releaseLock: jest.fn()
  })),
  _mockLockFailure: function() { _lockShouldFail = true; },
  _mockLockSuccess: function() { _lockShouldFail = false; }
};

// --- UrlFetchApp ---
global.UrlFetchApp = {
  fetch: jest.fn(() => ({
    getResponseCode: jest.fn(() => 200),
    getContentText: jest.fn(() => '{}')
  }))
};

// --- log_ (defined in 01_Core.gs) ---
// Available as a global so test files that don't load 01_Core.gs can still
// use files that call log_() without a ReferenceError.
global.log_ = function(context, message, level) {
  Logger.log('[' + (level || 'INFO') + '] ' + context + ': ' + message);
};

// --- Auth helpers (defined in 21_WebDashDataService.gs) ---
// Mocked globally so test files that don't load WebDashDataService can still
// test wrapper functions in 26_QAForum.gs, 27_TimelineService.gs, 28_FailsafeService.gs.
global._resolveCallerEmail = jest.fn(() => 'test@example.com');
global._requireStewardAuth = jest.fn(() => 'steward@example.com');

// --- Retry helper (defined in 06_Maintenance.gs) ---
// In tests, executeWithRetry just calls the function directly (no retry delay).
global.executeWithRetry = jest.fn(function (fn) { return fn(); });

module.exports = { createMockRange, createMockSheet, createMockSpreadsheet, createMockProtection };
