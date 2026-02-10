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
  formatDate: jest.fn((date, tz, fmt) => {
    if (fmt === 'yyyy-MM-dd') {
      return date.toISOString().slice(0, 10);
    }
    return date.toISOString();
  }),
  computeDigest: jest.fn((algo, data, charset) => {
    const hash = new Array(32).fill(0);
    for (let i = 0; i < data.length; i++) {
      hash[i % 32] = (hash[i % 32] + data.charCodeAt(i) + i) & 0xFF;
    }
    return hash;
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
  getScriptTimeZone: jest.fn(() => 'America/Los_Angeles')
};

// --- PropertiesService ---
const _scriptProperties = {};
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(key => _scriptProperties[key] || null),
    setProperty: jest.fn((key, val) => { _scriptProperties[key] = val; }),
    deleteProperty: jest.fn(key => { delete _scriptProperties[key]; })
  }))
};

// --- CacheService ---
const _cache = {};
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn(key => _cache[key] || null),
    put: jest.fn((key, val, ttl) => { _cache[key] = val; }),
    remove: jest.fn(key => { delete _cache[key]; })
  }))
};

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

function createMockSheet(name, data) {
  return {
    getName: jest.fn(() => name),
    getDataRange: jest.fn(() => createMockRange(data || [['header']])),
    getRange: jest.fn(() => createMockRange()),
    getLastRow: jest.fn(() => (data ? data.length : 1)),
    getLastColumn: jest.fn(() => (data && data[0] ? data[0].length : 1)),
    appendRow: jest.fn(),
    deleteRow: jest.fn(),
    deleteRows: jest.fn(),
    hideSheet: jest.fn(),
    showSheet: jest.fn(),
    setFrozenRows: jest.fn(),
    setColumnWidth: jest.fn(),
    clear: jest.fn()
  };
}

function createMockSpreadsheet(sheets) {
  const sheetMap = {};
  (sheets || []).forEach(s => { sheetMap[s.getName()] = s; });

  return {
    getSheetByName: jest.fn(name => sheetMap[name] || null),
    getSheets: jest.fn(() => sheets || []),
    toast: jest.fn(),
    insertSheet: jest.fn(name => {
      const s = createMockSheet(name);
      sheetMap[name] = s;
      return s;
    }),
    setActiveSheet: jest.fn(),
    moveActiveSheet: jest.fn(),
    getOwner: jest.fn(() => ({ getEmail: jest.fn(() => 'owner@example.com') }))
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
    build: jest.fn(() => ({}))
  }))
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
  XFrameOptionsMode: { DENY: 'DENY', ALLOWALL: 'ALLOWALL' }
};

// --- MailApp ---
global.MailApp = { sendEmail: jest.fn() };

// --- DriveApp ---
global.DriveApp = {
  createFolder: jest.fn(() => ({
    getId: jest.fn(() => 'folder-id'),
    getUrl: jest.fn(() => 'https://drive.google.com/folder-id'),
    setDescription: jest.fn(),
    getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
    createFolder: jest.fn()
  })),
  getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
  getRootFolder: jest.fn(() => ({ getId: jest.fn(() => 'root') }))
};

// --- CalendarApp ---
global.CalendarApp = {
  getAllCalendars: jest.fn(() => []),
  getCalendarsByName: jest.fn(() => []),
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
    create: jest.fn()
  }))
};

// --- LockService ---
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn()
  }))
};

module.exports = { createMockRange, createMockSheet, createMockSpreadsheet };
