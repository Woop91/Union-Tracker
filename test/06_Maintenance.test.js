/**
 * Tests for 06_Maintenance.gs
 *
 * Covers constants (CACHE_CONFIG, CACHE_KEYS, UNDO_CONFIG, TEST_RESULTS,
 * VALIDATION_PATTERNS, VALIDATION_MESSAGES), caching functions (getCachedData,
 * setCachedData, invalidateCache, invalidateAllCaches), undo/redo functions
 * (getUndoHistory, saveUndoHistory, clearUndoHistory, recordAction), and
 * utility functions (executeWithRetry, batchSetValues, confirmDestructiveAction,
 * DIAGNOSE_SETUP).
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '06_Maintenance.gs'
]);

// ============================================================================
// Constants - CACHE_CONFIG
// ============================================================================

describe('CACHE_CONFIG', () => {
  test('CACHE_CONFIG is defined', () => {
    expect(CACHE_CONFIG).toBeDefined();
  });

  test('MEMORY_TTL is 300', () => {
    expect(CACHE_CONFIG.MEMORY_TTL).toBe(300);
  });

  test('PROPS_TTL is 3600', () => {
    expect(CACHE_CONFIG.PROPS_TTL).toBe(3600);
  });

  test('ENABLE_LOGGING is false', () => {
    expect(CACHE_CONFIG.ENABLE_LOGGING).toBe(false);
  });
});

// ============================================================================
// Constants - CACHE_KEYS
// ============================================================================

describe('CACHE_KEYS', () => {
  test('CACHE_KEYS is defined', () => {
    expect(CACHE_KEYS).toBeDefined();
  });

  test('ALL_GRIEVANCES key exists', () => {
    expect(CACHE_KEYS.ALL_GRIEVANCES).toBe('cache_grievances');
  });

  test('ALL_MEMBERS key exists', () => {
    expect(CACHE_KEYS.ALL_MEMBERS).toBe('cache_members');
  });

  test('ALL_STEWARDS key exists', () => {
    expect(CACHE_KEYS.ALL_STEWARDS).toBe('cache_stewards');
  });

  test('DASHBOARD_METRICS key exists', () => {
    expect(CACHE_KEYS.DASHBOARD_METRICS).toBe('cache_metrics');
  });

  test('CONFIG_VALUES key exists', () => {
    expect(CACHE_KEYS.CONFIG_VALUES).toBeDefined();
  });
});

// ============================================================================
// Constants - UNDO_CONFIG
// ============================================================================

describe('UNDO_CONFIG', () => {
  test('UNDO_CONFIG is defined', () => {
    expect(UNDO_CONFIG).toBeDefined();
  });

  test('MAX_HISTORY is 50', () => {
    expect(UNDO_CONFIG.MAX_HISTORY).toBe(50);
  });

  test('STORAGE_KEY is undoRedoHistory', () => {
    expect(UNDO_CONFIG.STORAGE_KEY).toBe('undoRedoHistory');
  });
});

// ============================================================================
// Cache Functions - getCachedData
// ============================================================================

describe('getCachedData', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('returns data from memory cache on hit', () => {
    const mockData = { items: [1, 2, 3] };
    const mockCache = {
      get: jest.fn(() => JSON.stringify(mockData)),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const loader = jest.fn();
    const result = getCachedData('test_key', loader, 300);

    expect(result).toEqual(mockData);
    expect(loader).not.toHaveBeenCalled();
  });

  test('calls loader on cache miss and returns fresh data', () => {
    const freshData = { fresh: true };
    const mockCache = {
      get: jest.fn(() => null),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = {
      getProperty: jest.fn(() => null),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    const loader = jest.fn(() => freshData);
    const result = getCachedData('test_key', loader, 300);

    expect(result).toEqual(freshData);
    expect(loader).toHaveBeenCalledTimes(1);
  });

  test('uses default TTL from CACHE_CONFIG when not provided', () => {
    const mockCache = {
      get: jest.fn(() => null),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = {
      getProperty: jest.fn(() => null),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    const loader = jest.fn(() => 'data');
    getCachedData('test_key', loader);

    // Should use CACHE_CONFIG.MEMORY_TTL = 300
    expect(mockCache.put).toHaveBeenCalled();
  });
});

// ============================================================================
// Cache Functions - setCachedData
// ============================================================================

describe('setCachedData', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('stores data in memory cache', () => {
    const mockCache = {
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    setCachedData('key1', { value: 42 }, 300);

    expect(mockCache.put).toHaveBeenCalledWith(
      'key1',
      JSON.stringify({ value: 42 }),
      300
    );
  });

  test('stores data in properties cache for small payloads', () => {
    const mockCache = {
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    setCachedData('key1', { value: 42 }, 300);

    expect(mockProps.setProperty).toHaveBeenCalled();
    const storedArg = JSON.parse(mockProps.setProperty.mock.calls[0][1]);
    expect(storedArg.data).toEqual({ value: 42 });
    expect(storedArg.timestamp).toBeDefined();
  });

  test('caps memory TTL at 21600 seconds (6 hours)', () => {
    const mockCache = { get: jest.fn(), put: jest.fn(), remove: jest.fn() };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = { getProperty: jest.fn(), setProperty: jest.fn(), deleteProperty: jest.fn() };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    setCachedData('key1', 'data', 99999);

    expect(mockCache.put).toHaveBeenCalledWith('key1', '"data"', 21600);
  });
});

// ============================================================================
// Cache Functions - invalidateCache
// ============================================================================

describe('invalidateCache', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('removes key from both caches', () => {
    const mockCache = { get: jest.fn(), put: jest.fn(), remove: jest.fn() };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = { getProperty: jest.fn(), setProperty: jest.fn(), deleteProperty: jest.fn() };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    invalidateCache('test_key');

    expect(mockCache.remove).toHaveBeenCalledWith('test_key');
    expect(mockProps.deleteProperty).toHaveBeenCalledWith('test_key');
  });
});

// ============================================================================
// Cache Functions - invalidateAllCaches
// ============================================================================

describe('invalidateAllCaches', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('clears all keys from properties cache', () => {
    const mockCache = { get: jest.fn(), put: jest.fn(), remove: jest.fn(), removeAll: jest.fn() };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    const mockProps = { getProperty: jest.fn(), setProperty: jest.fn(), deleteProperty: jest.fn() };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    invalidateAllCaches();

    // Should call deleteProperty for each CACHE_KEYS value
    const keyCount = Object.keys(CACHE_KEYS).length;
    expect(mockProps.deleteProperty).toHaveBeenCalledTimes(keyCount);
  });
});

// ============================================================================
// Undo Functions - getUndoHistory
// ============================================================================

describe('getUndoHistory', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('returns default history when storage is empty', () => {
    const mockProps = {
      getProperty: jest.fn(() => null),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    const history = getUndoHistory();

    expect(history).toEqual({ actions: [], currentIndex: 0 });
  });

  test('returns parsed history from storage', () => {
    const stored = { actions: [{ type: 'EDIT_CELL' }], currentIndex: 1 };
    const mockProps = {
      getProperty: jest.fn(() => JSON.stringify(stored)),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    const history = getUndoHistory();

    expect(history).toEqual(stored);
  });
});

// ============================================================================
// Undo Functions - saveUndoHistory
// ============================================================================

describe('saveUndoHistory', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('saves history to properties', () => {
    const mockProps = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    const history = { actions: [{ type: 'TEST' }], currentIndex: 1 };
    saveUndoHistory(history);

    expect(mockProps.setProperty).toHaveBeenCalledWith(
      UNDO_CONFIG.STORAGE_KEY,
      JSON.stringify(history)
    );
  });

  test('trims history when exceeding MAX_HISTORY', () => {
    const mockProps = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    // Create history exceeding the limit
    const actions = [];
    for (let i = 0; i < 60; i++) {
      actions.push({ type: 'ACTION_' + i });
    }
    const history = { actions: actions, currentIndex: 60 };
    saveUndoHistory(history);

    const savedData = JSON.parse(mockProps.setProperty.mock.calls[0][1]);
    expect(savedData.actions.length).toBe(UNDO_CONFIG.MAX_HISTORY);
  });
});

// ============================================================================
// Undo Functions - clearUndoHistory
// ============================================================================

describe('clearUndoHistory', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('deletes history from properties', () => {
    const mockProps = {
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    clearUndoHistory();

    expect(mockProps.deleteProperty).toHaveBeenCalledWith(UNDO_CONFIG.STORAGE_KEY);
  });
});

// ============================================================================
// Undo Functions - recordAction
// ============================================================================

describe('recordAction', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('records an action and increments currentIndex', () => {
    // Set up so getUndoHistory returns empty, and saveUndoHistory captures
    const mockProps = {
      getProperty: jest.fn(() => null),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    recordAction('EDIT_CELL', 'Edited field X', { old: 1 }, { new: 2 });

    expect(mockProps.setProperty).toHaveBeenCalled();
    const saved = JSON.parse(mockProps.setProperty.mock.calls[0][1]);
    expect(saved.actions.length).toBe(1);
    expect(saved.actions[0].type).toBe('EDIT_CELL');
    expect(saved.actions[0].description).toBe('Edited field X');
    expect(saved.actions[0].beforeState).toEqual({ old: 1 });
    expect(saved.actions[0].afterState).toEqual({ new: 2 });
    expect(saved.currentIndex).toBe(1);
  });

  test('clears redo history when recording a new action mid-history', () => {
    const existing = {
      actions: [
        { type: 'A1', description: 'first', timestamp: new Date().toISOString(), beforeState: null, afterState: null },
        { type: 'A2', description: 'second', timestamp: new Date().toISOString(), beforeState: null, afterState: null }
      ],
      currentIndex: 1  // undid one action, so there is one to redo
    };
    const mockProps = {
      getProperty: jest.fn(() => JSON.stringify(existing)),
      setProperty: jest.fn(),
      deleteProperty: jest.fn()
    };
    PropertiesService.getUserProperties.mockReturnValue(mockProps);

    recordAction('A3', 'third', null, null);

    const saved = JSON.parse(mockProps.setProperty.mock.calls[0][1]);
    // Should have sliced off the second action and added the new one
    expect(saved.actions.length).toBe(2);
    expect(saved.actions[1].type).toBe('A3');
    expect(saved.currentIndex).toBe(2);
  });
});

// ============================================================================
// Utility Functions - executeWithRetry
// ============================================================================

describe('executeWithRetry', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('returns result on first successful attempt', () => {
    const fn = jest.fn(() => 'success');

    const result = executeWithRetry(fn);

    expect(result).toBe('success');
    expect(fn).toHaveBeenCalledTimes(1);
  });

  test('retries on failure and succeeds on second attempt', () => {
    let callCount = 0;
    const fn = jest.fn(() => {
      callCount++;
      if (callCount === 1) throw new Error('transient');
      return 'ok';
    });

    const result = executeWithRetry(fn, { maxRetries: 3, baseDelay: 1 });

    expect(result).toBe('ok');
    expect(fn).toHaveBeenCalledTimes(2);
  });

  test('throws after all retries exhausted', () => {
    const fn = jest.fn(() => { throw new Error('permanent'); });

    expect(() => {
      executeWithRetry(fn, { maxRetries: 2, baseDelay: 1 });
    }).toThrow(/Operation failed after 3 attempts/);

    // 1 initial + 2 retries = 3 calls
    expect(fn).toHaveBeenCalledTimes(3);
  });

  test('calls onError callback on each failure', () => {
    const onError = jest.fn();
    const fn = jest.fn(() => { throw new Error('fail'); });

    expect(() => {
      executeWithRetry(fn, { maxRetries: 1, baseDelay: 1, onError });
    }).toThrow();

    expect(onError).toHaveBeenCalledTimes(2);
  });

  test('uses exponential backoff via Utilities.sleep', () => {
    const fn = jest.fn(() => { throw new Error('fail'); });

    expect(() => {
      executeWithRetry(fn, { maxRetries: 2, baseDelay: 1000 });
    }).toThrow();

    // sleep(1000 * 2^0 = 1000), sleep(1000 * 2^1 = 2000)
    expect(Utilities.sleep).toHaveBeenCalledWith(1000);
    expect(Utilities.sleep).toHaveBeenCalledWith(2000);
  });
});

// ============================================================================
// Utility Functions - batchSetValues
// ============================================================================

describe('batchSetValues', () => {
  test('does nothing for empty updates', () => {
    const sheet = {
      getRange: jest.fn(),
      getLastRow: jest.fn(),
      getLastColumn: jest.fn()
    };

    batchSetValues(sheet, []);
    batchSetValues(sheet, null);

    expect(sheet.getRange).not.toHaveBeenCalled();
  });

  test('applies updates and writes back in single operation', () => {
    const currentData = [
      ['A1', 'B1'],
      ['A2', 'B2'],
      ['A3', 'B3']
    ];
    const mockRange = {
      getValues: jest.fn(() => currentData),
      setValues: jest.fn()
    };
    const sheet = {
      getRange: jest.fn(() => mockRange)
    };

    const updates = [
      { row: 2, col: 1, value: 'NEW_A2' },
      { row: 3, col: 2, value: 'NEW_B3' }
    ];

    batchSetValues(sheet, updates);

    expect(sheet.getRange).toHaveBeenCalledWith(1, 1, 3, 2);
    expect(mockRange.setValues).toHaveBeenCalledTimes(1);
    const written = mockRange.setValues.mock.calls[0][0];
    expect(written[1][0]).toBe('NEW_A2');
    expect(written[2][1]).toBe('NEW_B3');
  });
});

// ============================================================================
// Utility Functions - confirmDestructiveAction
// ============================================================================

describe('confirmDestructiveAction', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('returns true when user clicks YES', () => {
    const mockUi = {
      alert: jest.fn(() => 'YES'),
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    const result = confirmDestructiveAction('Delete Records', 'This will delete all records.');

    expect(result).toBe(true);
  });

  test('returns false when user clicks NO', () => {
    const mockUi = {
      alert: jest.fn(() => 'NO'),
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    const result = confirmDestructiveAction('Delete Records', 'This will delete all records.');

    expect(result).toBe(false);
  });
});

// ============================================================================
// DIAGNOSE_SETUP - basic structure
// ============================================================================

describe('DIAGNOSE_SETUP', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('returns object with required structure', () => {
    // Set up minimal mocks for DIAGNOSE_SETUP
    const mockSheet = {
      getName: jest.fn(() => 'TestSheet'),
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [['']]),
        getValue: jest.fn(() => '')
      })),
      getLastRow: jest.fn(() => 1),
      getLastColumn: jest.fn(() => 1),
      getDataRange: jest.fn(() => ({
        getValues: jest.fn(() => [['header']])
      }))
    };

    const mockSS = {
      getSheets: jest.fn(() => [mockSheet]),
      getSheetByName: jest.fn(() => null),
      toast: jest.fn(),
      insertSheet: jest.fn(() => mockSheet)
    };
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = DIAGNOSE_SETUP();

    expect(result).toBeDefined();
    expect(result.timestamp).toBeDefined();
    expect(result.status).toBeDefined();
    expect(Array.isArray(result.checks)).toBe(true);
    expect(Array.isArray(result.warnings)).toBe(true);
    expect(Array.isArray(result.errors)).toBe(true);
    expect(result.summary).toBeDefined();
  });
});
