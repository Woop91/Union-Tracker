/**
 * ============================================================================
 * PERFORMANCE CACHING & UNDO/REDO SYSTEM
 * ============================================================================
 * Data caching for performance + action history
 */

// ==================== CACHE CONFIGURATION ====================

var CACHE_CONFIG = {
  MEMORY_TTL: 300,
  PROPS_TTL: 3600,
  ENABLE_LOGGING: false
};

var CACHE_KEYS = {
  ALL_GRIEVANCES: 'cache_grievances',
  ALL_MEMBERS: 'cache_members',
  ALL_STEWARDS: 'cache_stewards',
  DASHBOARD_METRICS: 'cache_metrics'
};

// ==================== CACHING FUNCTIONS ====================

function getCachedData(key, loader, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;
  try {
    var memCache = CacheService.getScriptCache();
    var cached = memCache.get(key);
    if (cached) { if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT] ' + key); return JSON.parse(cached); }
    var propsCache = PropertiesService.getScriptProperties();
    var propsCached = propsCache.getProperty(key);
    if (propsCached) {
      var obj = JSON.parse(propsCached);
      if (obj.timestamp && (Date.now() - obj.timestamp) < (ttl * 1000)) {
        memCache.put(key, JSON.stringify(obj.data), ttl);
        return obj.data;
      }
    }
    if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE MISS] ' + key);
    var data = loader();
    setCachedData(key, data, ttl);
    return data;
  } catch (e) { Logger.log('Cache error: ' + e.message); return loader(); }
}

function setCachedData(key, data, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;
  try {
    var str = JSON.stringify(data);
    var memCache = CacheService.getScriptCache();
    memCache.put(key, str, Math.min(ttl, 21600));
    if (str.length < 400000) {
      var propsCache = PropertiesService.getScriptProperties();
      propsCache.setProperty(key, JSON.stringify({ data: data, timestamp: Date.now() }));
    }
  } catch (e) { Logger.log('Set cache error: ' + e.message); }
}

function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
    PropertiesService.getScriptProperties().deleteProperty(key);
    Logger.log('Cache invalidated: ' + key);
  } catch (e) { Logger.log('Invalidate error: ' + e.message); }
}

function invalidateAllCaches() {
  try {
    var keys = Object.keys(CACHE_KEYS).map(function(k) { return CACHE_KEYS[k]; });
    CacheService.getScriptCache().removeAll(keys);
    var props = PropertiesService.getScriptProperties();
    keys.forEach(function(k) { props.deleteProperty(k); });
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ All caches cleared', 'Cache', 3);
  } catch (e) { Logger.log('Clear all error: ' + e.message); }
}

function warmUpCaches() {
  SpreadsheetApp.getActiveSpreadsheet().toast('🔥 Warming caches...', 'Cache', -1);
  try {
    getCachedGrievances();
    getCachedMembers();
    getCachedStewards();
    getCachedDashboardMetrics();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Caches warmed', 'Cache', 3);
  } catch (e) { Logger.log('Warmup error: ' + e.message); }
}

function getCachedGrievances() {
  return getCachedData(CACHE_KEYS.ALL_GRIEVANCES, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 300);
}

function getCachedMembers() {
  return getCachedData(CACHE_KEYS.ALL_MEMBERS, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 600);
}

function getCachedStewards() {
  return getCachedData(CACHE_KEYS.ALL_STEWARDS, function() {
    var members = getCachedMembers();
    return members.filter(function(row) { return row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes'; });
  }, 600);
}

function getCachedDashboardMetrics() {
  return getCachedData(CACHE_KEYS.DASHBOARD_METRICS, function() {
    var grievances = getCachedGrievances();
    var metrics = { total: grievances.length, open: 0, closed: 0, overdue: 0, byStatus: {}, byIssueType: {}, bySteward: {} };
    var today = new Date(); today.setHours(0, 0, 0, 0);
    grievances.forEach(function(row) {
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var issue = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
      var steward = row[GRIEVANCE_COLS.STEWARD - 1];
      var deadline = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysTo = deadline ? Math.floor((new Date(deadline) - today) / (1000 * 60 * 60 * 24)) : null;
      if (status === 'Open') metrics.open++;
      if (status === 'Closed' || status === 'Resolved') metrics.closed++;
      if (daysTo !== null && daysTo < 0) metrics.overdue++;
      metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;
      if (issue) metrics.byIssueType[issue] = (metrics.byIssueType[issue] || 0) + 1;
      if (steward) metrics.bySteward[steward] = (metrics.bySteward[steward] || 0) + 1;
    });
    return metrics;
  }, 180);
}

function showCacheStatusDashboard() {
  var memCache = CacheService.getScriptCache();
  var propsCache = PropertiesService.getScriptProperties();
  var rows = Object.keys(CACHE_KEYS).map(function(name) {
    var key = CACHE_KEYS[name];
    var inMem = memCache.get(key) !== null;
    var inProps = propsCache.getProperty(key) !== null;
    var age = 'N/A';
    if (inProps) {
      try { var obj = JSON.parse(propsCache.getProperty(key)); if (obj.timestamp) age = Math.floor((Date.now() - obj.timestamp) / 1000) + 's'; } catch (e) {}
    }
    return '<tr><td>' + name + '</td><td>' + (inMem ? '✅' : '❌') + '</td><td>' + (inProps ? '✅' : '❌') + '</td><td>' + age + '</td></tr>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}table{width:100%;border-collapse:collapse;margin:20px 0}th{background:#1a73e8;color:white;padding:12px;text-align:left}td{padding:10px;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:5px}button.danger{background:#dc3545}</style></head><body><div class="container"><h2>🗄️ Cache Status</h2><table><tr><th>Cache</th><th>Memory</th><th>Props</th><th>Age</th></tr>' + rows + '</table><button onclick="google.script.run.withSuccessHandler(function(){location.reload()}).warmUpCaches()">🔥 Warm Up</button><button class="danger" onclick="google.script.run.withSuccessHandler(function(){location.reload()}).invalidateAllCaches()">🗑️ Clear All</button></div></body></html>'
  ).setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '🗄️ Cache Status');
}

// ==================== UNDO/REDO SYSTEM ====================

var UNDO_CONFIG = { MAX_HISTORY: 50, STORAGE_KEY: 'undoRedoHistory' };

function getUndoHistory() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty(UNDO_CONFIG.STORAGE_KEY);
  if (json) return JSON.parse(json);
  return { actions: [], currentIndex: 0 };
}

function saveUndoHistory(history) {
  var props = PropertiesService.getScriptProperties();
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }
  props.setProperty(UNDO_CONFIG.STORAGE_KEY, JSON.stringify(history));
}

function recordAction(type, description, beforeState, afterState) {
  var history = getUndoHistory();
  if (history.currentIndex < history.actions.length) history.actions = history.actions.slice(0, history.currentIndex);
  history.actions.push({ type: type, description: description, timestamp: new Date().toISOString(), beforeState: beforeState, afterState: afterState });
  history.currentIndex = history.actions.length;
  saveUndoHistory(history);
}

function recordCellEdit(row, col, oldValue, newValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colName = sheet.getRange(1, col).getValue();
  recordAction('EDIT_CELL', 'Edited ' + colName + ' in row ' + row, { row: row, col: col, value: oldValue, sheet: sheet.getName() }, { row: row, col: col, value: newValue, sheet: sheet.getName() });
}

function recordRowAddition(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  recordAction('ADD_ROW', 'Added row ' + row, null, { row: row, data: rowData, sheet: sheet.getName() });
}

function recordRowDeletion(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  recordAction('DELETE_ROW', 'Deleted row ' + row, { row: row, data: rowData, sheet: sheet.getName() }, null);
}

function undoLastAction() {
  var history = getUndoHistory();
  if (history.currentIndex === 0) throw new Error('Nothing to undo');
  var action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);
  history.currentIndex--;
  saveUndoHistory(history);
  SpreadsheetApp.getActiveSpreadsheet().toast('↩️ Undone: ' + action.description, 'Undo', 3);
}

function redoLastAction() {
  var history = getUndoHistory();
  if (history.currentIndex >= history.actions.length) throw new Error('Nothing to redo');
  var action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);
  history.currentIndex++;
  saveUndoHistory(history);
  SpreadsheetApp.getActiveSpreadsheet().toast('↪️ Redone: ' + action.description, 'Redo', 3);
}

function undoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex > targetIndex) undoLastAction();
}

function redoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) redoLastAction();
}

function applyState(state, actionType) {
  if (!state) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(state.sheet);
  if (!sheet) throw new Error('Sheet ' + state.sheet + ' not found');
  switch (actionType) {
    case 'EDIT_CELL': sheet.getRange(state.row, state.col).setValue(state.value); break;
    case 'ADD_ROW': if (state.row) sheet.deleteRow(state.row); break;
    case 'DELETE_ROW': if (state.row && state.data) { sheet.insertRowAfter(state.row - 1); sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]); } break;
    case 'BATCH_UPDATE': if (state.changes) state.changes.forEach(function(c) { sheet.getRange(c.row, c.col).setValue(c.oldValue); }); break;
  }
}

function clearUndoHistory() {
  PropertiesService.getScriptProperties().deleteProperty(UNDO_CONFIG.STORAGE_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ History cleared', 'Undo/Redo', 3);
}

function exportUndoHistoryToSheet() {
  var history = getUndoHistory();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Undo_History_Export');
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet('Undo_History_Export');
  var headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  if (history.actions.length > 0) {
    var rows = history.actions.map(function(a, i) { return [i + 1, a.type, a.description, new Date(a.timestamp).toLocaleString(), i < history.currentIndex ? 'Applied' : 'Undone']; });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  for (var c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);
  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

function showUndoRedoPanel() {
  var history = getUndoHistory();
  var rows = history.actions.length === 0 ? '<tr><td colspan="5" style="text-align:center;padding:40px;color:#999">No actions recorded</td></tr>' :
    history.actions.slice().reverse().map(function(a, i) {
      var idx = history.actions.length - i;
      var time = new Date(a.timestamp).toLocaleString();
      var canUndo = i < history.actions.length - history.currentIndex;
      return '<tr><td>' + idx + '</td><td><span class="badge ' + a.type.toLowerCase() + '">' + a.type + '</span></td><td>' + a.description + '</td><td style="font-size:12px;color:#666">' + time + '</td><td>' + (canUndo ? '<button onclick="undo(' + (idx - 1) + ')">↩️</button>' : '') + '</td></tr>';
    }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin:20px 0}.stat{background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;border-left:4px solid #1a73e8}.num{font-size:32px;font-weight:bold;color:#1a73e8}table{width:100%;border-collapse:collapse;margin:20px 0}th{background:#1a73e8;color:white;padding:12px;text-align:left}td{padding:12px;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;margin:2px}.badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:11px;font-weight:bold}.edit_cell{background:#e3f2fd;color:#1976d2}.add_row{background:#e8f5e9;color:#388e3c}.delete_row{background:#ffebee;color:#d32f2f}.quick{margin:20px 0;padding:20px;background:#f8f9fa;border-radius:8px}</style></head><body><div class="container"><h2>↩️ Undo/Redo History</h2><div class="stats"><div class="stat"><div class="num">' + history.actions.length + '</div><div>Total Actions</div></div><div class="stat"><div class="num">' + history.currentIndex + '</div><div>Current Position</div></div><div class="stat"><div class="num">' + (history.actions.length - history.currentIndex) + '</div><div>Available</div></div></div><div class="quick"><button onclick="performUndo()">↩️ Undo Last</button><button onclick="performRedo()">↪️ Redo</button><button onclick="clear()" style="background:#d32f2f">🗑️ Clear</button><button onclick="exp()" style="background:#00796b">📥 Export</button></div><table><tr><th>#</th><th>Type</th><th>Description</th><th>Time</th><th></th></tr>' + rows + '</table></div><script>function performUndo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("❌ "+e.message)}).undoLastAction()}function performRedo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("❌ "+e.message)}).redoLastAction()}function undo(i){google.script.run.withSuccessHandler(function(){location.reload()}).undoToIndex(i)}function clear(){if(confirm("Clear all history?")){google.script.run.withSuccessHandler(function(){location.reload()}).clearUndoHistory()}}function exp(){google.script.run.withSuccessHandler(function(url){alert("✅ Exported!");window.open(url,"_blank")}).exportUndoHistoryToSheet()}</script></body></html>'
  ).setWidth(800).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '↩️ Undo/Redo History');
}

function createGrievanceSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  return { timestamp: new Date().toISOString(), data: sheet.getDataRange().getValues(), lastRow: sheet.getLastRow(), lastColumn: sheet.getLastColumn() };
}

function restoreFromSnapshot(snapshot) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  if (snapshot.data.length > 1) sheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length).setValues(snapshot.data.slice(1));
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Snapshot restored', 'Undo/Redo', 3);
}
