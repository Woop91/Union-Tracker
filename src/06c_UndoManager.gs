/**
 * ============================================================================
 * UndoManager.gs - Undo/Redo System
 * ============================================================================
 *
 * This module handles the undo/redo functionality including:
 * - Action recording
 * - State management
 * - History navigation
 * - Snapshot management
 *
 * REFACTORED: Split from 06_Maintenance.gs for better maintainability
 *
 * @fileoverview Undo/redo system functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// UNDO CONFIGURATION
// ============================================================================

/**
 * Undo system configuration
 * @const {Object}
 */
var UNDO_CONFIG = {
  /** Maximum number of actions to store */
  MAX_HISTORY: 50,
  /** Storage key for undo history */
  STORAGE_KEY: 'undoRedoHistory'
};

// ============================================================================
// HISTORY MANAGEMENT
// ============================================================================

/**
 * Gets the undo history from storage
 * @returns {Object} History object with actions array and currentIndex
 */
function getUndoHistory() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty(UNDO_CONFIG.STORAGE_KEY);

  if (json) {
    return JSON.parse(json);
  }

  return { actions: [], currentIndex: 0 };
}

/**
 * Saves the undo history to storage
 * @param {Object} history - History object to save
 * @returns {void}
 */
function saveUndoHistory(history) {
  var props = PropertiesService.getScriptProperties();

  // Trim history if over limit
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }

  props.setProperty(UNDO_CONFIG.STORAGE_KEY, JSON.stringify(history));
}

/**
 * Clears all undo history
 * @returns {void}
 */
function clearUndoHistory() {
  PropertiesService.getScriptProperties().deleteProperty(UNDO_CONFIG.STORAGE_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('History cleared', 'Undo/Redo', 3);
}

// ============================================================================
// ACTION RECORDING
// ============================================================================

/**
 * Records an action for undo/redo
 * @param {string} type - Action type (EDIT_CELL, ADD_ROW, DELETE_ROW, etc.)
 * @param {string} description - Human-readable description
 * @param {Object} beforeState - State before the action
 * @param {Object} afterState - State after the action
 * @returns {void}
 */
function recordAction(type, description, beforeState, afterState) {
  var history = getUndoHistory();

  // Clear any redo history when new action is recorded
  if (history.currentIndex < history.actions.length) {
    history.actions = history.actions.slice(0, history.currentIndex);
  }

  history.actions.push({
    type: type,
    description: description,
    timestamp: new Date().toISOString(),
    beforeState: beforeState,
    afterState: afterState
  });

  history.currentIndex = history.actions.length;
  saveUndoHistory(history);
}

/**
 * Records a cell edit action
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {*} oldValue - Value before edit
 * @param {*} newValue - Value after edit
 * @returns {void}
 */
function recordCellEdit(row, col, oldValue, newValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colName = sheet.getRange(1, col).getValue();

  recordAction(
    'EDIT_CELL',
    'Edited ' + colName + ' in row ' + row,
    { row: row, col: col, value: oldValue, sheet: sheet.getName() },
    { row: row, col: col, value: newValue, sheet: sheet.getName() }
  );
}

/**
 * Records a row addition action
 * @param {number} row - Row number
 * @param {Array} rowData - Data in the new row
 * @returns {void}
 */
function recordRowAddition(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  recordAction(
    'ADD_ROW',
    'Added row ' + row,
    null,
    { row: row, data: rowData, sheet: sheet.getName() }
  );
}

/**
 * Records a row deletion action
 * @param {number} row - Row number
 * @param {Array} rowData - Data that was in the row
 * @returns {void}
 */
function recordRowDeletion(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  recordAction(
    'DELETE_ROW',
    'Deleted row ' + row,
    { row: row, data: rowData, sheet: sheet.getName() },
    null
  );
}

// ============================================================================
// UNDO/REDO OPERATIONS
// ============================================================================

/**
 * Undoes the last action
 * @returns {void}
 * @throws {Error} If nothing to undo
 */
function undoLastAction() {
  var history = getUndoHistory();

  if (history.currentIndex === 0) {
    throw new Error('Nothing to undo');
  }

  var action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);

  history.currentIndex--;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast('Undone: ' + action.description, 'Undo', 3);
}

/**
 * Redoes the last undone action
 * @returns {void}
 * @throws {Error} If nothing to redo
 */
function redoLastAction() {
  var history = getUndoHistory();

  if (history.currentIndex >= history.actions.length) {
    throw new Error('Nothing to redo');
  }

  var action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);

  history.currentIndex++;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast('Redone: ' + action.description, 'Redo', 3);
}

/**
 * Undoes all actions up to a specific index
 * @param {number} targetIndex - Target index to undo to
 * @returns {void}
 */
function undoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex > targetIndex) {
    undoLastAction();
    history = getUndoHistory();
  }
}

/**
 * Redoes all actions up to a specific index
 * @param {number} targetIndex - Target index to redo to
 * @returns {void}
 */
function redoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) {
    redoLastAction();
    history = getUndoHistory();
  }
}

/**
 * Applies a saved state to the spreadsheet
 * @param {Object} state - State to apply
 * @param {string} actionType - Type of action being applied
 * @returns {void}
 */
function applyState(state, actionType) {
  if (!state) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(state.sheet);

  if (!sheet) {
    throw new Error('Sheet ' + state.sheet + ' not found');
  }

  switch (actionType) {
    case 'EDIT_CELL':
      sheet.getRange(state.row, state.col).setValue(state.value);
      break;

    case 'ADD_ROW':
      if (state.row) {
        sheet.deleteRow(state.row);
      }
      break;

    case 'DELETE_ROW':
      if (state.row && state.data) {
        sheet.insertRowAfter(state.row - 1);
        sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]);
      }
      break;

    case 'BATCH_UPDATE':
      if (state.changes) {
        state.changes.forEach(function(c) {
          sheet.getRange(c.row, c.col).setValue(c.oldValue);
        });
      }
      break;
  }
}

// ============================================================================
// SNAPSHOT MANAGEMENT
// ============================================================================

/**
 * Creates a snapshot of the grievance log
 * @returns {Object} Snapshot object
 */
function createGrievanceSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    throw new Error('Grievance Log not found');
  }

  return {
    timestamp: new Date().toISOString(),
    data: sheet.getDataRange().getValues(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn()
  };
}

/**
 * Restores the grievance log from a snapshot
 * @param {Object} snapshot - Snapshot to restore
 * @returns {void}
 */
function restoreFromSnapshot(snapshot) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    throw new Error('Grievance Log not found');
  }

  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // Restore data
  if (snapshot.data.length > 1) {
    sheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length)
      .setValues(snapshot.data.slice(1));
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Snapshot restored', 'Undo/Redo', 3);
}

// ============================================================================
// EXPORT AND UI
// ============================================================================

/**
 * Exports undo history to a new sheet
 * @returns {string} URL of the exported sheet
 */
function exportUndoHistoryToSheet() {
  var history = getUndoHistory();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName('Undo_History_Export');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Undo_History_Export');
  }

  var headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#7c3aed')
    .setFontColor('#fff');

  if (history.actions.length > 0) {
    var rows = history.actions.map(function(a, i) {
      return [
        i + 1,
        a.type,
        a.description,
        new Date(a.timestamp).toLocaleString(),
        i < history.currentIndex ? 'Applied' : 'Undone'
      ];
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }

  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

/**
 * Shows the undo/redo panel
 * @returns {void}
 */
function showUndoRedoPanel() {
  var history = getUndoHistory();

  var rows = history.actions.length === 0
    ? '<tr><td colspan="5" style="text-align:center;padding:40px;color:#999">No actions recorded</td></tr>'
    : history.actions.slice().reverse().map(function(a, i) {
        var idx = history.actions.length - i;
        var time = new Date(a.timestamp).toLocaleString();
        var canUndo = idx <= history.currentIndex;

        return '<tr>' +
          '<td>' + idx + '</td>' +
          '<td><span class="badge ' + a.type.toLowerCase() + '">' + a.type + '</span></td>' +
          '<td>' + a.description + '</td>' +
          '<td style="font-size:12px;color:#666">' + time + '</td>' +
          '<td>' + (canUndo ? '<button onclick="undo(' + (idx - 1) + ')">↩️</button>' : '') + '</td>' +
          '</tr>';
      }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px}' +
    'h2{color:#7c3aed}' +
    '.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin:20px 0}' +
    '.stat{background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;border-left:4px solid #7c3aed}' +
    '.num{font-size:32px;font-weight:bold;color:#7c3aed}' +
    'table{width:100%;border-collapse:collapse;margin:20px 0}' +
    'th{background:#7c3aed;color:white;padding:12px;text-align:left}' +
    'td{padding:12px;border-bottom:1px solid #e0e0e0}' +
    'button{background:#7c3aed;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;margin:2px}' +
    '.badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:11px;font-weight:bold}' +
    '.edit_cell{background:#e3f2fd;color:#1976d2}' +
    '.add_row{background:#e8f5e9;color:#388e3c}' +
    '.delete_row{background:#ffebee;color:#d32f2f}' +
    '.actions{margin:20px 0;padding:20px;background:#f8f9fa;border-radius:8px}' +
    '.btn-danger{background:#ef4444}' +
    '.btn-success{background:#10b981}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>↩️ Undo/Redo History</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + history.actions.length + '</div><div>Total Actions</div></div>' +
    '<div class="stat"><div class="num">' + history.currentIndex + '</div><div>Current Position</div></div>' +
    '<div class="stat"><div class="num">' + (history.actions.length - history.currentIndex) + '</div><div>Can Redo</div></div>' +
    '</div>' +
    '<div class="actions">' +
    '<button onclick="performUndo()">↩️ Undo Last</button>' +
    '<button onclick="performRedo()">↪️ Redo</button>' +
    '<button class="btn-danger" onclick="clearHistory()">🗑️ Clear All</button>' +
    '<button class="btn-success" onclick="exportHistory()">📥 Export</button>' +
    '</div>' +
    '<table>' +
    '<tr><th>#</th><th>Type</th><th>Description</th><th>Time</th><th></th></tr>' +
    rows +
    '</table>' +
    '</div>' +
    '<script>' +
    'function performUndo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Error: "+e.message)}).undoLastAction()}' +
    'function performRedo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Error: "+e.message)}).redoLastAction()}' +
    'function undo(i){google.script.run.withSuccessHandler(function(){location.reload()}).undoToIndex(i)}' +
    'function clearHistory(){if(confirm("Clear all history?")){google.script.run.withSuccessHandler(function(){location.reload()}).clearUndoHistory()}}' +
    'function exportHistory(){google.script.run.withSuccessHandler(function(url){alert("Exported!");window.open(url,"_blank")}).exportUndoHistoryToSheet()}' +
    '</script></body></html>'
  ).setWidth(800).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '↩️ Undo/Redo History');
}
