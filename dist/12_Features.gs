/**
 * ============================================================================
 * 12_Features.gs - FEATURE MODULES
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Feature modules — Case Checklist Manager, Dynamic Engine, Member Leaders,
 *   Column Expansion, and Grievance Reminders. The checklist manager tracks
 *   action items per grievance case with status tracking. The Dynamic Engine
 *   handles self-healing formulas and auto-expanding column structures.
 *   Grievance reminders allow setting/getting/clearing deadline reminders
 *   per case.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   These features were grouped together because they all extend the core
 *   grievance/member functionality without being large enough for their own
 *   files. The checklist sheet is created on-demand (getOrCreateChecklistSheet)
 *   rather than during setup because not all organizations use case checklists.
 *   Header maps (CHECKLIST_HEADER_MAP_) use the same buildColsFromMap_()
 *   pattern as core sheets for consistency.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Case checklists stop working — stewards can't track action items per
 *   grievance. Grievance reminders fail — stewards miss deadline alerts.
 *   The Dynamic Engine stops self-healing formulas. Column expansion for
 *   new features fails.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, buildColsFromMap_, COLORS),
 *   00_Security.gs (escapeForFormula).
 *   Used by menu items in 03_, daily trigger reminders, and the SPA.
 *
 * @version 4.51.0
 * @license Free for use by non-profit collective bargaining groups and unions
 * ============================================================================
 */

// ============================================================================
// CHECKLIST SHEET MANAGEMENT
// ============================================================================

/**
 * Get or create the Case Checklist sheet
 * @returns {Sheet} The Case Checklist sheet
 */
function getOrCreateChecklistSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECKLIST_SHEET_NAME);

  if (!sheet) {
    sheet = createChecklistSheet_(ss);
  }

  return sheet;
}

/**
 * Create the Case Checklist sheet with proper structure
 * @param {Spreadsheet} ss - The spreadsheet
 * @returns {Sheet} The created sheet
 * @private
 */
function createChecklistSheet_(ss) {
  var sheet = ss.insertSheet(CHECKLIST_SHEET_NAME);

  // Headers — auto-derived from CHECKLIST_HEADER_MAP_
  var headers = getHeadersFromMap_(CHECKLIST_HEADER_MAP_);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground(COLORS.HEADER_BG || '#7C3AED')
    .setFontColor(COLORS.WHITE || '#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(CHECKLIST_COLS.CHECKLIST_ID, 100);
  sheet.setColumnWidth(CHECKLIST_COLS.CASE_ID, 120);
  sheet.setColumnWidth(CHECKLIST_COLS.ACTION_TYPE, 130);
  sheet.setColumnWidth(CHECKLIST_COLS.ITEM_TEXT, 350);
  sheet.setColumnWidth(CHECKLIST_COLS.CATEGORY, 120);
  sheet.setColumnWidth(CHECKLIST_COLS.REQUIRED, 80);
  sheet.setColumnWidth(CHECKLIST_COLS.COMPLETED, 90);
  sheet.setColumnWidth(CHECKLIST_COLS.COMPLETED_BY, 150);
  sheet.setColumnWidth(CHECKLIST_COLS.COMPLETED_DATE, 120);
  sheet.setColumnWidth(CHECKLIST_COLS.DUE_DATE, 100);
  sheet.setColumnWidth(CHECKLIST_COLS.NOTES, 250);
  sheet.setColumnWidth(CHECKLIST_COLS.SORT_ORDER, 80);

  // Add data validation for Required column (Yes/No)
  var requiredRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, CHECKLIST_COLS.REQUIRED, 1000, 1).setDataValidation(requiredRule);

  // Add checkboxes for Completed column
  sheet.getRange(2, CHECKLIST_COLS.COMPLETED, 1000, 1).insertCheckboxes();

  // Add data validation for Category column
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CHECKLIST_CATEGORIES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, CHECKLIST_COLS.CATEGORY, 1000, 1).setDataValidation(categoryRule);

  // Add data validation for Action Type column
  var actionTypes = ACTION_TYPE_CONFIG.map(function(config) { return config.value; });
  var actionTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(actionTypes, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, CHECKLIST_COLS.ACTION_TYPE, 1000, 1).setDataValidation(actionTypeRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // No protection - sheet is fully dynamic like all other tabs
  // Self-healing formulas in _Checklist_Calc handle calculations

  // Set tab color
  sheet.setTabColor(COLORS.CHART_YELLOW);

  return sheet;
}

// ============================================================================
// CHECKLIST ITEM MANAGEMENT
// ============================================================================

/**
 * Generate a unique checklist item ID
 * @param {number} [offset=0] - Offset to add for batch creation (prevents duplicate IDs)
 * @returns {string} Unique checklist ID (e.g., CL-00001)
 */
function generateChecklistId_(offset) {
  offset = offset || 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getOrCreateChecklistSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return 'CL-' + String(1 + offset).padStart(5, '0');
    }

    var ids = sheet.getRange(2, CHECKLIST_COLS.CHECKLIST_ID, lastRow - 1, 1).getValues();
    var maxNum = 0;

    for (var i = 0; i < ids.length; i++) {
      var id = ids[i][0];
      if (id && typeof id === 'string' && id.indexOf('CL-') === 0) {
        var numPart = parseInt(id.substring(3), 10);
        if (!isNaN(numPart) && numPart > maxNum) {
          maxNum = numPart;
        }
      }
    }

    return 'CL-' + String(maxNum + 1 + offset).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Create checklist items for a case from a template
 * @param {string} caseId - The case/grievance ID
 * @param {string} actionType - The action type (Grievance, Records Request, etc.)
 * @param {string} issueCategory - The issue category (for grievances only)
 * @returns {Object} Result with success status and items created
 */
function createChecklistFromTemplate(caseId, actionType, issueCategory) {
  if (!caseId) {
    return errorResponse('Case ID is required');
  }

  // Default to Grievance if not specified
  actionType = actionType || 'Grievance';

  // Get the template items
  var templateItems = getChecklistTemplate(actionType, issueCategory);

  if (!templateItems || templateItems.length === 0) {
    return errorResponse('No template found for this action type');
  }

  var sheet = getOrCreateChecklistSheet();
  var itemsCreated = [];

  // Acquire lock ONCE for the entire batch instead of per-item via generateChecklistId_
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Determine max existing ID once
    var lastRow = sheet.getLastRow();
    var maxNum = 0;
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, CHECKLIST_COLS.CHECKLIST_ID, lastRow - 1, 1).getValues();
      for (var k = 0; k < ids.length; k++) {
        var existId = ids[k][0];
        if (existId && typeof existId === 'string' && existId.indexOf('CL-') === 0) {
          var numPart = parseInt(existId.substring(3), 10);
          if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
        }
      }
    }

    // Prepare rows to add — derive sequential IDs from maxNum
    var rows = [];
    for (var i = 0; i < templateItems.length; i++) {
      var item = templateItems[i];
      var checklistId = 'CL-' + String(maxNum + 1 + i).padStart(5, '0');

      var row = [];
      setCol_(row, CHECKLIST_COLS.CHECKLIST_ID, checklistId);
      setCol_(row, CHECKLIST_COLS.CASE_ID, caseId);
      setCol_(row, CHECKLIST_COLS.ACTION_TYPE, actionType);
      setCol_(row, CHECKLIST_COLS.ITEM_TEXT, item.text);
      setCol_(row, CHECKLIST_COLS.CATEGORY, item.category);
      setCol_(row, CHECKLIST_COLS.REQUIRED, item.required ? 'Yes' : 'No');
      setCol_(row, CHECKLIST_COLS.COMPLETED, false);
      setCol_(row, CHECKLIST_COLS.COMPLETED_BY, '');
      setCol_(row, CHECKLIST_COLS.COMPLETED_DATE, '');
      setCol_(row, CHECKLIST_COLS.DUE_DATE, '');
      setCol_(row, CHECKLIST_COLS.NOTES, '');
      setCol_(row, CHECKLIST_COLS.SORT_ORDER, i + 1);

      rows.push(row);
      itemsCreated.push({
        id: checklistId,
        text: item.text,
        category: item.category,
        required: item.required
      });
    }

    // Add all rows at once for efficiency
    if (rows.length > 0) {
      var startRow = lastRow + 1;
      sheet.getRange(startRow, 1, rows.length, CHECKLIST_HEADER_MAP_.length).setValues(rows);

      // Add checkboxes to the new rows
      sheet.getRange(startRow, CHECKLIST_COLS.COMPLETED, rows.length, 1).insertCheckboxes();
    }
  } finally {
    lock.releaseLock();
  }

  // Update the checklist progress on the grievance
  updateChecklistProgress(caseId);

  // Log the action
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('CHECKLIST_CREATED', caseId, {
      actionType: actionType,
      issueCategory: issueCategory,
      itemCount: itemsCreated.length
    });
  }

  return {
    success: true,
    itemsCreated: itemsCreated,
    count: itemsCreated.length
  };
}

/**
 * Get all checklist items for a case
 * @param {string} caseId - The case/grievance ID
 * @returns {Array} Array of checklist item objects
 */
function getChecklistItems(caseId) {
  if (!caseId) {
    return [];
  }

  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, CHECKLIST_HEADER_MAP_.length).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (col_(row, CHECKLIST_COLS.CASE_ID) === caseId) {
      items.push({
        checklistId: col_(row, CHECKLIST_COLS.CHECKLIST_ID),
        caseId: col_(row, CHECKLIST_COLS.CASE_ID),
        actionType: col_(row, CHECKLIST_COLS.ACTION_TYPE),
        itemText: col_(row, CHECKLIST_COLS.ITEM_TEXT),
        category: col_(row, CHECKLIST_COLS.CATEGORY),
        required: col_(row, CHECKLIST_COLS.REQUIRED) === 'Yes',
        completed: col_(row, CHECKLIST_COLS.COMPLETED) === true,
        completedBy: col_(row, CHECKLIST_COLS.COMPLETED_BY),
        completedDate: col_(row, CHECKLIST_COLS.COMPLETED_DATE),
        dueDate: col_(row, CHECKLIST_COLS.DUE_DATE),
        notes: col_(row, CHECKLIST_COLS.NOTES),
        sortOrder: col_(row, CHECKLIST_COLS.SORT_ORDER),
        rowIndex: i + 2  // 1-indexed, accounting for header
      });
    }
  }

  // Sort by sort order
  items.sort(function(a, b) {
    return (a.sortOrder || 999) - (b.sortOrder || 999);
  });

  return items;
}

/**
 * Get checklist progress for a case
 * @param {string} caseId - The case/grievance ID
 * @returns {Object} Progress object with completed, total, percentage, and display string
 */
function getChecklistProgress(caseId) {
  var items = getChecklistItems(caseId);

  if (items.length === 0) {
    return {
      completed: 0,
      total: 0,
      percentage: 0,
      display: 'No checklist',
      requiredCompleted: 0,
      requiredTotal: 0
    };
  }

  var completed = 0;
  var requiredCompleted = 0;
  var requiredTotal = 0;

  for (var i = 0; i < items.length; i++) {
    if (items[i].completed) {
      completed++;
      if (items[i].required) {
        requiredCompleted++;
      }
    }
    if (items[i].required) {
      requiredTotal++;
    }
  }

  var percentage = Math.round((completed / items.length) * 100);

  return {
    completed: completed,
    total: items.length,
    percentage: percentage,
    display: completed + '/' + items.length + ' (' + percentage + '%)',
    requiredCompleted: requiredCompleted,
    requiredTotal: requiredTotal,
    allRequiredComplete: requiredCompleted >= requiredTotal
  };
}

/**
 * Update checklist progress on the Grievance Log
 * @param {string} caseId - The case/grievance ID
 */
function updateChecklistProgress(caseId) {
  var progress = getChecklistProgress(caseId);

  // Find the grievance row and update the progress column
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    return;
  }

  // Ensure sheet has enough columns for CHECKLIST_PROGRESS
  ensureMinimumColumns(grievanceSheet, getGrievanceHeaders().length);

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < grievanceIds.length; i++) {
    if (grievanceIds[i][0] === caseId) {
      var row = i + 2;
      grievanceSheet.getRange(row, GRIEVANCE_COLS.CHECKLIST_PROGRESS).setValue(progress.display);
      break;
    }
  }
}

/**
 * Mark a checklist item as completed or uncompleted
 * @param {string} checklistId - The checklist item ID
 * @param {boolean} completed - Whether the item is completed
 * @param {string} completedBy - Who completed the item (optional, defaults to current user)
 * @returns {Object} Result object
 */
function setChecklistItemCompleted(checklistId, completed, completedBy) {
  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return errorResponse('Checklist item not found');
  }

  var data = sheet.getRange(2, 1, lastRow - 1, CHECKLIST_HEADER_MAP_.length).getValues();

  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], CHECKLIST_COLS.CHECKLIST_ID) === checklistId) {
      var row = i + 2;
      var caseId = col_(data[i], CHECKLIST_COLS.CASE_ID);

      // Update completed status
      sheet.getRange(row, CHECKLIST_COLS.COMPLETED).setValue(completed);

      if (completed) {
        // Set completed by and date
        var user = completedBy || Session.getActiveUser().getEmail() || 'Unknown';
        sheet.getRange(row, CHECKLIST_COLS.COMPLETED_BY).setValue(user);
        sheet.getRange(row, CHECKLIST_COLS.COMPLETED_DATE).setValue(new Date());
      } else {
        // Clear completed by and date
        sheet.getRange(row, CHECKLIST_COLS.COMPLETED_BY).setValue('');
        sheet.getRange(row, CHECKLIST_COLS.COMPLETED_DATE).setValue('');
      }

      // Update progress on the grievance
      updateChecklistProgress(caseId);

      return { success: true, caseId: caseId };
    }
  }

  return errorResponse('Checklist item not found');
}

/**
 * Add a custom checklist item to a case
 * @param {string} caseId - The case/grievance ID
 * @param {string} itemText - The item description
 * @param {string} category - The category (Document, Meeting, etc.)
 * @param {boolean} required - Whether the item is required
 * @param {Date} dueDate - Optional due date
 * @returns {Object} Result with the new item
 */
function addChecklistItem(caseId, itemText, category, required, dueDate) {
  if (!caseId || !itemText) {
    return errorResponse('Case ID and item text are required');
  }

  // Validate category
  category = category || 'Other';
  if (CHECKLIST_CATEGORIES.indexOf(category) === -1) {
    category = 'Other';
  }

  var sheet = getOrCreateChecklistSheet();
  var checklistId = generateChecklistId_();

  // Get current max sort order for this case
  var items = getChecklistItems(caseId);
  var maxSortOrder = 0;
  for (var i = 0; i < items.length; i++) {
    if (items[i].sortOrder > maxSortOrder) {
      maxSortOrder = items[i].sortOrder;
    }
  }

  // Get action type from the grievance
  var actionType = 'Grievance';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet) {
    var lastRow = grievanceSheet.getLastRow();
    if (lastRow >= 2) {
      var grievanceData = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.ACTION_TYPE).getValues();
      for (var j = 0; j < grievanceData.length; j++) {
        if (grievanceData[j][0] === caseId) {
          actionType = col_(grievanceData[j], GRIEVANCE_COLS.ACTION_TYPE) || 'Grievance';
          break;
        }
      }
    }
  }

  var row = [
    checklistId,
    caseId,
    actionType,
    itemText,
    category,
    required ? 'Yes' : 'No',
    false,
    '',
    '',
    dueDate || '',
    '',
    maxSortOrder + 1
  ];

  sheet.appendRow(row);

  // Add checkbox to the new row
  var newRow = sheet.getLastRow();
  sheet.getRange(newRow, CHECKLIST_COLS.COMPLETED).insertCheckboxes();

  // Update progress
  updateChecklistProgress(caseId);

  return {
    success: true,
    item: {
      checklistId: checklistId,
      caseId: caseId,
      itemText: itemText,
      category: category,
      required: required
    }
  };
}

/**
 * Delete a checklist item
 * @param {string} checklistId - The checklist item ID
 * @returns {Object} Result object
 */
function deleteChecklistItem(checklistId) {
  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return errorResponse('Checklist item not found');
  }

  var numCols = Math.max(CHECKLIST_COLS.CHECKLIST_ID, CHECKLIST_COLS.CASE_ID);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], CHECKLIST_COLS.CHECKLIST_ID) === checklistId) {
      var caseId = col_(data[i], CHECKLIST_COLS.CASE_ID);
      var row = i + 2;
      sheet.deleteRow(row);

      // Update progress
      updateChecklistProgress(caseId);

      return { success: true, caseId: caseId };
    }
  }

  return errorResponse('Checklist item not found');
}

/**
 * Update a checklist item's text or properties
 * @param {string} checklistId - The checklist item ID
 * @param {Object} updates - Object with properties to update (itemText, category, required, dueDate, notes)
 * @returns {Object} Result object
 */
function updateChecklistItem(checklistId, updates) {
  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return errorResponse('Checklist item not found');
  }

  var data = sheet.getRange(2, 1, lastRow - 1, CHECKLIST_HEADER_MAP_.length).getValues();

  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], CHECKLIST_COLS.CHECKLIST_ID) === checklistId) {
      var row = i + 2;
      var rowData = data[i].slice(); // copy row; apply all changes in-memory

      if (updates.itemText !== undefined) {
        setCol_(rowData, CHECKLIST_COLS.ITEM_TEXT, escapeForFormula(updates.itemText));
      }
      if (updates.category !== undefined) {
        setCol_(rowData, CHECKLIST_COLS.CATEGORY, escapeForFormula(updates.category));
      }
      if (updates.required !== undefined) {
        setCol_(rowData, CHECKLIST_COLS.REQUIRED, updates.required ? 'Yes' : 'No');
      }
      if (updates.dueDate !== undefined) {
        setCol_(rowData, CHECKLIST_COLS.DUE_DATE, updates.dueDate);
      }
      if (updates.notes !== undefined) {
        setCol_(rowData, CHECKLIST_COLS.NOTES, escapeForFormula(updates.notes));
      }

      // Write entire row in one API call instead of one setValue per field
      sheet.getRange(row, 1, 1, CHECKLIST_HEADER_MAP_.length).setValues([rowData]);
      return { success: true };
    }
  }

  return errorResponse('Checklist item not found');
}
// ============================================================================
// CHECKLIST UI HELPERS
// ============================================================================

/**
 * Get checklist data formatted for UI display
 * @param {string} caseId - The case/grievance ID
 * @returns {Object} Formatted data for UI
 */
function getChecklistForUI(caseId) {
  var items = getChecklistItems(caseId);
  var progress = getChecklistProgress(caseId);

  // Group items by category
  var byCategory = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var cat = item.category || 'Other';
    if (!byCategory[cat]) {
      byCategory[cat] = [];
    }
    byCategory[cat].push(item);
  }

  return {
    caseId: caseId,
    items: items,
    byCategory: byCategory,
    progress: progress,
    categories: Object.keys(byCategory)
  };
}
// ============================================================================
// BATCH OPERATIONS
// ============================================================================

// ============================================================================
// ONOPEN / ONEDIT INTEGRATION
// ============================================================================

/**
 * Handle checkbox changes in the checklist sheet
 * Called from the main onEdit trigger
 * @param {Object} e - The edit event object
 */
function handleChecklistEdit(e) {
  if (!e || !e.range) {
    return;
  }

  var sheet = e.range.getSheet();
  if (sheet.getName() !== CHECKLIST_SHEET_NAME) {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Only handle edits to the Completed column (not header row)
  if (row < 2 || col !== CHECKLIST_COLS.COMPLETED) {
    return;
  }

  var completed = isTruthyValue(e.value);
  var caseId = sheet.getRange(row, CHECKLIST_COLS.CASE_ID).getValue();

  if (completed) {
    // Set completed by and date
    var user = Session.getActiveUser().getEmail() || 'Unknown';
    sheet.getRange(row, CHECKLIST_COLS.COMPLETED_BY).setValue(user);
    sheet.getRange(row, CHECKLIST_COLS.COMPLETED_DATE).setValue(new Date());
  } else {
    // Clear completed by and date
    sheet.getRange(row, CHECKLIST_COLS.COMPLETED_BY).setValue('');
    sheet.getRange(row, CHECKLIST_COLS.COMPLETED_DATE).setValue('');
  }

  // Update progress on the grievance
  if (caseId) {
    updateChecklistProgress(caseId);
  }
}

// ============================================================================
// CHECKLIST UI DIALOGS
// ============================================================================

/**
 * Show checklist dialog for the selected grievance
 */
function showChecklistDialog() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG && sheet.getName() !== 'Grievance Log') {
    SpreadsheetApp.getUi().alert('Please select a case in the Grievance Log sheet');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a case row');
    return;
  }

  var caseId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  if (!caseId) {
    SpreadsheetApp.getUi().alert('No case ID found in selected row');
    return;
  }

  showDialog_(getChecklistDialogHtml(caseId), 'Case Checklist: ' + caseId, 700, 600);
}

/**
 * Generate HTML for the checklist dialog
 * @param {string} caseId - The case/grievance ID
 * @returns {string} HTML content
 */
function getChecklistDialogHtml(caseId) {
  var checklistData = getChecklistForUI(caseId);
  var items = checklistData.items;
  var progress = checklistData.progress;

  // Build items HTML grouped by category
  var itemsHtml = '';
  var categories = checklistData.categories;

  if (items.length === 0) {
    itemsHtml = '<div class="no-items">No checklist items yet. Click "Add from Template" to create a checklist.</div>';
  } else {
    for (var c = 0; c < categories.length; c++) {
      var cat = categories[c];
      var catItems = checklistData.byCategory[cat];

      itemsHtml += '<div class="category-section">';
      itemsHtml += '<div class="category-header">' + escapeHtml(cat) + ' (' + catItems.length + ')</div>';

      for (var i = 0; i < catItems.length; i++) {
        var item = catItems[i];
        var checkedAttr = item.completed ? 'checked' : '';
        var completedClass = item.completed ? 'completed' : '';
        var requiredBadge = item.required ? '<span class="badge required">Required</span>' : '';
        var safeChecklistId = String(item.checklistId || '').replace(/[^a-zA-Z0-9_-]/g, '');

        itemsHtml += '<div class="checklist-item ' + completedClass + '">';
        itemsHtml += '  <label class="checkbox-container">';
        itemsHtml += '    <input type="checkbox" ' + checkedAttr + ' onchange="toggleItem(\'' + safeChecklistId + '\', this.checked)">';
        itemsHtml += '    <span class="checkmark"></span>';
        itemsHtml += '  </label>';
        itemsHtml += '  <div class="item-content">';
        itemsHtml += '    <div class="item-text">' + escapeHtml(item.itemText) + ' ' + requiredBadge + '</div>';
        if (item.completed && item.completedBy) {
          itemsHtml += '    <div class="item-meta">Completed by ' + escapeHtml(item.completedBy) + '</div>';
        }
        itemsHtml += '  </div>';
        itemsHtml += '  <button class="btn-icon" onclick="deleteItem(\'' + safeChecklistId + '\')" title="Delete item">x</button>';
        itemsHtml += '</div>';
      }

      itemsHtml += '</div>';
    }
  }

  // Calculate progress bar
  var progressPct = progress.percentage || 0;
  var progressColor = progressPct >= 100 ? '#10B981' : (progressPct >= 50 ? '#F59E0B' : '#EF4444');

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    getMobileOptimizedHead() +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; background: #F9FAFB; color: #1F2937; }' +
    '.container { padding: clamp(12px,3vw,20px); }' +
    '.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }' +
    '.progress-section { background: white; border-radius: 12px; padding: clamp(12px,3vw,16px); margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    '.progress-label { display: flex; justify-content: space-between; margin-bottom: 8px; font-size: clamp(12px,3vw,14px); }' +
    '.progress-bar { height: 8px; background: #E5E7EB; border-radius: 4px; overflow: hidden; }' +
    '.progress-fill { height: 100%; background: ' + progressColor + '; border-radius: 4px; transition: width 0.3s; }' +
    '.checklist-container { background: white; border-radius: 12px; padding: clamp(12px,3vw,16px); box-shadow: 0 1px 3px rgba(0,0,0,0.1); max-height: 350px; overflow-y: auto; -webkit-overflow-scrolling: touch; }' +
    '.category-section { margin-bottom: 16px; }' +
    '.category-header { font-size: clamp(10px,2.8vw,12px); font-weight: 600; color: #6B7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 1px solid #E5E7EB; }' +
    '.checklist-item { display: flex; align-items: center; padding: clamp(8px,2.5vw,10px); border-radius: 8px; margin-bottom: 6px; background: #F9FAFB; transition: all 0.2s; min-height: 44px; }' +
    '.checklist-item:hover { background: #F3F4F6; }' +
    '.checklist-item.completed { opacity: 0.7; }' +
    '.checklist-item.completed .item-text { text-decoration: line-through; color: #9CA3AF; }' +
    '.checkbox-container { position: relative; width: 24px; height: 24px; margin-right: 12px; flex-shrink: 0; }' +
    '.checkbox-container input { opacity: 0; width: 24px; height: 24px; cursor: pointer; }' +
    '.checkmark { position: absolute; top: 0; left: 0; width: 24px; height: 24px; background: white; border: 2px solid #D1D5DB; border-radius: 6px; pointer-events: none; }' +
    '.checkbox-container input:checked ~ .checkmark { background: #7C3AED; border-color: #7C3AED; }' +
    '.checkbox-container input:checked ~ .checkmark:after { content: ""; position: absolute; left: 7px; top: 3px; width: 6px; height: 12px; border: solid white; border-width: 0 2px 2px 0; transform: rotate(45deg); }' +
    '.item-content { flex: 1; }' +
    '.item-text { font-size: clamp(12px,3vw,14px); }' +
    '.item-meta { font-size: clamp(10px,2.5vw,11px); color: #9CA3AF; margin-top: 2px; }' +
    '.badge { font-size: clamp(9px,2.2vw,10px); padding: 2px 6px; border-radius: 4px; margin-left: 6px; }' +
    '.badge.required { background: #FEE2E2; color: #991B1B; }' +
    '.btn-icon { background: none; border: none; color: #9CA3AF; cursor: pointer; font-size: 16px; padding: 4px 8px; border-radius: 4px; min-height: 44px; min-width: 44px; display: flex; align-items: center; justify-content: center; }' +
    '.btn-icon:hover { background: #FEE2E2; color: #EF4444; }' +
    '.no-items { text-align: center; padding: 40px; color: #6B7280; }' +
    '.actions { display: flex; gap: 10px; margin-top: 20px; justify-content: space-between; flex-wrap: wrap; }' +
    '.btn { padding: 10px 16px; border-radius: 8px; font-size: clamp(12px,3vw,14px); font-weight: 500; cursor: pointer; border: none; transition: all 0.2s; min-height: 44px; }' +
    '.btn-primary { background: #7C3AED; color: white; }' +
    '.btn-primary:hover { background: #6D28D9; }' +
    '.btn-secondary { background: #E5E7EB; color: #374151; }' +
    '.btn-secondary:hover { background: #D1D5DB; }' +
    '.btn-success { background: #10B981; color: white; }' +
    '.btn-success:hover { background: #059669; }' +
    '.add-item-section { margin-top: 16px; padding-top: 16px; border-top: 1px solid #E5E7EB; }' +
    '.add-form { display: flex; gap: 8px; flex-wrap: wrap; }' +
    '.add-form input { flex: 1; padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 16px; min-height: 44px; min-width: 0; }' +
    '.add-form select { padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 16px; min-height: 44px; }' +
    '@media(max-width:480px){.add-form{flex-direction:column}.add-form input,.add-form select,.add-form .btn{width:100%}.actions{flex-direction:column}.actions .btn{width:100%}}' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '  <div class="progress-section">' +
    '    <div class="progress-label">' +
    '      <span>Progress</span>' +
    '      <span>' + progress.completed + ' of ' + progress.total + ' complete (' + progressPct + '%)</span>' +
    '    </div>' +
    '    <div class="progress-bar">' +
    '      <div class="progress-fill" style="width: ' + progressPct + '%"></div>' +
    '    </div>' +
    '    <div style="margin-top: 8px; font-size: 12px; color: #6B7280;">' +
    '      Required items: ' + progress.requiredCompleted + ' of ' + progress.requiredTotal + ' complete' +
    '    </div>' +
    '  </div>' +
    '  <div class="checklist-container">' +
    itemsHtml +
    '  </div>' +
    '  <div class="add-item-section">' +
    '    <div class="add-form">' +
    '      <input type="text" id="newItemText" placeholder="Add custom checklist item...">' +
    '      <select id="newItemCategory">' +
    '        <option value="Document">Document</option>' +
    '        <option value="Meeting">Meeting</option>' +
    '        <option value="Evidence">Evidence</option>' +
    '        <option value="Communication">Communication</option>' +
    '        <option value="Follow-up">Follow-up</option>' +
    '        <option value="Other">Other</option>' +
    '      </select>' +
    '      <button class="btn btn-success" onclick="addItem()">Add</button>' +
    '    </div>' +
    '  </div>' +
    '  <div class="actions">' +
    '    <div>' +
    (items.length === 0 ?
    '      <button class="btn btn-primary" onclick="createFromTemplate()">Add from Template</button>' : '') +
    '    </div>' +
    '    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '  </div>' +
    '</div>' +
    '<script>' +
    'var caseId = "' + String(caseId || '').replace(/['"\\<>&]/g, '') + '";' +
    'function toggleItem(checklistId, completed) {' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { location.reload(); })' +
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })' +
    '    .setChecklistItemCompleted(checklistId, completed);' +
    '}' +
    'function deleteItem(checklistId) {' +
    '  if (!confirm("Delete this checklist item?")) return;' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { location.reload(); })' +
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })' +
    '    .deleteChecklistItem(checklistId);' +
    '}' +
    'function addItem() {' +
    '  var text = document.getElementById("newItemText").value.trim();' +
    '  var category = document.getElementById("newItemCategory").value;' +
    '  if (!text) { alert("Please enter item text"); return; }' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { location.reload(); })' +
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })' +
    '    .addChecklistItem(caseId, text, category, false, null);' +
    '}' +
    'function createFromTemplate() {' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.success) { location.reload(); }' +
    '      else { alert("Error: " + result.error); }' +
    '    })' +
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })' +
    '    .createChecklistForCaseId(caseId);' +
    '}' +
    '</script>' +
    '</body></html>';
}

/**
 * Create checklist for a specific case by ID
 * Called from the checklist dialog with the known caseId
 * @param {string} caseId - The grievance/case ID
 * @returns {Object} Result with success status
 */
function createChecklistForCaseId(caseId) {
  if (!caseId) {
    return errorResponse('No case ID provided');
  }

  // Look up the grievance to get action type and category
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) {
    return errorResponse('Grievance Log not found');
  }

  var data = sheet.getDataRange().getValues();
  var actionType = 'Grievance';
  var issueCategory = '';

  for (var i = 1; i < data.length; i++) {
    if (String(col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID)) === String(caseId)) {
      actionType = col_(data[i], GRIEVANCE_COLS.ACTION_TYPE) || 'Grievance';
      issueCategory = col_(data[i], GRIEVANCE_COLS.ISSUE_CATEGORY) || '';
      break;
    }
  }

  return createChecklistFromTemplate(caseId, actionType, issueCategory);
}

// ============================================================================
// ACTION TYPE MANAGEMENT
// ============================================================================

/**
 * Update the Action Type column in the Grievance Log
 * Adds the column if it doesn't exist and sets up validation
 */
function setupActionTypeColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    return errorResponse('Grievance Log not found');
  }

  ensureMinimumColumns(grievanceSheet, getGrievanceHeaders().length);

  // Check if Action Type header exists
  var lastCol = grievanceSheet.getLastColumn();
  if (lastCol < 1) {
    // Sheet has no data columns yet — write grievance headers first
    var gHeaders = getGrievanceHeaders();
    grievanceSheet.getRange(1, 1, 1, gHeaders.length).setValues([gHeaders]);
    lastCol = gHeaders.length;
  }
  var headers = grievanceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var actionTypeCol = headers.indexOf('Action Type') + 1;

  if (actionTypeCol === 0) {
    // Add the column header — resolve from header map (v4.51.1)
    var atCol = resolveColumnByHeader_(grievanceSheet, 'Action Type', GRIEVANCE_COLS.ACTION_TYPE);
    var cpCol2 = resolveColumnByHeader_(grievanceSheet, 'Checklist Progress', GRIEVANCE_COLS.CHECKLIST_PROGRESS);
    grievanceSheet.getRange(1, atCol).setValue('Action Type');
    grievanceSheet.getRange(1, cpCol2).setValue('Checklist Progress');
    actionTypeCol = atCol;
  }

  // Set up data validation for Action Type
  var actionTypes = ACTION_TYPE_CONFIG.map(function(config) { return config.value; });
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(actionTypes, true)
    .setAllowInvalid(false)
    .build();

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow >= 2) {
    grievanceSheet.getRange(2, actionTypeCol, lastRow - 1, 1).setDataValidation(rule);

    // Set default value for existing rows without an action type
    var existingValues = grievanceSheet.getRange(2, actionTypeCol, lastRow - 1, 1).getValues();
    var output = existingValues.map(function(v) { return [v[0] || 'Grievance']; });
    grievanceSheet.getRange(2, actionTypeCol, output.length, 1).setValues(output);
  }

  return { success: true };
}
// ============================================================================
// CHECKLIST SHEET PROTECTION MANAGEMENT
// ============================================================================

/**
 * Unlock the Checklist sheet by removing all protections
 * Call this from Admin > Setup > Unlock Checklist Sheet menu
 */
function unlockChecklistSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECKLIST_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Checklist Sheet Not Found',
      'The Case Checklist sheet does not exist. It will be created when you first use a checklist feature.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Get all protections on the sheet
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  if (protections.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Already Unlocked',
      'The Checklist sheet has no protections. It is already fully editable.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Remove all protections
  var removed = 0;
  protections.forEach(function(protection) {
    if (protection.canEdit()) {
      protection.remove();
      removed++;
    }
  });

  // Also check for range protections
  var rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  rangeProtections.forEach(function(protection) {
    if (protection.canEdit()) {
      protection.remove();
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    'Checklist Sheet Unlocked',
    'Removed ' + removed + ' protection(s) from the Checklist sheet.\n\n' +
    'The sheet is now fully editable. Note: The protection was originally added to prevent ' +
    'accidental modification of the sheet structure (headers, formulas).\n\n' +
    'If you want to re-protect the sheet later, run the sheet setup function.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
// ============================================================================
// HIDDEN SHEET: _Checklist_Calc (Self-Healing Formulas)
// ============================================================================

/**
 * Setup the hidden _Checklist_Calc sheet with self-healing formulas
 * This follows the same pattern as other hidden sheets in the system
 * Formulas automatically calculate progress per case
 */
function setupChecklistCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = HIDDEN_SHEETS.CHECKLIST_CALC || '_Checklist_Calc';
  var sheet = ss.getSheetByName(sheetName);

  // Create if doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Clear and rebuild (self-healing)
  sheet.clear();

  // Headers
  var headers = [
    'Case ID',
    'Items Total',
    'Items Completed',
    'Progress %',
    'Required Total',
    'Required Complete',
    'All Required Done',
    'Display String'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY || '#F3F4F6');

  // Get the checklist sheet name for formula references
  var clSheet = "'" + (CHECKLIST_SHEET_NAME || 'Case Checklist') + "'";

  // Column letters for checklist columns — auto-derived from constants
  var caseIdCol = getColumnLetter(CHECKLIST_COLS.CASE_ID);
  var completedCol = getColumnLetter(CHECKLIST_COLS.COMPLETED);
  var requiredCol = getColumnLetter(CHECKLIST_COLS.REQUIRED);

  // Column A: Unique Case IDs from checklist
  // ARRAYFORMULA with UNIQUE to get all case IDs
  sheet.getRange('A2').setFormula(
    '=IFERROR(UNIQUE(FILTER(' + clSheet + '!' + caseIdCol + ':' + caseIdCol + ',' +
    clSheet + '!' + caseIdCol + ':' + caseIdCol + '<>"","")))'
  );

  // Column B: Total items per case (COUNTIF)
  sheet.getRange('B2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(' + clSheet + '!' + caseIdCol + ':' + caseIdCol + ',A2:A)))'
  );

  // Column C: Completed items per case (COUNTIFS where completed=TRUE)
  sheet.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(' + clSheet + '!' + caseIdCol + ':' + caseIdCol + ',A2:A,' +
    clSheet + '!' + completedCol + ':' + completedCol + ',TRUE)))'
  );

  // Column D: Progress percentage
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B=0,0,ROUND(C2:C/B2:B*100,0)))'
  );

  // Column E: Required items total per case
  sheet.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(' + clSheet + '!' + caseIdCol + ':' + caseIdCol + ',A2:A,' +
    clSheet + '!' + requiredCol + ':' + requiredCol + ',"Yes")))'
  );

  // Column F: Required items completed per case
  sheet.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(' + clSheet + '!' + caseIdCol + ':' + caseIdCol + ',A2:A,' +
    clSheet + '!' + requiredCol + ':' + requiredCol + ',"Yes",' +
    clSheet + '!' + completedCol + ':' + completedCol + ',TRUE)))'
  );

  // Column G: All required done? (Yes/No)
  sheet.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(E2:E=0,"N/A",IF(F2:F>=E2:E,"Yes","No"))))'
  );

  // Column H: Display string (e.g., "5/10 (50%)")
  sheet.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",C2:C&"/"&B2:B&" ("&D2:D&"%)"))'
  );

  // Set column widths
  sheet.setColumnWidth(1, 120);  // Case ID
  sheet.setColumnWidth(2, 80);   // Items Total
  sheet.setColumnWidth(3, 100);  // Items Completed
  sheet.setColumnWidth(4, 80);   // Progress %
  sheet.setColumnWidth(5, 100);  // Required Total
  sheet.setColumnWidth(6, 110);  // Required Complete
  sheet.setColumnWidth(7, 100);  // All Required Done
  sheet.setColumnWidth(8, 120);  // Display String

  // Hide the sheet (mobile-safe via Sheets API + protection)
  setSheetVeryHidden_(sheet);

  log_('setupChecklistCalcSheet', '_Checklist_Calc sheet setup complete with self-healing formulas');
  return sheet;
}

/**
 * Sync checklist progress from _Checklist_Calc to Grievance Log
 * Pushes calculated values from hidden sheet to the visible Checklist Progress column
 */
function syncChecklistCalcToGrievanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get hidden calc sheet
  var calcSheetName = HIDDEN_SHEETS.CHECKLIST_CALC || '_Checklist_Calc';
  var calcSheet = ss.getSheetByName(calcSheetName);

  if (!calcSheet) {
    // Auto-create if missing (self-healing)
    calcSheet = setupChecklistCalcSheet();
  }

  // Get grievance log
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet) {
    log_('syncChecklistCalcToGrievanceLog', 'Grievance Log not found');
    return errorResponse('Grievance Log not found');
  }

  ensureMinimumColumns(grievanceSheet, getGrievanceHeaders().length);

  var grievanceLastRow = grievanceSheet.getLastRow();
  if (grievanceLastRow < 2) {
    return { success: true, updatedCount: 0 };
  }

  // Build lookup from calc sheet (Case ID -> Display String)
  var calcLastRow = calcSheet.getLastRow();
  var progressLookup = {};

  if (calcLastRow >= 2) {
    var calcData = calcSheet.getRange(2, 1, calcLastRow - 1, 8).getValues();
    for (var i = 0; i < calcData.length; i++) {
      var caseId = calcData[i][0];
      var displayString = calcData[i][7];  // Column H: Display String
      if (caseId) {
        progressLookup[caseId] = displayString || 'No checklist';
      }
    }
  }

  // Get grievance IDs and update progress column
  var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, grievanceLastRow - 1, 1).getValues();
  var progressValues = [];
  var updatedCount = 0;

  for (var j = 0; j < grievanceIds.length; j++) {
    var gId = grievanceIds[j][0];
    var progress = progressLookup[gId] || 'No checklist';
    progressValues.push([progress]);
    if (progressLookup[gId]) {
      updatedCount++;
    }
  }

  // Batch write progress values — resolve from header (v4.51.1)
  if (progressValues.length > 0) {
    var cpCol = resolveColumnByHeader_(grievanceSheet, 'Checklist Progress', GRIEVANCE_COLS.CHECKLIST_PROGRESS);
    grievanceSheet.getRange(2, cpCol, progressValues.length, 1)
      .setValues(progressValues);
  }

  log_('Synced checklist progress', updatedCount + ' cases updated');
  return { success: true, updatedCount: updatedCount };
}
/**
 * ============================================================================
 * FEATURE EXTENSION: DYNAMIC ENGINE & MEMBER LEADERS
 * ============================================================================
 * This script handles dynamic column detection and organizational leadership.
 *
 * OPTIMIZATIONS:
 * - CacheService for header maps (5-minute TTL)
 * - Single sheet read for multiple operations
 * - Batch writes using setValues()
 * - Pre-computed column indices
 * - Minimized API calls via shared spreadsheet reference
 *
 * FEATURES:
 * ---------
 * 1. Member Leaders (Organizational Layer)
 *    - Creates a new hierarchy level for organizational "nodes"
 *    - Member Leaders connect members without case-management duties
 *    - Filtered out of Grievance dropdowns via IS_STEWARD column
 *    - Uses "Member Leader" value in IS_STEWARD column (column O)
 *
 * 2. Column Expansion (Dynamic Engine)
 *    - "No-Code" feature for adding custom columns
 *    - Header Scanner performs census of top row on dashboard open
 *    - Auto-generates form fields with matching CSS styling
 *    - Detection Zone: columns beyond Core data (Column 32+)
 *
 * 3. Self-Healing Hidden Architecture
 *    - Script-to-Sheet Bridge for formula management
 *    - Uses dedicated section in _Dashboard_Calc (row 50+)
 *    - Repair Trigger re-injects formulas if broken
 *    - Dynamic Referencing uses Col1, Col2 logic (not cell refs)
 *
 * INTEGRATION NOTES:
 * - Uses MEMBER_COLS from 01_Constants.gs for column positions
 * - Uses SHEETS constant for sheet names
 * - Integrates with getAllStewards() via getStewardsForGrievance()
 */

const EXTENSION_CONFIG = {
  HIDDEN_CALC_SHEET: "_Dashboard_Calc",
  DYNAMIC_FORMULA_ROW: 50,
  MEMBER_SHEET: SHEETS.MEMBER_DIR,
  GRIEVANCE_SHEET: SHEETS.GRIEVANCE_LOG,
  LEADER_ROLE_NAME: "Member Leader",
  CORE_COLUMN_COUNT: 32,
  CACHE_TTL_SECONDS: 300 // 5-minute cache for header maps
};

// Pre-computed column indices (0-based for array access)
// Uses MEMBER_COLS from 01_Core.gs (always loaded first in GAS alphabetical order)
const COL_IDX = {
  MEMBER_ID: MEMBER_COLS.MEMBER_ID - 1,
  FIRST_NAME: MEMBER_COLS.FIRST_NAME - 1,
  LAST_NAME: MEMBER_COLS.LAST_NAME - 1,
  EMAIL: MEMBER_COLS.EMAIL - 1,
  UNIT: MEMBER_COLS.UNIT - 1,
  WORK_LOCATION: MEMBER_COLS.WORK_LOCATION - 1,
  IS_STEWARD: MEMBER_COLS.IS_STEWARD - 1
};

// ============================================================================
// CACHING LAYER
// ============================================================================

/**
 * Gets cached header map or builds fresh one.
 * Uses CacheService for 5-minute TTL to reduce sheet reads.
 * @param {string} sheetName - Sheet to scan
 * @param {boolean} forceRefresh - Force cache refresh
 * @returns {Object} Header name -> column index map
 */
function getHeaderMap(sheetName, forceRefresh) {
  const cacheKey = 'headerMap_' + sheetName.replace(/\s/g, '_');
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (_e) { log_('getHeaderMap', 'Error parsing cached header map: ' + (_e.message || _e)); }
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {};

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = {};

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header && header.toString().trim() !== "") {
      headerMap[header.toString().trim()] = i + 1;
    }
  }

  // Cache for 5 minutes
  try {
    cache.put(cacheKey, JSON.stringify(headerMap), EXTENSION_CONFIG.CACHE_TTL_SECONDS);
  } catch (_e) { log_('getHeaderMap', 'Error writing cache: ' + (_e.message || _e)); }

  return headerMap;
}

// ============================================================================
// UNIFIED DATA LOADER
// ============================================================================

/**
 * Loads member directory data once and returns parsed results.
 * Optimized single-read function for multiple queries.
 * @param {Object} options - Query options
 * @returns {Object} Parsed member data with leaders, stewards, and full dataset
 */
function loadMemberData_(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEETS.MEMBER_DIR;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { leaders: [], stewards: [], data: [], headers: [], found: null };
  }

  const range = sheet.getDataRange();
  const data = range.getValues();

  if (data.length < 2) {
    return { leaders: [], stewards: [], data: data, headers: data[0] || [], found: null };
  }

  const headers = data[0];
  const leaders = [];
  const stewards = [];
  let foundMember = null;

  // Single pass through data
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const roleValue = row[COL_IDX.IS_STEWARD];
    const roleStr = roleValue ? roleValue.toString().trim() : '';
    const firstName = row[COL_IDX.FIRST_NAME] || '';
    const lastName = row[COL_IDX.LAST_NAME] || '';
    const fullName = (firstName + ' ' + lastName).trim();
    const memberId = row[COL_IDX.MEMBER_ID] || '';

    // Check for specific member lookup
    if (options.findMemberId && memberId === options.findMemberId) {
      foundMember = { row: i + 1, data: row, roleValue: roleStr };
    }

    // Categorize by role
    if (roleStr === EXTENSION_CONFIG.LEADER_ROLE_NAME) {
      if (fullName) {
        leaders.push({
          name: fullName,
          firstName: firstName,
          lastName: lastName,
          memberId: memberId,
          email: row[COL_IDX.EMAIL] || '',
          unit: row[COL_IDX.UNIT] || '',
          location: row[COL_IDX.WORK_LOCATION] || '',
          rowNumber: i + 1
        });
      }
    } else if (isTruthyValue(roleStr) || isTruthyValue(roleValue)) {
      const steward = { rowNumber: i + 1, fullName: fullName };
      for (let j = 0; j < headers.length; j++) {
        steward[headers[j]] = row[j];
      }
      stewards.push(steward);
    }
  }

  return {
    leaders: leaders,
    stewards: stewards,
    data: data,
    headers: headers,
    found: foundMember
  };
}

// ============================================================================
// PUBLIC API - MEMBER LEADERS
// ============================================================================
/**
 * Gets stewards for grievance dropdowns, EXCLUDING Member Leaders.
 * Use this instead of getAllStewards() when populating grievance assignment dropdowns.
 * @returns {Array} Array of steward objects (excludes Member Leaders)
 */
function getStewardsForGrievance() {
  return loadMemberData_().stewards;
}

// ============================================================================
// SELF-HEALING DYNAMIC FORMULAS
// ============================================================================

/**
 * Re-injects logic into a dedicated section of the hidden sheet.
 * Uses batch writes for performance.
 * @returns {Object} Result with success status
 */
function repairDynamicFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);

  if (!calcSheet) {
    calcSheet = ss.insertSheet(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
    setSheetVeryHidden_(calcSheet);
  }

  const startRow = EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW;
  const isStewardCol = MEMBER_COLS.IS_STEWARD;
  const colLetter = getColumnLetter(isStewardCol);

  // Build all values and formulas for batch write
  // Use safeSheetNameForFormula to prevent formula injection attacks
  const safeSheetName = safeSheetNameForFormula(SHEETS.MEMBER_DIR);
  const safeLeaderRole = String(EXTENSION_CONFIG.LEADER_ROLE_NAME || '').replace(/"/g, '""');

  const batchData = [
    ['=== DYNAMIC ENGINE FORMULAS ===', ''],
    ['Role Distribution:', ''],
    [`=QUERY(${safeSheetName}!A:ZZ, "SELECT Col${isStewardCol}, count(Col1) WHERE Col1 IS NOT NULL GROUP BY Col${isStewardCol} LABEL count(Col1) 'Total'", 1)`, ''],
    ['', ''],
    ['', ''],
    ['', ''],
    ['Member Leaders:', `=COUNTIF(${safeSheetName}!${colLetter}:${colLetter},"${safeLeaderRole}")`],
    ['Active Stewards:', `=COUNTIF(${safeSheetName}!${colLetter}:${colLetter},"Yes")`]
  ];

  // Single batch write
  calcSheet.getRange(startRow, 1, batchData.length, 2).setValues(batchData);

  ss.toast('Dynamic formulas injected at row ' + startRow, 'Self-Healing Complete', 5);

  return { success: true, startRow: startRow };
}

// ============================================================================
// COLUMN EXPANSION ENGINE
// ============================================================================

/**
 * Gets expansion column data with optional member data.
 * Optimized to minimize sheet reads.
 * Integration tests: see __tests__/12_Features.test.js
 * @param {string} memberId - Optional member ID to fetch data for
 * @returns {Object} Object with extraHeaders and memberData
 */
function getExpansionColumnData(memberId) {
  const sheetName = SHEETS.MEMBER_DIR;
  const headerMap = getHeaderMap(sheetName);
  const coreCount = EXTENSION_CONFIG.CORE_COLUMN_COUNT;

  // Identify extra columns (beyond core)
  const extraHeaders = [];
  const headerEntries = Object.entries(headerMap);

  for (let i = 0; i < headerEntries.length; i++) {
    const [name, col] = headerEntries[i];
    if (col > coreCount) {
      extraHeaders.push({
        name: name,
        column: col,
        letter: getColumnLetter(col)
      });
    }
  }

  // Sort by column position
  extraHeaders.sort((a, b) => a.column - b.column);

  // Fetch member data if requested and there are extra columns
  const memberData = {};
  if (memberId && extraHeaders.length > 0) {
    const result = loadMemberData_({ findMemberId: memberId });
    if (result.found) {
      for (let i = 0; i < extraHeaders.length; i++) {
        const extra = extraHeaders[i];
        memberData[extra.name] = result.found.data[extra.column - 1] || '';
      }
    }
  }

  return {
    extraHeaders: extraHeaders,
    memberData: memberData,
    coreColumnCount: coreCount
  };
}

// ============================================================================
// SETUP & CONFIGURATION
// ============================================================================

/**
 * Adds "Member Leader" option to the IS_STEWARD validation dropdown.
 * @returns {Object} Result with success status
 */
function setupMemberLeaderRole() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Add "Member Leader" directly to IS_STEWARD validation (not Config column E).
  // IS_STEWARD uses hardcoded validation; Config column E is reserved for INTEREST_* columns.
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory not found. Run CREATE_DASHBOARD first.');
    return errorResponse('Member Directory not found');
  }

  const leaderValues = ['Yes', 'No', EXTENSION_CONFIG.LEADER_ROLE_NAME];
  const isRange = memberSheet.getRange(2, MEMBER_COLS.IS_STEWARD, Math.max(1, memberSheet.getMaxRows() - 1), 1);
  const isRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(leaderValues, true)
    .setAllowInvalid(true)
    .build();
  isRange.setDataValidation(isRule);

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
      action: 'MEMBER_LEADER_SETUP',
      addedValue: EXTENSION_CONFIG.LEADER_ROLE_NAME,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  ss.toast('Member Leader role added to IS_STEWARD validation.', 'Setup Complete', 5);
  return { success: true };
}

// ============================================================================
// GRIEVANCE REMINDERS FEATURE
// ============================================================================
// Allows users to set two reminder dates with notes for meetings/follow-ups

/**
 * Gets all grievances with reminders due within specified days.
 * Optimized single-pass through grievance data.
 * @param {number} daysAhead - Number of days to look ahead (default 3)
 * @returns {Array} Array of grievances with due reminders
 */
function getDueReminders(daysAhead) {
  daysAhead = daysAhead || 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() + daysAhead);

  const dueReminders = [];

  // Pre-compute column indices
  const colGrievanceId = GRIEVANCE_COLS.GRIEVANCE_ID - 1;
  const colFirstName = GRIEVANCE_COLS.FIRST_NAME - 1;
  const colLastName = GRIEVANCE_COLS.LAST_NAME - 1;
  const colStatus = GRIEVANCE_COLS.STATUS - 1;
  const colSteward = GRIEVANCE_COLS.STEWARD - 1;
  const colR1Date = GRIEVANCE_COLS.REMINDER_1_DATE - 1;
  const colR1Note = GRIEVANCE_COLS.REMINDER_1_NOTE - 1;
  const colR2Date = GRIEVANCE_COLS.REMINDER_2_DATE - 1;
  const colR2Note = GRIEVANCE_COLS.REMINDER_2_NOTE - 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const grievanceId = row[colGrievanceId];
    const status = row[colStatus];

    // Skip closed/resolved grievances
    if (!grievanceId || GRIEVANCE_CLOSED_STATUSES.indexOf(status) !== -1) {
      continue;
    }

    const memberName = ((row[colFirstName] || '') + ' ' + (row[colLastName] || '')).trim();
    const steward = row[colSteward] || '';

    // Check reminder 1
    const r1Date = row[colR1Date];
    if (r1Date instanceof Date) {
      const r1DateNorm = new Date(r1Date);
      r1DateNorm.setHours(0, 0, 0, 0);

      if (r1DateNorm >= today && r1DateNorm <= cutoffDate) {
        const daysUntil = Math.ceil((r1DateNorm - today) / (1000 * 60 * 60 * 24));
        dueReminders.push({
          grievanceId: grievanceId,
          memberName: memberName,
          status: status,
          steward: steward,
          reminderNum: 1,
          date: r1Date,
          dateFormatted: Utilities.formatDate(r1Date, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          note: row[colR1Note] || '',
          daysUntil: daysUntil,
          isToday: daysUntil === 0,
          isOverdue: daysUntil < 0,
          rowNumber: i + 1
        });
      }
    }

    // Check reminder 2
    const r2Date = row[colR2Date];
    if (r2Date instanceof Date) {
      const r2DateNorm = new Date(r2Date);
      r2DateNorm.setHours(0, 0, 0, 0);

      if (r2DateNorm >= today && r2DateNorm <= cutoffDate) {
        const daysUntil = Math.ceil((r2DateNorm - today) / (1000 * 60 * 60 * 24));
        dueReminders.push({
          grievanceId: grievanceId,
          memberName: memberName,
          status: status,
          steward: steward,
          reminderNum: 2,
          date: r2Date,
          dateFormatted: Utilities.formatDate(r2Date, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          note: row[colR2Note] || '',
          daysUntil: daysUntil,
          isToday: daysUntil === 0,
          isOverdue: daysUntil < 0,
          rowNumber: i + 1
        });
      }
    }
  }

  // Sort by date (soonest first)
  dueReminders.sort(function(a, b) { return a.date - b.date; });

  return dueReminders;
}

// ============================================================================
// LOOKER STUDIO CONFIGURATION
// ============================================================================

/**
 * Looker Studio integration configuration.
 * Defines sheet names for data sources and anonymized exports.
 */
var LOOKER_CONFIG = {
  // Sheet names for Looker data sources (restricted to 3 source sheets)
  SHEETS: {
    GRIEVANCES: '_Looker_Grievances',
    MEMBERS: '_Looker_Members',
    SATISFACTION: '_Looker_Satisfaction'
  },
  // PII-FREE sheets for external/compliance use (no personally identifiable information)
  SHEETS_ANON: {
    GRIEVANCES: '_Looker_Anon_Grievances',
    MEMBERS: '_Looker_Anon_Members',
    SATISFACTION: '_Looker_Anon_Satisfaction'
  },
  // Source sheets (only these are allowed)
  ALLOWED_SOURCES: ['Member Directory', 'Grievance Log', 'Member Satisfaction'],
  // Refresh settings
  AUTO_REFRESH_HOUR: 6 // 6 AM daily refresh
};

// ============================================================================
// CATEGORIZATION / BUCKET HELPERS
// ============================================================================

/**
 * Categorizes days into buckets for privacy.
 * @param {number} days
 * @returns {string}
 * @private
 */
function getDaysBucket_(days) {
  if (days <= 7) return 'Within Week';
  if (days <= 30) return 'Within Month';
  if (days <= 90) return '1-3 Months';
  if (days <= 180) return '3-6 Months';
  if (days <= 365) return '6-12 Months';
  return 'Over 1 Year';
}

/**
 * Categorizes contact frequency based on days since last contact.
 * @param {Date|*} lastContact - Last contact date
 * @param {Date} now - Current date
 * @returns {string}
 * @private
 */
function getContactFrequencyCategory_(lastContact, now) {
  if (!(lastContact instanceof Date)) return 'No Contact';
  const days = Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24));
  if (days <= 30) return 'Active';
  if (days <= 90) return 'Regular';
  if (days <= 180) return 'Occasional';
  return 'Inactive';
}

/**
 * Calculates engagement level from multiple factors.
 * @param {number} volunteerHours
 * @param {Date|null} lastContact
 * @param {string} isSteward - 'Yes', 'No', or 'Member Leader'
 * @param {number} grievanceCount
 * @returns {string}
 * @private
 */
function getEngagementLevel_(volunteerHours, lastContact, isSteward, grievanceCount) {
  let score = 0;

  // Volunteer hours contribution
  if (volunteerHours > 50) score += 3;
  else if (volunteerHours > 20) score += 2;
  else if (volunteerHours > 5) score += 1;

  // Recent contact contribution
  if (lastContact instanceof Date) {
    const days = Math.ceil((new Date() - lastContact) / (1000 * 60 * 60 * 24));
    if (days <= 30) score += 2;
    else if (days <= 90) score += 1;
  }

  // Leadership role
  if (isTruthyValue(isSteward) || isSteward === 'Member Leader') score += 2;

  // Grievance involvement
  if (grievanceCount > 0) score += 1;

  // Map to engagement level
  if (score >= 6) return 'Highly Engaged';
  if (score >= 4) return 'Engaged';
  if (score >= 2) return 'Somewhat Engaged';
  if (score >= 1) return 'Low Engagement';
  return 'Not Engaged';
}

/**
 * Gets quarter string from date.
 * @param {Date} date
 * @returns {string} e.g. '2026-Q1'
 * @private
 */
function getQuarter_(date) {
  const month = date.getMonth();
  const quarter = Math.floor(month / 3) + 1;
  return date.getFullYear() + '-Q' + quarter;
}

/**
 * Gets outcome category from grievance status.
 * @param {string} status
 * @returns {string}
 * @private
 */
function getOutcomeCategory_(status) {
  if (status === GRIEVANCE_STATUS.WON) return 'Win';
  if (status === GRIEVANCE_STATUS.DENIED) return 'Loss';
  if (status === GRIEVANCE_STATUS.SETTLED) return 'Settlement';
  if (status === GRIEVANCE_STATUS.WITHDRAWN) return 'Withdrawn';
  if (status === GRIEVANCE_STATUS.CLOSED) return 'Closed';
  return 'Active';
}

// ============================================================================
// TRIGGERS & AUTOMATION
// ============================================================================
