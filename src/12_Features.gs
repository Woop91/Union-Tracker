/**
 * 509 Dashboard - Checklist Manager
 *
 * Manages case checklists for grievances and other action types.
 * Provides functions for creating, updating, and tracking checklist items.
 *
 * @version 4.3.0
 * @license Free for use by non-profit collective bargaining groups and unions
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

  // Set up headers
  var headers = [
    'Checklist ID',
    'Case ID',
    'Action Type',
    'Item',
    'Category',
    'Required',
    'Completed',
    'Completed By',
    'Completed Date',
    'Due Date',
    'Notes',
    'Sort Order'
  ];

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
    return { success: false, error: 'Case ID is required' };
  }

  // Default to Grievance if not specified
  actionType = actionType || 'Grievance';

  // Get the template items
  var templateItems = getChecklistTemplate(actionType, issueCategory);

  if (!templateItems || templateItems.length === 0) {
    return { success: false, error: 'No template found for this action type' };
  }

  var sheet = getOrCreateChecklistSheet();
  var itemsCreated = [];

  // Prepare rows to add
  var rows = [];
  for (var i = 0; i < templateItems.length; i++) {
    var item = templateItems[i];
    var checklistId = generateChecklistId_(i);

    var row = [];
    row[CHECKLIST_COLS.CHECKLIST_ID - 1] = checklistId;
    row[CHECKLIST_COLS.CASE_ID - 1] = caseId;
    row[CHECKLIST_COLS.ACTION_TYPE - 1] = actionType;
    row[CHECKLIST_COLS.ITEM_TEXT - 1] = item.text;
    row[CHECKLIST_COLS.CATEGORY - 1] = item.category;
    row[CHECKLIST_COLS.REQUIRED - 1] = item.required ? 'Yes' : 'No';
    row[CHECKLIST_COLS.COMPLETED - 1] = false;
    row[CHECKLIST_COLS.COMPLETED_BY - 1] = '';
    row[CHECKLIST_COLS.COMPLETED_DATE - 1] = '';
    row[CHECKLIST_COLS.DUE_DATE - 1] = '';
    row[CHECKLIST_COLS.NOTES - 1] = '';
    row[CHECKLIST_COLS.SORT_ORDER - 1] = i + 1;

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
    var lastRow = sheet.getLastRow();
    var startRow = lastRow + 1;
    sheet.getRange(startRow, 1, rows.length, 12).setValues(rows);

    // Add checkboxes to the new rows
    sheet.getRange(startRow, CHECKLIST_COLS.COMPLETED, rows.length, 1).insertCheckboxes();
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

  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[CHECKLIST_COLS.CASE_ID - 1] === caseId) {
      items.push({
        checklistId: row[CHECKLIST_COLS.CHECKLIST_ID - 1],
        caseId: row[CHECKLIST_COLS.CASE_ID - 1],
        actionType: row[CHECKLIST_COLS.ACTION_TYPE - 1],
        itemText: row[CHECKLIST_COLS.ITEM_TEXT - 1],
        category: row[CHECKLIST_COLS.CATEGORY - 1],
        required: row[CHECKLIST_COLS.REQUIRED - 1] === 'Yes',
        completed: row[CHECKLIST_COLS.COMPLETED - 1] === true,
        completedBy: row[CHECKLIST_COLS.COMPLETED_BY - 1],
        completedDate: row[CHECKLIST_COLS.COMPLETED_DATE - 1],
        dueDate: row[CHECKLIST_COLS.DUE_DATE - 1],
        notes: row[CHECKLIST_COLS.NOTES - 1],
        sortOrder: row[CHECKLIST_COLS.SORT_ORDER - 1],
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
    return { success: false, error: 'Checklist item not found' };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][CHECKLIST_COLS.CHECKLIST_ID - 1] === checklistId) {
      var row = i + 2;
      var caseId = data[i][CHECKLIST_COLS.CASE_ID - 1];

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

  return { success: false, error: 'Checklist item not found' };
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
    return { success: false, error: 'Case ID and item text are required' };
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
          actionType = grievanceData[j][GRIEVANCE_COLS.ACTION_TYPE - 1] || 'Grievance';
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
    return { success: false, error: 'Checklist item not found' };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === checklistId) {
      var caseId = data[i][1];
      var row = i + 2;
      sheet.deleteRow(row);

      // Update progress
      updateChecklistProgress(caseId);

      return { success: true, caseId: caseId };
    }
  }

  return { success: false, error: 'Checklist item not found' };
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
    return { success: false, error: 'Checklist item not found' };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][CHECKLIST_COLS.CHECKLIST_ID - 1] === checklistId) {
      var row = i + 2;

      if (updates.itemText !== undefined) {
        sheet.getRange(row, CHECKLIST_COLS.ITEM_TEXT).setValue(updates.itemText);
      }
      if (updates.category !== undefined) {
        sheet.getRange(row, CHECKLIST_COLS.CATEGORY).setValue(updates.category);
      }
      if (updates.required !== undefined) {
        sheet.getRange(row, CHECKLIST_COLS.REQUIRED).setValue(updates.required ? 'Yes' : 'No');
      }
      if (updates.dueDate !== undefined) {
        sheet.getRange(row, CHECKLIST_COLS.DUE_DATE).setValue(updates.dueDate);
      }
      if (updates.notes !== undefined) {
        sheet.getRange(row, CHECKLIST_COLS.NOTES).setValue(updates.notes);
      }

      return { success: true };
    }
  }

  return { success: false, error: 'Checklist item not found' };
}

/**
 * Delete all checklist items for a case
 * @param {string} caseId - The case/grievance ID
 * @returns {Object} Result with count of deleted items
 */
function deleteAllChecklistItems(caseId) {
  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { success: true, deletedCount: 0 };
  }

  var data = sheet.getRange(2, CHECKLIST_COLS.CASE_ID, lastRow - 1, 1).getValues();
  var rowsToDelete = [];

  // Find rows to delete (in reverse order to avoid index shifting)
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === caseId) {
      rowsToDelete.push(i + 2);
    }
  }

  // Delete rows in reverse order
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  return { success: true, deletedCount: rowsToDelete.length };
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

/**
 * Get all available checklist templates for a given action type
 * @param {string} actionType - The action type
 * @returns {Array} Array of template names
 */
function getAvailableTemplates(actionType) {
  var templates = CHECKLIST_TEMPLATES[actionType];
  if (!templates) {
    return ['_default'];
  }
  return Object.keys(templates);
}

// ============================================================================
// BATCH OPERATIONS
// ============================================================================

/**
 * Update checklist progress for all open cases
 * Useful for maintenance/sync operations
 */
function updateAllChecklistProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    return { success: false, error: 'Grievance Log not found' };
  }

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, updatedCount: 0 };
  }

  var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, lastRow - 1, 1).getValues();
  var updatedCount = 0;

  for (var i = 0; i < grievanceIds.length; i++) {
    var caseId = grievanceIds[i][0];
    if (caseId) {
      updateChecklistProgress(caseId);
      updatedCount++;
    }

    // Pause every 50 items to avoid timeout
    if (updatedCount % 50 === 0) {
      Utilities.sleep(100);
    }
  }

  return { success: true, updatedCount: updatedCount };
}

/**
 * Create checklists for all existing grievances that don't have one
 * Useful for migrating existing data
 */
function createChecklistsForExistingCases() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    return { success: false, error: 'Grievance Log not found' };
  }

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, createdCount: 0 };
  }

  var data = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.CHECKLIST_PROGRESS).getValues();
  var createdCount = 0;

  for (var i = 0; i < data.length; i++) {
    var caseId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var actionType = data[i][GRIEVANCE_COLS.ACTION_TYPE - 1] || 'Grievance';
    var issueCategory = data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
    var checklistProgress = data[i][GRIEVANCE_COLS.CHECKLIST_PROGRESS - 1];

    // Skip if already has a checklist
    if (checklistProgress && checklistProgress !== 'No checklist') {
      continue;
    }

    // Check if items exist
    var existingItems = getChecklistItems(caseId);
    if (existingItems.length > 0) {
      // Just update the progress display
      updateChecklistProgress(caseId);
      continue;
    }

    // Create checklist from template
    var result = createChecklistFromTemplate(caseId, actionType, issueCategory);
    if (result.success) {
      createdCount++;
    }

    // Pause every 10 items to avoid timeout
    if (createdCount % 10 === 0) {
      Utilities.sleep(100);
    }
  }

  return { success: true, createdCount: createdCount };
}

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

  var completed = e.value === 'TRUE' || e.value === true;
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

  var html = HtmlService.createHtmlOutput(getChecklistDialogHtml(caseId))
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Case Checklist: ' + caseId);
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
      itemsHtml += '<div class="category-header">' + cat + ' (' + catItems.length + ')</div>';

      for (var i = 0; i < catItems.length; i++) {
        var item = catItems[i];
        var checkedAttr = item.completed ? 'checked' : '';
        var completedClass = item.completed ? 'completed' : '';
        var requiredBadge = item.required ? '<span class="badge required">Required</span>' : '';

        itemsHtml += '<div class="checklist-item ' + completedClass + '">';
        itemsHtml += '  <label class="checkbox-container">';
        itemsHtml += '    <input type="checkbox" ' + checkedAttr + ' onchange="toggleItem(\'' + item.checklistId + '\', this.checked)">';
        itemsHtml += '    <span class="checkmark"></span>';
        itemsHtml += '  </label>';
        itemsHtml += '  <div class="item-content">';
        itemsHtml += '    <div class="item-text">' + item.itemText + ' ' + requiredBadge + '</div>';
        if (item.completed && item.completedBy) {
          itemsHtml += '    <div class="item-meta">Completed by ' + item.completedBy + '</div>';
        }
        itemsHtml += '  </div>';
        itemsHtml += '  <button class="btn-icon" onclick="deleteItem(\'' + item.checklistId + '\')" title="Delete item">x</button>';
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
    'var caseId = "' + caseId + '";' +
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
    '    .createChecklistForSelectedCase();' +
    '}' +
    '</script>' +
    '</body></html>';
}

/**
 * Create checklist for the currently selected case
 * Called from the checklist dialog
 */
function createChecklistForSelectedCase() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();

  if (row <= 1) {
    return { success: false, error: 'No case selected' };
  }

  var caseId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  var actionType = sheet.getRange(row, GRIEVANCE_COLS.ACTION_TYPE).getValue() || 'Grievance';
  var issueCategory = sheet.getRange(row, GRIEVANCE_COLS.ISSUE_CATEGORY).getValue();

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
    return { success: false, error: 'Grievance Log not found' };
  }

  // Check if Action Type header exists
  var headers = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues()[0];
  var actionTypeCol = headers.indexOf('Action Type') + 1;

  if (actionTypeCol === 0) {
    // Add the column header
    grievanceSheet.getRange(1, GRIEVANCE_COLS.ACTION_TYPE).setValue('Action Type');
    grievanceSheet.getRange(1, GRIEVANCE_COLS.CHECKLIST_PROGRESS).setValue('Checklist Progress');
    actionTypeCol = GRIEVANCE_COLS.ACTION_TYPE;
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
    for (var i = 0; i < existingValues.length; i++) {
      if (!existingValues[i][0]) {
        grievanceSheet.getRange(i + 2, actionTypeCol).setValue('Grievance');
      }
    }
  }

  return { success: true };
}

/**
 * Get a summary of cases by action type
 * @returns {Object} Summary with counts by action type
 */
function getCasesByActionType() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    return {};
  }

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  var data = grievanceSheet.getRange(2, GRIEVANCE_COLS.ACTION_TYPE, lastRow - 1, 1).getValues();
  var statusData = grievanceSheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1).getValues();

  var summary = {};
  var openStatuses = ['Open', 'Pending Info', 'Appealed', 'In Arbitration'];

  for (var i = 0; i < data.length; i++) {
    var actionType = data[i][0] || 'Grievance';
    var status = statusData[i][0];

    if (!summary[actionType]) {
      summary[actionType] = { total: 0, open: 0, closed: 0 };
    }

    summary[actionType].total++;

    if (openStatuses.indexOf(status) !== -1) {
      summary[actionType].open++;
    } else {
      summary[actionType].closed++;
    }
  }

  return summary;
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

/**
 * Re-protect the Checklist sheet (keeps data rows editable)
 * Only protects the header row
 */
function reprotectChecklistSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECKLIST_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', 'Checklist sheet not found', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // First remove any existing protections
  var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  existingProtections.forEach(function(p) { if (p.canEdit()) p.remove(); });

  // Add protection that only protects header row
  var protection = sheet.protect().setDescription('Checklist Sheet Structure');
  var lastCol = sheet.getLastColumn() || 12;
  var lastRow = Math.max(sheet.getLastRow(), 1000);

  // Allow editing everything except row 1 (header)
  protection.setUnprotectedRanges([sheet.getRange(2, 1, lastRow - 1, lastCol)]);

  SpreadsheetApp.getUi().alert(
    'Checklist Sheet Protected',
    'The header row is now protected. All data rows remain editable.',
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

  // Column letters for checklist columns
  var caseIdCol = 'B';      // CHECKLIST_COLS.CASE_ID = 2
  var completedCol = 'G';   // CHECKLIST_COLS.COMPLETED = 7
  var requiredCol = 'F';    // CHECKLIST_COLS.REQUIRED = 6

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

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Checklist_Calc sheet setup complete with self-healing formulas');
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
    Logger.log('Grievance Log not found');
    return { success: false, error: 'Grievance Log not found' };
  }

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

  // Batch write progress values
  if (progressValues.length > 0) {
    grievanceSheet.getRange(2, GRIEVANCE_COLS.CHECKLIST_PROGRESS, progressValues.length, 1)
      .setValues(progressValues);
  }

  Logger.log('Synced checklist progress: ' + updatedCount + ' cases updated');
  return { success: true, updatedCount: updatedCount };
}

/**
 * Repair/rebuild the checklist calc sheet
 * Called by repairAllHiddenSheets()
 */
function repairChecklistCalcSheet() {
  return setupChecklistCalcSheet();
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
 *    - Uses "Member Leader" value in IS_STEWARD column (column N)
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
  MEMBER_SHEET: "Member Directory",
  GRIEVANCE_SHEET: "Grievance Log",
  LEADER_ROLE_NAME: "Member Leader",
  CORE_COLUMN_COUNT: 32,
  CACHE_TTL_SECONDS: 300 // 5-minute cache for header maps
};

// Pre-computed column indices (0-based for array access)
const COL_IDX = {
  MEMBER_ID: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.MEMBER_ID : 1) - 1,
  FIRST_NAME: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.FIRST_NAME : 2) - 1,
  LAST_NAME: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.LAST_NAME : 3) - 1,
  EMAIL: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.EMAIL : 8) - 1,
  UNIT: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.UNIT : 5) - 1,
  WORK_LOCATION: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.WORK_LOCATION : 6) - 1,
  IS_STEWARD: (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.IS_STEWARD : 14) - 1
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
      } catch (e) {
        // Cache corrupted, rebuild
      }
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
  } catch (e) {
    // Cache write failed, continue without caching
  }

  return headerMap;
}

/**
 * Invalidates header map cache for a sheet.
 * Call after structural changes (adding/removing columns).
 * @param {string} sheetName - Sheet to invalidate
 */
function invalidateHeaderCache(sheetName) {
  const cacheKey = 'headerMap_' + sheetName.replace(/\s/g, '_');
  CacheService.getScriptCache().remove(cacheKey);
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
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
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
    } else if (roleStr === 'Yes' || roleValue === true) {
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
 * Fetches all members tagged as 'Member Leader'.
 * These individuals are organizational and do not appear in Grievance lists.
 * @returns {Array} Array of leader objects
 */
function getMemberLeaders() {
  return loadMemberData_().leaders;
}

/**
 * Gets stewards for grievance dropdowns, EXCLUDING Member Leaders.
 * Use this instead of getAllStewards() when populating grievance assignment dropdowns.
 * @returns {Array} Array of steward objects (excludes Member Leaders)
 */
function getStewardsForGrievance() {
  return loadMemberData_().stewards;
}

/**
 * Gets both leaders and stewards in a single optimized call.
 * Use when you need both datasets to avoid duplicate sheet reads.
 * @returns {Object} { leaders: Array, stewards: Array }
 */
function getLeadersAndStewards() {
  const result = loadMemberData_();
  return { leaders: result.leaders, stewards: result.stewards };
}

/**
 * Checks if a member is a Member Leader (not a Steward).
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if member is a Member Leader
 */
function isMemberLeader(memberId) {
  if (!memberId) return false;
  const result = loadMemberData_({ findMemberId: memberId });
  return result.found && result.found.roleValue === EXTENSION_CONFIG.LEADER_ROLE_NAME;
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
    calcSheet.hideSheet();
  }

  const startRow = EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW;
  const isStewardCol = (typeof MEMBER_COLS !== 'undefined' ? MEMBER_COLS.IS_STEWARD : 14);
  const colLetter = typeof getColumnLetter === 'function' ? getColumnLetter(isStewardCol) : 'N';

  // Build all values and formulas for batch write
  // Use safeSheetNameForFormula to prevent formula injection attacks
  const safeSheetName = typeof safeSheetNameForFormula === 'function'
    ? safeSheetNameForFormula(EXTENSION_CONFIG.MEMBER_SHEET)
    : "'" + String(EXTENSION_CONFIG.MEMBER_SHEET).replace(/'/g, "''") + "'";
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
 * @param {string} memberId - Optional member ID to fetch data for
 * @returns {Object} Object with extraHeaders and memberData
 */
function getExpansionColumnData(memberId) {
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
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
        letter: typeof getColumnLetter === 'function' ? getColumnLetter(col) : ''
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

/**
 * Generates HTML form fields for custom columns.
 * Uses array join for efficient string building.
 * @param {string} memberId - Optional member ID to pre-fill values
 * @returns {string} HTML string with form fields
 */
function generateExpansionFieldsHtml(memberId) {
  const expansionData = getExpansionColumnData(memberId);

  if (expansionData.extraHeaders.length === 0) {
    return '<!-- No custom columns detected -->';
  }

  const parts = [
    '<div class="expansion-fields" style="margin-top:16px;padding-top:16px;border-top:1px solid #E5E7EB;">',
    '<div class="expansion-header" style="font-weight:600;color:#7C3AED;margin-bottom:12px;">Custom Fields</div>'
  ];

  const fields = expansionData.extraHeaders;
  for (let i = 0; i < fields.length; i++) {
    const field = fields[i];
    const value = (expansionData.memberData[field.name] || '').toString().replace(/"/g, '&quot;');
    const fieldId = 'custom_' + field.name.replace(/[^a-zA-Z0-9]/g, '_');

    parts.push(
      '<div class="form-group" style="margin-bottom:12px;">',
      '<label class="form-label" style="display:block;font-size:13px;color:#374151;margin-bottom:4px;">',
      field.name,
      '</label>',
      '<input type="text" class="form-input" id="', fieldId, '" name="', fieldId,
      '" data-column="', field.column, '" value="', value,
      '" style="width:100%;padding:8px 12px;border:1px solid #D1D5DB;border-radius:6px;font-family:Roboto,sans-serif;">',
      '</div>'
    );
  }

  parts.push('</div>');
  return parts.join('');
}

/**
 * Saves custom column data for a member using batch write.
 * @param {string} memberId - The member ID
 * @param {Object} customData - Object with field names as keys
 * @returns {Object} Result with success status
 */
function saveExpansionData(memberId, customData) {
  if (!memberId || !customData) {
    return { success: false, error: 'Member ID and data required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) || EXTENSION_CONFIG.MEMBER_SHEET;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, error: 'Sheet not found' };
  }

  const headerMap = getHeaderMap(sheetName);
  const result = loadMemberData_({ findMemberId: memberId });

  if (!result.found) {
    return { success: false, error: 'Member not found' };
  }

  const memberRow = result.found.row;
  const coreCount = EXTENSION_CONFIG.CORE_COLUMN_COUNT;

  // Collect updates for batch write
  const updates = [];
  const fieldNames = Object.keys(customData);

  for (let i = 0; i < fieldNames.length; i++) {
    const fieldName = fieldNames[i];
    const col = headerMap[fieldName];
    if (col && col > coreCount) {
      updates.push({ col: col, value: customData[fieldName] });
    }
  }

  if (updates.length === 0) {
    return { success: true, updated: 0 };
  }

  // Sort by column for potential range optimization
  updates.sort((a, b) => a.col - b.col);

  // Check if columns are contiguous for single batch write
  const minCol = updates[0].col;
  const maxCol = updates[updates.length - 1].col;

  if (maxCol - minCol + 1 === updates.length) {
    // Contiguous columns - single batch write
    const values = updates.map(u => u.value);
    sheet.getRange(memberRow, minCol, 1, values.length).setValues([values]);
  } else {
    // Non-contiguous - use individual writes (still fewer than per-field)
    for (let i = 0; i < updates.length; i++) {
      sheet.getRange(memberRow, updates[i].col).setValue(updates[i].value);
    }
  }

  return { success: true, updated: updates.length };
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
  const configSheetName = (typeof SHEETS !== 'undefined' && SHEETS.CONFIG) || 'Config';
  const configSheet = ss.getSheetByName(configSheetName);

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found');
    return { success: false, error: 'Config sheet not found' };
  }

  const yesNoCol = (typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.YES_NO) || 5;
  const existingValues = configSheet.getRange(3, yesNoCol, 10, 1).getValues().flat().filter(Boolean);

  if (existingValues.indexOf(EXTENSION_CONFIG.LEADER_ROLE_NAME) !== -1) {
    ss.toast('Member Leader role already configured', 'Setup Complete', 3);
    return { success: true, alreadyExists: true };
  }

  const nextRow = existingValues.length + 3;
  configSheet.getRange(nextRow, yesNoCol).setValue(EXTENSION_CONFIG.LEADER_ROLE_NAME);

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
      action: 'MEMBER_LEADER_SETUP',
      addedValue: EXTENSION_CONFIG.LEADER_ROLE_NAME,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  ss.toast('Member Leader role added to Config.', 'Setup Complete', 5);
  return { success: true, addedAt: nextRow };
}

/**
 * Runs all Dynamic Engine setup tasks.
 * @returns {Object} Result with setup details
 */
function setupDynamicEngine() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Dynamic Engine',
    'This will:\n\n' +
    '1. Add "Member Leader" to IS_STEWARD dropdown options\n' +
    '2. Inject self-healing formulas into _Dashboard_Calc\n' +
    '3. Scan for custom columns beyond Column AF\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  // Invalidate cache before setup
  invalidateHeaderCache(EXTENSION_CONFIG.MEMBER_SHEET);

  const roleResult = setupMemberLeaderRole();
  const formulaResult = repairDynamicFormulas();
  const expansionData = getExpansionColumnData();

  const summary =
    'Dynamic Engine Setup Complete:\n\n' +
    '• Member Leader Role: ' + (roleResult.success ? 'OK' : 'Failed') + '\n' +
    '• Self-Healing Formulas: ' + (formulaResult.success ? 'Injected at row ' + formulaResult.startRow : 'Failed') + '\n' +
    '• Custom Columns Found: ' + expansionData.extraHeaders.length;

  ui.alert('Setup Complete', summary, ui.ButtonSet.OK);

  return {
    success: true,
    roleSetup: roleResult,
    formulas: formulaResult,
    expansionColumns: expansionData.extraHeaders.length
  };
}

/**
 * Gets Dynamic Engine status for display in admin panels.
 * Optimized single-pass data collection.
 * @returns {Object} Status object with feature states
 */
function getDynamicEngineStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Single load for leaders
  const memberData = loadMemberData_();

  // Get expansion columns (uses cached header map)
  const expansionData = getExpansionColumnData();

  // Check self-healing status
  const calcSheet = ss.getSheetByName(EXTENSION_CONFIG.HIDDEN_CALC_SHEET);
  let formulasInjected = false;

  if (calcSheet) {
    const checkValue = calcSheet.getRange(EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW, 1).getValue();
    formulasInjected = checkValue && checkValue.toString().indexOf('DYNAMIC ENGINE') !== -1;
  }

  return {
    memberLeaders: {
      count: memberData.leaders.length,
      names: memberData.leaders.map(function(l) { return l.name; })
    },
    stewards: {
      count: memberData.stewards.length
    },
    expansionColumns: {
      count: expansionData.extraHeaders.length,
      headers: expansionData.extraHeaders.map(function(h) { return h.name; })
    },
    selfHealing: {
      active: formulasInjected,
      sheet: EXTENSION_CONFIG.HIDDEN_CALC_SHEET,
      row: EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW
    }
  };
}

// ============================================================================
// GRIEVANCE REMINDERS FEATURE
// ============================================================================
// Allows users to set two reminder dates with notes for meetings/follow-ups

/**
 * Sets a reminder for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @param {number} reminderNum - Which reminder (1 or 2)
 * @param {Date|string} reminderDate - The reminder date
 * @param {string} reminderNote - Brief note about the reminder
 * @returns {Object} Result with success status
 */
function setGrievanceReminder(grievanceId, reminderNum, reminderDate, reminderNote) {
  if (!grievanceId || (reminderNum !== 1 && reminderNum !== 2)) {
    return { success: false, error: 'Invalid grievance ID or reminder number' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  if (!sheet) {
    return { success: false, error: 'Grievance Log sheet not found' };
  }

  // Find the grievance row
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, error: 'Grievance not found: ' + grievanceId };
  }

  // Determine which columns to update
  const dateCol = reminderNum === 1 ? GRIEVANCE_COLS.REMINDER_1_DATE : GRIEVANCE_COLS.REMINDER_2_DATE;
  const noteCol = reminderNum === 1 ? GRIEVANCE_COLS.REMINDER_1_NOTE : GRIEVANCE_COLS.REMINDER_2_NOTE;

  // Parse date if string
  let parsedDate = reminderDate;
  if (typeof reminderDate === 'string' && reminderDate) {
    parsedDate = new Date(reminderDate);
    if (isNaN(parsedDate.getTime())) {
      return { success: false, error: 'Invalid date format' };
    }
  }

  // Batch update both columns
  const updates = [[parsedDate || '', reminderNote || '']];
  sheet.getRange(rowIndex, dateCol, 1, 2).setValues(updates);

  // Log the action
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      action: 'REMINDER_SET',
      reminderNum: reminderNum,
      reminderDate: parsedDate ? parsedDate.toISOString() : null,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, grievanceId: grievanceId, reminderNum: reminderNum };
}

/**
 * Gets reminders for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @returns {Object} Reminder data or null if not found
 */
function getGrievanceReminders(grievanceId) {
  if (!grievanceId) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      return {
        grievanceId: grievanceId,
        memberName: (data[i][GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' +
                    (data[i][GRIEVANCE_COLS.LAST_NAME - 1] || ''),
        status: data[i][GRIEVANCE_COLS.STATUS - 1] || '',
        reminder1: {
          date: data[i][GRIEVANCE_COLS.REMINDER_1_DATE - 1] || null,
          note: data[i][GRIEVANCE_COLS.REMINDER_1_NOTE - 1] || ''
        },
        reminder2: {
          date: data[i][GRIEVANCE_COLS.REMINDER_2_DATE - 1] || null,
          note: data[i][GRIEVANCE_COLS.REMINDER_2_NOTE - 1] || ''
        }
      };
    }
  }

  return null;
}

/**
 * Clears a reminder for a grievance.
 * @param {string} grievanceId - The grievance ID
 * @param {number} reminderNum - Which reminder to clear (1 or 2)
 * @returns {Object} Result with success status
 */
function clearGrievanceReminder(grievanceId, reminderNum) {
  return setGrievanceReminder(grievanceId, reminderNum, '', '');
}

/**
 * Gets all grievances with reminders due within specified days.
 * Optimized single-pass through grievance data.
 * @param {number} daysAhead - Number of days to look ahead (default 3)
 * @returns {Array} Array of grievances with due reminders
 */
function getDueReminders(daysAhead) {
  daysAhead = daysAhead || 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');

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
    if (!grievanceId || status === 'Closed' || status === 'Won' ||
        status === 'Denied' || status === 'Withdrawn' || status === 'Settled') {
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

/**
 * Shows the reminder management dialog for a grievance.
 * Can be called from Quick Actions or menu.
 * @param {string} grievanceId - Optional grievance ID (uses selected row if not provided)
 */
function showReminderDialog(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // If no grievanceId provided, try to get from selected row
  if (!grievanceId) {
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== (SHEETS.GRIEVANCE_LOG || 'Grievance Log')) {
      ui.alert('Please select a grievance in the Grievance Log sheet, or provide a Grievance ID.');
      return;
    }

    const row = sheet.getActiveRange().getRow();
    if (row <= 1) {
      ui.alert('Please select a grievance row (not the header).');
      return;
    }

    grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  }

  if (!grievanceId) {
    ui.alert('No grievance selected.');
    return;
  }

  // Get current reminders
  const reminders = getGrievanceReminders(grievanceId);
  if (!reminders) {
    ui.alert('Grievance not found: ' + grievanceId);
    return;
  }

  // Build the HTML dialog
  const html = buildReminderDialogHtml_(grievanceId, reminders);
  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  ui.showModalDialog(dialog, '⏰ Reminders: ' + grievanceId);
}

/**
 * Builds the HTML for the reminder dialog.
 * @private
 */
function buildReminderDialogHtml_(grievanceId, reminders) {
  const r1Date = reminders.reminder1.date instanceof Date
    ? Utilities.formatDate(reminders.reminder1.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';
  const r2Date = reminders.reminder2.date instanceof Date
    ? Utilities.formatDate(reminders.reminder2.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';

  return `<!DOCTYPE html>
<html>
<head>
  ${getMobileOptimizedHead()}
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Google Sans', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; padding: clamp(12px,3vw,20px); color: #F8FAFC; }
    .header { margin-bottom: 20px; }
    .header h2 { font-size: clamp(15px,4vw,18px); color: #A78BFA; margin-bottom: 4px; }
    .header .member { font-size: clamp(12px,3vw,14px); color: #94A3B8; }
    .reminder-card { background: rgba(255,255,255,0.05); border: 1px solid #334155; border-radius: 12px; padding: clamp(12px,3vw,16px); margin-bottom: 16px; }
    .reminder-card h3 { font-size: clamp(12px,3vw,14px); color: #7C3AED; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
    .reminder-card h3::before { content: '⏰'; }
    .form-group { margin-bottom: 12px; }
    .form-label { display: block; font-size: clamp(11px,2.8vw,12px); color: #94A3B8; margin-bottom: 4px; }
    .form-input { width: 100%; padding: 10px 12px; border: 1px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-size: 16px; min-height: 44px; }
    .form-input:focus { border-color: #7C3AED; outline: none; }
    .form-input::placeholder { color: #64748B; }
    .btn-row { display: flex; gap: 8px; margin-top: 8px; }
    .btn { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: clamp(12px,3vw,13px); font-weight: 500; transition: all 0.2s; min-height: 44px; }
    .btn-clear { background: #334155; color: #F8FAFC; }
    .btn-clear:hover { background: #475569; }
    .actions { display: flex; gap: 12px; margin-top: 20px; justify-content: flex-end; flex-wrap: wrap; }
    .btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; padding: 12px 24px; }
    .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }
    .btn-secondary { background: #334155; color: #F8FAFC; padding: 12px 24px; }
    .status { padding: 8px 12px; border-radius: 6px; margin-top: 12px; font-size: clamp(11px,2.8vw,13px); display: none; }
    .status.success { display: block; background: rgba(16,185,129,0.2); border: 1px solid #10B981; }
    .status.error { display: block; background: rgba(239,68,68,0.2); border: 1px solid #EF4444; }
    @media(max-width:480px) { .actions { flex-direction: column; } .actions .btn { width: 100%; text-align: center; } }
  </style>
</head>
<body>
  <div class="header">
    <h2>Grievance ${grievanceId}</h2>
    <div class="member">${reminders.memberName} • ${reminders.status}</div>
  </div>

  <div class="reminder-card">
    <h3>Reminder 1</h3>
    <div class="form-group">
      <label class="form-label">Date</label>
      <input type="date" class="form-input" id="r1Date" value="${r1Date}">
    </div>
    <div class="form-group">
      <label class="form-label">Note (e.g., "Schedule Step II meeting")</label>
      <input type="text" class="form-input" id="r1Note" value="${reminders.reminder1.note.replace(/"/g, '&quot;')}" placeholder="Brief description...">
    </div>
    <div class="btn-row">
      <button class="btn btn-clear" onclick="clearReminder(1)">Clear</button>
    </div>
  </div>

  <div class="reminder-card">
    <h3>Reminder 2</h3>
    <div class="form-group">
      <label class="form-label">Date</label>
      <input type="date" class="form-input" id="r2Date" value="${r2Date}">
    </div>
    <div class="form-group">
      <label class="form-label">Note</label>
      <input type="text" class="form-input" id="r2Note" value="${reminders.reminder2.note.replace(/"/g, '&quot;')}" placeholder="Brief description...">
    </div>
    <div class="btn-row">
      <button class="btn btn-clear" onclick="clearReminder(2)">Clear</button>
    </div>
  </div>

  <div id="statusArea"></div>

  <div class="actions">
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn btn-primary" onclick="saveReminders()">Save Reminders</button>
  </div>

  <script>
    var grievanceId = '${grievanceId}';

    function saveReminders() {
      var r1Date = document.getElementById('r1Date').value;
      var r1Note = document.getElementById('r1Note').value;
      var r2Date = document.getElementById('r2Date').value;
      var r2Note = document.getElementById('r2Note').value;

      showStatus('Saving...', false);

      // Save reminder 1
      google.script.run
        .withSuccessHandler(function() {
          // Save reminder 2
          google.script.run
            .withSuccessHandler(function() {
              showStatus('Reminders saved!', false);
              setTimeout(function() { google.script.host.close(); }, 1000);
            })
            .withFailureHandler(function(e) { showStatus('Error: ' + e.message, true); })
            .setGrievanceReminder(grievanceId, 2, r2Date, r2Note);
        })
        .withFailureHandler(function(e) { showStatus('Error: ' + e.message, true); })
        .setGrievanceReminder(grievanceId, 1, r1Date, r1Note);
    }

    function clearReminder(num) {
      if (num === 1) {
        document.getElementById('r1Date').value = '';
        document.getElementById('r1Note').value = '';
      } else {
        document.getElementById('r2Date').value = '';
        document.getElementById('r2Note').value = '';
      }
    }

    function showStatus(msg, isError) {
      var area = document.getElementById('statusArea');
      area.className = 'status ' + (isError ? 'error' : 'success');
      area.textContent = msg;
    }
  </script>
</body>
</html>`;
}

/**
 * Checks for due reminders and sends notifications.
 * Call this from a daily time-driven trigger.
 * @param {number} daysAhead - Days to look ahead (default 1 = today only)
 * @returns {Object} Result with notification count
 */
function checkAndNotifyReminders(daysAhead) {
  daysAhead = daysAhead || 1;

  const dueReminders = getDueReminders(daysAhead);

  if (dueReminders.length === 0) {
    return { success: true, notified: 0, message: 'No reminders due' };
  }

  // Group reminders by steward for consolidated notifications
  const bysteward = {};
  for (let i = 0; i < dueReminders.length; i++) {
    const reminder = dueReminders[i];
    const steward = reminder.steward || 'Unassigned';
    if (!bysteward[steward]) {
      bysteward[steward] = [];
    }
    bysteward[steward].push(reminder);
  }

  // Show toast summary
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayCount = dueReminders.filter(function(r) { return r.isToday; }).length;
  const upcomingCount = dueReminders.length - todayCount;

  let message = '';
  if (todayCount > 0) {
    message += todayCount + ' reminder(s) due TODAY';
  }
  if (upcomingCount > 0) {
    message += (message ? ', ' : '') + upcomingCount + ' upcoming';
  }

  ss.toast(message, '⏰ Grievance Reminders', 10);

  return {
    success: true,
    notified: dueReminders.length,
    todayCount: todayCount,
    upcomingCount: upcomingCount,
    reminders: dueReminders
  };
}

/**
 * Installs the daily reminder check trigger.
 * Call once during setup to enable automatic reminder notifications.
 */
function installReminderTrigger() {
  // Remove existing reminder triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndNotifyReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install new daily trigger at 8 AM
  ScriptApp.newTrigger('checkAndNotifyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily reminder check installed (8 AM)',
    'Trigger Installed',
    5
  );

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.TRIGGER_INSTALLED, {
      triggerName: 'checkAndNotifyReminders',
      schedule: 'Daily at 8 AM',
      installedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true };
}

/**
 * Gets reminder summary for dashboard display.
 * @returns {Object} Summary with counts and upcoming reminders
 */
function getReminderSummary() {
  const todayReminders = getDueReminders(0); // Due today
  const weekReminders = getDueReminders(7);  // Due within a week

  return {
    dueToday: todayReminders.length,
    dueThisWeek: weekReminders.length,
    upcoming: weekReminders.slice(0, 5) // Top 5 for dashboard
  };
}



/**
 * ============================================================================
 * LOOKER STUDIO INTEGRATION
 * ============================================================================
 * Provides a read-only data layer for Google Looker Studio (Data Studio).
 * Creates optimized hidden sheets that Looker can connect to as data sources.
 *
 * RESTRICTED DATA SOURCES:
 * This integration ONLY exports data from:
 * - Member Directory → _Looker_Members
 * - Grievance Log → _Looker_Grievances
 * - Member Satisfaction → _Looker_Satisfaction
 *
 * ARCHITECTURE:
 * - Does NOT modify existing sheets or modal dashboards
 * - Creates separate hidden "_Looker_*" sheets for data export
 * - Looker connects to these sheets as native Google Sheets data sources
 * - Data is refreshed on-demand or via scheduled trigger
 *
 * USAGE:
 * 1. Run setupLookerIntegration() once to create sheets
 * 2. Run refreshLookerData() to update data (or install trigger)
 * 3. Connect Looker Studio to the spreadsheet, select _Looker_* sheets
 */

const LOOKER_CONFIG = {
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
// SETUP & MANAGEMENT
// ============================================================================

/**
 * Sets up Looker Studio integration by creating hidden data sheets.
 * Only creates sheets for allowed data sources.
 * @returns {Object} Result with created sheet names
 */
function setupLookerIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Looker Studio Integration',
    'This will create hidden data sheets for Looker Studio:\n\n' +
    '• _Looker_Members - From Member Directory\n' +
    '• _Looker_Grievances - From Grievance Log\n' +
    '• _Looker_Satisfaction - From Member Satisfaction Survey\n\n' +
    'Only these 3 source sheets will be accessible to Looker.\n' +
    'Existing sheets will NOT be modified.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  const created = [];

  // Create each Looker sheet
  for (const key in LOOKER_CONFIG.SHEETS) {
    const sheetName = LOOKER_CONFIG.SHEETS[key];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.hideSheet();
      created.push(sheetName);
    }
  }

  // Initialize headers for each sheet
  initializeLookerGrievancesSheet_();
  initializeLookerMembersSheet_();
  initializeLookerSatisfactionSheet_();

  // Do initial data population
  refreshLookerData();

  const summary = created.length > 0
    ? 'Created sheets: ' + created.join(', ')
    : 'All Looker sheets already exist';

  ui.alert(
    'Setup Complete',
    summary + '\n\nData has been populated from:\n' +
    '• Member Directory\n• Grievance Log\n• Member Satisfaction\n\n' +
    'Connect Looker Studio to this spreadsheet and select the _Looker_* sheets as data sources.',
    ui.ButtonSet.OK
  );

  // Log setup
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_INITIALIZED, {
      action: 'LOOKER_SETUP',
      sheetsCreated: created,
      allowedSources: LOOKER_CONFIG.ALLOWED_SOURCES,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, created: created };
}

/**
 * Initializes the Looker Grievances sheet with headers.
 * @private
 */
function initializeLookerGrievancesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.GRIEVANCES);
  if (!sheet) return;

  const headers = [
    'Grievance ID', 'Member ID', 'Member Name', 'Status', 'Current Step',
    'Incident Date', 'Date Filed', 'Date Closed', 'Days Open',
    'Days to Deadline', 'Is Overdue', 'Issue Category', 'Articles',
    'Unit', 'Location', 'Assigned Steward', 'Resolution',
    'Action Type', 'Checklist Progress',
    'Has Reminder 1', 'Reminder 1 Date', 'Has Reminder 2', 'Reminder 2 Date',
    'Filed Month', 'Filed Quarter', 'Filed Year',
    'Closed Month', 'Closed Quarter', 'Closed Year',
    'Outcome Category', 'Time to Resolution Days',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the Looker Members sheet with headers.
 * @private
 */
function initializeLookerMembersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);
  if (!sheet) return;

  const headers = [
    'Member ID', 'Full Name', 'First Name', 'Last Name',
    'Job Title', 'Unit', 'Location',
    'Is Steward', 'Is Member Leader', 'Assigned Steward',
    'Email', 'Phone', 'Preferred Communication',
    'Has Open Grievance', 'Total Grievances', 'Grievances Won', 'Grievances Lost',
    'Last Contact Date', 'Days Since Contact',
    'Volunteer Hours', 'Last Virtual Meeting', 'Last In-Person Meeting',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the Looker Satisfaction sheet with headers.
 * @private
 */
function initializeLookerSatisfactionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SATISFACTION);
  if (!sheet) return;

  const headers = [
    'Response ID', 'Timestamp', 'Response Month', 'Response Quarter', 'Response Year',
    'Worksite', 'Role', 'Shift', 'Time in Role', 'Has Steward Contact',
    'Overall Satisfaction Avg', 'Steward Rating Avg', 'Steward Access Avg',
    'Chapter Effectiveness Avg', 'Leadership Avg', 'Contract Enforcement Avg',
    'Communication Avg', 'Member Voice Avg', 'Value Action Avg', 'Scheduling Avg',
    'Satisfied with Rep', 'Trust Union', 'Feel Protected', 'Would Recommend',
    'Filed Grievance', 'Representation Avg',
    'Verification Status', 'Matched Member ID', 'Quarter Period',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

// ============================================================================
// DATA REFRESH FUNCTIONS
// ============================================================================

/**
 * Refreshes all Looker data sheets.
 * Only pulls from allowed source sheets.
 * @returns {Object} Result with refresh counts
 */
function refreshLookerData() {
  const startTime = new Date();

  const grievanceCount = refreshLookerGrievances_();
  const memberCount = refreshLookerMembers_();
  const satisfactionCount = refreshLookerSatisfaction_();

  const duration = (new Date() - startTime) / 1000;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Members: ${memberCount}, Grievances: ${grievanceCount}, Survey: ${satisfactionCount}`,
    'Looker Data Refreshed',
    5
  );

  return {
    success: true,
    sources: LOOKER_CONFIG.ALLOWED_SOURCES,
    grievances: grievanceCount,
    members: memberCount,
    satisfaction: satisfactionCount,
    durationSeconds: duration
  };
}

/**
 * Refreshes the Looker Grievances sheet from Grievance Log.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerGrievances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.GRIEVANCES);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = GRIEVANCE_COLS;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const grievanceId = row[cols.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    const firstName = row[cols.FIRST_NAME - 1] || '';
    const lastName = row[cols.LAST_NAME - 1] || '';
    const status = row[cols.STATUS - 1] || '';
    const dateFiled = row[cols.DATE_FILED - 1];
    const dateClosed = row[cols.DATE_CLOSED - 1];
    const daysToDeadline = row[cols.DAYS_TO_DEADLINE - 1];
    const r1Date = row[cols.REMINDER_1_DATE - 1];
    const r2Date = row[cols.REMINDER_2_DATE - 1];

    // Compute derived fields
    const isOverdue = typeof daysToDeadline === 'number' && daysToDeadline < 0;
    const outcomeCategory = getOutcomeCategory_(status);
    const timeToResolution = (dateFiled instanceof Date && dateClosed instanceof Date)
      ? Math.ceil((dateClosed - dateFiled) / (1000 * 60 * 60 * 24))
      : '';

    // Date dimensions
    const filedMonth = dateFiled instanceof Date ? Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const filedQuarter = dateFiled instanceof Date ? getQuarter_(dateFiled) : '';
    const filedYear = dateFiled instanceof Date ? dateFiled.getFullYear() : '';
    const closedMonth = dateClosed instanceof Date ? Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const closedQuarter = dateClosed instanceof Date ? getQuarter_(dateClosed) : '';
    const closedYear = dateClosed instanceof Date ? dateClosed.getFullYear() : '';

    exportData.push([
      grievanceId,
      row[cols.MEMBER_ID - 1] || '',
      (firstName + ' ' + lastName).trim(),
      status,
      row[cols.CURRENT_STEP - 1] || '',
      row[cols.INCIDENT_DATE - 1] || '',
      dateFiled || '',
      dateClosed || '',
      row[cols.DAYS_OPEN - 1] || '',
      daysToDeadline || '',
      isOverdue ? 'Yes' : 'No',
      row[cols.ISSUE_CATEGORY - 1] || '',
      row[cols.ARTICLES - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.LOCATION - 1] || '',
      row[cols.STEWARD - 1] || '',
      row[cols.RESOLUTION - 1] || '',
      row[cols.ACTION_TYPE - 1] || 'Grievance',
      row[cols.CHECKLIST_PROGRESS - 1] || '',
      r1Date ? 'Yes' : 'No',
      r1Date || '',
      r2Date ? 'Yes' : 'No',
      r2Date || '',
      filedMonth,
      filedQuarter,
      filedYear,
      closedMonth,
      closedQuarter,
      closedYear,
      outcomeCategory,
      timeToResolution,
      now
    ]);
  }

  // Clear old data and write new
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes the Looker Members sheet from Member Directory.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerMembers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.MEMBER_DIR || 'Member Directory');
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);

  if (!sourceSheet || !targetSheet) return 0;

  const memberData = sourceSheet.getDataRange().getValues();
  if (memberData.length < 2) return 0;

  // Build grievance stats by member
  const grievanceStats = {};
  if (grievanceSheet) {
    const gData = grievanceSheet.getDataRange().getValues();
    for (let i = 1; i < gData.length; i++) {
      const memberId = gData[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      const status = gData[i][GRIEVANCE_COLS.STATUS - 1];
      if (!memberId) continue;

      if (!grievanceStats[memberId]) {
        grievanceStats[memberId] = { total: 0, won: 0, lost: 0, open: 0 };
      }
      grievanceStats[memberId].total++;
      if (status === 'Won') grievanceStats[memberId].won++;
      if (status === 'Denied') grievanceStats[memberId].lost++;
      if (status === 'Open' || status === 'Pending Info' || status === 'Appealed' || status === 'In Arbitration') {
        grievanceStats[memberId].open++;
      }
    }
  }

  const now = new Date();
  const exportData = [];
  const cols = MEMBER_COLS;

  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const memberId = row[cols.MEMBER_ID - 1];
    if (!memberId) continue;

    const firstName = row[cols.FIRST_NAME - 1] || '';
    const lastName = row[cols.LAST_NAME - 1] || '';
    const isSteward = row[cols.IS_STEWARD - 1];
    const lastContact = row[cols.RECENT_CONTACT_DATE - 1];
    const stats = grievanceStats[memberId] || { total: 0, won: 0, lost: 0, open: 0 };

    // Days since last contact
    const daysSinceContact = lastContact instanceof Date
      ? Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24))
      : '';

    exportData.push([
      memberId,
      (firstName + ' ' + lastName).trim(),
      firstName,
      lastName,
      row[cols.JOB_TITLE - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.WORK_LOCATION - 1] || '',
      isSteward === 'Yes' ? 'Yes' : 'No',
      isSteward === 'Member Leader' ? 'Yes' : 'No',
      row[cols.ASSIGNED_STEWARD - 1] || '',
      row[cols.EMAIL - 1] || '',
      row[cols.PHONE - 1] || '',
      row[cols.PREFERRED_COMM - 1] || '',
      stats.open > 0 ? 'Yes' : 'No',
      stats.total,
      stats.won,
      stats.lost,
      lastContact || '',
      daysSinceContact,
      parseFloat(row[cols.VOLUNTEER_HOURS - 1]) || 0,
      row[cols.LAST_VIRTUAL_MTG - 1] || '',
      row[cols.LAST_INPERSON_MTG - 1] || '',
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes the Looker Satisfaction sheet from Member Satisfaction survey.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerSatisfaction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.SATISFACTION || '📊 Member Satisfaction');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.SATISFACTION);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = typeof SATISFACTION_COLS !== 'undefined' ? SATISFACTION_COLS : {};

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const timestamp = row[0]; // First column is always timestamp
    if (!timestamp) continue;

    // Generate response ID
    const responseId = 'SR' + String(i).padStart(5, '0');

    // Date dimensions
    const respMonth = timestamp instanceof Date ? Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const respQuarter = timestamp instanceof Date ? getQuarter_(timestamp) : '';
    const respYear = timestamp instanceof Date ? timestamp.getFullYear() : '';

    // Calculate section averages using SATISFACTION_COLS if available
    const overallSatAvg = calculateSectionAvg_(row, [6, 7, 8, 9]); // Q6-Q9
    const stewardRatingAvg = calculateSectionAvg_(row, [10, 11, 12, 13, 14, 15, 16]); // Q10-Q16
    const stewardAccessAvg = calculateSectionAvg_(row, [18, 19, 20]); // Q18-Q20
    const chapterAvg = calculateSectionAvg_(row, [21, 22, 23, 24, 25]); // Q21-Q25
    const leadershipAvg = calculateSectionAvg_(row, [26, 27, 28, 29, 30, 31]); // Q26-Q31
    const contractAvg = calculateSectionAvg_(row, [32, 33, 34, 35]); // Q32-Q35
    const commAvg = calculateSectionAvg_(row, [41, 42, 43, 44, 45]); // Q41-Q45
    const voiceAvg = calculateSectionAvg_(row, [46, 47, 48, 49, 50]); // Q46-Q50
    const valueAvg = calculateSectionAvg_(row, [51, 52, 53, 54, 55]); // Q51-Q55
    const schedAvg = calculateSectionAvg_(row, [56, 57, 58, 59, 60, 61, 62]); // Q56-Q62
    const repAvg = calculateSectionAvg_(row, [37, 38, 39, 40]); // Q37-Q40

    // Get individual key questions (0-indexed)
    const worksite = row[1] || '';
    const role = row[2] || '';
    const shift = row[3] || '';
    const timeInRole = row[4] || '';
    const hasStewardContact = row[5] || '';
    const satisfiedRep = row[6] || '';
    const trustUnion = row[7] || '';
    const feelProtected = row[8] || '';
    const wouldRecommend = row[9] || '';
    const filedGrievance = row[36] || '';

    // Verification columns (if they exist)
    const verificationStatus = cols.VERIFIED ? (row[cols.VERIFIED - 1] || '') : '';
    const matchedMemberId = cols.MATCHED_MEMBER_ID ? (row[cols.MATCHED_MEMBER_ID - 1] || '') : '';
    const quarterPeriod = cols.QUARTER ? (row[cols.QUARTER - 1] || '') : respQuarter;

    exportData.push([
      responseId,
      timestamp,
      respMonth,
      respQuarter,
      respYear,
      worksite,
      role,
      shift,
      timeInRole,
      hasStewardContact,
      overallSatAvg,
      stewardRatingAvg,
      stewardAccessAvg,
      chapterAvg,
      leadershipAvg,
      contractAvg,
      commAvg,
      voiceAvg,
      valueAvg,
      schedAvg,
      satisfiedRep,
      trustUnion,
      feelProtected,
      wouldRecommend,
      filedGrievance,
      repAvg,
      verificationStatus,
      matchedMemberId,
      quarterPeriod,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

// ============================================================================
// PII-FREE LOOKER INTEGRATION (Anonymized Data)
// ============================================================================
// For external stakeholders, compliance, or public-facing dashboards.
// Excludes: Names, Emails, Phones, Addresses, Member IDs

/**
 * Sets up PII-free Looker sheets for external/compliance use.
 * @returns {Object} Result with created sheet names
 */
function setupLookerAnonIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup PII-Free Looker Integration',
    'This will create ANONYMIZED data sheets for external Looker reports:\n\n' +
    '• _Looker_Anon_Members - Aggregated member data (no names/contact)\n' +
    '• _Looker_Anon_Grievances - Case data (no member info)\n' +
    '• _Looker_Anon_Satisfaction - Survey data (already anonymous)\n\n' +
    'These sheets contain NO personally identifiable information (PII).\n' +
    'Safe for external dashboards and compliance reporting.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, cancelled: true };
  }

  const created = [];

  // Create each anonymous Looker sheet
  for (const key in LOOKER_CONFIG.SHEETS_ANON) {
    const sheetName = LOOKER_CONFIG.SHEETS_ANON[key];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.hideSheet();
      created.push(sheetName);
    }
  }

  // Initialize headers for each sheet
  initializeLookerAnonGrievancesSheet_();
  initializeLookerAnonMembersSheet_();
  initializeLookerAnonSatisfactionSheet_();

  // Do initial data population
  refreshLookerAnonData();

  const summary = created.length > 0
    ? 'Created sheets: ' + created.join(', ')
    : 'All PII-free Looker sheets already exist';

  ui.alert(
    'PII-Free Setup Complete',
    summary + '\n\nThese sheets contain NO personally identifiable information:\n' +
    '• No names, emails, or phone numbers\n' +
    '• No member IDs (uses anonymous hashes)\n' +
    '• Safe for external/public dashboards',
    ui.ButtonSet.OK
  );

  // Log setup
  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_INITIALIZED, {
      action: 'LOOKER_ANON_SETUP',
      sheetsCreated: created,
      piiExcluded: true,
      performedBy: Session.getActiveUser().getEmail()
    });
  }

  return { success: true, created: created, piiExcluded: true };
}

/**
 * Initializes the anonymized Grievances sheet headers.
 * @private
 */
function initializeLookerAnonGrievancesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.GRIEVANCES);
  if (!sheet) return;

  // No Member Name, No Steward Name - only aggregation-safe fields
  const headers = [
    'Case Hash', 'Status', 'Current Step',
    'Incident Month', 'Incident Quarter', 'Incident Year',
    'Filed Month', 'Filed Quarter', 'Filed Year',
    'Closed Month', 'Closed Quarter', 'Closed Year',
    'Days Open', 'Days to Deadline', 'Is Overdue',
    'Issue Category', 'Articles Violated',
    'Unit', 'Location', 'Action Type',
    'Outcome Category', 'Time to Resolution Days',
    'Has Reminder', 'Checklist Progress',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the anonymized Members sheet headers.
 * @private
 */
function initializeLookerAnonMembersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.MEMBERS);
  if (!sheet) return;

  // Aggregation-focused: No names, no contact info, no member IDs
  const headers = [
    'Member Hash', 'Job Title', 'Unit', 'Location',
    'Is Steward', 'Is Member Leader',
    'Has Open Grievance', 'Total Grievances', 'Grievances Won', 'Grievances Lost',
    'Days Since Contact Bucket', 'Contact Frequency Category',
    'Volunteer Hours Bucket', 'Engagement Level',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Initializes the anonymized Satisfaction sheet headers.
 * @private
 */
function initializeLookerAnonSatisfactionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.SATISFACTION);
  if (!sheet) return;

  // Survey data is already mostly anonymous, just remove any potential PII linkages
  const headers = [
    'Response Hash', 'Response Month', 'Response Quarter', 'Response Year',
    'Worksite', 'Role Category', 'Shift', 'Tenure Bucket',
    'Has Steward Contact',
    'Overall Satisfaction Avg', 'Steward Rating Avg', 'Steward Access Avg',
    'Chapter Effectiveness Avg', 'Leadership Avg', 'Contract Enforcement Avg',
    'Communication Avg', 'Member Voice Avg', 'Value Action Avg', 'Scheduling Avg',
    'Representation Avg',
    'Satisfied Bucket', 'Trust Bucket', 'Protected Bucket', 'Recommend Bucket',
    'Filed Grievance',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#334155')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

/**
 * Refreshes all PII-free Looker data sheets.
 * @returns {Object} Result with refresh counts
 */
function refreshLookerAnonData() {
  const startTime = new Date();

  const grievanceCount = refreshLookerAnonGrievances_();
  const memberCount = refreshLookerAnonMembers_();
  const satisfactionCount = refreshLookerAnonSatisfaction_();

  const duration = (new Date() - startTime) / 1000;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Anon Members: ${memberCount}, Grievances: ${grievanceCount}, Survey: ${satisfactionCount}`,
    'PII-Free Looker Data Refreshed',
    5
  );

  return {
    success: true,
    piiExcluded: true,
    grievances: grievanceCount,
    members: memberCount,
    satisfaction: satisfactionCount,
    durationSeconds: duration
  };
}

/**
 * Refreshes anonymized Grievances sheet - NO member names or IDs.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonGrievances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.GRIEVANCES);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];
  const cols = GRIEVANCE_COLS;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const grievanceId = row[cols.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    const status = row[cols.STATUS - 1] || '';
    const incidentDate = row[cols.INCIDENT_DATE - 1];
    const dateFiled = row[cols.DATE_FILED - 1];
    const dateClosed = row[cols.DATE_CLOSED - 1];
    const daysToDeadline = row[cols.DAYS_TO_DEADLINE - 1];
    const r1Date = row[cols.REMINDER_1_DATE - 1];
    const r2Date = row[cols.REMINDER_2_DATE - 1];

    // Generate anonymous hash (not reversible)
    const caseHash = generateAnonHash_(grievanceId);

    // Compute derived fields
    const isOverdue = typeof daysToDeadline === 'number' && daysToDeadline < 0;
    const outcomeCategory = getOutcomeCategory_(status);
    const timeToResolution = (dateFiled instanceof Date && dateClosed instanceof Date)
      ? Math.ceil((dateClosed - dateFiled) / (1000 * 60 * 60 * 24))
      : '';

    // Date dimensions (incident)
    const incMonth = incidentDate instanceof Date ? Utilities.formatDate(incidentDate, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const incQuarter = incidentDate instanceof Date ? getQuarter_(incidentDate) : '';
    const incYear = incidentDate instanceof Date ? incidentDate.getFullYear() : '';

    // Date dimensions (filed)
    const filedMonth = dateFiled instanceof Date ? Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const filedQuarter = dateFiled instanceof Date ? getQuarter_(dateFiled) : '';
    const filedYear = dateFiled instanceof Date ? dateFiled.getFullYear() : '';

    // Date dimensions (closed)
    const closedMonth = dateClosed instanceof Date ? Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const closedQuarter = dateClosed instanceof Date ? getQuarter_(dateClosed) : '';
    const closedYear = dateClosed instanceof Date ? dateClosed.getFullYear() : '';

    exportData.push([
      caseHash,
      status,
      row[cols.CURRENT_STEP - 1] || '',
      incMonth,
      incQuarter,
      incYear,
      filedMonth,
      filedQuarter,
      filedYear,
      closedMonth,
      closedQuarter,
      closedYear,
      row[cols.DAYS_OPEN - 1] || '',
      daysToDeadline || '',
      isOverdue ? 'Yes' : 'No',
      row[cols.ISSUE_CATEGORY - 1] || '',
      row[cols.ARTICLES - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.LOCATION - 1] || '',
      row[cols.ACTION_TYPE - 1] || 'Grievance',
      outcomeCategory,
      timeToResolution,
      (r1Date || r2Date) ? 'Yes' : 'No',
      row[cols.CHECKLIST_PROGRESS - 1] || '',
      now
    ]);
  }

  // Clear old data and write new
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes anonymized Members sheet - NO names, emails, phones, or IDs.
 * Uses bucketed/categorized values for privacy.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonMembers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.MEMBER_DIR || 'Member Directory');
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG || 'Grievance Log');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.MEMBERS);

  if (!sourceSheet || !targetSheet) return 0;

  const memberData = sourceSheet.getDataRange().getValues();
  if (memberData.length < 2) return 0;

  // Build grievance stats by member
  const grievanceStats = {};
  if (grievanceSheet) {
    const gData = grievanceSheet.getDataRange().getValues();
    for (let i = 1; i < gData.length; i++) {
      const memberId = gData[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      const status = gData[i][GRIEVANCE_COLS.STATUS - 1];
      if (!memberId) continue;

      if (!grievanceStats[memberId]) {
        grievanceStats[memberId] = { total: 0, won: 0, lost: 0, open: 0 };
      }
      grievanceStats[memberId].total++;
      if (status === 'Won') grievanceStats[memberId].won++;
      if (status === 'Denied') grievanceStats[memberId].lost++;
      if (status === 'Open' || status === 'Pending Info' || status === 'Appealed' || status === 'In Arbitration') {
        grievanceStats[memberId].open++;
      }
    }
  }

  const now = new Date();
  const exportData = [];
  const cols = MEMBER_COLS;

  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const memberId = row[cols.MEMBER_ID - 1];
    if (!memberId) continue;

    const isSteward = row[cols.IS_STEWARD - 1];
    const lastContact = row[cols.RECENT_CONTACT_DATE - 1];
    const stats = grievanceStats[memberId] || { total: 0, won: 0, lost: 0, open: 0 };
    const volunteerHours = parseFloat(row[cols.VOLUNTEER_HOURS - 1]) || 0;

    // Generate anonymous hash
    const memberHash = generateAnonHash_(memberId);

    // Days since last contact - BUCKETED for privacy
    let daysSinceContactBucket = 'Unknown';
    if (lastContact instanceof Date) {
      const days = Math.ceil((now - lastContact) / (1000 * 60 * 60 * 24));
      daysSinceContactBucket = getDaysBucket_(days);
    }

    // Contact frequency category
    const contactFreq = getContactFrequencyCategory_(lastContact, now);

    // Volunteer hours bucket
    const volunteerBucket = getVolunteerHoursBucket_(volunteerHours);

    // Engagement level
    const engagementLevel = getEngagementLevel_(volunteerHours, lastContact, isSteward, stats.total);

    exportData.push([
      memberHash,
      row[cols.JOB_TITLE - 1] || '',
      row[cols.UNIT - 1] || '',
      row[cols.WORK_LOCATION - 1] || '',
      isSteward === 'Yes' ? 'Yes' : 'No',
      isSteward === 'Member Leader' ? 'Yes' : 'No',
      stats.open > 0 ? 'Yes' : 'No',
      stats.total,
      stats.won,
      stats.lost,
      daysSinceContactBucket,
      contactFreq,
      volunteerBucket,
      engagementLevel,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Refreshes anonymized Satisfaction sheet - surveys are already mostly anonymous.
 * @private
 * @returns {number} Number of rows exported
 */
function refreshLookerAnonSatisfaction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEETS.SATISFACTION || '📊 Member Satisfaction');
  const targetSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS_ANON.SATISFACTION);

  if (!sourceSheet || !targetSheet) return 0;

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) return 0;

  const now = new Date();
  const exportData = [];

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const timestamp = row[0];
    if (!timestamp) continue;

    // Generate anonymous response hash
    const responseHash = generateAnonHash_('SR' + i + (timestamp instanceof Date ? timestamp.getTime() : ''));

    // Date dimensions
    const respMonth = timestamp instanceof Date ? Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM') : '';
    const respQuarter = timestamp instanceof Date ? getQuarter_(timestamp) : '';
    const respYear = timestamp instanceof Date ? timestamp.getFullYear() : '';

    // Calculate section averages
    const overallSatAvg = calculateSectionAvg_(row, [6, 7, 8, 9]);
    const stewardRatingAvg = calculateSectionAvg_(row, [10, 11, 12, 13, 14, 15, 16]);
    const stewardAccessAvg = calculateSectionAvg_(row, [18, 19, 20]);
    const chapterAvg = calculateSectionAvg_(row, [21, 22, 23, 24, 25]);
    const leadershipAvg = calculateSectionAvg_(row, [26, 27, 28, 29, 30, 31]);
    const contractAvg = calculateSectionAvg_(row, [32, 33, 34, 35]);
    const commAvg = calculateSectionAvg_(row, [41, 42, 43, 44, 45]);
    const voiceAvg = calculateSectionAvg_(row, [46, 47, 48, 49, 50]);
    const valueAvg = calculateSectionAvg_(row, [51, 52, 53, 54, 55]);
    const schedAvg = calculateSectionAvg_(row, [56, 57, 58, 59, 60, 61, 62]);
    const repAvg = calculateSectionAvg_(row, [37, 38, 39, 40]);

    // Get individual questions - BUCKETED for aggregation
    const worksite = row[1] || '';
    const role = categorizeRole_(row[2] || '');
    const shift = row[3] || '';
    const timeInRole = categorizeTenure_(row[4] || '');
    const hasStewardContact = row[5] || '';
    const filedGrievance = row[36] || '';

    // Bucket satisfaction scores (1-10 → Low/Medium/High)
    const satisfiedBucket = getScoreBucket_(row[6]);
    const trustBucket = getScoreBucket_(row[7]);
    const protectedBucket = getScoreBucket_(row[8]);
    const recommendBucket = getScoreBucket_(row[9]);

    exportData.push([
      responseHash,
      respMonth,
      respQuarter,
      respYear,
      worksite,
      role,
      shift,
      timeInRole,
      hasStewardContact,
      overallSatAvg,
      stewardRatingAvg,
      stewardAccessAvg,
      chapterAvg,
      leadershipAvg,
      contractAvg,
      commAvg,
      voiceAvg,
      valueAvg,
      schedAvg,
      repAvg,
      satisfiedBucket,
      trustBucket,
      protectedBucket,
      recommendBucket,
      filedGrievance,
      now
    ]);
  }

  // Clear and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  }

  if (exportData.length > 0) {
    targetSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }

  return exportData.length;
}

/**
 * Generates an anonymous, non-reversible hash from an identifier.
 * @private
 */
function generateAnonHash_(id) {
  // Use a simple hash that can't be reversed to original ID
  const salt = 'anon509data';
  const combined = salt + String(id);
  let hash = 0;
  for (let i = 0; i < combined.length; i++) {
    const char = combined.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return 'A' + Math.abs(hash).toString(36).toUpperCase().substring(0, 8);
}

/**
 * Categorizes days into buckets for privacy.
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
 * Categorizes contact frequency.
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
 * Categorizes volunteer hours into buckets.
 * @private
 */
function getVolunteerHoursBucket_(hours) {
  if (hours === 0) return 'None';
  if (hours <= 5) return '1-5 Hours';
  if (hours <= 20) return '6-20 Hours';
  if (hours <= 50) return '21-50 Hours';
  return '50+ Hours';
}

/**
 * Calculates engagement level from multiple factors.
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
  if (isSteward === 'Yes' || isSteward === 'Member Leader') score += 2;

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
 * Categorizes role into broader categories.
 * @private
 */
function categorizeRole_(role) {
  const roleLower = String(role).toLowerCase();
  if (roleLower.includes('steward') || roleLower.includes('leader')) return 'Leadership';
  if (roleLower.includes('nurse') || roleLower.includes('rn') || roleLower.includes('lpn')) return 'Nursing';
  if (roleLower.includes('tech') || roleLower.includes('aide')) return 'Technical/Support';
  if (roleLower.includes('admin') || roleLower.includes('clerk')) return 'Administrative';
  return 'Other';
}

/**
 * Categorizes tenure into buckets.
 * @private
 */
function categorizeTenure_(tenure) {
  const tenureLower = String(tenure).toLowerCase();
  if (tenureLower.includes('less than') || tenureLower.includes('< 1')) return 'New (< 1 year)';
  if (tenureLower.includes('1-3') || tenureLower.includes('1 to 3')) return '1-3 Years';
  if (tenureLower.includes('3-5') || tenureLower.includes('3 to 5')) return '3-5 Years';
  if (tenureLower.includes('5-10') || tenureLower.includes('5 to 10')) return '5-10 Years';
  return '10+ Years';
}

/**
 * Converts numeric score (1-10) to bucket.
 * @private
 */
function getScoreBucket_(score) {
  const num = parseFloat(score);
  if (isNaN(num)) return 'No Response';
  if (num >= 8) return 'High (8-10)';
  if (num >= 5) return 'Medium (5-7)';
  return 'Low (1-4)';
}

/**
 * Shows help dialog for PII-free Looker connection.
 */
function showLookerAnonConnectionHelp() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const html = `<!DOCTYPE html>
<html>
<head>
  ${getMobileOptimizedHead()}
  <style>
    body { font-family: 'Google Sans', Roboto, sans-serif; padding: 20px; color: #1F2937; }
    h2 { color: #059669; margin-bottom: 16px; }
    .pii-badge { display: inline-block; background: #059669; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; margin-bottom: 16px; }
    .step { background: #F3F4F6; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; }
    .step-num { display: inline-block; width: 24px; height: 24px; background: #059669; color: white; border-radius: 50%; text-align: center; line-height: 24px; margin-right: 8px; font-size: 12px; }
    code { background: #E5E7EB; padding: 2px 6px; border-radius: 4px; font-size: 12px; }
    .sheets { margin: 16px 0; }
    .sheet-name { display: inline-block; background: #334155; color: #F8FAFC; padding: 4px 12px; border-radius: 4px; margin: 4px; font-size: 13px; }
    .excluded { margin: 16px 0; padding: 12px; background: #ECFDF5; border: 1px solid #059669; border-radius: 8px; }
    .excluded h4 { color: #059669; margin-bottom: 8px; font-size: 13px; }
    .excluded ul { margin: 0; padding-left: 20px; font-size: 12px; color: #065F46; }
    .btn { display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #059669, #047857); color: white; text-decoration: none; border-radius: 6px; margin-top: 12px; }
  </style>
</head>
<body>
  <h2>PII-Free Looker Connection</h2>
  <span class="pii-badge">✓ NO PERSONAL DATA</span>

  <div class="excluded">
    <h4>Excluded from these sheets:</h4>
    <ul>
      <li>Member names (first, last, full)</li>
      <li>Email addresses</li>
      <li>Phone numbers</li>
      <li>Member IDs (uses anonymous hashes)</li>
      <li>Steward names</li>
    </ul>
  </div>

  <div class="step">
    <span class="step-num">1</span>
    Open <a href="https://lookerstudio.google.com" target="_blank">Looker Studio</a> and create a new report
  </div>

  <div class="step">
    <span class="step-num">2</span>
    Add a data source → Select <strong>Google Sheets</strong>
  </div>

  <div class="step">
    <span class="step-num">3</span>
    Select these ANONYMIZED sheets:
    <div class="sheets">
      <span class="sheet-name">_Looker_Anon_Members</span>
      <span class="sheet-name">_Looker_Anon_Grievances</span>
      <span class="sheet-name">_Looker_Anon_Satisfaction</span>
    </div>
  </div>

  <a class="btn" href="https://lookerstudio.google.com/reporting/create" target="_blank">
    Open Looker Studio →
  </a>

  <p style="margin-top: 20px; font-size: 11px; color: #6B7280;">
    Safe for: External stakeholders, public dashboards, compliance reporting
  </p>
</body>
</html>`;

  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(480)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(dialog, 'PII-Free Looker Connection');
}

/**
 * Gets status of PII-free Looker integration.
 * @returns {Object} Status info
 */
function getLookerAnonStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = {
    isSetup: true,
    piiExcluded: true,
    sheets: {},
    recordCounts: {}
  };

  for (const key in LOOKER_CONFIG.SHEETS_ANON) {
    const sheetName = LOOKER_CONFIG.SHEETS_ANON[key];
    const sheet = ss.getSheetByName(sheetName);

    status.sheets[key] = {
      exists: !!sheet,
      name: sheetName,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
    };

    if (!sheet) status.isSetup = false;
    status.recordCounts[key] = status.sheets[key].rows;
  }

  return status;
}

/**
 * Installs trigger to refresh both standard and PII-free Looker data.
 */
function installLookerAllRefreshTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'refreshLookerData' || fn === 'refreshLookerAnonData' || fn === 'refreshAllLookerData_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install combined refresh trigger
  ScriptApp.newTrigger('refreshAllLookerData_')
    .timeBased()
    .everyDays(1)
    .atHour(LOOKER_CONFIG.AUTO_REFRESH_HOUR)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily refresh installed for both standard and PII-free data',
    'Looker Triggers Installed',
    5
  );

  return { success: true };
}

/**
 * Refreshes both standard and PII-free Looker data.
 * @private
 */
function refreshAllLookerData_() {
  refreshLookerData();
  refreshLookerAnonData();
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Calculates average of numeric values in specified columns.
 * @private
 */
function calculateSectionAvg_(row, colIndices) {
  let sum = 0;
  let count = 0;

  for (let i = 0; i < colIndices.length; i++) {
    const val = row[colIndices[i]];
    const num = parseFloat(val);
    if (!isNaN(num) && num >= 1 && num <= 10) {
      sum += num;
      count++;
    }
  }

  return count > 0 ? Math.round((sum / count) * 10) / 10 : '';
}

/**
 * Gets quarter string from date.
 * @private
 */
function getQuarter_(date) {
  const month = date.getMonth();
  const quarter = Math.floor(month / 3) + 1;
  return date.getFullYear() + '-Q' + quarter;
}

/**
 * Gets outcome category from status.
 * @private
 */
function getOutcomeCategory_(status) {
  if (status === 'Won') return 'Win';
  if (status === 'Denied') return 'Loss';
  if (status === 'Settled') return 'Settlement';
  if (status === 'Withdrawn') return 'Withdrawn';
  if (status === 'Closed') return 'Closed';
  return 'Active';
}

// ============================================================================
// TRIGGERS & AUTOMATION
// ============================================================================

/**
 * Installs daily auto-refresh trigger for Looker data.
 */
function installLookerRefreshTrigger() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshLookerData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Install new trigger
  ScriptApp.newTrigger('refreshLookerData')
    .timeBased()
    .everyDays(1)
    .atHour(LOOKER_CONFIG.AUTO_REFRESH_HOUR)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily refresh installed (' + LOOKER_CONFIG.AUTO_REFRESH_HOUR + ':00 AM)',
    'Looker Trigger Installed',
    5
  );

  return { success: true };
}

/**
 * Gets the Looker Studio connection URL for this spreadsheet.
 * @returns {string} URL to create new Looker Studio report
 */
function getLookerConnectionUrl() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return 'https://lookerstudio.google.com/reporting/create?c.reportId=&r.reportName=509%20Dashboard%20Report&ds.connector=googleSheets&ds.spreadsheetId=' + ssId;
}

/**
 * Shows dialog with Looker Studio connection instructions.
 * Lists only the 3 allowed data sources.
 */
function showLookerConnectionHelp() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const html = `<!DOCTYPE html>
<html>
<head>
  ${getMobileOptimizedHead()}
  <style>
    body { font-family: 'Google Sans', Roboto, sans-serif; padding: 20px; color: #1F2937; }
    h2 { color: #7C3AED; margin-bottom: 16px; }
    .step { background: #F3F4F6; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; }
    .step-num { display: inline-block; width: 24px; height: 24px; background: #7C3AED; color: white; border-radius: 50%; text-align: center; line-height: 24px; margin-right: 8px; font-size: 12px; }
    code { background: #E5E7EB; padding: 2px 6px; border-radius: 4px; font-size: 12px; }
    .sheets { margin: 16px 0; }
    .sheet-name { display: inline-block; background: #1E293B; color: #F8FAFC; padding: 4px 12px; border-radius: 4px; margin: 4px; font-size: 13px; }
    .source-info { font-size: 11px; color: #6B7280; display: block; margin-top: 2px; }
    .btn { display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; text-decoration: none; border-radius: 6px; margin-top: 12px; }
    .btn:hover { opacity: 0.9; }
    .restricted { background: #FEF3C7; border: 1px solid #F59E0B; padding: 10px; border-radius: 6px; margin: 16px 0; font-size: 13px; }
  </style>
</head>
<body>
  <h2>Connect to Looker Studio</h2>

  <div class="restricted">
    <strong>Restricted Data Access:</strong> Only Member Directory, Grievance Log, and Member Satisfaction data is available to Looker Studio.
  </div>

  <div class="step">
    <span class="step-num">1</span>
    Open <a href="https://lookerstudio.google.com" target="_blank">Looker Studio</a> and create a new report
  </div>

  <div class="step">
    <span class="step-num">2</span>
    Add a data source → Select <strong>Google Sheets</strong>
  </div>

  <div class="step">
    <span class="step-num">3</span>
    Find this spreadsheet or paste the URL:<br>
    <code style="word-break: break-all; display: block; margin-top: 8px;">${ssUrl}</code>
  </div>

  <div class="step">
    <span class="step-num">4</span>
    Select one of these sheets as your data source:
    <div class="sheets">
      <span class="sheet-name">_Looker_Members<span class="source-info">← Member Directory</span></span>
      <span class="sheet-name">_Looker_Grievances<span class="source-info">← Grievance Log</span></span>
      <span class="sheet-name">_Looker_Satisfaction<span class="source-info">← Survey Data</span></span>
    </div>
  </div>

  <div class="step">
    <span class="step-num">5</span>
    Add multiple data sources to create relationships between members and grievances
  </div>

  <a class="btn" href="https://lookerstudio.google.com/reporting/create" target="_blank">
    Open Looker Studio →
  </a>

  <p style="margin-top: 20px; font-size: 12px; color: #6B7280;">
    Spreadsheet ID: <code>${ssId}</code>
  </p>
</body>
</html>`;

  const dialog = HtmlService.createHtmlOutput(html)
    .setWidth(520)
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(dialog, 'Looker Studio Connection');
}

/**
 * Gets Looker integration status.
 * @returns {Object} Status of Looker sheets and data
 */
function getLookerStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = {
    isSetup: true,
    allowedSources: LOOKER_CONFIG.ALLOWED_SOURCES,
    sheets: {},
    lastRefresh: null,
    recordCounts: {}
  };

  for (const key in LOOKER_CONFIG.SHEETS) {
    const sheetName = LOOKER_CONFIG.SHEETS[key];
    const sheet = ss.getSheetByName(sheetName);

    status.sheets[key] = {
      exists: !!sheet,
      name: sheetName,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
    };

    if (!sheet) status.isSetup = false;
    status.recordCounts[key] = status.sheets[key].rows;
  }

  // Get last refresh time from first data sheet
  const memberSheet = ss.getSheetByName(LOOKER_CONFIG.SHEETS.MEMBERS);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    const lastCol = memberSheet.getLastColumn();
    const lastUpdate = memberSheet.getRange(2, lastCol).getValue();
    if (lastUpdate instanceof Date) {
      status.lastRefresh = lastUpdate;
    }
  }

  return status;
}
