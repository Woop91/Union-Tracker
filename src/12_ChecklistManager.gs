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

  // Protect header row
  var protection = sheet.protect().setDescription('Checklist Sheet Structure');
  protection.setUnprotectedRanges([sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length)]);

  return sheet;
}

// ============================================================================
// CHECKLIST ITEM MANAGEMENT
// ============================================================================

/**
 * Generate a unique checklist item ID
 * @returns {string} Unique checklist ID (e.g., CL-00001)
 */
function generateChecklistId_() {
  var sheet = getOrCreateChecklistSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'CL-00001';
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

  return 'CL-' + String(maxNum + 1).padStart(5, '0');
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
    var checklistId = generateChecklistId_();

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
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; background: #F9FAFB; color: #1F2937; }' +
    '.container { padding: 20px; }' +
    '.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }' +
    '.progress-section { background: white; border-radius: 12px; padding: 16px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    '.progress-label { display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px; }' +
    '.progress-bar { height: 8px; background: #E5E7EB; border-radius: 4px; overflow: hidden; }' +
    '.progress-fill { height: 100%; background: ' + progressColor + '; border-radius: 4px; transition: width 0.3s; }' +
    '.checklist-container { background: white; border-radius: 12px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); max-height: 350px; overflow-y: auto; }' +
    '.category-section { margin-bottom: 16px; }' +
    '.category-header { font-size: 12px; font-weight: 600; color: #6B7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 1px solid #E5E7EB; }' +
    '.checklist-item { display: flex; align-items: center; padding: 10px; border-radius: 8px; margin-bottom: 6px; background: #F9FAFB; transition: all 0.2s; }' +
    '.checklist-item:hover { background: #F3F4F6; }' +
    '.checklist-item.completed { opacity: 0.7; }' +
    '.checklist-item.completed .item-text { text-decoration: line-through; color: #9CA3AF; }' +
    '.checkbox-container { position: relative; width: 24px; height: 24px; margin-right: 12px; flex-shrink: 0; }' +
    '.checkbox-container input { opacity: 0; width: 24px; height: 24px; cursor: pointer; }' +
    '.checkmark { position: absolute; top: 0; left: 0; width: 24px; height: 24px; background: white; border: 2px solid #D1D5DB; border-radius: 6px; pointer-events: none; }' +
    '.checkbox-container input:checked ~ .checkmark { background: #7C3AED; border-color: #7C3AED; }' +
    '.checkbox-container input:checked ~ .checkmark:after { content: ""; position: absolute; left: 7px; top: 3px; width: 6px; height: 12px; border: solid white; border-width: 0 2px 2px 0; transform: rotate(45deg); }' +
    '.item-content { flex: 1; }' +
    '.item-text { font-size: 14px; }' +
    '.item-meta { font-size: 11px; color: #9CA3AF; margin-top: 2px; }' +
    '.badge { font-size: 10px; padding: 2px 6px; border-radius: 4px; margin-left: 6px; }' +
    '.badge.required { background: #FEE2E2; color: #991B1B; }' +
    '.btn-icon { background: none; border: none; color: #9CA3AF; cursor: pointer; font-size: 16px; padding: 4px 8px; border-radius: 4px; }' +
    '.btn-icon:hover { background: #FEE2E2; color: #EF4444; }' +
    '.no-items { text-align: center; padding: 40px; color: #6B7280; }' +
    '.actions { display: flex; gap: 10px; margin-top: 20px; justify-content: space-between; }' +
    '.btn { padding: 10px 16px; border-radius: 8px; font-size: 14px; font-weight: 500; cursor: pointer; border: none; transition: all 0.2s; }' +
    '.btn-primary { background: #7C3AED; color: white; }' +
    '.btn-primary:hover { background: #6D28D9; }' +
    '.btn-secondary { background: #E5E7EB; color: #374151; }' +
    '.btn-secondary:hover { background: #D1D5DB; }' +
    '.btn-success { background: #10B981; color: white; }' +
    '.btn-success:hover { background: #059669; }' +
    '.add-item-section { margin-top: 16px; padding-top: 16px; border-top: 1px solid #E5E7EB; }' +
    '.add-form { display: flex; gap: 8px; }' +
    '.add-form input { flex: 1; padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; }' +
    '.add-form select { padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; }' +
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
