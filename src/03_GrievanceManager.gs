/**
 * ============================================================================
 * GrievanceManager.gs - Grievance Lifecycle Management
 * ============================================================================
 *
 * This module handles all grievance-related operations including:
 * - Creating new grievances
 * - Advancing grievance steps
 * - Deadline calculations based on Article 23A
 * - Status updates and lifecycle management
 * - Batch recalculation of deadlines
 *
 * SEPARATION OF CONCERNS: This file contains ONLY grievance business logic.
 * UI components are in UIService.gs, integrations in Integrations.gs.
 *
 * @fileoverview Grievance lifecycle management
 * @version 2.0.0
 * @requires Constants.gs
 */

// ============================================================================
// GRIEVANCE CREATION
// ============================================================================

/**
 * Starts a new grievance for a member
 * @param {Object} grievanceData - Initial grievance data
 * @return {Object} Result with grievance ID or error
 */
function startNewGrievance(grievanceData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);

    if (!grievanceSheet) {
      throw new Error('Grievance Tracker sheet not found');
    }

    // Validate required fields
    const validation = validateGrievanceData(grievanceData);
    if (!validation.valid) {
      return { success: false, error: validation.error };
    }

    // Generate new grievance ID
    const grievanceId = getNextGrievanceId(grievanceSheet);

    // Calculate initial deadlines
    const filingDate = grievanceData.filingDate ? new Date(grievanceData.filingDate) : new Date();
    const deadlines = calculateInitialDeadlines(filingDate);

    // Prepare row data
    const rowData = [
      grievanceId,                                    // Grievance ID
      grievanceData.memberId || '',                   // Member ID
      grievanceData.memberName || '',                 // Member Name
      filingDate,                                     // Filing Date
      grievanceData.grievanceType || '',              // Grievance Type
      grievanceData.articleViolated || '',            // Article Violated
      grievanceData.description || '',                // Description
      1,                                              // Current Step (starts at 1)
      filingDate,                                     // Step 1 Date
      deadlines.step1Due,                             // Step 1 Due
      'Pending',                                      // Step 1 Status
      '',                                             // Step 2 Date
      '',                                             // Step 2 Due
      '',                                             // Step 2 Status
      '',                                             // Step 3 Date
      '',                                             // Step 3 Due
      '',                                             // Step 3 Status
      '',                                             // Arbitration Date
      '',                                             // Resolution
      GRIEVANCE_OUTCOMES.PENDING,                     // Outcome
      '',                                             // Drive Folder
      grievanceData.notes || '',                      // Notes
      GRIEVANCE_STATUS.OPEN,                          // Status
      new Date()                                      // Last Updated
    ];

    // Append to sheet
    grievanceSheet.appendRow(rowData);

    // Log the creation
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_CREATED, {
      grievanceId: grievanceId,
      memberId: grievanceData.memberId,
      createdBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      message: `Grievance ${grievanceId} created successfully`
    };

  } catch (error) {
    console.error('Error creating grievance:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Handles grievance form submission from dialog
 * @param {Object} formData - Form data from UI
 * @return {Object} Result object
 */
function onGrievanceFormSubmit(formData) {
  // Map form data to grievance data structure
  const grievanceData = {
    memberId: formData.memberId,
    memberName: formData.memberName,
    filingDate: formData.filingDate,
    grievanceType: formData.grievanceType,
    articleViolated: formData.articleViolated,
    description: formData.description,
    notes: formData.notes
  };

  const result = startNewGrievance(grievanceData);

  if (result.success) {
    // Optionally create Drive folder
    if (formData.createFolder) {
      setupDriveFolderForGrievance(result.grievanceId);
    }

    // Optionally sync to calendar
    if (formData.syncCalendar) {
      syncSingleGrievanceToCalendar(result.grievanceId);
    }

    showToast(`Grievance ${result.grievanceId} created!`, 'Success');
  }

  return result;
}

/**
 * Validates grievance data before creation
 * @param {Object} data - Grievance data to validate
 * @return {Object} Validation result
 */
function validateGrievanceData(data) {
  if (!data.memberName && !data.memberId) {
    return { valid: false, error: 'Member name or ID is required' };
  }

  if (!data.grievanceType) {
    return { valid: false, error: 'Grievance type is required' };
  }

  if (!data.description || data.description.trim().length < 10) {
    return { valid: false, error: 'Description must be at least 10 characters' };
  }

  return { valid: true };
}

/**
 * Gets the next available grievance ID
 * @param {Sheet} sheet - Grievance tracker sheet
 * @return {string} Next grievance ID
 */
function getNextGrievanceId(sheet) {
  const data = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  let maxSequence = 0;

  // Find highest sequence number for current year
  for (let i = 1; i < data.length; i++) {
    const id = data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID];
    if (id && typeof id === 'string') {
      const match = id.match(/^GRV-(\d{4})-(\d{4})$/);
      if (match && parseInt(match[1]) === currentYear) {
        maxSequence = Math.max(maxSequence, parseInt(match[2]));
      }
    }
  }

  return generateGrievanceId(maxSequence + 1);
}

// ============================================================================
// DEADLINE CALCULATIONS - Based on Article 23A
// ============================================================================

/**
 * Calculates initial deadlines for a new grievance
 * @param {Date} filingDate - The grievance filing date
 * @return {Object} Calculated deadline dates
 */
function calculateInitialDeadlines(filingDate) {
  const step1Due = addBusinessDays(filingDate, DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);

  return {
    step1Due: step1Due
  };
}

/**
 * Calculates deadline for advancing to next step
 * @param {number} currentStep - Current grievance step (1-3)
 * @param {Date} currentStepDate - Date of current step
 * @return {Date} Deadline for next step
 */
function calculateNextStepDeadline(currentStep, currentStepDate) {
  let daysToAdd;

  switch (currentStep) {
    case 1:
      daysToAdd = DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL;
      break;
    case 2:
      daysToAdd = DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL;
      break;
    case 3:
      daysToAdd = DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND;
      break;
    default:
      return null;
  }

  return addBusinessDays(currentStepDate, daysToAdd);
}

/**
 * Calculates response deadline for a given step
 * @param {number} step - Grievance step
 * @param {Date} stepDate - Date step was initiated
 * @return {Date} Response deadline
 */
function calculateResponseDeadline(step, stepDate) {
  let daysToAdd;

  switch (step) {
    case 1:
      daysToAdd = DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE;
      break;
    case 2:
      daysToAdd = DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE;
      break;
    case 3:
      daysToAdd = DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE;
      break;
    default:
      return null;
  }

  return addBusinessDays(stepDate, daysToAdd);
}

/**
 * Adds business days to a date (excludes weekends)
 * @param {Date} startDate - Starting date
 * @param {number} days - Number of business days to add
 * @return {Date} Resulting date
 */
function addBusinessDays(startDate, days) {
  const result = new Date(startDate);
  let addedDays = 0;

  while (addedDays < days) {
    result.setDate(result.getDate() + 1);
    const dayOfWeek = result.getDay();
    // Skip weekends (0 = Sunday, 6 = Saturday)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      addedDays++;
    }
  }

  return result;
}

/**
 * Calculates days remaining until a deadline
 * @param {Date} deadline - Deadline date
 * @return {number} Days remaining (negative if overdue)
 */
function getDaysUntilDeadline(deadline) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const deadlineDate = new Date(deadline);
  deadlineDate.setHours(0, 0, 0, 0);

  const diffTime = deadlineDate - today;
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

// ============================================================================
// STEP ADVANCEMENT
// ============================================================================

/**
 * Advances a grievance to the next step
 * @param {string} grievanceId - The grievance ID to advance
 * @param {Object} options - Additional options (response, notes, etc.)
 * @return {Object} Result object
 */
function advanceGrievanceStep(grievanceId, options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    const data = sheet.getDataRange().getValues();

    // Find the grievance row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
        rowIndex = i + 1; // 1-indexed for sheet operations
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Grievance not found' };
    }

    const currentStep = data[rowIndex - 1][GRIEVANCE_COLUMNS.CURRENT_STEP];
    const nextStep = currentStep + 1;

    if (nextStep > 4) {
      return { success: false, error: 'Grievance is already at arbitration level' };
    }

    const today = new Date();
    const responseDue = calculateResponseDeadline(nextStep, today);

    // Update current step status to completed/appealed
    const currentStepStatusCol = getStepStatusColumn(currentStep);
    sheet.getRange(rowIndex, currentStepStatusCol).setValue(options.currentStepOutcome || 'Appealed');

    // Update next step columns
    if (nextStep <= 3) {
      const nextStepDateCol = getStepDateColumn(nextStep);
      const nextStepDueCol = nextStepDateCol + 1;
      const nextStepStatusCol = nextStepDateCol + 2;

      sheet.getRange(rowIndex, nextStepDateCol).setValue(today);
      sheet.getRange(rowIndex, nextStepDueCol).setValue(responseDue);
      sheet.getRange(rowIndex, nextStepStatusCol).setValue('Pending');
    } else {
      // Arbitration
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.ARBITRATION_DATE + 1).setValue(today);
    }

    // Update current step and status
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.CURRENT_STEP + 1).setValue(nextStep);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(
      nextStep === 4 ? GRIEVANCE_STATUS.AT_ARBITRATION : GRIEVANCE_STATUS.APPEALED
    );
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(today);

    // Add notes if provided
    if (options.notes) {
      const existingNotes = data[rowIndex - 1][GRIEVANCE_COLUMNS.NOTES] || '';
      const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                       `[${timestamp}] Step ${currentStep} -> ${nextStep}: ${options.notes}`;
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(newNotes);
    }

    // Log the advancement
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_STEP_ADVANCED, {
      grievanceId: grievanceId,
      fromStep: currentStep,
      toStep: nextStep,
      advancedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      newStep: nextStep,
      message: `Grievance advanced to Step ${nextStep}`
    };

  } catch (error) {
    console.error('Error advancing grievance:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets column index for step date
 * @param {number} step - Step number
 * @return {number} 1-indexed column number
 */
function getStepDateColumn(step) {
  switch (step) {
    case 1: return GRIEVANCE_COLUMNS.STEP_1_DATE + 1;
    case 2: return GRIEVANCE_COLUMNS.STEP_2_DATE + 1;
    case 3: return GRIEVANCE_COLUMNS.STEP_3_DATE + 1;
    default: return null;
  }
}

/**
 * Gets column index for step status
 * @param {number} step - Step number
 * @return {number} 1-indexed column number
 */
function getStepStatusColumn(step) {
  switch (step) {
    case 1: return GRIEVANCE_COLUMNS.STEP_1_STATUS + 1;
    case 2: return GRIEVANCE_COLUMNS.STEP_2_STATUS + 1;
    case 3: return GRIEVANCE_COLUMNS.STEP_3_STATUS + 1;
    default: return null;
  }
}

// ============================================================================
// BATCH OPERATIONS
// ============================================================================

/**
 * Recalculates all grievance deadlines in batches
 * Used when deadline rules change or for data repair
 * @return {Object} Result with count of updated grievances
 */
function recalcAllGrievancesBatched() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  let updatedCount = 0;
  let batchCount = 0;
  const startTime = new Date().getTime();

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    // Check execution time limit
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
      console.log('Approaching time limit, stopping batch');
      break;
    }

    // Batch pause
    if (batchCount >= BATCH_LIMITS.MAX_ROWS_PER_BATCH) {
      Utilities.sleep(BATCH_LIMITS.PAUSE_BETWEEN_BATCHES_MS);
      batchCount = 0;
    }

    const row = data[i];
    const status = row[GRIEVANCE_COLUMNS.STATUS];

    // Only recalculate open/pending grievances
    if (status === GRIEVANCE_STATUS.OPEN ||
        status === GRIEVANCE_STATUS.PENDING ||
        status === GRIEVANCE_STATUS.APPEALED) {

      const currentStep = row[GRIEVANCE_COLUMNS.CURRENT_STEP];
      const stepDate = row[getStepDateColumn(currentStep) - 1]; // 0-indexed for data array

      if (stepDate instanceof Date) {
        const newDue = calculateResponseDeadline(currentStep, stepDate);
        const dueColumn = getStepDateColumn(currentStep) + 1; // Due is next column after date

        sheet.getRange(i + 1, dueColumn).setValue(newDue);
        updatedCount++;
      }
    }

    batchCount++;
  }

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'RECALC_DEADLINES',
    grievancesUpdated: updatedCount,
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    updatedCount: updatedCount,
    message: `Recalculated deadlines for ${updatedCount} grievances`
  };
}

/**
 * Updates status for multiple grievances at once
 * @param {string[]} grievanceIds - Array of grievance IDs
 * @param {string} newStatus - New status to set
 * @param {string} notes - Optional notes to add
 * @return {Object} Result object
 */
function bulkUpdateGrievanceStatus(grievanceIds, newStatus, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  let updatedCount = 0;
  const today = new Date();
  const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

  for (let i = 1; i < data.length; i++) {
    const grievanceId = data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID];

    if (grievanceIds.includes(grievanceId)) {
      const rowIndex = i + 1;

      // Update status
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(newStatus);
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(today);

      // Add notes if provided
      if (notes) {
        const existingNotes = data[i][GRIEVANCE_COLUMNS.NOTES] || '';
        const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                         `[${timestamp}] Bulk status update to "${newStatus}": ${notes}`;
        sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(newNotes);
      }

      updatedCount++;
    }
  }

  return {
    success: true,
    updatedCount: updatedCount,
    message: `Updated status for ${updatedCount} grievances`
  };
}

// ============================================================================
// GRIEVANCE QUERIES
// ============================================================================

/**
 * Gets grievance data by ID
 * @param {string} grievanceId - The grievance ID
 * @return {Object|null} Grievance data or null if not found
 */
function getGrievanceById(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
      const grievance = {};
      headers.forEach((header, index) => {
        grievance[header] = data[i][index];
      });
      return grievance;
    }
  }

  return null;
}

/**
 * Gets all open grievances
 * @return {Array} Array of open grievance objects
 */
function getOpenGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const openStatuses = [
    GRIEVANCE_STATUS.OPEN,
    GRIEVANCE_STATUS.PENDING,
    GRIEVANCE_STATUS.APPEALED,
    GRIEVANCE_STATUS.AT_ARBITRATION
  ];

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][GRIEVANCE_COLUMNS.STATUS];
    if (openStatuses.includes(status)) {
      const grievance = {};
      headers.forEach((header, index) => {
        grievance[header] = data[i][index];
      });
      results.push(grievance);
    }
  }

  return results;
}

/**
 * Gets grievances with upcoming deadlines
 * @param {number} daysAhead - Number of days to look ahead
 * @return {Array} Array of grievances with deadline info
 */
function getUpcomingDeadlines(daysAhead) {
  const openGrievances = getOpenGrievances();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() + daysAhead);

  const upcoming = [];

  openGrievances.forEach(g => {
    const currentStep = g['Current Step'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.CURRENT_STEP]];
    let deadline;

    // Get the due date for current step
    switch (currentStep) {
      case 1:
        deadline = g['Step 1 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_1_DUE]];
        break;
      case 2:
        deadline = g['Step 2 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_2_DUE]];
        break;
      case 3:
        deadline = g['Step 3 Due'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.STEP_3_DUE]];
        break;
    }

    if (deadline instanceof Date) {
      const deadlineDate = new Date(deadline);
      deadlineDate.setHours(0, 0, 0, 0);

      if (deadlineDate <= cutoffDate) {
        const daysLeft = Math.ceil((deadlineDate - today) / (1000 * 60 * 60 * 24));
        upcoming.push({
          grievanceId: g['Grievance ID'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.GRIEVANCE_ID]],
          memberName: g['Member Name'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.MEMBER_NAME]],
          step: `Step ${currentStep}`,
          deadline: deadline,
          date: Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
          daysLeft: daysLeft
        });
      }
    }
  });

  // Sort by deadline (soonest first)
  upcoming.sort((a, b) => a.deadline - b.deadline);

  return upcoming;
}

/**
 * Gets grievance statistics for dashboard
 * @return {Object} Statistics object
 */
function getGrievanceStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  const currentYear = new Date().getFullYear();
  let open = 0;
  let pending = 0;
  let resolved = 0;
  let closedThisYear = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][GRIEVANCE_COLUMNS.STATUS];
    const lastUpdated = data[i][GRIEVANCE_COLUMNS.LAST_UPDATED];

    switch (status) {
      case GRIEVANCE_STATUS.OPEN:
        open++;
        break;
      case GRIEVANCE_STATUS.PENDING:
      case GRIEVANCE_STATUS.APPEALED:
        pending++;
        break;
      case GRIEVANCE_STATUS.RESOLVED:
      case GRIEVANCE_STATUS.CLOSED:
        if (lastUpdated instanceof Date && lastUpdated.getFullYear() === currentYear) {
          closedThisYear++;
        }
        resolved++;
        break;
    }
  }

  return {
    open: open,
    pending: pending,
    resolved: resolved,
    closedThisYear: closedThisYear,
    total: data.length - 1
  };
}

// ============================================================================
// GRIEVANCE RESOLUTION
// ============================================================================

/**
 * Resolves a grievance with outcome
 * @param {string} grievanceId - The grievance ID
 * @param {string} outcome - Resolution outcome
 * @param {string} resolution - Resolution description
 * @param {string} notes - Additional notes
 * @return {Object} Result object
 */
function resolveGrievance(grievanceId, outcome, resolution, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Grievance not found' };
    }

    const today = new Date();
    const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');

    // Update resolution fields
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.RESOLUTION + 1).setValue(resolution);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.OUTCOME + 1).setValue(outcome);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(GRIEVANCE_STATUS.RESOLVED);
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(today);

    // Mark current step as completed
    const currentStep = data[rowIndex - 1][GRIEVANCE_COLUMNS.CURRENT_STEP];
    const stepStatusCol = getStepStatusColumn(currentStep);
    if (stepStatusCol) {
      sheet.getRange(rowIndex, stepStatusCol).setValue(outcome);
    }

    // Add resolution notes
    if (notes) {
      const existingNotes = data[rowIndex - 1][GRIEVANCE_COLUMNS.NOTES] || '';
      const newNotes = existingNotes + (existingNotes ? '\n' : '') +
                       `[${timestamp}] RESOLVED - ${outcome}: ${notes}`;
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(newNotes);
    }

    // Log the resolution
    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      action: 'RESOLVED',
      outcome: outcome,
      resolvedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      grievanceId: grievanceId,
      message: `Grievance ${grievanceId} resolved as "${outcome}"`
    };

  } catch (error) {
    console.error('Error resolving grievance:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// UI DIALOG TRIGGERS
// ============================================================================

/**
 * Shows new grievance dialog
 */
function showNewGrievanceDialog() {
  const html = HtmlService.createHtmlOutput(getNewGrievanceFormHtml())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);

  SpreadsheetApp.getUi().showModalDialog(html, 'New Grievance');
}

/**
 * Shows edit grievance dialog for selected row
 */
function showEditGrievanceDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.GRIEVANCE_TRACKER) {
    showAlert('Please select a grievance in the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    showAlert('Please select a grievance row', 'No Selection');
    return;
  }

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const grievanceId = data[GRIEVANCE_COLUMNS.GRIEVANCE_ID];

  const html = HtmlService.createHtmlOutput(getEditGrievanceFormHtml(grievanceId))
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);

  SpreadsheetApp.getUi().showModalDialog(html, `Edit Grievance: ${grievanceId}`);
}

/**
 * Shows bulk status update dialog
 */
function showBulkStatusUpdate() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.GRIEVANCE_TRACKER) {
    showAlert('Please use this from the Grievance Tracker sheet', 'Wrong Sheet');
    return;
  }

  const openGrievances = getOpenGrievances();
  const items = openGrievances.map(g => ({
    id: g['Grievance ID'] || g[Object.keys(g)[GRIEVANCE_COLUMNS.GRIEVANCE_ID]],
    label: `${g['Grievance ID']} - ${g['Member Name']}`,
    selected: false
  }));

  showMultiSelectDialog('Select Grievances to Update', items, 'handleBulkStatusSelection');
}

/**
 * Generates HTML for new grievance form
 * @return {string} HTML content
 */
function getNewGrievanceFormHtml() {
  // Get member list for dropdown
  const members = getMemberList();
  const memberOptions = members.map(m =>
    `<option value="${m.id}">${m.name} (${m.department})</option>`
  ).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .form-actions { margin-top: 20px; display: flex; justify-content: flex-end; gap: 10px; }
        .checkbox-row { display: flex; gap: 20px; margin-top: 15px; }
        .checkbox-label { display: flex; align-items: center; gap: 8px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <form id="grievanceForm">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Member *</label>
              <select class="form-select" id="memberId" required>
                <option value="">Select a member...</option>
                ${memberOptions}
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Filing Date *</label>
              <input type="date" class="form-input" id="filingDate" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Grievance Type *</label>
              <select class="form-select" id="grievanceType" required>
                <option value="">Select type...</option>
                <option value="Contract Violation">Contract Violation</option>
                <option value="Discipline">Discipline</option>
                <option value="Discharge">Discharge</option>
                <option value="Working Conditions">Working Conditions</option>
                <option value="Safety">Safety</option>
                <option value="Other">Other</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Article Violated</label>
              <input type="text" class="form-input" id="articleViolated"
                     placeholder="e.g., Article 23A, Section 5">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Description *</label>
            <textarea class="form-textarea" id="description" required
                      placeholder="Describe the grievance in detail..."></textarea>
          </div>

          <div class="form-group">
            <label class="form-label">Initial Notes</label>
            <textarea class="form-textarea" id="notes" rows="2"
                      placeholder="Any additional notes..."></textarea>
          </div>

          <div class="checkbox-row">
            <label class="checkbox-label">
              <input type="checkbox" id="createFolder" checked>
              Create Drive folder for documents
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="syncCalendar" checked>
              Sync deadlines to calendar
            </label>
          </div>

          <div class="form-actions">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="submit" class="btn btn-primary">Create Grievance</button>
          </div>
        </form>
      </div>

      <script>
        // Set default date to today
        document.getElementById('filingDate').valueAsDate = new Date();

        document.getElementById('grievanceForm').addEventListener('submit', function(e) {
          e.preventDefault();

          const memberSelect = document.getElementById('memberId');
          const formData = {
            memberId: memberSelect.value,
            memberName: memberSelect.options[memberSelect.selectedIndex].text.split(' (')[0],
            filingDate: document.getElementById('filingDate').value,
            grievanceType: document.getElementById('grievanceType').value,
            articleViolated: document.getElementById('articleViolated').value,
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value,
            createFolder: document.getElementById('createFolder').checked,
            syncCalendar: document.getElementById('syncCalendar').checked
          };

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                google.script.host.close();
              } else {
                alert('Error: ' + result.error);
              }
            })
            .withFailureHandler(function(e) {
              alert('Error: ' + e.message);
            })
            .onGrievanceFormSubmit(formData);
        });
      </script>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for edit grievance form
 * @param {string} grievanceId - The grievance ID to edit
 * @return {string} HTML content
 */
function getEditGrievanceFormHtml(grievanceId) {
  const grievance = getGrievanceById(grievanceId);

  if (!grievance) {
    return '<p>Grievance not found</p>';
  }

  // Similar to new form but pre-populated with existing data
  // Abbreviated for space - would include full edit form
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-actions { margin-top: 20px; display: flex; justify-content: flex-end; gap: 10px; }
        .status-badge { padding: 4px 12px; border-radius: 12px; font-size: 12px; }
        .status-open { background: #e8f0fe; color: #1967d2; }
        .status-pending { background: #fef7e0; color: #ea8600; }
        .info-box { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <div class="info-box">
          <strong>Grievance ID:</strong> ${grievanceId}<br>
          <strong>Current Step:</strong> ${grievance['Current Step'] || 1}<br>
          <strong>Status:</strong> <span class="status-badge status-open">
            ${grievance['Status'] || 'Open'}</span>
        </div>

        <form id="editForm">
          <div class="form-group">
            <label class="form-label">Description</label>
            <textarea class="form-textarea" id="description">${grievance['Description'] || ''}</textarea>
          </div>

          <div class="form-group">
            <label class="form-label">Notes</label>
            <textarea class="form-textarea" id="notes">${grievance['Notes'] || ''}</textarea>
          </div>

          <div class="form-actions">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="button" class="btn btn-danger" onclick="advanceStep()">
              Advance Step
            </button>
            <button type="submit" class="btn btn-primary">Save Changes</button>
          </div>
        </form>
      </div>

      <script>
        const grievanceId = '${grievanceId}';

        document.getElementById('editForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const updates = {
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value
          };
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .updateGrievance(grievanceId, updates);
        });

        function advanceStep() {
          if (confirm('Advance this grievance to the next step?')) {
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  alert(result.message);
                  google.script.host.close();
                } else {
                  alert('Error: ' + result.error);
                }
              })
              .advanceGrievanceStep(grievanceId, {});
          }
        }
      </script>
    </body>
    </html>
  `;
}
