/**
 * ============================================================================
 * Main.gs - Dashboard Entry Point & Triggers
 * ============================================================================
 *
 * This is the main entry point for the Union Steward Dashboard.
 * It contains trigger functions and coordinates between all modules.
 *
 * Module Architecture:
 * - Constants.gs     : Configuration and constants (single source of truth)
 * - UIService.gs     : Dialogs, sidebars, and UI components
 * - GrievanceManager.gs : Grievance lifecycle management
 * - Integrations.gs  : Drive, Calendar, and email services
 * - Maintenance.gs   : Admin tools and diagnostics
 * - FormulaService.gs: Hidden sheet and formula logic
 *
 * Build Instructions:
 * During development, keep files separate. Use build.js to merge all files
 * into ConsolidatedDashboard.gs for deployment:
 *   node build.js
 *
 * @fileoverview Main entry point and trigger functions
 * @version 2.0.0
 * @author Dashboard Team
 */

// ============================================================================
// TRIGGER FUNCTIONS
// ============================================================================

/**
 * Runs when the spreadsheet is opened
 * Sets up the custom menu and initializes the dashboard
 */
function onOpen() {
  try {
    // Create the dashboard menu
    createDashboardMenu();

    // Show welcome toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Dashboard loaded successfully',
      '🏛️ Union Dashboard',
      3
    );

  } catch (error) {
    console.error('Error in onOpen:', error);
    // Still try to create a basic menu
    SpreadsheetApp.getUi()
      .createMenu('Union Dashboard')
      .addItem('Initialize Dashboard', 'initializeDashboard')
      .addToUi();
  }
}

/**
 * Runs when a cell is edited
 * Handles auto-calculations and validations
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // Handle edits in Grievance Tracker
    if (sheetName === SHEET_NAMES.GRIEVANCE_TRACKER) {
      handleGrievanceEdit(e);
    }

    // Handle edits in Member Directory
    if (sheetName === SHEET_NAMES.MEMBER_DIRECTORY) {
      handleMemberEdit(e);
    }

  } catch (error) {
    console.error('Error in onEdit:', error);
    // Don't show error to user for automatic functions
  }
}

/**
 * Handles edits to the Grievance Tracker sheet
 * @param {Object} e - The edit event object
 */
function handleGrievanceEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Skip header row
  if (row <= 1) return;

  const sheet = e.range.getSheet();

  // Auto-update Last Updated timestamp
  if (col !== GRIEVANCE_COLUMNS.LAST_UPDATED + 1) {
    sheet.getRange(row, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(new Date());
  }

  // If status changed, check for auto-actions
  if (col === GRIEVANCE_COLUMNS.STATUS + 1) {
    const settings = getSettings();
    const grievanceId = sheet.getRange(row, GRIEVANCE_COLUMNS.GRIEVANCE_ID + 1).getValue();

    // Auto-sync to calendar if enabled
    if (settings.autoSyncCalendar && grievanceId) {
      syncSingleGrievanceToCalendar(grievanceId);
    }
  }

  // If step dates changed, recalculate deadlines
  const stepDateColumns = [
    GRIEVANCE_COLUMNS.STEP_1_DATE + 1,
    GRIEVANCE_COLUMNS.STEP_2_DATE + 1,
    GRIEVANCE_COLUMNS.STEP_3_DATE + 1
  ];

  if (stepDateColumns.includes(col)) {
    const step = stepDateColumns.indexOf(col) + 1;
    const stepDate = e.value;

    if (stepDate) {
      const deadline = calculateResponseDeadline(step, new Date(stepDate));
      if (deadline) {
        sheet.getRange(row, col + 1).setValue(deadline); // Due column is next to date
      }
    }
  }
}

/**
 * Handles edits to the Member Directory sheet
 * @param {Object} e - The edit event object
 */
function handleMemberEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Skip header row
  if (row <= 1) return;

  const sheet = e.range.getSheet();

  // Auto-update Last Updated timestamp
  if (col !== MEMBER_COLUMNS.LAST_UPDATED + 1) {
    sheet.getRange(row, MEMBER_COLUMNS.LAST_UPDATED + 1).setValue(new Date());
  }

  // Validate email format
  if (col === MEMBER_COLUMNS.EMAIL + 1) {
    const email = e.value;
    if (email && !VALIDATION_RULES.EMAIL_PATTERN.test(email)) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Warning: Email format may be invalid',
        'Validation',
        3
      );
    }
  }

  // Validate phone format
  if (col === MEMBER_COLUMNS.PHONE + 1) {
    const phone = e.value;
    if (phone && !VALIDATION_RULES.PHONE_PATTERN.test(phone)) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Warning: Phone format may be invalid (expected: XXX-XXX-XXXX)',
        'Validation',
        3
      );
    }
  }
}

/**
 * Runs on a time-based trigger (daily)
 * Handles scheduled tasks like deadline reminders
 */
function dailyTrigger() {
  try {
    const settings = getSettings();

    // Send deadline reminders if enabled
    if (settings.emailReminders) {
      sendDeadlineReminders(settings.reminderDays);
    }

    // Log the trigger run
    logAuditEvent('DAILY_TRIGGER', {
      timestamp: new Date().toISOString(),
      remindersSent: settings.emailReminders
    });

  } catch (error) {
    console.error('Error in dailyTrigger:', error);
  }
}

// ============================================================================
// INITIALIZATION FUNCTIONS
// ============================================================================

/**
 * Initializes the dashboard for first-time setup
 * Creates all required sheets and configurations
 */
function initializeDashboard() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Initialize Dashboard',
    'This will set up the Union Steward Dashboard with all required sheets and configurations. ' +
    'Existing data will not be affected. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    showToast('Initializing dashboard...', 'Setup');

    // Create visible sheets if missing
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    for (const [key, sheetName] of Object.entries(SHEET_NAMES)) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        setupSheetStructure(sheet, key);
        showToast(`Created sheet: ${sheetName}`, 'Setup');
      }
    }

    // Setup hidden calculation sheets
    const hiddenResult = setupAllHiddenSheets();
    showToast(`Set up ${hiddenResult.created} hidden sheets`, 'Setup');

    // Setup triggers
    setupTriggers();

    // Create menu
    createDashboardMenu();

    // Log initialization
    logAuditEvent('DASHBOARD_INITIALIZED', {
      sheetsCreated: Object.keys(SHEET_NAMES).length,
      hiddenSheetsCreated: hiddenResult.created,
      initializedBy: Session.getActiveUser().getEmail()
    });

    ui.alert(
      'Setup Complete',
      'The Union Steward Dashboard has been initialized successfully!\n\n' +
      'Use the "Union Dashboard" menu to access all features.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Error initializing dashboard:', error);
    ui.alert('Error', 'Failed to initialize: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Sets up time-based triggers for the dashboard
 */
function setupTriggers() {
  // Remove existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create daily trigger at 8 AM
  ScriptApp.newTrigger('dailyTrigger')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

/**
 * Removes all triggers (for cleanup)
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
}

// ============================================================================
// HELP & DOCUMENTATION
// ============================================================================

/**
 * Shows the help dialog
 */
function showHelpDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .help-container { padding: 20px; }
        .help-section { margin-bottom: 25px; }
        .help-title { font-size: 18px; font-weight: 600; margin-bottom: 10px; color: #1a73e8; }
        .help-text { color: #5f6368; line-height: 1.6; }
        .shortcut-list { margin: 10px 0; }
        .shortcut-item { display: flex; margin: 8px 0; }
        .shortcut-key { background: #f1f3f4; padding: 4px 8px; border-radius: 4px;
                        font-family: monospace; margin-right: 10px; min-width: 120px; }
        .version { color: #999; font-size: 12px; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="help-container">
        <div class="help-section">
          <div class="help-title">Union Steward Dashboard</div>
          <div class="help-text">
            A comprehensive tool for managing union grievances, member records,
            and tracking deadlines based on your collective bargaining agreement.
          </div>
        </div>

        <div class="help-section">
          <div class="help-title">Key Features</div>
          <div class="help-text">
            <ul>
              <li><strong>Grievance Tracking</strong> - Manage grievances through all steps with automatic deadline calculations</li>
              <li><strong>Member Directory</strong> - Maintain member records with contact information and union status</li>
              <li><strong>Calendar Integration</strong> - Sync deadlines to Google Calendar for reminders</li>
              <li><strong>Drive Integration</strong> - Auto-create folders for grievance documentation</li>
              <li><strong>Self-Healing Formulas</strong> - Dashboard statistics update automatically</li>
            </ul>
          </div>
        </div>

        <div class="help-section">
          <div class="help-title">Quick Access</div>
          <div class="shortcut-list">
            <div class="shortcut-item">
              <span class="shortcut-key">Union Dashboard menu</span>
              <span>Access all dashboard features</span>
            </div>
            <div class="shortcut-item">
              <span class="shortcut-key">Search > Quick Search</span>
              <span>Fast search across all records</span>
            </div>
            <div class="shortcut-item">
              <span class="shortcut-key">Grievances > New</span>
              <span>File a new grievance</span>
            </div>
            <div class="shortcut-item">
              <span class="shortcut-key">Admin > Diagnostics</span>
              <span>Check system health</span>
            </div>
          </div>
        </div>

        <div class="help-section">
          <div class="help-title">Deadline Rules (Article 23A)</div>
          <div class="help-text">
            <ul>
              <li><strong>Step 1</strong> - ${DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE} days for management response</li>
              <li><strong>Step 2</strong> - ${DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL} days to appeal, ${DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE} days for response</li>
              <li><strong>Step 3</strong> - ${DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL} days to appeal, ${DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE} days for response</li>
              <li><strong>Arbitration</strong> - ${DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND} days to demand after Step 3</li>
            </ul>
          </div>
        </div>

        <div class="help-section">
          <div class="help-title">Need Help?</div>
          <div class="help-text">
            For technical issues, run <strong>Admin Tools > System Diagnostics</strong> to check for problems.
            Use <strong>Repair Dashboard</strong> to fix common issues automatically.
          </div>
        </div>

        <div class="version">
          Dashboard Version 2.0.0 | Modular Architecture
        </div>

        <div style="margin-top: 20px; text-align: right;">
          <button class="btn btn-primary" onclick="google.script.host.close()">Got it!</button>
        </div>
      </div>
    </body>
    </html>
  `).setWidth(550).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Documentation');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets the current version info
 * @return {Object} Version information
 */
function getVersionInfo() {
  return {
    version: '2.0.0',
    architecture: 'Modular Multi-File',
    modules: [
      'Constants.gs',
      'UIService.gs',
      'GrievanceManager.gs',
      'Integrations.gs',
      'Maintenance.gs',
      'FormulaService.gs',
      'Main.gs'
    ],
    lastUpdated: new Date().toISOString()
  };
}

/**
 * Updates a grievance with new data
 * @param {string} grievanceId - The grievance ID
 * @param {Object} updates - Fields to update
 * @return {Object} Result object
 */
function updateGrievance(grievanceId, updates) {
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

    // Update each provided field
    if (updates.description !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.DESCRIPTION + 1).setValue(updates.description);
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(updates.notes);
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(updates.status);
    }

    // Update timestamp
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(new Date());

    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      updates: Object.keys(updates),
      updatedBy: Session.getActiveUser().getEmail()
    });

    return { success: true, message: 'Grievance updated successfully' };

  } catch (error) {
    console.error('Error updating grievance:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Handles bulk status selection from multi-select dialog
 * @param {string[]} selectedIds - Selected grievance IDs
 */
function handleBulkStatusSelection(selectedIds) {
  if (!selectedIds || selectedIds.length === 0) {
    showAlert('No grievances selected', 'Bulk Update');
    return;
  }

  // Show status selection dialog
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
    </head>
    <body style="padding: 20px;">
      <h3>Update ${selectedIds.length} Grievances</h3>
      <div class="form-group">
        <label class="form-label">New Status</label>
        <select class="form-select" id="newStatus">
          ${Object.values(GRIEVANCE_STATUS).map(s =>
            `<option value="${s}">${s}</option>`
          ).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Notes (optional)</label>
        <textarea class="form-textarea" id="notes" rows="2"></textarea>
      </div>
      <div style="margin-top: 20px; text-align: right;">
        <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        <button class="btn btn-primary" onclick="applyUpdate()">Update All</button>
      </div>
      <script>
        const ids = ${JSON.stringify(selectedIds)};
        function applyUpdate() {
          const status = document.getElementById('newStatus').value;
          const notes = document.getElementById('notes').value;
          google.script.run
            .withSuccessHandler(function(r) {
              alert(r.success ? r.message : 'Error: ' + r.error);
              google.script.host.close();
            })
            .bulkUpdateGrievanceStatus(ids, status, notes);
        }
      </script>
    </body>
    </html>
  `).setWidth(400).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Status Update');
}

// ============================================================================
// MEMBER MANAGEMENT
// ============================================================================

/**
 * Shows new member dialog
 */
function showNewMemberDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <form id="memberForm">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">First Name *</label>
              <input type="text" class="form-input" id="firstName" required>
            </div>
            <div class="form-group">
              <label class="form-label">Last Name *</label>
              <input type="text" class="form-input" id="lastName" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Employee ID</label>
              <input type="text" class="form-input" id="employeeId" placeholder="XX000000">
            </div>
            <div class="form-group">
              <label class="form-label">Department *</label>
              <select class="form-select" id="department" required>
                <option value="">Select department...</option>
              </select>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Job Title</label>
              <input type="text" class="form-input" id="jobTitle">
            </div>
            <div class="form-group">
              <label class="form-label">Hire Date</label>
              <input type="date" class="form-input" id="hireDate">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Email</label>
              <input type="email" class="form-input" id="email">
            </div>
            <div class="form-group">
              <label class="form-label">Phone</label>
              <input type="tel" class="form-input" id="phone" placeholder="XXX-XXX-XXXX">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Union Status</label>
              <select class="form-select" id="unionStatus">
                <option value="Full Member">Full Member</option>
                <option value="Agency Fee">Agency Fee</option>
                <option value="Non-Member">Non-Member</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Status</label>
              <select class="form-select" id="status">
                <option value="Active">Active</option>
                <option value="On Leave">On Leave</option>
                <option value="Terminated">Terminated</option>
              </select>
            </div>
          </div>

          <div style="margin-top: 20px; text-align: right;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn btn-primary">Add Member</button>
          </div>
        </form>
      </div>

      <script>
        // Load departments
        google.script.run
          .withSuccessHandler(function(depts) {
            const select = document.getElementById('department');
            depts.forEach(function(d) {
              const opt = document.createElement('option');
              opt.value = d;
              opt.textContent = d;
              select.appendChild(opt);
            });
            // Add "Other" option
            const other = document.createElement('option');
            other.value = 'Other';
            other.textContent = 'Other';
            select.appendChild(other);
          })
          .getDepartmentList();

        document.getElementById('memberForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const memberData = {
            firstName: document.getElementById('firstName').value,
            lastName: document.getElementById('lastName').value,
            employeeId: document.getElementById('employeeId').value,
            department: document.getElementById('department').value,
            jobTitle: document.getElementById('jobTitle').value,
            hireDate: document.getElementById('hireDate').value,
            email: document.getElementById('email').value,
            phone: document.getElementById('phone').value,
            unionStatus: document.getElementById('unionStatus').value,
            status: document.getElementById('status').value
          };

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('Member added successfully!');
                google.script.host.close();
              } else {
                alert('Error: ' + result.error);
              }
            })
            .addNewMember(memberData);
        });
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Member');
}

/**
 * Adds a new member to the directory
 * @param {Object} memberData - Member information
 * @return {Object} Result object
 */
function addNewMember(memberData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

    if (!sheet) {
      return { success: false, error: 'Member Directory sheet not found' };
    }

    // Generate new member ID
    const lastRow = sheet.getLastRow();
    const newId = `MEM-${String(lastRow).padStart(5, '0')}`;

    // Prepare row data
    const rowData = [
      newId,
      memberData.firstName,
      memberData.lastName,
      memberData.employeeId,
      memberData.department,
      memberData.jobTitle,
      memberData.hireDate ? new Date(memberData.hireDate) : '',
      memberData.hireDate ? new Date(memberData.hireDate) : '', // Seniority same as hire
      memberData.email,
      memberData.phone,
      memberData.status || 'Active',
      memberData.unionStatus || 'Full Member',
      '',  // Notes
      new Date()  // Last Updated
    ];

    sheet.appendRow(rowData);

    logAuditEvent(AUDIT_EVENTS.MEMBER_ADDED, {
      memberId: newId,
      name: `${memberData.firstName} ${memberData.lastName}`,
      addedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      memberId: newId,
      message: 'Member added successfully'
    };

  } catch (error) {
    console.error('Error adding member:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Starts a grievance for a specific member (from quick action)
 */
function startGrievanceForMember() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.MEMBER_DIRECTORY) {
    showAlert('Please select a member in the Member Directory', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    showAlert('Please select a member row', 'No Selection');
    return;
  }

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const memberId = data[MEMBER_COLUMNS.ID];
  const memberName = `${data[MEMBER_COLUMNS.FIRST_NAME]} ${data[MEMBER_COLUMNS.LAST_NAME]}`;

  // Open new grievance dialog pre-populated with member info
  const html = HtmlService.createHtmlOutput(`
    <script>
      // Store member info for the form
      sessionStorage.setItem('prefillMemberId', '${memberId}');
      sessionStorage.setItem('prefillMemberName', '${memberName}');
      google.script.host.close();
      google.script.run.showNewGrievanceDialog();
    </script>
  `).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Loading...');
}
