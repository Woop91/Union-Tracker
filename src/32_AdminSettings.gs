/**
 * ============================================================================
 * 32_AdminSettings.gs — Admin Settings Sidebar
 * ============================================================================
 *
 * Provides a GUI sidebar for editing Config sheet settings. Replaces
 * direct Config sheet editing with a tabbed form interface grouped by
 * category (Organization, Employment & Lists, Grievance Settings,
 * Deadlines, Integrations, Branding & UX).
 *
 * WHAT THIS FILE DOES:
 *   Server-side functions for the Admin Settings sidebar. Reads from and
 *   writes to the Config sheet via CONFIG_COLS constants. Enforces admin
 *   authorization, input sanitization (escapeForFormula), and audit logging.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   The Config sheet has 65+ columns of horizontal data requiring extensive
 *   scrolling. This sidebar provides a clean form UI with inline descriptions,
 *   field grouping, and sensitive field masking — without changing any existing
 *   Config reading infrastructure (getConfigValue_, ConfigReader, etc.).
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The sidebar won't open or save settings. Config sheet still works directly
 *   as before — this is purely additive UI. No existing functionality affected.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, CONFIG_COLS), 00_Security.gs (escapeHtml,
 *               escapeForFormula), 00_DataAccess.gs (withScriptLock_),
 *               06_Maintenance.gs (logAuditEvent), 20_WebDashConfigReader.gs
 *               (ConfigReader.refreshConfig), 08a_SheetSetup.gs (getConfigValues,
 *               setupDataValidations).
 *   Used by:    AdminSettings.html (sidebar UI), 03_UIComponents.gs (menu item).
 *
 * @fileoverview Admin Settings sidebar server-side functions
 * @version 4.51.0
 * @requires 01_Core.gs
 * @requires 00_Security.gs
 */

// ============================================================================
// SETTINGS SCHEMA — Defines tabs, fields, types, and descriptions
// ============================================================================
// The schema drives both server-side reads/writes and client-side rendering.
// Types: 'text', 'number', 'url', 'email', 'toggle', 'list'
// Flags: sensitive (masked for non-owners), readOnly (cannot be edited)

var ADMIN_SETTINGS_SCHEMA_ = [
  {
    id: 'org',
    label: 'Organization',
    icon: '\uD83C\uDFE2',
    fields: [
      { key: 'ORG_NAME',            label: 'Organization Name',             type: 'text',  desc: 'Full name of your organization' },
      { key: 'LOCAL_NUMBER',         label: 'Local Number',                  type: 'text',  desc: 'Union local number' },
      { key: 'UNION_PARENT',        label: 'Union Parent',                  type: 'text',  desc: 'Parent union body (e.g. AFL-CIO)' },
      { key: 'STATE_REGION',        label: 'State/Region',                  type: 'text',  desc: 'Geographic region' },
      { key: 'ORG_WEBSITE',         label: 'Organization Website',          type: 'url',   desc: 'Public website URL' },
      { key: 'MAIN_ADDRESS',        label: 'Main Office Address',           type: 'text',  desc: 'Primary office address' },
      { key: 'MAIN_PHONE',          label: 'Main Phone',                    type: 'text',  desc: 'Main office phone number' },
      { key: 'MAIN_FAX',            label: 'Main Fax',                      type: 'text',  desc: 'Main office fax number' },
      { key: 'MAIN_CONTACT_NAME',   label: 'Main Contact Name',             type: 'text',  desc: 'Primary contact person' },
      { key: 'MAIN_CONTACT_EMAIL',  label: 'Main Contact Email',            type: 'email', desc: 'Primary contact email' },
      { key: 'CHIEF_STEWARD_EMAIL', label: 'Chief Steward Email',           type: 'email', desc: 'Chief steward email for notifications' },
      { key: 'CONTRACT_NAME',       label: 'Contract Name',                 type: 'text',  desc: 'Name of the collective bargaining agreement' },
      { key: 'CONTRACT_GRIEVANCE',  label: 'Contract Article (Grievance)',  type: 'text',  desc: 'Article number for grievance procedures' },
      { key: 'CONTRACT_DISCIPLINE', label: 'Contract Article (Discipline)', type: 'text',  desc: 'Article number for discipline procedures' },
      { key: 'CONTRACT_WORKLOAD',   label: 'Contract Article (Workload)',   type: 'text',  desc: 'Article number for workload provisions' },
      { key: 'ORG_ABBREV',          label: 'Organization Abbreviation',     type: 'text',  desc: 'Short abbreviation for the organization' }
    ]
  },
  {
    id: 'lists',
    label: 'Employment & Lists',
    icon: '\uD83D\uDCCB',
    fields: [
      { key: 'JOB_TITLES',             label: 'Job Titles',              type: 'list', desc: 'Available job titles for member profiles' },
      { key: 'OFFICE_LOCATIONS',        label: 'Office Locations',        type: 'list', desc: 'Work locations for member profiles' },
      { key: 'UNITS',                   label: 'Units',                   type: 'list', desc: 'Organizational units or departments' },
      { key: 'UNIT_CODES',             label: 'Unit Codes',              type: 'list', desc: 'Format: "Unit Name:CODE" (e.g. Main Station:MS)' },
      { key: 'OFFICE_DAYS',            label: 'Office Days',             type: 'list', desc: 'Days of the week for scheduling' },
      { key: 'OFFICE_ADDRESSES',       label: 'Office Addresses',        type: 'list', desc: 'Full addresses for office locations' },
      { key: 'SUPERVISORS',            label: 'Supervisors',             type: 'list', desc: 'Supervisor names for member profiles' },
      { key: 'MANAGERS',               label: 'Directors',               type: 'list', desc: 'Director names for member profiles' },
      { key: 'STEWARDS',               label: 'Stewards',                type: 'list', desc: 'Auto-synced with Member Directory steward status', note: 'Changes may be overwritten by steward promotions/demotions' },
      { key: 'STEWARD_COMMITTEES',     label: 'Steward Committees',      type: 'list', desc: 'Committee names for steward assignments' },
      { key: 'COMM_METHODS',           label: 'Communication Methods',   type: 'list', desc: 'Contact preference options (e.g. Email, Phone, Text)' },
      { key: 'BEST_TIMES',            label: 'Best Times to Contact',   type: 'list', desc: 'Time preferences for member contact' },
      { key: 'SURVEY_PRIORITY_OPTIONS', label: 'Survey Priority Options', type: 'list', desc: 'Priority level choices for survey questions' },
      { key: 'DUES_STATUSES',          label: 'Dues Statuses',            type: 'list', desc: 'Dues payment status options for member profiles' }
    ]
  },
  {
    id: 'grievance',
    label: 'Grievance Settings',
    icon: '\u2696\uFE0F',
    fields: [
      { key: 'GRIEVANCE_STATUS',       label: 'Grievance Status Values', type: 'list', desc: 'Status options for grievance tracking (e.g. Open, Won, Denied)' },
      { key: 'GRIEVANCE_STEP',         label: 'Grievance Step Values',   type: 'list', desc: 'Step levels (e.g. Step I, Step II, Arbitration)' },
      { key: 'ISSUE_CATEGORY',         label: 'Issue Categories',        type: 'list', desc: 'Types of grievance issues (multi-select enabled)' },
      { key: 'ARTICLES',               label: 'Articles Violated',       type: 'list', desc: 'Contract article references (multi-select enabled)' },
      { key: 'GRIEVANCE_COORDINATORS', label: 'Grievance Coordinators',  type: 'list', desc: 'People who coordinate grievance processing' },
      { key: 'ESCALATION_STATUSES',    label: 'Escalation Statuses',     type: 'list', desc: 'Status values that trigger escalation alerts' },
      { key: 'ESCALATION_STEPS',       label: 'Escalation Steps',        type: 'list', desc: 'Step values that trigger escalation alerts' }
    ]
  },
  {
    id: 'deadlines',
    label: 'Deadlines',
    icon: '\u23F0',
    fields: [
      { key: 'FILING_DEADLINE_DAYS',   label: 'Filing Deadline',          type: 'number', desc: 'Days to file a grievance after incident',              min: 1, max: 365 },
      { key: 'STEP1_RESPONSE_DAYS',    label: 'Step I Response',          type: 'number', desc: 'Days for management Step I response',                  min: 1, max: 365 },
      { key: 'STEP2_APPEAL_DAYS',      label: 'Step II Appeal',           type: 'number', desc: 'Days to appeal to Step II',                            min: 1, max: 365 },
      { key: 'STEP2_RESPONSE_DAYS',    label: 'Step II Response',         type: 'number', desc: 'Days for management Step II response',                 min: 1, max: 365 },
      { key: 'STEP3_APPEAL_DAYS',      label: 'Step III Appeal',          type: 'number', desc: 'Days to appeal to Step III',                           min: 1, max: 365 },
      { key: 'STEP3_RESPONSE_DAYS',    label: 'Step III Response',        type: 'number', desc: 'Days for management Step III response',                min: 1, max: 365 },
      { key: 'ARBITRATION_DEMAND_DAYS', label: 'Arbitration Demand',      type: 'number', desc: 'Days to demand arbitration',                           min: 1, max: 365 },
      { key: 'ALERT_DAYS',             label: 'Alert Days Before Deadline', type: 'text', desc: 'Comma-separated days before deadline to send alerts (e.g. 3,7,14)' },
      { key: 'GRIEVANCE_ARCHIVE_DAYS', label: 'Grievance Archive Days',  type: 'number', desc: 'Days before closed grievances are auto-archived',      min: 30, max: 3650 },
      { key: 'AUDIT_ARCHIVE_DAYS',     label: 'Audit Log Archive Days',  type: 'number', desc: 'Days before audit log entries are exported and pruned', min: 30, max: 3650 }
    ]
  },
  {
    id: 'integrations',
    label: 'Integrations',
    icon: '\uD83D\uDD17',
    fields: [
      { key: 'GRIEVANCE_FORM_URL',       label: 'Grievance Form URL',         type: 'url',   desc: 'Link to external grievance submission form' },
      { key: 'CONTACT_FORM_URL',         label: 'Contact Form URL',           type: 'url',   desc: 'Link to external contact form' },
      { key: 'MOBILE_DASHBOARD_URL',     label: 'Mobile Dashboard URL',       type: 'url',   desc: 'Auto-generated web app URL',                      readOnly: true },
      { key: 'DRIVE_FOLDER_ID',          label: 'Google Drive Folder ID',     type: 'text',  desc: 'Root Drive folder for documents',                  sensitive: true },
      { key: 'CALENDAR_ID',              label: 'Google Calendar ID',         type: 'text',  desc: 'Calendar for deadline tracking',                   sensitive: true },
      { key: 'ARCHIVE_FOLDER_ID',        label: 'Archive Folder ID',          type: 'text',  desc: 'Drive folder for archived documents',              sensitive: true },
      { key: 'TEMPLATE_ID',              label: 'Template Document ID',       type: 'text',  desc: 'Google Doc template for grievance PDFs',           sensitive: true },
      { key: 'PDF_FOLDER_ID',            label: 'PDF Folder ID',              type: 'text',  desc: 'Drive folder for generated PDFs',                  sensitive: true },
      { key: 'PAST_SURVEYS_FOLDER_ID',   label: 'Past Surveys Folder ID',    type: 'text',  desc: 'Drive folder for archived survey periods',         sensitive: true },
      { key: 'DASHBOARD_ROOT_FOLDER_ID', label: 'Dashboard Root Folder ID',  type: 'text',  desc: 'Set by CREATE_DASHBOARD \u2014 do not edit',       sensitive: true, readOnly: true },
      { key: 'GRIEVANCES_FOLDER_ID',     label: 'Grievances Folder ID',      type: 'text',  desc: 'Set by CREATE_DASHBOARD \u2014 do not edit',       sensitive: true, readOnly: true },
      { key: 'RESOURCES_FOLDER_ID',      label: 'Resources Folder ID',       type: 'text',  desc: 'Set by CREATE_DASHBOARD \u2014 do not edit',       sensitive: true, readOnly: true },
      { key: 'MINUTES_FOLDER_ID',        label: 'Minutes Folder ID',         type: 'text',  desc: 'Set by CREATE_DASHBOARD \u2014 do not edit',       sensitive: true, readOnly: true },
      { key: 'EVENT_CHECKIN_FOLDER_ID',  label: 'Event Check-In Folder ID',  type: 'text',  desc: 'Set by CREATE_DASHBOARD \u2014 do not edit',       sensitive: true, readOnly: true },
      { key: 'ADMIN_EMAILS',             label: 'Admin Emails',              type: 'list',  desc: 'Email addresses for admin notifications',           sensitive: true },
      { key: 'NOTIFICATION_RECIPIENTS',  label: 'Notification Recipients',   type: 'list',  desc: 'Email addresses for system notifications',         sensitive: true },
      { key: 'TEST_NOTIFY_EMAIL',        label: 'Test Runner Notify Email',  type: 'email', desc: 'Email for test failure notifications',              sensitive: true }
    ]
  },
  {
    id: 'branding',
    label: 'Branding & UX',
    icon: '\uD83C\uDFA8',
    fields: [
      { key: 'ACCENT_HUE',             label: 'Accent Hue',                type: 'number', desc: 'Color hue (0\u2013360) for UI theme accent',         min: 0, max: 360 },
      { key: 'LOGO_INITIALS',          label: 'Logo Initials',             type: 'text',   desc: 'Short text (2\u20133 chars) shown in the logo' },
      { key: 'STEWARD_LABEL',          label: 'Steward Label',             type: 'text',   desc: 'Custom label for steward role (e.g. "Rep")' },
      { key: 'MEMBER_LABEL',           label: 'Member Label',              type: 'text',   desc: 'Custom label for member role' },
      { key: 'MAGIC_LINK_EXPIRY_DAYS', label: 'Magic Link Expiry Days',    type: 'number', desc: 'Days before magic login links expire',               min: 1, max: 90 },
      { key: 'COOKIE_DURATION_DAYS',   label: 'Cookie Duration Days',      type: 'number', desc: 'Days before session cookies expire',                 min: 1, max: 365 },
      { key: 'INSIGHTS_CACHE_TTL_MIN', label: 'Insights Cache (Minutes)',  type: 'number', desc: 'Minutes to cache Insights page data',                min: 1, max: 60 },
      { key: 'BROADCAST_SCOPE_ALL',    label: 'Broadcast: All Members',    type: 'toggle', desc: 'Allow stewards to message ALL members (not just assigned)' },
      { key: 'ENABLE_CORRELATION',     label: 'Enable Correlation Engine', type: 'toggle', desc: 'Enable correlation analysis on the Insights page' },
      { key: 'SHOW_GRIEVANCES',        label: 'Show Grievances',           type: 'toggle', desc: 'Show grievance tracking (cases, new grievance, stats). Reload required.' },
      { key: 'ENABLE_TAB_MODALS',     label: 'Enable Tab Modals',         type: 'toggle', desc: 'Enable tab-based modal dialogs instead of sidebar modals. Reload required.' },
      { key: 'CUSTOM_LINK_1_NAME',     label: 'Custom Link 1 Name',        type: 'text',   desc: 'Display name for first custom sidebar link' },
      { key: 'CUSTOM_LINK_1_URL',      label: 'Custom Link 1 URL',         type: 'url',    desc: 'URL for first custom sidebar link' },
      { key: 'CUSTOM_LINK_2_NAME',     label: 'Custom Link 2 Name',        type: 'text',   desc: 'Display name for second custom sidebar link' },
      { key: 'CUSTOM_LINK_2_URL',      label: 'Custom Link 2 URL',         type: 'url',    desc: 'URL for second custom sidebar link' }
    ]
  }
];

// ============================================================================
// SIDEBAR ENTRY POINT
// ============================================================================

/**
 * Opens the Admin Settings sidebar.
 * Called from the Admin menu in the spreadsheet.
 */
function showAdminSettingsSidebar() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    SpreadsheetApp.getUi().alert('Could not identify your account. Please sign in and try again.');
    return;
  }

  if (!_adminIsAuthorized_(email)) {
    SpreadsheetApp.getUi().alert('Admin access required. Only the spreadsheet owner and designated admins can access settings.');
    return;
  }

  var html = HtmlService.createHtmlOutputFromFile('AdminSettings')
    .setTitle('Admin Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// AUTHORIZATION
// ============================================================================

/**
 * Checks whether the given email is authorized to access Admin Settings.
 *
 * @param {string} email - The user's email address
 * @returns {boolean} true if authorized
 * @private
 */
function _adminIsAuthorized_(email) {
  if (!email) return false;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;

    // Spreadsheet owner always has access
    var owner = ss.getOwner();
    if (owner && owner.getEmail().toLowerCase() === email.toLowerCase()) return true;

    // Check ADMIN_EMAILS list in Config sheet
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return false; // fail-closed: no Config = no access
    var adminEmails = getConfigValues(configSheet, CONFIG_COLS.ADMIN_EMAILS);
    var emailLower = email.toLowerCase();
    for (var i = 0; i < adminEmails.length; i++) {
      if (String(adminEmails[i]).trim().toLowerCase() === emailLower) return true;
    }
    return false;
  } catch (e) {
    log_('_adminIsAuthorized_ error', e.message);
    return false; // fail-closed on error
  }
}

// ============================================================================
// READ FUNCTIONS
// ============================================================================

/**
 * Returns the settings schema and current values for all tabs.
 * Called by the sidebar HTML via google.script.run.
 *
 * @returns {Object} { success, settings, schema, isOwner } or { success: false, error }
 */
function adminGetSettings() {
  var email = Session.getActiveUser().getEmail();
  if (!_adminIsAuthorized_(email)) {
    return { success: false, error: 'Admin access required' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return { success: false, error: 'Config sheet not found' };

    var owner = ss.getOwner();
    var isOwner = owner ? owner.getEmail().toLowerCase() === email.toLowerCase() : false;

    // Batch read: entire row 3 + full data block for list columns
    var lastCol = configSheet.getLastColumn();
    var lastRow = configSheet.getLastRow();
    var row3 = lastCol > 0 ? configSheet.getRange(3, 1, 1, lastCol).getValues()[0] : [];
    var allData = (lastRow >= 3 && lastCol > 0)
      ? configSheet.getRange(3, 1, lastRow - 2, lastCol).getValues()
      : [];

    var settings = {};

    for (var t = 0; t < ADMIN_SETTINGS_SCHEMA_.length; t++) {
      var tab = ADMIN_SETTINGS_SCHEMA_[t];
      for (var f = 0; f < tab.fields.length; f++) {
        var field = tab.fields[f];
        var col = CONFIG_COLS[field.key];
        if (!col) continue;

        if (field.type === 'list') {
          var listVals = [];
          for (var r = 0; r < allData.length; r++) {
            var cellVal = allData[r][col - 1];
            if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== '') {
              listVals.push(String(cellVal).trim());
            }
          }
          settings[field.key] = listVals;
        } else {
          var rawVal = (col <= row3.length) ? row3[col - 1] : '';
          var strVal = (rawVal !== null && rawVal !== undefined) ? String(rawVal).trim() : '';

          // Mask sensitive fields for non-owners
          if (field.sensitive && !isOwner && strVal) {
            settings[field.key] = '\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022';
          } else {
            settings[field.key] = escapeHtml(strVal);
          }
        }
      }
    }

    return {
      success: true,
      settings: settings,
      schema: ADMIN_SETTINGS_SCHEMA_,
      isOwner: isOwner
    };
  } catch (e) {
    log_('adminGetSettings error', e.message);
    return { success: false, error: 'Failed to load settings: ' + e.message };
  }
}

// ============================================================================
// WRITE FUNCTIONS
// ============================================================================

/**
 * Saves one or more single-value settings to the Config sheet.
 * Called by the sidebar when saving a tab's form fields.
 *
 * @param {Object} changes - { CONFIG_COLS_KEY: newValue, ... }
 * @returns {Object} { success: true, saved: number } or error
 */
function adminSaveSettings(changes) {
  var email = Session.getActiveUser().getEmail();
  if (!_adminIsAuthorized_(email)) {
    return { success: false, error: 'Admin access required' };
  }

  if (!changes || typeof changes !== 'object') {
    return { success: false, error: 'Invalid changes object' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return { success: false, error: 'Config sheet not found' };

    var owner = ss.getOwner();
    var isOwner = owner ? owner.getEmail().toLowerCase() === email.toLowerCase() : false;
    var saved = 0;

    var keys = Object.keys(changes);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      var col = CONFIG_COLS[key];
      if (!col) continue;

      var fieldDef = _findFieldDef_(key);
      if (!fieldDef) continue;
      if (fieldDef.readOnly) continue;
      if (fieldDef.sensitive && !isOwner) continue;
      if (fieldDef.type === 'list') continue; // Lists use adminSaveListValues

      var newValue = changes[key];

      // Validate URLs
      if (fieldDef.type === 'url' && newValue && !/^https?:\/\//i.test(newValue)) {
        return { success: false, error: fieldDef.label + ' must start with http:// or https://' };
      }

      // Validate numbers
      if (fieldDef.type === 'number' && newValue !== '' && newValue !== null) {
        var num = Number(newValue);
        if (isNaN(num)) return { success: false, error: fieldDef.label + ' must be a number' };
        if (fieldDef.min !== undefined && num < fieldDef.min) {
          return { success: false, error: fieldDef.label + ' minimum is ' + fieldDef.min };
        }
        if (fieldDef.max !== undefined && num > fieldDef.max) {
          return { success: false, error: fieldDef.label + ' maximum is ' + fieldDef.max };
        }
      }

      // Toggle normalization
      if (fieldDef.type === 'toggle') {
        newValue = (newValue === true || newValue === 'yes') ? 'yes' : 'no';
      }

      var oldValue = configSheet.getRange(3, col).getValue();
      configSheet.getRange(3, col).setValue(escapeForFormula(String(newValue)));
      saved++;

      logAuditEvent('ADMIN_SETTINGS_CHANGE', {
        field: key,
        oldValue: String(oldValue),
        newValue: String(newValue),
        changedBy: email
      });
    }

    // Bust caches after writes
    if (saved > 0) {
      try { ConfigReader.refreshConfig(); } catch (_e) { log_('Cache refresh', (_e.message || _e)); }
    }

    return { success: true, saved: saved };
  } catch (e) {
    log_('adminSaveSettings error', e.message);
    return { success: false, error: 'Save failed: ' + e.message };
  }
}

/**
 * Saves a list of values to a Config column (replaces all existing values).
 * Used for dropdown/list fields like Job Titles, Stewards, etc.
 *
 * @param {string} configColKey - CONFIG_COLS key (e.g. 'JOB_TITLES')
 * @param {string[]} values - Array of string values
 * @returns {Object} { success: true, count: number } or error
 */
function adminSaveListValues(configColKey, values) {
  var email = Session.getActiveUser().getEmail();
  if (!_adminIsAuthorized_(email)) {
    return { success: false, error: 'Admin access required' };
  }

  var col = CONFIG_COLS[configColKey];
  if (!col) return { success: false, error: 'Unknown config key: ' + configColKey };

  var fieldDef = _findFieldDef_(configColKey);
  if (!fieldDef || fieldDef.type !== 'list') {
    return { success: false, error: configColKey + ' is not a list field' };
  }
  if (fieldDef.readOnly) return { success: false, error: configColKey + ' is read-only' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = ss.getOwner(); var isOwner = owner ? owner.getEmail().toLowerCase() === email.toLowerCase() : false;
  if (fieldDef.sensitive && !isOwner) {
    return { success: false, error: 'Only the spreadsheet owner can edit ' + fieldDef.label };
  }

  if (!Array.isArray(values)) {
    return { success: false, error: 'Values must be an array' };
  }

  try {
    var result = { success: false, error: 'Lock timeout' };

    withScriptLock_(function() {
      var configSheet = ss.getSheetByName(SHEETS.CONFIG);
      if (!configSheet) { result = { success: false, error: 'Config sheet not found' }; return; }

      // Read old values for audit log
      var oldValues = getConfigValues(configSheet, col);

      // Clear existing values in this column (rows 3+)
      var lastRow = configSheet.getLastRow();
      if (lastRow >= 3) {
        configSheet.getRange(3, col, lastRow - 2, 1).clearContent();
      }

      // Sanitize and filter blanks
      var sanitized = [];
      for (var i = 0; i < values.length; i++) {
        var v = String(values[i]).trim();
        if (v) sanitized.push(v);
      }

      // Write new values
      if (sanitized.length > 0) {
        var writeData = sanitized.map(function(v) { return [escapeForFormula(v)]; });
        configSheet.getRange(3, col, writeData.length, 1).setValues(writeData);
      }

      // Re-sync data validations so dropdowns reflect new values
      try { setupDataValidations(); } catch (_e) { log_('Validation sync', (_e.message || _e)); }

      // Bust caches
      try { ConfigReader.refreshConfig(); } catch (_e) { log_('Cache refresh', (_e.message || _e)); }

      logAuditEvent('ADMIN_SETTINGS_LIST_CHANGE', {
        field: configColKey,
        oldCount: oldValues.length,
        newCount: sanitized.length,
        changedBy: email
      });

      result = { success: true, count: sanitized.length };
    });

    return result;
  } catch (e) {
    if (e.message && e.message.indexOf('lock') !== -1) {
      return { success: false, error: 'Another save is in progress. Please try again.' };
    }
    log_('adminSaveListValues error', e.message);
    return { success: false, error: 'Save failed: ' + e.message };
  }
}

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Finds a field definition in the schema by CONFIG_COLS key.
 * @param {string} key - CONFIG_COLS key
 * @returns {Object|null} Field definition or null
 * @private
 */
function _findFieldDef_(key) {
  for (var t = 0; t < ADMIN_SETTINGS_SCHEMA_.length; t++) {
    var fields = ADMIN_SETTINGS_SCHEMA_[t].fields;
    for (var f = 0; f < fields.length; f++) {
      if (fields[f].key === key) return fields[f];
    }
  }
  return null;
}
