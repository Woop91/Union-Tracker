/**
 * ============================================================================
 * 04b_AccessibilityFeatures.gs - Shared Styles & Productivity Tools
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Two responsibilities:
 *   1. Shared CSS — getCommonStyles() returns a CSS string used by ALL dialog
 *      and sidebar HTML (button, form, input, and responsive styles).
 *   2. Productivity & accessibility features — Pomodoro timer, Quick Capture
 *      notepad, Comfort View control panel, break reminders, theme manager, smart
 *      dashboard, and Member Import dialog (CSV parsing + column mapping).
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Centralizes CSS so all dialogs have consistent styling via UI_THEME
 *   constants from 01_Core.gs. The productivity tools are grouped here
 *   because they share the same CSS foundation and target accessibility /
 *   quality-of-life needs rather than core union workflows.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   All dialogs and sidebars appear unstyled (raw HTML). Pomodoro timer,
 *   Quick Capture, Comfort View tools, and Member Import stop working. Core
 *   grievance and member operations are unaffected.
 *
 * DEPENDENCIES:
 *   - Depends on: 01_Core.gs (UI_THEME, getMobileOptimizedHead)
 *   - Used by:    Every file that generates dialog/sidebar HTML
 *                 (04a, 04d, 04e, 08b, 08c, 09_, 11_, 13_, 14_)
 *
 * @fileoverview Common CSS styles, productivity tools, and accessibility features
 * @requires 01_Core.gs
 */

// ============================================================================
// COMMON UI UTILITIES
// ============================================================================

/**
 * Gets common CSS styles used across all dialogs
 * @return {string} CSS styles
 */
function getCommonStyles() {
  return `
    <style>
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      body {
        font-family: 'Google Sans', Roboto, Arial, sans-serif;
        font-size: 14px;
        color: ${UI_THEME.TEXT_PRIMARY};
        line-height: 1.5;
      }
      .btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 10px 20px;
        border: none;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: background 0.2s, transform 0.1s;
      }
      .btn:hover {
        transform: translateY(-1px);
      }
      .btn:active {
        transform: translateY(0);
      }
      .btn-primary {
        background: ${UI_THEME.PRIMARY_COLOR};
        color: white;
      }
      .btn-primary:hover {
        background: ${SHEET_COLORS.DIALOG_ACCENT_DARK};
      }
      .btn-secondary {
        background: #f1f3f4;
        color: ${UI_THEME.TEXT_PRIMARY};
      }
      .btn-secondary:hover {
        background: #e8eaed;
      }
      .btn-danger {
        background: ${UI_THEME.DANGER_COLOR};
        color: white;
      }
      .btn-danger:hover {
        background: #c5221f;
      }
      .btn-success {
        background: ${UI_THEME.SECONDARY_COLOR};
        color: white;
      }
      .form-group {
        margin-bottom: 15px;
      }
      .form-label {
        display: block;
        font-weight: 500;
        margin-bottom: 5px;
        color: ${UI_THEME.TEXT_PRIMARY};
      }
      .form-input {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        transition: border-color 0.2s;
      }
      .form-input:focus {
        outline: none;
        border-color: ${UI_THEME.PRIMARY_COLOR};
      }
      .form-select {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        background: white;
      }
      .form-textarea {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        min-height: 100px;
        resize: vertical;
      }
      .alert {
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 15px;
      }
      .alert-info {
        background: #e8f0fe;
        color: #1967d2;
      }
      .alert-warning {
        background: #fef7e0;
        color: #ea8600;
      }
      .alert-error {
        background: #fce8e6;
        color: #c5221f;
      }
      .alert-success {
        background: #e6f4ea;
        color: #137333;
      }

      /* Mobile-responsive enhancements for common styles */
      @media (max-width: ${MOBILE_CONFIG.MOBILE_BREAKPOINT}px) {
        .btn {
          min-height: ${MOBILE_CONFIG.TOUCH_TARGET_SIZE};
          padding: 12px 16px;
          font-size: clamp(13px, 3.5vw, 15px);
        }
        .form-input, .form-select, .form-textarea {
          min-height: ${MOBILE_CONFIG.TOUCH_TARGET_SIZE};
          font-size: 16px; /* Prevents iOS zoom */
        }
        .alert {
          padding: 10px 12px;
          font-size: clamp(12px, 3vw, 14px);
        }
      }
    </style>
    ${getMobileOptimizedHead()}
  `;
}

// navigateToSheet() - REMOVED DUPLICATE - see line 565 for main definition

/**
 * Navigates to a specific record (row) in the appropriate sheet
 * @param {string} id - Record ID
 * @param {string} type - Record type ('member' or 'grievance')
 */
/**
 * ============================================================================
 * COMFORT VIEW ACCESSIBILITY & THEMING
 * ============================================================================
 * Features for neurodivergent users + theme customization
 */

// Theme Configuration — THEME_PRESETS in 03_UIComponents.gs is the single source of truth.
// showThemeManager() below delegates to showThemePresetPicker() for unified UI.

// ==================== COMFORT VIEW SETTINGS ====================

// ==================== VISUAL HELPERS ====================

// ==================== FOCUS MODE ====================

/**
 * Activate Focus Mode - distraction-free view for focused work
 *
 * WHAT IT DOES:
 * - Hides all sheets except the one you're currently viewing
 * - Removes gridlines to reduce visual clutter
 * - Creates a clean, focused work environment
 *
 * HOW TO EXIT:
 * - Use menu: Comfort View > Exit Focus Mode
 * - Or run: deactivateFocusMode()
 *
 * BEST FOR:
 * - Deep work on a single task
 * - Reducing cognitive load
 * - Preventing tab-switching distractions
 */

// ==================== POMODORO TIMER ====================

/**
 * Starts a simple Pomodoro timer using Google Sheets toast notifications
 * Shows a 25-minute work session reminder
 */
function startPomodoroTimer() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show starting message
  ss.toast('🍅 Pomodoro started! Focus for 25 minutes.\n\nYou\'ll get a notification when the session ends.', 'Pomodoro Timer', 10);

  // Store the start time
  PropertiesService.getUserProperties().setProperty('pomodoroStart', new Date().toISOString());

  // For immediate feedback, show a modal with timer info
  var html = HtmlService.createHtmlOutput(
    `<!DOCTYPE html><html><head><base target="_top">${getMobileOptimizedHead()}<style>
body{font-family:Arial,sans-serif;padding:30px;text-align:center;background:linear-gradient(135deg,#ff6b6b,#ee5a24)}
.timer{font-size:72px;color:white;text-shadow:2px 2px 4px rgba(0,0,0,0.3)}
.label{color:white;font-size:18px;margin-top:10px;opacity:0.9}
.tip{background:rgba(255,255,255,0.2);padding:15px;border-radius:8px;margin-top:20px;color:white}
button{background:white;color:#ee5a24;border:none;padding:12px 24px;border-radius:25px;font-size:16px;cursor:pointer;margin-top:20px}
</style></head><body>
<div class="timer" id="time">25:00</div>
<div class="label">🍅 Focus Time Remaining</div>
<div class="tip">💡 Tip: Close this window and focus on your task.<br>A toast notification will remind you when time is up.</div>
<button onclick="google.script.host.close()">Start Focusing</button>
<script>
var seconds = 25 * 60;
var timer = setInterval(function(){
  seconds--;
  if(seconds <= 0){clearInterval(timer);document.getElementById("time").innerHTML="Done!";return;}
  var m = Math.floor(seconds/60);
  var s = seconds % 60;
  document.getElementById("time").innerHTML = m + ":" + (s < 10 ? "0" : "") + s;
},1000);
</script></body></html>`
  ).setWidth(350).setHeight(350);

  ui.showModelessDialog(html, '🍅 Pomodoro Timer');
}
// ==================== QUICK CAPTURE NOTEPAD ====================

/**
 * Quick Capture Notepad - Fast note-taking without losing focus
 * Notes are stored per-user in Script Properties
 */

/**
 * Gets the current user's quick capture notes
 * @returns {string} The saved notes or empty string
 */
function getQuickCaptureNotes() {
  var userProps = PropertiesService.getUserProperties();
  return userProps.getProperty('quickCaptureNotes') || '';
}

/**
 * Saves quick capture notes for the current user
 * @param {string} notes - The notes to save
 * @returns {Object} Result object with success status
 */
function saveQuickCaptureNotes(notes) {
  try {
    var userProps = PropertiesService.getUserProperties();
    var sanitized = typeof escapeForFormula === 'function' ? escapeForFormula(notes || '') : (notes || '');
    userProps.setProperty('quickCaptureNotes', sanitized);
    userProps.setProperty('quickCaptureLastSaved', new Date().toISOString());
    return { success: true, message: 'Notes saved' };
  } catch (e) {
    return errorResponse(e.message);
  }
}

/**
 * Clears the quick capture notes
 * @returns {Object} Result object with success status
 */
function clearQuickCaptureNotes() {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('quickCaptureNotes');
    userProps.deleteProperty('quickCaptureLastSaved');
    return { success: true, message: 'Notes cleared' };
  } catch (e) {
    return errorResponse(e.message);
  }
}
/**
 * Shows the Quick Capture Notepad dialog
 * A fast notepad for capturing thoughts without losing focus
 */
function showQuickCaptureNotepad() {
  showDialog_(
    '<html><head>' + getMobileOptimizedHead() + '</head><body><style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", Roboto, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; padding: 16px; color: #F8FAFC; }' +
    '.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }' +
    'h3 { font-size: 18px; display: flex; align-items: center; gap: 8px; }' +
    '.meta { font-size: 12px; color: #64748B; }' +
    'textarea { width: 100%; height: 280px; padding: 12px; border: 2px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-size: 14px; font-family: inherit; resize: none; outline: none; }' +
    'textarea:focus { border-color: #7C3AED; }' +
    'textarea::placeholder { color: #64748B; }' +
    '.btn-row { display: flex; gap: 8px; margin-top: 12px; }' +
    'button { flex: 1; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; }' +
    '.btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }' +
    '.btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }' +
    '.btn-secondary { background: #334155; color: #F8FAFC; }' +
    '.btn-danger { background: #DC2626; color: white; }' +
    '.status { margin-top: 8px; font-size: 12px; color: #10B981; text-align: center; opacity: 0; transition: opacity 0.3s; }' +
    '.status.show { opacity: 1; }' +
    '</style>' +
    '<div class="header">' +
    '  <h3>📝 Quick Capture</h3>' +
    '  <span class="meta" id="meta"></span>' +
    '</div>' +
    '<textarea id="notes" placeholder="Capture your thoughts quickly...\\n\\nUse this notepad to jot down ideas, reminders, or notes without losing focus on your current task.\\n\\nYour notes are auto-saved when you click Save."></textarea>' +
    '<div class="btn-row">' +
    '  <button class="btn-primary" onclick="saveNotes()">💾 Save</button>' +
    '  <button class="btn-secondary" onclick="copyNotes()">📋 Copy</button>' +
    '  <button class="btn-danger" onclick="clearNotes()">🗑️ Clear</button>' +
    '  <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div class="status" id="status"></div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'var notesEl = document.getElementById("notes");' +
    'var statusEl = document.getElementById("status");' +
    'var metaEl = document.getElementById("meta");' +
    '' +
    'google.script.run.withSuccessHandler(function(notes) {' +
    '  notesEl.value = notes || "";' +
    '  updateMeta();' +
    '}).getQuickCaptureNotes();' +
    '' +
    'notesEl.addEventListener("input", updateMeta);' +
    '' +
    'function updateMeta() {' +
    '  var text = notesEl.value;' +
    '  var words = text.trim() ? text.trim().split(/\\s+/).length : 0;' +
    '  metaEl.textContent = words + " words, " + text.length + " chars";' +
    '}' +
    '' +
    'function showStatus(msg, isError) {' +
    '  statusEl.textContent = msg;' +
    '  statusEl.style.color = isError ? "#EF4444" : "#10B981";' +
    '  statusEl.classList.add("show");' +
    '  setTimeout(function() { statusEl.classList.remove("show"); }, 2000);' +
    '}' +
    '' +
    'function saveNotes() {' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    showStatus(result.success ? "✅ Notes saved!" : "❌ " + escapeHtml(result.message), !result.success);' +
    '  }).saveQuickCaptureNotes(notesEl.value);' +
    '}' +
    '' +
    'function copyNotes() {' +
    '  navigator.clipboard.writeText(notesEl.value).then(function() {' +
    '    showStatus("📋 Copied to clipboard!");' +
    '  }).catch(function() {' +
    '    showStatus("❌ Failed to copy", true);' +
    '  });' +
    '}' +
    '' +
    'function clearNotes() {' +
    '  if (confirm("Clear all notes? This cannot be undone.")) {' +
    '    notesEl.value = "";' +
    '    google.script.run.withSuccessHandler(function(result) {' +
    '      showStatus(result.success ? "🗑️ Notes cleared" : "❌ " + escapeHtml(result.message), !result.success);' +
    '      updateMeta();' +
    '    }).clearQuickCaptureNotes();' +
    '  }' +
    '}' +
    '' +
    'notesEl.addEventListener("keydown", function(e) {' +
    '  if ((e.ctrlKey || e.metaKey) && e.key === "s") {' +
    '    e.preventDefault();' +
    '    saveNotes();' +
    '  }' +
    '});' +
    '</script></body></html>',
    '📝 Quick Capture Notepad', 500, 450);
}
// Dead code removed: getImportDialogHtml_() — zero callers in src

/**
 * Processes member import from CSV data
 * @param {string} csvData - Raw CSV data
 * @returns {Object} Result with success status and count/error
 */
function processMemberImport(csvData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      return errorResponse('Member Directory sheet not found');
    }

    // Parse CSV
    var lines = csvData.split(/\r?\n/).filter(function(l) { return l.trim(); });
    if (lines.length < 2) {
      return errorResponse('Need at least header and one data row');
    }

    // Parse header and map columns
    var headers = parseCSVLine_(lines[0]);
    var columnMap = mapImportColumns_(headers);

    if (!columnMap.firstName || !columnMap.lastName) {
      return errorResponse('Required columns missing: First Name, Last Name');
    }

    // Process data rows
    var importedCount = 0;
    var allRows = [];
    var numCols = sheet.getLastColumn();
    var _eff = typeof escapeForFormula === 'function' ? escapeForFormula : function(v) { return v; };
    for (var i = 1; i < lines.length; i++) {
      var values = parseCSVLine_(lines[i]);
      if (!values || values.length === 0) continue;

      // Build row data — escapeForFormula on all string inputs to prevent formula injection
      var rowData = [];
      setCol_(rowData, MEMBER_COLS.FIRST_NAME, _eff(values[columnMap.firstName] || ''));
      setCol_(rowData, MEMBER_COLS.LAST_NAME, _eff(values[columnMap.lastName] || ''));
      setCol_(rowData, MEMBER_COLS.EMAIL, columnMap.email !== undefined ? _eff(values[columnMap.email] || '') : '');
      setCol_(rowData, MEMBER_COLS.PHONE, columnMap.phone !== undefined ? _eff(values[columnMap.phone] || '') : '');
      setCol_(rowData, MEMBER_COLS.JOB_TITLE, columnMap.jobTitle !== undefined ? _eff(values[columnMap.jobTitle] || '') : '');
      setCol_(rowData, MEMBER_COLS.UNIT, columnMap.unit !== undefined ? _eff(values[columnMap.unit] || '') : '');
      setCol_(rowData, MEMBER_COLS.WORK_LOCATION, columnMap.workLocation !== undefined ? _eff(values[columnMap.workLocation] || '') : '');
      setCol_(rowData, MEMBER_COLS.SUPERVISOR, columnMap.supervisor !== undefined ? _eff(values[columnMap.supervisor] || '') : '');
      setCol_(rowData, MEMBER_COLS.IS_STEWARD, columnMap.isSteward !== undefined ? (values[columnMap.isSteward] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'No');
      setCol_(rowData, MEMBER_COLS.DUES_STATUS, columnMap.duesPaying !== undefined ? (values[columnMap.duesPaying] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'Yes');

      // Fill empty cells
      while (rowData.length < numCols) {
        rowData.push('');
      }

      allRows.push(rowData);
      importedCount++;
    }

    // Batch append all rows at once
    if (allRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
    }

    // Generate Member IDs for imported rows
    if (typeof generateMissingMemberIDs === 'function') {
      generateMissingMemberIDs();
    }

    return { success: true, count: importedCount };

  } catch (e) {
    return errorResponse(e.message);
  }
}

/**
 * Parses a single CSV line handling quoted values
 * KNOWN LIMITATION: This CSV parser does not handle newlines within quoted fields.
 * Quoted fields containing \n will be split across lines, producing corrupt data.
 * For files with newlines in fields, use a pre-processor or convert to XLSX format.
 * @private
 */
function parseCSVLine_(line) {
  var result = [];
  var cell = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (c === '"') {
      if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"') {
        cell += '"';
        i++; // skip the second quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      result.push(cell.trim());
      cell = '';
    } else {
      cell += c;
    }
  }
  result.push(cell.trim());
  return result;
}

/**
 * Maps CSV headers to member column indices
 * @private
 */
function mapImportColumns_(headers) {
  var map = {};
  var headerLower = headers.map(function(h) { return (h || '').toLowerCase().replace(/[^a-z]/g, ''); });

  for (var i = 0; i < headerLower.length; i++) {
    var h = headerLower[i];
    if (h === 'firstname' || h === 'first') map.firstName = i;
    else if (h === 'lastname' || h === 'last') map.lastName = i;
    else if (h === 'email' || h === 'emailaddress') map.email = i;
    else if (h === 'phone' || h === 'phonenumber') map.phone = i;
    else if (h === 'jobtitle' || h === 'title' || h === 'position') map.jobTitle = i;
    else if (h === 'unit' || h === 'department' || h === 'dept') map.unit = i;
    else if (h === 'worklocation' || h === 'location' || h === 'worksite') map.workLocation = i;
    else if (h === 'supervisor' || h === 'manager') map.supervisor = i;
    else if (h === 'issteward' || h === 'steward') map.isSteward = i;
    else if (h === 'duespaying' || h === 'dues') map.duesPaying = i;
  }

  return map;
}

// Dead code removed: showBreakReminder() — zero callers in src

// ==================== THEME MANAGEMENT ====================

/**
 * Resets to default theme using theme system
 * NOTE: Renamed to avoid duplicate. Use resetToDefaultTheme() for hard reset to defaults.
 * This version uses the theme system; resetToDefaultTheme() clears all styling.
 */

// ==================== COMFORT VIEW CONTROL PANEL ====================

// Dead code removed: showComfortViewControlPanel() — zero callers in src

/**
 * Opens the Theme Manager dialog for selecting and applying sheet themes.
 * @returns {void}
 */
/**
 * Theme Manager — delegates to the unified showThemePresetPicker() in 03_UIComponents.gs.
 * Kept as a named function because Comfort View menu references it by name.
 */
function showThemeManager() {
  showThemePresetPicker();
}

// ==================== SETUP DEFAULTS ====================

/**
 * Setup Comfort View defaults with options dialog
 * User can choose which settings to apply and settings can be undone
 */

/**
 * Apply Comfort View defaults with selected options
 * @param {Object} options - Selected options
 */

/**
 * Undo Comfort View defaults - restore original settings
 */
/**
 * ============================================================================
 * MOBILE INTERFACE & QUICK ACTIONS
 * ============================================================================
 * Mobile-optimized views and context-aware quick actions
 * Includes automatic device detection for responsive experience
 */

// ==================== DEVICE DETECTION ====================
// Note: MOBILE_CONFIG is now defined in 01_Core.gs as a shared constant

// Dead code removed: showSmartDashboard(), getSmartDashboardHtml() — zero callers in src

/**
 * Check if the current context appears to be mobile
 * Note: This is a server-side heuristic based on available info
 * Real detection happens client-side in the HTML
 */

// ==================== MOBILE DASHBOARD ====================

// ==================== QUICK ACTIONS ====================

/**
 * Show context-aware Quick Actions menu
 *
 * HOW IT WORKS:
 * Quick Actions provides contextual shortcuts based on your current selection.
 *
 * AVAILABLE ON:
 * - Member Directory: Start new grievance, send email, view history, copy ID
 * - Grievance Log: Sync to calendar, setup folder, update status, copy ID
 *
 * HOW TO USE:
 * 1. Navigate to Member Directory or Grievance Log
 * 2. Click on any data row (not the header)
 * 3. Run Quick Actions from the menu
 * 4. A popup will show relevant actions for that row
 */

