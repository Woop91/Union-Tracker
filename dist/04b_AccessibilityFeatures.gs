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
        background: #1557b0;
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

/**
 * Shows a toast notification
 * @param {string} message - Message to display
 * @param {string} title - Optional title
 */

/**
 * Shows a confirmation dialog
 * @param {string} message - Confirmation message
 * @param {string} title - Dialog title
 * @return {boolean} True if user confirmed
 */

/**
 * Shows an alert dialog
 * @param {string} message - Alert message
 * @param {string} title - Dialog title
 */

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

// Theme Configuration — delegates to THEME_PRESETS (03_UIComponents.gs) for unified keys
var THEME_CONFIG = {
  THEMES: {
    'default': { name: 'Dark Slate', icon: '🌑', headerBackground: '#1e293b' },
    'union-blue': { name: 'Union Blue', icon: '💙', headerBackground: '#1e40af' },
    'forest': { name: 'Forest Green', icon: '💚', headerBackground: '#166534' },
    'midnight': { name: 'Midnight Purple', icon: '💜', headerBackground: '#581c87' },
    'crimson': { name: 'Crimson', icon: '❤️', headerBackground: '#991b1b' },
    'ocean': { name: 'Ocean Teal', icon: '🌊', headerBackground: '#115e59' }
  },
  DEFAULT_THEME: 'default'
};

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
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:Arial,sans-serif;padding:30px;text-align:center;background:linear-gradient(135deg,#ff6b6b,#ee5a24)}' +
    '.timer{font-size:72px;color:white;text-shadow:2px 2px 4px rgba(0,0,0,0.3)}' +
    '.label{color:white;font-size:18px;margin-top:10px;opacity:0.9}' +
    '.tip{background:rgba(255,255,255,0.2);padding:15px;border-radius:8px;margin-top:20px;color:white}' +
    'button{background:white;color:#ee5a24;border:none;padding:12px 24px;border-radius:25px;font-size:16px;cursor:pointer;margin-top:20px}' +
    '</style></head><body>' +
    '<div class="timer" id="time">25:00</div>' +
    '<div class="label">🍅 Focus Time Remaining</div>' +
    '<div class="tip">💡 Tip: Close this window and focus on your task.<br>A toast notification will remind you when time is up.</div>' +
    '<button onclick="google.script.host.close()">Start Focusing</button>' +
    '<script>' +
    'var seconds = 25 * 60;' +
    'var timer = setInterval(function(){' +
    '  seconds--;' +
    '  if(seconds <= 0){clearInterval(timer);document.getElementById("time").innerHTML="Done!";return;}' +
    '  var m = Math.floor(seconds/60);' +
    '  var s = seconds % 60;' +
    '  document.getElementById("time").innerHTML = m + ":" + (s < 10 ? "0" : "") + s;' +
    '},1000);' +
    '</script></body></html>'
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
    userProps.setProperty('quickCaptureNotes', notes || '');
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
  var html = HtmlService.createHtmlOutput(
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
    '</script>'
  )
  .setWidth(500)
  .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '📝 Quick Capture Notepad');
}
/**
 * Generates HTML for import dialog
 * @private
 */
function getImportDialogHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Roboto", Arial, sans-serif; padding: 20px; background: #f5f5f5; }' +
    '.container { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }' +
    'h2 { color: #1a73e8; margin-bottom: 15px; font-size: 18px; }' +
    '.section { margin-bottom: 20px; }' +
    '.section-title { font-weight: bold; color: #333; margin-bottom: 8px; font-size: 14px; }' +
    'textarea { width: 100%; height: 200px; border: 2px solid #e0e0e0; border-radius: 6px; padding: 12px; font-family: monospace; font-size: 12px; resize: vertical; }' +
    'textarea:focus { border-color: #1a73e8; outline: none; }' +
    '.format-hint { background: #e8f0fe; padding: 12px; border-radius: 6px; font-size: 12px; color: #1967d2; margin-bottom: 15px; }' +
    '.format-hint code { background: #fff; padding: 2px 6px; border-radius: 3px; }' +
    '.preview { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 12px; max-height: 150px; overflow-y: auto; font-size: 12px; display: none; }' +
    '.preview-row { padding: 4px 0; border-bottom: 1px solid #eee; }' +
    '.preview-row:last-child { border-bottom: none; }' +
    '.btn-row { display: flex; gap: 10px; margin-top: 15px; }' +
    'button { padding: 10px 20px; border-radius: 6px; font-size: 14px; cursor: pointer; border: none; font-weight: 500; }' +
    '.btn-primary { background: #1a73e8; color: white; }' +
    '.btn-primary:hover { background: #1557b0; }' +
    '.btn-secondary { background: #f1f3f4; color: #5f6368; }' +
    '.btn-secondary:hover { background: #e8eaed; }' +
    '.status { padding: 10px; border-radius: 6px; margin-top: 10px; display: none; }' +
    '.status.success { background: #e6f4ea; color: #137333; }' +
    '.status.error { background: #fce8e6; color: #c5221f; }' +
    '.count { font-weight: bold; color: #1a73e8; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>📥 Import Members from CSV</h2>' +
    '<div class="format-hint">' +
    '<strong>Required columns:</strong> First Name, Last Name<br>' +
    '<strong>Optional:</strong> Email, Phone, Job Title, Unit, Work Location, Supervisor, Is Steward, Dues Paying<br>' +
    '<strong>Format:</strong> Paste CSV with headers in first row. Use comma separator.' +
    '</div>' +
    '<div class="section">' +
    '<div class="section-title">Paste CSV Data:</div>' +
    '<textarea id="csvData" placeholder="First Name,Last Name,Email,Phone,Job Title,Unit&#10;John,Doe,john@example.com,555-1234,Clerk,Admin&#10;Jane,Smith,jane@example.com,555-5678,Manager,Operations"></textarea>' +
    '</div>' +
    '<div class="section">' +
    '<div class="section-title">Preview (<span id="rowCount" class="count">0</span> rows):</div>' +
    '<div id="preview" class="preview"></div>' +
    '</div>' +
    '<div id="status" class="status"></div>' +
    '<div class="btn-row">' +
    '<button class="btn-secondary" onclick="previewData()">👁️ Preview</button>' +
    '<button class="btn-primary" onclick="importData()">📥 Import</button>' +
    '<button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '</div></div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'function previewData() {' +
    '  var csv = document.getElementById("csvData").value.trim();' +
    '  if (!csv) { showStatus("Please paste CSV data first", "error"); return; }' +
    '  var rows = parseCSV(csv);' +
    '  if (rows.length < 2) { showStatus("Need at least a header row and one data row", "error"); return; }' +
    '  document.getElementById("rowCount").textContent = rows.length - 1;' +
    '  var previewHtml = "<div class=\\"preview-row\\"><strong>" + rows[0].map(function(c){return escapeHtml(c)}).join(" | ") + "</strong></div>";' +
    '  for (var i = 1; i < Math.min(rows.length, 6); i++) {' +
    '    previewHtml += "<div class=\\"preview-row\\">" + rows[i].map(function(c){return escapeHtml(c)}).join(" | ") + "</div>";' +
    '  }' +
    '  if (rows.length > 6) previewHtml += "<div class=\\"preview-row\\">... and " + (rows.length - 6) + " more rows</div>";' +
    '  document.getElementById("preview").innerHTML = previewHtml;' +
    '  document.getElementById("preview").style.display = "block";' +
    '  showStatus("Preview ready. Click Import to add " + (rows.length - 1) + " members.", "success");' +
    '}' +
    'function parseCSV(csv) {' +
    '  var lines = csv.split(/\\r?\\n/);' +
    '  return lines.filter(function(line) { return line.trim(); }).map(function(line) {' +
    '    var result = []; var cell = ""; var inQuotes = false;' +
    '    for (var i = 0; i < line.length; i++) {' +
    '      var c = line[i];' +
    '      if (c === "\\"") { inQuotes = !inQuotes; }' +
    '      else if (c === "," && !inQuotes) { result.push(cell.trim()); cell = ""; }' +
    '      else { cell += c; }' +
    '    }' +
    '    result.push(cell.trim());' +
    '    return result;' +
    '  });' +
    '}' +
    'function importData() {' +
    '  var csv = document.getElementById("csvData").value.trim();' +
    '  if (!csv) { showStatus("Please paste CSV data first", "error"); return; }' +
    '  showStatus("Importing...", "success");' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    if (result.success) {' +
    '      showStatus("✅ Successfully imported " + result.count + " members!", "success");' +
    '      setTimeout(function() { google.script.host.close(); }, 2000);' +
    '    } else {' +
    '      showStatus("❌ " + escapeHtml(result.error), "error");' +
    '    }' +
    '  }).withFailureHandler(function(err) {' +
    '    showStatus("❌ Error: " + escapeHtml(err.message), "error");' +
    '  }).processMemberImport(csv);' +
    '}' +
    'function showStatus(msg, type) {' +
    '  var el = document.getElementById("status");' +
    '  el.textContent = msg;' +
    '  el.className = "status " + type;' +
    '  el.style.display = "block";' +
    '}' +
    '</script></body></html>';
}

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
    for (var i = 1; i < lines.length; i++) {
      var values = parseCSVLine_(lines[i]);
      if (!values || values.length === 0) continue;

      // Build row data
      var rowData = [];
      rowData[MEMBER_COLS.FIRST_NAME - 1] = values[columnMap.firstName] || '';
      rowData[MEMBER_COLS.LAST_NAME - 1] = values[columnMap.lastName] || '';
      rowData[MEMBER_COLS.EMAIL - 1] = columnMap.email !== undefined ? values[columnMap.email] || '' : '';
      rowData[MEMBER_COLS.PHONE - 1] = columnMap.phone !== undefined ? values[columnMap.phone] || '' : '';
      rowData[MEMBER_COLS.JOB_TITLE - 1] = columnMap.jobTitle !== undefined ? values[columnMap.jobTitle] || '' : '';
      rowData[MEMBER_COLS.UNIT - 1] = columnMap.unit !== undefined ? values[columnMap.unit] || '' : '';
      rowData[MEMBER_COLS.WORK_LOCATION - 1] = columnMap.workLocation !== undefined ? values[columnMap.workLocation] || '' : '';
      rowData[MEMBER_COLS.SUPERVISOR - 1] = columnMap.supervisor !== undefined ? values[columnMap.supervisor] || '' : '';
      rowData[MEMBER_COLS.IS_STEWARD - 1] = columnMap.isSteward !== undefined ? (values[columnMap.isSteward] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'No';
      rowData[MEMBER_COLS.DUES_PAYING - 1] = columnMap.duesPaying !== undefined ? (values[columnMap.duesPaying] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'Yes';

      // Fill empty cells
      while (rowData.length < sheet.getLastColumn()) {
        rowData.push('');
      }

      // Append row
      sheet.appendRow(rowData);
      importedCount++;
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
 * @private
 */
function parseCSVLine_(line) {
  var result = [];
  var cell = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (c === '"') {
      inQuotes = !inQuotes;
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

/**
 * Displays a toast notification reminding the user to take a break.
 * @returns {void}
 */
function showBreakReminder() {
  SpreadsheetApp.getActiveSpreadsheet().toast('💆 Time for a break! Stretch and rest your eyes.', 'Break Reminder', 10);
}

// ==================== THEME MANAGEMENT ====================

/**
 * Resets to default theme using theme system
 * NOTE: Renamed to avoid duplicate. Use resetToDefaultTheme() for hard reset to defaults.
 * This version uses the theme system; resetToDefaultTheme() clears all styling.
 */

// ==================== COMFORT VIEW CONTROL PANEL ====================

/**
 * Opens the Comfort View control panel dialog with accessibility toggles.
 * @returns {void}
 */
function showComfortViewControlPanel() {
  var settings = getComfortViewSettings();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8;border-bottom:3px solid #1a73e8;padding-bottom:10px}.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px;border-left:4px solid #1a73e8}.row{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;margin:5px}button:hover{background:#1557b0}button.sec{background:#6c757d}</style></head><body><div class="container"><h2>♿ Comfort View Panel</h2><div class="section"><div class="row"><span>Zebra Stripes</span><button onclick="google.script.run.toggleZebraStripes();setTimeout(function(){location.reload()},1000)">' + (settings.zebraStripes ? '✅ On' : 'Off') + '</button></div><div class="row"><span>Gridlines</span><button onclick="google.script.run.toggleGridlinesComfortView();setTimeout(function(){location.reload()},1000)">' + (settings.gridlines ? '✅ Visible' : 'Hidden') + '</button></div><div class="row"><span>Focus Mode</span><button onclick="google.script.run.activateFocusMode();google.script.host.close()">🎯 Activate</button></div></div><div class="section"><div class="row"><span>Quick Capture</span><button onclick="google.script.run.showQuickCaptureNotepad()">📝 Open</button></div><div class="row"><span>Pomodoro Timer</span><button onclick="google.script.run.startPomodoroTimer();google.script.host.close()">⏱️ Start</button></div></div><button class="sec" onclick="google.script.run.resetComfortViewSettings();google.script.host.close()">🔄 Reset</button><button class="sec" onclick="google.script.host.close()">Close</button></div></body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '♿ Comfort View Panel');
}

/**
 * Opens the Theme Manager dialog for selecting and applying sheet themes.
 * @returns {void}
 */
function showThemeManager() {
  var current = getCurrentTheme();
  var themeCards = Object.keys(THEME_CONFIG.THEMES).map(function(key) {
    var t = THEME_CONFIG.THEMES[key];
    return '<div style="background:#f8f9fa;padding:15px;border-radius:8px;cursor:pointer;border:3px solid ' + (current === key ? '#1a73e8' : 'transparent') + '" onclick="select(\'' + key + '\')">' +
      '<div style="font-size:32px;text-align:center">' + t.icon + '</div>' +
      '<div style="text-align:center;font-weight:bold">' + t.name + '</div>' +
      '<div style="height:20px;background:' + t.headerBackground + ';border-radius:4px;margin-top:10px"></div></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:5px}button.sec{background:#6c757d}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:15px;margin:20px 0}</style></head><body><div class="container"><h2>🎨 Theme Manager</h2><div class="grid">' + themeCards + '</div><button onclick="apply()">✅ Apply Theme</button><button class="sec" onclick="google.script.host.close()">Close</button></div><script>var sel="' + current + '";function select(k){sel=k;document.querySelectorAll(".grid>div").forEach(function(d){d.style.border="3px solid transparent"});event.currentTarget.style.border="3px solid #1a73e8"}function apply(){google.script.run.withSuccessHandler(function(){alert("Theme applied!");google.script.host.close()}).applyThemePreset(sel)}</script></body></html>'
  ).setWidth(450).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🎨 Theme Manager');
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

/**
 * Shows a smart dashboard that automatically detects the device type
 * and displays the appropriate interface (mobile or desktop)
 */
function showSmartDashboard() {
  var html = HtmlService.createHtmlOutput(getSmartDashboardHtml())
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Dashboard Pend');
}

/**
 * Returns the HTML for the smart dashboard with device detection
 */
function getSmartDashboardHtml() {
  var stats = getMobileDashboardStats();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    getMobileOptimizedHead() +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Responsive container
    '.container{padding:15px;max-width:1200px;margin:0 auto}' +

    // Header - responsive
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +
    '.device-badge{display:inline-block;padding:4px 12px;background:rgba(255,255,255,0.2);border-radius:20px;font-size:11px;margin-top:8px}' +

    // Stats grid - responsive
    '.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,6vw,36px);font-weight:bold;color:#1a73e8}' +
    '.stat-label{font-size:clamp(11px,2.5vw,13px);color:#666;text-transform:uppercase;margin-top:5px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Action buttons - responsive grid
    '.actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}' +
    '.action-btn{background:white;border:none;padding:16px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;' +
    'min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';transition:all 0.2s}' +
    '.action-btn:hover{background:#e8f0fe;transform:translateX(4px)}' +
    '.action-btn:active{transform:scale(0.98)}' +
    '.action-icon{font-size:24px;width:44px;height:44px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px;flex-shrink:0}' +
    '.action-label{font-weight:500}' +
    '.action-desc{font-size:12px;color:#666;margin-top:2px}' +

    // FAB (Floating Action Button)
    '.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;' +
    'border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer;z-index:1000}' +
    '.fab:hover{background:#1557b0}' +

    // Desktop-only elements
    '.desktop-only{display:none}' +

    // Mobile-specific adjustments
    '@media (max-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:1fr}' +
    '  .container{padding:10px}' +
    '  .header{padding:15px}' +
    '}' +

    // Tablet adjustments
    '@media (min-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px) and (max-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '}' +

    // Desktop view
    '@media (min-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(4,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '  .desktop-only{display:block}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header with dynamic device badge
    '<div class="header">' +
    '<h1>📋 Dashboard Pend</h1>' +
    '<div class="subtitle">Pending Actions & Quick Overview</div>' +
    '<div class="device-badge" id="deviceBadge">Detecting device...</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section
    '<div class="stats">' +
    '<div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div>' +
    '</div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<button class="action-btn" onclick="google.script.run.showMobileGrievanceList()">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">View Grievances</div><div class="action-desc">Browse and filter all grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()">' +
    '<div class="action-icon">👤</div>' +
    '<div><div class="action-label">My Cases</div><div class="action-desc">View your assigned grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showQuickActionsMenu()">' +
    '<div class="action-icon">⚡</div>' +
    '<div><div class="action-label">Row Actions</div><div class="action-desc">Quick actions for selected row</div></div>' +
    '</button>' +

    '</div>' +

    // Desktop-only additional info
    '<div class="desktop-only">' +
    '<div class="section-title">ℹ️ Dashboard Info</div>' +
    '<p style="color:#666;font-size:14px;padding:15px;background:white;border-radius:8px;">' +
    'This responsive dashboard automatically adjusts to your screen size. ' +
    'On mobile devices, you\'ll see a touch-optimized interface with larger buttons. ' +
    'Use the menu items above to manage grievances and member information.' +
    '</p>' +
    '</div>' +

    '</div>' +

    // FAB for refresh
    '<button class="fab" onclick="location.reload()" title="Refresh">🔄</button>' +

    // Device detection script
    '<script>' +
    'function detectDevice(){' +
    '  var w=window.innerWidth;' +
    '  var badge=document.getElementById("deviceBadge");' +
    '  var isTouchDevice="ontouchstart" in window||navigator.maxTouchPoints>0;' +
    '  var userAgent=navigator.userAgent.toLowerCase();' +
    '  var isMobileUA=/android|webos|iphone|ipad|ipod|blackberry|iemobile|opera mini/i.test(userAgent);' +
    '  ' +
    '  if(w<' + MOBILE_CONFIG.MOBILE_BREAKPOINT + '||isMobileUA){' +
    '    badge.textContent="📱 Mobile View";' +
    '    badge.style.background="rgba(76,175,80,0.3)";' +
    '  }else if(w<' + MOBILE_CONFIG.TABLET_BREAKPOINT + '){' +
    '    badge.textContent="📱 Tablet View";' +
    '    badge.style.background="rgba(255,152,0,0.3)";' +
    '  }else{' +
    '    badge.textContent="🖥️ Desktop View";' +
    '    badge.style.background="rgba(33,150,243,0.3)";' +
    '  }' +
    '  ' +
    '  if(isTouchDevice){' +
    '    document.body.classList.add("touch-device");' +
    '  }' +
    '}' +
    'detectDevice();' +
    'window.addEventListener("resize",detectDevice);' +
    '</script>' +

    '</body></html>';
}

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

