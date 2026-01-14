/**
 * ============================================================================
 * COMFORT VIEW ACCESSIBILITY & THEMING
 * ============================================================================
 * Features for neurodivergent users + theme customization
 */

// Comfort View Configuration
var COMFORT_VIEW_CONFIG = {
  FOCUS_MODE_COLORS: { background: '#f5f5f5', header: '#4a4a4a', accent: '#6b9bd1' },
  HIGH_CONTRAST: { background: '#ffffff', header: '#000000', accent: '#0066cc' },
  PASTEL: { background: '#fef9e7', header: '#85929e', accent: '#7fb3d5' }
};

// Theme Configuration
var THEME_CONFIG = {
  THEMES: {
    LIGHT: { name: 'Light', icon: '☀️', background: '#ffffff', headerBackground: '#1a73e8', headerText: '#ffffff', evenRow: '#f8f9fa', oddRow: '#ffffff', text: '#202124', accent: '#1a73e8' },
    DARK: { name: 'Dark', icon: '🌙', background: '#202124', headerBackground: '#35363a', headerText: '#e8eaed', evenRow: '#292a2d', oddRow: '#202124', text: '#e8eaed', accent: '#8ab4f8' },
    PURPLE: { name: '509 Purple', icon: '💜', background: '#ffffff', headerBackground: '#5B4B9E', headerText: '#ffffff', evenRow: '#E8E3F3', oddRow: '#ffffff', text: '#1F2937', accent: '#6B5CA5' },
    GREEN: { name: 'Union Green', icon: '💚', background: '#ffffff', headerBackground: '#059669', headerText: '#ffffff', evenRow: '#D1FAE5', oddRow: '#ffffff', text: '#1F2937', accent: '#10B981' }
  },
  DEFAULT_THEME: 'LIGHT'
};

// ==================== COMFORT VIEW SETTINGS ====================

function getADHDSettings() {
  var props = PropertiesService.getUserProperties();
  var settingsJSON = props.getProperty('adhdSettings');
  if (settingsJSON) return JSON.parse(settingsJSON);
  return { theme: 'default', fontSize: 10, zebraStripes: false, gridlines: true, reducedMotion: false, breakInterval: 0 };
}

function saveADHDSettings(settings) {
  var props = PropertiesService.getUserProperties();
  var current = getADHDSettings();
  var newSettings = Object.assign({}, current, settings);
  props.setProperty('adhdSettings', JSON.stringify(newSettings));
  applyADHDSettings(newSettings);
}

function applyADHDSettings(settings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (settings.fontSize) sheet.getDataRange().setFontSize(parseInt(settings.fontSize));
    if (settings.zebraStripes) applyZebraStripes(sheet);
    if (settings.gridlines !== undefined) sheet.setHiddenGridlines(!settings.gridlines);
  });
}

function resetADHDSettings() {
  PropertiesService.getUserProperties().deleteProperty('adhdSettings');
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Comfort View settings reset', 'Settings', 3);
}

// ==================== VISUAL HELPERS ====================

function hideAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name !== SHEETS.CONFIG && name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) {
      sheet.setHiddenGridlines(true);
    }
  });
  SpreadsheetApp.getUi().alert('✅ Gridlines hidden on dashboards!');
}

function showAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) { sheet.showGridlines(); });
  SpreadsheetApp.getUi().alert('✅ Gridlines shown on all sheets.');
}

function toggleGridlinesADHD() {
  var settings = getADHDSettings();
  settings.gridlines = !settings.gridlines;
  saveADHDSettings(settings);
}

function applyZebraStripes(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var bandings = range.getBandings();
  if (bandings.length > 0) bandings[0].remove();
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function removeZebraStripes(sheet) {
  sheet.getBandings().forEach(function(b) { b.remove(); });
}

function toggleZebraStripes() {
  var settings = getADHDSettings();
  settings.zebraStripes = !settings.zebraStripes;
  saveADHDSettings(settings);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    if (settings.zebraStripes) applyZebraStripes(sheet);
    else removeZebraStripes(sheet);
  });
  ss.toast(settings.zebraStripes ? '✅ Zebra stripes enabled' : '🔕 Zebra stripes disabled', 'Visual', 3);
}

function toggleReducedMotion() {
  var settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
}

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
function activateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var active = ss.getActiveSheet();

  // Count visible sheets to warn user
  var visibleSheets = ss.getSheets().filter(function(s) { return !s.isSheetHidden(); });

  if (visibleSheets.length === 1) {
    ui.alert('🎯 Focus Mode',
      'Focus mode is already active (only one sheet visible).\n\n' +
      'To exit focus mode, use:\n' +
      'Comfort View Menu > Exit Focus Mode',
      ui.ButtonSet.OK);
    return;
  }

  // Hide all sheets except active
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName() !== active.getName()) sheet.hideSheet();
  });
  active.setHiddenGridlines(true);

  ui.alert('🎯 Focus Mode Activated',
    'You are now in Focus Mode on: "' + active.getName() + '"\n\n' +
    'WHAT THIS DOES:\n' +
    '• Hides all other sheets to reduce distractions\n' +
    '• Removes gridlines for a cleaner view\n\n' +
    'TO EXIT:\n' +
    '• Use menu: ♿ Comfort View > Exit Focus Mode\n' +
    '• Or run: deactivateFocusMode()\n\n' +
    '💡 Tip: Focus mode helps with deep work and reduces cognitive load.',
    ui.ButtonSet.OK);
}

function deactivateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    if (sheet.isSheetHidden()) sheet.showSheet();
  });
  var settings = getADHDSettings();
  ss.getActiveSheet().setHiddenGridlines(!settings.gridlines);
  ss.toast('✅ Focus mode deactivated', 'Focus Mode', 3);
}

// ==================== QUICK CAPTURE & TIMER ====================

function getQuickCaptureNotes() {
  return PropertiesService.getUserProperties().getProperty('quickCaptureNotes') || '';
}

function saveQuickCaptureNotes(notes) {
  PropertiesService.getUserProperties().setProperty('quickCaptureNotes', notes);
}

function showQuickCaptureNotepad() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}textarea{width:100%;height:300px;padding:10px;border:2px solid #ddd;border-radius:4px}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:10px 5px 0 0}</style></head><body><h3>📝 Quick Capture</h3><textarea id="notes" placeholder="Type your thoughts..."></textarea><br><button onclick="save()">💾 Save</button><button onclick="google.script.host.close()">Close</button><script>google.script.run.withSuccessHandler(function(n){document.getElementById("notes").value=n||""}).getQuickCaptureNotes();function save(){google.script.run.saveQuickCaptureNotes(document.getElementById("notes").value);alert("Saved!")}</script></body></html>'
  ).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Quick Capture');
}

function startPomodoroTimer() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:40px;text-align:center;background:#1a73e8;color:white}.timer{font-size:72px;font-weight:bold;margin:40px 0;font-family:monospace}button{background:white;color:#1a73e8;border:none;padding:15px 30px;font-size:16px;border-radius:8px;cursor:pointer;margin:10px}</style></head><body><h2>🍅 Pomodoro Timer</h2><div id="status">Focus Session</div><div class="timer" id="timer">25:00</div><button onclick="toggle()">▶️ Start</button><button onclick="google.script.host.close()">Close</button><script>var left=25*60,running=false,iv;function toggle(){if(running){clearInterval(iv);running=false}else{running=true;iv=setInterval(function(){if(left>0){left--;var m=Math.floor(left/60),s=left%60;document.getElementById("timer").textContent=(m<10?"0":"")+m+":"+(s<10?"0":"")+s}else{clearInterval(iv);alert("Session complete!")}},1000)}}</script></body></html>'
  ).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModelessDialog(html, '🍅 Pomodoro Timer');
}

function setBreakReminders(minutes) {
  var settings = getADHDSettings();
  settings.breakInterval = minutes;
  saveADHDSettings(settings);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'showBreakReminder') ScriptApp.deleteTrigger(t);
  });
  if (minutes > 0) {
    ScriptApp.newTrigger('showBreakReminder').timeBased().everyMinutes(minutes).create();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Break reminders: every ' + minutes + ' min', 'Comfort View', 3);
  }
}

function showBreakReminder() {
  SpreadsheetApp.getActiveSpreadsheet().toast('💆 Time for a break! Stretch and rest your eyes.', 'Break Reminder', 10);
}

// ==================== THEME MANAGEMENT ====================

function getCurrentTheme() {
  return PropertiesService.getUserProperties().getProperty('currentTheme') || THEME_CONFIG.DEFAULT_THEME;
}

function applyTheme(themeKey, scope) {
  scope = scope || 'all';
  var theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) throw new Error('Invalid theme');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = scope === 'all' ? ss.getSheets() : [ss.getActiveSheet()];
  sheets.forEach(function(sheet) { applyThemeToSheet(sheet, theme); });
  PropertiesService.getUserProperties().setProperty('currentTheme', themeKey);
  ss.toast(theme.icon + ' ' + theme.name + ' theme applied!', 'Theme', 3);
}

function applyThemeToSheet(sheet, theme) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(theme.headerBackground).setFontColor(theme.headerText).setFontWeight('bold');
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    sheet.getBandings().forEach(function(b) { b.remove(); });
    var banding = dataRange.applyRowBanding();
    banding.setFirstRowColor(theme.oddRow).setSecondRowColor(theme.evenRow);
    dataRange.setFontColor(theme.text);
  }
  sheet.setTabColor(theme.accent);
}

function previewTheme(themeKey) {
  var theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) throw new Error('Invalid theme');
  applyThemeToSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), theme);
  SpreadsheetApp.getActiveSpreadsheet().toast('👁️ Previewing ' + theme.name, 'Preview', 5);
}

function resetToDefaultTheme() {
  applyTheme(THEME_CONFIG.DEFAULT_THEME, 'all');
}

function quickToggleDarkMode() {
  var current = getCurrentTheme();
  applyTheme(current === 'LIGHT' ? 'DARK' : 'LIGHT', 'all');
}

// ==================== COMFORT VIEW CONTROL PANEL ====================

function showADHDControlPanel() {
  var settings = getADHDSettings();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8;border-bottom:3px solid #1a73e8;padding-bottom:10px}.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px;border-left:4px solid #1a73e8}.row{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;margin:5px}button:hover{background:#1557b0}button.sec{background:#6c757d}</style></head><body><div class="container"><h2>♿ Comfort View Panel</h2><div class="section"><div class="row"><span>Zebra Stripes</span><button onclick="google.script.run.toggleZebraStripes();setTimeout(function(){location.reload()},1000)">' + (settings.zebraStripes ? '✅ On' : 'Off') + '</button></div><div class="row"><span>Gridlines</span><button onclick="google.script.run.toggleGridlinesADHD();setTimeout(function(){location.reload()},1000)">' + (settings.gridlines ? '✅ Visible' : 'Hidden') + '</button></div><div class="row"><span>Focus Mode</span><button onclick="google.script.run.activateFocusMode();google.script.host.close()">🎯 Activate</button></div></div><div class="section"><div class="row"><span>Quick Capture</span><button onclick="google.script.run.showQuickCaptureNotepad()">📝 Open</button></div><div class="row"><span>Pomodoro Timer</span><button onclick="google.script.run.startPomodoroTimer();google.script.host.close()">⏱️ Start</button></div></div><button class="sec" onclick="google.script.run.resetADHDSettings();google.script.host.close()">🔄 Reset</button><button class="sec" onclick="google.script.host.close()">Close</button></div></body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '♿ Comfort View Panel');
}

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
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:5px}button.sec{background:#6c757d}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:15px;margin:20px 0}</style></head><body><div class="container"><h2>🎨 Theme Manager</h2><div class="grid">' + themeCards + '</div><button onclick="apply()">✅ Apply Theme</button><button class="sec" onclick="google.script.host.close()">Close</button></div><script>var sel="' + current + '";function select(k){sel=k;document.querySelectorAll(".grid>div").forEach(function(d){d.style.border="3px solid transparent"});event.currentTarget.style.border="3px solid #1a73e8"}function apply(){google.script.run.withSuccessHandler(function(){alert("Theme applied!");google.script.host.close()}).applyTheme(sel,"all")}</script></body></html>'
  ).setWidth(450).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🎨 Theme Manager');
}

// ==================== SETUP DEFAULTS ====================

/**
 * Setup Comfort View defaults with options dialog
 * User can choose which settings to apply and settings can be undone
 */
function setupADHDDefaults() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px;max-width:500px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.option{display:flex;align-items:center;padding:12px;margin:8px 0;background:#f8f9fa;border-radius:8px;cursor:pointer}' +
    '.option:hover{background:#e8f0fe}' +
    '.option input{margin-right:12px;width:18px;height:18px}' +
    '.option-text{flex:1}' +
    '.option-label{font-weight:bold;font-size:14px}' +
    '.option-desc{font-size:12px;color:#666;margin-top:2px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 24px;border:none;border-radius:4px;cursor:pointer;font-size:14px;flex:1}' +
    '.primary{background:#1a73e8;color:white}' +
    '.secondary{background:#e0e0e0;color:#333}' +
    '.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:15px;font-size:13px}' +
    '</style></head><body><div class="container">' +
    '<h2>🎨 Comfort View Setup</h2>' +
    '<div class="info">💡 These settings can be undone anytime via the Comfort View Panel or by running "Undo Comfort View"</div>' +
    '<div class="option" onclick="toggle(\'gridlines\')"><input type="checkbox" id="gridlines" checked><div class="option-text"><div class="option-label">Hide Gridlines</div><div class="option-desc">Reduce visual clutter by hiding sheet gridlines</div></div></div>' +
    '<div class="option" onclick="toggle(\'zebra\')"><input type="checkbox" id="zebra"><div class="option-text"><div class="option-label">Zebra Stripes</div><div class="option-desc">Alternating row colors for easier reading</div></div></div>' +
    '<div class="option" onclick="toggle(\'fontSize\')"><input type="checkbox" id="fontSize"><div class="option-text"><div class="option-label">Larger Font (12pt)</div><div class="option-desc">Increase default font size for better readability</div></div></div>' +
    '<div class="option" onclick="toggle(\'focus\')"><input type="checkbox" id="focus"><div class="option-text"><div class="option-label">Focus Mode</div><div class="option-desc">Hide all sheets except the active one</div></div></div>' +
    '<div class="buttons">' +
    '<button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="primary" onclick="apply()">Apply Settings</button>' +
    '</div></div>' +
    '<script>' +
    'function toggle(id){var cb=document.getElementById(id);cb.checked=!cb.checked}' +
    'function apply(){' +
    'var opts={gridlines:document.getElementById("gridlines").checked,zebra:document.getElementById("zebra").checked,fontSize:document.getElementById("fontSize").checked,focus:document.getElementById("focus").checked};' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()}).applyADHDDefaultsWithOptions(opts)}' +
    '</script></body></html>'
  ).setWidth(500).setHeight(450);
  ui.showModalDialog(html, '🎨 Comfort View Setup');
}

/**
 * Apply Comfort View defaults with selected options
 * @param {Object} options - Selected options
 */
function applyADHDDefaultsWithOptions(options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var applied = [];

  try {
    // Store previous settings for undo
    var prevSettings = {
      gridlinesWereHidden: [],
      zebraWasApplied: false,
      previousFontSize: 10
    };

    var sheets = ss.getSheets();

    if (options.gridlines) {
      sheets.forEach(function(sheet) {
        var name = sheet.getName();
        if (name !== SHEETS.CONFIG && name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) {
          sheet.setHiddenGridlines(true);
        }
      });
      applied.push('✅ Gridlines hidden on dashboard sheets');
    }

    if (options.zebra) {
      sheets.forEach(function(sheet) {
        applyZebraStripes(sheet);
      });
      saveADHDSettings({zebraStripes: true});
      applied.push('✅ Zebra stripes applied');
    }

    if (options.fontSize) {
      sheets.forEach(function(sheet) {
        if (sheet.getLastRow() > 0) {
          sheet.getDataRange().setFontSize(12);
        }
      });
      saveADHDSettings({fontSize: 12});
      applied.push('✅ Font size increased to 12pt');
    }

    if (options.focus) {
      activateFocusMode();
      applied.push('✅ Focus mode activated');
    }

    ss.toast(applied.join('\n'), '🎨 Setup Complete', 5);

  } catch (e) {
    SpreadsheetApp.getUi().alert('⚠️ Error: ' + e.message);
  }
}

/**
 * Undo Comfort View defaults - restore original settings
 */
function undoADHDDefaults() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('↩️ Undo Comfort View',
    'This will:\n\n' +
    '• Show all gridlines\n' +
    '• Remove zebra stripes\n' +
    '• Reset font size to 10pt\n' +
    '• Exit focus mode (show all sheets)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Show all gridlines
    ss.getSheets().forEach(function(sheet) {
      sheet.setHiddenGridlines(false);
    });

    // Remove zebra stripes
    ss.getSheets().forEach(function(sheet) {
      removeZebraStripes(sheet);
    });

    // Reset font size
    ss.getSheets().forEach(function(sheet) {
      if (sheet.getLastRow() > 0) {
        sheet.getDataRange().setFontSize(10);
      }
    });

    // Exit focus mode
    deactivateFocusMode();

    // Reset stored settings
    resetADHDSettings();

    ui.alert('↩️ Undo Complete',
      'Comfort View defaults have been reset:\n\n' +
      '✅ Gridlines restored\n' +
      '✅ Zebra stripes removed\n' +
      '✅ Font size reset to 10pt\n' +
      '✅ Focus mode deactivated',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('⚠️ Error: ' + e.message);
  }
}
