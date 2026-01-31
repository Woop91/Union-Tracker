/**
 * ============================================================================
 * ThemeService.gs - Theme Management and Visual Settings
 * ============================================================================
 *
 * This module handles all theme-related functions including:
 * - Theme application (dark mode, light mode, etc.)
 * - Visual settings persistence
 * - Comfort view / ADHD-friendly settings
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Theme management and visual settings functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// THEME APPLICATION
// ============================================================================

/**
 * Applies the system theme to all sheets
 * @returns {void}
 */
function APPLY_SYSTEM_THEME() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    // Skip hidden calculation sheets
    if (sheetName.indexOf('_Calc') === 0) return;

    applyThemeToSheet_(sheet);
  });

  ss.toast('Theme applied to all sheets!', 'Theme', 3);
}

/**
 * Applies theme styling to a single sheet
 * @param {Sheet} sheet - The sheet to style
 * @returns {void}
 * @private
 */
function applyThemeToSheet_(sheet) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastCol < 1 || lastRow < 1) return;

  // Apply header styling (row 1)
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Apply alternating row colors for data rows
  if (lastRow > 1) {
    for (var row = 2; row <= lastRow; row++) {
      var rowRange = sheet.getRange(row, 1, 1, lastCol);
      if (row % 2 === 0) {
        rowRange.setBackground('#f1f5f9');
      } else {
        rowRange.setBackground('#ffffff');
      }
    }
  }
}

/**
 * Applies global styling to the spreadsheet
 * @returns {void}
 */
function applyGlobalStyling() {
  APPLY_SYSTEM_THEME();
}

/**
 * Resets all sheets to default theme
 * @returns {void}
 */
function resetToDefaultTheme() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName.indexOf('_Calc') === 0) return;

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    if (lastCol < 1 || lastRow < 1) return;

    // Reset header to default
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange
      .setBackground(COLORS.PRIMARY_PURPLE)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold');

    // Reset data rows to white
    if (lastRow > 1) {
      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange
        .setBackground('#ffffff')
        .setFontColor('#1e293b');
    }
  });

  ss.toast('Theme reset to default!', 'Theme', 3);
}

/**
 * Saves a visual setting
 * @param {string} setting - The setting name
 * @param {*} value - The setting value
 * @returns {void}
 */
function saveVisualSetting(setting, value) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('visual_' + setting, JSON.stringify(value));
}

/**
 * Applies a dashboard theme
 * @param {string} theme - Theme name ('default', 'dark', 'light', 'contrast')
 * @returns {void}
 */
function applyDashboardTheme(theme) {
  saveVisualSetting('theme', theme);
  SpreadsheetApp.getActiveSpreadsheet().toast('Theme changed to ' + theme, 'Theme', 2);
}

/**
 * Toggles dark mode
 * @returns {void}
 */
function toggleDarkMode() {
  var props = PropertiesService.getUserProperties();
  var isDark = props.getProperty('visual_darkMode') === 'true';
  props.setProperty('visual_darkMode', (!isDark).toString());
  SpreadsheetApp.getActiveSpreadsheet().toast(
    isDark ? 'Dark mode disabled' : 'Dark mode enabled',
    'Theme', 2
  );
}

/**
 * Shows the theme settings dialog
 * @returns {void}
 */
function showThemeSettings() {
  showVisualControlPanel();
}

// ============================================================================
// COMFORT VIEW / ADHD-FRIENDLY SETTINGS
// ============================================================================

/**
 * Gets ADHD/Comfort View settings
 * @returns {Object} Settings object
 */
function getADHDSettings() {
  var props = PropertiesService.getUserProperties();
  var settings = props.getProperty('adhdSettings');
  return settings ? JSON.parse(settings) : getDefaultADHDSettings_();
}

/**
 * Gets default ADHD settings
 * @returns {Object} Default settings
 * @private
 */
function getDefaultADHDSettings_() {
  return {
    zebraStripes: true,
    reducedMotion: false,
    focusMode: false,
    highContrast: false,
    largeText: false,
    hideGridlines: false
  };
}

/**
 * Saves ADHD/Comfort View settings
 * @param {Object} settings - The settings to save
 * @returns {void}
 */
function saveADHDSettings(settings) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('adhdSettings', JSON.stringify(settings));
}

/**
 * Applies ADHD/Comfort View settings
 * @param {Object} settings - The settings to apply
 * @returns {void}
 */
function applyADHDSettings(settings) {
  if (settings.zebraStripes) {
    applyZebraStripesToAllSheets_();
  }
  if (settings.hideGridlines) {
    hideAllGridlines();
  }
  saveADHDSettings(settings);
}

/**
 * Resets ADHD/Comfort View settings to defaults
 * @returns {void}
 */
function resetADHDSettings() {
  saveADHDSettings(getDefaultADHDSettings_());
  SpreadsheetApp.getActiveSpreadsheet().toast('Comfort View settings reset', 'Settings', 2);
}

/**
 * Hides gridlines on all sheets
 * @returns {void}
 */
function hideAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    sheet.setHiddenGridlines(true);
  });
}

/**
 * Shows gridlines on all sheets
 * @returns {void}
 */
function showAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    sheet.setHiddenGridlines(false);
  });
}

/**
 * Toggles gridlines visibility
 * @returns {void}
 */
function toggleGridlinesADHD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  var hidden = currentSheet.hasHiddenGridlines();
  currentSheet.setHiddenGridlines(!hidden);
}

/**
 * Applies zebra stripes to a sheet
 * @param {Sheet} sheet - The sheet to style
 * @returns {void}
 */
function applyZebraStripes(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return;

  for (var row = 2; row <= lastRow; row++) {
    var rowRange = sheet.getRange(row, 1, 1, lastCol);
    if (row % 2 === 0) {
      rowRange.setBackground('#f1f5f9');
    } else {
      rowRange.setBackground('#ffffff');
    }
  }
}

/**
 * Applies zebra stripes to all sheets
 * @returns {void}
 * @private
 */
function applyZebraStripesToAllSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getName().indexOf('_Calc') !== 0) {
      applyZebraStripes(sheet);
    }
  });
}

/**
 * Removes zebra stripes from a sheet
 * @param {Sheet} sheet - The sheet to clear
 * @returns {void}
 */
function removeZebraStripes(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return;

  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.setBackground('#ffffff');
}

/**
 * Toggles zebra stripes on active sheet
 * @returns {void}
 */
function toggleZebraStripes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if zebra stripes are applied (row 2 has different color than row 3)
  var row2Color = sheet.getRange(2, 1).getBackground();
  var row3Color = sheet.getRange(3, 1).getBackground();

  if (row2Color !== row3Color) {
    removeZebraStripes(sheet);
  } else {
    applyZebraStripes(sheet);
  }
}

/**
 * Toggles reduced motion setting
 * @returns {void}
 */
function toggleReducedMotion() {
  var settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    settings.reducedMotion ? 'Reduced motion enabled' : 'Reduced motion disabled',
    'Accessibility', 2
  );
}

/**
 * Activates focus mode
 * @returns {void}
 */
function activateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selection = sheet.getSelection();

  if (!selection) {
    ss.toast('Please select a cell first', 'Focus Mode', 2);
    return;
  }

  var settings = getADHDSettings();
  settings.focusMode = true;
  saveADHDSettings(settings);

  ss.toast('Focus mode activated. Press Esc to exit.', 'Focus Mode', 3);
}

/**
 * Deactivates focus mode
 * @returns {void}
 */
function deactivateFocusMode() {
  var settings = getADHDSettings();
  settings.focusMode = false;
  saveADHDSettings(settings);

  SpreadsheetApp.getActiveSpreadsheet().toast('Focus mode deactivated', 'Focus Mode', 2);
}

/**
 * Gets the current theme
 * @returns {string} Current theme name
 */
function getCurrentTheme() {
  var props = PropertiesService.getUserProperties();
  return props.getProperty('visual_theme') || 'default';
}

/**
 * Applies a theme
 * @param {string} themeKey - Theme key
 * @param {string} scope - 'all' or 'current'
 * @returns {void}
 */
function applyTheme(themeKey, scope) {
  saveVisualSetting('theme', themeKey);

  if (scope === 'all') {
    APPLY_SYSTEM_THEME();
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    applyThemeToSheet_(sheet);
  }
}

/**
 * Applies theme to a specific sheet
 * @param {Sheet} sheet - The sheet
 * @param {Object} theme - Theme object
 * @returns {void}
 */
function applyThemeToSheet(sheet, theme) {
  applyThemeToSheet_(sheet);
}

/**
 * Previews a theme
 * @param {string} themeKey - Theme key to preview
 * @returns {void}
 */
function previewTheme(themeKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  applyThemeToSheet_(sheet);
  ss.toast('Previewing ' + themeKey + ' theme', 'Theme Preview', 2);
}

/**
 * Resets to default theme via system
 * @returns {void}
 * @private
 */
function resetToDefaultThemeViaSystem_() {
  resetToDefaultTheme();
}

/**
 * Quick toggle for dark mode
 * @returns {void}
 */
function quickToggleDarkMode() {
  toggleDarkMode();
}

/**
 * Sets up ADHD/Comfort View defaults
 * @returns {void}
 */
function setupADHDDefaults() {
  var settings = getDefaultADHDSettings_();
  settings.zebraStripes = true;
  applyADHDSettings(settings);
  SpreadsheetApp.getActiveSpreadsheet().toast('Comfort View defaults applied', 'Settings', 3);
}

/**
 * Applies ADHD defaults with options
 * @param {Object} options - Options to apply
 * @returns {void}
 */
function applyADHDDefaultsWithOptions(options) {
  var settings = getADHDSettings();
  Object.assign(settings, options);
  applyADHDSettings(settings);
}

/**
 * Undoes ADHD/Comfort View defaults
 * @returns {void}
 */
function undoADHDDefaults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.getName().indexOf('_Calc') !== 0) {
      removeZebraStripes(sheet);
      sheet.setHiddenGridlines(false);
    }
  });

  resetADHDSettings();
  ss.toast('Comfort View settings removed', 'Settings', 3);
}

/**
 * Refreshes all visual elements
 * @returns {void}
 */
function refreshAllVisuals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Reapply theme
  APPLY_SYSTEM_THEME();

  // Reapply ADHD settings if enabled
  var settings = getADHDSettings();
  if (settings.zebraStripes) {
    applyZebraStripesToAllSheets_();
  }

  ss.toast('All visuals refreshed!', 'Refresh', 3);
}
