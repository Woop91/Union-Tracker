/**
 * ============================================================================
 * 16_DashboardEnhancements.gs - DASHBOARD ENHANCEMENT FEATURES
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Eight enhancement features for the steward/member dashboards:
 *   (1) Custom date ranges for trend filtering, (2) Chart export to Drive
 *   as PNG images, (3) Scheduled email reports with charts, (4) Web push
 *   notifications (service worker), (5) Multi-user chart collaboration,
 *   (6) Saved chart presets, (7) Advanced filtering options, (8) Multi-level
 *   drill-down capabilities. saveChartImageToDrive() handles server-side
 *   chart image storage.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   These enhancements extend the base dashboards without modifying core
 *   dashboard code. Chart export validates input (name sanitization, 10MB
 *   size limit) before saving to Drive. Each feature is independently
 *   callable — they don't depend on each other.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Chart exports fail. Scheduled email reports stop. Custom date filtering
 *   reverts to defaults. Drill-down navigation breaks. Core dashboard
 *   display still works — these are supplementary features.
 *
 * DEPENDENCIES:
 *   Depends on DriveApp (chart image storage), 01_Core.gs.
 *   Used by the SPA dashboard views and menu items.
 *
 * @version 4.33.0
 * @license Free for use by non-profit collective bargaining groups and unions
 * ============================================================================
 */

// ============================================================================
// 1. CUSTOM DATE RANGES - Server-side helpers
// ============================================================================
// ============================================================================
// 2. EXPORT INDIVIDUAL CHARTS AS IMAGES - Server-side download handler
// ============================================================================

/**
 * Saves an exported chart image (base64 PNG) to the user's Drive
 * @param {string} chartName - Name identifier for the chart
 * @param {string} base64Data - Base64-encoded PNG data (without data: prefix)
 * @returns {string} Drive file URL
 */
function saveChartImageToDrive(chartName, base64Data) {
  if (!chartName || typeof chartName !== 'string') throw new Error('Invalid chart name');
  if (!base64Data || typeof base64Data !== 'string') throw new Error('Invalid image data');
  chartName = chartName.replace(/[^a-zA-Z0-9_\- ]/g, '_').substring(0, 100);
  if (base64Data.length > 10 * 1024 * 1024) throw new Error('Image data too large');

  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    'image/png',
    chartName + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.png'
  );

  var folder = getOrCreateExportFolder_();
  var file = folder.createFile(blob);
  return file.getUrl();
}

/**
 * Gets or creates the Dashboard Exports folder in Drive
 * @returns {Folder} Google Drive folder
 * @private
 */
function getOrCreateExportFolder_() {
  var folderName = 'Dashboard Exports';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}
// ============================================================================
// 3. SCHEDULED EMAIL REPORTS
// ============================================================================

/**
 * Schedules a recurring dashboard email report
 * @param {Object} config - Report configuration
 * @param {string} config.email - Recipient email address
 * @param {string} config.frequency - 'daily', 'weekly', or 'monthly'
 * @param {string[]} config.sections - Sections to include ('summary','charts','trends','satisfaction')
 * @param {boolean} config.includePII - Whether to include PII data
 * @returns {Object} Schedule confirmation with trigger ID
 */
function scheduleEmailReport(config) {
  if (!config || !config.email) {
    return { success: false, error: 'Email address is required' };
  }

  // Validate email format
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(config.email)) {
    return { success: false, error: 'Invalid email address format' };
  }

  // Authorization check — PII reports may only be sent to stewards or admins
  var recipientRole = typeof getUserRole_ === 'function' ? getUserRole_(config.email) : null;
  if (recipientRole !== 'admin' && recipientRole !== 'steward') {
    var callerEmail = '';
    try { callerEmail = Session.getActiveUser().getEmail(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    // Allow sending to self even if not steward/admin (non-PII only)
    if (config.includePII && callerEmail.toLowerCase() !== config.email.toLowerCase()) {
      return { success: false, error: 'Reports containing PII can only be sent to stewards or admins.' };
    }
  }

  // Store report config in script properties
  var props = PropertiesService.getScriptProperties();
  var schedules = JSON.parse(props.getProperty('report_schedules') || '[]');

  var schedule = {
    id: 'rpt_' + Date.now(),
    email: config.email,
    frequency: config.frequency || 'weekly',
    sections: config.sections || ['summary', 'charts'],
    includePII: config.includePII || false,
    createdBy: Session.getActiveUser().getEmail(),
    createdAt: new Date().toISOString(),
    active: true
  };

  schedules.push(schedule);
  props.setProperty('report_schedules', JSON.stringify(schedules));

  // Install trigger if not already installed
  installReportTrigger_();

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('report:scheduled', schedule);
  }

  return { success: true, schedule: schedule };
}

/**
 * Sends all scheduled email reports (called by trigger)
 */
function sendScheduledReports() {
  var props = PropertiesService.getScriptProperties();
  var schedules = JSON.parse(props.getProperty('report_schedules') || '[]');

  var now = new Date();
  var dayOfWeek = now.getDay(); // 0=Sun
  var dayOfMonth = now.getDate();

  for (var i = 0; i < schedules.length; i++) {
    var sched = schedules[i];
    if (!sched.active) continue;

    var shouldSend = false;
    if (sched.frequency === 'daily') {
      shouldSend = true;
    } else if (sched.frequency === 'weekly' && dayOfWeek === 1) {
      shouldSend = true;
    } else if (sched.frequency === 'monthly' && dayOfMonth === 1) {
      shouldSend = true;
    }

    if (shouldSend) {
      try {
        sendDashboardReportEmail_(sched);
      } catch (e) {
        Logger.log('Failed to send report to ' + sched.email + ': ' + e.message);
      }
    }
  }
}

/**
 * Sends a single dashboard report email
 * @param {Object} schedule - The schedule config
 * @private
 */
function sendDashboardReportEmail_(schedule) {
  var data = JSON.parse(getUnifiedDashboardData(schedule.includePII));
  var sections = schedule.sections;

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#1e293b;color:#f8fafc;padding:24px;border-radius:12px">';
  html += '<h1 style="color:#3b82f6;font-size:20px;margin-bottom:16px">Dashboard Report</h1>';
  html += '<p style="color:#94a3b8;font-size:12px">Generated ' + new Date().toLocaleDateString() + '</p>';

  if (sections.indexOf('summary') >= 0) {
    html += '<div style="background:#334155;padding:16px;border-radius:8px;margin:16px 0">';
    html += '<h2 style="font-size:14px;color:#e2e8f0;margin-bottom:12px">Summary Metrics</h2>';
    html += '<table style="width:100%;color:#e2e8f0;font-size:13px">';
    html += '<tr><td>Total Members</td><td style="text-align:right;font-weight:bold">' + escapeHtml(String(data.totalMembers)) + '</td></tr>';
    html += '<tr><td>Open Cases</td><td style="text-align:right;font-weight:bold">' + escapeHtml(String(data.openCases)) + '</td></tr>';
    html += '<tr><td>Win Rate</td><td style="text-align:right;font-weight:bold">' + escapeHtml(String(data.winRate)) + '%</td></tr>';
    html += '<tr><td>Morale Score</td><td style="text-align:right;font-weight:bold">' + escapeHtml(String(data.moraleScore)) + '/10</td></tr>';
    html += '</table></div>';
  }

  if (sections.indexOf('trends') >= 0 && data.monthlyFilings) {
    html += '<div style="background:#334155;padding:16px;border-radius:8px;margin:16px 0">';
    html += '<h2 style="font-size:14px;color:#e2e8f0;margin-bottom:12px">Monthly Filings Trend</h2>';
    html += '<table style="width:100%;color:#e2e8f0;font-size:13px">';
    for (var m = 0; m < data.monthlyFilings.length; m++) {
      html += '<tr><td>' + escapeHtml(String(data.monthlyFilings[m].month)) + '</td><td style="text-align:right">' + escapeHtml(String(data.monthlyFilings[m].count)) + ' filed</td></tr>';
    }
    html += '</table></div>';
  }

  if (sections.indexOf('satisfaction') >= 0) {
    html += '<div style="background:#334155;padding:16px;border-radius:8px;margin:16px 0">';
    html += '<h2 style="font-size:14px;color:#e2e8f0;margin-bottom:12px">Satisfaction</h2>';
    html += '<p style="color:#e2e8f0">Overall Morale: <strong>' + escapeHtml(String(data.moraleScore)) + '/10</strong></p>';
    html += '</div>';
  }

  html += '<p style="color:#64748b;font-size:11px;margin-top:24px;text-align:center">This is an automated report. Manage your subscriptions in the dashboard settings.</p>';
  html += '</div>';

  MailApp.sendEmail({
    to: schedule.email,
    subject: 'Dashboard Report - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy'),
    htmlBody: html
  });

}

/**
 * Installs the daily trigger for report dispatch
 * @private
 */
function installReportTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendScheduledReports') {
      return; // Already installed
    }
  }
  ScriptApp.newTrigger('sendScheduledReports')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}

/**
 * Gets all scheduled reports for the current user
 * @returns {string} JSON array of schedules
 */
function getScheduledReports() {
  var props = PropertiesService.getScriptProperties();
  var schedules = JSON.parse(props.getProperty('report_schedules') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  var userSchedules = schedules.filter(function(s) {
    return s.createdBy === userEmail;
  });

  return JSON.stringify(userSchedules);
}

/**
 * Removes a scheduled report
 * @param {string} scheduleId - The schedule ID to remove
 * @returns {Object} Removal confirmation
 */
function removeScheduledReport(scheduleId) {
  var props = PropertiesService.getScriptProperties();
  var schedules = JSON.parse(props.getProperty('report_schedules') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  schedules = schedules.filter(function(s) {
    return !(s.id === scheduleId && s.createdBy === userEmail);
  });

  props.setProperty('report_schedules', JSON.stringify(schedules));

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('report:removed', { scheduleId: scheduleId });
  }

  return { success: true };
}
// ============================================================================
// 4. NOTIFICATIONS (v4.22.0 — rerouted to sheet-based system in 05_Integrations.gs)
// ============================================================================
// getUserNotifications() and markNotificationRead() removed — used ScriptProperties
// storage which is orphaned by the Notifications sheet system (v4.13.0+).
// broadcastStewardNotification() removed — no active callers; steward compose
// form handles broadcast directly via sendWebAppNotification().
//
// pushNotification() is kept because saveSharedView() (below) calls it.
// It now writes to the Notifications sheet instead of ScriptProperties.

/**
 * Server-side push for a single-user notification.
 * Called internally by saveSharedView() and EventBus trigger handlers.
 * Writes to the 📢 Notifications sheet (same system as sendWebAppNotification).
 * @param {string} userEmail - Target user email
 * @param {Object} notification - { title, body, type }
 * @returns {Object} { success, id }
 */
function pushNotification(userEmail, notification) {
  try {
    if (!userEmail || !notification || !notification.title) {
      return { success: false, error: 'userEmail and notification.title are required.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) {
      if (typeof createNotificationsSheet === 'function') {
        sheet = createNotificationsSheet(ss);
      } else {
        return { success: false, error: 'Notifications sheet not found.' };
      }
    }

    var allData = sheet.getDataRange().getValues();
    var maxNum = 0;
    var C = NOTIFICATIONS_COLS;
    for (var i = 1; i < allData.length; i++) {
      var existId = String(allData[i][C.NOTIFICATION_ID - 1] || '');
      var match = existId.match(/NOTIF-(\d+)/);
      if (match) { var num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
    }
    var nextId = 'NOTIF-' + String(maxNum + 1).padStart(3, '0');

    var tz = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    sheet.appendRow([
      nextId,
      userEmail,                                      // RECIPIENT
      'System',                                       // TYPE
      String(notification.title || ''),               // TITLE
      String(notification.body || ''),                // MESSAGE
      'Normal',                                       // PRIORITY
      Session.getActiveUser().getEmail() || 'system', // SENT_BY
      'System',                                       // SENT_BY_NAME
      today,                                          // CREATED_DATE
      '',                                             // EXPIRES_DATE
      '',                                             // DISMISSED_BY
      'Active',                                       // STATUS
      'Dismissible'                                   // DISMISS_MODE
    ]);

    if (typeof EventBus !== 'undefined' && EventBus.emit) {
      EventBus.emit('notification:pushed', { userId: userEmail, id: nextId });
    }

    return { success: true, id: nextId };
  } catch (e) {
    Logger.log('pushNotification error: ' + e.message);
    return { success: false, error: e.message };
  }
}
// ============================================================================
// 5. MULTI-USER COLLABORATION ON CHART SELECTIONS
// ============================================================================

/**
 * Saves a shared dashboard view configuration
 * @param {Object} view - View configuration
 * @param {string} view.name - View name
 * @param {string[]} view.selectedCharts - Array of chart IDs to display
 * @param {Object} view.filters - Active filters
 * @param {string[]} view.sharedWith - Email addresses to share with
 * @returns {Object} Saved view with ID
 */
function saveSharedView(view) {
  return withScriptLock_(function() {
  var props = PropertiesService.getScriptProperties();
  var views = JSON.parse(props.getProperty('shared_views') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  var sharedView = {
    id: 'view_' + Date.now(),
    name: escapeForFormula(view.name || 'Untitled View'),
    selectedCharts: view.selectedCharts || [],
    filters: view.filters || {},
    dateRange: view.dateRange || null,
    createdBy: userEmail,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    sharedWith: view.sharedWith || [],
    comments: []
  };

  views.push(sharedView);
  props.setProperty('shared_views', JSON.stringify(views));

  // Notify shared users
  for (var i = 0; i < sharedView.sharedWith.length; i++) {
    pushNotification(sharedView.sharedWith[i], {
      title: 'Dashboard View Shared',
      body: userEmail + ' shared "' + sharedView.name + '" with you',
      type: 'info'
    });
  }

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('collaboration:viewCreated', { viewId: sharedView.id, name: sharedView.name });
  }

  return { success: true, view: sharedView };
  });
}

/**
 * Gets all shared views accessible to the current user
 * @returns {string} JSON array of shared views
 */
function getSharedViews() {
  var props = PropertiesService.getScriptProperties();
  var views = JSON.parse(props.getProperty('shared_views') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  var accessible = views.filter(function(v) {
    return v.createdBy === userEmail ||
           (v.sharedWith && v.sharedWith.indexOf(userEmail) >= 0);
  });

  return JSON.stringify(accessible);
}


/**
 * Deletes a shared view (only owner can delete)
 * @param {string} viewId - View ID to delete
 * @returns {Object} Deletion result
 */
function deleteSharedView(viewId) {
  var props = PropertiesService.getScriptProperties();
  var views = JSON.parse(props.getProperty('shared_views') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  var filtered = views.filter(function(v) {
    return !(v.id === viewId && v.createdBy === userEmail);
  });

  if (filtered.length === views.length) {
    return { success: false, error: 'View not found or not authorized' };
  }

  props.setProperty('shared_views', JSON.stringify(filtered));

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('collaboration:viewDeleted', { viewId: viewId });
  }

  return { success: true };
}
// ============================================================================
// 6. SAVED CHART CONFIGURATIONS (PRESETS)
// ============================================================================

/**
 * Saves a chart configuration preset
 * @param {Object} preset - Preset configuration
 * @param {string} preset.name - Preset display name
 * @param {string[]} preset.visibleCharts - Chart IDs to show
 * @param {Object} preset.chartOptions - Per-chart overrides (colors, types)
 * @param {Object} preset.layout - Layout configuration (grid columns, ordering)
 * @param {Object} preset.filters - Pre-applied filters
 * @returns {Object} Saved preset with ID
 */
function saveChartPreset(preset) {
  return withScriptLock_(function() {
  var props = PropertiesService.getUserProperties();
  var presets = JSON.parse(props.getProperty('chart_presets') || '[]');

  var chartPreset = {
    id: 'preset_' + Date.now(),
    name: preset.name || 'Custom Preset',
    visibleCharts: preset.visibleCharts || [],
    chartOptions: preset.chartOptions || {},
    layout: preset.layout || { columns: 2 },
    filters: preset.filters || {},
    dateRange: preset.dateRange || null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  presets.push(chartPreset);
  props.setProperty('chart_presets', JSON.stringify(presets));

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('preset:saved', chartPreset);
  }

  return { success: true, preset: chartPreset };
  });
}

/**
 * Gets all saved chart presets for the current user
 * @returns {string} JSON array of presets
 */
function getChartPresets() {
  var props = PropertiesService.getUserProperties();
  return props.getProperty('chart_presets') || '[]';
}

/**
 * Deletes a chart preset
 * @param {string} presetId - Preset ID to delete
 * @returns {Object} Deletion result
 */
function deleteChartPreset(presetId) {
  var props = PropertiesService.getUserProperties();
  var presets = JSON.parse(props.getProperty('chart_presets') || '[]');

  presets = presets.filter(function(p) { return p.id !== presetId; });
  props.setProperty('chart_presets', JSON.stringify(presets));

  if (typeof EventBus !== 'undefined' && EventBus.emit) {
    EventBus.emit('preset:deleted', { presetId: presetId });
  }

  return { success: true };
}

// ============================================================================
// 7. ADVANCED FILTERING OPTIONS
// ============================================================================

/**
 * Applies advanced filters to dashboard data and returns filtered results
 * @param {boolean} isPII - Whether to include PII data
 * @param {Object} filters - Filter configuration
 * @param {string[]} [filters.statuses] - Filter by grievance statuses
 * @param {string[]} [filters.locations] - Filter by locations
 * @param {string[]} [filters.stewards] - Filter by steward names
 * @param {string[]} [filters.categories] - Filter by grievance categories
 * @param {string[]} [filters.units] - Filter by units
 * @param {string} [filters.dateFrom] - Start date (ISO)
 * @param {string} [filters.dateTo] - End date (ISO)
 * @param {string} [filters.searchText] - Text search across all fields
 * @param {number[]} [filters.steps] - Filter by grievance steps (1,2,3,4)
 * @param {number} [filters.minMorale] - Minimum morale score
 * @param {number} [filters.maxMorale] - Maximum morale score
 * @returns {string} JSON filtered dashboard data
 */
function getFilteredDashboardData(isPII, filters) {
  var fullData = JSON.parse(getUnifiedDashboardData(isTruthyValue(isPII)));

  if (!filters || Object.keys(filters).length === 0) {
    return JSON.stringify(fullData);
  }

  // Apply date range filter first (reuse existing function)
  if (filters.dateFrom || filters.dateTo) {
    fullData = JSON.parse(getUnifiedDashboardDataWithDateRange(
      isPII, 0, filters.dateFrom || '2000-01-01', filters.dateTo || new Date().toISOString()
    ));
  }

  // Filter status distribution to only selected statuses
  if (filters.statuses && filters.statuses.length > 0) {
    var filteredStatus = {};
    for (var sKey in fullData.statusDistribution) {
      if (filters.statuses.indexOf(sKey) >= 0) {
        filteredStatus[sKey] = fullData.statusDistribution[sKey];
      }
    }
    fullData.statusDistribution = filteredStatus;
  }

  // Filter location breakdown
  if (filters.locations && filters.locations.length > 0) {
    var filteredLocations = {};
    for (var locKey in fullData.locationBreakdown) {
      if (filters.locations.indexOf(locKey) >= 0) {
        filteredLocations[locKey] = fullData.locationBreakdown[locKey];
      }
    }
    fullData.locationBreakdown = filteredLocations;
  }

  // Filter unit breakdown
  if (filters.units && filters.units.length > 0) {
    var filteredUnits = {};
    for (var uKey in fullData.unitBreakdown) {
      if (filters.units.indexOf(uKey) >= 0) {
        filteredUnits[uKey] = fullData.unitBreakdown[uKey];
      }
    }
    fullData.unitBreakdown = filteredUnits;
  }

  // Filter grievances by category
  if (filters.categories && filters.categories.length > 0 && fullData.grievancesByCategory) {
    var filteredCats = {};
    for (var catKey in fullData.grievancesByCategory) {
      if (filters.categories.indexOf(catKey) >= 0) {
        filteredCats[catKey] = fullData.grievancesByCategory[catKey];
      }
    }
    fullData.grievancesByCategory = filteredCats;
  }

  // Filter drill-down data by steward
  if (filters.stewards && filters.stewards.length > 0 && fullData.chartDrillDown) {
    for (var ddKey in fullData.chartDrillDown.statusByCase) {
      fullData.chartDrillDown.statusByCase[ddKey] = fullData.chartDrillDown.statusByCase[ddKey].filter(function(c) {
        return !c.steward || filters.stewards.indexOf(c.steward) >= 0;
      });
    }
  }

  // Filter by step progression
  if (filters.steps && filters.steps.length > 0) {
    var filteredSteps = { step1: 0, step2: 0, step3: 0, arb: 0 };
    if (filters.steps.indexOf(1) >= 0) filteredSteps.step1 = fullData.stepProgression.step1;
    if (filters.steps.indexOf(2) >= 0) filteredSteps.step2 = fullData.stepProgression.step2;
    if (filters.steps.indexOf(3) >= 0) filteredSteps.step3 = fullData.stepProgression.step3;
    if (filters.steps.indexOf(4) >= 0) filteredSteps.arb = fullData.stepProgression.arb;
    fullData.stepProgression = filteredSteps;
  }

  // Text search filter on drill-down case data
  if (filters.searchText) {
    var searchLower = filters.searchText.toLowerCase();
    if (fullData.chartDrillDown) {
      for (var drillKey in fullData.chartDrillDown) {
        var drillGroup = fullData.chartDrillDown[drillKey];
        for (var gKey in drillGroup) {
          drillGroup[gKey] = drillGroup[gKey].filter(function(item) {
            var searchable = JSON.stringify(item).toLowerCase();
            return searchable.indexOf(searchLower) >= 0;
          });
        }
      }
    }
  }

  fullData.filtersApplied = filters;
  return JSON.stringify(fullData);
}
// ============================================================================
// 8. DRILL-DOWN CAPABILITIES (Multi-level hierarchical)
// ============================================================================


