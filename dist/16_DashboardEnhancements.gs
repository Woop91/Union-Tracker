/**
 * ============================================================================
 * 16_DashboardEnhancements.gs - Dashboard Enhancement Features
 * ============================================================================
 *
 * Eight dashboard features that extend the public/executive dashboards:
 *
 * 1. Custom Date Ranges for Trend Analysis (client + server filter)
 * 2. Export Individual Charts as Images (client-side Chart.js canvas export)
 * 3. Scheduled Email Reports with Selected Charts
 * 4. Mobile App Push Notifications (web push via service worker)
 * 5. Multi-User Collaboration on Chart Selections
 * 6. Saved Chart Configurations (Presets)
 * 7. Advanced Filtering Options
 * 8. Drill-Down Capabilities (multi-level hierarchical)
 *
 * @version 4.9.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// 1. CUSTOM DATE RANGES - Server-side helpers
// ============================================================================

/**
 * Returns available date range presets for the UI picker
 * @returns {Object[]} Array of preset definitions
 */
function getDateRangePresets() {
  return [
    { id: 'last7',    label: 'Last 7 Days',    days: 7 },
    { id: 'last30',   label: 'Last 30 Days',   days: 30 },
    { id: 'last90',   label: 'Last 90 Days',   days: 90 },
    { id: 'last180',  label: 'Last 6 Months',  days: 180 },
    { id: 'lastYear', label: 'Last Year',       days: 365 },
    { id: 'ytd',      label: 'Year to Date',    days: -1 },
    { id: 'custom',   label: 'Custom Range',    days: 0 }
  ];
}

/**
 * Gets trend comparison data for two date ranges (period-over-period)
 * @param {boolean} isPII - Include PII data
 * @param {string} fromDate1 - First period start (ISO)
 * @param {string} toDate1 - First period end (ISO)
 * @param {string} fromDate2 - Second period start (ISO)
 * @param {string} toDate2 - Second period end (ISO)
 * @returns {string} JSON with both periods and calculated deltas
 */
function getTrendComparisonData(isPII, fromDate1, toDate1, fromDate2, toDate2) {
  var period1 = JSON.parse(getUnifiedDashboardDataWithDateRange(isPII, 0, fromDate1, toDate1));
  var period2 = JSON.parse(getUnifiedDashboardDataWithDateRange(isPII, 0, fromDate2, toDate2));

  var comparison = {
    period1: {
      label: fromDate1 + ' to ' + toDate1,
      totalCases: period1.totalCases || 0,
      openCases: period1.openCases || 0,
      winRate: period1.winRate || 0,
      moraleScore: period1.moraleScore || 0,
      totalMembers: period1.totalMembers || 0
    },
    period2: {
      label: fromDate2 + ' to ' + toDate2,
      totalCases: period2.totalCases || 0,
      openCases: period2.openCases || 0,
      winRate: period2.winRate || 0,
      moraleScore: period2.moraleScore || 0,
      totalMembers: period2.totalMembers || 0
    },
    deltas: {}
  };

  // Calculate percentage deltas
  var metrics = ['totalCases', 'openCases', 'winRate', 'moraleScore', 'totalMembers'];
  for (var i = 0; i < metrics.length; i++) {
    var key = metrics[i];
    var v1 = comparison.period1[key];
    var v2 = comparison.period2[key];
    comparison.deltas[key] = v1 === 0 ? 0 : Math.round(((v2 - v1) / v1) * 100);
  }

  return JSON.stringify(comparison);
}


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

  EventBus.emit('report:scheduled', { scheduleId: schedule.id, email: schedule.email });
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

  EventBus.emit('report:sent', { scheduleId: schedule.id, email: schedule.email });
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
  EventBus.emit('report:removed', { scheduleId: scheduleId });
  return { success: true };
}


// ============================================================================
// 4. MOBILE APP PUSH NOTIFICATIONS (In-app notification center)
// ============================================================================

/**
 * Gets pending notifications for the current user
 * @returns {string} JSON array of notifications
 */
function getUserNotifications() {
  var props = PropertiesService.getUserProperties();
  var notifications = JSON.parse(props.getProperty('dashboard_notifications') || '[]');

  // Prune notifications older than 30 days
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  notifications = notifications.filter(function(n) {
    return new Date(n.timestamp) > cutoff;
  });

  props.setProperty('dashboard_notifications', JSON.stringify(notifications));
  return JSON.stringify(notifications);
}

/**
 * Adds a notification for a specific user (server-side push)
 * @param {string} userEmail - Target user email
 * @param {Object} notification - Notification data
 * @param {string} notification.title - Notification title
 * @param {string} notification.body - Notification message body
 * @param {string} notification.type - Type: 'alert','info','success','warning'
 * @param {string} [notification.link] - Optional link/action URL
 * @returns {Object} Push result
 */
function pushNotification(userEmail, notification) {
  // For cross-user notifications, use ScriptProperties with user key
  var scriptProps = PropertiesService.getScriptProperties();
  var key = 'notifications_' + userEmail.replace(/[^a-zA-Z0-9]/g, '_');
  var notifications = JSON.parse(scriptProps.getProperty(key) || '[]');

  var note = {
    id: 'notif_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
    title: notification.title,
    body: notification.body,
    type: notification.type || 'info',
    link: notification.link || '',
    timestamp: new Date().toISOString(),
    read: false
  };

  notifications.unshift(note);

  // Keep max 50 notifications
  if (notifications.length > 50) {
    notifications = notifications.slice(0, 50);
  }

  scriptProps.setProperty(key, JSON.stringify(notifications));
  EventBus.emit('notification:pushed', { userId: userEmail, notificationId: note.id });
  return { success: true, id: note.id };
}

/**
 * Marks a notification as read
 * @param {string} notificationId - Notification ID
 * @returns {Object} Update result
 */
function markNotificationRead(notificationId) {
  var userEmail = Session.getActiveUser().getEmail();
  var scriptProps = PropertiesService.getScriptProperties();
  var key = 'notifications_' + userEmail.replace(/[^a-zA-Z0-9]/g, '_');
  var notifications = JSON.parse(scriptProps.getProperty(key) || '[]');

  for (var i = 0; i < notifications.length; i++) {
    if (notifications[i].id === notificationId) {
      notifications[i].read = true;
      break;
    }
  }

  scriptProps.setProperty(key, JSON.stringify(notifications));
  return { success: true };
}

/**
 * Sends notification to all stewards (bulk push)
 * @param {Object} notification - Notification data (title, body, type)
 * @returns {Object} Broadcast result with count
 */
function broadcastStewardNotification(notification) {
  var data = JSON.parse(getUnifiedDashboardData(true));
  var stewards = data.stewardPerformance || [];
  var sent = 0;

  for (var i = 0; i < stewards.length; i++) {
    if (stewards[i].email) {
      pushNotification(stewards[i].email, notification);
      sent++;
    }
  }

  EventBus.emit('notification:broadcast', { count: sent, type: notification.type });
  return { success: true, sentCount: sent };
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
  var props = PropertiesService.getScriptProperties();
  var views = JSON.parse(props.getProperty('shared_views') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  var sharedView = {
    id: 'view_' + Date.now(),
    name: view.name || 'Untitled View',
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

  EventBus.emit('collaboration:viewCreated', { viewId: sharedView.id });
  return { success: true, view: sharedView };
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
 * Adds a comment to a shared view
 * @param {string} viewId - View ID
 * @param {string} commentText - Comment content
 * @returns {Object} Updated view
 */
function addViewComment(viewId, commentText) {
  var props = PropertiesService.getScriptProperties();
  var views = JSON.parse(props.getProperty('shared_views') || '[]');
  var userEmail = Session.getActiveUser().getEmail();

  for (var i = 0; i < views.length; i++) {
    if (views[i].id === viewId) {
      var comment = {
        id: 'cmt_' + Date.now(),
        author: userEmail,
        text: commentText,
        timestamp: new Date().toISOString()
      };
      views[i].comments.push(comment);
      views[i].updatedAt = new Date().toISOString();

      props.setProperty('shared_views', JSON.stringify(views));

      // Notify the view owner if commenter is not the owner
      if (views[i].createdBy !== userEmail) {
        pushNotification(views[i].createdBy, {
          title: 'New Comment on View',
          body: userEmail + ' commented on "' + views[i].name + '"',
          type: 'info'
        });
      }

      EventBus.emit('collaboration:commentAdded', { viewId: viewId, commentId: comment.id });
      return { success: true, comment: comment };
    }
  }

  return { success: false, error: 'View not found' };
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
  EventBus.emit('collaboration:viewDeleted', { viewId: viewId });
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

  EventBus.emit('preset:saved', { presetId: chartPreset.id, name: chartPreset.name });
  return { success: true, preset: chartPreset };
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

  EventBus.emit('preset:deleted', { presetId: presetId });
  return { success: true };
}

/**
 * Updates an existing chart preset
 * @param {string} presetId - Preset ID to update
 * @param {Object} updates - Partial preset data to merge
 * @returns {Object} Updated preset
 */
function updateChartPreset(presetId, updates) {
  var props = PropertiesService.getUserProperties();
  var presets = JSON.parse(props.getProperty('chart_presets') || '[]');

  for (var i = 0; i < presets.length; i++) {
    if (presets[i].id === presetId) {
      if (updates.name) presets[i].name = updates.name;
      if (updates.visibleCharts) presets[i].visibleCharts = updates.visibleCharts;
      if (updates.chartOptions) presets[i].chartOptions = updates.chartOptions;
      if (updates.layout) presets[i].layout = updates.layout;
      if (updates.filters) presets[i].filters = updates.filters;
      if (updates.dateRange !== undefined) presets[i].dateRange = updates.dateRange;
      presets[i].updatedAt = new Date().toISOString();

      props.setProperty('chart_presets', JSON.stringify(presets));
      EventBus.emit('preset:updated', { presetId: presetId });
      return { success: true, preset: presets[i] };
    }
  }

  return { success: false, error: 'Preset not found' };
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

/**
 * Returns available filter options based on current data
 * @param {boolean} isPII - Whether to include PII data
 * @returns {string} JSON with available filter values per dimension
 */
function getAvailableFilterOptions(isPII) {
  var data = JSON.parse(getUnifiedDashboardData(isTruthyValue(isPII)));

  var options = {
    statuses: Object.keys(data.statusDistribution || {}),
    locations: Object.keys(data.locationBreakdown || {}),
    units: Object.keys(data.unitBreakdown || {}),
    categories: Object.keys(data.grievancesByCategory || {}),
    stewards: [],
    steps: [1, 2, 3, 4]
  };

  // Extract steward names from drill-down data
  if (data.chartDrillDown && data.chartDrillDown.stewardByCase) {
    options.stewards = Object.keys(data.chartDrillDown.stewardByCase);
  }

  // Extract from steward performance if available
  if (data.stewardPerformance) {
    for (var i = 0; i < data.stewardPerformance.length; i++) {
      var name = data.stewardPerformance[i].name;
      if (name && options.stewards.indexOf(name) < 0) {
        options.stewards.push(name);
      }
    }
  }

  return JSON.stringify(options);
}

/**
 * Saves a named filter preset for quick reuse
 * @param {string} name - Filter preset name
 * @param {Object} filters - Filter configuration object
 * @returns {Object} Saved filter preset with ID
 */
function saveFilterPreset(name, filters) {
  var props = PropertiesService.getUserProperties();
  var presets = JSON.parse(props.getProperty('filter_presets') || '[]');

  var preset = {
    id: 'filt_' + Date.now(),
    name: name,
    filters: filters,
    createdAt: new Date().toISOString()
  };

  presets.push(preset);
  props.setProperty('filter_presets', JSON.stringify(presets));

  EventBus.emit('filter:presetSaved', { presetId: preset.id, name: name });
  return { success: true, preset: preset };
}

/**
 * Gets saved filter presets
 * @returns {string} JSON array of filter presets
 */
function getFilterPresets() {
  var props = PropertiesService.getUserProperties();
  return props.getProperty('filter_presets') || '[]';
}

/**
 * Deletes a saved filter preset
 * @param {string} presetId - Preset ID to delete
 * @returns {Object} Deletion result
 */
function deleteFilterPreset(presetId) {
  var props = PropertiesService.getUserProperties();
  var presets = JSON.parse(props.getProperty('filter_presets') || '[]');
  presets = presets.filter(function(p) { return p.id !== presetId; });
  props.setProperty('filter_presets', JSON.stringify(presets));
  return { success: true };
}


// ============================================================================
// 8. DRILL-DOWN CAPABILITIES (Multi-level hierarchical)
// ============================================================================

/**
 * Gets detailed drill-down data for a specific chart segment with multiple levels
 * @param {boolean} isPII - Whether to include PII data
 * @param {string} chartType - Chart category: 'status','location','category','unit','steward','step'
 * @param {string} primaryKey - First-level key (e.g. status name, location name)
 * @param {string} [secondaryKey] - Second-level key for deeper drill (e.g. sub-location, sub-category)
 * @returns {string} JSON drill-down data with items and sub-groups
 */
function getMultiLevelDrillDown(isPII, chartType, primaryKey, secondaryKey) {
  var data = JSON.parse(getUnifiedDashboardData(isTruthyValue(isPII)));
  var drillDown = data.chartDrillDown || {};
  var result = {
    chartType: chartType,
    primaryKey: primaryKey,
    secondaryKey: secondaryKey || null,
    items: [],
    subGroups: {},
    breadcrumbs: [{ label: 'All Data', key: null }],
    totalCount: 0
  };

  // Level 1: Get primary data
  var primaryItems = [];
  if (chartType === 'status' && drillDown.statusByCase) {
    primaryItems = drillDown.statusByCase[primaryKey] || [];
    result.breadcrumbs.push({ label: 'Status: ' + primaryKey, key: primaryKey });
  } else if (chartType === 'location' && drillDown.locationByCase) {
    primaryItems = drillDown.locationByCase[primaryKey] || [];
    result.breadcrumbs.push({ label: 'Location: ' + primaryKey, key: primaryKey });
  } else if (chartType === 'category' && drillDown.categoryByCase) {
    primaryItems = drillDown.categoryByCase[primaryKey] || [];
    result.breadcrumbs.push({ label: 'Category: ' + primaryKey, key: primaryKey });
  } else if (chartType === 'unit' && drillDown.unitByMember) {
    primaryItems = drillDown.unitByMember[primaryKey] || [];
    result.breadcrumbs.push({ label: 'Unit: ' + primaryKey, key: primaryKey });
  } else if (chartType === 'steward' && drillDown.stewardByCase) {
    primaryItems = drillDown.stewardByCase[primaryKey] || [];
    result.breadcrumbs.push({ label: 'Steward: ' + primaryKey, key: primaryKey });
  }

  // Level 2: Sub-group by secondary dimension
  if (!secondaryKey) {
    // Show sub-groups available for deeper drill
    var subGroupDimensions = getSubGroupDimensions_(chartType);
    for (var d = 0; d < subGroupDimensions.length; d++) {
      var dim = subGroupDimensions[d];
      var groups = {};
      for (var i = 0; i < primaryItems.length; i++) {
        var item = primaryItems[i];
        var groupKey = item[dim.field] || 'Unknown';
        if (!groups[groupKey]) groups[groupKey] = [];
        groups[groupKey].push(item);
      }
      result.subGroups[dim.label] = {};
      for (var gk in groups) {
        result.subGroups[dim.label][gk] = groups[gk].length;
      }
    }
    result.items = primaryItems;
  } else {
    // Level 2 drill: filter primary items by secondary key
    result.breadcrumbs.push({ label: secondaryKey, key: secondaryKey });
    var subDims = getSubGroupDimensions_(chartType);
    for (var s = 0; s < subDims.length; s++) {
      var field = subDims[s].field;
      var filtered = primaryItems.filter(function(item) {
        return item[field] === secondaryKey;
      });
      if (filtered.length > 0) {
        primaryItems = filtered;
        break;
      }
    }
    result.items = primaryItems;
  }

  result.totalCount = result.items.length;
  return JSON.stringify(result);
}

/**
 * Returns which fields can be used for sub-grouping based on chart type
 * @param {string} chartType - Primary chart type
 * @returns {Object[]} Array of { field, label } for sub-group dimensions
 * @private
 */
function getSubGroupDimensions_(chartType) {
  var allDimensions = [
    { field: 'status', label: 'By Status' },
    { field: 'location', label: 'By Location' },
    { field: 'steward', label: 'By Steward' },
    { field: 'category', label: 'By Category' },
    { field: 'step', label: 'By Step' }
  ];

  // Exclude the primary dimension from sub-groups
  return allDimensions.filter(function(d) {
    return d.field !== chartType;
  });
}

/**
 * Gets a summary overview for drill-down (counts per sub-group)
 * @param {boolean} isPII - Whether to include PII data
 * @param {string} chartType - Chart category
 * @returns {string} JSON summary with counts per group
 */
function getDrillDownSummary(isPII, chartType) {
  var data = JSON.parse(getUnifiedDashboardData(isTruthyValue(isPII)));
  var drillDown = data.chartDrillDown || {};
  var summary = { chartType: chartType, groups: {} };

  var sourceMap = {
    status: drillDown.statusByCase,
    location: drillDown.locationByCase,
    category: drillDown.categoryByCase,
    unit: drillDown.unitByMember,
    steward: drillDown.stewardByCase
  };

  var source = sourceMap[chartType] || {};
  for (var key in source) {
    summary.groups[key] = {
      count: source[key].length,
      label: key
    };
  }

  return JSON.stringify(summary);
}
