/**
 * ============================================================================
 * 29_TrendAlertService.gs — Trend Alert Detection Service
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Detects unusual patterns in grievance, satisfaction, and performance data
 *   and surfaces them as actionable alerts on the steward dashboard. Runs via
 *   daily trigger and stores alerts in a hidden _Trend_Alerts sheet.
 *
 * DETECTION ALGORITHMS:
 *   - Grievance spike at a location (7-day count > 2x 30-day daily average)
 *   - Satisfaction score drop (>15% quarter-over-quarter)
 *   - Steward win rate decline (>20% drop over 90 days)
 *   - Deadline miss rate spike (current month > 3-month average by >50%)
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Proactive alerting helps stewards identify problems early rather than
 *   discovering them in retrospective data. Deduplication prevents repeated
 *   alerts for the same issue. Alerts are acknowledgeable/resolvable to
 *   track steward response.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Trend alerts stop generating. The steward dashboard still works — alerts
 *   section simply shows empty. No data loss risk.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, RSVP_STATUS, GRIEVANCE_COLS),
 *   00_Security.gs (escapeForFormula, logAuditEvent)
 *
 * @version 4.36.0
 */

// ============================================================================
// TREND ALERT SERVICE
// ============================================================================

var TrendAlertService = (function () {

  var HEADERS = ['ID', 'Type', 'Severity', 'Title', 'Message', 'Data', 'Created', 'Status', 'Acknowledged By', 'Acknowledged At'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.TREND_ALERTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.TREND_ALERTS);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Generates a deduplication key for an alert to avoid repeats.
   * @param {string} type - Alert type
   * @param {string} dataSignature - Unique data identifier (e.g., location name)
   * @returns {string} Dedup key
   */
  function _dedupKey(type, dataSignature) {
    return type + ':' + (dataSignature || '').toLowerCase().trim();
  }

  /**
   * Checks if an alert with this dedup key already exists (Active or Acknowledged).
   * @param {Array} existingAlerts - Current alert rows
   * @param {string} key - Dedup key
   * @returns {boolean}
   */
  function _alertExists(existingAlerts, key) {
    for (var i = 0; i < existingAlerts.length; i++) {
      var row = existingAlerts[i];
      var status = String(row[7]).trim();
      if (status === 'Resolved') continue;
      var existingKey = _dedupKey(String(row[1]), String(row[5]));
      if (existingKey === key) return true;
    }
    return false;
  }

  /**
   * Detects grievance spikes at specific locations.
   * Alert triggers when a location's 7-day count > 2x the 30-day daily average,
   * with a minimum threshold of 3 filings in 7 days to avoid noise.
   * @returns {Array} Array of alert objects
   */
  function _detectGrievanceSpike() {
    var alerts = [];
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return alerts;
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet || sheet.getLastRow() < 2) return alerts;
      var data = sheet.getDataRange().getValues();
      var headers = data[0];

      // Find Date Filed and Work Location columns
      var dateCol = -1, locCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (h === 'date filed') dateCol = c;
        if (h === 'work location' || h === 'location') locCol = c;
      }
      if (dateCol === -1 || locCol === -1) return alerts;

      var now = new Date();
      var day7ago = new Date(now.getTime() - 7 * 86400000);
      var day30ago = new Date(now.getTime() - 30 * 86400000);

      // Count grievances per location for 7-day and 30-day windows
      var locations = {}; // { locName: { count7d, count30d } }
      for (var i = 1; i < data.length; i++) {
        var filed = data[i][dateCol];
        if (!(filed instanceof Date)) continue;
        var loc = String(data[i][locCol] || '').trim();
        if (!loc) continue;

        if (!locations[loc]) locations[loc] = { count7d: 0, count30d: 0 };
        if (filed >= day30ago) {
          locations[loc].count30d++;
          if (filed >= day7ago) {
            locations[loc].count7d++;
          }
        }
      }

      // Check each location for spikes
      var locNames = Object.keys(locations);
      for (var li = 0; li < locNames.length; li++) {
        var name = locNames[li];
        var counts = locations[name];

        // Minimum threshold: at least 3 in the 7-day window
        if (counts.count7d < 3) continue;

        // 30-day daily average, then expected 7-day count at that rate
        var dailyAvg = counts.count30d / 30;
        var expected7d = dailyAvg * 7;

        // Spike if 7-day count > 2x the expected 7-day volume
        if (expected7d > 0 && counts.count7d > expected7d * 2) {
          alerts.push({
            type: 'GRIEVANCE_SPIKE',
            severity: 'HIGH',
            title: 'Grievance Spike: ' + name,
            message: counts.count7d + ' grievances filed at ' + name + ' in the last 7 days (vs ' + Math.round(expected7d) + ' expected based on 30-day average of ' + dailyAvg.toFixed(1) + '/day).',
            data: name
          });
        }
      }
    } catch (e) {
      Logger.log('_detectGrievanceSpike error: ' + e.message);
    }
    return alerts;
  }

  /**
   * Detects satisfaction score drops > 15% quarter-over-quarter.
   * @returns {Array} Array of alert objects
   */
  function _detectSatisfactionDrop() {
    var alerts = [];
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return alerts;
      var sheet = ss.getSheetByName(SHEETS.SATISFACTION || '_Satisfaction');
      if (!sheet || sheet.getLastRow() < 2) return alerts;
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      if (data.length < 2) return alerts;

      // Group responses by quarter
      var quarters = {};
      for (var i = 0; i < data.length; i++) {
        var ts = data[i][0];
        if (!(ts instanceof Date)) continue;
        var qKey = ts.getFullYear() + '-Q' + (Math.floor(ts.getMonth() / 3) + 1);
        if (!quarters[qKey]) quarters[qKey] = { total: 0, count: 0 };
        // Average numeric values across columns (skip timestamp col 0)
        var rowSum = 0;
        var rowCount = 0;
        for (var c = 1; c < data[i].length; c++) {
          var val = parseFloat(data[i][c]);
          if (!isNaN(val) && val >= 1 && val <= 5) {
            rowSum += val;
            rowCount++;
          }
        }
        if (rowCount > 0) {
          quarters[qKey].total += rowSum / rowCount;
          quarters[qKey].count++;
        }
      }

      var qKeys = Object.keys(quarters).sort();
      if (qKeys.length < 2) return alerts;

      var prev = quarters[qKeys[qKeys.length - 2]];
      var curr = quarters[qKeys[qKeys.length - 1]];
      var prevAvg = prev.count > 0 ? prev.total / prev.count : 0;
      var currAvg = curr.count > 0 ? curr.total / curr.count : 0;

      if (prevAvg > 0 && currAvg > 0) {
        var dropPct = ((prevAvg - currAvg) / prevAvg) * 100;
        if (dropPct > 15) {
          alerts.push({
            type: 'SATISFACTION_DROP',
            severity: 'HIGH',
            title: 'Satisfaction Score Drop',
            message: 'Average satisfaction dropped ' + Math.round(dropPct) + '% from ' + qKeys[qKeys.length - 2] + ' (' + prevAvg.toFixed(1) + ') to ' + qKeys[qKeys.length - 1] + ' (' + currAvg.toFixed(1) + ').',
            data: qKeys[qKeys.length - 1]
          });
        }
      }
    } catch (e) {
      Logger.log('_detectSatisfactionDrop error: ' + e.message);
    }
    return alerts;
  }

  /**
   * Detects steward win rate decline > 20% over 90-day windows.
   * @returns {Array} Array of alert objects
   */
  function _detectWinRateDecline() {
    var alerts = [];
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return alerts;
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet || sheet.getLastRow() < 2) return alerts;
      var data = sheet.getDataRange().getValues();
      var headers = data[0];

      // Find columns
      var statusCol = -1, resolvedDateCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (h === 'status') statusCol = c;
        if (h === 'resolution date' || h === 'resolved date' || h === 'date resolved') resolvedDateCol = c;
      }
      if (statusCol === -1) return alerts;

      var now = new Date();
      var day90ago = new Date(now.getTime() - 90 * 86400000);
      var day180ago = new Date(now.getTime() - 180 * 86400000);

      var recentResolved = 0, recentWon = 0;
      var priorResolved = 0, priorWon = 0;

      for (var i = 1; i < data.length; i++) {
        var status = String(data[i][statusCol]).toLowerCase().trim();
        if (status !== 'settled' && status !== 'withdrawn' && status !== 'denied' && status !== 'resolved') continue;

        var resolvedDate = resolvedDateCol >= 0 ? data[i][resolvedDateCol] : null;
        if (!(resolvedDate instanceof Date)) continue;

        var isWon = (status === 'settled');

        if (resolvedDate >= day90ago) {
          recentResolved++;
          if (isWon) recentWon++;
        } else if (resolvedDate >= day180ago) {
          priorResolved++;
          if (isWon) priorWon++;
        }
      }

      if (priorResolved >= 5 && recentResolved >= 3) {
        var priorRate = priorWon / priorResolved;
        var recentRate = recentWon / recentResolved;
        if (priorRate > 0 && recentRate < priorRate) {
          var declinePct = ((priorRate - recentRate) / priorRate) * 100;
          if (declinePct > 20) {
            alerts.push({
              type: 'WIN_RATE_DECLINE',
              severity: 'MEDIUM',
              title: 'Win Rate Decline',
              message: 'Win rate dropped from ' + Math.round(priorRate * 100) + '% (prior 90 days) to ' + Math.round(recentRate * 100) + '% (recent 90 days) — a ' + Math.round(declinePct) + '% decline.',
              data: 'win_rate_' + now.toISOString().substring(0, 7)
            });
          }
        }
      }
    } catch (e) {
      Logger.log('_detectWinRateDecline error: ' + e.message);
    }
    return alerts;
  }

  /**
   * Detects deadline miss rate increases > 50% vs 3-month average.
   * @returns {Array} Array of alert objects
   */
  function _detectDeadlineMissSpike() {
    var alerts = [];
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return alerts;
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet || sheet.getLastRow() < 2) return alerts;
      var data = sheet.getDataRange().getValues();
      var headers = data[0];

      var deadlineCol = -1, statusCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (h === 'next action due' || h === 'deadline') deadlineCol = c;
        if (h === 'status') statusCol = c;
      }
      if (deadlineCol === -1) return alerts;

      var now = new Date();
      var thisMonth = now.getFullYear() * 12 + now.getMonth();
      var monthCounts = {}; // { monthKey: { total, missed } }

      for (var i = 1; i < data.length; i++) {
        var deadline = data[i][deadlineCol];
        if (!(deadline instanceof Date)) continue;
        var status = String(data[i][statusCol] || '').toLowerCase().trim();
        if (status === 'resolved' || status === 'settled' || status === 'withdrawn') continue;

        var monthKey = deadline.getFullYear() * 12 + deadline.getMonth();
        if (!monthCounts[monthKey]) monthCounts[monthKey] = { total: 0, missed: 0 };
        monthCounts[monthKey].total++;
        if (deadline < now) monthCounts[monthKey].missed++;
      }

      // Compare current month miss rate vs 3-month average
      var currMisses = monthCounts[thisMonth] ? monthCounts[thisMonth].missed : 0;
      var currTotal = monthCounts[thisMonth] ? monthCounts[thisMonth].total : 0;
      var avgMisses = 0, avgCount = 0;
      for (var m = 1; m <= 3; m++) {
        var mk = thisMonth - m;
        if (monthCounts[mk]) {
          avgMisses += monthCounts[mk].missed;
          avgCount++;
        }
      }

      if (avgCount > 0 && currTotal >= 3) {
        var avgMissRate = avgMisses / avgCount;
        if (avgMissRate > 0 && currMisses > avgMissRate * 1.5) {
          alerts.push({
            type: 'DEADLINE_MISS_SPIKE',
            severity: 'HIGH',
            title: 'Deadline Miss Rate Spike',
            message: currMisses + ' deadline misses this month vs ' + Math.round(avgMissRate) + '/month average (3-month). That\'s a ' + Math.round(((currMisses - avgMissRate) / avgMissRate) * 100) + '% increase.',
            data: 'deadline_' + now.toISOString().substring(0, 7)
          });
        }
      }
    } catch (e) {
      Logger.log('_detectDeadlineMissSpike error: ' + e.message);
    }
    return alerts;
  }

  /**
   * Runs all detection algorithms, deduplicates, and writes new alerts.
   * Called by daily trigger.
   * @returns {number} Count of new alerts generated
   */
  function runDetection() {
    var sheet = initSheet();
    var existingAlerts = [];
    if (sheet.getLastRow() >= 2) {
      existingAlerts = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    }

    var newAlerts = [];
    newAlerts = newAlerts.concat(_detectGrievanceSpike());
    newAlerts = newAlerts.concat(_detectSatisfactionDrop());
    newAlerts = newAlerts.concat(_detectWinRateDecline());
    newAlerts = newAlerts.concat(_detectDeadlineMissSpike());

    var written = 0;
    for (var i = 0; i < newAlerts.length; i++) {
      var alert = newAlerts[i];
      var key = _dedupKey(alert.type, alert.data);
      if (_alertExists(existingAlerts, key)) continue;

      sheet.appendRow([
        Utilities.getUuid().substring(0, 8),
        escapeForFormula(alert.type),
        escapeForFormula(alert.severity),
        escapeForFormula(alert.title),
        escapeForFormula(alert.message),
        escapeForFormula(alert.data || ''),
        new Date(),
        'Active',
        '',
        ''
      ]);
      written++;
    }

    if (written > 0 && typeof logAuditEvent === 'function') {
      logAuditEvent('TREND_ALERTS_GENERATED', { newAlerts: written });
    }
    return written;
  }

  /**
   * Returns all active (non-resolved) trend alerts.
   * @returns {Array} Alert objects
   */
  function getActiveAlerts() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.TREND_ALERTS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var alerts = [];
    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][7]).trim();
      if (status === 'Resolved') continue;
      alerts.push({
        id: data[i][0],
        type: data[i][1],
        severity: data[i][2],
        title: data[i][3],
        message: data[i][4],
        data: data[i][5],
        created: data[i][6],
        status: status,
        acknowledgedBy: data[i][8],
        acknowledgedAt: data[i][9]
      });
    }
    return alerts;
  }

  /**
   * Acknowledges an alert.
   * @param {string} alertId - Alert ID
   * @param {string} stewardEmail - Who acknowledged
   * @returns {Object} { success }
   */
  function acknowledgeAlert(alertId, stewardEmail) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName(SHEETS.TREND_ALERTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === alertId) {
        sheet.getRange(i + 1, 8).setValue('Acknowledged');
        sheet.getRange(i + 1, 9).setValue(escapeForFormula(stewardEmail));
        sheet.getRange(i + 1, 10).setValue(new Date());
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('TREND_ALERT_ACKNOWLEDGED', { alertId: alertId, by: typeof maskEmail === 'function' ? maskEmail(stewardEmail) : stewardEmail });
        }
        return { success: true };
      }
    }
    return { success: false };
  }

  /**
   * Resolves an alert (marks it as handled).
   * @param {string} alertId - Alert ID
   * @returns {Object} { success }
   */
  function resolveAlert(alertId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName(SHEETS.TREND_ALERTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === alertId) {
        sheet.getRange(i + 1, 8).setValue('Resolved');
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('TREND_ALERT_RESOLVED', { alertId: alertId });
        }
        return { success: true };
      }
    }
    return { success: false };
  }

  return {
    initSheet: initSheet,
    runDetection: runDetection,
    getActiveAlerts: getActiveAlerts,
    acknowledgeAlert: acknowledgeAlert,
    resolveAlert: resolveAlert
  };
})();

// ============================================================================
// TREND ALERT GLOBAL WRAPPERS (v4.36.0)
// ============================================================================

/** @param {string} sessionToken @returns {Array} Active trend alerts. */
function dataGetTrendAlerts(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return [];
  return TrendAlertService.getActiveAlerts();
}

/** @param {string} sessionToken @param {string} alertId @returns {Object} Result. */
function dataAcknowledgeTrendAlert(sessionToken, alertId) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  return TrendAlertService.acknowledgeAlert(alertId, s);
}

/** @param {string} sessionToken @param {string} alertId @returns {Object} Result. */
function dataResolveTrendAlert(sessionToken, alertId) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  return TrendAlertService.resolveAlert(alertId);
}

/** Daily trigger handler — runs all trend detection algorithms. No auth needed. */
function triggerDailyTrendDetection() {
  TrendAlertService.runDetection();
}
