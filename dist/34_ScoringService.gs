/* exported ScoringService */
/**
 * ScoringService — 3-dimension member scoring used by the Org Health Tree.
 *
 * Dimensions: engagement (default 70%), profile (20%), grievance (10%).
 * Configurable via Config sheet scoreWeightEngagement / scoreWeightProfile /
 * scoreWeightGrievance. Weights are normalized in calculateCompositeScore.
 *
 * NOTE (v4.55.1 T12-BUG-02): this service is intentionally distinct from
 * EngagementService (30_EngagementService.gs). The two services use different
 * dimensions and different weights and target different UIs:
 *   - ScoringService → Org Health Tree (high-level org health overview)
 *   - EngagementService → Member Report Card + Steward Scoreboard (activity detail)
 * A future version may consolidate them, but for now they coexist. If you edit
 * weights or add dimensions, update BOTH services and their documentation.
 */
var ScoringService = (function() {
  'use strict';

  function calculateEngagementScore(volunteerHours, openRate, committees) {
    var config = ConfigReader.getConfig();
    var maxHours = config.maxVolunteerHours || 20;
    var hoursScore = Math.min((Number(volunteerHours) || 0) / maxHours, 1) * 100;
    var openRateScore = Math.min(Math.max(Number(openRate) || 0, 0), 100);
    var committeeScore = (committees && String(committees).trim().length > 0) ? 100 : 0;
    return (hoursScore + openRateScore + committeeScore) / 3;
  }

  function calculateProfileScore(fields) {
    var keys = ['phone', 'email', 'street', 'city', 'state', 'zip', 'shirtSize', 'preferredComm', 'bestTime', 'officeDays'];
    var filled = 0;
    for (var i = 0; i < keys.length; i++) {
      if (fields[keys[i]] && String(fields[keys[i]]).trim().length > 0) filled++;
    }
    return (filled / keys.length) * 100;
  }

  function calculateGrievanceScore(hasOpenGrievance, grievanceStatus, daysToDeadline, direction) {
    if (direction === 'Positive') {
      var closedStatuses = ['closed', 'withdrawn'];
      if (hasOpenGrievance && closedStatuses.indexOf(String(grievanceStatus).toLowerCase()) === -1) return 100;
      return 50;
    }
    if (!hasOpenGrievance) return 100;
    var days = Number(daysToDeadline) || 0;
    // v4.55.1 A08-BUG-01: progressive scaling above 7 days so 8-day and 60-day grievances
    // do not both return 70. Past deadline: 0. 0-7 days: 30-40. 7-30: 50-70. 30-90: 70-85. 90+: 90.
    if (days < 0) return 0;
    if (days <= 7) return 30 + Math.round((days / 7) * 10);
    if (days <= 30) return 50 + Math.round(((days - 7) / 23) * 20);
    if (days <= 90) return 70 + Math.round(((days - 30) / 60) * 15);
    return 90;
  }

  function calculateCompositeScore(engagement, profile, grievance) {
    var config = ConfigReader.getConfig();
    var wE = (Number(config.scoreWeightEngagement) || 70) / 100;
    var wP = (Number(config.scoreWeightProfile) || 20) / 100;
    var wG = (Number(config.scoreWeightGrievance) || 10) / 100;
    // v4.55.1: normalize weights if they don't sum to 1 so composite is always 0-100
    var sum = wE + wP + wG;
    if (sum > 0 && Math.abs(sum - 1) > 0.001) {
      wE = wE / sum;
      wP = wP / sum;
      wG = wG / sum;
    }
    return Math.round(engagement * wE + profile * wP + grievance * wG);
  }

  function getScoreColor(score) {
    var config = ConfigReader.getConfig();
    var green = Number(config.scoreThresholdGreen) || 70;
    var yellow = Number(config.scoreThresholdYellow) || 40;
    if (score >= green) return '#4CAF50';
    if (score >= yellow) return '#FFC107';
    return '#f44336';
  }

  function autoAssignMembers(members, stewards) {
    if (!stewards || stewards.length === 0) return [];
    var unassigned = members.filter(function(m) { return !m.steward || !m.steward.trim(); });
    var results = [];
    for (var i = 0; i < unassigned.length; i++) {
      results.push({ email: unassigned[i].email, steward: stewards[i % stewards.length] });
    }
    return results;
  }

  function locationBasedAssign(members, locationMapStr, stewards) {
    var map = {};
    if (locationMapStr) {
      locationMapStr.split(',').forEach(function(pair) {
        var parts = pair.split(':');
        if (parts.length === 2) map[parts[0].trim().toLowerCase()] = parts[1].trim();
      });
    }
    var unmatched = [];
    var results = [];
    var unassigned = members.filter(function(m) { return !m.steward || !m.steward.trim(); });
    for (var i = 0; i < unassigned.length; i++) {
      var loc = String(unassigned[i].location || '').trim().toLowerCase();
      if (map[loc]) { results.push({ email: unassigned[i].email, steward: map[loc] }); }
      else { unmatched.push(unassigned[i]); }
    }
    for (var j = 0; j < unmatched.length; j++) {
      results.push({ email: unmatched[j].email, steward: stewards[j % stewards.length] });
    }
    return results;
  }

  return {
    calculateEngagementScore: calculateEngagementScore,
    calculateProfileScore: calculateProfileScore,
    calculateGrievanceScore: calculateGrievanceScore,
    calculateCompositeScore: calculateCompositeScore,
    getScoreColor: getScoreColor,
    autoAssignMembers: autoAssignMembers,
    locationBasedAssign: locationBasedAssign,
  };
})();
