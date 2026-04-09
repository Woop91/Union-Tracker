/* exported ScoringService */
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
    if (days < 0) return 0;
    if (days < 7) return 40;
    return 70;
  }

  function calculateCompositeScore(engagement, profile, grievance) {
    var config = ConfigReader.getConfig();
    var wE = (Number(config.scoreWeightEngagement) || 70) / 100;
    var wP = (Number(config.scoreWeightProfile) || 20) / 100;
    var wG = (Number(config.scoreWeightGrievance) || 10) / 100;
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
