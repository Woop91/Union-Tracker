// ============================================================================
// 10. UNIFIED WEB APP DASHBOARDS (v4.4.0)
// ============================================================================
// Both Steward and Member dashboards are web apps with identical dark theme.
// Only difference: Member dashboard hides PII (names, contact info).
// Features: Clickable pill boxes, enhanced analytics, directory trends,
// engagement metrics, Google Drive resources tab.
// ============================================================================

/**
 * Shows the Member Dashboard (web app URL)
 * Opens the unified dashboard in member mode (no PII)
 */
function showPublicMemberDashboard() {
  var url = ScriptApp.getService().getUrl() + '?mode=member';
  var html = HtmlService.createHtmlOutput(
    '<html><head>' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;background:#0f172a;color:#f8fafc;margin:0;padding:16px}' +
    '.icon{font-size:clamp(36px,10vw,48px);margin-bottom:16px}h1{margin:0 0 8px;font-size:clamp(18px,5vw,24px);text-align:center}p{color:#94a3b8;margin:0 0 24px;text-align:center;max-width:400px;line-height:1.5;font-size:clamp(13px,3.5vw,15px);padding:0 8px}' +
    'a.open-link{background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:600;display:inline-block;min-height:44px;line-height:20px;text-align:center}a.open-link:hover{background:#2563eb}' +
    '.copy-btn{background:#475569;cursor:pointer;border:none;padding:10px 16px;border-radius:8px;color:white;font-size:clamp(11px,3vw,13px);min-height:44px}' +
    '.url{background:#1e293b;padding:12px;border-radius:8px;font-family:monospace;font-size:clamp(10px,2.5vw,12px);word-break:break-all;max-width:90%;margin-bottom:16px;border:1px solid #334155;width:100%}' +
    '.btn-row{display:flex;gap:8px;flex-wrap:wrap;justify-content:center}' +
    '@media(max-width:480px){.btn-row{flex-direction:column;width:100%}a.open-link,.copy-btn{width:100%;text-align:center}}' +
    '</style></head><body><div class="icon">👥</div><h1>Member Dashboard</h1>' +
    '<p>Open the Member Dashboard web app. This version hides sensitive personal information while showing full analytics.</p>' +
    '<div class="url" id="url">' + url + '</div>' +
    '<div class="btn-row"><a class="open-link" href="' + url + '" target="_blank">Open Dashboard</a>' +
    '<button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'url\').textContent);this.textContent=\'Copied!\';setTimeout(function(){document.querySelector(\'.copy-btn\').textContent=\'Copy URL\'},2000)">Copy URL</button></div>' +
    '</body></html>'
  ).setWidth(500).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Dashboard');
}

/**
 * Gets comprehensive dashboard data for unified dashboard (v4.4.0)
 * @param {boolean} includePII - Whether to include personally identifiable information
 * @returns {string} JSON with all dashboard data
 */
function getUnifiedDashboardData(includePII) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var data = {
    mode: includePII ? 'steward' : 'member',
    timestamp: new Date().toISOString(),

    // KPIs
    totalMembers: 0,
    stewardCount: 0,
    totalGrievances: 0,
    openGrievances: 0,
    wins: 0,
    losses: 0,
    settled: 0,
    winRate: 0,
    overdueCount: 0,
    moraleScore: 7.5,

    // Trend comparison data (previous 30-day period)
    prevOpenCases: 0,
    prevWinRate: 0,
    prevOverdueCount: 0,
    prevMoraleScore: 0,
    prevTotalGrievances: 0,

    // Lists for drill-down (steward mode only)
    stewardList: [],
    memberList: [],
    overdueList: [],
    openCasesList: [],
    myCases: [],           // Current user's assigned cases (v4.4.0)
    currentUserEmail: '',  // For My Cases identification

    // Performance metrics
    stewardRatio: 'N/A',   // Member:Steward ratio
    topPerformers: [],     // Top performing stewards

    // Breakdowns
    unitBreakdown: {},
    locationBreakdown: {},
    officeDaysBreakdown: {},
    stewardWorkload: [],
    hotZones: [],
    // Enhanced Hot Spots (multiple types)
    hotSpots: {
      grievance: [],        // Locations with 3+ active grievances
      dissatisfaction: [],  // Locations/units with satisfaction < 5
      lowEngagement: [],    // Locations/units with engagement < 30%
      overdueConcentration: [] // Locations with multiple overdue cases
    },

    // Participation/Engagement by dimensions (for heatmaps)
    participationByUnit: {},      // { unitName: { emailRate, meetingRate, count } }
    participationByLocation: {},  // { location: { emailRate, meetingRate, count } }

    // Satisfaction by dimensions
    satisfactionByUnit: {},       // { unitName: { score, count } }
    satisfactionByLocation: {},   // { location: { score, count } }
    satisfactionByOfficeDays: {}, // { day: { score, count } }

    // Member Directory Trends (NEW)
    directoryTrends: {
      recentUpdates: [],      // Members who updated contact info recently
      staleContacts: [],      // Members who haven't updated in 90+ days
      missingEmail: 0,
      missingPhone: 0,
      totalWithEmail: 0,
      totalWithPhone: 0
    },

    // Engagement Metrics (from Member Directory columns Q-W)
    engagement: {
      emailOpenRate: 0,           // Average from OPEN_RATE column (S)
      virtualMeetingRate: 0,      // % with recent virtual meeting (Q)
      inPersonMeetingRate: 0,     // % with recent in-person meeting (R)
      totalVolunteerHours: 0,     // Sum from VOLUNTEER_HOURS column (T)
      unionInterestLocal: 0,      // % interested in local issues (U)
      unionInterestChapter: 0,    // % interested in chapter activities (V)
      unionInterestAllied: 0,     // % interested in allied movements (W)
      surveyResponseRate: 0,
      recentMeetingAttendees: []  // Members who attended recently
    },

    // Bargaining Data (Enhanced)
    step1DenialRate: 0,
    step2DenialRate: 0,
    avgSettlementDays: 0,
    topViolatedArticle: 'N/A',
    articleViolations: {},
    stepProgression: { step1: 0, step2: 0, step3: 0, arb: 0 },
    resolutionByStep: { step1Resolved: 0, step2Resolved: 0, step3Resolved: 0, arbResolved: 0 },
    stepOutcomes: {
      step1: { won: 0, denied: 0, settled: 0, withdrawn: 0, pending: 0 },
      step2: { won: 0, denied: 0, settled: 0, withdrawn: 0, pending: 0 },
      step3: { won: 0, denied: 0, settled: 0, withdrawn: 0, pending: 0 },
      arb: { won: 0, denied: 0, settled: 0, withdrawn: 0, pending: 0 }
    },
    avgDaysAtStep: { step1: 0, step2: 0, step3: 0, arb: 0 },
    managementResponseTime: 0,
    grievancesByCategory: {},
    monthlyFilings: [],
    monthlyResolved: [],  // v4.4.0 - For Filed vs Resolved chart
    recentGrievances: [],  // Last 10 grievances for display
    stepCaseDetails: { step1: [], step2: [], step3: [], arb: [] },  // Cases at each step

    // Chart Data
    statusDistribution: { open: 0, pending: 0, won: 0, denied: 0, settled: 0, withdrawn: 0 },
    monthlyTrend: [],
    sentimentTrend: [],

    // Chart drill-down data (details for each chart segment)
    chartDrillDown: {
      statusByCase: { open: [], pending: [], won: [], denied: [], settled: [], withdrawn: [] },
      locationByCase: {},      // { location: [cases] }
      categoryByCase: {},      // { category: [cases] }
      unitByMember: {},        // { unit: [members] }
      stewardByCase: {}        // { steward: [cases] }
    },

    // Satisfaction Survey Data (Enhanced with individual question scores)
    satisfactionData: {
      responseCount: 0,
      sections: [
        { name: 'Overall Satisfaction', key: 'overall', score: 0, questions: [
          { label: 'Satisfied with Rep', key: 'q6', score: 0 },
          { label: 'Trust Union', key: 'q7', score: 0 },
          { label: 'Feel Protected', key: 'q8', score: 0 },
          { label: 'Recommend to Colleague', key: 'q9', score: 0 }
        ]},
        { name: 'Steward Ratings', key: 'steward', score: 0, questions: [
          { label: 'Timely Response', key: 'q10', score: 0 },
          { label: 'Treated with Respect', key: 'q11', score: 0 },
          { label: 'Explained Options', key: 'q12', score: 0 },
          { label: 'Followed Through', key: 'q13', score: 0 },
          { label: 'Advocated Effectively', key: 'q14', score: 0 },
          { label: 'Safe to Raise Concerns', key: 'q15', score: 0 },
          { label: 'Maintained Confidentiality', key: 'q16', score: 0 }
        ]},
        { name: 'Chapter Effectiveness', key: 'chapter', score: 0, questions: [
          { label: 'Understands Workplace Issues', key: 'q21', score: 0 },
          { label: 'Communicates Effectively', key: 'q22', score: 0 },
          { label: 'Organizes Events/Actions', key: 'q23', score: 0 },
          { label: 'Easy to Reach', key: 'q24', score: 0 },
          { label: 'Fair Representation', key: 'q25', score: 0 }
        ]},
        { name: 'Local Leadership', key: 'leadership', score: 0, questions: [
          { label: 'Decisions Explained Clearly', key: 'q26', score: 0 },
          { label: 'Understand Decision Process', key: 'q27', score: 0 },
          { label: 'Financial Transparency', key: 'q28', score: 0 },
          { label: 'Leadership Accountable', key: 'q29', score: 0 },
          { label: 'Fair Internal Processes', key: 'q30', score: 0 },
          { label: 'Welcomes Member Opinions', key: 'q31', score: 0 }
        ]},
        { name: 'Contract Enforcement', key: 'contract', score: 0, questions: [
          { label: 'Actively Enforces Contract', key: 'q32', score: 0 },
          { label: 'Realistic Timelines', key: 'q33', score: 0 },
          { label: 'Clear Updates on Cases', key: 'q34', score: 0 },
          { label: 'Frontline Workers Priority', key: 'q35', score: 0 }
        ]},
        { name: 'Communication Quality', key: 'communication', score: 0, questions: [
          { label: 'Clear & Actionable', key: 'q41', score: 0 },
          { label: 'Enough Information', key: 'q42', score: 0 },
          { label: 'Easy to Find Info', key: 'q43', score: 0 },
          { label: 'Reaches All Shifts', key: 'q44', score: 0 },
          { label: 'Meetings Worth Attending', key: 'q45', score: 0 }
        ]},
        { name: 'Member Voice', key: 'voice', score: 0, questions: [
          { label: 'Voice Matters', key: 'q46', score: 0 },
          { label: 'Seeks Member Input', key: 'q47', score: 0 },
          { label: 'Treated with Dignity', key: 'q48', score: 0 },
          { label: 'Newer Members Supported', key: 'q49', score: 0 },
          { label: 'Conflicts Handled Respectfully', key: 'q50', score: 0 }
        ]},
        { name: 'Value & Action', key: 'value', score: 0, questions: [
          { label: 'Good Value for Dues', key: 'q51', score: 0 },
          { label: 'Priorities Match Needs', key: 'q52', score: 0 },
          { label: 'Prepared to Mobilize', key: 'q53', score: 0 },
          { label: 'Believe We Can Win', key: 'q55', score: 0 }
        ]}
      ],
      trendByMonth: [],
      questionBreakdown: {},
      allQuestionScores: {}  // Individual question scores
    },

    // Google Drive Resources
    driveResources: {
      folderId: '',
      folderUrl: '',
      recentFiles: []
    },

    // Meeting Notes (for member dashboard - completed meetings with published notes)
    meetingNotes: []
  };

  var now = new Date();
  var ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
  var thirtyDaysAgo = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
  var sixMonthsAgo = new Date(now.getTime() - (180 * 24 * 60 * 60 * 1000));

  // Engagement metric accumulators
  var openRates = [];
  var virtualMeetingCount = 0;
  var inPersonMeetingCount = 0;
  var totalVolunteerHours = 0;
  var interestLocalCount = 0;
  var interestChapterCount = 0;
  var interestAlliedCount = 0;

  // Engagement by unit/location accumulators
  var engagementByUnit = {};    // { unit: { openRates: [], meetings: 0, count: 0 } }
  var engagementByLocation = {}; // { location: { openRates: [], meetings: 0, count: 0 } }

  // Process Members
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    var headers = memberData[0];

    for (var m = 1; m < memberData.length; m++) {
      var memberId = memberData[m][MEMBER_COLS.MEMBER_ID - 1];
      if (!memberId) continue;

      data.totalMembers++;
      var firstName = memberData[m][MEMBER_COLS.FIRST_NAME - 1] || '';
      var lastName = memberData[m][MEMBER_COLS.LAST_NAME - 1] || '';
      var name = (firstName + ' ' + lastName).trim() || 'Unknown';
      var email = memberData[m][MEMBER_COLS.EMAIL - 1] || '';
      var phone = memberData[m][MEMBER_COLS.PHONE - 1] || '';
      var location = memberData[m][MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      var unit = memberData[m][MEMBER_COLS.UNIT - 1] || 'Unknown';
      var isSteward = memberData[m][MEMBER_COLS.IS_STEWARD - 1] === 'Yes';
      var lastUpdated = memberData[m][MEMBER_COLS.RECENT_CONTACT_DATE - 1];

      // Engagement metrics from columns Q-W
      var lastVirtualMtg = memberData[m][MEMBER_COLS.LAST_VIRTUAL_MTG - 1];
      var lastInPersonMtg = memberData[m][MEMBER_COLS.LAST_INPERSON_MTG - 1];
      var openRate = memberData[m][MEMBER_COLS.OPEN_RATE - 1];
      var volunteerHours = memberData[m][MEMBER_COLS.VOLUNTEER_HOURS - 1];
      var interestLocal = memberData[m][MEMBER_COLS.INTEREST_LOCAL - 1];
      var interestChapter = memberData[m][MEMBER_COLS.INTEREST_CHAPTER - 1];
      var interestAllied = memberData[m][MEMBER_COLS.INTEREST_ALLIED - 1];

      // Track email open rates (if numeric, including 0)
      if ((openRate !== '' && openRate !== null && openRate !== undefined) && !isNaN(parseFloat(openRate))) {
        openRates.push(parseFloat(openRate));
      }

      // Track volunteer hours (if numeric, including 0)
      if ((volunteerHours !== '' && volunteerHours !== null && volunteerHours !== undefined) && !isNaN(parseFloat(volunteerHours))) {
        totalVolunteerHours += parseFloat(volunteerHours);
      }

      // Track meeting attendance (within last 6 months)
      if (lastVirtualMtg instanceof Date && lastVirtualMtg > sixMonthsAgo) {
        virtualMeetingCount++;
        if (includePII) {
          data.engagement.recentMeetingAttendees.push({ id: memberId, name: name, type: 'Virtual', date: lastVirtualMtg });
        }
      }
      if (lastInPersonMtg instanceof Date && lastInPersonMtg > sixMonthsAgo) {
        inPersonMeetingCount++;
        if (includePII) {
          data.engagement.recentMeetingAttendees.push({ id: memberId, name: name, type: 'In-Person', date: lastInPersonMtg });
        }
      }

      // Track union interest (Yes/True/true values)
      if (interestLocal && (interestLocal === 'Yes' || interestLocal === true || interestLocal === 'TRUE' || interestLocal === 'true')) {
        interestLocalCount++;
      }
      if (interestChapter && (interestChapter === 'Yes' || interestChapter === true || interestChapter === 'TRUE' || interestChapter === 'true')) {
        interestChapterCount++;
      }
      if (interestAllied && (interestAllied === 'Yes' || interestAllied === true || interestAllied === 'TRUE' || interestAllied === 'true')) {
        interestAlliedCount++;
      }

      // Track office days breakdown
      var officeDays = memberData[m][MEMBER_COLS.OFFICE_DAYS - 1] || '';
      if (officeDays && officeDays !== 'N/A') {
        var days = officeDays.split(',');
        for (var di = 0; di < days.length; di++) {
          var day = days[di].trim();
          if (day) {
            if (!data.officeDaysBreakdown[day]) data.officeDaysBreakdown[day] = 0;
            data.officeDaysBreakdown[day]++;
          }
        }
      }

      // Track engagement by unit
      if (!engagementByUnit[unit]) engagementByUnit[unit] = { openRates: [], meetings: 0, count: 0 };
      engagementByUnit[unit].count++;
      if ((openRate !== '' && openRate !== null && openRate !== undefined) && !isNaN(parseFloat(openRate))) engagementByUnit[unit].openRates.push(parseFloat(openRate));
      if ((lastVirtualMtg instanceof Date && lastVirtualMtg > sixMonthsAgo) ||
          (lastInPersonMtg instanceof Date && lastInPersonMtg > sixMonthsAgo)) {
        engagementByUnit[unit].meetings++;
      }

      // Track engagement by location
      if (!engagementByLocation[location]) engagementByLocation[location] = { openRates: [], meetings: 0, count: 0 };
      engagementByLocation[location].count++;
      if ((openRate !== '' && openRate !== null && openRate !== undefined) && !isNaN(parseFloat(openRate))) engagementByLocation[location].openRates.push(parseFloat(openRate));
      if ((lastVirtualMtg instanceof Date && lastVirtualMtg > sixMonthsAgo) ||
          (lastInPersonMtg instanceof Date && lastInPersonMtg > sixMonthsAgo)) {
        engagementByLocation[location].meetings++;
      }

      // Count stewards (steward PII is always shown - they're public union reps)
      if (isSteward) {
        data.stewardCount++;
        data.stewardList.push({ id: memberId, name: name, location: location, unit: unit, email: email, phone: phone });
      }

      // Member list for drill-down (members hidden in non-PII mode)
      if (includePII) {
        data.memberList.push({ id: memberId, name: name, location: location, unit: unit, isSteward: isSteward });
      }

      // Location and unit breakdowns
      if (!data.locationBreakdown[location]) data.locationBreakdown[location] = 0;
      data.locationBreakdown[location]++;
      if (!data.unitBreakdown[unit]) data.unitBreakdown[unit] = 0;
      data.unitBreakdown[unit]++;

      // Unit drill-down by member
      if (!data.chartDrillDown.unitByMember[unit]) data.chartDrillDown.unitByMember[unit] = [];
      data.chartDrillDown.unitByMember[unit].push({
        id: includePII ? memberId : memberId.substring(0, 4) + '***',
        name: includePII ? name : 'Member',
        location: location,
        isSteward: isSteward
      });

      // Contact info tracking
      if (email) data.directoryTrends.totalWithEmail++;
      else data.directoryTrends.missingEmail++;
      if (phone) data.directoryTrends.totalWithPhone++;
      else data.directoryTrends.missingPhone++;

      // Recent updates vs stale contacts
      if (lastUpdated instanceof Date) {
        if (lastUpdated > thirtyDaysAgo) {
          if (includePII) {
            data.directoryTrends.recentUpdates.push({ id: memberId, name: name, date: lastUpdated });
          } else {
            data.directoryTrends.recentUpdates.push({ id: memberId.substring(0, 4) + '***', name: 'Member', date: lastUpdated });
          }
        }
        if (lastUpdated < ninetyDaysAgo) {
          if (includePII) {
            data.directoryTrends.staleContacts.push({ id: memberId, name: name, lastUpdate: lastUpdated });
          } else {
            data.directoryTrends.staleContacts.push({ id: memberId.substring(0, 4) + '***', name: 'Member', lastUpdate: lastUpdated });
          }
        }
      }
    }

    // Calculate engagement metrics after processing all members
    if (data.totalMembers > 0) {
      data.engagement.emailOpenRate = openRates.length > 0
        ? Math.round((openRates.reduce(function(a,b){return a+b;},0) / openRates.length) * 10) / 10
        : 0;
      data.engagement.virtualMeetingRate = Math.round((virtualMeetingCount / data.totalMembers) * 100);
      data.engagement.inPersonMeetingRate = Math.round((inPersonMeetingCount / data.totalMembers) * 100);
      data.engagement.totalVolunteerHours = Math.round(totalVolunteerHours);
      data.engagement.unionInterestLocal = Math.round((interestLocalCount / data.totalMembers) * 100);
      data.engagement.unionInterestChapter = Math.round((interestChapterCount / data.totalMembers) * 100);
      data.engagement.unionInterestAllied = Math.round((interestAlliedCount / data.totalMembers) * 100);

      // Calculate Steward:Member ratio
      if (data.stewardCount > 0) {
        var ratio = Math.round(data.totalMembers / data.stewardCount);
        data.stewardRatio = ratio + ':1';
      }

      // Calculate participation rates by unit
      for (var unitKey in engagementByUnit) {
        var u = engagementByUnit[unitKey];
        var avgRate = u.openRates.length > 0 ? Math.round(u.openRates.reduce(function(a,b){return a+b;},0) / u.openRates.length) : 0;
        var meetingPct = u.count > 0 ? Math.round((u.meetings / u.count) * 100) : 0;
        data.participationByUnit[unitKey] = { emailRate: avgRate, meetingRate: meetingPct, count: u.count };
      }

      // Calculate participation rates by location
      for (var locKey in engagementByLocation) {
        var loc = engagementByLocation[locKey];
        var avgLocRate = loc.openRates.length > 0 ? Math.round(loc.openRates.reduce(function(a,b){return a+b;},0) / loc.openRates.length) : 0;
        var locMeetingPct = loc.count > 0 ? Math.round((loc.meetings / loc.count) * 100) : 0;
        data.participationByLocation[locKey] = { emailRate: avgLocRate, meetingRate: locMeetingPct, count: loc.count };
      }
    }
  }

  // Get current user email for My Cases
  try {
    data.currentUserEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    data.currentUserEmail = '';
  }

  // Process Grievances
  var stewardCases = {};
  var locationCases = {};
  var locationOverdue = {};  // Track overdue by location
  var step1Total = 0, step1Denials = 0;
  var step2Total = 0, step2Denials = 0;
  var settlementDays = [];
  var stepDays = { step1: [], step2: [], step3: [], arb: [] };  // Days at each step
  var mgmtResponseDays = [];
  var monthlyFilingsMap = {};
  var monthlyResolvedMap = {};  // v4.4.0 - Track resolved by month

  // Previous period tracking (30-60 days ago vs current 30 days)
  var prevPeriodStart = new Date(now.getTime() - (60 * 24 * 60 * 60 * 1000));
  var prevPeriodEnd = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
  var prevOpenCases = 0, prevWins = 0, prevLosses = 0, prevSettled = 0, prevOverdue = 0;

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    for (var g = 1; g < grievanceData.length; g++) {
      var grievanceId = grievanceData[g][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      if (!grievanceId) continue;

      data.totalGrievances++;
      var status = (grievanceData[g][GRIEVANCE_COLS.STATUS - 1] || '').toString();
      var steward = grievanceData[g][GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned';
      var gLocation = grievanceData[g][GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
      var article = grievanceData[g][GRIEVANCE_COLS.ARTICLES - 1];
      var category = grievanceData[g][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';
      var gFirstName = grievanceData[g][GRIEVANCE_COLS.FIRST_NAME - 1] || '';
      var gLastName = grievanceData[g][GRIEVANCE_COLS.LAST_NAME - 1] || '';
      var memberName = (gFirstName + ' ' + gLastName).trim() || 'Unknown';
      var currentStep = grievanceData[g][GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step 1';
      var dateFiled = grievanceData[g][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = grievanceData[g][GRIEVANCE_COLS.DATE_CLOSED - 1];

      // Create case summary for drill-down
      var caseSummary = {
        id: grievanceId,
        member: includePII ? memberName : 'Member',
        steward: steward,
        location: gLocation,
        category: category,
        step: currentStep,
        dateFiled: dateFiled instanceof Date ? dateFiled.toLocaleDateString() : '',
        dateClosed: dateClosed instanceof Date ? dateClosed.toLocaleDateString() : ''
      };

      // Track previous period metrics (cases active 30-60 days ago)
      var wasOpenInPrevPeriod = dateFiled instanceof Date && dateFiled <= prevPeriodEnd &&
        (!dateClosed || (dateClosed instanceof Date && dateClosed > prevPeriodEnd));
      var closedInPrevPeriod = dateClosed instanceof Date && dateClosed >= prevPeriodStart && dateClosed <= prevPeriodEnd;

      // Status distribution with drill-down data
      var statusKey = '';
      switch(status.toLowerCase()) {
        case 'open':
          data.statusDistribution.open++; data.openGrievances++; statusKey = 'open';
          if (wasOpenInPrevPeriod) prevOpenCases++;
          break;
        case 'pending info': case 'pending':
          data.statusDistribution.pending++; data.openGrievances++; statusKey = 'pending';
          if (wasOpenInPrevPeriod) prevOpenCases++;
          break;
        case 'won': case 'sustained': case 'favorable':
          data.statusDistribution.won++; data.wins++; statusKey = 'won';
          if (closedInPrevPeriod) prevWins++;
          break;
        case 'denied': case 'lost':
          data.statusDistribution.denied++; data.losses++; statusKey = 'denied';
          if (closedInPrevPeriod) prevLosses++;
          break;
        case 'settled':
          data.statusDistribution.settled++; data.settled++; statusKey = 'settled';
          if (closedInPrevPeriod) prevSettled++;
          break;
        case 'withdrawn':
          data.statusDistribution.withdrawn++; statusKey = 'withdrawn'; break;
      }

      // Add to drill-down arrays
      if (statusKey && data.chartDrillDown.statusByCase[statusKey]) {
        data.chartDrillDown.statusByCase[statusKey].push(caseSummary);
      }

      // Location drill-down
      if (!data.chartDrillDown.locationByCase[gLocation]) data.chartDrillDown.locationByCase[gLocation] = [];
      data.chartDrillDown.locationByCase[gLocation].push(caseSummary);

      // Steward drill-down
      if (!data.chartDrillDown.stewardByCase[steward]) data.chartDrillDown.stewardByCase[steward] = [];
      data.chartDrillDown.stewardByCase[steward].push(caseSummary);

      // Category breakdown and drill-down
      if (!data.grievancesByCategory[category]) data.grievancesByCategory[category] = 0;
      data.grievancesByCategory[category]++;
      if (!data.chartDrillDown.categoryByCase[category]) data.chartDrillDown.categoryByCase[category] = [];
      data.chartDrillDown.categoryByCase[category].push(caseSummary);

      // Step progression with detailed tracking
      var statusLower = status.toLowerCase();
      var currentStepKey = null;
      if (currentStep.toLowerCase().includes('step 1')) {
        data.stepProgression.step1++;
        currentStepKey = 'step1';
      } else if (currentStep.toLowerCase().includes('step 2')) {
        data.stepProgression.step2++;
        currentStepKey = 'step2';
      } else if (currentStep.toLowerCase().includes('step 3')) {
        data.stepProgression.step3++;
        currentStepKey = 'step3';
      } else if (currentStep.toLowerCase().includes('arb')) {
        data.stepProgression.arb++;
        currentStepKey = 'arb';
      }

      // Track outcomes at each step
      if (currentStepKey) {
        var displayMemberStep = includePII ? memberName : 'Member';
        var caseDetail = { id: grievanceId, member: displayMemberStep, category: category, location: gLocation, status: status, daysOpen: grievanceData[g][GRIEVANCE_COLS.DAYS_OPEN - 1] || 0 };
        data.stepCaseDetails[currentStepKey].push(caseDetail);

        if (statusLower === 'won') data.stepOutcomes[currentStepKey].won++;
        else if (statusLower === 'denied') data.stepOutcomes[currentStepKey].denied++;
        else if (statusLower === 'settled') data.stepOutcomes[currentStepKey].settled++;
        else if (statusLower === 'withdrawn') data.stepOutcomes[currentStepKey].withdrawn++;
        else data.stepOutcomes[currentStepKey].pending++;
      }

      // Track Step 2 denial rate
      if (grievanceData[g][GRIEVANCE_COLS.STEP2_RCVD - 1]) {
        step2Total++;
        if (status !== 'Won' && grievanceData[g][GRIEVANCE_COLS.STEP3_APPEAL_FILED - 1]) step2Denials++;
      }

      // Open cases list for drill-down (member hidden in non-PII, steward always shown)
      if (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info' || status.toLowerCase() === 'pending') {
        var displayMember = includePII ? memberName : 'Member';
        data.openCasesList.push({ id: grievanceId, member: displayMember, steward: steward, step: currentStep, location: gLocation, category: category });

        // My Cases - check if current user is the assigned steward (v4.4.0)
        if (includePII && data.currentUserEmail && steward) {
          var stewardLower = steward.toLowerCase();
          var emailLower = data.currentUserEmail.toLowerCase();
          var emailName = emailLower.split('@')[0].replace(/[._]/g, ' ');
          // Match by full steward name, email prefix, or partial match
          if (stewardLower.indexOf(emailName) >= 0 || emailLower.indexOf(stewardLower.replace(/\s+/g, '.')) >= 0 ||
              stewardLower.split(' ').some(function(part) { return emailName.indexOf(part) >= 0 && part.length > 2; })) {
            data.myCases.push({
              id: grievanceId,
              member: memberName,
              step: currentStep,
              location: gLocation,
              category: category,
              status: status,
              daysOpen: grievanceData[g][GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
              nextDue: grievanceData[g][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] || ''
            });
          }
        }
      }

      // Steward workload
      if (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info' || status.toLowerCase() === 'pending') {
        if (!stewardCases[steward]) stewardCases[steward] = 0;
        stewardCases[steward]++;
      }

      // Hot zones (locations with active cases)
      if (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info' || status.toLowerCase() === 'pending') {
        if (!locationCases[gLocation]) locationCases[gLocation] = 0;
        locationCases[gLocation]++;
      }

      // Article violations
      if (article) {
        var articles = article.toString().split(',');
        articles.forEach(function(art) {
          art = art.trim();
          if (art) {
            if (!data.articleViolations[art]) data.articleViolations[art] = 0;
            data.articleViolations[art]++;
          }
        });
      }

      // Step 1 denial rate
      if (grievanceData[g][GRIEVANCE_COLS.STEP1_RCVD - 1]) {
        step1Total++;
        if (status !== 'Won' && grievanceData[g][GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1]) step1Denials++;
      }

      // Settlement time
      if (dateFiled instanceof Date && dateClosed instanceof Date) {
        var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
        if (days > 0) settlementDays.push(days);
      }

      // Monthly filings
      if (dateFiled instanceof Date) {
        var monthKey = dateFiled.toLocaleString('default', { month: 'short', year: 'numeric' });
        if (!monthlyFilingsMap[monthKey]) monthlyFilingsMap[monthKey] = 0;
        monthlyFilingsMap[monthKey]++;
      }

      // Monthly resolved (v4.4.0)
      if (dateClosed instanceof Date) {
        var resolvedMonthKey = dateClosed.toLocaleString('default', { month: 'short', year: 'numeric' });
        if (!monthlyResolvedMap[resolvedMonthKey]) monthlyResolvedMap[resolvedMonthKey] = 0;
        monthlyResolvedMap[resolvedMonthKey]++;
      }

      // Check overdue
      var step1Due = grievanceData[g][GRIEVANCE_COLS.STEP1_DUE - 1];
      var daysToDeadline = grievanceData[g][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
      var isOverdue = (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0)) ||
                      (step1Due && new Date(step1Due) < now && (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info'));

      if (isOverdue && (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info' || status.toLowerCase() === 'pending')) {
        data.overdueCount++;
        var displayMemberOverdue = includePII ? memberName : 'Member';
        data.overdueList.push({ id: grievanceId, member: displayMemberOverdue, steward: steward, step: currentStep, dueDate: step1Due });
        // Track overdue by location for hot spots
        if (!locationOverdue[gLocation]) locationOverdue[gLocation] = 0;
        locationOverdue[gLocation]++;
      }

      // Track recent grievances for display (last 10)
      if (data.recentGrievances.length < 10) {
        var displayMemberRecent = includePII ? memberName : 'Member';
        data.recentGrievances.push({ id: grievanceId, member: displayMemberRecent, category: category, status: status, step: currentStep, location: gLocation, dateFiled: dateFiled });
      }
    }
  }

  // Calculate derived metrics
  var totalClosed = data.wins + data.losses + data.settled;
  data.winRate = totalClosed > 0 ? Math.round((data.wins / totalClosed) * 100) : 0;
  data.step1DenialRate = step1Total > 0 ? Math.round((step1Denials / step1Total) * 100) : 0;
  data.step2DenialRate = step2Total > 0 ? Math.round((step2Denials / step2Total) * 100) : 0;
  data.avgSettlementDays = settlementDays.length > 0 ? Math.round(settlementDays.reduce(function(a,b){return a+b;},0) / settlementDays.length) : 0;

  // Calculate previous period comparison metrics
  data.prevOpenCases = prevOpenCases;
  var prevTotalClosed = prevWins + prevLosses + prevSettled;
  data.prevWinRate = prevTotalClosed > 0 ? Math.round((prevWins / prevTotalClosed) * 100) : data.winRate;
  data.prevOverdueCount = prevOverdue;
  data.prevTotalGrievances = data.totalGrievances; // Approximation - would need historical data for accurate count

  // Build steward workload array
  for (var s in stewardCases) {
    var count = stewardCases[s];
    var statusLabel = count > 8 ? 'OVERLOAD' : count > 5 ? 'Heavy' : 'Available';
    var color = count > 8 ? '#ef4444' : count > 5 ? '#f59e0b' : '#22c55e';
    // Steward names always shown (they're public union reps)
    data.stewardWorkload.push({ name: s, count: count, status: statusLabel, color: color });
  }
  data.stewardWorkload.sort(function(a,b){return b.count - a.count;});

  // Build hot zones (locations with 3+ active cases)
  for (var loc in locationCases) {
    if (locationCases[loc] >= 3) {
      data.hotZones.push({ location: loc, count: locationCases[loc] });
      data.hotSpots.grievance.push({ name: loc, type: 'location', count: locationCases[loc], reason: 'High grievance activity' });
    }
  }
  data.hotZones.sort(function(a,b){return b.count - a.count;});

  // Build overdue concentration hot spots (locations with 2+ overdue)
  for (var overdLoc in locationOverdue) {
    if (locationOverdue[overdLoc] >= 2) {
      data.hotSpots.overdueConcentration.push({ name: overdLoc, type: 'location', count: locationOverdue[overdLoc], reason: 'Multiple overdue cases' });
    }
  }
  data.hotSpots.overdueConcentration.sort(function(a,b){return b.count - a.count;});

  // Top violated article
  var maxViolations = 0;
  for (var art in data.articleViolations) {
    if (data.articleViolations[art] > maxViolations) {
      maxViolations = data.articleViolations[art];
      data.topViolatedArticle = art;
    }
  }

  // Get Top Performers from hidden steward performance sheet (v4.4.0)
  var perfSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
  if (perfSheet && perfSheet.getLastRow() > 1) {
    try {
      var perfData = perfSheet.getRange(2, 1, Math.min(perfSheet.getLastRow() - 1, 20), 10).getValues();
      data.topPerformers = perfData
        .filter(function(row) { return row[0] && row[9]; })
        .map(function(row) {
          return {
            name: row[0],  // Steward names always shown (public union reps)
            totalCases: row[1] || 0,
            active: row[2] || 0,
            closed: row[3] || 0,
            won: row[4] || 0,
            winRate: row[5] || 0,
            avgDays: row[6] || 0,
            score: row[9] || 0
          };
        })
        .sort(function(a, b) { return b.score - a.score; })
        .slice(0, 5);
    } catch (e) {
      Logger.log('Error reading steward performance: ' + e.message);
    }
  }

  // Monthly filings for trend
  var allMonths = Object.keys(monthlyFilingsMap).concat(Object.keys(monthlyResolvedMap));
  var uniqueMonths = allMonths.filter(function(item, pos) { return allMonths.indexOf(item) === pos; });
  var sortedMonths = uniqueMonths.sort(function(a, b) {
    return new Date(a) - new Date(b);
  }).slice(-12);
  sortedMonths.forEach(function(month) {
    data.monthlyFilings.push({ month: month, count: monthlyFilingsMap[month] || 0 });
    data.monthlyResolved.push({ month: month, count: monthlyResolvedMap[month] || 0 });  // v4.4.0
  });

  // Process Satisfaction Survey
  if (satSheet && satSheet.getLastRow() > 1) {
    var satData = satSheet.getDataRange().getValues();
    var trustScores = [];
    var monthlyTrust = {};
    var sectionScores = { overall: [], steward: [], chapter: [], leadership: [], contract: [], communication: [], voice: [], value: [] };
    var satByWorksite = {};  // { worksite: { scores: [], count: 0 } }
    var satByRole = {};      // { role: { scores: [], count: 0 } }

    // Individual question score accumulators
    var questionScores = {
      q6: [], q7: [], q8: [], q9: [],  // Overall
      q10: [], q11: [], q12: [], q13: [], q14: [], q15: [], q16: [],  // Steward
      q21: [], q22: [], q23: [], q24: [], q25: [],  // Chapter
      q26: [], q27: [], q28: [], q29: [], q30: [], q31: [],  // Leadership
      q32: [], q33: [], q34: [], q35: [],  // Contract
      q41: [], q42: [], q43: [], q44: [], q45: [],  // Communication
      q46: [], q47: [], q48: [], q49: [], q50: [],  // Voice
      q51: [], q52: [], q53: [], q55: []  // Value
    };

    for (var i = 1; i < satData.length; i++) {
      if (!satData[i][0]) continue;
      data.satisfactionData.responseCount++;

      // Get worksite and role for breakdown
      var worksite = (satData[i][SATISFACTION_COLS.Q1_WORKSITE - 1] || 'Unknown').toString().trim();
      var role = (satData[i][SATISFACTION_COLS.Q2_ROLE - 1] || 'Unknown').toString().trim();

      var trustVal = parseFloat(satData[i][7]);
      var timestamp = satData[i][0];

      if (!isNaN(trustVal) && trustVal >= 1 && trustVal <= 10) {
        trustScores.push(trustVal);
        if (timestamp) {
          var date = new Date(timestamp);
          var monthKey = date.toLocaleString('default', { month: 'short' });
          if (!monthlyTrust[monthKey]) monthlyTrust[monthKey] = { sum: 0, count: 0 };
          monthlyTrust[monthKey].sum += trustVal;
          monthlyTrust[monthKey].count++;
        }
      }

      // Collect individual question scores (columns are 0-indexed, so Q6 is column 6)
      var qVal;
      // Overall Satisfaction questions
      qVal = parseFloat(satData[i][6]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q6.push(qVal);
      qVal = parseFloat(satData[i][7]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q7.push(qVal);
      qVal = parseFloat(satData[i][8]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q8.push(qVal);
      qVal = parseFloat(satData[i][9]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q9.push(qVal);
      // Steward questions
      qVal = parseFloat(satData[i][10]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q10.push(qVal);
      qVal = parseFloat(satData[i][11]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q11.push(qVal);
      qVal = parseFloat(satData[i][12]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q12.push(qVal);
      qVal = parseFloat(satData[i][13]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q13.push(qVal);
      qVal = parseFloat(satData[i][14]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q14.push(qVal);
      qVal = parseFloat(satData[i][15]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q15.push(qVal);
      qVal = parseFloat(satData[i][16]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q16.push(qVal);
      // Chapter questions
      qVal = parseFloat(satData[i][21]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q21.push(qVal);
      qVal = parseFloat(satData[i][22]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q22.push(qVal);
      qVal = parseFloat(satData[i][23]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q23.push(qVal);
      qVal = parseFloat(satData[i][24]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q24.push(qVal);
      qVal = parseFloat(satData[i][25]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q25.push(qVal);
      // Leadership questions
      qVal = parseFloat(satData[i][26]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q26.push(qVal);
      qVal = parseFloat(satData[i][27]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q27.push(qVal);
      qVal = parseFloat(satData[i][28]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q28.push(qVal);
      qVal = parseFloat(satData[i][29]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q29.push(qVal);
      qVal = parseFloat(satData[i][30]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q30.push(qVal);
      qVal = parseFloat(satData[i][31]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q31.push(qVal);
      // Contract questions
      qVal = parseFloat(satData[i][32]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q32.push(qVal);
      qVal = parseFloat(satData[i][33]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q33.push(qVal);
      qVal = parseFloat(satData[i][34]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q34.push(qVal);
      qVal = parseFloat(satData[i][35]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q35.push(qVal);
      // Communication questions
      qVal = parseFloat(satData[i][41]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q41.push(qVal);
      qVal = parseFloat(satData[i][42]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q42.push(qVal);
      qVal = parseFloat(satData[i][43]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q43.push(qVal);
      qVal = parseFloat(satData[i][44]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q44.push(qVal);
      qVal = parseFloat(satData[i][45]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q45.push(qVal);
      // Voice questions
      qVal = parseFloat(satData[i][46]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q46.push(qVal);
      qVal = parseFloat(satData[i][47]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q47.push(qVal);
      qVal = parseFloat(satData[i][48]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q48.push(qVal);
      qVal = parseFloat(satData[i][49]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q49.push(qVal);
      qVal = parseFloat(satData[i][50]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q50.push(qVal);
      // Value questions
      qVal = parseFloat(satData[i][51]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q51.push(qVal);
      qVal = parseFloat(satData[i][52]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q52.push(qVal);
      qVal = parseFloat(satData[i][53]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q53.push(qVal);
      qVal = parseFloat(satData[i][55]); if (!isNaN(qVal) && qVal >= 1 && qVal <= 10) questionScores.q55.push(qVal);

      // Section averages
      var avgOverall = parseFloat(satData[i][71]) || 0;
      var avgSteward = parseFloat(satData[i][72]) || 0;
      var avgChapter = parseFloat(satData[i][74]) || 0;
      var avgLeadership = parseFloat(satData[i][75]) || 0;
      var avgContract = parseFloat(satData[i][76]) || 0;
      var avgComm = parseFloat(satData[i][78]) || 0;
      var avgVoice = parseFloat(satData[i][79]) || 0;
      var avgValue = parseFloat(satData[i][80]) || 0;

      if (avgOverall === 0) {
        var q6 = parseFloat(satData[i][6]) || 0, q7 = parseFloat(satData[i][7]) || 0;
        var q8 = parseFloat(satData[i][8]) || 0, q9 = parseFloat(satData[i][9]) || 0;
        avgOverall = (q6 + q7 + q8 + q9) / 4;
      }

      if (avgOverall > 0) sectionScores.overall.push(avgOverall);
      if (avgSteward > 0) sectionScores.steward.push(avgSteward);
      if (avgChapter > 0) sectionScores.chapter.push(avgChapter);
      if (avgLeadership > 0) sectionScores.leadership.push(avgLeadership);
      if (avgContract > 0) sectionScores.contract.push(avgContract);
      if (avgComm > 0) sectionScores.communication.push(avgComm);
      if (avgVoice > 0) sectionScores.voice.push(avgVoice);
      if (avgValue > 0) sectionScores.value.push(avgValue);

      // Track satisfaction by worksite (location)
      if (worksite && avgOverall > 0) {
        if (!satByWorksite[worksite]) satByWorksite[worksite] = { scores: [], count: 0 };
        satByWorksite[worksite].scores.push(avgOverall);
        satByWorksite[worksite].count++;
      }

      // Track satisfaction by role (similar to unit)
      if (role && avgOverall > 0) {
        if (!satByRole[role]) satByRole[role] = { scores: [], count: 0 };
        satByRole[role].scores.push(avgOverall);
        satByRole[role].count++;
      }
    }

    function avg(arr) { return arr.length > 0 ? Math.round((arr.reduce(function(a,b){return a+b;},0) / arr.length) * 10) / 10 : 0; }
    data.satisfactionData.sections[0].score = avg(sectionScores.overall);
    data.satisfactionData.sections[1].score = avg(sectionScores.steward);
    data.satisfactionData.sections[2].score = avg(sectionScores.chapter);
    data.satisfactionData.sections[3].score = avg(sectionScores.leadership);
    data.satisfactionData.sections[4].score = avg(sectionScores.contract);
    data.satisfactionData.sections[5].score = avg(sectionScores.communication);
    data.satisfactionData.sections[6].score = avg(sectionScores.voice);
    data.satisfactionData.sections[7].score = avg(sectionScores.value);

    // Calculate individual question scores and populate section questions
    data.satisfactionData.allQuestionScores = {};
    for (var qKey in questionScores) {
      data.satisfactionData.allQuestionScores[qKey] = avg(questionScores[qKey]);
    }
    // Populate scores in sections
    data.satisfactionData.sections[0].questions[0].score = avg(questionScores.q6);
    data.satisfactionData.sections[0].questions[1].score = avg(questionScores.q7);
    data.satisfactionData.sections[0].questions[2].score = avg(questionScores.q8);
    data.satisfactionData.sections[0].questions[3].score = avg(questionScores.q9);
    data.satisfactionData.sections[1].questions[0].score = avg(questionScores.q10);
    data.satisfactionData.sections[1].questions[1].score = avg(questionScores.q11);
    data.satisfactionData.sections[1].questions[2].score = avg(questionScores.q12);
    data.satisfactionData.sections[1].questions[3].score = avg(questionScores.q13);
    data.satisfactionData.sections[1].questions[4].score = avg(questionScores.q14);
    data.satisfactionData.sections[1].questions[5].score = avg(questionScores.q15);
    data.satisfactionData.sections[1].questions[6].score = avg(questionScores.q16);
    data.satisfactionData.sections[2].questions[0].score = avg(questionScores.q21);
    data.satisfactionData.sections[2].questions[1].score = avg(questionScores.q22);
    data.satisfactionData.sections[2].questions[2].score = avg(questionScores.q23);
    data.satisfactionData.sections[2].questions[3].score = avg(questionScores.q24);
    data.satisfactionData.sections[2].questions[4].score = avg(questionScores.q25);
    data.satisfactionData.sections[3].questions[0].score = avg(questionScores.q26);
    data.satisfactionData.sections[3].questions[1].score = avg(questionScores.q27);
    data.satisfactionData.sections[3].questions[2].score = avg(questionScores.q28);
    data.satisfactionData.sections[3].questions[3].score = avg(questionScores.q29);
    data.satisfactionData.sections[3].questions[4].score = avg(questionScores.q30);
    data.satisfactionData.sections[3].questions[5].score = avg(questionScores.q31);
    data.satisfactionData.sections[4].questions[0].score = avg(questionScores.q32);
    data.satisfactionData.sections[4].questions[1].score = avg(questionScores.q33);
    data.satisfactionData.sections[4].questions[2].score = avg(questionScores.q34);
    data.satisfactionData.sections[4].questions[3].score = avg(questionScores.q35);
    data.satisfactionData.sections[5].questions[0].score = avg(questionScores.q41);
    data.satisfactionData.sections[5].questions[1].score = avg(questionScores.q42);
    data.satisfactionData.sections[5].questions[2].score = avg(questionScores.q43);
    data.satisfactionData.sections[5].questions[3].score = avg(questionScores.q44);
    data.satisfactionData.sections[5].questions[4].score = avg(questionScores.q45);
    data.satisfactionData.sections[6].questions[0].score = avg(questionScores.q46);
    data.satisfactionData.sections[6].questions[1].score = avg(questionScores.q47);
    data.satisfactionData.sections[6].questions[2].score = avg(questionScores.q48);
    data.satisfactionData.sections[6].questions[3].score = avg(questionScores.q49);
    data.satisfactionData.sections[6].questions[4].score = avg(questionScores.q50);
    data.satisfactionData.sections[7].questions[0].score = avg(questionScores.q51);
    data.satisfactionData.sections[7].questions[1].score = avg(questionScores.q52);
    data.satisfactionData.sections[7].questions[2].score = avg(questionScores.q53);
    data.satisfactionData.sections[7].questions[3].score = avg(questionScores.q55);

    if (trustScores.length > 0) {
      data.moraleScore = Math.round((trustScores.reduce(function(a,b){return a+b;},0) / trustScores.length) * 10) / 10;
    }

    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    months.forEach(function(month) {
      if (monthlyTrust[month]) {
        data.sentimentTrend.push({ month: month, score: Math.round((monthlyTrust[month].sum / monthlyTrust[month].count) * 10) / 10 });
      }
    });

    // Engagement: Survey response rate
    data.engagement.surveyResponseRate = data.totalMembers > 0 ? Math.round((data.satisfactionData.responseCount / data.totalMembers) * 100) : 0;

    // Calculate satisfaction by worksite (maps to location)
    for (var ws in satByWorksite) {
      var wsScore = avg(satByWorksite[ws].scores);
      data.satisfactionByLocation[ws] = {
        score: wsScore,
        count: satByWorksite[ws].count
      };
      // Track dissatisfaction hot spots (score < 5)
      if (wsScore > 0 && wsScore < 5) {
        data.hotSpots.dissatisfaction.push({ name: ws, type: 'location', score: wsScore, count: satByWorksite[ws].count, reason: 'Low satisfaction score' });
      }
    }
    data.hotSpots.dissatisfaction.sort(function(a,b){return a.score - b.score;});

    // Calculate satisfaction by role (maps to unit)
    for (var rl in satByRole) {
      data.satisfactionByUnit[rl] = {
        score: avg(satByRole[rl].scores),
        count: satByRole[rl].count
      };
    }
  }

  // Build low engagement hot spots from participation data
  for (var engLoc in data.participationByLocation) {
    var engData = data.participationByLocation[engLoc];
    var avgEngagement = (engData.emailRate + engData.meetingRate) / 2;
    if (avgEngagement < 30 && engData.count >= 5) {  // Less than 30% engagement, at least 5 members
      data.hotSpots.lowEngagement.push({ name: engLoc, type: 'location', engagement: Math.round(avgEngagement), count: engData.count, reason: 'Low member engagement' });
    }
  }
  for (var engUnit in data.participationByUnit) {
    var engUnitData = data.participationByUnit[engUnit];
    var avgUnitEngagement = (engUnitData.emailRate + engUnitData.meetingRate) / 2;
    if (avgUnitEngagement < 30 && engUnitData.count >= 5) {
      data.hotSpots.lowEngagement.push({ name: engUnit, type: 'unit', engagement: Math.round(avgUnitEngagement), count: engUnitData.count, reason: 'Low member engagement' });
    }
  }
  data.hotSpots.lowEngagement.sort(function(a,b){return a.engagement - b.engagement;});

  // Get Google Drive resources folder from Config tab (column AU / 47)
  try {
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    var archiveFolderId = '';

    // Try to read from Config tab first
    if (configSheet) {
      archiveFolderId = configSheet.getRange(3, CONFIG_COLS.ARCHIVE_FOLDER_ID).getValue();
      if (archiveFolderId) {
        archiveFolderId = String(archiveFolderId).trim();
      }
    }

    // Fallback to COMMAND_CONFIG if not in Config sheet
    if (!archiveFolderId && COMMAND_CONFIG && COMMAND_CONFIG.ARCHIVE_FOLDER_ID) {
      archiveFolderId = COMMAND_CONFIG.ARCHIVE_FOLDER_ID;
    }

    if (archiveFolderId) {
      data.driveResources.folderId = archiveFolderId;
      data.driveResources.folderUrl = 'https://drive.google.com/drive/folders/' + archiveFolderId;

      var folder = DriveApp.getFolderById(archiveFolderId);
      var files = folder.getFiles();
      var fileCount = 0;
      while (files.hasNext() && fileCount < 10) {
        var file = files.next();
        data.driveResources.recentFiles.push({
          name: file.getName(),
          url: file.getUrl(),
          type: file.getMimeType(),
          updated: file.getLastUpdated()
        });
        fileCount++;
      }
    }
  } catch (e) {
    Logger.log('Error accessing Drive resources: ' + e.message);
  }

  // Populate Meeting Notes (completed meetings with notes docs, chronological order)
  try {
    var meetingSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (meetingSheet && meetingSheet.getLastRow() > 1) {
      var meetingData = meetingSheet.getDataRange().getValues();
      var seenMeetings = {};
      var yesterday = new Date(now.getTime() - (24 * 60 * 60 * 1000));
      yesterday.setHours(0, 0, 0, 0);

      for (var mi = 1; mi < meetingData.length; mi++) {
        var mtgId = String(meetingData[mi][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
        if (!mtgId || seenMeetings[mtgId]) continue;
        seenMeetings[mtgId] = true;

        var mtgDate = meetingData[mi][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
        if (!(mtgDate instanceof Date)) continue;

        var mtgDay = new Date(mtgDate);
        mtgDay.setHours(0, 0, 0, 0);

        // Only show notes for past meetings (day after meeting = available to members)
        if (mtgDay > yesterday) continue;

        var notesUrl = String(meetingData[mi][MEETING_CHECKIN_COLS.NOTES_DOC_URL - 1] || '');
        if (!notesUrl) continue;

        data.meetingNotes.push({
          id: mtgId,
          name: String(meetingData[mi][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || ''),
          date: mtgDate.toLocaleDateString(),
          dateSort: mtgDate.getTime(),
          type: String(meetingData[mi][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || ''),
          notesUrl: notesUrl
        });
      }

      // Sort chronologically (newest first)
      data.meetingNotes.sort(function(a, b) { return b.dateSort - a.dateSort; });
    }
  } catch (e) {
    Logger.log('Error loading meeting notes: ' + e.message);
  }

  return JSON.stringify(data);
}

/**
 * API function for web app to get dashboard data
 * @param {boolean} isPII - Include PII (steward mode)
 * @returns {string} JSON dashboard data
 */
function getUnifiedDashboardDataAPI(isPII) {
  return getUnifiedDashboardData(isPII === true || isPII === 'true');
}

/**
 * API function with date range filtering support
 * @param {boolean} isPII - Include PII (steward mode)
 * @param {number} days - Number of days to filter (0 = all data)
 * @param {string} fromDate - Custom start date (ISO string)
 * @param {string} toDate - Custom end date (ISO string)
 * @returns {string} JSON dashboard data filtered by date range
 */
function getUnifiedDashboardDataWithDateRange(isPII, days, fromDate, toDate) {
  var fullData = JSON.parse(getUnifiedDashboardData(isPII === true || isPII === 'true'));

  // If no filtering requested, return full data
  if (!days && !fromDate) {
    return JSON.stringify(fullData);
  }

  var now = new Date();
  var startDate, endDate;

  if (fromDate && toDate) {
    startDate = new Date(fromDate);
    endDate = new Date(toDate);
  } else if (days > 0) {
    startDate = new Date(now.getTime() - (days * 24 * 60 * 60 * 1000));
    endDate = now;
  } else {
    return JSON.stringify(fullData);
  }

  // Filter monthly filings to date range
  fullData.monthlyFilings = fullData.monthlyFilings.filter(function(m) {
    var monthDate = new Date(m.month + ' 1');
    return monthDate >= startDate && monthDate <= endDate;
  });

  fullData.monthlyResolved = fullData.monthlyResolved.filter(function(m) {
    var monthDate = new Date(m.month + ' 1');
    return monthDate >= startDate && monthDate <= endDate;
  });

  // Filter drill-down data by date
  for (var statusKey in fullData.chartDrillDown.statusByCase) {
    fullData.chartDrillDown.statusByCase[statusKey] = fullData.chartDrillDown.statusByCase[statusKey].filter(function(c) {
      if (!c.dateFiled) return true;
      var filedDate = new Date(c.dateFiled);
      return filedDate >= startDate && filedDate <= endDate;
    });
  }

  // Recalculate counts based on filtered data
  fullData.statusDistribution = { open: 0, pending: 0, won: 0, denied: 0, settled: 0, withdrawn: 0 };
  for (var sk in fullData.chartDrillDown.statusByCase) {
    fullData.statusDistribution[sk] = fullData.chartDrillDown.statusByCase[sk].length;
  }

  fullData.dateRangeApplied = {
    days: days,
    from: startDate.toISOString(),
    to: endDate.toISOString()
  };

  return JSON.stringify(fullData);
}

/**
 * Generates the unified dashboard HTML for web app (v4.4.0)
 * @param {boolean} isPII - Whether to include PII (steward mode)
 * @returns {string} Complete HTML for the web app
 */
function getUnifiedDashboardHtml(isPII) {
  var mode = isPII ? 'steward' : 'member';
  var title = isPII ? '509 STEWARD COMMAND CENTER' : '509 MEMBER DASHBOARD';
  var badge = isPII ? '<span class="pii-badge">INTERNAL USE - CONTAINS PII</span>' : '<span class="member-badge">MEMBER VIEW</span>';

  return '<!DOCTYPE html>' +
    '<html lang="en"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<title>' + title + '</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>' +
    '<style>' +
    // CSS Reset & Base
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:"Inter",sans-serif;background:linear-gradient(180deg,#0f172a 0%,#1e293b 100%);color:#f8fafc;min-height:100vh;overflow-x:hidden}' +

    // Header
    '.header{background:linear-gradient(135deg,#1e293b 0%,#334155 100%);padding:16px 20px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid rgba(255,255,255,0.1);position:sticky;top:0;z-index:100;backdrop-filter:blur(10px)}' +
    '.header h1{font-size:18px;font-weight:700;color:#60a5fa;display:flex;align-items:center;gap:10px}' +
    '.header .material-icons{font-size:24px}' +
    '.pii-badge{background:rgba(239,68,68,0.2);color:#fca5a5;padding:6px 12px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px}' +
    '.member-badge{background:rgba(34,197,94,0.2);color:#86efac;padding:6px 12px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px}' +

    // Tabs
    '.tabs{display:flex;gap:2px;padding:0 16px;background:rgba(0,0,0,0.3);overflow-x:auto;-webkit-overflow-scrolling:touch}' +
    '.tabs::-webkit-scrollbar{display:none}' +
    '.tab{padding:14px 16px;cursor:pointer;font-size:11px;font-weight:600;color:#94a3b8;border-bottom:3px solid transparent;transition:all 0.2s;white-space:nowrap;text-transform:uppercase;letter-spacing:0.5px;touch-action:manipulation;-webkit-tap-highlight-color:transparent}' +
    '.tab:hover{color:#e2e8f0;background:rgba(255,255,255,0.05)}' +
    '.tab.active{color:#60a5fa;border-bottom-color:#60a5fa;background:rgba(96,165,250,0.1)}' +

    // Main content
    '.content{padding:16px;max-width:1200px;margin:0 auto;min-height:calc(100vh - 140px)}' +
    '.tab-content{display:none;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}' +

    // KPI Cards - Clickable
    '.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}' +
    '.kpi-card{background:linear-gradient(135deg,rgba(30,41,59,0.9) 0%,rgba(51,65,85,0.9) 100%);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:16px;text-align:center;cursor:pointer;transition:all 0.2s}' +
    '.kpi-card:hover{transform:translateY(-2px);border-color:rgba(96,165,250,0.5);box-shadow:0 8px 25px rgba(0,0,0,0.3)}' +
    '.kpi-card:active{transform:scale(0.98)}' +
    '.kpi-card.alert{border-color:#ef4444;background:linear-gradient(135deg,rgba(127,29,29,0.3) 0%,rgba(30,41,59,0.9) 100%)}' +
    '.kpi-card.clickable::after{content:"Click for details";display:block;font-size:9px;color:#64748b;margin-top:6px;opacity:0;transition:opacity 0.2s}' +
    '.kpi-card.clickable:hover::after{opacity:1}' +
    '.kpi-label{font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px}' +
    '.kpi-value{font-size:32px;font-weight:900;line-height:1}' +
    '.kpi-value.green{color:#34d399}.kpi-value.red{color:#f87171}.kpi-value.blue{color:#60a5fa}.kpi-value.yellow{color:#fbbf24}.kpi-value.purple{color:#a78bfa}' +

    // Charts
    '.charts-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;margin-bottom:16px}' +
    '.chart-card{background:rgba(30,41,59,0.8);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:16px}' +
    '.chart-title{font-size:13px;font-weight:600;color:#e2e8f0;margin-bottom:14px;display:flex;align-items:center;gap:8px;text-transform:uppercase;letter-spacing:0.5px}' +
    '.chart-title .material-icons{font-size:20px;color:#60a5fa}' +
    'canvas{max-height:220px!important}' +

    // Lists
    '.list-container{max-height:300px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#475569 transparent}' +
    '.list-container::-webkit-scrollbar{width:6px}' +
    '.list-container::-webkit-scrollbar-thumb{background:#475569;border-radius:3px}' +
    '.list-item{display:flex;justify-content:space-between;align-items:center;padding:12px;border-bottom:1px solid rgba(255,255,255,0.05);transition:background 0.2s}' +
    '.list-item:hover{background:rgba(255,255,255,0.03)}' +
    '.list-item:last-child{border-bottom:none}' +
    '.badge{padding:5px 12px;border-radius:20px;font-size:11px;font-weight:600}' +

    // Bargaining Cards
    '.bargain-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}' +
    '.bargain-card{background:linear-gradient(135deg,rgba(251,191,36,0.1) 0%,rgba(30,41,59,0.9) 100%);border:1px solid rgba(251,191,36,0.3);border-radius:12px;padding:16px;text-align:center}' +
    '.bargain-label{font-size:10px;color:#fbbf24;text-transform:uppercase;letter-spacing:1px}' +
    '.bargain-value{font-size:28px;font-weight:800;color:#fcd34d;margin-top:8px}' +
    '.bargain-status{font-size:11px;color:#94a3b8;margin-top:6px}' +

    // Hot Zones
    '.hot-zone{display:flex;justify-content:space-between;align-items:center;padding:14px;background:linear-gradient(90deg,rgba(239,68,68,0.15) 0%,rgba(30,41,59,0.8) 100%);border-radius:10px;margin-bottom:10px;border-left:4px solid #ef4444}' +

    // Satisfaction
    '.sat-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px}' +
    '.sat-response-count{background:rgba(96,165,250,0.2);color:#60a5fa;padding:8px 16px;border-radius:20px;font-size:12px;font-weight:600}' +
    '.sat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px}' +
    '.sat-section{background:rgba(30,41,59,0.8);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:16px}' +
    '.sat-section-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}' +
    '.sat-section-name{font-size:13px;font-weight:600;color:#e2e8f0}' +
    '.sat-section-score{font-size:22px;font-weight:900}' +
    '.sat-score-bar{height:8px;background:rgba(255,255,255,0.1);border-radius:4px;overflow:hidden;margin-bottom:12px}' +
    '.sat-score-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.sat-questions{font-size:11px;color:#94a3b8;line-height:1.7}' +

    // Directory Trends
    '.trend-card{background:rgba(30,41,59,0.8);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:16px;margin-bottom:16px}' +
    '.trend-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}' +
    '.trend-title{font-size:14px;font-weight:600;color:#e2e8f0;display:flex;align-items:center;gap:8px}' +
    '.trend-value{font-size:20px;font-weight:800}' +
    '.trend-list{max-height:150px;overflow-y:auto}' +
    '.trend-item{padding:8px 0;border-bottom:1px solid rgba(255,255,255,0.05);font-size:12px;display:flex;justify-content:space-between}' +

    // Resources
    '.resource-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px}' +
    '.resource-card{background:rgba(30,41,59,0.8);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px;text-align:center;cursor:pointer;transition:all 0.2s;text-decoration:none;color:inherit}' +
    '.resource-card:hover{transform:translateY(-2px);border-color:rgba(96,165,250,0.5);box-shadow:0 8px 25px rgba(0,0,0,0.3)}' +
    '.resource-icon{font-size:40px;margin-bottom:12px}' +
    '.resource-title{font-size:14px;font-weight:600;color:#e2e8f0;margin-bottom:4px}' +
    '.resource-desc{font-size:11px;color:#94a3b8}' +

    // Modal
    '.modal-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.8);z-index:200;align-items:center;justify-content:center;padding:20px}' +
    '.modal-overlay.active{display:flex}' +
    '.modal{background:#1e293b;border-radius:16px;max-width:600px;width:100%;max-height:80vh;overflow:hidden;border:1px solid rgba(255,255,255,0.1);animation:modalIn 0.2s}' +
    '@keyframes modalIn{from{opacity:0;transform:scale(0.95)}to{opacity:1;transform:scale(1)}}' +
    '.modal-header{padding:16px 20px;border-bottom:1px solid rgba(255,255,255,0.1);display:flex;justify-content:space-between;align-items:center}' +
    '.modal-title{font-size:16px;font-weight:700;color:#60a5fa}' +
    '.modal-close{background:none;border:none;color:#94a3b8;font-size:24px;cursor:pointer;padding:4px;min-width:44px;min-height:44px;touch-action:manipulation}' +
    '.modal-close:hover{color:#f8fafc}' +
    '.modal-body{padding:20px;max-height:calc(80vh - 60px);overflow-y:auto}' +
    '.modal-list-item{padding:12px;background:rgba(15,23,42,0.5);border-radius:8px;margin-bottom:8px;display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center}' +
    '.modal-list-id{font-weight:700;color:#60a5fa;font-size:12px}' +
    '.modal-list-name{font-size:13px;color:#e2e8f0}' +
    '.modal-list-meta{font-size:11px;color:#94a3b8}' +

    // Footer
    '.footer{padding:12px 20px;border-top:1px solid rgba(255,255,255,0.1);display:flex;justify-content:space-between;align-items:center;font-size:11px;color:#64748b;background:rgba(15,23,42,0.8)}' +
    '.btn{padding:10px 20px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;border:none;transition:all 0.2s}' +
    '.btn-primary{background:#3b82f6;color:white}.btn-primary:hover{background:#2563eb}' +
    '.btn-secondary{background:#475569;color:white}.btn-secondary:hover{background:#64748b}' +
    '.btn-sm{padding:6px 12px;font-size:11px}' +

    // Loading
    '.loading{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px;color:#94a3b8}' +
    '.spinner{width:40px;height:40px;border:3px solid rgba(96,165,250,0.3);border-top-color:#60a5fa;border-radius:50%;animation:spin 1s linear infinite;margin-bottom:16px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Trend Arrows
    '.trend-arrow{font-size:14px;margin-left:4px}.trend-arrow.up{color:#22c55e}.trend-arrow.down{color:#ef4444}.trend-arrow.flat{color:#94a3b8}' +
    '.trend-pct{font-size:10px;margin-left:2px}' +

    // Goal Progress
    '.goal-bar{height:6px;background:rgba(255,255,255,0.1);border-radius:3px;margin-top:8px;overflow:hidden}' +
    '.goal-fill{height:100%;border-radius:3px;transition:width 0.5s}' +
    '.goal-label{display:flex;justify-content:space-between;font-size:9px;color:#64748b;margin-top:4px}' +
    // Pinned Metrics
    '.pinned-section{margin-bottom:16px;padding:16px;background:rgba(59,130,246,0.1);border:1px solid rgba(59,130,246,0.3);border-radius:12px}' +
    '.pinned-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}' +
    '.pinned-title{font-size:12px;font-weight:600;color:#60a5fa;display:flex;align-items:center;gap:6px}' +
    '.pinned-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px}' +
    '.pinned-metric{background:#1e293b;padding:12px;border-radius:8px;text-align:center;position:relative}' +
    '.pinned-metric .unpin-btn{position:absolute;top:4px;right:4px;background:none;border:none;color:#64748b;cursor:pointer;font-size:14px;opacity:0;transition:opacity 0.2s}' +
    '.pinned-metric:hover .unpin-btn{opacity:1}' +
    '.pin-btn{position:absolute;top:4px;right:4px;background:none;border:none;color:#64748b;cursor:pointer;font-size:14px;opacity:0;transition:opacity 0.2s}' +
    '.kpi-card:hover .pin-btn{opacity:1}' +
    '.kpi-card.pinned .pin-btn{opacity:1;color:#3b82f6}' +
    // Chart Drill-down
    '.chart-card canvas{cursor:pointer}' +
    '.drill-down-list{max-height:300px;overflow-y:auto}' +
    '.drill-down-item{display:flex;justify-content:space-between;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,0.05);font-size:12px}' +
    '.drill-down-item:hover{background:rgba(255,255,255,0.05)}' +

    // Settings Panel
    '.settings-panel{position:fixed;right:-320px;top:0;bottom:0;width:320px;background:#1e293b;border-left:1px solid rgba(255,255,255,0.1);z-index:250;transition:right 0.3s;overflow-y:auto;padding:20px}' +
    '.settings-panel.open{right:0}' +
    '.settings-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:240;display:none}' +
    '.settings-overlay.open{display:block}' +
    '.settings-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid rgba(255,255,255,0.1)}' +
    '.settings-section{margin-bottom:24px}' +
    '.settings-title{font-size:12px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px}' +
    '.setting-row{display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid rgba(255,255,255,0.05)}' +
    '.setting-label{font-size:13px;color:#e2e8f0}' +
    '.toggle{width:44px;height:24px;background:#475569;border-radius:12px;position:relative;cursor:pointer;transition:background 0.2s}' +
    '.toggle.on{background:#22c55e}' +
    '.toggle::after{content:"";position:absolute;width:20px;height:20px;background:white;border-radius:50%;top:2px;left:2px;transition:left 0.2s}' +
    '.toggle.on::after{left:22px}' +

    // Alert Center
    '.alert-center{position:fixed;right:-400px;top:0;bottom:0;width:400px;background:#1e293b;border-left:1px solid rgba(255,255,255,0.1);z-index:250;transition:right 0.3s;overflow-y:auto}' +
    '.alert-center.open{right:0}' +
    '.alert-header{padding:16px 20px;border-bottom:1px solid rgba(255,255,255,0.1);display:flex;justify-content:space-between;align-items:center;background:rgba(15,23,42,0.8);position:sticky;top:0}' +
    '.alert-item{padding:16px 20px;border-bottom:1px solid rgba(255,255,255,0.05)}' +
    '.alert-item.critical{border-left:4px solid #ef4444}' +
    '.alert-item.warning{border-left:4px solid #f59e0b}' +
    '.alert-item.info{border-left:4px solid #3b82f6}' +
    '.alert-item.success{border-left:4px solid #22c55e}' +
    '.alert-title{font-size:13px;font-weight:600;color:#e2e8f0;margin-bottom:4px}' +
    '.alert-desc{font-size:12px;color:#94a3b8}' +
    '.alert-time{font-size:10px;color:#64748b;margin-top:6px}' +
    '.alert-badge{background:#ef4444;color:white;font-size:10px;padding:2px 6px;border-radius:10px;position:absolute;top:-5px;right:-5px}' +

    // Date Range Filter
    '.date-filter{display:flex;align-items:center;gap:8px;padding:8px 16px;background:rgba(0,0,0,0.2);border-bottom:1px solid rgba(255,255,255,0.05)}' +
    '.date-btn{padding:6px 12px;border-radius:6px;font-size:11px;background:#334155;color:#94a3b8;border:none;cursor:pointer;transition:all 0.2s;touch-action:manipulation;-webkit-tap-highlight-color:transparent}' +
    '.date-btn:hover,.date-btn.active{background:#475569;color:#e2e8f0}' +
    '.date-input{padding:6px 10px;border-radius:6px;font-size:11px;background:#1e293b;color:#f8fafc;border:1px solid #475569;width:120px}' +

    // Last Updated
    '.last-updated{font-size:10px;color:#64748b;display:flex;align-items:center;gap:4px}' +
    '.last-updated .material-icons{font-size:12px}' +
    '.pulse{animation:pulse 2s infinite}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}' +

    // Pinned Metrics
    '.pinned-section{background:linear-gradient(135deg,rgba(96,165,250,0.1) 0%,rgba(30,41,59,0.9) 100%);border:1px solid rgba(96,165,250,0.3);border-radius:12px;padding:16px;margin-bottom:20px}' +
    '.pinned-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}' +
    '.pinned-title{font-size:12px;font-weight:600;color:#60a5fa;text-transform:uppercase}' +
    '.pin-btn{background:none;border:none;color:#64748b;cursor:pointer;font-size:18px;transition:color 0.2s}' +
    '.pin-btn:hover,.pin-btn.pinned{color:#fbbf24}' +

    // Print Styles
    '@media print{body{background:white!important;color:black!important}.header,.tabs,.footer,.settings-panel,.alert-center,.btn{display:none!important}.content{padding:0!important;max-width:100%!important}.chart-card,.kpi-card,.trend-card{background:white!important;border:1px solid #ddd!important;break-inside:avoid}.kpi-value,.trend-value{color:black!important}}' +

    // High Contrast Mode
    'body.high-contrast{background:#000!important}' +
    'body.high-contrast .kpi-card,body.high-contrast .chart-card,body.high-contrast .trend-card{background:#111!important;border-color:#fff!important}' +
    'body.high-contrast .kpi-value,body.high-contrast .trend-value,body.high-contrast .chart-title{color:#fff!important}' +
    'body.high-contrast .kpi-label,body.high-contrast .trend-title{color:#ccc!important}' +

    // Large Text Mode
    'body.large-text{font-size:16px}' +
    'body.large-text .kpi-value{font-size:42px}' +
    'body.large-text .kpi-label{font-size:12px}' +
    'body.large-text .chart-title{font-size:16px}' +
    'body.large-text .tab{font-size:13px;padding:16px 20px}' +
    'body.large-text .list-item{font-size:15px;padding:16px}' +

    // Keyboard Shortcut Hint
    '.shortcut-hint{position:fixed;bottom:80px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,0.9);color:white;padding:12px 24px;border-radius:8px;font-size:13px;z-index:300;display:none}' +
    '.shortcut-hint.show{display:block;animation:fadeInOut 2s forwards}' +
    '@keyframes fadeInOut{0%{opacity:0}10%{opacity:1}90%{opacity:1}100%{opacity:0}}' +

    // Responsive
    '@media(max-width:768px){.header h1{font-size:14px}.kpi-grid{grid-template-columns:repeat(2,1fr)}.charts-row{grid-template-columns:1fr}.bargain-grid{grid-template-columns:repeat(2,1fr)}.settings-panel,.alert-center{width:100%;right:-100%}.settings-panel.open,.alert-center.open{right:0}}' +

    // View Mode Toggle Pill
    '.view-toggle{display:flex;background:rgba(255,255,255,0.1);border-radius:20px;padding:3px;gap:2px}' +
    '.view-toggle button{padding:6px 12px;border:none;border-radius:17px;font-size:11px;cursor:pointer;transition:all 0.2s;background:transparent;color:#94a3b8;display:flex;align-items:center;gap:4px;touch-action:manipulation;-webkit-tap-highlight-color:transparent;user-select:none}' +
    '.view-toggle button.active{background:#3b82f6;color:white}' +
    '.view-toggle button .material-icons{font-size:14px}' +
    '@media(max-width:768px){.view-toggle button{padding:8px 14px;min-height:44px}}' +

    // Mobile Mode (forced mobile layout)
    'body.mobile-mode .header{flex-direction:column;gap:12px;padding:12px 16px}' +
    'body.mobile-mode .header h1{font-size:16px}' +
    'body.mobile-mode .header>div{width:100%;justify-content:space-between}' +
    'body.mobile-mode .tabs{flex-wrap:nowrap;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;padding:0 8px}' +
    'body.mobile-mode .tabs::-webkit-scrollbar{display:none}' +
    'body.mobile-mode .tab{flex-shrink:0;padding:12px 16px;font-size:12px}' +
    'body.mobile-mode .kpi-grid{grid-template-columns:repeat(2,1fr)!important;gap:10px}' +
    'body.mobile-mode .kpi-card{padding:14px}' +
    'body.mobile-mode .kpi-value{font-size:28px}' +
    'body.mobile-mode .kpi-label{font-size:10px}' +
    'body.mobile-mode .charts-row{grid-template-columns:1fr!important;gap:12px}' +
    'body.mobile-mode .chart-card{padding:14px}' +
    'body.mobile-mode .chart-title{font-size:12px}' +
    'body.mobile-mode .list-item{padding:14px;font-size:13px}' +
    'body.mobile-mode .date-filter{flex-wrap:wrap;gap:6px;padding:10px 12px}' +
    'body.mobile-mode .date-btn{padding:8px 14px;font-size:12px}' +
    'body.mobile-mode .date-input{width:100px}' +
    'body.mobile-mode .pinned-section{padding:12px}' +
    'body.mobile-mode .pinned-grid{grid-template-columns:repeat(2,1fr)}' +
    'body.mobile-mode .modal{width:95%;max-height:85vh;margin:10px}' +
    'body.mobile-mode .settings-panel,body.mobile-mode .alert-center{width:100%;right:-100%}' +
    'body.mobile-mode .settings-panel.open,body.mobile-mode .alert-center.open{right:0}' +
    'body.mobile-mode .btn{padding:12px 20px;font-size:13px}' +
    'body.mobile-mode .content{padding:12px}' +
    'body.mobile-mode .trend-grid{grid-template-columns:repeat(2,1fr)}' +
    'body.mobile-mode .bargain-grid{grid-template-columns:repeat(2,1fr)}' +
    'body.mobile-mode .heatmap-grid{grid-template-columns:repeat(2,1fr)}' +
    'body.mobile-mode .satisfaction-grid{grid-template-columns:1fr}' +
    'body.mobile-mode .faq-container{padding:12px}' +
    'body.mobile-mode .faq-item{padding:12px}' +
    'body.mobile-mode .compare-grid{grid-template-columns:repeat(2,1fr)}' +
    'body.mobile-mode .last-updated{display:none}' +

    // Desktop Mode (forced desktop layout)
    'body.desktop-mode .kpi-grid{grid-template-columns:repeat(6,1fr)!important}' +
    'body.desktop-mode .charts-row{grid-template-columns:repeat(2,1fr)!important}' +
    'body.desktop-mode .tab{padding:14px 24px}' +
    'body.desktop-mode .tabs{flex-wrap:nowrap}' +

    // Auto-detect small screens
    '@media(max-width:600px){body:not(.desktop-mode) .header{flex-direction:column;gap:10px}body:not(.desktop-mode) .tabs{overflow-x:auto}body:not(.desktop-mode) .tab{flex-shrink:0;font-size:11px;padding:10px 14px}body:not(.desktop-mode) .kpi-grid{grid-template-columns:repeat(2,1fr)}body:not(.desktop-mode) .charts-row{grid-template-columns:1fr}body:not(.desktop-mode) .date-filter{flex-wrap:wrap}body:not(.desktop-mode) .last-updated{display:none}}' +
    '</style></head><body>' +

    // Settings Overlay
    '<div class="settings-overlay" onclick="toggleSettings()" ontouchend="event.preventDefault();toggleSettings()"></div>' +

    // Settings Panel
    '<div class="settings-panel" id="settingsPanel">' +
    '<div class="settings-header"><span style="font-size:16px;font-weight:700;color:#e2e8f0">Settings</span><button onclick="toggleSettings()" ontouchend="event.preventDefault();toggleSettings()" style="background:none;border:none;color:#94a3b8;font-size:24px;cursor:pointer;min-width:44px;min-height:44px;touch-action:manipulation">&times;</button></div>' +
    '<div class="settings-section"><div class="settings-title">Display</div>' +
    '<div class="setting-row"><span class="setting-label">High Contrast Mode</span><div class="toggle" id="highContrastToggle" onclick="toggleHighContrast()"></div></div>' +
    '<div class="setting-row"><span class="setting-label">Large Text Mode</span><div class="toggle" id="largeTextToggle" onclick="toggleLargeText()"></div></div>' +
    '</div>' +
    '<div class="settings-section"><div class="settings-title">Auto Refresh</div>' +
    '<div class="setting-row"><span class="setting-label">Enable Auto-Refresh</span><div class="toggle" id="autoRefreshToggle" onclick="toggleAutoRefresh()"></div></div>' +
    '<div class="setting-row"><span class="setting-label">Interval</span><select id="refreshInterval" onchange="setRefreshInterval()" style="padding:6px;border-radius:6px;background:#334155;color:#e2e8f0;border:none"><option value="60000">1 minute</option><option value="300000" selected>5 minutes</option><option value="600000">10 minutes</option></select></div>' +
    '</div>' +
    '<div class="settings-section"><div class="settings-title">Goals</div>' +
    '<div class="setting-row"><span class="setting-label">Win Rate Target</span><input type="number" id="goalWinRate" value="75" min="0" max="100" style="width:60px;padding:6px;border-radius:6px;background:#334155;color:#e2e8f0;border:none;text-align:center" onchange="saveGoals()">%</div>' +
    '<div class="setting-row"><span class="setting-label">Morale Target</span><input type="number" id="goalMorale" value="8" min="1" max="10" step="0.1" style="width:60px;padding:6px;border-radius:6px;background:#334155;color:#e2e8f0;border:none;text-align:center" onchange="saveGoals()">/10</div>' +
    '<div class="setting-row"><span class="setting-label">Response Rate Target</span><input type="number" id="goalResponse" value="50" min="0" max="100" style="width:60px;padding:6px;border-radius:6px;background:#334155;color:#e2e8f0;border:none;text-align:center" onchange="saveGoals()">%</div>' +
    '</div>' +
    '<div class="settings-section"><div class="settings-title">Actions</div>' +
    '<button class="btn btn-secondary" style="width:100%;margin-bottom:8px" onclick="printDashboard()"><i class="material-icons" style="font-size:14px;vertical-align:middle;margin-right:4px">print</i>Print Dashboard</button>' +
    '<button class="btn btn-secondary" style="width:100%;margin-bottom:8px" onclick="exportAllData()"><i class="material-icons" style="font-size:14px;vertical-align:middle;margin-right:4px">download</i>Export All Data</button>' +
    '<button class="btn btn-secondary" style="width:100%" onclick="scheduleReport()"><i class="material-icons" style="font-size:14px;vertical-align:middle;margin-right:4px">schedule</i>Schedule Report</button>' +
    '</div>' +
    '<div class="settings-section"><div class="settings-title">Keyboard Shortcuts</div>' +
    '<div style="font-size:11px;color:#94a3b8;line-height:2">' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">1-9</kbd> Switch tabs</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">/</kbd> Focus search</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">R</kbd> Refresh data</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">P</kbd> Print view</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">A</kbd> Alert center</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">S</kbd> Settings</div>' +
    '<div><kbd style="background:#334155;padding:2px 6px;border-radius:4px">Esc</kbd> Close panels</div>' +
    '</div></div>' +
    '</div>' +

    // Alert Center
    '<div class="alert-center" id="alertCenter">' +
    '<div class="alert-header"><span style="font-size:16px;font-weight:700;color:#e2e8f0">Alert Center</span><button onclick="toggleAlerts()" ontouchend="event.preventDefault();toggleAlerts()" style="background:none;border:none;color:#94a3b8;font-size:24px;cursor:pointer;min-width:44px;min-height:44px;touch-action:manipulation">&times;</button></div>' +
    '<div id="alertList"></div>' +
    '</div>' +

    // Shortcut Hint
    '<div class="shortcut-hint" id="shortcutHint"></div>' +

    // Header with view toggle and buttons
    '<div class="header"><h1><i class="material-icons">analytics</i>' + title + '</h1><div style="display:flex;align-items:center;gap:12px">' +
    // View Mode Toggle Pill
    '<div class="view-toggle" id="viewToggle">' +
    '<button id="autoViewBtn" class="active" onclick="setViewMode(\\x27auto\\x27)" ontouchend="event.preventDefault();setViewMode(\\x27auto\\x27)"><i class="material-icons">auto_fix_high</i>Auto</button>' +
    '<button id="desktopViewBtn" onclick="setViewMode(\\x27desktop\\x27)" ontouchend="event.preventDefault();setViewMode(\\x27desktop\\x27)"><i class="material-icons">computer</i>Desktop</button>' +
    '<button id="mobileViewBtn" onclick="setViewMode(\\x27mobile\\x27)" ontouchend="event.preventDefault();setViewMode(\\x27mobile\\x27)"><i class="material-icons">smartphone</i>Mobile</button>' +
    '</div>' +
    '<div class="last-updated" id="lastUpdated"><i class="material-icons">schedule</i><span>Loading...</span></div>' +
    badge +
    '<button onclick="toggleAlerts()" style="background:none;border:none;color:#94a3b8;cursor:pointer;position:relative"><i class="material-icons">notifications</i><span class="alert-badge" id="alertBadge" style="display:none">0</span></button>' +
    '<button onclick="toggleSettings()" style="background:none;border:none;color:#94a3b8;cursor:pointer"><i class="material-icons">settings</i></button>' +
    '</div></div>' +

    // Date Range Filter
    '<div class="date-filter">' +
    '<span style="font-size:11px;color:#64748b;margin-right:8px">Period:</span>' +
    '<button class="date-btn active" onclick="setDateRange(7,this)">7D</button>' +
    '<button class="date-btn" onclick="setDateRange(30,this)">30D</button>' +
    '<button class="date-btn" onclick="setDateRange(90,this)">90D</button>' +
    '<button class="date-btn" onclick="setDateRange(365,this)">1Y</button>' +
    '<button class="date-btn" onclick="setDateRange(0,this)">All</button>' +
    '<span style="margin-left:auto;font-size:11px;color:#64748b">Custom:</span>' +
    '<input type="date" class="date-input" id="dateFrom" onchange="applyCustomDateRange()">' +
    '<span style="color:#64748b">to</span>' +
    '<input type="date" class="date-input" id="dateTo" onchange="applyCustomDateRange()">' +
    '</div>' +

    // Tabs (My Cases and Help only visible in steward mode, Compare in both)
    '<div class="tabs">' +
    '<div class="tab active" onclick="showTab(\'overview\')">Overview</div>' +
    (isPII ? '<div class="tab" onclick="showTab(\'mycases\')">My Cases</div>' : '') +
    '<div class="tab" onclick="showTab(\'workload\')">Workload</div>' +
    '<div class="tab" onclick="showTab(\'analytics\')">Analytics</div>' +
    '<div class="tab" onclick="showTab(\'directory\')">Directory</div>' +
    '<div class="tab" onclick="showTab(\'hotspots\')">Hot Spots</div>' +
    '<div class="tab" onclick="showTab(\'bargaining\')">Bargaining</div>' +
    '<div class="tab" onclick="showTab(\'satisfaction\')">Satisfaction</div>' +
    '<div class="tab" onclick="showTab(\'resources\')">Resources</div>' +
    '<div class="tab" onclick="showTab(\'meetingnotes\')">Meeting Notes</div>' +
    '<div class="tab" onclick="showTab(\'compare\')">Compare</div>' +
    (isPII ? '<div class="tab" onclick="showTab(\'help\')">Help</div>' : '') +
    '</div>' +

    // Content
    '<div class="content"><div id="main-content"><div class="loading"><div class="spinner"></div>Loading dashboard data...</div></div></div>' +

    // Modal
    '<div class="modal-overlay" id="modal" onclick="if(event.target===this)closeModal()" ontouchend="if(event.target===this){event.preventDefault();closeModal()}">' +
    '<div class="modal">' +
    '<div class="modal-header"><span class="modal-title" id="modal-title">Details</span><button class="modal-close" onclick="closeModal()" ontouchend="event.preventDefault();closeModal()">&times;</button></div>' +
    '<div class="modal-body" id="modal-body"></div>' +
    '</div></div>' +

    // Footer with Help/FAQ button
    '<div class="footer"><span>Data refreshes on load | v4.4.0</span><div style="display:flex;gap:8px"><button class="btn btn-secondary" onclick="showFAQ()"><i class="material-icons" style="font-size:14px;vertical-align:middle;margin-right:4px">help</i>Help</button><button class="btn btn-secondary" onclick="location.reload()">Refresh</button></div></div>' +

    // JavaScript
    '<script>' +
    'function escapeHtml(t){if(t==null)return"";return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#x27;").replace(/\\//g,"&#x2F;");}' +
    'var dashData=null;var isPII=' + isPII + ';' +
    'window.onload=function(){google.script.run.withSuccessHandler(render).withFailureHandler(showError).getUnifiedDashboardDataAPI(isPII)};' +
    'function showError(e){document.getElementById("main-content").innerHTML="<div class=\\"loading\\">Error: "+escapeHtml(e.message)+"</div>"}' +
    'function showTab(tab){document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});document.querySelector(".tab[onclick*=\\x27"+tab+"\\x27]").classList.add("active");renderTab(tab)}' +
    'function render(json){dashData=JSON.parse(json);renderTab("overview");setTimeout(renderPinnedSection,100)}' +

    // Modal functions
    'function openModal(title,content){document.getElementById("modal-title").textContent=title;document.getElementById("modal-body").innerHTML=content;document.getElementById("modal").classList.add("active")}' +
    'function closeModal(){document.getElementById("modal").classList.remove("active")}' +

    // Show list in modal with search (v4.4.0)
    'function showList(type){' +
    'var d=dashData,items=[],title="",listHtml="";' +
    'if(type==="members"){title="All Members ("+d.totalMembers+")";items=d.memberList.map(function(m){return{id:m.id,name:m.name+(m.isSteward?" (Steward)":""),meta:m.location+" | "+m.unit}})}' +
    'else if(type==="stewards"){title="Stewards ("+d.stewardCount+")";items=d.stewardList.map(function(s){return{id:s.id,name:s.name,meta:s.location+" | "+s.unit}})}' +
    'else if(type==="open"){title="Open Cases ("+d.openGrievances+")";items=d.openCasesList.map(function(c){return{id:c.id,name:c.member+" - "+c.category,meta:c.step+" | "+c.steward}})}' +
    'else if(type==="overdue"){title="Overdue Cases ("+d.overdueCount+")";items=d.overdueList.map(function(c){return{id:c.id,name:c.member,meta:c.step+" | "+c.steward}})}' +
    'if(items.length===0){listHtml="<p style=\\"color:#94a3b8\\">No data available in this view.</p>"}' +
    'else{listHtml="<input type=\\"text\\" id=\\"modalSearch\\" placeholder=\\"Search...\\" oninput=\\"filterModalList()\\" style=\\"width:100%;padding:10px;margin-bottom:12px;border:1px solid #475569;border-radius:8px;background:#1e293b;color:#f8fafc;font-size:13px\\"><div id=\\"modalListItems\\" style=\\"max-height:350px;overflow-y:auto\\">";' +
    'items.forEach(function(item){listHtml+="<div class=\\"modal-list-item\\" data-search=\\""+escapeHtml(((item.id||"")+" "+(item.name||"")+" "+(item.meta||"")).toLowerCase())+"\\"><span class=\\"modal-list-id\\">"+escapeHtml(item.id)+"</span><span class=\\"modal-list-name\\">"+escapeHtml(item.name)+"</span><span class=\\"modal-list-meta\\">"+escapeHtml(item.meta)+"</span></div>"});' +
    'listHtml+="</div>"}' +
    'openModal(title,listHtml)}' +
    'function filterModalList(){var q=document.getElementById("modalSearch").value.toLowerCase();document.querySelectorAll("#modalListItems .modal-list-item").forEach(function(el){el.style.display=el.getAttribute("data-search").indexOf(q)>=0?"flex":"none"})}' +

    // Render tab content
    'function renderTab(tab){' +
    'var d=dashData,html="";' +

    // Overview Tab with trend arrows, goal progress, pin buttons, and enhanced insights
    'if(tab==="overview"){' +
    'var goals=JSON.parse(localStorage.getItem("509_goals")||"{}");' +
    'var winTarget=goals.winRate||75,moraleTarget=goals.morale||8,responseTarget=goals.response||50;' +
    'var isPinned=function(k){return pinnedMetrics.some(function(p){return p.key===k})};' +
    // Primary KPIs
    'html="<div class=\\"kpi-grid\\">";' +
    'html+="<div class=\\"kpi-card clickable "+(isPinned("members")?"pinned":"")+"\\" onclick=\\"showList(\\x27members\\x27)\\"><button class=\\"pin-btn\\" onclick=\\"event.stopPropagation();togglePin(\\x27members\\x27,\\x27Members\\x27,"+d.totalMembers+",\\x27#3b82f6\\x27)\\"><i class=\\"material-icons\\">"+(isPinned("members")?"push_pin":"push_pin")+"</i></button><div class=\\"kpi-label\\">Members</div><div class=\\"kpi-value blue\\">"+d.totalMembers+"</div></div>";' +
    'html+="<div class=\\"kpi-card clickable "+(isPinned("stewards")?"pinned":"")+"\\" onclick=\\"showList(\\x27stewards\\x27)\\"><button class=\\"pin-btn\\" onclick=\\"event.stopPropagation();togglePin(\\x27stewards\\x27,\\x27Stewards\\x27,"+d.stewardCount+",\\x27#a855f7\\x27)\\"><i class=\\"material-icons\\">push_pin</i></button><div class=\\"kpi-label\\">Stewards</div><div class=\\"kpi-value purple\\">"+d.stewardCount+"</div></div>";' +
    'html+="<div class=\\"kpi-card clickable "+(isPinned("openCases")?"pinned":"")+"\\" onclick=\\"showList(\\x27open\\x27)\\"><button class=\\"pin-btn\\" onclick=\\"event.stopPropagation();togglePin(\\x27openCases\\x27,\\x27Open Cases\\x27,"+d.openGrievances+",\\x27#f59e0b\\x27)\\"><i class=\\"material-icons\\">push_pin</i></button><div class=\\"kpi-label\\">Open Cases</div><div class=\\"kpi-value yellow\\">"+d.openGrievances+(d.prevOpenCases!==undefined?getTrendArrow(d.openGrievances,d.prevOpenCases):"")+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(isPinned("winRate")?"pinned":"")+"\\"><button class=\\"pin-btn\\" onclick=\\"togglePin(\\x27winRate\\x27,\\x27Win Rate\\x27,\\x27"+d.winRate+"%\\x27,\\x27#22c55e\\x27)\\"><i class=\\"material-icons\\">push_pin</i></button><div class=\\"kpi-label\\">Win Rate <span style=\\"font-size:9px;color:#64748b\\">(Goal: "+winTarget+"%)</span></div><div class=\\"kpi-value green\\">"+d.winRate+"%"+(d.prevWinRate!==undefined?getTrendArrow(d.winRate,d.prevWinRate):"")+"</div>"+getGoalBar(d.winRate,winTarget)+"</div>";' +
    'html+="<div class=\\"kpi-card clickable "+(d.overdueCount>0?"alert":"")+" "+(isPinned("overdue")?"pinned":"")+"\\" onclick=\\"showList(\\x27overdue\\x27)\\"><button class=\\"pin-btn\\" onclick=\\"event.stopPropagation();togglePin(\\x27overdue\\x27,\\x27Overdue\\x27,"+d.overdueCount+",\\x27#ef4444\\x27)\\"><i class=\\"material-icons\\">push_pin</i></button><div class=\\"kpi-label\\">Overdue</div><div class=\\"kpi-value red\\">"+d.overdueCount+(d.prevOverdueCount!==undefined?getTrendArrow(d.prevOverdueCount,d.overdueCount):"")+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(isPinned("morale")?"pinned":"")+"\\"><button class=\\"pin-btn\\" onclick=\\"togglePin(\\x27morale\\x27,\\x27Morale\\x27,"+d.moraleScore+",\\x27#22c55e\\x27)\\"><i class=\\"material-icons\\">push_pin</i></button><div class=\\"kpi-label\\">Morale <span style=\\"font-size:9px;color:#64748b\\">(Goal: "+moraleTarget+")</span></div><div class=\\"kpi-value "+(d.moraleScore>=7?"green":d.moraleScore>=5?"yellow":"red")+"\\">"+d.moraleScore+(d.prevMoraleScore!==undefined?getTrendArrow(d.moraleScore,d.prevMoraleScore):"")+"</div>"+getGoalBar(d.moraleScore*10,moraleTarget*10)+"</div>";' +
    'html+="</div>";' +
    // Quick Insights Panel
    'html+="<div class=\\"chart-card\\" style=\\"margin-bottom:16px;background:linear-gradient(135deg,rgba(96,165,250,0.08),rgba(139,92,246,0.08))\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">lightbulb</i>Quick Insights</div><div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px\\">";' +
    // Insight 1: Grievance Status
    'var totalActive=d.openGrievances+(d.pendingCases||0);' +
    'html+="<div style=\\"padding:12px;background:rgba(96,165,250,0.1);border-radius:8px;border-left:3px solid #60a5fa\\"><div style=\\"font-size:11px;color:#60a5fa;font-weight:600\\">GRIEVANCE STATUS</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+totalActive+" active cases | "+d.totalGrievances+" total all-time</div><div style=\\"font-size:10px;color:#94a3b8;margin-top:2px\\">Won: "+d.wins+" | Denied: "+d.losses+" | Settled: "+d.settled+"</div></div>";' +
    // Insight 2: Hot Spot Alert
    'var hsCount=d.hotSpots?(d.hotSpots.grievance?d.hotSpots.grievance.length:0)+(d.hotSpots.dissatisfaction?d.hotSpots.dissatisfaction.length:0):d.hotZones.length;' +
    'var hsColor=hsCount>0?"#ef4444":"#22c55e";' +
    'html+="<div style=\\"padding:12px;background:rgba("+(hsCount>0?"239,68,68":"34,197,94")+",0.1);border-radius:8px;border-left:3px solid "+hsColor+"\\"><div style=\\"font-size:11px;color:"+hsColor+";font-weight:600\\">HOT SPOTS</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+(hsCount>0?hsCount+" areas need attention":"No hot spots - All clear!")+"</div><div style=\\"font-size:10px;color:#94a3b8;margin-top:2px\\">Grievance clusters, dissatisfaction, low engagement</div></div>";' +
    // Insight 3: Engagement Summary
    'var avgEngagement=Math.round((d.engagement.emailOpenRate+d.engagement.virtualMeetingRate+d.engagement.inPersonMeetingRate)/3);' +
    'var engColor=avgEngagement>=40?"#22c55e":avgEngagement>=25?"#f59e0b":"#ef4444";' +
    'html+="<div style=\\"padding:12px;background:rgba("+(avgEngagement>=40?"34,197,94":avgEngagement>=25?"245,158,11":"239,68,68")+",0.1);border-radius:8px;border-left:3px solid "+engColor+"\\"><div style=\\"font-size:11px;color:"+engColor+";font-weight:600\\">ENGAGEMENT</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+avgEngagement+"% average engagement</div><div style=\\"font-size:10px;color:#94a3b8;margin-top:2px\\">Email: "+d.engagement.emailOpenRate+"% | Meetings: "+(d.engagement.virtualMeetingRate+d.engagement.inPersonMeetingRate)+"%</div></div>";' +
    // Insight 4: Bargaining Position
    'var bargainColor=d.step1DenialRate>60?"#ef4444":d.step1DenialRate>40?"#f59e0b":"#22c55e";' +
    'html+="<div style=\\"padding:12px;background:rgba("+(d.step1DenialRate>60?"239,68,68":d.step1DenialRate>40?"245,158,11":"34,197,94")+",0.1);border-radius:8px;border-left:3px solid "+bargainColor+"\\"><div style=\\"font-size:11px;color:"+bargainColor+";font-weight:600\\">BARGAINING POSITION</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">Step 1 Denial: "+d.step1DenialRate+"% | Avg "+d.avgSettlementDays+" days</div><div style=\\"font-size:10px;color:#94a3b8;margin-top:2px\\">"+(d.step1DenialRate>60?"High management hostility":"Normal bargaining environment")+"</div></div>";' +
    'html+="</div></div>";' +
    // Secondary Metrics Row
    'html+="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(auto-fit,minmax(100px,1fr));margin-bottom:16px\\">";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Grievances</div><div class=\\"kpi-value blue\\">"+d.totalGrievances+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Member:Steward</div><div class=\\"kpi-value purple\\">"+d.stewardRatio+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Survey Response</div><div class=\\"kpi-value yellow\\">"+d.engagement.surveyResponseRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Email Open Rate</div><div class=\\"kpi-value "+(d.engagement.emailOpenRate>=50?"green":"yellow")+"\\">"+d.engagement.emailOpenRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Step 1 Denial</div><div class=\\"kpi-value "+(d.step1DenialRate>60?"red":"yellow")+"\\">"+d.step1DenialRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Avg Settlement</div><div class=\\"kpi-value blue\\">"+d.avgSettlementDays+" days</div></div>";' +
    'html+="</div>";' +
    // Charts
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">pie_chart</i>Case Status <span style=\\"font-size:10px;color:#64748b\\">(click segment)</span></div><canvas id=\\"statusChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">trending_up</i>Morale Trend</div><canvas id=\\"trendChart\\"></canvas></div></div>";' +
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">location_on</i>Members by Location <span style=\\"font-size:10px;color:#64748b\\">(click bar)</span></div><canvas id=\\"locationChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">show_chart</i>Filed vs Resolved</div><canvas id=\\"filingChart\\"></canvas></div></div>";' +
    'document.getElementById("main-content").innerHTML=html;renderOverviewCharts()' +
    '}' +

    // My Cases Tab (Steward mode only) - v4.4.0
    'else if(tab==="mycases"){' +
    'html="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(3,1fr)\\"><div class=\\"kpi-card\\"><div class=\\"kpi-label\\">My Active Cases</div><div class=\\"kpi-value blue\\">"+d.myCases.length+"</div></div>";' +
    'var urgentCount=d.myCases.filter(function(c){return c.daysOpen>30||c.status.toLowerCase().indexOf("pending")>=0}).length;' +
    'html+="<div class=\\"kpi-card "+(urgentCount>0?"alert":"")+"\\"><div class=\\"kpi-label\\">Needs Attention</div><div class=\\"kpi-value "+(urgentCount>0?"red":"green")+"\\">"+urgentCount+"</div></div>";' +
    'var avgDays=d.myCases.length>0?Math.round(d.myCases.reduce(function(s,c){return s+(c.daysOpen||0)},0)/d.myCases.length):0;' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Avg Days Open</div><div class=\\"kpi-value "+(avgDays>45?"red":avgDays>30?"yellow":"green")+"\\">"+avgDays+"</div></div></div>";' +
    // Status filter buttons
    'html+="<div style=\\"display:flex;gap:8px;margin:12px 0\\"><button class=\\"btn btn-primary btn-sm\\" onclick=\\"filterMyCases(\\x27all\\x27)\\">All</button><button class=\\"btn btn-secondary btn-sm\\" onclick=\\"filterMyCases(\\x27Open\\x27)\\">Open</button><button class=\\"btn btn-secondary btn-sm\\" onclick=\\"filterMyCases(\\x27Pending\\x27)\\">Pending</button><button class=\\"btn btn-secondary btn-sm\\" style=\\"background:#ef4444\\" onclick=\\"filterMyCases(\\x27Overdue\\x27)\\">Overdue</button></div>";' +
    'if(d.myCases.length===0){html+="<div class=\\"chart-card\\" style=\\"text-align:center;padding:60px\\"><i class=\\"material-icons\\" style=\\"font-size:48px;color:#22c55e\\">check_circle</i><p style=\\"color:#94a3b8;margin-top:16px\\">No active cases assigned to you. Great job!</p></div>"}' +
    'else{html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">folder_open</i>My Assigned Cases</div><div id=\\"myCasesList\\" class=\\"list-container\\" style=\\"max-height:400px\\">";' +
    'd.myCases.forEach(function(c){var statusColor=c.status.toLowerCase()==="open"?"#f59e0b":c.status.toLowerCase().indexOf("pending")>=0?"#ef4444":"#64748b";var isOverdue=c.daysOpen>30||c.status.toLowerCase().indexOf("pending")>=0;' +
    'html+="<div class=\\"my-case-item list-item\\" data-status=\\""+c.status+"\\" data-overdue=\\""+(isOverdue?"yes":"no")+"\\" style=\\"flex-wrap:wrap\\"><div style=\\"display:flex;justify-content:space-between;width:100%\\"><span style=\\"font-weight:600\\">"+c.id+"</span><span class=\\"badge\\" style=\\"background:"+statusColor+";color:white\\">"+c.status+(isOverdue?" ⚠":"")+"</span></div>";' +
    'html+="<div style=\\"width:100%;margin-top:6px;font-size:12px;color:#cbd5e1\\">"+c.member+" | "+c.category+"</div>";' +
    'html+="<div style=\\"width:100%;display:flex;justify-content:space-between;margin-top:4px;font-size:11px;color:#94a3b8\\"><span>"+c.step+" | "+c.location+"</span><span>"+c.daysOpen+" days</span></div></div>"});' +
    'html+="</div></div>"}' +
    'document.getElementById("main-content").innerHTML=html' +
    '}' +

    // Workload Tab (Enhanced with ratio and top performers)
    'else if(tab==="workload"){' +
    'var totalCases=d.stewardWorkload.reduce(function(s,w){return s+w.count},0);' +
    'var avgCases=d.stewardWorkload.length>0?(totalCases/d.stewardWorkload.length).toFixed(1):0;' +
    'var overloaded=d.stewardWorkload.filter(function(w){return w.status==="OVERLOAD"}).length;' +
    'html="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(5,1fr)\\"><div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Stewards</div><div class=\\"kpi-value blue\\">"+d.stewardCount+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Member:Steward</div><div class=\\"kpi-value purple\\">"+d.stewardRatio+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Active Cases</div><div class=\\"kpi-value yellow\\">"+totalCases+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Avg per Steward</div><div class=\\"kpi-value green\\">"+avgCases+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(overloaded>0?"alert":"")+"\\"><div class=\\"kpi-label\\">Overloaded</div><div class=\\"kpi-value red\\">"+overloaded+"</div></div></div>";' +
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">assignment_ind</i>Steward Caseload</div><div class=\\"list-container\\">";' +
    'd.stewardWorkload.forEach(function(w){html+="<div class=\\"list-item\\"><span>"+w.name+"</span><span class=\\"badge\\" style=\\"background:"+w.color+";color:white\\">"+w.count+" cases - "+w.status+"</span></div>"});' +
    'html+="</div></div>";' +
    // Top Performers section
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">emoji_events</i>Top Performers</div><div class=\\"list-container\\">";' +
    'if(d.topPerformers&&d.topPerformers.length>0){d.topPerformers.forEach(function(p,i){html+="<div class=\\"list-item\\"><span>"+(i+1)+". "+p.name+"</span><span class=\\"badge\\" style=\\"background:#22c55e;color:white\\">Score: "+p.score+" | Win: "+p.winRate+"%</span></div>"})}' +
    'else{html+="<p style=\\"color:#94a3b8;text-align:center;padding:20px\\">Performance data not yet calculated</p>"}' +
    'html+="</div></div></div>";' +
    'document.getElementById("main-content").innerHTML=html' +
    '}' +

    // Analytics Tab (Enhanced with Engagement Metrics)
    'else if(tab==="analytics"){' +
    'html="<div class=\\"kpi-grid\\"><div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Grievances</div><div class=\\"kpi-value blue\\">"+d.totalGrievances+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Won</div><div class=\\"kpi-value green\\">"+d.wins+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Denied</div><div class=\\"kpi-value red\\">"+d.losses+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Settled</div><div class=\\"kpi-value purple\\">"+d.settled+"</div></div>";' +
    'var responseTarget=(JSON.parse(localStorage.getItem("509_goals")||"{}").response)||50;' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Survey Response <span style=\\"font-size:9px;color:#64748b\\">(Goal: "+responseTarget+"%)</span></div><div class=\\"kpi-value yellow\\">"+d.engagement.surveyResponseRate+"%</div>"+getGoalBar(d.engagement.surveyResponseRate,responseTarget)+"</div></div>";' +
    // Engagement Metrics Section
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:20px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">trending_up</i>Member Engagement Metrics</h3>";' +
    'html+="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(auto-fit,minmax(120px,1fr))\\"><div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Email Open Rate</div><div class=\\"kpi-value "+(d.engagement.emailOpenRate>=50?"green":d.engagement.emailOpenRate>=30?"yellow":"red")+"\\">"+d.engagement.emailOpenRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Virtual Mtg Att.</div><div class=\\"kpi-value "+(d.engagement.virtualMeetingRate>=40?"green":d.engagement.virtualMeetingRate>=20?"yellow":"red")+"\\">"+d.engagement.virtualMeetingRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">In-Person Mtg Att.</div><div class=\\"kpi-value "+(d.engagement.inPersonMeetingRate>=30?"green":d.engagement.inPersonMeetingRate>=15?"yellow":"red")+"\\">"+d.engagement.inPersonMeetingRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Vol. Hours</div><div class=\\"kpi-value purple\\">"+d.engagement.totalVolunteerHours+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Local Interest</div><div class=\\"kpi-value blue\\">"+d.engagement.unionInterestLocal+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Chapter Interest</div><div class=\\"kpi-value blue\\">"+d.engagement.unionInterestChapter+"%</div></div></div>";' +
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">donut_large</i>Unit Distribution</div><canvas id=\\"unitChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">bar_chart</i>Outcomes</div><canvas id=\\"outcomeChart\\"></canvas></div></div>";' +
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">category</i>Cases by Category</div><canvas id=\\"categoryChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">stairs</i>Step Progression</div><canvas id=\\"stepChart\\"></canvas></div></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">mood</i>Member Satisfaction: "+d.moraleScore+"/10</div><div style=\\"height:20px;background:rgba(255,255,255,0.1);border-radius:10px;overflow:hidden;margin-top:12px\\"><div style=\\"height:100%;width:"+(d.moraleScore*10)+"%;background:linear-gradient(90deg,#22c55e,#3b82f6)\\"></div></div></div>";' +
    'document.getElementById("main-content").innerHTML=html;renderAnalyticsCharts()' +
    '}' +

    // Directory Trends Tab (Enhanced with Full Analytics)
    'else if(tab==="directory"){' +
    'var dt=d.directoryTrends;var eng=d.engagement;' +
    // Number formatting function
    'function fmt(n){return n>=1000?n.toLocaleString():n}' +
    'html="<div class=\\"kpi-grid\\"><div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Members</div><div class=\\"kpi-value blue\\">"+fmt(d.totalMembers)+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">With Email</div><div class=\\"kpi-value green\\">"+fmt(dt.totalWithEmail)+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Missing Email</div><div class=\\"kpi-value "+(dt.missingEmail>0?"red":"green")+"\\">"+fmt(dt.missingEmail)+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Missing Phone</div><div class=\\"kpi-value "+(dt.missingPhone>0?"yellow":"green")+"\\">"+fmt(dt.missingPhone)+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Email Open Rate</div><div class=\\"kpi-value "+(eng.emailOpenRate>=50?"green":eng.emailOpenRate>=30?"yellow":"red")+"\\">"+eng.emailOpenRate+"%</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Meeting Attendance</div><div class=\\"kpi-value purple\\">"+(eng.virtualMeetingRate+eng.inPersonMeetingRate)+"%</div></div></div>";' +
    // Section: Employee Distribution
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:24px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">people</i>Employee Distribution</h3>";' +
    'html+="<div class=\\"charts-row\\">";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">location_on</i>Employees by Office Location</div><canvas id=\\"empByLocationChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">business</i>Employees by Unit</div><canvas id=\\"empByUnitChart\\"></canvas></div></div>";' +
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">event</i>Employees by Office Days</div><canvas id=\\"empByOfficeDaysChart\\"></canvas></div>";' +
    // Section: Participation Heatmap
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:24px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#22c55e\\">insights</i>Participation Rate Heatmap</h3>";' +
    'html+="<div style=\\"display:flex;gap:4px;margin-bottom:12px;font-size:10px;align-items:center\\"><span style=\\"color:#94a3b8\\">Low</span><div style=\\"display:flex;height:12px\\"><div style=\\"width:20px;background:#fee2e2\\"></div><div style=\\"width:20px;background:#fef3c7\\"></div><div style=\\"width:20px;background:#d9f99d\\"></div><div style=\\"width:20px;background:#bbf7d0\\"></div><div style=\\"width:20px;background:#86efac\\"></div></div><span style=\\"color:#94a3b8\\">High</span></div>";' +
    // Participation by Unit heatmap
    'html+="<div class=\\"chart-card\\" style=\\"margin-bottom:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">grid_view</i>Email/Meeting Participation by Unit</div><div class=\\"heatmap-grid\\" style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin-top:12px\\">";' +
    'var partUnits=Object.keys(d.participationByUnit).slice(0,12);' +
    'partUnits.forEach(function(u){var p=d.participationByUnit[u];var avgPart=(p.emailRate+p.meetingRate)/2;var heatBg=avgPart>=60?"#86efac":avgPart>=40?"#d9f99d":avgPart>=20?"#fef3c7":"#fee2e2";var textCol=avgPart>=40?"#166534":"#991b1b";html+="<div style=\\"background:"+heatBg+";padding:10px;border-radius:8px\\"><div style=\\"font-weight:600;color:"+textCol+";font-size:12px;margin-bottom:4px\\">"+u+"</div><div style=\\"display:flex;justify-content:space-between;font-size:11px;color:"+textCol+"\\"><span>Email: "+p.emailRate+"%</span><span>Mtg: "+p.meetingRate+"%</span></div><div style=\\"font-size:10px;color:"+textCol+";opacity:0.7;margin-top:2px\\">"+fmt(p.count)+" members</div></div>"});' +
    'if(partUnits.length===0)html+="<p style=\\"color:#64748b;text-align:center;padding:20px;grid-column:1/-1\\">No participation data available</p>";' +
    'html+="</div></div>";' +
    // Participation by Location heatmap
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">place</i>Email/Meeting Participation by Location</div><div class=\\"heatmap-grid\\" style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin-top:12px\\">";' +
    'var partLocs=Object.keys(d.participationByLocation).slice(0,12);' +
    'partLocs.forEach(function(l){var p=d.participationByLocation[l];var avgPart=(p.emailRate+p.meetingRate)/2;var heatBg=avgPart>=60?"#86efac":avgPart>=40?"#d9f99d":avgPart>=20?"#fef3c7":"#fee2e2";var textCol=avgPart>=40?"#166534":"#991b1b";html+="<div style=\\"background:"+heatBg+";padding:10px;border-radius:8px\\"><div style=\\"font-weight:600;color:"+textCol+";font-size:12px;margin-bottom:4px\\">"+l+"</div><div style=\\"display:flex;justify-content:space-between;font-size:11px;color:"+textCol+"\\"><span>Email: "+p.emailRate+"%</span><span>Mtg: "+p.meetingRate+"%</span></div><div style=\\"font-size:10px;color:"+textCol+";opacity:0.7;margin-top:2px\\">"+fmt(p.count)+" members</div></div>"});' +
    'if(partLocs.length===0)html+="<p style=\\"color:#64748b;text-align:center;padding:20px;grid-column:1/-1\\">No participation data available</p>";' +
    'html+="</div></div>";' +
    // Section: Member Satisfaction Analysis
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:24px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#a78bfa\\">sentiment_satisfied</i>Member Satisfaction Analysis</h3>";' +
    'html+="<div class=\\"charts-row\\">";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">work</i>Satisfaction by Unit/Role</div><canvas id=\\"satByUnitChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">location_city</i>Satisfaction by Work Location</div><canvas id=\\"satByLocationChart\\"></canvas></div></div>";' +
    // Satisfaction matrix (unit x location)
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">table_chart</i>Satisfaction by Unit & Location Matrix</div>";' +
    'var satUnits=Object.keys(d.satisfactionByUnit).slice(0,8);var satLocs=Object.keys(d.satisfactionByLocation).slice(0,8);' +
    'if(satUnits.length>0&&satLocs.length>0){' +
    'html+="<div style=\\"overflow-x:auto;margin-top:12px\\"><table style=\\"width:100%;border-collapse:collapse;font-size:11px\\"><tr><th style=\\"padding:8px;background:#1e293b;color:#94a3b8;text-align:left\\">Unit / Location</th>";' +
    'satLocs.forEach(function(l){html+="<th style=\\"padding:8px;background:#1e293b;color:#94a3b8;text-align:center;min-width:80px\\">"+l.substring(0,12)+"</th>"});' +
    'html+="</tr>";' +
    'satUnits.forEach(function(u){html+="<tr><td style=\\"padding:8px;background:#0f172a;color:#e2e8f0;font-weight:500\\">"+u+"</td>";satLocs.forEach(function(l){var uScore=d.satisfactionByUnit[u]?d.satisfactionByUnit[u].score:0;var lScore=d.satisfactionByLocation[l]?d.satisfactionByLocation[l].score:0;var avgS=(uScore+lScore)/2;var cellBg=avgS>=7?"rgba(34,197,94,0.3)":avgS>=5?"rgba(245,158,11,0.3)":"rgba(239,68,68,0.3)";var cellCol=avgS>=7?"#22c55e":avgS>=5?"#f59e0b":"#ef4444";html+="<td style=\\"padding:8px;background:"+cellBg+";color:"+cellCol+";text-align:center;font-weight:600\\">"+avgS.toFixed(1)+"</td>"});html+="</tr>"});' +
    'html+="</table></div>"}' +
    'else{html+="<p style=\\"color:#64748b;text-align:center;padding:30px\\">Not enough satisfaction data for matrix view</p>"}' +
    'html+="</div>";' +
    // Contact Updates Section
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:24px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#f59e0b\\">contact_phone</i>Contact Updates</h3>";' +
    'html+="<div class=\\"charts-row\\">";' +
    'html+="<div class=\\"trend-card\\"><div class=\\"trend-header\\"><span class=\\"trend-title\\"><i class=\\"material-icons\\" style=\\"color:#22c55e\\">update</i>Recent Updates (30 days)</span><span class=\\"trend-value green\\">"+fmt(dt.recentUpdates.length)+"</span></div><div class=\\"trend-list\\">";' +
    'dt.recentUpdates.slice(0,10).forEach(function(m){html+="<div class=\\"trend-item\\"><span>"+m.name+" ("+m.id+")</span><span style=\\"color:#64748b\\">"+new Date(m.date).toLocaleDateString()+"</span></div>"});' +
    'if(dt.recentUpdates.length===0)html+="<p style=\\"color:#64748b;text-align:center;padding:20px\\">No recent updates</p>";' +
    'html+="</div></div>";' +
    'html+="<div class=\\"trend-card\\"><div class=\\"trend-header\\"><span class=\\"trend-title\\"><i class=\\"material-icons\\" style=\\"color:#f59e0b\\">warning</i>Stale Contacts (90+ days)</span><span class=\\"trend-value yellow\\">"+fmt(dt.staleContacts.length)+"</span></div><div class=\\"trend-list\\">";' +
    'dt.staleContacts.slice(0,10).forEach(function(m){html+="<div class=\\"trend-item\\"><span>"+m.name+" ("+m.id+")</span><span style=\\"color:#64748b\\">"+new Date(m.lastUpdate).toLocaleDateString()+"</span></div>"});' +
    'if(dt.staleContacts.length===0)html+="<p style=\\"color:#64748b;text-align:center;padding:20px\\">All contacts up to date!</p>";' +
    'html+="</div></div></div>";' +
    // Meeting Attendees
    'html+="<div class=\\"charts-row\\"><div class=\\"trend-card\\"><div class=\\"trend-header\\"><span class=\\"trend-title\\"><i class=\\"material-icons\\" style=\\"color:#a78bfa\\">groups</i>Recent Meeting Attendees</span><span class=\\"trend-value purple\\">"+fmt(eng.recentMeetingAttendees.length)+"</span></div><div class=\\"trend-list\\">";' +
    'eng.recentMeetingAttendees.slice(0,8).forEach(function(m){html+="<div class=\\"trend-item\\"><span>"+m.name+"</span><span style=\\"color:"+(m.type==="Virtual"?"#60a5fa":"#22c55e")+"\\">"+m.type+"</span></div>"});' +
    'if(eng.recentMeetingAttendees.length===0)html+="<p style=\\"color:#64748b;text-align:center;padding:20px\\">No recent meeting attendance data</p>";' +
    'html+="</div></div><div class=\\"chart-card\\"></div></div>";' +
    'document.getElementById("main-content").innerHTML=html;renderDirectoryCharts()' +
    '}' +

    // Hot Spots Tab - Multi-Type with Explanations
    'else if(tab==="hotspots"){' +
    'var hs=d.hotSpots||{grievance:[],dissatisfaction:[],lowEngagement:[],overdueConcentration:[]};' +
    'var maxCases=d.hotZones.length>0?Math.max.apply(null,d.hotZones.map(function(h){return h.count})):0;' +
    'var minCases=d.hotZones.length>0?Math.min.apply(null,d.hotZones.map(function(h){return h.count})):0;' +
    'function getHeatColor(count){if(maxCases===minCases)return"#FEF3C7";var pct=(count-minCases)/(maxCases-minCases);if(pct<=0.5){var r=Math.round(209+(254-209)*pct*2);var g=Math.round(250+(243-250)*pct*2);var b=Math.round(229+(199-229)*pct*2);return"rgb("+r+","+g+","+b+")"}else{var p=(pct-0.5)*2;var r=Math.round(254+(252-254)*p);var g=Math.round(243+(165-243)*p);var b=Math.round(199+(165-199)*p);return"rgb("+r+","+g+","+b+")"}}' +
    // Explanation header
    'html="<div class=\\"chart-card\\" style=\\"margin-bottom:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">info</i>What are Hot Spots?</div>";' +
    'html+="<p style=\\"color:#94a3b8;margin:0 0 12px;font-size:12px;line-height:1.5\\">Hot Spots identify areas requiring immediate attention. These are locations or units with concentrated issues that may indicate systemic problems, management hostility, or member concerns that need focused resources.</p>";' +
    'html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px\\">";' +
    'html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid #ef4444\\"><div style=\\"font-weight:600;color:#ef4444;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">local_fire_department</i> Grievance Hot Spots</div><div style=\\"color:#94a3b8;font-size:11px;margin-top:4px\\">Locations with 3+ active grievance cases - indicates potential management issues</div></div>";' +
    'html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid #f59e0b\\"><div style=\\"font-weight:600;color:#f59e0b;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">sentiment_dissatisfied</i> Dissatisfaction Hot Spots</div><div style=\\"color:#94a3b8;font-size:11px;margin-top:4px\\">Areas with satisfaction scores below 5/10 - members unhappy with representation</div></div>";' +
    'html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid #a78bfa\\"><div style=\\"font-weight:600;color:#a78bfa;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">trending_down</i> Low Engagement Hot Spots</div><div style=\\"color:#94a3b8;font-size:11px;margin-top:4px\\">Areas with less than 30% email/meeting engagement - outreach needed</div></div>";' +
    'html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid #60a5fa\\"><div style=\\"font-weight:600;color:#60a5fa;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">schedule</i> Overdue Hot Spots</div><div style=\\"color:#94a3b8;font-size:11px;margin-top:4px\\">Locations with 2+ overdue cases - deadline management needed</div></div>";' +
    'html+="</div></div>";' +
    // Summary KPIs
    'var totalHotSpots=(hs.grievance?hs.grievance.length:0)+(hs.dissatisfaction?hs.dissatisfaction.length:0)+(hs.lowEngagement?hs.lowEngagement.length:0)+(hs.overdueConcentration?hs.overdueConcentration.length:0);' +
    'html+="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(4,1fr);margin-bottom:16px\\">";' +
    'html+="<div class=\\"kpi-card "+(hs.grievance&&hs.grievance.length>0?"alert":"")+"\\"><div class=\\"kpi-label\\">Grievance Hot Spots</div><div class=\\"kpi-value "+(hs.grievance&&hs.grievance.length>0?"red":"green")+"\\">"+(hs.grievance?hs.grievance.length:0)+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(hs.dissatisfaction&&hs.dissatisfaction.length>0?"alert":"")+"\\"><div class=\\"kpi-label\\">Dissatisfaction</div><div class=\\"kpi-value "+(hs.dissatisfaction&&hs.dissatisfaction.length>0?"yellow":"green")+"\\">"+(hs.dissatisfaction?hs.dissatisfaction.length:0)+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(hs.lowEngagement&&hs.lowEngagement.length>0?"":"")+"\\"><div class=\\"kpi-label\\">Low Engagement</div><div class=\\"kpi-value "+(hs.lowEngagement&&hs.lowEngagement.length>0?"purple":"green")+"\\">"+(hs.lowEngagement?hs.lowEngagement.length:0)+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(hs.overdueConcentration&&hs.overdueConcentration.length>0?"alert":"")+"\\"><div class=\\"kpi-label\\">Overdue Concentration</div><div class=\\"kpi-value "+(hs.overdueConcentration&&hs.overdueConcentration.length>0?"blue":"green")+"\\">"+(hs.overdueConcentration?hs.overdueConcentration.length:0)+"</div></div>";' +
    'html+="</div>";' +
    // Grievance Hot Zones (original)
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\" style=\\"color:#ef4444\\">local_fire_department</i>Grievance Hot Zones (3+ Active Cases)</div>";' +
    'html+="<div style=\\"display:flex;gap:4px;margin:12px 0;font-size:10px;align-items:center\\"><span style=\\"color:#94a3b8\\">Low</span><div style=\\"display:flex;height:12px\\"><div style=\\"width:20px;background:#D1FAE5\\"></div><div style=\\"width:20px;background:#FEF3C7\\"></div><div style=\\"width:20px;background:#FCA5A5\\"></div></div><span style=\\"color:#94a3b8\\">High</span></div>";' +
    'if(d.hotZones.length===0){html+="<div style=\\"text-align:center;padding:30px;color:#22c55e\\"><i class=\\"material-icons\\" style=\\"font-size:36px\\">check_circle</i><p style=\\"margin-top:8px;font-size:13px\\">No grievance hot zones - All clear!</p></div>"}' +
    'else{d.hotZones.sort(function(a,b){return b.count-a.count}).forEach(function(h){var bgColor=getHeatColor(h.count);var textColor=(h.count-minCases)/(maxCases-minCases||1)>0.6?"#fff":"#1e293b";html+="<div class=\\"hot-zone\\" style=\\"background:"+bgColor+";border-left-color:"+bgColor+"\\"><span style=\\"color:"+textColor+"\\">"+h.location+"</span><span class=\\"badge\\" style=\\"background:rgba(0,0,0,0.2);color:"+textColor+"\\">"+h.count+" cases</span></div>"})}' +
    'html+="</div>";' +
    // Dissatisfaction Hot Spots
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\" style=\\"color:#f59e0b\\">sentiment_dissatisfied</i>Dissatisfaction Hot Spots (Score < 5)</div>";' +
    'if(!hs.dissatisfaction||hs.dissatisfaction.length===0){html+="<div style=\\"text-align:center;padding:30px;color:#22c55e\\"><i class=\\"material-icons\\" style=\\"font-size:36px\\">sentiment_satisfied</i><p style=\\"margin-top:8px;font-size:13px\\">No dissatisfaction hot spots - Members are satisfied!</p></div>"}' +
    'else{html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;margin-top:12px\\">";hs.dissatisfaction.forEach(function(h){var scoreColor=h.score<3?"#ef4444":"#f59e0b";html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid "+scoreColor+"\\"><div style=\\"font-weight:600;color:#e2e8f0;font-size:13px\\">"+h.name+"</div><div style=\\"display:flex;justify-content:space-between;margin-top:6px\\"><span style=\\"color:#94a3b8;font-size:11px\\">"+h.count+" responses</span><span style=\\"color:"+scoreColor+";font-weight:700\\">"+h.score+"/10</span></div></div>"});html+="</div>"}' +
    'html+="</div>";' +
    // Low Engagement Hot Spots
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\" style=\\"color:#a78bfa\\">trending_down</i>Low Engagement Hot Spots (< 30%)</div>";' +
    'if(!hs.lowEngagement||hs.lowEngagement.length===0){html+="<div style=\\"text-align:center;padding:30px;color:#22c55e\\"><i class=\\"material-icons\\" style=\\"font-size:36px\\">groups</i><p style=\\"margin-top:8px;font-size:13px\\">No low engagement areas - Good outreach!</p></div>"}' +
    'else{html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;margin-top:12px\\">";hs.lowEngagement.forEach(function(h){var engColor=h.engagement<15?"#ef4444":"#a78bfa";html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid "+engColor+"\\"><div style=\\"font-weight:600;color:#e2e8f0;font-size:13px\\">"+h.name+"</div><div style=\\"color:#64748b;font-size:10px\\">"+h.type+"</div><div style=\\"display:flex;justify-content:space-between;margin-top:6px\\"><span style=\\"color:#94a3b8;font-size:11px\\">"+h.count+" members</span><span style=\\"color:"+engColor+";font-weight:700\\">"+h.engagement+"%</span></div></div>"});html+="</div>"}' +
    'html+="</div>";' +
    // Cases by Location chart
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">location_on</i>Cases by Location</div><canvas id=\\"hotspotChart\\"></canvas></div>";' +
    'document.getElementById("main-content").innerHTML=html;renderHotspotChart()' +
    '}' +

    // Bargaining Tab (Comprehensive)
    'else if(tab==="bargaining"){' +
    // KPI Grid - Key Metrics
    'html="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(6,1fr);margin-bottom:16px\\">";' +
    'html+="<div class=\\"kpi-card "+(d.step1DenialRate>60?"alert":"")+"\\"><div class=\\"kpi-label\\">Step 1 Denial Rate</div><div class=\\"kpi-value "+(d.step1DenialRate>60?"red":"yellow")+"\\">"+d.step1DenialRate+"%</div><div style=\\"font-size:9px;color:#64748b\\">"+(d.step1DenialRate>60?"High Hostility":"Normal Range")+"</div></div>";' +
    'html+="<div class=\\"kpi-card "+(d.step2DenialRate>50?"alert":"")+"\\"><div class=\\"kpi-label\\">Step 2 Denial Rate</div><div class=\\"kpi-value "+(d.step2DenialRate>50?"red":"yellow")+"\\">"+(d.step2DenialRate||0)+"%</div><div style=\\"font-size:9px;color:#64748b\\">"+(d.step2DenialRate>50?"Escalation Pattern":"Expected Range")+"</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Avg Settlement Time</div><div class=\\"kpi-value blue\\">"+d.avgSettlementDays+"</div><div style=\\"font-size:9px;color:#64748b\\">Days to Resolution</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Most Violated Article</div><div class=\\"kpi-value purple\\" style=\\"font-size:16px\\">"+d.topViolatedArticle+"</div><div style=\\"font-size:9px;color:#64748b\\">Focus Area</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Grievances</div><div class=\\"kpi-value blue\\">"+d.totalGrievances+"</div><div style=\\"font-size:9px;color:#64748b\\">All Time</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Win Rate</div><div class=\\"kpi-value green\\">"+d.winRate+"%</div><div style=\\"font-size:9px;color:#64748b\\">Won + Settled</div></div>";' +
    'html+="</div>";' +
    // Step Progression Detail
    'html+="<div class=\\"chart-card\\" style=\\"margin-bottom:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">stairs</i>Grievance Step Progression - Detailed View</div>";' +
    'html+="<p style=\\"color:#94a3b8;font-size:11px;margin:0 0 12px\\">Track cases through each step of the grievance process with outcome breakdown.</p>";' +
    'html+="<div style=\\"overflow-x:auto\\"><table style=\\"width:100%;border-collapse:collapse;font-size:12px\\"><thead><tr style=\\"background:#1e293b\\"><th style=\\"padding:10px;text-align:left;color:#94a3b8\\">Step</th><th style=\\"padding:10px;text-align:center;color:#60a5fa\\">Active Cases</th><th style=\\"padding:10px;text-align:center;color:#22c55e\\">Won</th><th style=\\"padding:10px;text-align:center;color:#ef4444\\">Denied</th><th style=\\"padding:10px;text-align:center;color:#a78bfa\\">Settled</th><th style=\\"padding:10px;text-align:center;color:#f59e0b\\">Pending</th><th style=\\"padding:10px;text-align:center;color:#64748b\\">Withdrawn</th><th style=\\"padding:10px;text-align:center;color:#22c55e\\">Success Rate</th></tr></thead><tbody>";' +
    'var steps=[{name:"Step 1 - Informal",key:"step1",desc:"Initial meeting with supervisor"},{name:"Step 2 - Formal",key:"step2",desc:"Written grievance to management"},{name:"Step 3 - Review",key:"step3",desc:"Higher-level management review"},{name:"Arbitration",key:"arb",desc:"Third-party arbitrator decision"}];' +
    'steps.forEach(function(s){var count=d.stepProgression[s.key]||0;var outcomes=d.stepOutcomes?d.stepOutcomes[s.key]:{won:0,denied:0,settled:0,pending:0,withdrawn:0};var total=(outcomes.won||0)+(outcomes.denied||0)+(outcomes.settled||0);var successRate=total>0?Math.round(((outcomes.won||0)+(outcomes.settled||0))/total*100):0;html+="<tr style=\\"border-bottom:1px solid #1e293b\\"><td style=\\"padding:10px\\"><div style=\\"color:#e2e8f0;font-weight:600\\">"+s.name+"</div><div style=\\"color:#64748b;font-size:10px\\">"+s.desc+"</div></td><td style=\\"padding:10px;text-align:center;color:#60a5fa;font-weight:700;font-size:16px\\">"+count+"</td><td style=\\"padding:10px;text-align:center;color:#22c55e\\">"+(outcomes.won||0)+"</td><td style=\\"padding:10px;text-align:center;color:#ef4444\\">"+(outcomes.denied||0)+"</td><td style=\\"padding:10px;text-align:center;color:#a78bfa\\">"+(outcomes.settled||0)+"</td><td style=\\"padding:10px;text-align:center;color:#f59e0b\\">"+(outcomes.pending||0)+"</td><td style=\\"padding:10px;text-align:center;color:#64748b\\">"+(outcomes.withdrawn||0)+"</td><td style=\\"padding:10px;text-align:center\\"><span style=\\"background:"+(successRate>=50?"rgba(34,197,94,0.2)":"rgba(239,68,68,0.2)")+";color:"+(successRate>=50?"#22c55e":"#ef4444")+";padding:4px 8px;border-radius:4px;font-weight:600\\">"+successRate+"%</span></td></tr>"});' +
    'html+="</tbody></table></div></div>";' +
    // Cases at Each Step - Detail lists
    'html+="<div class=\\"charts-row\\">";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\" style=\\"color:#f59e0b\\">folder_open</i>Cases at Step 1 ("+d.stepProgression.step1+")</div><div class=\\"list-container\\" style=\\"max-height:200px\\">";' +
    'if(d.stepCaseDetails&&d.stepCaseDetails.step1&&d.stepCaseDetails.step1.length>0){d.stepCaseDetails.step1.slice(0,8).forEach(function(c){html+="<div class=\\"list-item\\"><span style=\\"font-weight:600\\">"+c.id+"</span><span style=\\"color:#94a3b8;font-size:11px\\">"+c.category+" | "+c.location+"</span></div>"})}else{html+="<p style=\\"color:#64748b;text-align:center;padding:20px;font-size:12px\\">No cases at Step 1</p>"}' +
    'html+="</div></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\" style=\\"color:#ef4444\\">folder_open</i>Cases at Step 2 ("+d.stepProgression.step2+")</div><div class=\\"list-container\\" style=\\"max-height:200px\\">";' +
    'if(d.stepCaseDetails&&d.stepCaseDetails.step2&&d.stepCaseDetails.step2.length>0){d.stepCaseDetails.step2.slice(0,8).forEach(function(c){html+="<div class=\\"list-item\\"><span style=\\"font-weight:600\\">"+c.id+"</span><span style=\\"color:#94a3b8;font-size:11px\\">"+c.category+" | "+c.location+"</span></div>"})}else{html+="<p style=\\"color:#64748b;text-align:center;padding:20px;font-size:12px\\">No cases at Step 2</p>"}' +
    'html+="</div></div></div>";' +
    // Charts row
    'html+="<div class=\\"charts-row\\"><div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">article</i>Violations by Article</div><canvas id=\\"bargainChart\\"></canvas></div>";' +
    'html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">donut_large</i>Step Distribution</div><canvas id=\\"stepDistChart\\"></canvas></div></div>";' +
    // Recent Grievances
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">history</i>Recent Grievances</div><div class=\\"list-container\\" style=\\"max-height:250px\\">";' +
    'if(d.recentGrievances&&d.recentGrievances.length>0){d.recentGrievances.forEach(function(g){var statusColor=g.status.toLowerCase()==="won"?"#22c55e":g.status.toLowerCase()==="denied"?"#ef4444":g.status.toLowerCase()==="settled"?"#a78bfa":"#f59e0b";html+="<div class=\\"list-item\\" style=\\"flex-wrap:wrap\\"><div style=\\"display:flex;justify-content:space-between;width:100%\\"><span style=\\"font-weight:600\\">"+g.id+"</span><span class=\\"badge\\" style=\\"background:"+statusColor+";color:white\\">"+g.status+"</span></div><div style=\\"width:100%;margin-top:4px;font-size:11px;color:#94a3b8\\">"+g.member+" | "+g.category+" | "+g.location+"</div></div>"})}else{html+="<p style=\\"color:#64748b;text-align:center;padding:20px\\">No recent grievances</p>"}' +
    'html+="</div></div>";' +
    'document.getElementById("main-content").innerHTML=html;renderBargainCharts()' +
    '}' +

    // Satisfaction Tab (Enhanced with Insights)
    'else if(tab==="satisfaction"){' +
    'var sat=d.satisfactionData||{responseCount:0,sections:[]};' +
    'var surveyRate=d.engagement?d.engagement.surveyResponseRate:0;' +
    'html="<div class=\\"sat-header\\"><h2 style=\\"color:#e2e8f0;font-size:16px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#22c55e\\">sentiment_satisfied</i>Member Satisfaction Analysis</h2><span class=\\"sat-response-count\\">"+sat.responseCount+" Responses ("+surveyRate+"% response rate)</span></div>";' +
    'if(!sat.sections||sat.sections.length===0){html+="<div class=\\"chart-card\\" style=\\"text-align:center;padding:60px\\"><i class=\\"material-icons\\" style=\\"font-size:48px;color:#64748b\\">poll</i><p style=\\"color:#94a3b8;margin-top:16px\\">No satisfaction survey data available yet.</p></div>"}' +
    'else{' +
    // Calculate insights
    'var avgScore=sat.sections.reduce(function(s,sec){return s+sec.score},0)/sat.sections.length;' +
    'var sorted=sat.sections.slice().sort(function(a,b){return b.score-a.score});' +
    'var best=sorted[0];var worst=sorted[sorted.length-1];' +
    'var avgColor=avgScore>=7?"#22c55e":avgScore>=5?"#f59e0b":"#ef4444";' +
    // KPI row with overall score and insights
    'html+="<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(4,1fr);margin-bottom:20px\\">";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Overall Score</div><div class=\\"kpi-value\\" style=\\"color:"+avgColor+"\\">"+avgScore.toFixed(1)+"</div><div style=\\"font-size:10px;color:#64748b\\">out of 10</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Best Area</div><div style=\\"font-size:14px;font-weight:700;color:#22c55e\\">"+best.name+"</div><div style=\\"font-size:10px;color:#64748b\\">"+best.score+"/10</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Needs Focus</div><div style=\\"font-size:14px;font-weight:700;color:"+(worst.score<5?"#ef4444":"#f59e0b")+"\\">"+worst.name+"</div><div style=\\"font-size:10px;color:#64748b\\">"+worst.score+"/10</div></div>";' +
    'html+="<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Morale Score</div><div class=\\"kpi-value\\" style=\\"color:"+(d.moraleScore>=7?"#22c55e":d.moraleScore>=5?"#f59e0b":"#ef4444")+"\\">"+d.moraleScore+"</div><div style=\\"font-size:10px;color:#64748b\\">Trust Index</div></div></div>";' +
    // Insights box
    'html+="<div class=\\"chart-card\\" style=\\"margin-bottom:16px;background:linear-gradient(135deg,rgba(96,165,250,0.1),rgba(139,92,246,0.1))\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">lightbulb</i>Key Insights</div><div style=\\"display:grid;grid-template-columns:1fr 1fr;gap:12px\\">";' +
    'html+="<div style=\\"padding:12px;background:rgba(34,197,94,0.1);border-radius:8px;border-left:3px solid #22c55e\\"><div style=\\"font-size:11px;color:#22c55e;font-weight:600\\">STRENGTH</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+best.name+" rated highest at "+best.score+"/10</div></div>";' +
    'html+="<div style=\\"padding:12px;background:rgba("+(worst.score<5?"239,68,68":"245,158,11")+",0.1);border-radius:8px;border-left:3px solid "+(worst.score<5?"#ef4444":"#f59e0b")+"\\"><div style=\\"font-size:11px;color:"+(worst.score<5?"#ef4444":"#f59e0b")+";font-weight:600\\">OPPORTUNITY</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+worst.name+" at "+worst.score+"/10 needs attention</div></div>";' +
    'var highCount=sat.sections.filter(function(s){return s.score>=7}).length;var lowCount=sat.sections.filter(function(s){return s.score<5}).length;' +
    'html+="<div style=\\"padding:12px;background:rgba(96,165,250,0.1);border-radius:8px;border-left:3px solid #60a5fa\\"><div style=\\"font-size:11px;color:#60a5fa;font-weight:600\\">SUMMARY</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+highCount+" of 8 areas rated Good (7+)</div></div>";' +
    'html+="<div style=\\"padding:12px;background:rgba(139,92,246,0.1);border-radius:8px;border-left:3px solid #8b5cf6\\"><div style=\\"font-size:11px;color:#8b5cf6;font-weight:600\\">ACTION ITEMS</div><div style=\\"font-size:13px;color:#e2e8f0;margin-top:4px\\">"+(lowCount>0?lowCount+" areas need immediate improvement":"All areas above minimum threshold")+"</div></div></div></div>";' +
    // Section details grid with individual question breakdown
    'html+="<h3 style=\\"color:#e2e8f0;font-size:14px;margin:16px 0 12px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">assessment</i>Section Breakdown with Question Scores</h3>";' +
    'html+="<div class=\\"sat-grid\\">";' +
    'sat.sections.forEach(function(section,sIdx){' +
    'var scoreColor=section.score>=7?"#22c55e":section.score>=5?"#f59e0b":"#ef4444";' +
    'var pct=(section.score/10)*100;' +
    'var rank=sorted.indexOf(section)+1;' +
    'html+="<div class=\\"sat-section\\" style=\\"cursor:pointer\\" onclick=\\"toggleSectionDetail("+sIdx+")\\"><div class=\\"sat-section-header\\"><span class=\\"sat-section-name\\">"+section.name+"<span style=\\"font-size:10px;color:#64748b;margin-left:6px\\">#"+rank+"</span></span><span class=\\"sat-section-score\\" style=\\"color:"+scoreColor+"\\">"+(section.score||0)+"/10 <i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">expand_more</i></span></div>";' +
    'html+="<div class=\\"sat-score-bar\\"><div class=\\"sat-score-fill\\" style=\\"width:"+pct+"%;background:"+scoreColor+"\\"></div></div>";' +
    // Individual question breakdown
    'html+="<div id=\\"section-detail-"+sIdx+"\\" class=\\"sat-questions-detail\\" style=\\"display:none;margin-top:12px;border-top:1px solid #334155;padding-top:12px\\">";' +
    'if(section.questions&&section.questions.length>0){' +
    'section.questions.forEach(function(q){' +
    'var qScore=q.score||0;var qColor=qScore>=7?"#22c55e":qScore>=5?"#f59e0b":"#ef4444";var qPct=(qScore/10)*100;' +
    'html+="<div style=\\"margin-bottom:8px\\"><div style=\\"display:flex;justify-content:space-between;align-items:center;margin-bottom:4px\\"><span style=\\"font-size:11px;color:#cbd5e1\\">"+q.label+"</span><span style=\\"font-size:12px;font-weight:600;color:"+qColor+"\\">"+qScore.toFixed(1)+"</span></div><div style=\\"height:4px;background:rgba(255,255,255,0.1);border-radius:2px\\"><div style=\\"height:100%;width:"+qPct+"%;background:"+qColor+";border-radius:2px\\"></div></div></div>"' +
    '})}' +
    'html+="</div></div>"});' +
    'html+="</div>";' +
    // Full Question Breakdown Table
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:20px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">list_alt</i>Complete Survey Question Breakdown</div>";' +
    'html+="<p style=\\"color:#94a3b8;font-size:11px;margin:0 0 12px\\">All "+sat.responseCount+" survey responses analyzed. Click a section above to expand individual questions.</p>";' +
    'html+="<div style=\\"overflow-x:auto;max-height:400px;overflow-y:auto\\"><table style=\\"width:100%;border-collapse:collapse;font-size:11px\\"><thead style=\\"position:sticky;top:0;background:#0f172a\\"><tr><th style=\\"padding:10px;text-align:left;color:#94a3b8;border-bottom:1px solid #334155\\">Section</th><th style=\\"padding:10px;text-align:left;color:#94a3b8;border-bottom:1px solid #334155\\">Question</th><th style=\\"padding:10px;text-align:center;color:#94a3b8;border-bottom:1px solid #334155;width:80px\\">Score</th><th style=\\"padding:10px;text-align:left;color:#94a3b8;border-bottom:1px solid #334155;width:120px\\">Rating</th></tr></thead><tbody>";' +
    'sat.sections.forEach(function(section){' +
    'if(section.questions&&section.questions.length>0){' +
    'section.questions.forEach(function(q,qIdx){' +
    'var qScore=q.score||0;var qColor=qScore>=7?"#22c55e":qScore>=5?"#f59e0b":"#ef4444";var qPct=(qScore/10)*100;' +
    'html+="<tr style=\\"border-bottom:1px solid #1e293b\\"><td style=\\"padding:8px;color:#64748b\\">"+(qIdx===0?section.name:"")+"</td><td style=\\"padding:8px;color:#e2e8f0\\">"+q.label+"</td><td style=\\"padding:8px;text-align:center;font-weight:600;color:"+qColor+"\\">"+qScore.toFixed(1)+"</td><td style=\\"padding:8px\\"><div style=\\"height:6px;background:rgba(255,255,255,0.1);border-radius:3px;min-width:80px\\"><div style=\\"height:100%;width:"+qPct+"%;background:"+qColor+";border-radius:3px\\"></div></div></td></tr>"' +
    '})}});' +
    'html+="</tbody></table></div></div>";' +
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:20px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">bar_chart</i>Section Scores Comparison</div><canvas id=\\"satChart\\"></canvas></div>"}' +
    'document.getElementById("main-content").innerHTML=html;if(sat.sections&&sat.sections.length>0)renderSatisfactionChart()' +
    '}' +

    // Resources Tab (Enhanced with Steward Directory)
    'else if(tab==="resources"){' +
    'var dr=d.driveResources;' +
    'html="<h2 style=\\"color:#e2e8f0;font-size:16px;margin-bottom:20px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">folder</i>Union Resources</h2>";' +
    'html+="<div class=\\"resource-grid\\">";' +
    'if(dr.folderUrl){html+="<a href=\\""+dr.folderUrl+"\\" target=\\"_blank\\" class=\\"resource-card\\"><div class=\\"resource-icon\\">📁</div><div class=\\"resource-title\\">Google Drive Folder</div><div class=\\"resource-desc\\">Access all union documents</div></a>"}' +
    'html+="<div class=\\"resource-card\\" onclick=\\"alert(\\x27Coming soon: Contract PDF\\x27)\\"><div class=\\"resource-icon\\">📜</div><div class=\\"resource-title\\">Contract</div><div class=\\"resource-desc\\">Current collective bargaining agreement</div></div>";' +
    'html+="<div class=\\"resource-card\\" onclick=\\"alert(\\x27Coming soon: Forms\\x27)\\"><div class=\\"resource-icon\\">📝</div><div class=\\"resource-title\\">Forms</div><div class=\\"resource-desc\\">Grievance forms and templates</div></div>";' +
    'html+="</div>";' +
    // Steward Contact Directory with Smart Search
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:20px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">contacts</i>Steward Contact Directory</div>";' +
    'html+="<div style=\\"position:relative\\"><input type=\\"text\\" id=\\"stewardSearch\\" placeholder=\\"Start typing to search stewards...\\" oninput=\\"smartFilterStewards()\\" onfocus=\\"showSearchSuggestions()\\" autocomplete=\\"off\\" style=\\"width:100%;padding:10px 10px 10px 36px;margin:12px 0;border:1px solid #475569;border-radius:8px;background:#1e293b;color:#f8fafc;font-size:13px\\"><i class=\\"material-icons\\" style=\\"position:absolute;left:10px;top:50%;transform:translateY(-50%);color:#64748b;font-size:18px\\">search</i><div id=\\"searchSuggestions\\" style=\\"display:none;position:absolute;top:100%;left:0;right:0;background:#1e293b;border:1px solid #475569;border-radius:8px;max-height:200px;overflow-y:auto;z-index:100;box-shadow:0 4px 12px rgba(0,0,0,0.3)\\"></div></div>";' +
    'html+="<div style=\\"display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap\\"><span style=\\"color:#64748b;font-size:11px\\">Quick filters:</span>";' +
    'var locations=[...new Set(d.stewardList.map(function(s){return s.location}))].filter(Boolean).slice(0,5);' +
    'locations.forEach(function(loc){html+="<button class=\\"btn btn-sm\\" style=\\"font-size:10px;padding:4px 8px\\" onclick=\\"quickFilterSteward(\\x27"+loc+"\\x27)\\">"+loc+"</button>"});' +
    'html+="<button class=\\"btn btn-sm\\" style=\\"font-size:10px;padding:4px 8px;background:#475569\\" onclick=\\"clearStewardFilter()\\">Clear</button></div>";' +
    'html+="<div id=\\"stewardList\\" class=\\"list-container\\" style=\\"max-height:300px\\">";' +
    'if(d.stewardList&&d.stewardList.length>0){d.stewardList.forEach(function(s){' +
    'html+="<div class=\\"steward-contact list-item\\" data-search=\\""+((s.name||"")+" "+(s.location||"")+" "+(s.unit||"")).toLowerCase()+"\\" style=\\"flex-wrap:wrap\\"><div style=\\"display:flex;justify-content:space-between;width:100%\\"><span style=\\"font-weight:600;color:#e2e8f0\\">"+s.name+"</span><button class=\\"btn btn-sm\\" style=\\"background:#22c55e;color:white;padding:4px 8px\\" onclick=\\"saveStewardContact(\\x27"+s.name+"\\x27,\\x27"+(s.email||"")+"\\x27,\\x27"+(s.phone||"")+"\\x27)\\">Save Contact</button></div>";' +
    'html+="<div style=\\"width:100%;margin-top:6px;font-size:12px;color:#94a3b8\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">location_on</i> "+s.location+" | "+s.unit+"</div>";' +
    'if(s.email){html+="<div style=\\"width:100%;margin-top:4px;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle;color:#60a5fa\\">email</i> <a href=\\"mailto:"+s.email+"\\" style=\\"color:#60a5fa\\">"+s.email+"</a></div>"}' +
    'if(s.phone){html+="<div style=\\"width:100%;margin-top:4px;font-size:12px\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle;color:#22c55e\\">phone</i> <a href=\\"tel:"+s.phone+"\\" style=\\"color:#22c55e\\">"+s.phone+"</a></div>"}' +
    'html+="</div>"})}' +
    'else{html+="<p style=\\"color:#94a3b8;text-align:center;padding:20px\\">No steward contacts available</p>"}' +
    'html+="</div></div>";' +
    'if(dr.recentFiles&&dr.recentFiles.length>0){' +
    'html+="<div class=\\"chart-card\\" style=\\"margin-top:20px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">description</i>Recent Files</div><div class=\\"list-container\\">";' +
    'dr.recentFiles.forEach(function(f){html+="<a href=\\""+f.url+"\\" target=\\"_blank\\" class=\\"list-item\\" style=\\"text-decoration:none;color:inherit\\"><span>"+f.name+"</span><span class=\\"badge\\" style=\\"background:#475569;color:white\\">Open</span></a>"});' +
    'html+="</div></div>"}' +
    'document.getElementById("main-content").innerHTML=html' +
    '}' +

    // Meeting Notes Tab - Chronological meeting notes with view-only links
    'else if(tab==="meetingnotes"){' +
    'var notes=d.meetingNotes||[];' +
    'html="<h2 style=\\"color:#e2e8f0;font-size:16px;margin-bottom:20px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">description</i>Meeting Notes</h2>";' +
    'html+="<div class=\\"chart-card\\" style=\\"margin-bottom:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">info</i>About Meeting Notes</div><p style=\\"color:#94a3b8;font-size:12px;line-height:1.6\\">Meeting notes are published the day after each meeting. Click the link to view the notes (read-only). For questions, contact your steward.</p></div>";' +
    'if(notes.length===0){html+="<div class=\\"chart-card\\" style=\\"text-align:center;padding:60px\\"><i class=\\"material-icons\\" style=\\"font-size:48px;color:#64748b\\">notes</i><p style=\\"color:#94a3b8;margin-top:16px\\">No meeting notes available yet.</p></div>"}' +
    'else{html+="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">history</i>Meeting History ("+notes.length+" meetings)</div>";' +
    'html+="<input type=\\"text\\" id=\\"notesSearch\\" placeholder=\\"Search meetings...\\" oninput=\\"filterMeetingNotes()\\" style=\\"width:100%;padding:10px 10px 10px 36px;margin:12px 0;border:1px solid #475569;border-radius:8px;background:#1e293b;color:#f8fafc;font-size:13px\\">";' +
    'html+="<div id=\\"notesList\\" class=\\"list-container\\" style=\\"max-height:500px\\">";' +
    'notes.forEach(function(n){' +
    'html+="<div class=\\"meeting-note-item list-item\\" data-search=\\""+((n.name||"")+" "+(n.date||"")+" "+(n.type||"")).toLowerCase()+"\\" style=\\"flex-wrap:wrap;padding:16px\\"><div style=\\"display:flex;justify-content:space-between;width:100%;align-items:center\\"><div><span style=\\"font-weight:600;color:#e2e8f0;font-size:14px\\">"+escapeHtml(n.name)+"</span></div><span class=\\"badge\\" style=\\"background:#475569;color:#e2e8f0\\">"+escapeHtml(n.type)+"</span></div>";' +
    'html+="<div style=\\"width:100%;margin-top:8px;display:flex;justify-content:space-between;align-items:center\\"><span style=\\"font-size:12px;color:#94a3b8\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle\\">event</i> "+escapeHtml(n.date)+"</span><a href=\\""+n.notesUrl+"\\" target=\\"_blank\\" class=\\"btn btn-sm\\" style=\\"background:#3b82f6;color:white;text-decoration:none;display:flex;align-items:center;gap:4px\\"><i class=\\"material-icons\\" style=\\"font-size:14px\\">open_in_new</i>View Notes</a></div></div>"});' +
    'html+="</div></div>"}' +
    'document.getElementById("main-content").innerHTML=html' +
    '}' +

    // Compare Tab - Comprehensive Dashboard Comparison Tool
    'else if(tab==="compare"){' +
    'function fmt(n){return typeof n==="number"?(n>=1000?n.toLocaleString():n):n}' +
    'function getChangeIndicator(curr,prev,inverse){if(!prev||prev===0)return"";var pct=Math.round(((curr-prev)/Math.abs(prev))*100);var isUp=pct>0;var color=inverse?(isUp?"#ef4444":"#22c55e"):(isUp?"#22c55e":"#ef4444");if(pct===0)color="#64748b";return" <span style=\\"font-size:10px;color:"+color+"\\">"+(isUp?"+":"")+pct+"%</span>"}' +
    // Header with explanation
    'html="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">compare_arrows</i>Dashboard Comparison</div>";' +
    'html+="<div style=\\"background:#1e293b;padding:16px;border-radius:8px;margin-bottom:16px\\"><h3 style=\\"color:#e2e8f0;margin:0 0 8px 0;font-size:14px\\">What is this?</h3><p style=\\"color:#94a3b8;margin:0;font-size:12px;line-height:1.5\\">The Compare tab analyzes key metrics across different dimensions: <strong>current vs previous period</strong>, <strong>step-by-step grievance progression</strong>, <strong>satisfaction breakdowns</strong>, and <strong>denial rates</strong>. Use this to identify patterns, track improvements, and prepare for bargaining.</p></div>";' +
    // Section 1: Current vs Previous Period
    'html+="<h3 style=\\"color:#e2e8f0;margin:20px 0 12px;font-size:14px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">trending_up</i>Current Period vs Previous 30 Days</h3>";' +
    'html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px\\">";' +
    'var compareMetrics=[{label:"Open Cases",curr:d.openGrievances,prev:d.prevOpenCases,inverse:true},{label:"Win Rate",curr:d.winRate,prev:d.prevWinRate,suffix:"%"},{label:"Overdue Cases",curr:d.overdueCount,prev:d.prevOverdueCount,inverse:true},{label:"Morale Score",curr:d.moraleScore,prev:d.prevMoraleScore}];' +
    'compareMetrics.forEach(function(m){var suffix=m.suffix||"";html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;text-align:center\\"><div style=\\"color:#94a3b8;font-size:11px;margin-bottom:4px\\">"+m.label+"</div><div style=\\"display:flex;justify-content:center;align-items:baseline;gap:8px\\"><span style=\\"font-size:24px;font-weight:700;color:#e2e8f0\\">"+m.curr+suffix+"</span>"+getChangeIndicator(m.curr,m.prev,m.inverse)+"</div><div style=\\"color:#64748b;font-size:10px;margin-top:4px\\">Previous: "+(m.prev||0)+suffix+"</div></div>"});' +
    'html+="</div>";' +
    // Section 2: Grievance Step Comparison
    'html+="<h3 style=\\"color:#e2e8f0;margin:20px 0 12px;font-size:14px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#a78bfa\\">stairs</i>Grievance Step-by-Step Comparison</h3>";' +
    'html+="<div style=\\"overflow-x:auto\\"><table style=\\"width:100%;border-collapse:collapse;font-size:12px\\"><thead><tr style=\\"background:#1e293b\\"><th style=\\"padding:12px;text-align:left;color:#94a3b8;border-bottom:1px solid #334155\\">Step</th><th style=\\"padding:12px;text-align:center;color:#94a3b8;border-bottom:1px solid #334155\\">Active</th><th style=\\"padding:12px;text-align:center;color:#22c55e;border-bottom:1px solid #334155\\">Won</th><th style=\\"padding:12px;text-align:center;color:#ef4444;border-bottom:1px solid #334155\\">Denied</th><th style=\\"padding:12px;text-align:center;color:#a78bfa;border-bottom:1px solid #334155\\">Settled</th><th style=\\"padding:12px;text-align:center;color:#f59e0b;border-bottom:1px solid #334155\\">Pending</th></tr></thead><tbody>";' +
    'var steps=[{name:"Step 1",key:"step1",count:d.stepProgression.step1},{name:"Step 2",key:"step2",count:d.stepProgression.step2},{name:"Step 3",key:"step3",count:d.stepProgression.step3},{name:"Arbitration",key:"arb",count:d.stepProgression.arb}];' +
    'steps.forEach(function(s){var outcomes=d.stepOutcomes?d.stepOutcomes[s.key]:{won:0,denied:0,settled:0,pending:0};html+="<tr style=\\"border-bottom:1px solid #1e293b\\"><td style=\\"padding:12px;color:#e2e8f0;font-weight:600\\">"+s.name+"</td><td style=\\"padding:12px;text-align:center;color:#60a5fa;font-weight:700\\">"+s.count+"</td><td style=\\"padding:12px;text-align:center;color:#22c55e\\">"+(outcomes.won||0)+"</td><td style=\\"padding:12px;text-align:center;color:#ef4444\\">"+(outcomes.denied||0)+"</td><td style=\\"padding:12px;text-align:center;color:#a78bfa\\">"+(outcomes.settled||0)+"</td><td style=\\"padding:12px;text-align:center;color:#f59e0b\\">"+(outcomes.pending||0)+"</td></tr>"});' +
    'html+="</tbody></table></div>";' +
    // Section 3: Satisfaction Section Comparison
    'html+="<h3 style=\\"color:#e2e8f0;margin:24px 0 12px;font-size:14px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#22c55e\\">sentiment_satisfied</i>Satisfaction Section Comparison</h3>";' +
    'var sat=d.satisfactionData;' +
    'if(sat&&sat.sections&&sat.sections.length>0){' +
    'html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:20px\\">";' +
    'sat.sections.forEach(function(sec){var scoreColor=sec.score>=7?"#22c55e":sec.score>=5?"#f59e0b":"#ef4444";html+="<div style=\\"background:#1e293b;padding:12px;border-radius:8px;border-left:3px solid "+scoreColor+"\\"><div style=\\"color:#94a3b8;font-size:10px;margin-bottom:4px\\">"+sec.name+"</div><div style=\\"font-size:20px;font-weight:700;color:"+scoreColor+"\\">"+sec.score+"<span style=\\"font-size:11px;color:#64748b\\">/10</span></div></div>"});' +
    'html+="</div>"}' +
    'else{html+="<p style=\\"color:#64748b;text-align:center;padding:20px\\">No satisfaction data available</p>"}' +
    // Section 4: Denial Rate Analysis
    'html+="<h3 style=\\"color:#e2e8f0;margin:24px 0 12px;font-size:14px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#ef4444\\">block</i>Denial Rate Analysis</h3>";' +
    'html+="<div style=\\"display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:20px\\">";' +
    'html+="<div style=\\"background:linear-gradient(135deg,rgba(239,68,68,0.15),rgba(239,68,68,0.05));padding:16px;border-radius:8px;text-align:center\\"><div style=\\"color:#94a3b8;font-size:11px\\">Step 1 Denial Rate</div><div style=\\"font-size:28px;font-weight:700;color:"+(d.step1DenialRate>60?"#ef4444":"#f59e0b")+"\\">"+d.step1DenialRate+"%</div><div style=\\"color:#64748b;font-size:10px\\">"+(d.step1DenialRate>60?"High Hostility":"Normal Range")+"</div></div>";' +
    'html+="<div style=\\"background:linear-gradient(135deg,rgba(239,68,68,0.15),rgba(239,68,68,0.05));padding:16px;border-radius:8px;text-align:center\\"><div style=\\"color:#94a3b8;font-size:11px\\">Step 2 Denial Rate</div><div style=\\"font-size:28px;font-weight:700;color:"+(d.step2DenialRate>50?"#ef4444":"#f59e0b")+"\\">"+(d.step2DenialRate||0)+"%</div><div style=\\"color:#64748b;font-size:10px\\">"+(d.step2DenialRate>50?"Escalation Pattern":"Expected Range")+"</div></div>";' +
    'html+="<div style=\\"background:linear-gradient(135deg,rgba(96,165,250,0.15),rgba(96,165,250,0.05));padding:16px;border-radius:8px;text-align:center\\"><div style=\\"color:#94a3b8;font-size:11px\\">Avg Settlement Days</div><div style=\\"font-size:28px;font-weight:700;color:#60a5fa\\">"+d.avgSettlementDays+"</div><div style=\\"color:#64748b;font-size:10px\\">days to resolution</div></div>";' +
    'html+="</div>";' +
    // Section 5: Export metrics
    'html+="<h3 style=\\"color:#e2e8f0;margin:24px 0 12px;font-size:14px;display:flex;align-items:center;gap:8px\\"><i class=\\"material-icons\\" style=\\"color:#f59e0b\\">download</i>Export Metrics</h3>";' +
    'var metrics=[' +
    '{id:"totalMembers",label:"Total Members",value:d.totalMembers,category:"Membership"},' +
    '{id:"stewardCount",label:"Steward Count",value:d.stewardCount,category:"Membership"},' +
    '{id:"totalCases",label:"Total Grievances",value:d.totalGrievances,category:"Grievances"},' +
    '{id:"openCases",label:"Open Cases",value:d.openGrievances,category:"Grievances"},' +
    '{id:"winRate",label:"Win Rate",value:d.winRate+"%",category:"Grievances"},' +
    '{id:"step1Denial",label:"Step 1 Denial Rate",value:d.step1DenialRate+"%",category:"Bargaining"},' +
    '{id:"step2Denial",label:"Step 2 Denial Rate",value:(d.step2DenialRate||0)+"%",category:"Bargaining"},' +
    '{id:"avgSettlement",label:"Avg Settlement Days",value:d.avgSettlementDays,category:"Bargaining"},' +
    '{id:"moraleScore",label:"Morale Score",value:d.moraleScore+"/10",category:"Satisfaction"},' +
    '{id:"surveyResponseRate",label:"Survey Response Rate",value:d.engagement.surveyResponseRate+"%",category:"Engagement"}' +
    '];' +
    'html+="<div style=\\"display:flex;gap:12px;margin-bottom:12px;flex-wrap:wrap\\"><button class=\\"btn btn-sm\\" onclick=\\"selectAllMetrics()\\">Select All</button><button class=\\"btn btn-sm\\" onclick=\\"clearAllMetrics()\\">Clear All</button><button class=\\"btn\\" style=\\"background:#22c55e\\" onclick=\\"exportComparison()\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle;margin-right:4px\\">download</i>Export CSV</button></div>";' +
    'html+="<div style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px\\">";' +
    'metrics.forEach(function(m){html+="<label style=\\"display:flex;align-items:center;gap:8px;background:#1e293b;padding:8px 10px;border-radius:6px;cursor:pointer;font-size:11px\\"><input type=\\"checkbox\\" class=\\"metric-check\\" data-id=\\""+m.id+"\\" data-label=\\""+m.label+"\\" data-value=\\""+m.value+"\\" data-category=\\""+m.category+"\\"><span style=\\"flex:1\\">"+m.label+"</span><span style=\\"font-weight:600;color:#60a5fa\\">"+fmt(m.value)+"</span></label>"});' +
    'html+="</div></div>";' +
    'document.getElementById("main-content").innerHTML=html;' +
    'document.querySelectorAll(".metric-check").forEach(function(cb){cb.addEventListener("change",updateComparisonPreview)})' +
    '}' +

    // Help Tab - Comprehensive FAQ & Tutorials (Steward only)
    'else if(tab==="help"&&isPII){' +
    'var faqData=[' +
    '{cat:"Getting Started",q:"How do I navigate the dashboard?",a:"Use the tabs at the top to switch between views. Overview shows key metrics, Workload shows steward assignments, Analytics shows trends and charts, and other tabs provide detailed breakdowns.",tutorial:true},' +
    '{cat:"Getting Started",q:"What is the difference between Steward and Member dashboards?",a:"The Steward Dashboard shows full PII (names, contact info) and has the Help tab. The Member Dashboard anonymizes data for public sharing but shows the same analytics.",tutorial:true},' +
    '{cat:"Getting Started",q:"How often does the data refresh?",a:"Data refreshes each time you load the dashboard. Click the Refresh button in the footer to reload the latest data.",tutorial:false},' +
    '{cat:"Grievance Management",q:"How do I track my assigned cases?",a:"Go to the My Cases tab to see all grievances assigned to you. Filter by status using the buttons at the top.",tutorial:true},' +
    '{cat:"Grievance Management",q:"What do the grievance statuses mean?",a:"Open = Active case, Pending = Awaiting response, Won = Resolved in member favor, Denied = Management decision upheld, Settled = Negotiated resolution, Withdrawn = Case dropped.",tutorial:false},' +
    '{cat:"Grievance Management",q:"How is the win rate calculated?",a:"Win Rate = (Won + Settled) / Total Resolved cases. Only closed cases are counted.",tutorial:false},' +
    '{cat:"Analytics",q:"What does the Morale Score represent?",a:"The Morale Score is the average Trust Score (Q7) from member satisfaction surveys, on a scale of 1-10.",tutorial:false},' +
    '{cat:"Analytics",q:"How do I interpret the Hot Spots?",a:"Hot Spots show locations with 3+ active cases. Red/orange colors indicate higher case counts. Focus resources on these areas.",tutorial:true},' +
    '{cat:"Analytics",q:"What are the engagement metrics?",a:"Email Open Rate = % who opened union emails. Meeting Attendance = % who attended meetings in last 6 months. Survey Response = % who completed satisfaction survey.",tutorial:false},' +
    '{cat:"Member Directory",q:"How do I find a specific member?",a:"Use the search field in the steward contact section or the filter in list views. Search works on names, locations, and units.",tutorial:false},' +
    '{cat:"Member Directory",q:"What do the participation heatmaps show?",a:"Heatmaps show engagement levels by unit or location. Green = high engagement (60%+), Yellow = moderate (40-60%), Red = low (<40%).",tutorial:true},' +
    '{cat:"Satisfaction",q:"How are satisfaction scores calculated?",a:"Each section score is the average of related survey questions (1-10 scale). Overall Score combines all sections weighted equally.",tutorial:false},' +
    '{cat:"Satisfaction",q:"What is the Key Insights section?",a:"Key Insights automatically identifies your strongest area, biggest opportunity, and suggests action items based on survey data.",tutorial:true},' +
    '{cat:"Comparison Tool",q:"How do I use the Compare tab?",a:"Select metrics you want to compare using checkboxes, then click Export to download a CSV file with your selected data.",tutorial:true},' +
    '{cat:"Comparison Tool",q:"Can I export data for reports?",a:"Yes! Use the Compare tab to select metrics, then Export. The CSV can be opened in Excel or Google Sheets for further analysis.",tutorial:false},' +
    '{cat:"Troubleshooting",q:"Data is not loading",a:"Try clicking Refresh in the footer. If that fails, close and reopen the dashboard. Check your network connection.",tutorial:false},' +
    '{cat:"Troubleshooting",q:"Charts are not displaying",a:"Ensure JavaScript is enabled in your browser. Try a different browser (Chrome recommended). Clear browser cache if issues persist.",tutorial:false}' +
    '];' +
    'var features=[' +
    '{name:"Overview Tab",desc:"At-a-glance KPIs including total members, open cases, win rate, morale score. Quick status cards and trend charts.",icon:"dashboard"},' +
    '{name:"My Cases Tab",desc:"View and filter grievances assigned to you. See case details, timelines, and status updates.",icon:"assignment_ind"},' +
    '{name:"Workload Tab",desc:"Steward workload distribution and top performers. Identify imbalances and reassignment opportunities.",icon:"work"},' +
    '{name:"Analytics Tab",desc:"Deep-dive charts: cases by category, step progression, outcomes, and member satisfaction trends.",icon:"analytics"},' +
    '{name:"Directory Tab",desc:"Member distribution by location, unit, and office days. Participation heatmaps and satisfaction breakdowns.",icon:"people"},' +
    '{name:"Hot Spots Tab",desc:"Locations with high case concentrations. Heatmap colors show severity. Prioritize these areas.",icon:"whatshot"},' +
    '{name:"Bargaining Tab",desc:"Contract-related metrics for bargaining preparation. Category breakdowns and resolution patterns.",icon:"handshake"},' +
    '{name:"Satisfaction Tab",desc:"Member survey results with section scores, key insights, and actionable recommendations.",icon:"sentiment_satisfied"},' +
    '{name:"Resources Tab",desc:"Quick links to contract, forms, and resources. Searchable steward contact directory.",icon:"folder"},' +
    '{name:"Compare Tab",desc:"Select multiple metrics to compare side-by-side. Export data as CSV for reporting.",icon:"compare"}' +
    '];' +
    'html="<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">search</i>Search Help</div>";' +
    'html+="<input type=\\"text\\" id=\\"faqSearch\\" placeholder=\\"Search FAQs, tutorials, and features...\\" oninput=\\"filterFAQ()\\" style=\\"width:100%;padding:12px;border:1px solid #475569;border-radius:8px;background:#1e293b;color:#f8fafc;font-size:14px;margin-bottom:8px\\">";' +
    'html+="<div style=\\"display:flex;gap:8px;flex-wrap:wrap\\"><button class=\\"btn btn-sm faq-filter active\\" onclick=\\"setFAQFilter(\\x27all\\x27,this)\\">All</button><button class=\\"btn btn-sm faq-filter\\" onclick=\\"setFAQFilter(\\x27tutorial\\x27,this)\\">Tutorials Only</button><button class=\\"btn btn-sm faq-filter\\" onclick=\\"setFAQFilter(\\x27features\\x27,this)\\">Features</button></div></div>";' +
    // Feature Breakdown
    'html+="<div id=\\"features-section\\" class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">apps</i>Feature Breakdown</div><div class=\\"feature-grid\\" style=\\"display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px\\">";' +
    'features.forEach(function(f){html+="<div class=\\"feature-item\\" style=\\"background:#1e293b;padding:12px;border-radius:8px;border:1px solid #334155\\"><div style=\\"display:flex;align-items:center;gap:8px;margin-bottom:6px\\"><i class=\\"material-icons\\" style=\\"color:#60a5fa\\">"+f.icon+"</i><span style=\\"font-weight:600;color:#e2e8f0\\">"+f.name+"</span></div><p style=\\"color:#94a3b8;font-size:12px;margin:0;line-height:1.4\\">"+f.desc+"</p></div>"});' +
    'html+="</div></div>";' +
    // FAQ Section
    'html+="<div id=\\"faq-section\\" class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">help_outline</i>Frequently Asked Questions</div><div id=\\"faq-list\\">";' +
    'var cats=["Getting Started","Grievance Management","Analytics","Member Directory","Satisfaction","Comparison Tool","Troubleshooting"];' +
    'cats.forEach(function(cat){' +
    'var catFaqs=faqData.filter(function(f){return f.cat===cat});' +
    'if(catFaqs.length>0){html+="<div class=\\"faq-category\\" style=\\"margin-bottom:16px\\"><div style=\\"font-weight:600;color:#a78bfa;margin-bottom:8px;font-size:13px\\">"+cat+"</div>";' +
    'catFaqs.forEach(function(f){html+="<div class=\\"faq-item\\" data-searchable=\\""+f.q.toLowerCase()+" "+f.a.toLowerCase()+"\\" data-tutorial=\\""+(f.tutorial?"yes":"no")+"\\" style=\\"background:#1e293b;padding:12px;border-radius:8px;margin-bottom:8px;border-left:3px solid "+(f.tutorial?"#22c55e":"#475569")+"\\"><div style=\\"font-weight:500;color:#e2e8f0;margin-bottom:6px;display:flex;align-items:center;gap:6px\\">"+(f.tutorial?"<span style=\\"background:#22c55e;color:white;font-size:9px;padding:2px 6px;border-radius:4px\\">TUTORIAL</span>":"")+""+f.q+"</div><p style=\\"color:#94a3b8;font-size:12px;margin:0;line-height:1.5\\">"+f.a+"</p></div>"});' +
    'html+="</div>"}});' +
    'html+="</div></div>";' +
    'document.getElementById("main-content").innerHTML=html' +
    '}' +
    '}' +

    // Chart rendering functions
    'function renderOverviewCharts(){' +
    'var d=dashData;' +
    // Status chart with drill-down
    'var statusKeys=["open","pending","won","denied","settled","withdrawn"];' +
    'var statusChart=new Chart(document.getElementById("statusChart"),{type:"doughnut",data:{labels:["Open","Pending","Won","Denied","Settled","Withdrawn"],datasets:[{data:[d.statusDistribution.open,d.statusDistribution.pending,d.statusDistribution.won,d.statusDistribution.denied,d.statusDistribution.settled,d.statusDistribution.withdrawn],backgroundColor:["#3b82f6","#f59e0b","#22c55e","#ef4444","#8b5cf6","#64748b"]}]},options:{responsive:true,onClick:function(e,el){if(el.length>0){var idx=el[0].index;showChartDrillDown("status",statusKeys[idx],statusChart.data.labels[idx]+" Cases")}},plugins:{legend:{position:"right",labels:{color:"#cbd5e1",font:{size:11}}}}}});' +
    'var trendLabels=d.sentimentTrend.length>0?d.sentimentTrend.map(function(t){return t.month}):["Jan","Feb","Mar"];' +
    'var trendData=d.sentimentTrend.length>0?d.sentimentTrend.map(function(t){return t.score}):[7,7.2,7.5];' +
    'new Chart(document.getElementById("trendChart"),{type:"line",data:{labels:trendLabels,datasets:[{label:"Trust Score",data:trendData,borderColor:"#a78bfa",backgroundColor:"rgba(167,139,250,0.2)",fill:true,tension:0.4}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{min:0,max:10,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    'var locLabels=Object.keys(d.locationBreakdown).slice(0,8);' +
    'var locData=locLabels.map(function(l){return d.locationBreakdown[l]});' +
    // Location chart with drill-down
    'var locChart=new Chart(document.getElementById("locationChart"),{type:"bar",data:{labels:locLabels,datasets:[{label:"Members",data:locData,backgroundColor:"#3b82f6"}]},options:{responsive:true,indexAxis:"y",onClick:function(e,el){if(el.length>0){var idx=el[0].index;showChartDrillDown("unit",locLabels[idx],"Members in "+locLabels[idx])}},plugins:{legend:{display:false}},scales:{x:{ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1"}}}}});' +
    'var filingLabels=d.monthlyFilings.map(function(f){return f.month});' +
    'var filingData=d.monthlyFilings.map(function(f){return f.count});' +
    'var resolvedData=d.monthlyResolved?d.monthlyResolved.map(function(f){return f.count}):[];' +
    'new Chart(document.getElementById("filingChart"),{type:"line",data:{labels:filingLabels.length>0?filingLabels:["Jan","Feb","Mar"],datasets:[{label:"Filed",data:filingData.length>0?filingData:[2,3,1],borderColor:"#f59e0b",backgroundColor:"rgba(245,158,11,0.1)",fill:true,tension:0.4},{label:"Resolved",data:resolvedData.length>0?resolvedData:[1,2,2],borderColor:"#22c55e",backgroundColor:"rgba(34,197,94,0.1)",fill:true,tension:0.4}]},options:{responsive:true,plugins:{legend:{display:true,labels:{color:"#cbd5e1"}}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}})' +
    '}' +

    'function renderAnalyticsCharts(){' +
    'var d=dashData;' +
    'var unitLabels=Object.keys(d.unitBreakdown).slice(0,8);' +
    'var unitData=unitLabels.map(function(u){return d.unitBreakdown[u]});' +
    // Unit chart with drill-down
    'var unitChart=new Chart(document.getElementById("unitChart"),{type:"doughnut",data:{labels:unitLabels,datasets:[{data:unitData,backgroundColor:["#3b82f6","#22c55e","#f59e0b","#ef4444","#8b5cf6","#ec4899","#06b6d4","#84cc16"]}]},options:{responsive:true,onClick:function(e,el){if(el.length>0){var idx=el[0].index;showChartDrillDown("unit",unitLabels[idx],"Members in "+unitLabels[idx])}},plugins:{legend:{position:"right",labels:{color:"#cbd5e1",font:{size:10}}}}}});' +
    // Outcomes chart with drill-down
    'var outcomeKeys=["won","denied","settled"];' +
    'new Chart(document.getElementById("outcomeChart"),{type:"bar",data:{labels:["Won","Denied","Settled"],datasets:[{label:"Cases",data:[d.wins,d.losses,d.settled],backgroundColor:["#22c55e","#ef4444","#8b5cf6"]}]},options:{responsive:true,onClick:function(e,el){if(el.length>0){var idx=el[0].index;showChartDrillDown("status",outcomeKeys[idx],["Won","Denied","Settled"][idx]+" Cases")}},plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    'var catLabels=Object.keys(d.grievancesByCategory).slice(0,6);' +
    'var catData=catLabels.map(function(c){return d.grievancesByCategory[c]});' +
    // Category chart with drill-down
    'new Chart(document.getElementById("categoryChart"),{type:"bar",data:{labels:catLabels.length>0?catLabels:["Discipline","Contract","Safety"],datasets:[{label:"Cases",data:catData.length>0?catData:[3,5,2],backgroundColor:"#06b6d4"}]},options:{responsive:true,onClick:function(e,el){if(el.length>0&&catLabels.length>0){var idx=el[0].index;showChartDrillDown("category",catLabels[idx],catLabels[idx]+" Cases")}},plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8",maxRotation:45}}}}});' +
    'new Chart(document.getElementById("stepChart"),{type:"bar",data:{labels:["Step 1","Step 2","Step 3","Arbitration"],datasets:[{label:"Cases",data:[d.stepProgression.step1,d.stepProgression.step2,d.stepProgression.step3,d.stepProgression.arb],backgroundColor:["#3b82f6","#f59e0b","#ef4444","#8b5cf6"]}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}})' +
    '}' +

    'function renderDirectoryCharts(){' +
    'var d=dashData;' +
    // Employees by Location chart
    'var locLabels=Object.keys(d.locationBreakdown).slice(0,10);' +
    'var locData=locLabels.map(function(l){return d.locationBreakdown[l]});' +
    'if(document.getElementById("empByLocationChart")){new Chart(document.getElementById("empByLocationChart"),{type:"bar",data:{labels:locLabels,datasets:[{label:"Employees",data:locData,backgroundColor:"#3b82f6"}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{color:"#94a3b8",callback:function(v){return v>=1000?v.toLocaleString():v}}},y:{ticks:{color:"#cbd5e1"}}}}})}' +
    // Employees by Unit chart
    'var unitLabels=Object.keys(d.unitBreakdown).slice(0,10);' +
    'var unitData=unitLabels.map(function(u){return d.unitBreakdown[u]});' +
    'if(document.getElementById("empByUnitChart")){new Chart(document.getElementById("empByUnitChart"),{type:"bar",data:{labels:unitLabels,datasets:[{label:"Employees",data:unitData,backgroundColor:"#8b5cf6"}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{color:"#94a3b8",callback:function(v){return v>=1000?v.toLocaleString():v}}},y:{ticks:{color:"#cbd5e1"}}}}})}' +
    // Employees by Office Days chart
    'var dayOrder=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];' +
    'var dayLabels=Object.keys(d.officeDaysBreakdown).sort(function(a,b){return dayOrder.indexOf(a)-dayOrder.indexOf(b)});' +
    'var dayData=dayLabels.map(function(day){return d.officeDaysBreakdown[day]});' +
    'if(document.getElementById("empByOfficeDaysChart")&&dayLabels.length>0){new Chart(document.getElementById("empByOfficeDaysChart"),{type:"bar",data:{labels:dayLabels,datasets:[{label:"Employees",data:dayData,backgroundColor:["#ef4444","#f59e0b","#22c55e","#3b82f6","#8b5cf6","#ec4899","#06b6d4"]}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8",callback:function(v){return v>=1000?v.toLocaleString():v}}},x:{ticks:{color:"#cbd5e1"}}}}})}' +
    // Satisfaction by Unit/Role chart
    'var satUnitLabels=Object.keys(d.satisfactionByUnit).slice(0,8);' +
    'var satUnitData=satUnitLabels.map(function(u){return d.satisfactionByUnit[u].score});' +
    'var satUnitColors=satUnitData.map(function(s){return s>=7?"#22c55e":s>=5?"#f59e0b":"#ef4444"});' +
    'if(document.getElementById("satByUnitChart")&&satUnitLabels.length>0){new Chart(document.getElementById("satByUnitChart"),{type:"bar",data:{labels:satUnitLabels,datasets:[{label:"Score",data:satUnitData,backgroundColor:satUnitColors}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{min:0,max:10,ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1",font:{size:10}}}}}})}' +
    // Satisfaction by Location chart
    'var satLocLabels=Object.keys(d.satisfactionByLocation).slice(0,8);' +
    'var satLocData=satLocLabels.map(function(l){return d.satisfactionByLocation[l].score});' +
    'var satLocColors=satLocData.map(function(s){return s>=7?"#22c55e":s>=5?"#f59e0b":"#ef4444"});' +
    'if(document.getElementById("satByLocationChart")&&satLocLabels.length>0){new Chart(document.getElementById("satByLocationChart"),{type:"bar",data:{labels:satLocLabels,datasets:[{label:"Score",data:satLocData,backgroundColor:satLocColors}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{min:0,max:10,ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1",font:{size:10}}}}}})}' +
    '}' +

    'function renderHotspotChart(){' +
    'var d=dashData;' +
    'var locLabels=Object.keys(d.locationBreakdown).slice(0,10);' +
    'var locData=locLabels.map(function(l){return d.locationBreakdown[l]});' +
    'var maxVal=Math.max.apply(null,locData)||1;var minVal=Math.min.apply(null,locData)||0;' +
    'function heatColor(v){var pct=(v-minVal)/(maxVal-minVal||1);if(pct<=0.5){return"rgba("+(209+Math.round((254-209)*pct*2))+","+(250+Math.round((243-250)*pct*2))+","+(229+Math.round((199-229)*pct*2))+",0.9)"}else{var p=(pct-0.5)*2;return"rgba("+(254+Math.round((252-254)*p))+","+(243+Math.round((165-243)*p))+","+(199+Math.round((165-199)*p))+",0.9)"}}' +
    'var barColors=locData.map(function(v){return heatColor(v)});' +
    'new Chart(document.getElementById("hotspotChart"),{type:"bar",data:{labels:locLabels,datasets:[{label:"Cases",data:locData,backgroundColor:barColors,borderColor:barColors.map(function(c){return c.replace("0.9","1")}),borderWidth:1}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1"}}}}})' +
    '}' +

    'function renderBargainCharts(){' +
    'var d=dashData;' +
    'var artLabels=Object.keys(d.articleViolations).slice(0,8);' +
    'var artData=artLabels.map(function(a){return d.articleViolations[a]});' +
    'new Chart(document.getElementById("bargainChart"),{type:"bar",data:{labels:artLabels.length>0?artLabels:["Art 5","Art 7","Art 12"],datasets:[{label:"Violations",data:artData.length>0?artData:[3,5,2],backgroundColor:"#fbbf24"}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    'new Chart(document.getElementById("stepDistChart"),{type:"doughnut",data:{labels:["Step 1","Step 2","Step 3","Arbitration"],datasets:[{data:[d.stepProgression.step1,d.stepProgression.step2,d.stepProgression.step3,d.stepProgression.arb],backgroundColor:["#3b82f6","#f59e0b","#ef4444","#8b5cf6"]}]},options:{responsive:true,plugins:{legend:{position:"right",labels:{color:"#cbd5e1"}}}}})' +
    '}' +

    'function renderSatisfactionChart(){' +
    'var d=dashData;' +
    'var labels=d.satisfactionData.sections.map(function(s){return s.name});' +
    'var scores=d.satisfactionData.sections.map(function(s){return s.score});' +
    'var colors=scores.map(function(s){return s>=7?"#22c55e":s>=5?"#f59e0b":"#ef4444"});' +
    'new Chart(document.getElementById("satChart"),{type:"bar",data:{labels:labels,datasets:[{label:"Score",data:scores,backgroundColor:colors}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{min:0,max:10,ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1",font:{size:10}}}}}})' +
    '}' +

    // FAQ/Help Modal function (v4.4.0)
    'function showFAQ(){' +
    'var faq=[' +
    '{cat:"Getting Started",items:[' +
    '{q:"What is this dashboard?",a:"This is the Union Dashboard providing real-time analytics on grievances, members, and steward workload."},' +
    '{q:"How is data updated?",a:"Data refreshes each time you open the dashboard. Click Refresh to reload."},' +
    '{q:"What\\x27s the difference between Steward and Member dashboards?",a:"Steward dashboard shows full PII (names, IDs). Member dashboard anonymizes data for public sharing."}' +
    ']},' +
    '{cat:"Dashboard Tabs",items:[' +
    '{q:"What is the My Cases tab?",a:"Shows your personally assigned grievances as a steward. Only visible in Steward mode."},' +
    '{q:"What are Hot Spots?",a:"Locations with 3 or more active grievances that need management attention."},' +
    '{q:"What are Engagement Metrics?",a:"Email open rates, meeting attendance, volunteer hours, and union interest levels from Member Directory."},' +
    '{q:"What is the Morale Score?",a:"Average trust/satisfaction score from the Member Satisfaction survey (1-10 scale)."}' +
    ']},' +
    '{cat:"Understanding Metrics",items:[' +
    '{q:"How is Win Rate calculated?",a:"(Won Cases / Total Closed Cases) × 100. Settled cases count as closed but not won."},' +
    '{q:"What does OVERLOAD mean?",a:"A steward with more than 8 active cases. Heavy = 5-8 cases."},' +
    '{q:"What is Step 1 Denial Rate?",a:"Percentage of grievances denied at Step 1 that had to escalate to Step 2."},' +
    '{q:"What is Avg Settlement Time?",a:"Average number of days from filing to closure for resolved grievances."}' +
    ']},' +
    '{cat:"Taking Action",items:[' +
    '{q:"How do I file a new grievance?",a:"Go to Member Directory in the spreadsheet > check the Start Grievance box for a member."},' +
    '{q:"How do I get help?",a:"Use the Help menu in the spreadsheet or contact your chapter leadership."},' +
    '{q:"Can I export this data?",a:"Use your browser\\x27s print function or screenshot the dashboard."}' +
    ']}' +
    '];' +
    'var html="<div style=\\"max-height:400px;overflow-y:auto;padding:8px\\">";' +
    'faq.forEach(function(section){' +
    'html+="<div style=\\"font-size:12px;font-weight:600;color:#22c55e;margin:16px 0 8px;padding:6px 10px;background:rgba(34,197,94,0.1);border-radius:6px\\">"+section.cat+"</div>";' +
    'section.items.forEach(function(item){' +
    'html+="<div style=\\"border-left:3px solid #3b82f6;padding-left:12px;margin-bottom:12px\\">";' +
    'html+="<div style=\\"font-weight:500;color:#e2e8f0;margin-bottom:4px;font-size:13px\\">"+item.q+"</div>";' +
    'html+="<div style=\\"color:#94a3b8;font-size:12px;line-height:1.5\\">"+item.a+"</div></div>"})});' +
    'html+="</div>";' +
    'openModal("Help & FAQ",html)' +
    '}' +

    // Filter My Cases by status (v4.4.0)
    'function filterMyCases(status){' +
    'document.querySelectorAll(".my-case-item").forEach(function(el){' +
    'if(status==="all"){el.style.display="flex"}' +
    'else if(status==="Overdue"){el.style.display=el.getAttribute("data-overdue")==="yes"?"flex":"none"}' +
    'else{el.style.display=el.getAttribute("data-status").toLowerCase().indexOf(status.toLowerCase())>=0?"flex":"none"}' +
    '})' +
    '}' +

    // Filter Meeting Notes by search
    'function filterMeetingNotes(){' +
    'var q=(document.getElementById("notesSearch")||{value:""}).value.toLowerCase();' +
    'document.querySelectorAll(".meeting-note-item").forEach(function(el){' +
    'el.style.display=(el.getAttribute("data-search")||"").indexOf(q)>=0?"flex":"none"' +
    '})' +
    '}' +

    // Filter Stewards by search (v4.4.0)
    'function filterStewards(){' +
    'var q=document.getElementById("stewardSearch").value.toLowerCase();' +
    'document.querySelectorAll(".steward-contact").forEach(function(el){' +
    'el.style.display=el.getAttribute("data-search").indexOf(q)>=0?"flex":"none"' +
    '})' +
    '}' +
    // Smart steward search with autocomplete
    'function smartFilterStewards(){' +
    'var q=document.getElementById("stewardSearch").value.toLowerCase();' +
    'var suggestions=document.getElementById("searchSuggestions");' +
    'filterStewards();' +
    'if(q.length<2){suggestions.style.display="none";return}' +
    'var matches=[];' +
    'document.querySelectorAll(".steward-contact").forEach(function(el){' +
    'var searchable=el.getAttribute("data-search")||"";' +
    'if(searchable.indexOf(q)>=0){' +
    'var name=el.querySelector("span[style*=\\"font-weight\\"]");' +
    'if(name)matches.push(name.textContent)}' +
    '});' +
    'if(matches.length>0&&matches.length<10){' +
    'suggestions.innerHTML=matches.slice(0,5).map(function(m){return "<div style=\\"padding:8px 12px;cursor:pointer;border-bottom:1px solid #334155\\" onmouseover=\\"this.style.background=\\x27#334155\\x27\\" onmouseout=\\"this.style.background=\\x27transparent\\x27\\" onclick=\\"selectStewardSuggestion(\\x27"+m.replace(/\\x27/g,"")+"\\x27)\\"><i class=\\"material-icons\\" style=\\"font-size:14px;vertical-align:middle;margin-right:6px;color:#60a5fa\\">person</i>"+escapeHtml(m)+"</div>"}).join("");' +
    'suggestions.style.display="block"}else{suggestions.style.display="none"}' +
    '}' +
    'function selectStewardSuggestion(name){' +
    'document.getElementById("stewardSearch").value=name;' +
    'document.getElementById("searchSuggestions").style.display="none";' +
    'filterStewards()' +
    '}' +
    'function showSearchSuggestions(){' +
    'var q=document.getElementById("stewardSearch").value;' +
    'if(q.length>=2)smartFilterStewards()' +
    '}' +
    'function quickFilterSteward(loc){' +
    'document.getElementById("stewardSearch").value=loc;' +
    'filterStewards();' +
    'document.getElementById("searchSuggestions").style.display="none"' +
    '}' +
    'function clearStewardFilter(){' +
    'document.getElementById("stewardSearch").value="";' +
    'filterStewards();' +
    'var s=document.getElementById("searchSuggestions");if(s)s.style.display="none"' +
    '}' +
    'document.addEventListener("click",function(e){var s=document.getElementById("searchSuggestions");if(s&&!e.target.closest("#stewardSearch")&&!e.target.closest("#searchSuggestions"))s.style.display="none"});' +

    // Save steward contact (v4.4.0)
    'function saveStewardContact(name,email,phone){' +
    'var vcard="BEGIN:VCARD\\nVERSION:3.0\\nFN:"+name+"\\nORG:SEIU Local 509\\nTITLE:Union Steward\\n";' +
    'if(email)vcard+="EMAIL:"+email+"\\n";' +
    'if(phone)vcard+="TEL:"+phone+"\\n";' +
    'vcard+="END:VCARD";' +
    'var blob=new Blob([vcard],{type:"text/vcard"});' +
    'var url=URL.createObjectURL(blob);' +
    'var a=document.createElement("a");' +
    'a.href=url;a.download=name.replace(/\\s+/g,"_")+".vcf";' +
    'document.body.appendChild(a);a.click();document.body.removeChild(a);' +
    'URL.revokeObjectURL(url)' +
    '}' +

    // Compare Tab Functions
    'function selectAllMetrics(){document.querySelectorAll(".metric-check").forEach(function(cb){cb.checked=true});updateComparisonPreview()}' +
    'function clearAllMetrics(){document.querySelectorAll(".metric-check").forEach(function(cb){cb.checked=false});updateComparisonPreview()}' +
    'function updateComparisonPreview(){' +
    'var selected=[];document.querySelectorAll(".metric-check:checked").forEach(function(cb){selected.push({id:cb.dataset.id,label:cb.dataset.label,value:cb.dataset.value,category:cb.dataset.category})});' +
    'var preview=document.getElementById("comparison-preview");var content=document.getElementById("preview-content");' +
    'if(selected.length===0){preview.style.display="none";return}' +
    'preview.style.display="block";' +
    'var html="<p style=\\"color:#94a3b8;margin-bottom:12px\\">"+selected.length+" metrics selected</p>";' +
    'html+="<table style=\\"width:100%;border-collapse:collapse;font-size:12px\\"><tr style=\\"background:#1e293b\\"><th style=\\"padding:8px;text-align:left;color:#94a3b8\\">Category</th><th style=\\"padding:8px;text-align:left;color:#94a3b8\\">Metric</th><th style=\\"padding:8px;text-align:right;color:#94a3b8\\">Value</th></tr>";' +
    'selected.forEach(function(m,i){html+="<tr style=\\"background:"+(i%2===0?"#0f172a":"#1e293b")+"\\"><td style=\\"padding:8px;color:#60a5fa\\">"+m.category+"</td><td style=\\"padding:8px;color:#e2e8f0\\">"+m.label+"</td><td style=\\"padding:8px;text-align:right;font-weight:600;color:#22c55e\\">"+m.value+"</td></tr>"});' +
    'html+="</table>";content.innerHTML=html' +
    '}' +
    'function exportComparison(){' +
    'var selected=[];document.querySelectorAll(".metric-check:checked").forEach(function(cb){selected.push({label:cb.dataset.label,value:cb.dataset.value,category:cb.dataset.category})});' +
    'if(selected.length===0){alert("Please select at least one metric to export");return}' +
    'var csv="Category,Metric,Value\\n";' +
    'selected.forEach(function(m){csv+="\\""+m.category+"\\",\\""+m.label+"\\",\\""+m.value+"\\"\\n"});' +
    'var blob=new Blob([csv],{type:"text/csv"});' +
    'var url=URL.createObjectURL(blob);' +
    'var a=document.createElement("a");' +
    'a.href=url;a.download="509_metrics_comparison_"+new Date().toISOString().split("T")[0]+".csv";' +
    'document.body.appendChild(a);a.click();document.body.removeChild(a);' +
    'URL.revokeObjectURL(url)' +
    '}' +

    // Satisfaction Tab Functions - Toggle section details
    'function toggleSectionDetail(sIdx){' +
    'var detail=document.getElementById("section-detail-"+sIdx);' +
    'if(detail){detail.style.display=detail.style.display==="none"?"block":"none"}' +
    '}' +

    // Help Tab Functions (FAQ search and filter)
    'var currentFAQFilter="all";' +
    'function filterFAQ(){' +
    'var query=document.getElementById("faqSearch").value.toLowerCase();' +
    'document.querySelectorAll(".faq-item").forEach(function(item){' +
    'var searchable=item.dataset.searchable;' +
    'var isTutorial=item.dataset.tutorial==="yes";' +
    'var matchesQuery=!query||searchable.indexOf(query)>=0;' +
    'var matchesFilter=currentFAQFilter==="all"||(currentFAQFilter==="tutorial"&&isTutorial);' +
    'item.style.display=(matchesQuery&&matchesFilter)?"block":"none"' +
    '});' +
    'document.querySelectorAll(".feature-item").forEach(function(item){' +
    'var text=(item.textContent||"").toLowerCase();' +
    'var matchesQuery=!query||text.indexOf(query)>=0;' +
    'var matchesFilter=currentFAQFilter==="all"||currentFAQFilter==="features";' +
    'item.style.display=(matchesQuery&&matchesFilter)?"block":"none"' +
    '});' +
    'document.getElementById("features-section").style.display=(currentFAQFilter==="tutorial")?"none":"block";' +
    'document.getElementById("faq-section").style.display=(currentFAQFilter==="features")?"none":"block"' +
    '}' +
    'function setFAQFilter(filter,btn){' +
    'currentFAQFilter=filter;' +
    'document.querySelectorAll(".faq-filter").forEach(function(b){b.classList.remove("active")});' +
    'btn.classList.add("active");' +
    'filterFAQ()' +
    '}' +

    // Settings Panel Functions
    'function toggleSettings(){' +
    'document.getElementById("settingsPanel").classList.toggle("open");' +
    'document.querySelector(".settings-overlay").classList.toggle("open")' +
    '}' +
    // View Mode Toggle
    'function setViewMode(mode){' +
    'document.body.classList.remove("mobile-mode","desktop-mode");' +
    'document.querySelectorAll(".view-toggle button").forEach(function(b){b.classList.remove("active")});' +
    'if(mode==="mobile"){document.body.classList.add("mobile-mode");document.getElementById("mobileViewBtn").classList.add("active")}' +
    'else if(mode==="desktop"){document.body.classList.add("desktop-mode");document.getElementById("desktopViewBtn").classList.add("active")}' +
    'else{document.getElementById("autoViewBtn").classList.add("active")}' +
    'localStorage.setItem("509_viewMode",mode);' +
    'showHint(mode.charAt(0).toUpperCase()+mode.slice(1)+" view")' +
    '}' +

    'function toggleHighContrast(){' +
    'document.body.classList.toggle("high-contrast");' +
    'document.getElementById("highContrastToggle").classList.toggle("on");' +
    'localStorage.setItem("509_highContrast",document.body.classList.contains("high-contrast"))' +
    '}' +
    'function toggleLargeText(){' +
    'document.body.classList.toggle("large-text");' +
    'document.getElementById("largeTextToggle").classList.toggle("on");' +
    'localStorage.setItem("509_largeText",document.body.classList.contains("large-text"))' +
    '}' +
    'var autoRefreshTimer=null;' +
    'function toggleAutoRefresh(){' +
    'var toggle=document.getElementById("autoRefreshToggle");' +
    'toggle.classList.toggle("on");' +
    'if(toggle.classList.contains("on")){startAutoRefresh()}else{stopAutoRefresh()}' +
    'localStorage.setItem("509_autoRefresh",toggle.classList.contains("on"))' +
    '}' +
    'function startAutoRefresh(){' +
    'var interval=parseInt(document.getElementById("refreshInterval").value);' +
    'autoRefreshTimer=setInterval(function(){location.reload()},interval)' +
    '}' +
    'function stopAutoRefresh(){if(autoRefreshTimer){clearInterval(autoRefreshTimer);autoRefreshTimer=null}}' +
    'function setRefreshInterval(){stopAutoRefresh();if(document.getElementById("autoRefreshToggle").classList.contains("on")){startAutoRefresh()}}' +
    'function saveGoals(){' +
    'var goals={winRate:parseInt(document.getElementById("goalWinRate").value),morale:parseFloat(document.getElementById("goalMorale").value),response:parseInt(document.getElementById("goalResponse").value)};' +
    'localStorage.setItem("509_goals",JSON.stringify(goals));' +
    'if(typeof renderOverviewWithGoals==="function")renderOverviewWithGoals()' +
    '}' +
    'function loadGoals(){' +
    'try{var g=JSON.parse(localStorage.getItem("509_goals")||"{}");' +
    'if(g.winRate)document.getElementById("goalWinRate").value=g.winRate;' +
    'if(g.morale)document.getElementById("goalMorale").value=g.morale;' +
    'if(g.response)document.getElementById("goalResponse").value=g.response}catch(e){}' +
    '}' +

    // Alert Center Functions
    'function toggleAlerts(){document.getElementById("alertCenter").classList.toggle("open")}' +
    'function populateAlerts(){' +
    'var d=dashData;if(!d)return;var alerts=[];var now=new Date();' +
    'if(d.overdueCount>0){alerts.push({type:"critical",title:"Overdue Grievances",desc:d.overdueCount+" case(s) past deadline",time:"Action required"})}' +
    'if(d.deadlines&&d.deadlines.length>0){d.deadlines.forEach(function(dl){alerts.push({type:"warning",title:"Upcoming Deadline",desc:dl.id+": "+dl.step,time:dl.daysUntil+" days"})})}' +
    'if(d.winRate<50){alerts.push({type:"warning",title:"Win Rate Below Target",desc:"Current: "+d.winRate+"% (Target: 75%)",time:"Review cases"})}' +
    'if(d.moraleScore<6){alerts.push({type:"warning",title:"Low Morale Score",desc:"Current: "+d.moraleScore+"/10",time:"Survey feedback"})}' +
    'if(d.engagement&&d.engagement.emailOpenRate<30){alerts.push({type:"info",title:"Low Email Engagement",desc:d.engagement.emailOpenRate+"% open rate",time:"Consider outreach"})}' +
    'if(d.hotZones&&d.hotZones.length>0){alerts.push({type:"info",title:"Active Hot Spots",desc:d.hotZones.length+" location(s) with 3+ cases",time:"Monitor closely"})}' +
    'if(alerts.length===0){alerts.push({type:"success",title:"All Clear",desc:"No alerts at this time",time:"Looking good!"})}' +
    'var html="";alerts.forEach(function(a){html+="<div class=\\"alert-item "+a.type+"\\"><div class=\\"alert-title\\">"+a.title+"</div><div class=\\"alert-desc\\">"+a.desc+"</div><div class=\\"alert-time\\">"+a.time+"</div></div>"});' +
    'document.getElementById("alertList").innerHTML=html;' +
    'var badge=document.getElementById("alertBadge");var criticalCount=alerts.filter(function(a){return a.type==="critical"||a.type==="warning"}).length;' +
    'if(criticalCount>0){badge.textContent=criticalCount;badge.style.display="block"}else{badge.style.display="none"}' +
    '}' +

    // Date Range Functions
    'var currentDateRange=7;' +
    'function setDateRange(days,btn){' +
    'currentDateRange=days;' +
    'document.querySelectorAll(".date-btn").forEach(function(b){b.classList.remove("active")});' +
    'btn.classList.add("active");' +
    'document.getElementById("dateFrom").value="";document.getElementById("dateTo").value="";' +
    'applyDateFilter()' +
    '}' +
    'function applyCustomDateRange(){' +
    'var from=document.getElementById("dateFrom").value;var to=document.getElementById("dateTo").value;' +
    'if(from&&to){document.querySelectorAll(".date-btn").forEach(function(b){b.classList.remove("active")});applyDateFilter(from,to)}' +
    '}' +
    'function applyDateFilter(from,to){' +
    'showHint("Loading filtered data...");' +
    'google.script.run.withSuccessHandler(function(json){' +
    'dashData=JSON.parse(json);renderTab(document.querySelector(".tab.active").textContent.toLowerCase().replace(/\\s+/g,""));showHint("Date filter applied")' +
    '}).withFailureHandler(function(e){showHint("Filter error: "+e.message)}).getUnifiedDashboardDataWithDateRange(isPII,currentDateRange,from,to)' +
    '}' +

    // Pinned Metrics Functions
    'var pinnedMetrics=JSON.parse(localStorage.getItem("509_pinned")||"[]");' +
    'function togglePin(metricKey,label,value,color){' +
    'var idx=pinnedMetrics.findIndex(function(p){return p.key===metricKey});' +
    'if(idx>=0){pinnedMetrics.splice(idx,1);showHint("Unpinned: "+label)}' +
    'else{pinnedMetrics.push({key:metricKey,label:label,value:value,color:color});showHint("Pinned: "+label)}' +
    'localStorage.setItem("509_pinned",JSON.stringify(pinnedMetrics));' +
    'renderPinnedSection()' +
    '}' +
    'function renderPinnedSection(){' +
    'var existing=document.getElementById("pinnedSection");if(existing)existing.remove();' +
    'if(pinnedMetrics.length===0)return;' +
    'var html="<div id=\\"pinnedSection\\" class=\\"pinned-section\\"><div class=\\"pinned-header\\"><span class=\\"pinned-title\\"><i class=\\"material-icons\\" style=\\"font-size:16px\\">push_pin</i>Pinned Metrics</span><button onclick=\\"clearAllPinned()\\" class=\\"btn btn-secondary btn-sm\\">Clear All</button></div><div class=\\"pinned-grid\\">";' +
    'pinnedMetrics.forEach(function(p){' +
    'var val=p.value;if(dashData){' +
    'if(p.key==="members")val=dashData.totalMembers;' +
    'else if(p.key==="stewards")val=dashData.stewardCount;' +
    'else if(p.key==="openCases")val=dashData.openGrievances;' +
    'else if(p.key==="winRate")val=dashData.winRate+"%";' +
    'else if(p.key==="overdue")val=dashData.overdueCount;' +
    'else if(p.key==="morale")val=dashData.moraleScore;' +
    '}' +
    'html+="<div class=\\"pinned-metric\\"><button class=\\"unpin-btn\\" onclick=\\"togglePin(\\x27"+p.key+"\\x27,\\x27"+p.label+"\\x27)\\"><i class=\\"material-icons\\">close</i></button><div style=\\"font-size:10px;color:#94a3b8\\">"+p.label+"</div><div style=\\"font-size:24px;font-weight:700;color:"+p.color+"\\">"+val+"</div></div>"});' +
    'html+="</div></div>";' +
    'var content=document.getElementById("main-content");' +
    'if(content)content.insertAdjacentHTML("afterbegin",html)' +
    '}' +
    'function clearAllPinned(){pinnedMetrics=[];localStorage.removeItem("509_pinned");renderPinnedSection();showHint("All pins cleared")}' +

    // Chart Drill-Down Functions
    'function showChartDrillDown(type,key,title){' +
    'var d=dashData;if(!d||!d.chartDrillDown)return;' +
    'var items=[];' +
    'if(type==="status"&&d.chartDrillDown.statusByCase[key])items=d.chartDrillDown.statusByCase[key];' +
    'else if(type==="location"&&d.chartDrillDown.locationByCase[key])items=d.chartDrillDown.locationByCase[key];' +
    'else if(type==="category"&&d.chartDrillDown.categoryByCase[key])items=d.chartDrillDown.categoryByCase[key];' +
    'else if(type==="unit"&&d.chartDrillDown.unitByMember[key])items=d.chartDrillDown.unitByMember[key];' +
    'else if(type==="steward"&&d.chartDrillDown.stewardByCase[key])items=d.chartDrillDown.stewardByCase[key];' +
    'if(items.length===0){openModal(title,"<p style=\\"color:#94a3b8\\">No data available</p>");return}' +
    'var html="<div class=\\"drill-down-list\\">";' +
    'items.forEach(function(item){' +
    'if(item.id&&item.member){html+="<div class=\\"drill-down-item\\"><span><strong>"+item.id+"</strong> - "+item.member+"</span><span style=\\"color:#94a3b8\\">"+item.steward+" | "+item.location+"</span></div>"}' +
    'else if(item.id&&item.name){html+="<div class=\\"drill-down-item\\"><span>"+item.name+"</span><span style=\\"color:#94a3b8\\">"+item.location+(item.isSteward?" (Steward)":"")+"</span></div>"}' +
    '});' +
    'html+="</div>";' +
    'openModal(title+" ("+items.length+")",html)' +
    '}' +

    // Print & Export Functions
    'function printDashboard(){toggleSettings();setTimeout(function(){window.print()},300)}' +
    'function exportAllData(){' +
    'var d=dashData;if(!d){alert("No data available");return}' +
    'var csv="509 Dashboard Full Export\\n\\n";' +
    'csv+="SUMMARY METRICS\\n";' +
    'csv+="Metric,Value\\n";' +
    'csv+="Total Members,"+d.totalMembers+"\\n";' +
    'csv+="Steward Count,"+d.stewardCount+"\\n";' +
    'csv+="Total Grievances,"+d.totalCases+"\\n";' +
    'csv+="Open Cases,"+d.openCases+"\\n";' +
    'csv+="Win Rate,"+d.winRate+"%\\n";' +
    'csv+="Morale Score,"+d.moraleScore+"/10\\n";' +
    'csv+="\\nGRIEVANCES BY STATUS\\n";' +
    'csv+="Status,Count\\n";' +
    'csv+="Open,"+d.statusDistribution.open+"\\n";' +
    'csv+="Pending,"+d.statusDistribution.pending+"\\n";' +
    'csv+="Won,"+d.statusDistribution.won+"\\n";' +
    'csv+="Denied,"+d.statusDistribution.denied+"\\n";' +
    'csv+="Settled,"+d.statusDistribution.settled+"\\n";' +
    'csv+="\\nMEMBERS BY LOCATION\\n";' +
    'csv+="Location,Count\\n";' +
    'Object.keys(d.locationBreakdown).forEach(function(loc){csv+="\\""+loc+"\\","+d.locationBreakdown[loc]+"\\n"});' +
    'csv+="\\nMEMBERS BY UNIT\\n";' +
    'csv+="Unit,Count\\n";' +
    'Object.keys(d.unitBreakdown).forEach(function(u){csv+="\\""+u+"\\","+d.unitBreakdown[u]+"\\n"});' +
    'var blob=new Blob([csv],{type:"text/csv"});var url=URL.createObjectURL(blob);' +
    'var a=document.createElement("a");a.href=url;a.download="509_full_export_"+new Date().toISOString().split("T")[0]+".csv";' +
    'document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);' +
    'toggleSettings()' +
    '}' +
    'function scheduleReport(){' +
    'var email=prompt("Enter email address for scheduled reports:");' +
    'if(email){alert("Report scheduling requires server-side setup.\\n\\nEmail: "+email+"\\n\\nContact your administrator to configure automated reports.");toggleSettings()}' +
    '}' +

    // Keyboard Shortcuts
    'function showHint(msg){var hint=document.getElementById("shortcutHint");hint.textContent=msg;hint.classList.remove("show");void hint.offsetWidth;hint.classList.add("show")}' +
    'document.addEventListener("keydown",function(e){' +
    'if(e.target.tagName==="INPUT"||e.target.tagName==="TEXTAREA"||e.target.tagName==="SELECT")return;' +
    'var key=e.key.toLowerCase();' +
    'if(key>="1"&&key<="9"){' +
    'var tabs=document.querySelectorAll(".tab");var idx=parseInt(key)-1;' +
    'if(tabs[idx]){tabs[idx].click();showHint("Tab "+(idx+1)+": "+tabs[idx].textContent)}e.preventDefault()' +
    '}' +
    'else if(key==="/"){' +
    'var search=document.getElementById("stewardSearch")||document.getElementById("faqSearch");' +
    'if(search){search.focus();showHint("Search focused")}e.preventDefault()' +
    '}' +
    'else if(key==="r"&&!e.ctrlKey&&!e.metaKey){location.reload();e.preventDefault()}' +
    'else if(key==="p"&&!e.ctrlKey&&!e.metaKey){printDashboard();e.preventDefault()}' +
    'else if(key==="a"){toggleAlerts();showHint("Alert Center");e.preventDefault()}' +
    'else if(key==="s"&&!e.ctrlKey&&!e.metaKey){toggleSettings();showHint("Settings");e.preventDefault()}' +
    'else if(key==="escape"){' +
    'document.getElementById("settingsPanel").classList.remove("open");' +
    'document.querySelector(".settings-overlay").classList.remove("open");' +
    'document.getElementById("alertCenter").classList.remove("open");' +
    'closeModal()' +
    '}' +
    '});' +

    // Last Updated Indicator
    'function updateLastUpdated(){' +
    'var el=document.getElementById("lastUpdated");if(el){' +
    'var now=new Date();var time=now.toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"});' +
    'el.innerHTML="<i class=\\"material-icons\\">schedule</i><span>Updated "+time+"</span>"' +
    '}}' +

    // Trend Arrow Helper
    'function getTrendArrow(current,previous){' +
    'if(previous===0||previous===undefined)return"<span class=\\"trend-arrow flat\\">―</span>";' +
    'var pct=Math.round(((current-previous)/previous)*100);' +
    'if(pct>5)return"<span class=\\"trend-arrow up\\">↑</span><span class=\\"trend-pct\\">+"+pct+"%</span>";' +
    'if(pct<-5)return"<span class=\\"trend-arrow down\\">↓</span><span class=\\"trend-pct\\">"+pct+"%</span>";' +
    'return"<span class=\\"trend-arrow flat\\">―</span>"' +
    '}' +

    // Goal Progress Helper
    'function getGoalBar(current,target,color){' +
    'var pct=Math.min(100,Math.round((current/target)*100));' +
    'var barColor=pct>=100?"#22c55e":pct>=75?"#3b82f6":pct>=50?"#f59e0b":"#ef4444";' +
    'return"<div class=\\"goal-bar\\"><div class=\\"goal-fill\\" style=\\"width:"+pct+"%;background:"+barColor+"\\"></div></div><div class=\\"goal-label\\"><span>"+current+"</span><span>Target: "+target+"</span></div>"' +
    '}' +

    // Initialize on load
    'window.addEventListener("load",function(){' +
    'updateLastUpdated();' +
    'loadGoals();' +
    // Load saved view mode
    'var savedView=localStorage.getItem("509_viewMode");' +
    'if(savedView&&savedView!=="auto"){setViewMode(savedView)}' +
    'if(localStorage.getItem("509_highContrast")==="true"){document.body.classList.add("high-contrast");document.getElementById("highContrastToggle").classList.add("on")}' +
    'if(localStorage.getItem("509_largeText")==="true"){document.body.classList.add("large-text");document.getElementById("largeTextToggle").classList.add("on")}' +
    'if(localStorage.getItem("509_autoRefresh")==="true"){document.getElementById("autoRefreshToggle").classList.add("on");startAutoRefresh()}' +
    'setTimeout(populateAlerts,1000)' +
    '});' +

    '</script></body></html>';
}
