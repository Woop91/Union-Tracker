/**
 * ============================================================================
 * 17_CorrelationEngine.gs - Cross-Dimensional Correlation Engine
 * ============================================================================
 *
 * Statistical correlation engine for union dashboard analytics. Computes
 * relationships between member engagement, grievance outcomes, satisfaction
 * scores, and organizational dimensions.
 *
 * Design principles:
 *   - ~10 curated, union-relevant correlation pairs (not an N×N matrix)
 *   - Confidence reporting alongside every coefficient
 *   - Plain-language insight strings for non-statistician users
 *   - "Associated with" framing, never "causes"
 *
 * @version 4.9.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// STATISTICAL PRIMITIVES
// ============================================================================

/**
 * Calculates the arithmetic mean of a numeric array
 * @param {number[]} arr - Input values
 * @returns {number} Mean, or 0 for empty arrays
 */
function statMean_(arr) {
  if (!arr || arr.length === 0) return 0;
  var sum = 0;
  for (var i = 0; i < arr.length; i++) sum += arr[i];
  return sum / arr.length;
}

/**
 * Calculates the standard deviation of a numeric array
 * @param {number[]} arr - Input values
 * @returns {number} Population standard deviation
 */
function statStdDev_(arr) {
  if (!arr || arr.length < 2) return 0;
  var mean = statMean_(arr);
  var sumSq = 0;
  for (var i = 0; i < arr.length; i++) {
    var diff = arr[i] - mean;
    sumSq += diff * diff;
  }
  return Math.sqrt(sumSq / arr.length);
}

/**
 * Calculates Pearson correlation coefficient between two numeric arrays.
 * Returns r in [-1, 1]. Requires paired arrays of equal length.
 *
 * @param {number[]} x - First variable values
 * @param {number[]} y - Second variable values
 * @returns {number} Pearson r, or 0 if computation is impossible
 */
function pearsonCorrelation_(x, y) {
  var n = Math.min(x.length, y.length);
  if (n < 3) return 0; // Need at least 3 pairs for meaningful correlation

  var meanX = statMean_(x.slice(0, n));
  var meanY = statMean_(y.slice(0, n));

  var sumXY = 0, sumX2 = 0, sumY2 = 0;
  for (var i = 0; i < n; i++) {
    var dx = x[i] - meanX;
    var dy = y[i] - meanY;
    sumXY += dx * dy;
    sumX2 += dx * dx;
    sumY2 += dy * dy;
  }

  var denom = Math.sqrt(sumX2 * sumY2);
  if (denom === 0) return 0;
  return sumXY / denom;
}

/**
 * Calculates Spearman rank correlation (for ordinal or non-normal data).
 * Converts values to ranks then applies Pearson on ranks.
 *
 * @param {number[]} x - First variable values
 * @param {number[]} y - Second variable values
 * @returns {number} Spearman rho in [-1, 1]
 */
function spearmanCorrelation_(x, y) {
  var n = Math.min(x.length, y.length);
  if (n < 3) return 0;
  return pearsonCorrelation_(toRanks_(x.slice(0, n)), toRanks_(y.slice(0, n)));
}

/**
 * Converts an array of values to their fractional ranks.
 * Ties receive the average of their positions.
 *
 * @param {number[]} arr - Input values
 * @returns {number[]} Rank array (1-based)
 * @private
 */
function toRanks_(arr) {
  var indexed = [];
  for (var i = 0; i < arr.length; i++) {
    indexed.push({ val: arr[i], idx: i });
  }
  indexed.sort(function(a, b) { return a.val - b.val; });

  var ranks = new Array(arr.length);
  var i2 = 0;
  while (i2 < indexed.length) {
    var j = i2;
    while (j < indexed.length && indexed[j].val === indexed[i2].val) j++;
    var avgRank = (i2 + j + 1) / 2; // Average rank for ties
    for (var k = i2; k < j; k++) {
      ranks[indexed[k].idx] = avgRank;
    }
    i2 = j;
  }
  return ranks;
}

/**
 * Calculates a simple chi-square statistic for a 2D contingency table.
 * Used for categorical-vs-categorical associations (e.g., location vs outcome).
 *
 * @param {Object} contingency - { rowKey: { colKey: count } }
 * @returns {Object} { chiSquare, cramersV, degreesOfFreedom }
 */
function chiSquareTest_(contingency) {
  var rowKeys = Object.keys(contingency);
  if (rowKeys.length < 2) return { chiSquare: 0, cramersV: 0, degreesOfFreedom: 0 };

  // Collect all column keys
  var colKeySet = {};
  for (var r = 0; r < rowKeys.length; r++) {
    var cols = contingency[rowKeys[r]];
    for (var c in cols) {
      if (cols.hasOwnProperty(c)) colKeySet[c] = true;
    }
  }
  var colKeys = Object.keys(colKeySet);
  if (colKeys.length < 2) return { chiSquare: 0, cramersV: 0, degreesOfFreedom: 0 };

  // Build observed matrix and marginals
  var rowTotals = [];
  var colTotals = new Array(colKeys.length);
  for (var ci = 0; ci < colKeys.length; ci++) colTotals[ci] = 0;
  var grandTotal = 0;

  var observed = [];
  for (var ri = 0; ri < rowKeys.length; ri++) {
    var row = [];
    var rowSum = 0;
    for (var cj = 0; cj < colKeys.length; cj++) {
      var val = (contingency[rowKeys[ri]] && contingency[rowKeys[ri]][colKeys[cj]]) || 0;
      row.push(val);
      rowSum += val;
      colTotals[cj] += val;
    }
    observed.push(row);
    rowTotals.push(rowSum);
    grandTotal += rowSum;
  }

  if (grandTotal === 0) return { chiSquare: 0, cramersV: 0, degreesOfFreedom: 0 };

  // Calculate chi-square
  var chiSq = 0;
  for (var ri2 = 0; ri2 < rowKeys.length; ri2++) {
    for (var cj2 = 0; cj2 < colKeys.length; cj2++) {
      var expected = (rowTotals[ri2] * colTotals[cj2]) / grandTotal;
      if (expected > 0) {
        var diff = observed[ri2][cj2] - expected;
        chiSq += (diff * diff) / expected;
      }
    }
  }

  var df = (rowKeys.length - 1) * (colKeys.length - 1);
  var minDim = Math.min(rowKeys.length - 1, colKeys.length - 1);
  var cramersV = minDim > 0 ? Math.sqrt(chiSq / (grandTotal * minDim)) : 0;

  return { chiSquare: Math.round(chiSq * 100) / 100, cramersV: Math.round(cramersV * 1000) / 1000, degreesOfFreedom: df };
}

/**
 * Classifies correlation strength and returns a plain-language label.
 *
 * @param {number} r - Correlation coefficient (Pearson r or Spearman rho)
 * @param {number} n - Sample size
 * @returns {Object} { strength: string, confidence: string, reliable: boolean }
 */
function classifyCorrelation_(r, n) {
  var absR = Math.abs(r);
  var strength;

  if (absR >= 0.7) strength = 'strong';
  else if (absR >= 0.4) strength = 'moderate';
  else if (absR >= 0.2) strength = 'weak';
  else strength = 'negligible';

  // Simple confidence based on sample size
  var confidence;
  var reliable;
  if (n >= 30) {
    confidence = 'high';
    reliable = absR >= 0.2;
  } else if (n >= 15) {
    confidence = 'moderate';
    reliable = absR >= 0.35;
  } else if (n >= 5) {
    confidence = 'low';
    reliable = absR >= 0.5;
  } else {
    confidence = 'insufficient';
    reliable = false;
  }

  return { strength: strength, confidence: confidence, reliable: reliable };
}


// ============================================================================
// DATA EXTRACTION HELPERS
// ============================================================================

/**
 * Extracts paired numeric arrays from per-group aggregate data.
 * Used to build [x, y] pairs from objects like { "Location A": { rate: 0.6, count: 10 }, ... }
 *
 * @param {Object} groupedData - Keyed object with numeric fields
 * @param {string} fieldX - Field name for X dimension
 * @param {string} fieldY - Field name for Y dimension
 * @param {number} [minCount] - Minimum sample size per group to include
 * @returns {Object} { x: number[], y: number[], labels: string[], n: number }
 * @private
 */
function extractPairs_(groupedData, fieldX, fieldY, minCount) {
  minCount = minCount || 1;
  var x = [], y = [], labels = [];

  for (var key in groupedData) {
    if (!groupedData.hasOwnProperty(key)) continue;
    var group = groupedData[key];
    var count = group.count || group.n || 1;
    if (count < minCount) continue;

    var valX = parseFloat(group[fieldX]);
    var valY = parseFloat(group[fieldY]);
    if (isNaN(valX) || isNaN(valY)) continue;

    x.push(valX);
    y.push(valY);
    labels.push(key);
  }

  return { x: x, y: y, labels: labels, n: x.length };
}


// ============================================================================
// CURATED CORRELATION PAIRS
// ============================================================================

/**
 * Runs all curated correlation analyses against the current dashboard data.
 * Returns structured results with plain-language insights.
 *
 * @param {boolean} isPII - Whether to include PII in drill-down details
 * @returns {string} JSON array of correlation results
 */
function getCorrelationInsights(isPII) {
  var data = JSON.parse(getUnifiedDashboardData(isTruthyValue(isPII)));
  var insights = [];

  // 1. Location Satisfaction vs Grievance Rate
  insights.push(correlateLocationSatVsGrievance_(data));

  // 2. Steward Win Rate vs Assigned Member Engagement
  insights.push(correlateStewardWinVsEngagement_(data));

  // 3. Issue Category vs Resolution Time
  insights.push(correlateCategoryVsResolutionTime_(data));

  // 4. Meeting Attendance vs Satisfaction
  insights.push(correlateEngagementVsSatisfaction_(data));

  // 5. Location Engagement vs Grievance Concentration
  insights.push(correlateEngagementVsGrievance_(data));

  // 6. Article Violated vs Win Rate
  insights.push(correlateArticleVsOutcome_(data));

  // 7. Step Level vs Resolution Time
  insights.push(correlateStepVsTime_(data));

  // 8. Unit Size vs Satisfaction
  insights.push(correlateUnitSizeVsSatisfaction_(data));

  // 9. Steward Caseload vs Win Rate
  insights.push(correlateCaseloadVsWinRate_(data));

  // 10. Volunteer Hours vs Meeting Attendance (by location)
  insights.push(correlateVolunteerVsEngagement_(data));

  // Filter out any null results (insufficient data)
  insights = insights.filter(function(i) { return i !== null; });

  return JSON.stringify(insights);
}


// ============================================================================
// INDIVIDUAL CORRELATION IMPLEMENTATIONS
// ============================================================================

/**
 * Correlation 1: Location satisfaction score vs grievance filing rate
 * Question: Are dissatisfied locations filing more grievances?
 */
function correlateLocationSatVsGrievance_(data) {
  var satByLoc = data.satisfactionByLocation || {};
  var locBreakdown = data.locationBreakdown || {};
  var drillDown = (data.chartDrillDown && data.chartDrillDown.locationByCase) || {};

  var x = [], y = [], labels = [];

  for (var loc in satByLoc) {
    if (!satByLoc.hasOwnProperty(loc)) continue;
    var satScore = satByLoc[loc].score || 0;
    var memberCount = locBreakdown[loc] || 0;
    if (memberCount < 3) continue; // Need enough members

    var grievanceCount = drillDown[loc] ? drillDown[loc].length : 0;
    var grievanceRate = grievanceCount / memberCount;

    x.push(satScore);
    y.push(grievanceRate);
    labels.push(loc);
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'loc_sat_vs_grievance',
    title: 'Satisfaction vs Grievance Rate by Location',
    description: 'Examines whether locations with lower satisfaction scores are associated with higher grievance filing rates.',
    xLabel: 'Satisfaction Score (1-10)',
    yLabel: 'Grievances per Member',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('satisfaction', 'grievance filing rate', r, cls, 'location')
  };
}

/**
 * Correlation 2: Steward win rate vs member engagement of their assigned members
 * Question: Do effective stewards have more engaged members (or vice versa)?
 */
function correlateStewardWinVsEngagement_(data) {
  var stewards = data.stewardPerformance || [];
  var partByLoc = data.participationByLocation || {};

  if (stewards.length < 3) return null;

  var x = [], y = [], labels = [];

  for (var i = 0; i < stewards.length; i++) {
    var s = stewards[i];
    if (!s.totalCases || s.totalCases < 2) continue;
    var winRate = s.winRate || 0;

    // Approximate engagement from location-level data
    var engagementRate = 0;
    var locData = partByLoc[s.location] || partByLoc[s.name];
    if (locData) {
      engagementRate = ((locData.emailRate || 0) + (locData.meetingRate || 0)) / 2;
    }

    x.push(winRate);
    y.push(engagementRate);
    labels.push(s.name || 'Steward ' + i);
  }

  if (x.length < 3) return null;

  var r = spearmanCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'steward_win_vs_engagement',
    title: 'Steward Win Rate vs Member Engagement',
    description: 'Examines whether stewards with higher win rates are associated with more engaged membership in their area.',
    xLabel: 'Win Rate (%)',
    yLabel: 'Engagement Rate (%)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('steward win rate', 'member engagement', r, cls, 'steward')
  };
}

/**
 * Correlation 3: Issue category vs average resolution time
 * Question: Which types of grievances take longest to resolve?
 */
function correlateCategoryVsResolutionTime_(data) {
  var categories = data.grievancesByCategory || {};
  var drillDown = (data.chartDrillDown && data.chartDrillDown.categoryByCase) || {};

  var catData = {};
  for (var cat in drillDown) {
    if (!drillDown.hasOwnProperty(cat)) continue;
    var cases = drillDown[cat];
    var resolvedDays = [];
    var winCount = 0;

    for (var i = 0; i < cases.length; i++) {
      if (cases[i].daysOpen && cases[i].daysOpen > 0) {
        resolvedDays.push(cases[i].daysOpen);
      }
      if (cases[i].status === 'Won' || cases[i].status === 'won') {
        winCount++;
      }
    }

    if (resolvedDays.length >= 2) {
      catData[cat] = {
        avgDays: statMean_(resolvedDays),
        winRate: cases.length > 0 ? (winCount / cases.length) * 100 : 0,
        count: cases.length
      };
    }
  }

  var catKeys = Object.keys(catData);
  if (catKeys.length < 3) return null;

  var x = [], y = [], labels = [];
  for (var ci = 0; ci < catKeys.length; ci++) {
    x.push(catData[catKeys[ci]].count);
    y.push(catData[catKeys[ci]].avgDays);
    labels.push(catKeys[ci]);
  }

  var r = spearmanCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'category_vs_resolution',
    title: 'Grievance Category vs Resolution Time',
    description: 'Examines which types of grievances are associated with longer resolution times.',
    xLabel: 'Case Volume',
    yLabel: 'Avg Days to Resolution',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    categoryBreakdown: catData,
    insight: generateInsight_('case volume', 'resolution time', r, cls, 'category')
  };
}

/**
 * Correlation 4: Location engagement (email + meeting) vs satisfaction score
 * Question: Do more engaged locations report higher satisfaction?
 */
function correlateEngagementVsSatisfaction_(data) {
  var partByLoc = data.participationByLocation || {};
  var satByLoc = data.satisfactionByLocation || {};

  var x = [], y = [], labels = [];

  for (var loc in partByLoc) {
    if (!partByLoc.hasOwnProperty(loc)) continue;
    var part = partByLoc[loc];
    if (!part.count || part.count < 3) continue;

    var engagement = ((part.emailRate || 0) + (part.meetingRate || 0)) / 2;
    var sat = satByLoc[loc] ? satByLoc[loc].score : null;
    if (sat === null || sat === 0) continue;

    x.push(engagement);
    y.push(sat);
    labels.push(loc);
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'engagement_vs_satisfaction',
    title: 'Member Engagement vs Satisfaction by Location',
    description: 'Examines whether locations with higher engagement (email open rates, meeting attendance) are associated with higher satisfaction scores.',
    xLabel: 'Engagement Rate (%)',
    yLabel: 'Satisfaction Score (1-10)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('member engagement', 'satisfaction', r, cls, 'location')
  };
}

/**
 * Correlation 5: Location engagement rate vs grievance concentration
 * Question: Do low-engagement areas have more grievances (hidden problems)?
 */
function correlateEngagementVsGrievance_(data) {
  var partByLoc = data.participationByLocation || {};
  var locBreakdown = data.locationBreakdown || {};
  var drillDown = (data.chartDrillDown && data.chartDrillDown.locationByCase) || {};

  var x = [], y = [], labels = [];

  for (var loc in partByLoc) {
    if (!partByLoc.hasOwnProperty(loc)) continue;
    var part = partByLoc[loc];
    if (!part.count || part.count < 3) continue;

    var engagement = ((part.emailRate || 0) + (part.meetingRate || 0)) / 2;
    var memberCount = locBreakdown[loc] || part.count;
    var grievanceCount = drillDown[loc] ? drillDown[loc].length : 0;
    var grievanceRate = grievanceCount / memberCount;

    x.push(engagement);
    y.push(grievanceRate);
    labels.push(loc);
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'engagement_vs_grievance',
    title: 'Engagement Rate vs Grievance Concentration',
    description: 'Examines whether locations with lower engagement are associated with more grievances per member.',
    xLabel: 'Engagement Rate (%)',
    yLabel: 'Grievances per Member',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('engagement rate', 'grievance concentration', r, cls, 'location')
  };
}

/**
 * Correlation 6: Article violated vs win rate
 * Question: Which contract articles have the strongest enforcement track record?
 */
function correlateArticleVsOutcome_(data) {
  var drillDown = (data.chartDrillDown && data.chartDrillDown.statusByCase) || {};

  // Build per-article outcome tallies
  var articleOutcomes = {};
  var allStatuses = ['open', 'pending', 'won', 'denied', 'settled', 'withdrawn'];

  for (var si = 0; si < allStatuses.length; si++) {
    var status = allStatuses[si];
    var cases = drillDown[status] || [];
    for (var ci = 0; ci < cases.length; ci++) {
      var articles = cases[ci].articles || cases[ci].article || '';
      if (!articles) continue;
      var artList = String(articles).split(',');
      for (var ai = 0; ai < artList.length; ai++) {
        var art = artList[ai].trim();
        if (!art) continue;
        if (!articleOutcomes[art]) articleOutcomes[art] = { total: 0, won: 0, denied: 0, settled: 0 };
        articleOutcomes[art].total++;
        if (status === 'won') articleOutcomes[art].won++;
        if (status === 'denied') articleOutcomes[art].denied++;
        if (status === 'settled') articleOutcomes[art].settled++;
      }
    }
  }

  var artKeys = Object.keys(articleOutcomes).filter(function(k) {
    return articleOutcomes[k].total >= 2;
  });

  if (artKeys.length < 3) {
    // Not enough articles for correlation, but still useful as breakdown
    return {
      id: 'article_vs_outcome',
      title: 'Contract Article vs Win Rate',
      description: 'Shows which contract articles have the highest win rates when grieved.',
      xLabel: 'Case Volume',
      yLabel: 'Win Rate (%)',
      r: 0,
      direction: 'n/a',
      strength: 'insufficient data',
      confidence: 'insufficient',
      reliable: false,
      sampleSize: artKeys.length,
      dataPoints: [],
      articleBreakdown: articleOutcomes,
      insight: 'Not enough article types with multiple cases to compute a correlation. Review the breakdown for individual article performance.'
    };
  }

  var x = [], y = [], labels = [];
  for (var ki = 0; ki < artKeys.length; ki++) {
    var ao = articleOutcomes[artKeys[ki]];
    x.push(ao.total);
    y.push(ao.total > 0 ? (ao.won / ao.total) * 100 : 0);
    labels.push(artKeys[ki]);
  }

  var r = spearmanCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'article_vs_outcome',
    title: 'Contract Article vs Win Rate',
    description: 'Examines whether more frequently violated articles are associated with different win rates.',
    xLabel: 'Times Violated',
    yLabel: 'Win Rate (%)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    articleBreakdown: articleOutcomes,
    insight: generateInsight_('violation frequency', 'win rate', r, cls, 'contract article')
  };
}

/**
 * Correlation 7: Grievance step level vs average resolution time
 * Question: How much longer do cases take at higher steps?
 */
function correlateStepVsTime_(data) {
  var avgDays = data.avgDaysAtStep || {};
  var stepProg = data.stepProgression || {};

  var stepMap = [
    { label: 'Step I', step: 1, days: avgDays.step1 || 0, count: stepProg.step1 || 0 },
    { label: 'Step II', step: 2, days: avgDays.step2 || 0, count: stepProg.step2 || 0 },
    { label: 'Step III', step: 3, days: avgDays.step3 || 0, count: stepProg.step3 || 0 },
    { label: 'Arbitration', step: 4, days: avgDays.arb || 0, count: stepProg.arb || 0 }
  ];

  var x = [], y = [], labels = [];
  for (var i = 0; i < stepMap.length; i++) {
    if (stepMap[i].count > 0) {
      x.push(stepMap[i].step);
      y.push(stepMap[i].days);
      labels.push(stepMap[i].label);
    }
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'step_vs_time',
    title: 'Grievance Step Level vs Resolution Time',
    description: 'Examines how resolution time increases at higher grievance steps.',
    xLabel: 'Step Level',
    yLabel: 'Avg Days to Resolve',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    stepBreakdown: stepMap,
    insight: generateInsight_('step level', 'resolution time', r, cls, 'grievance step')
  };
}

/**
 * Correlation 8: Unit size vs satisfaction score
 * Question: Are larger or smaller units more satisfied?
 */
function correlateUnitSizeVsSatisfaction_(data) {
  var unitBreakdown = data.unitBreakdown || {};
  var satByUnit = data.satisfactionByUnit || {};

  var x = [], y = [], labels = [];

  for (var unit in unitBreakdown) {
    if (!unitBreakdown.hasOwnProperty(unit)) continue;
    var count = unitBreakdown[unit] || 0;
    if (count < 2) continue;

    var sat = satByUnit[unit] ? satByUnit[unit].score : null;
    if (sat === null || sat === 0) continue;

    x.push(count);
    y.push(sat);
    labels.push(unit);
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'unit_size_vs_satisfaction',
    title: 'Unit Size vs Satisfaction Score',
    description: 'Examines whether larger units are associated with different satisfaction levels than smaller ones.',
    xLabel: 'Members in Unit',
    yLabel: 'Satisfaction Score (1-10)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('unit size', 'satisfaction', r, cls, 'unit')
  };
}

/**
 * Correlation 9: Steward caseload vs win rate
 * Question: Are overloaded stewards less effective?
 */
function correlateCaseloadVsWinRate_(data) {
  var stewards = data.stewardPerformance || [];
  if (stewards.length < 3) return null;

  var x = [], y = [], labels = [];

  for (var i = 0; i < stewards.length; i++) {
    var s = stewards[i];
    if (!s.totalCases || s.totalCases < 1) continue;
    x.push(s.activeCases || s.totalCases);
    y.push(s.winRate || 0);
    labels.push(s.name || 'Steward ' + i);
  }

  if (x.length < 3) return null;

  var r = spearmanCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'caseload_vs_winrate',
    title: 'Steward Caseload vs Win Rate',
    description: 'Examines whether stewards carrying more cases are associated with lower win rates.',
    xLabel: 'Active Cases',
    yLabel: 'Win Rate (%)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('caseload size', 'win rate', r, cls, 'steward')
  };
}

/**
 * Correlation 10: Volunteer hours rate vs engagement rate by location
 * Question: Are locations with more volunteer activity also more engaged overall?
 */
function correlateVolunteerVsEngagement_(data) {
  var partByLoc = data.participationByLocation || {};

  // We need volunteer data per location — approximate from engagement data
  var x = [], y = [], labels = [];

  for (var loc in partByLoc) {
    if (!partByLoc.hasOwnProperty(loc)) continue;
    var part = partByLoc[loc];
    if (!part.count || part.count < 3) continue;

    var emailRate = part.emailRate || 0;
    var meetingRate = part.meetingRate || 0;

    // Use email rate as one proxy, meeting rate as another
    x.push(emailRate);
    y.push(meetingRate);
    labels.push(loc);
  }

  if (x.length < 3) return null;

  var r = pearsonCorrelation_(x, y);
  var cls = classifyCorrelation_(r, x.length);

  return {
    id: 'email_vs_meeting_engagement',
    title: 'Email Engagement vs Meeting Attendance',
    description: 'Examines whether locations with higher email open rates also have higher meeting attendance, indicating consistent engagement patterns.',
    xLabel: 'Email Open Rate (%)',
    yLabel: 'Meeting Attendance Rate (%)',
    r: Math.round(r * 1000) / 1000,
    direction: r < 0 ? 'inverse' : 'positive',
    strength: cls.strength,
    confidence: cls.confidence,
    reliable: cls.reliable,
    sampleSize: x.length,
    dataPoints: buildDataPoints_(labels, x, y),
    insight: generateInsight_('email engagement', 'meeting attendance', r, cls, 'location')
  };
}


// ============================================================================
// INSIGHT GENERATION
// ============================================================================

/**
 * Generates a plain-language insight string for a correlation result.
 *
 * @param {string} varX - Human label for X variable
 * @param {string} varY - Human label for Y variable
 * @param {number} r - Correlation coefficient
 * @param {Object} cls - Classification from classifyCorrelation_()
 * @param {string} dimension - Grouping dimension (e.g., 'location', 'steward')
 * @returns {string} Plain-language insight
 * @private
 */
function generateInsight_(varX, varY, r, cls, dimension) {
  if (!cls.reliable) {
    if (cls.confidence === 'insufficient') {
      return 'Not enough data points across ' + dimension + 's to draw conclusions. Revisit when more data is available.';
    }
    return 'The association between ' + varX + ' and ' + varY + ' across ' + dimension +
           's is too weak (r=' + (Math.round(r * 100) / 100) + ') or sample too small (n=' +
           ') to be meaningful.';
  }

  var direction = r > 0 ? 'higher' : 'lower';
  var assoc = r > 0 ? 'tend to also have higher' : 'tend to have lower';

  if (cls.strength === 'strong') {
    return 'Strong pattern: ' + dimension + 's with ' + direction + ' ' + varX +
           ' ' + assoc + ' ' + varY + '. This is one of the clearest relationships in the data (' +
           cls.confidence + ' confidence).';
  }

  if (cls.strength === 'moderate') {
    return 'Notable trend: ' + dimension + 's with ' + direction + ' ' + varX +
           ' are generally associated with ' + (r > 0 ? 'higher' : 'lower') + ' ' + varY +
           '. Worth monitoring (' + cls.confidence + ' confidence).';
  }

  return 'Slight tendency: ' + dimension + 's with ' + direction + ' ' + varX +
         ' show a weak association with ' + (r > 0 ? 'higher' : 'lower') + ' ' + varY +
         '. Not strong enough to act on alone (' + cls.confidence + ' confidence).';
}

/**
 * Builds data point objects for scatter plot rendering.
 *
 * @param {string[]} labels - Point labels (e.g., location names)
 * @param {number[]} x - X values
 * @param {number[]} y - Y values
 * @returns {Object[]} Array of { label, x, y }
 * @private
 */
function buildDataPoints_(labels, x, y) {
  var points = [];
  for (var i = 0; i < labels.length; i++) {
    points.push({
      label: labels[i],
      x: Math.round(x[i] * 100) / 100,
      y: Math.round(y[i] * 100) / 100
    });
  }
  return points;
}


// ============================================================================
// HOTSPOT INTEGRATION
// ============================================================================

/**
 * Generates correlation-based alerts for the hotspot/alert system.
 * Only returns alerts for strong, reliable correlations that suggest action.
 *
 * @param {boolean} isPII - Whether to include PII
 * @returns {string} JSON array of alert objects
 */
function getCorrelationAlerts(isPII) {
  var insightsJson = getCorrelationInsights(isPII);
  var insights = JSON.parse(insightsJson);
  var alerts = [];

  for (var i = 0; i < insights.length; i++) {
    var insight = insights[i];

    // Only alert on strong or moderate reliable correlations
    if (!insight.reliable) continue;
    if (insight.strength !== 'strong' && insight.strength !== 'moderate') continue;

    var severity = insight.strength === 'strong' ? 'high' : 'medium';
    var icon = insight.direction === 'inverse' ? 'trending_down' : 'trending_up';

    alerts.push({
      id: 'corr_' + insight.id,
      type: 'correlation',
      severity: severity,
      icon: icon,
      title: insight.title,
      message: insight.insight,
      r: insight.r,
      sampleSize: insight.sampleSize,
      actionable: insight.strength === 'strong'
    });
  }

  EventBus.emit('correlation:alertsGenerated', { count: alerts.length });
  return JSON.stringify(alerts);
}


// ============================================================================
// SUMMARY / OVERVIEW
// ============================================================================

/**
 * Returns a summary of all correlations with their status.
 * Lighter-weight than full insights — for the overview card.
 *
 * @param {boolean} isPII - Whether to include PII
 * @returns {string} JSON summary object
 */
function getCorrelationSummary(isPII) {
  var insights = JSON.parse(getCorrelationInsights(isPII));

  var summary = {
    total: insights.length,
    strong: 0,
    moderate: 0,
    weak: 0,
    negligible: 0,
    insufficientData: 0,
    topInsights: [],
    actionableCount: 0
  };

  for (var i = 0; i < insights.length; i++) {
    var ins = insights[i];
    if (ins.strength === 'strong') { summary.strong++; summary.actionableCount++; }
    else if (ins.strength === 'moderate') { summary.moderate++; summary.actionableCount++; }
    else if (ins.strength === 'weak') summary.weak++;
    else if (ins.strength === 'insufficient data') summary.insufficientData++;
    else summary.negligible++;

    // Top insights = reliable ones sorted by |r|
    if (ins.reliable) {
      summary.topInsights.push({
        id: ins.id,
        title: ins.title,
        r: ins.r,
        strength: ins.strength,
        insight: ins.insight
      });
    }
  }

  // Sort top insights by |r| descending
  summary.topInsights.sort(function(a, b) {
    return Math.abs(b.r) - Math.abs(a.r);
  });

  // Keep only top 5
  summary.topInsights = summary.topInsights.slice(0, 5);

  return JSON.stringify(summary);
}
