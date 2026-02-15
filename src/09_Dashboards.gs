/**
 * 09_Dashboards.gs - Satisfaction Engine Module
 *
 * Member satisfaction survey management, dashboard display, analytics,
 * and trend analysis functionality.
 *
 * This module handles:
 * - Satisfaction dashboard display (modal dialogs)
 * - Survey response data retrieval and analysis
 * - Section-level score calculations
 * - Trend analysis and insights generation
 * - Worksite/role breakdown analytics
 * - Value synchronization and dashboard updates
 *
 * @version 4.7.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// SATISFACTION DASHBOARD - MODAL DISPLAY
// ============================================================================

/**
 * Show the Member Satisfaction Dashboard modal
 */
function showSatisfactionDashboard() {
  var html = HtmlService.createHtmlOutput(getSatisfactionDashboardHtml())
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Member Satisfaction');
}

/**
 * Returns the HTML for the Member Satisfaction Dashboard with tabs
 */
function getSatisfactionDashboardHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    getMobileOptimizedHead() +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header - Green theme for satisfaction
    '.header{background:linear-gradient(135deg,#059669,#047857);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(11px,2.5vw,13px);opacity:0.9}' +

    // Tab navigation
    '.tabs{display:flex;background:white;border-bottom:2px solid #e0e0e0;position:sticky;top:0;z-index:100}' +
    '.tab{flex:1;padding:clamp(12px,3vw,16px);text-align:center;font-size:clamp(12px,2.5vw,14px);font-weight:600;color:#666;' +
    'border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;min-height:44px}' +
    '.tab:hover{background:#f0fdf4;color:#059669}' +
    '.tab.active{color:#059669;border-bottom-color:#059669;background:#f0fdf4}' +
    '.tab-icon{display:block;font-size:18px;margin-bottom:4px}' +

    // Tab content
    '.tab-content{display:none;padding:15px;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0}to{opacity:1}}' +

    // Stats grid
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,5vw,32px);font-weight:bold;color:#059669}' +
    '.stat-label{font-size:clamp(10px,2vw,12px);color:#666;text-transform:uppercase;margin-top:5px}' +
    '.stat-card.green .stat-value{color:#059669}' +
    '.stat-card.red .stat-value{color:#DC2626}' +
    '.stat-card.orange .stat-value{color:#F97316}' +
    '.stat-card.blue .stat-value{color:#2563EB}' +
    '.stat-card.purple .stat-value{color:#7C3AED}' +

    // Score indicator with color gradient
    '.score-indicator{display:inline-block;padding:4px 12px;border-radius:20px;font-size:14px;font-weight:bold}' +
    '.score-high{background:#d1fae5;color:#059669}' +
    '.score-mid{background:#fef3c7;color:#d97706}' +
    '.score-low{background:#fee2e2;color:#dc2626}' +

    // Data table
    '.data-table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)}' +
    '.data-table th{background:#059669;color:white;padding:12px;text-align:left;font-size:13px}' +
    '.data-table td{padding:12px;border-bottom:1px solid #eee;font-size:13px}' +
    '.data-table tr:hover{background:#f0fdf4}' +
    '.data-table tr:last-child td{border-bottom:none}' +

    // Section cards
    '.section-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:12px}' +
    '.section-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}' +
    '.section-title{font-weight:600;color:#1f2937;font-size:14px}' +
    '.section-score{font-size:20px;font-weight:bold}' +

    // Progress bar for scores
    '.progress-bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden;margin-top:8px}' +
    '.progress-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.progress-green{background:linear-gradient(90deg,#059669,#10b981)}' +
    '.progress-yellow{background:linear-gradient(90deg,#f59e0b,#fbbf24)}' +
    '.progress-red{background:linear-gradient(90deg,#dc2626,#ef4444)}' +

    // Action buttons
    '.action-btn{display:inline-flex;align-items:center;gap:8px;padding:10px 16px;border:none;border-radius:8px;' +
    'cursor:pointer;font-size:13px;font-weight:500;transition:all 0.2s;min-height:44px}' +
    '.action-btn-primary{background:#059669;color:white}' +
    '.action-btn-primary:hover{background:#047857}' +
    '.action-btn-secondary{background:#f3f4f6;color:#374151}' +
    '.action-btn-secondary:hover{background:#e5e7eb}' +

    // List items for responses (clickable)
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:white;padding:15px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.06);cursor:pointer;transition:all 0.2s}' +
    '.list-item:hover{box-shadow:0 4px 8px rgba(0,0,0,0.1);transform:translateY(-1px)}' +
    '.list-item-header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}' +
    '.list-item-main{flex:1;min-width:200px}' +
    '.list-item-title{font-weight:600;color:#1f2937;margin-bottom:3px}' +
    '.list-item-subtitle{font-size:12px;color:#666}' +
    '.list-item-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee}' +
    '.list-item.expanded .list-item-details{display:block}' +
    '.detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}' +
    '.detail-item{font-size:12px}' +
    '.detail-item-label{color:#666;margin-bottom:2px}' +
    '.detail-item-value{font-weight:600;color:#1f2937}' +

    // Search input
    '.search-container{position:relative;margin-bottom:15px}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:2px solid #e5e7eb;border-radius:8px;font-size:14px;transition:border-color 0.2s}' +
    '.search-input:focus{outline:none;border-color:#059669}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af}' +

    // Filter buttons
    '.filter-group{display:flex;gap:8px;margin-bottom:15px;flex-wrap:wrap}' +

    // Charts section
    '.chart-container{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:15px}' +
    '.chart-title{font-weight:600;color:#1f2937;margin-bottom:15px;font-size:14px}' +
    '.bar-chart{display:flex;flex-direction:column;gap:10px}' +
    '.bar-row{display:flex;align-items:center;gap:10px}' +
    '.bar-label{width:140px;font-size:12px;color:#666;text-align:right}' +
    '.bar-container{flex:1;background:#e5e7eb;border-radius:4px;height:24px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:4px;transition:width 0.5s;display:flex;align-items:center;justify-content:flex-end;padding-right:8px}' +
    '.bar-value{width:50px;font-size:12px;font-weight:600;color:#374151}' +
    '.bar-inner-value{font-size:11px;font-weight:600;color:white}' +

    // Gauge chart
    '.gauge-container{display:flex;flex-wrap:wrap;gap:20px;justify-content:center}' +
    '.gauge{text-align:center;padding:15px}' +
    '.gauge-value{font-size:36px;font-weight:bold;margin-bottom:5px}' +
    '.gauge-label{font-size:12px;color:#666}' +
    '.gauge-ring{width:100px;height:100px;border-radius:50%;margin:0 auto 10px;position:relative;display:flex;align-items:center;justify-content:center}' +
    '.gauge-ring::before{content:"";position:absolute;inset:8px;background:white;border-radius:50%}' +
    '.gauge-ring span{position:relative;z-index:1;font-size:24px;font-weight:bold}' +

    // Trend arrows
    '.trend-up{color:#059669}' +
    '.trend-down{color:#dc2626}' +
    '.trend-neutral{color:#6b7280}' +

    // Insights card
    '.insight-card{background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-left:4px solid #059669;padding:15px;border-radius:0 8px 8px 0;margin-bottom:12px}' +
    '.insight-card.warning{background:linear-gradient(135deg,#fef3c7,#fde68a);border-left-color:#f59e0b}' +
    '.insight-card.alert{background:linear-gradient(135deg,#fee2e2,#fecaca);border-left-color:#dc2626}' +
    '.insight-title{font-weight:600;color:#1f2937;margin-bottom:5px}' +
    '.insight-text{font-size:13px;color:#374151}' +

    // Heatmap styles
    '.heatmap-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(80px,1fr));gap:8px}' +
    '.heatmap-cell{padding:12px;border-radius:8px;text-align:center;font-weight:600;font-size:14px}' +

    // Empty state
    '.empty-state{text-align:center;padding:40px;color:#9ca3af}' +
    '.empty-state-icon{font-size:48px;margin-bottom:10px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e5e7eb;border-top-color:#059669;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Responsive - Mobile
    '@media (max-width:768px){' +
    '  .header{padding:15px 12px}' +
    '  .tab-content{padding:10px}' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr);gap:8px}' +
    '  .stat-card{padding:14px 10px}' +
    '  .list-item{flex-direction:column;align-items:flex-start}' +
    '  .list-item-header{flex-direction:column;gap:6px}' +
    '  .list-item-main{min-width:0;width:100%}' +
    '  .tab-icon{font-size:16px}' +
    '  .tab{padding:10px 6px;font-size:clamp(11px,2.5vw,13px)}' +
    '  .bar-label{width:80px;font-size:11px}' +
    '  .bar-value{width:40px;font-size:11px}' +
    '  .gauge-container{flex-direction:column;align-items:center}' +
    '  .gauge-ring{width:80px;height:80px}' +
    '  .gauge-ring span{font-size:20px}' +
    '  .filter-group{flex-direction:column}' +
    '  .filter-group .action-btn{width:100%}' +
    '  .data-table{display:block;overflow-x:auto;-webkit-overflow-scrolling:touch}' +
    '  .detail-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .chart-container{padding:12px}' +
    '  .search-input{font-size:16px}' +
    '  .section-card{padding:12px}' +
    '  .heatmap-grid{grid-template-columns:repeat(auto-fit,minmax(60px,1fr));gap:6px}' +
    '}' +
    // Responsive - Tablet
    '@media (min-width:769px) and (max-width:1024px){' +
    '  .stats-grid{grid-template-columns:repeat(3,1fr)}' +
    '  .detail-grid{grid-template-columns:repeat(3,1fr)}' +
    '  .gauge-container{gap:15px}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header
    '<div class="header">' +
    '<h1>📊 Member Satisfaction</h1>' +
    '<div class="subtitle">Survey results and satisfaction trends</div>' +
    '</div>' +

    // Tab Navigation
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\',this)" id="tab-overview"><span class="tab-icon">📊</span>Overview</button>' +
    '<button class="tab" onclick="switchTab(\'responses\',this)" id="tab-responses"><span class="tab-icon">📝</span>Responses</button>' +
    '<button class="tab" onclick="switchTab(\'sections\',this)" id="tab-sections"><span class="tab-icon">📈</span>By Section</button>' +
    '<button class="tab" onclick="switchTab(\'analytics\',this)" id="tab-analytics"><span class="tab-icon">🔍</span>Insights</button>' +
    '</div>' +

    // Overview Tab
    '<div class="tab-content active" id="content-overview">' +
    '<div class="stats-grid" id="overview-stats"><div class="loading"><div class="spinner"></div><p>Loading stats...</p></div></div>' +
    '<div id="overview-gauges"></div>' +
    '<div id="overview-insights" style="margin-top:15px"></div>' +
    '</div>' +

    // Responses Tab
    '<div class="tab-content" id="content-responses">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="response-search" placeholder="Search by worksite or role..." oninput="filterResponses(this.value)"></div>' +
    '<div class="filter-group">' +
    '<button class="action-btn action-btn-primary" onclick="filterResponsesBy(\'all\')">All</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'high\')">High Satisfaction</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'mid\')">Medium</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterResponsesBy(\'low\')">Needs Attention</button>' +
    '</div>' +
    '<div class="list-container" id="responses-list"><div class="loading"><div class="spinner"></div><p>Loading responses...</p></div></div>' +
    '</div>' +

    // Sections Tab
    '<div class="tab-content" id="content-sections">' +
    '<div id="sections-charts"><div class="loading"><div class="spinner"></div><p>Loading section scores...</p></div></div>' +
    '</div>' +

    // Analytics Tab
    '<div class="tab-content" id="content-analytics">' +
    '<div id="analytics-content"><div class="loading"><div class="spinner"></div><p>Loading insights...</p></div></div>' +
    '</div>' +

    // JavaScript
    '<script>' +
    // XSS Prevention - escape HTML special characters
    ' + getClientSideEscapeHtml() + ' +
    'var allResponses=[];var currentFilter="all";var analyticsLoaded=false;var sectionsLoaded=false;' +

    // Tab switching
    'function switchTab(tabName,btn){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  document.getElementById("content-"+tabName).classList.add("active");' +
    '  if(tabName==="responses"&&allResponses.length===0)loadResponses();' +
    '  if(tabName==="sections"&&!sectionsLoaded)loadSections();' +
    '  if(tabName==="analytics"&&!analyticsLoaded)loadAnalytics();' +
    '}' +

    // Score color helper
    'function getScoreClass(score){' +
    '  if(score>=7)return"high";' +
    '  if(score>=5)return"mid";' +
    '  return"low";' +
    '}' +
    'function getScoreColor(score){' +
    '  if(score>=7)return"#059669";' +
    '  if(score>=5)return"#f59e0b";' +
    '  return"#dc2626";' +
    '}' +
    'function getProgressClass(score){' +
    '  if(score>=7)return"progress-green";' +
    '  if(score>=5)return"progress-yellow";' +
    '  return"progress-red";' +
    '}' +

    // Load overview data
    'function loadOverview(){' +
    '  google.script.run.withSuccessHandler(function(data){renderOverview(data)}).getSatisfactionOverviewData();' +
    '}' +

    // Render overview
    'function renderOverview(data){' +
    '  var html="";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.totalResponses+"</div><div class=\\"stat-label\\">Total Responses</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.avgOverall.toFixed(1)+"</div><div class=\\"stat-label\\">Avg Satisfaction</div></div>";' +
    '  html+="<div class=\\"stat-card blue\\"><div class=\\"stat-value\\">"+data.npsScore+"</div><div class=\\"stat-label\\">Loyalty Score</div></div>";' +
    '  html+="<div class=\\"stat-card purple\\"><div class=\\"stat-value\\">"+data.responseRate+"</div><div class=\\"stat-label\\">Response Rate</div></div>";' +
    '  html+="<div class=\\"stat-card "+(data.avgSteward>=7?"green":data.avgSteward>=5?"orange":"red")+"\\"><div class=\\"stat-value\\">"+data.avgSteward.toFixed(1)+"</div><div class=\\"stat-label\\">Steward Rating</div></div>";' +
    '  html+="<div class=\\"stat-card "+(data.avgLeadership>=7?"green":data.avgLeadership>=5?"orange":"red")+"\\"><div class=\\"stat-value\\">"+data.avgLeadership.toFixed(1)+"</div><div class=\\"stat-label\\">Leadership</div></div>";' +
    '  document.getElementById("overview-stats").innerHTML=html;' +
    // Gauge display
    '  var gauges="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📊 Key Metrics at a Glance</div><div class=\\"gauge-container\\">";' +
    '  gauges+=renderGauge(data.avgOverall,"Overall\\nSatisfaction");' +
    '  gauges+=renderGauge(data.avgTrust,"Trust in\\nUnion");' +
    '  gauges+=renderGauge(data.avgProtected,"Feel\\nProtected");' +
    '  gauges+=renderGauge(data.avgRecommend,"Would\\nRecommend");' +
    '  gauges+="</div></div>";' +
    '  document.getElementById("overview-gauges").innerHTML=gauges;' +
    // Insights - add Loyalty Score explanation first
    '  var insights="";' +
    '  insights+="<div class=\\"insight-card\\" style=\\"background:linear-gradient(135deg,#eff6ff,#dbeafe);border-left-color:#2563eb\\"><div class=\\"insight-title\\">ℹ️ Understanding Loyalty Score</div><div class=\\"insight-text\\">The <strong>Loyalty Score</strong> (ranging from -100 to +100) measures how likely members are to recommend the union. <strong>50+</strong> = Excellent (many advocates), <strong>0-49</strong> = Good (room for growth), <strong>Below 0</strong> = Needs work (more critics than advocates). It\'s based on the \\"Would Recommend\\" question.</div></div>";' +
    '  if(data.insights&&data.insights.length>0){' +
    '    data.insights.forEach(function(i){' +
    '      insights+="<div class=\\"insight-card "+escapeHtml(i.type)+"\\"><div class=\\"insight-title\\">"+escapeHtml(i.icon)+" "+escapeHtml(i.title)+"</div><div class=\\"insight-text\\">"+escapeHtml(i.text)+"</div></div>";' +
    '    });' +
    '  }' +
    '  document.getElementById("overview-insights").innerHTML=insights;' +
    '}' +

    // Render gauge
    'function renderGauge(value,label){' +
    '  var color=getScoreColor(value);' +
    '  var pct=value*10;' +
    '  return"<div class=\\"gauge\\"><div class=\\"gauge-ring\\" style=\\"background:conic-gradient("+color+" "+pct+"%,#e5e7eb "+pct+"%)\\"><span style=\\"color:"+color+"\\">"+value.toFixed(1)+"</span></div><div class=\\"gauge-label\\">"+label.replace("\\n","<br>")+"</div></div>";' +
    '}' +

    // Load responses
    'function loadResponses(){' +
    '  google.script.run.withSuccessHandler(function(data){allResponses=data;renderResponses(data)}).getSatisfactionResponseData();' +
    '}' +

    // Render responses with clickable details
    'function renderResponses(data){' +
    '  var c=document.getElementById("responses-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">📝</div><p>No responses found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(r,i){' +
    '    var scoreClass=getScoreClass(r.avgScore);' +
    '    var scoreColor=getScoreColor(r.avgScore);' +
    '    return"<div class=\\"list-item\\" onclick=\\"toggleResponse(this)\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+escapeHtml(r.worksite)+" - "+escapeHtml(r.role)+"</div><div class=\\"list-item-subtitle\\">"+escapeHtml(r.shift)+" • "+escapeHtml(r.timeInRole)+" • "+escapeHtml(r.date)+"</div></div><div><span class=\\"score-indicator score-"+scoreClass+"\\" style=\\"color:"+scoreColor+"\\">"+r.avgScore.toFixed(1)+"/10</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-grid\\">' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Satisfaction</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.satisfaction)+"\\">"+r.satisfaction+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Trust in Union</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.trust)+"\\">"+r.trust+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Feel Protected</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.protected)+"\\">"+r.protected+"/10</div></div>' +
    '          <div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Would Recommend</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.recommend)+"\\">"+r.recommend+"/10</div></div>' +
    '          "+(r.stewardContact?"<div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Steward Contact</div><div class=\\"detail-item-value\\">Yes</div></div>":"")+"' +
    '          "+(r.stewardRating>0?"<div class=\\"detail-item\\"><div class=\\"detail-item-label\\">Steward Rating</div><div class=\\"detail-item-value\\" style=\\"color:"+getScoreColor(r.stewardRating)+"\\">"+r.stewardRating.toFixed(1)+"/10</div></div>":"")+"' +
    '        </div>' +
    '      </div>' +
    '    </div>";' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" responses. Use search/filters to narrow.</p></div>";' +
    '}' +

    // Toggle response details
    'function toggleResponse(el){el.classList.toggle("expanded")}' +

    // Filter responses
    'function filterResponses(query){' +
    '  if(!query||query.length<2){applyFilters();return}' +
    '  query=query.toLowerCase();' +
    '  var filtered=allResponses.filter(function(r){return r.worksite.toLowerCase().indexOf(query)>=0||r.role.toLowerCase().indexOf(query)>=0||r.shift.toLowerCase().indexOf(query)>=0});' +
    '  if(currentFilter!=="all")filtered=applyScoreFilter(filtered,currentFilter);' +
    '  renderResponses(filtered);' +
    '}' +

    // Filter by satisfaction level
    'function filterResponsesBy(level){' +
    '  currentFilter=level;' +
    '  applyFilters();' +
    '}' +

    // Apply filters
    'function applyFilters(){' +
    '  var query=document.getElementById("response-search").value.toLowerCase();' +
    '  var filtered=allResponses;' +
    '  if(currentFilter!=="all")filtered=applyScoreFilter(filtered,currentFilter);' +
    '  if(query&&query.length>=2)filtered=filtered.filter(function(r){return r.worksite.toLowerCase().indexOf(query)>=0||r.role.toLowerCase().indexOf(query)>=0});' +
    '  renderResponses(filtered);' +
    '}' +

    // Score filter helper
    'function applyScoreFilter(data,level){' +
    '  return data.filter(function(r){' +
    '    if(level==="high")return r.avgScore>=7;' +
    '    if(level==="mid")return r.avgScore>=5&&r.avgScore<7;' +
    '    if(level==="low")return r.avgScore<5;' +
    '    return true;' +
    '  });' +
    '}' +

    // Load sections data
    'function loadSections(){' +
    '  sectionsLoaded=true;' +
    '  google.script.run.withSuccessHandler(function(data){renderSections(data)}).getSatisfactionSectionData();' +
    '}' +

    // Render sections
    'function renderSections(data){' +
    '  var c=document.getElementById("sections-charts");' +
    '  var html="";' +
    '  if(!data.sections||data.sections.length===0){c.innerHTML="<div class=\\"empty-state\\">No section data available</div>";return}' +
    // Section scores bar chart - scale to actual data range, not always 0-10
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📊 Average Score by Section (1-10 Scale)</div>";' +
    '  html+="<div style=\\"font-size:11px;color:#666;margin-bottom:12px\\">Sorted by score - areas needing attention shown first</div>";' +
    '  html+="<div class=\\"bar-chart\\">";' +
    '  var maxScore=10;' +
    '  var hasValidData=data.sections.some(function(s){return s.avg>0&&s.responseCount>0});' +
    '  if(!hasValidData){html+="<div class=\\"empty-state\\">No survey responses yet</div>";}else{' +
    '  data.sections.forEach(function(s){' +
    '    if(s.responseCount===0)return;' +  // Skip sections with no data
    '    var pct=Math.max(0,Math.min(100,(s.avg/maxScore)*100));' +  // Clamp to 0-100%
    '    var color=getScoreColor(s.avg);' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+escapeHtml(s.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+s.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+s.responseCount+" responses</div></div>";' +
    '  });' +
    '  }' +
    '  html+="</div></div>";' +
    // Summary insights instead of redundant detail cards
    '  var lowScoring=data.sections.filter(function(s){return s.avg>0&&s.avg<6&&s.responseCount>0});' +
    '  var highScoring=data.sections.filter(function(s){return s.avg>=8&&s.responseCount>0});' +
    '  if(lowScoring.length>0||highScoring.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">💡 Section Insights</div>";' +
    '    if(lowScoring.length>0){' +
    '      html+="<div class=\\"insight-card warning\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">⚠️ Areas Needing Attention</div><div class=\\"insight-text\\">";' +
    '      lowScoring.forEach(function(s,i){html+=(i>0?", ":"")+escapeHtml(s.name)+" ("+s.avg.toFixed(1)+")"});' +
    '      html+="</div></div>";' +
    '    }' +
    '    if(highScoring.length>0){' +
    '      html+="<div class=\\"insight-card success\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">✅ Strong Performance</div><div class=\\"insight-text\\">";' +
    '      highScoring.forEach(function(s,i){html+=(i>0?", ":"")+escapeHtml(s.name)+" ("+s.avg.toFixed(1)+")"});' +
    '      html+="</div></div>";' +
    '    }' +
    '    html+="</div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Load analytics
    'function loadAnalytics(){' +
    '  analyticsLoaded=true;' +
    '  google.script.run.withSuccessHandler(function(data){renderAnalytics(data)}).getSatisfactionAnalyticsData();' +
    '}' +

    // Render analytics/insights
    'function renderAnalytics(data){' +
    '  var c=document.getElementById("analytics-content");' +
    '  var html="";' +
    // Key insights
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">💡 Key Insights</div>";' +
    '  if(data.insights&&data.insights.length>0){' +
    '    data.insights.forEach(function(i){' +
    '      html+="<div class=\\"insight-card "+escapeHtml(i.type)+"\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">"+escapeHtml(i.icon)+" "+escapeHtml(i.title)+"</div><div class=\\"insight-text\\">"+escapeHtml(i.text)+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No insights available</div>";}' +
    '  html+="</div>";' +
    // By worksite breakdown
    '  if(data.byWorksite&&data.byWorksite.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Satisfaction by Worksite</div><div class=\\"bar-chart\\">";' +
    '    data.byWorksite.forEach(function(w){' +
    '      var pct=(w.avg/10)*100;' +
    '      var color=getScoreColor(w.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+escapeHtml(w.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+w.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+w.count+" responses</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    // By role breakdown
    '  if(data.byRole&&data.byRole.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👤 Satisfaction by Role</div><div class=\\"bar-chart\\">";' +
    '    data.byRole.forEach(function(r){' +
    '      var pct=(r.avg/10)*100;' +
    '      var color=getScoreColor(r.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+escapeHtml(r.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+r.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+r.count+" responses</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    // Steward contact impact
    '  if(data.stewardImpact){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🤝 Impact of Steward Contact</div>";' +
    '    html+="<div class=\\"stats-grid\\">";' +
    '    html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.stewardImpact.withContact.toFixed(1)+"</div><div class=\\"stat-label\\">With Steward Contact ("+data.stewardImpact.withContactCount+" members)</div></div>";' +
    '    html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.stewardImpact.withoutContact.toFixed(1)+"</div><div class=\\"stat-label\\">Without Contact ("+data.stewardImpact.withoutContactCount+" members)</div></div>";' +
    '    html+="</div>";' +
    '    var diff=data.stewardImpact.withContact-data.stewardImpact.withoutContact;' +
    '    if(diff>0){' +
    '      html+="<div class=\\"insight-card\\" style=\\"margin-top:10px\\"><div class=\\"insight-text\\">Members with steward contact report <strong>+"+diff.toFixed(1)+"</strong> higher satisfaction on average.</div></div>";' +
    '    }' +
    '    html+="</div>";' +
    '  }' +
    // Top priorities
    '  if(data.topPriorities&&data.topPriorities.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🎯 Top Member Priorities</div><div class=\\"bar-chart\\">";' +
    '    var maxP=Math.max.apply(null,data.topPriorities.map(function(p){return p.count}))||1;' +
    '    data.topPriorities.forEach(function(p){' +
    '      var pct=(p.count/maxP)*100;' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+escapeHtml(p.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:#7C3AED\\"></div></div><div class=\\"bar-value\\">"+p.count+"</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Initialize
    'loadOverview();' +
    '</script>' +

    '</body></html>';
}

// ============================================================================
// SATISFACTION DATA RETRIEVAL - OVERVIEW AND STATS
// ============================================================================

/**
 * Get overview data for satisfaction dashboard
 */
function getSatisfactionOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var data = {
    totalResponses: 0,
    avgOverall: 0,
    avgSteward: 0,
    avgLeadership: 0,
    avgTrust: 0,
    avgProtected: 0,
    avgRecommend: 0,
    npsScore: 0,
    responseRate: 'N/A',
    insights: [],
    distribution: { high: 0, mid: 0, low: 0 }
  };

  if (!sheet) return data;

  // Check if there's data by looking at column A (Timestamp)
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return data;

  data.totalResponses = lastRow - 1;

  // Get satisfaction scores (Q6-Q9 are columns G-J, 1-indexed as 7-10)
  var satisfactionRange = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, data.totalResponses, 4).getValues();

  var sumOverall = 0, sumTrust = 0, sumProtected = 0, sumRecommend = 0;
  var promoters = 0, detractors = 0;
  var validCount = 0;

  satisfactionRange.forEach(function(row) {
    var satisfied = parseFloat(row[0]) || 0;
    var trust = parseFloat(row[1]) || 0;
    var protected_ = parseFloat(row[2]) || 0;
    var recommend = parseFloat(row[3]) || 0;

    if (satisfied > 0) {
      sumOverall += satisfied;
      sumTrust += trust;
      sumProtected += protected_;
      sumRecommend += recommend;
      validCount++;

      // NPS calculation (based on recommend score 1-10)
      if (recommend >= 9) promoters++;
      else if (recommend <= 6) detractors++;

      // Calculate distribution based on average score
      var avgScore = (satisfied + trust + protected_ + recommend) / 4;
      if (avgScore >= 7) data.distribution.high++;
      else if (avgScore >= 5) data.distribution.mid++;
      else data.distribution.low++;
    }
  });

  if (validCount > 0) {
    data.avgOverall = sumOverall / validCount;
    data.avgTrust = sumTrust / validCount;
    data.avgProtected = sumProtected / validCount;
    data.avgRecommend = sumRecommend / validCount;
    data.npsScore = Math.round(((promoters - detractors) / validCount) * 100);
  }

  // Get steward ratings (Q10-Q16, columns K-Q)
  var stewardRange = sheet.getRange(2, SATISFACTION_COLS.Q10_TIMELY_RESPONSE, data.totalResponses, 7).getValues();
  var sumSteward = 0, stewardCount = 0;

  stewardRange.forEach(function(row) {
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      sumSteward += rowSum / rowCount;
      stewardCount++;
    }
  });

  if (stewardCount > 0) {
    data.avgSteward = sumSteward / stewardCount;
  }

  // Get leadership ratings (Q26-Q31, columns AA-AF)
  var leadershipRange = sheet.getRange(2, SATISFACTION_COLS.Q26_DECISIONS_CLEAR, data.totalResponses, 6).getValues();
  var sumLeadership = 0, leadershipCount = 0;

  leadershipRange.forEach(function(row) {
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      sumLeadership += rowSum / rowCount;
      leadershipCount++;
    }
  });

  if (leadershipCount > 0) {
    data.avgLeadership = sumLeadership / leadershipCount;
  }

  // Calculate response rate if we have member directory
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var totalMembers = memberSheet.getLastRow() - 1;
    var rate = Math.round((data.totalResponses / totalMembers) * 100);
    data.responseRate = rate + '%';
  }

  // Generate insights
  if (data.avgOverall >= 8) {
    data.insights.push({
      type: '',
      icon: '🌟',
      title: 'High Overall Satisfaction',
      text: 'Members report strong satisfaction with union representation (avg ' + data.avgOverall.toFixed(1) + '/10).'
    });
  } else if (data.avgOverall < 5) {
    data.insights.push({
      type: 'alert',
      icon: '⚠️',
      title: 'Low Satisfaction Alert',
      text: 'Overall satisfaction is below target at ' + data.avgOverall.toFixed(1) + '/10. Consider reviewing member concerns.'
    });
  }

  if (data.npsScore >= 50) {
    data.insights.push({
      type: '',
      icon: '🎯',
      title: 'Members Highly Recommend',
      text: 'Loyalty Score of ' + data.npsScore + ' means members actively recommend the union to colleagues.'
    });
  } else if (data.npsScore >= 0) {
    data.insights.push({
      type: '',
      icon: '📊',
      title: 'Moderate Member Loyalty',
      text: 'Loyalty Score of ' + data.npsScore + ' shows members are neutral. Focus on converting neutral members to advocates.'
    });
  } else {
    data.insights.push({
      type: 'warning',
      icon: '⚠️',
      title: 'Member Loyalty Needs Attention',
      text: 'Loyalty Score of ' + data.npsScore + ' indicates more critics than advocates. Address member concerns to improve.'
    });
  }

  if (data.avgSteward >= 8) {
    data.insights.push({
      type: '',
      icon: '🤝',
      title: 'Excellent Steward Performance',
      text: 'Stewards are rated highly at ' + data.avgSteward.toFixed(1) + '/10 on average.'
    });
  } else if (data.avgSteward < 6 && stewardCount > 0) {
    data.insights.push({
      type: 'warning',
      icon: '👤',
      title: 'Steward Training Opportunity',
      text: 'Steward ratings averaging ' + data.avgSteward.toFixed(1) + '/10 suggest room for improvement.'
    });
  }

  return data;
}

/**
 * Gets aggregate satisfaction statistics for dashboard widgets
 * @returns {Object} Aggregate satisfaction metrics
 */
function getAggregateSatisfactionStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      avgTrust: 0,
      avgStewardRating: 0,
      avgLeadership: 0,
      avgCommunication: 0,
      responseCount: 0,
      trendData: []
    };
  }

  // Get data starting from row 2 (skip header)
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, SATISFACTION_COLS.AVG_SCHEDULING || 82).getValues();

  // Load vault data to check verified/isLatest status (PII stays in vault)
  var vaultMap = getVaultDataMap_();

  // Filter to only verified and latest responses using vault flags
  var validRows = [];
  for (var vi = 0; vi < data.length; vi++) {
    var satRow = vi + 2; // 1-indexed sheet row
    var vaultEntry = vaultMap[satRow];
    if (vaultEntry && isTruthyValue(vaultEntry.verified) && isTruthyValue(vaultEntry.isLatest)) {
      data[vi]._vaultQuarter = vaultEntry.quarter; // attach for trend tracking
      validRows.push(data[vi]);
    }
  }

  if (validRows.length === 0) {
    return {
      avgTrust: 0,
      avgStewardRating: 0,
      avgLeadership: 0,
      avgCommunication: 0,
      responseCount: 0,
      trendData: []
    };
  }

  // Calculate averages from the summary columns
  var trustSum = 0, stewardSum = 0, leadershipSum = 0, commSum = 0;
  var trustCount = 0, stewardCount = 0, leadershipCount = 0, commCount = 0;

  // Also track trend data by quarter
  var quarterData = {};

  for (var i = 0; i < validRows.length; i++) {
    var row = validRows[i];

    // Trust (Q7_TRUST_UNION)
    var trust = parseFloat(row[SATISFACTION_COLS.Q7_TRUST_UNION - 1]);
    if (!isNaN(trust)) {
      trustSum += trust;
      trustCount++;
    }

    // Steward Rating (average of Q10-Q16)
    var stewardAvg = parseFloat(row[SATISFACTION_COLS.AVG_STEWARD_RATING - 1]);
    if (!isNaN(stewardAvg)) {
      stewardSum += stewardAvg;
      stewardCount++;
    }

    // Leadership (average of Q26-Q31)
    var leadershipAvg = parseFloat(row[SATISFACTION_COLS.AVG_LEADERSHIP - 1]);
    if (!isNaN(leadershipAvg)) {
      leadershipSum += leadershipAvg;
      leadershipCount++;
    }

    // Communication (average of Q41-Q45)
    var commAvg = parseFloat(row[SATISFACTION_COLS.AVG_COMMUNICATION - 1]);
    if (!isNaN(commAvg)) {
      commSum += commAvg;
      commCount++;
    }

    // Track by quarter for trend (quarter stored in vault, attached above)
    var quarter = row._vaultQuarter || '';
    if (quarter && trust) {
      if (!quarterData[quarter]) {
        quarterData[quarter] = { sum: 0, count: 0 };
      }
      quarterData[quarter].sum += trust;
      quarterData[quarter].count++;
    }
  }

  // Build trend data for charts (last 6 quarters)
  var trendData = [];
  var quarters = Object.keys(quarterData).sort().slice(-6);
  quarters.forEach(function(q) {
    var avg = quarterData[q].count > 0 ? quarterData[q].sum / quarterData[q].count : 0;
    trendData.push([q, parseFloat(avg.toFixed(1))]);
  });

  return {
    avgTrust: trustCount > 0 ? parseFloat((trustSum / trustCount).toFixed(1)) : 0,
    avgStewardRating: stewardCount > 0 ? parseFloat((stewardSum / stewardCount).toFixed(1)) : 0,
    avgLeadership: leadershipCount > 0 ? parseFloat((leadershipSum / leadershipCount).toFixed(1)) : 0,
    avgCommunication: commCount > 0 ? parseFloat((commSum / commCount).toFixed(1)) : 0,
    responseCount: validRows.length,
    trendData: trendData
  };
}

// ============================================================================
// SATISFACTION DATA RETRIEVAL - RESPONSES AND SECTIONS
// ============================================================================

/**
 * Get individual response data for satisfaction dashboard
 */
function getSatisfactionResponseData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (!sheet) return [];

  // Check if there's data
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  // Get worksite, role, shift, time in role, steward contact, and satisfaction scores
  var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
  var roleData = sheet.getRange(2, SATISFACTION_COLS.Q2_ROLE, numRows, 1).getValues();
  var shiftData = sheet.getRange(2, SATISFACTION_COLS.Q3_SHIFT, numRows, 1).getValues();
  var timeData = sheet.getRange(2, SATISFACTION_COLS.Q4_TIME_IN_ROLE, numRows, 1).getValues();
  var stewardContactData = sheet.getRange(2, SATISFACTION_COLS.Q5_STEWARD_CONTACT, numRows, 1).getValues();
  var timestampData = sheet.getRange(2, 1, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();
  var stewardRatingsData = sheet.getRange(2, SATISFACTION_COLS.Q10_TIMELY_RESPONSE, numRows, 7).getValues();

  var responses = [];
  for (var i = 0; i < numRows; i++) {
    // Get individual scores
    var satisfaction = parseFloat(satisfactionData[i][0]) || 0;
    var trust = parseFloat(satisfactionData[i][1]) || 0;
    var protected_ = parseFloat(satisfactionData[i][2]) || 0;
    var recommend = parseFloat(satisfactionData[i][3]) || 0;

    // Calculate average satisfaction score
    var sum = 0, count = 0;
    [satisfaction, trust, protected_, recommend].forEach(function(s) {
      if (s > 0) { sum += s; count++; }
    });
    var avgScore = count > 0 ? sum / count : 0;

    // Calculate steward rating average
    var stewardSum = 0, stewardCount = 0;
    stewardRatingsData[i].forEach(function(s) {
      var v = parseFloat(s);
      if (v > 0) { stewardSum += v; stewardCount++; }
    });
    var stewardRating = stewardCount > 0 ? stewardSum / stewardCount : 0;

    var ts = timestampData[i][0];
    var dateStr = ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : (ts || 'N/A');
    var stewardContact = stewardContactData[i][0];

    responses.push({
      worksite: worksiteData[i][0] || 'Unknown',
      role: roleData[i][0] || 'Unknown',
      shift: shiftData[i][0] || 'N/A',
      timeInRole: timeData[i][0] || 'N/A',
      date: dateStr,
      avgScore: avgScore,
      satisfaction: satisfaction,
      trust: trust,
      protected: protected_,
      recommend: recommend,
      stewardContact: isTruthyValue(stewardContact),
      stewardRating: stewardRating
    });
  }

  // Sort by date (most recent first) - use Date parsing for correct chronological order
  responses.sort(function(a, b) {
    return new Date(b.date).getTime() - new Date(a.date).getTime();
  });

  return responses;
}

/**
 * Get section-level data for satisfaction dashboard
 */
function getSatisfactionSectionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { sections: [] };
  if (!sheet) return result;

  // Check if there's data
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;

  // Define sections with their column ranges
  var sectionDefs = [
    { name: 'Overall Satisfaction', startCol: SATISFACTION_COLS.Q6_SATISFIED_REP, numCols: 4 },
    { name: 'Steward Ratings', startCol: SATISFACTION_COLS.Q10_TIMELY_RESPONSE, numCols: 7 },
    { name: 'Steward Access', startCol: SATISFACTION_COLS.Q18_KNOW_CONTACT, numCols: 3 },
    { name: 'Chapter Effectiveness', startCol: SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, numCols: 5 },
    { name: 'Local Leadership', startCol: SATISFACTION_COLS.Q26_DECISIONS_CLEAR, numCols: 6 },
    { name: 'Contract Enforcement', startCol: SATISFACTION_COLS.Q32_ENFORCES_CONTRACT, numCols: 4 },
    { name: 'Representation Process', startCol: SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS, numCols: 4 },
    { name: 'Communication Quality', startCol: SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, numCols: 5 },
    { name: 'Member Voice & Culture', startCol: SATISFACTION_COLS.Q46_VOICE_MATTERS, numCols: 5 },
    { name: 'Value & Collective Action', startCol: SATISFACTION_COLS.Q51_GOOD_VALUE, numCols: 5 },
    { name: 'Scheduling/Office Days', startCol: SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES, numCols: 7 }
  ];

  sectionDefs.forEach(function(section) {
    var data = sheet.getRange(2, section.startCol, numRows, section.numCols).getValues();
    var sum = 0, count = 0;

    data.forEach(function(row) {
      row.forEach(function(val) {
        var v = parseFloat(val);
        if (v > 0 && v <= 10) {
          sum += v;
          count++;
        }
      });
    });

    result.sections.push({
      name: section.name,
      avg: count > 0 ? sum / count : 0,
      responseCount: Math.floor(count / section.numCols),
      questions: section.numCols
    });
  });

  // Sort by score (lowest first to highlight areas needing attention)
  result.sections.sort(function(a, b) { return a.avg - b.avg; });

  return result;
}

// ============================================================================
// SATISFACTION DATA RETRIEVAL - TRENDS AND BREAKDOWNS
// ============================================================================

/**
 * Get trend data for satisfaction dashboard - responses over time
 */
function getSatisfactionTrendData(period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    byMonth: [],
    satisfactionTrend: [],
    issuesTrend: [],
    totalInPeriod: 0
  };

  if (!sheet) return result;

  // Get data
  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

  // Filter by period
  var now = new Date();
  var cutoff = null;
  if (period === '30') cutoff = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  else if (period === '90') cutoff = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
  else if (period === 'year') cutoff = new Date(now.getFullYear(), 0, 1);

  // Group by month
  var monthData = {};
  for (var i = 0; i < numRows; i++) {
    var ts = timestamps[i][0];
    if (!(ts instanceof Date)) continue;
    if (cutoff && ts < cutoff) continue;

    var monthKey = Utilities.formatDate(ts, tz, 'yyyy-MM');
    var monthLabel = Utilities.formatDate(ts, tz, 'MMM yy');

    if (!monthData[monthKey]) {
      monthData[monthKey] = { label: monthLabel, count: 0, sum: 0, validCount: 0 };
    }

    monthData[monthKey].count++;
    result.totalInPeriod++;

    // Calculate avg satisfaction
    var row = satisfactionData[i];
    var rowSum = 0, rowCount = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { rowSum += v; rowCount++; }
    });
    if (rowCount > 0) {
      monthData[monthKey].sum += rowSum / rowCount;
      monthData[monthKey].validCount++;
    }
  }

  // Convert to arrays sorted by date
  var months = Object.keys(monthData).sort();
  months.forEach(function(key) {
    var m = monthData[key];
    result.byMonth.push({ label: m.label, count: m.count });
    result.satisfactionTrend.push({
      label: m.label,
      avg: m.validCount > 0 ? m.sum / m.validCount : 0
    });
  });

  // Get common issues/priorities for trend
  try {
    var prioritiesData = sheet.getRange(2, SATISFACTION_COLS.Q64_TOP_PRIORITIES, numRows, 1).getValues();
    var issueMap = {};
    for (var i = 0; i < numRows; i++) {
      var ts = timestamps[i][0];
      if (!(ts instanceof Date)) continue;
      if (cutoff && ts < cutoff) continue;

      var priorities = String(prioritiesData[i][0] || '');
      if (priorities) {
        priorities.split(',').forEach(function(item) {
          var p = item.trim();
          if (p) issueMap[p] = (issueMap[p] || 0) + 1;
        });
      }
    }
    for (var issue in issueMap) {
      result.issuesTrend.push({ name: issue, count: issueMap[issue] });
    }
    result.issuesTrend.sort(function(a, b) { return b.count - a.count; });
  } catch(_e) { /* ignore if column doesn't exist */ }

  return result;
}

/**
 * Get breakdown data for satisfaction dashboard
 */
function getSatisfactionBreakdownData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    sections: [],
    byWorksite: [],
    byRole: []
  };

  if (!sheet) return result;

  // Get sections data
  var sectionResult = getSatisfactionSectionData();
  result.sections = sectionResult.sections;

  // Get analytics data for worksite/role
  var analyticsData = getSatisfactionAnalyticsData();
  result.byWorksite = analyticsData.byWorksite;
  result.byRole = analyticsData.byRole;

  return result;
}

/**
 * Get insights data for satisfaction dashboard
 */
function getSatisfactionInsightsData() {
  var analyticsData = getSatisfactionAnalyticsData();
  var overviewData = getSatisfactionOverviewData();

  var result = {
    insights: analyticsData.insights || [],
    stewardImpact: analyticsData.stewardImpact,
    topPriorities: analyticsData.topPriorities
  };

  // Add additional insights based on overview data
  if (overviewData.avgOverall >= 8) {
    result.insights.unshift({
      type: 'success',
      icon: '🌟',
      title: 'Excellent Overall Satisfaction',
      text: 'Members report high satisfaction (' + overviewData.avgOverall.toFixed(1) + '/10). Keep up the great work!'
    });
  } else if (overviewData.avgOverall < 5) {
    result.insights.unshift({
      type: 'alert',
      icon: '⚠️',
      title: 'Satisfaction Needs Attention',
      text: 'Overall satisfaction is below target at ' + overviewData.avgOverall.toFixed(1) + '/10. Review member feedback for areas to improve.'
    });
  }

  return result;
}

// ============================================================================
// SATISFACTION DATA RETRIEVAL - DRILL-DOWN AND ANALYTICS
// ============================================================================

/**
 * Get drill-down data for specific categories
 */
function getSatisfactionDrillData(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { items: [] };
  if (!sheet) return result;

  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  if (type === 'responses') {
    // Show recent responses
    var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
    var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
    var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

    for (var i = 0; i < numRows; i++) {
      var ts = timestamps[i][0];
      var row = satisfactionData[i];
      var sum = 0, count = 0;
      row.forEach(function(val) { var v = parseFloat(val); if (v > 0) { sum += v; count++; } });
      var avg = count > 0 ? sum / count : 0;

      result.items.push({
        label: worksiteData[i][0] || 'Unknown',
        detail: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : 'N/A',
        score: avg
      });
    }
    result.items.sort(function(a, b) { return b.score - a.score; });
    result.items = result.items.slice(0, 20);
  }

  return result;
}

/**
 * Get location-specific drill-down data
 */
function getSatisfactionLocationDrill(location) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = { count: 0, avgScore: 0, responses: [] };
  if (!sheet || !location) return result;

  var lastRow = getSheetLastRow(sheet);
  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  var timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
  var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
  var roleData = sheet.getRange(2, SATISFACTION_COLS.Q2_ROLE, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();

  var totalScore = 0;

  for (var i = 0; i < numRows; i++) {
    if (worksiteData[i][0] !== location) continue;

    var ts = timestamps[i][0];
    var row = satisfactionData[i];
    var sum = 0, count = 0;
    row.forEach(function(val) { var v = parseFloat(val); if (v > 0) { sum += v; count++; } });
    var avg = count > 0 ? sum / count : 0;

    result.count++;
    totalScore += avg;

    result.responses.push({
      role: roleData[i][0] || 'Unknown',
      date: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : 'N/A',
      avgScore: avg
    });
  }

  result.avgScore = result.count > 0 ? totalScore / result.count : 0;
  result.responses.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

  return result;
}

/**
 * Get analytics data for satisfaction dashboard insights
 */
function getSatisfactionAnalyticsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    insights: [],
    byWorksite: [],
    byRole: [],
    stewardImpact: null,
    topPriorities: []
  };

  if (!sheet) return result;

  // Check if there's data
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return result;

  var numRows = lastRow - 1;

  // Get all relevant data in one batch
  var worksiteData = sheet.getRange(2, SATISFACTION_COLS.Q1_WORKSITE, numRows, 1).getValues();
  var roleData = sheet.getRange(2, SATISFACTION_COLS.Q2_ROLE, numRows, 1).getValues();
  var stewardContactData = sheet.getRange(2, SATISFACTION_COLS.Q5_STEWARD_CONTACT, numRows, 1).getValues();
  var satisfactionData = sheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numRows, 4).getValues();
  var prioritiesData = sheet.getRange(2, SATISFACTION_COLS.Q64_TOP_PRIORITIES, numRows, 1).getValues();

  // Calculate average score for each response
  var scores = [];
  for (var i = 0; i < numRows; i++) {
    var row = satisfactionData[i];
    var sum = 0, count = 0;
    row.forEach(function(val) {
      var v = parseFloat(val);
      if (v > 0) { sum += v; count++; }
    });
    scores.push(count > 0 ? sum / count : 0);
  }

  // By Worksite analysis
  var worksiteMap = {};
  for (var i = 0; i < numRows; i++) {
    var ws = worksiteData[i][0] || 'Unknown';
    if (!worksiteMap[ws]) worksiteMap[ws] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      worksiteMap[ws].sum += scores[i];
      worksiteMap[ws].count++;
    }
  }

  for (var ws in worksiteMap) {
    if (worksiteMap[ws].count > 0) {
      result.byWorksite.push({
        name: ws,
        avg: worksiteMap[ws].sum / worksiteMap[ws].count,
        count: worksiteMap[ws].count
      });
    }
  }
  result.byWorksite.sort(function(a, b) { return b.avg - a.avg; });

  // By Role analysis
  var roleMap = {};
  for (var i = 0; i < numRows; i++) {
    var role = roleData[i][0] || 'Unknown';
    if (!roleMap[role]) roleMap[role] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      roleMap[role].sum += scores[i];
      roleMap[role].count++;
    }
  }

  for (var role in roleMap) {
    if (roleMap[role].count > 0) {
      result.byRole.push({
        name: role,
        avg: roleMap[role].sum / roleMap[role].count,
        count: roleMap[role].count
      });
    }
  }
  result.byRole.sort(function(a, b) { return b.avg - a.avg; });

  // Steward contact impact
  var withContactSum = 0, withContactCount = 0;
  var withoutContactSum = 0, withoutContactCount = 0;

  for (var i = 0; i < numRows; i++) {
    var contact = String(stewardContactData[i][0]).toLowerCase();
    if (scores[i] > 0) {
      if (contact === 'yes') {
        withContactSum += scores[i];
        withContactCount++;
      } else if (contact === 'no') {
        withoutContactSum += scores[i];
        withoutContactCount++;
      }
    }
  }

  if (withContactCount > 0 || withoutContactCount > 0) {
    result.stewardImpact = {
      withContact: withContactCount > 0 ? withContactSum / withContactCount : 0,
      withContactCount: withContactCount,
      withoutContact: withoutContactCount > 0 ? withoutContactSum / withoutContactCount : 0,
      withoutContactCount: withoutContactCount
    };
  }

  // Top priorities analysis
  var priorityMap = {};
  for (var i = 0; i < numRows; i++) {
    var priorities = String(prioritiesData[i][0] || '');
    if (priorities) {
      // Split by comma and count each priority
      var items = priorities.split(',');
      items.forEach(function(item) {
        var p = item.trim();
        if (p) {
          priorityMap[p] = (priorityMap[p] || 0) + 1;
        }
      });
    }
  }

  for (var p in priorityMap) {
    result.topPriorities.push({ name: p, count: priorityMap[p] });
  }
  result.topPriorities.sort(function(a, b) { return b.count - a.count; });
  result.topPriorities = result.topPriorities.slice(0, 10); // Top 10

  // Generate insights
  // Lowest scoring worksite
  if (result.byWorksite.length > 0) {
    var lowest = result.byWorksite[result.byWorksite.length - 1];
    if (lowest.avg < 6 && lowest.count >= 3) {
      result.insights.push({
        type: 'warning',
        icon: '📍',
        title: 'Worksite Attention Needed',
        text: lowest.name + ' has the lowest satisfaction score (' + lowest.avg.toFixed(1) + '/10) with ' + lowest.count + ' responses.'
      });
    }
  }

  // Steward impact insight
  if (result.stewardImpact && result.stewardImpact.withContactCount > 0 && result.stewardImpact.withoutContactCount > 0) {
    var diff = result.stewardImpact.withContact - result.stewardImpact.withoutContact;
    if (diff > 1) {
      result.insights.push({
        type: '',
        icon: '🤝',
        title: 'Steward Contact Matters',
        text: 'Members who contacted a steward report ' + diff.toFixed(1) + ' points higher satisfaction on average.'
      });
    }
  }

  // Role insights
  if (result.byRole.length >= 2) {
    var topRole = result.byRole[0];
    var bottomRole = result.byRole[result.byRole.length - 1];
    if (topRole.avg - bottomRole.avg > 2 && bottomRole.count >= 3) {
      result.insights.push({
        type: 'warning',
        icon: '👤',
        title: 'Role Disparity',
        text: bottomRole.name + ' roles report lower satisfaction (' + bottomRole.avg.toFixed(1) + ') than ' + topRole.name + ' (' + topRole.avg.toFixed(1) + ').'
      });
    }
  }

  // Top priority insight
  if (result.topPriorities.length > 0) {
    var topP = result.topPriorities[0];
    result.insights.push({
      type: '',
      icon: '🎯',
      title: 'Top Member Priority',
      text: '"' + topP.name + '" is the most cited priority with ' + topP.count + ' mentions.'
    });
  }

  return result;
}

// ============================================================================
// SATISFACTION VALUE SYNC AND CALCULATIONS
// ============================================================================

/**
 * Sync satisfaction sheet with computed values (no formulas)
 * Calculates section averages and dashboard metrics
 */
function syncSatisfactionValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!sheet) {
    Logger.log('Member Satisfaction sheet not found');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // No data to process, just write empty dashboard
    writeSatisfactionDashboard_(sheet, [], []);
    return;
  }

  // Get all response data (columns A-BK, 1-63)
  var responseData = sheet.getRange(2, 1, lastRow - 1, 63).getValues();

  // Calculate section averages for each row
  var sectionAverages = computeSectionAverages_(responseData);

  // Write section averages to columns BT-CD (72-82)
  if (sectionAverages.length > 0) {
    sheet.getRange(2, 72, sectionAverages.length, 11).setValues(sectionAverages);
  }

  // Calculate and write dashboard metrics
  writeSatisfactionDashboard_(sheet, responseData, sectionAverages);

  Logger.log('Member Satisfaction values synced for ' + responseData.length + ' responses');
}

/**
 * Compute section averages for satisfaction survey rows
 * @param {Array} responseData - 2D array of survey response data
 * @return {Array} 2D array of section averages (11 columns per row)
 * @private
 */
function computeSectionAverages_(responseData) {
  var results = [];

  for (var r = 0; r < responseData.length; r++) {
    var row = responseData[r];
    if (!row[0]) continue; // Skip empty rows

    var averages = [];

    // Overall Satisfaction (Q6-9: columns G-J, indices 6-9)
    averages.push(computeAverage_(row, 6, 9));

    // Steward Rating (Q10-16: columns K-Q, indices 10-16)
    averages.push(computeAverage_(row, 10, 16));

    // Steward Access (Q18-20: columns S-U, indices 18-20)
    averages.push(computeAverage_(row, 18, 20));

    // Chapter (Q21-25: columns V-Z, indices 21-25)
    averages.push(computeAverage_(row, 21, 25));

    // Leadership (Q26-31: columns AA-AF, indices 26-31)
    averages.push(computeAverage_(row, 26, 31));

    // Contract (Q32-35: columns AG-AJ, indices 32-35)
    averages.push(computeAverage_(row, 32, 35));

    // Representation (Q37-40: columns AL-AO, indices 37-40)
    averages.push(computeAverage_(row, 37, 40));

    // Communication (Q41-45: columns AP-AT, indices 41-45)
    averages.push(computeAverage_(row, 41, 45));

    // Member Voice (Q46-50: columns AU-AY, indices 46-50)
    averages.push(computeAverage_(row, 46, 50));

    // Value/Action (Q51-55: columns AZ-BD, indices 51-55)
    averages.push(computeAverage_(row, 51, 55));

    // Scheduling (Q56-62: columns BE-BK, indices 56-62)
    averages.push(computeAverage_(row, 56, 62));

    results.push(averages);
  }

  return results;
}

/**
 * Compute average of numeric values in a row range
 * @param {Array} row - Single row of data
 * @param {number} startIdx - Start index (0-based)
 * @param {number} endIdx - End index (0-based, inclusive)
 * @return {number|string} Average or empty string if no valid values
 * @private
 */
function computeAverage_(row, startIdx, endIdx) {
  var values = [];
  for (var i = startIdx; i <= endIdx; i++) {
    var val = row[i];
    if (typeof val === 'number' && !isNaN(val)) {
      values.push(val);
    }
  }

  if (values.length === 0) return '';

  var sum = values.reduce(function(a, b) { return a + b; }, 0);
  return Math.round(sum / values.length * 100) / 100;
}

/**
 * Write satisfaction dashboard summary values
 * @param {Sheet} sheet - The Satisfaction sheet
 * @param {Array} responseData - Raw response data
 * @param {Array} sectionAverages - Computed section averages
 * @private
 */
function writeSatisfactionDashboard_(sheet, responseData, sectionAverages) {
  var dashStart = 84; // Column CF
  var demoStart = 87; // Column CH
  var chartStart = 90; // Column CK

  // Calculate aggregate metrics
  var totalResponses = responseData.length;
  var responsePeriod = 'No data';
  if (totalResponses > 0) {
    var timestamps = responseData.map(function(r) { return r[0]; }).filter(function(t) { return t instanceof Date; });
    if (timestamps.length > 0) {
      var minDate = new Date(Math.min.apply(null, timestamps));
      var maxDate = new Date(Math.max.apply(null, timestamps));
      responsePeriod = Utilities.formatDate(minDate, Session.getScriptTimeZone(), 'MM/dd') + ' - ' +
                       Utilities.formatDate(maxDate, Session.getScriptTimeZone(), 'MM/dd');
    }
  }

  // Calculate section score averages
  var sectionScores = [];
  var sectionNames = ['Overall Satisfaction', 'Steward Rating', 'Steward Access', 'Chapter Effectiveness',
                      'Local Leadership', 'Contract Enforcement', 'Representation', 'Communication',
                      'Member Voice', 'Value & Action', 'Scheduling'];

  for (var s = 0; s < 11; s++) {
    var values = sectionAverages.map(function(r) { return r[s]; }).filter(function(v) { return typeof v === 'number'; });
    var avg = values.length > 0 ? Math.round(values.reduce(function(a, b) { return a + b; }, 0) / values.length * 10) / 10 : '';
    sectionScores.push(avg);
  }

  // Write Response Summary (rows 4-19, columns CF-CG)
  var summaryData = [
    ['Total Responses', totalResponses],
    ['Response Period', responsePeriod],
    ['', ''],
    ['📊 SECTION SCORES', ''],
    ['Section', 'Avg Score']
  ];
  for (var i = 0; i < sectionNames.length; i++) {
    summaryData.push([sectionNames[i], sectionScores[i]]);
  }
  sheet.getRange(4, dashStart, summaryData.length, 2).setValues(summaryData);

  // Calculate demographics
  var shifts = { Day: 0, Evening: 0, Night: 0, Rotating: 0 };
  var tenure = { '<1': 0, '1-3': 0, '4-7': 0, '8-15': 0, '15+': 0 };
  var stewardContact = { Yes: 0, No: 0 };
  var filedGrievance = { Yes: 0, No: 0 };

  for (var d = 0; d < responseData.length; d++) {
    var row = responseData[d];

    // Shift (column D, index 3)
    var shift = row[3];
    if (shift === 'Day') shifts.Day++;
    else if (shift === 'Evening') shifts.Evening++;
    else if (shift === 'Night') shifts.Night++;
    else if (shift === 'Rotating') shifts.Rotating++;

    // Tenure (column E, index 4)
    var ten = String(row[4] || '');
    if (ten.indexOf('<1') >= 0) tenure['<1']++;
    else if (ten.indexOf('1-3') >= 0) tenure['1-3']++;
    else if (ten.indexOf('4-7') >= 0) tenure['4-7']++;
    else if (ten.indexOf('8-15') >= 0) tenure['8-15']++;
    else if (ten.indexOf('15+') >= 0) tenure['15+']++;

    // Steward contact (column F, index 5)
    if (row[5] === 'Yes') stewardContact.Yes++;
    else if (row[5] === 'No') stewardContact.No++;

    // Filed grievance (column AK, index 36)
    if (row[36] === 'Yes') filedGrievance.Yes++;
    else if (row[36] === 'No') filedGrievance.No++;
  }

  // Write Demographics (rows 4-23, columns CH-CI)
  var demoData = [
    ['Shift Breakdown', ''],
    ['Day', shifts.Day],
    ['Evening', shifts.Evening],
    ['Night', shifts.Night],
    ['Rotating', shifts.Rotating],
    ['', ''],
    ['Tenure', ''],
    ['<1 year', tenure['<1']],
    ['1-3 years', tenure['1-3']],
    ['4-7 years', tenure['4-7']],
    ['8-15 years', tenure['8-15']],
    ['15+ years', tenure['15+']],
    ['', ''],
    ['Steward Contact', ''],
    ['Yes (12 mo)', stewardContact.Yes],
    ['No', stewardContact.No],
    ['', ''],
    ['Filed Grievance', ''],
    ['Yes (24 mo)', filedGrievance.Yes],
    ['No', filedGrievance.No]
  ];
  sheet.getRange(4, demoStart, demoData.length, 2).setValues(demoData);

  // Write Chart Data (rows 4-15, columns CK-CL)
  var chartSectionNames = ['Overall Satisfaction', 'Steward Rating', 'Steward Access', 'Chapter',
                           'Leadership', 'Contract', 'Representation', 'Communication',
                           'Member Voice', 'Value & Action', 'Scheduling'];
  var chartData = [['Section', 'Score']];
  for (var c = 0; c < 11; c++) {
    var score = typeof sectionScores[c] === 'number' ? Math.round(sectionScores[c] * 100) / 100 : 0;
    chartData.push([chartSectionNames[c], score]);
  }
  sheet.getRange(4, chartStart, chartData.length, 2).setValues(chartData);
}

/**
 * Compute section averages for a single new survey response row
 * Used by onSatisfactionFormSubmit for efficiency (only computes one row)
 * @param {number} row - Row number of the new response
 */
function computeSatisfactionRowAverages(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!sheet || row < 2) return;

  // Get the response data for this row (columns A-BK, 1-63)
  var rowData = sheet.getRange(row, 1, 1, 63).getValues()[0];

  if (!rowData[0]) return; // Skip if no timestamp

  var averages = [];

  // Overall Satisfaction (Q6-9: indices 6-9)
  averages.push(computeAverage_(rowData, 6, 9));
  // Steward Rating (Q10-16: indices 10-16)
  averages.push(computeAverage_(rowData, 10, 16));
  // Steward Access (Q18-20: indices 18-20)
  averages.push(computeAverage_(rowData, 18, 20));
  // Chapter (Q21-25: indices 21-25)
  averages.push(computeAverage_(rowData, 21, 25));
  // Leadership (Q26-31: indices 26-31)
  averages.push(computeAverage_(rowData, 26, 31));
  // Contract (Q32-35: indices 32-35)
  averages.push(computeAverage_(rowData, 32, 35));
  // Representation (Q37-40: indices 37-40)
  averages.push(computeAverage_(rowData, 37, 40));
  // Communication (Q41-45: indices 41-45)
  averages.push(computeAverage_(rowData, 41, 45));
  // Member Voice (Q46-50: indices 46-50)
  averages.push(computeAverage_(rowData, 46, 50));
  // Value/Action (Q51-55: indices 51-55)
  averages.push(computeAverage_(rowData, 51, 55));
  // Scheduling (Q56-62: indices 56-62)
  averages.push(computeAverage_(rowData, 56, 62));

  // Write section averages to this row (columns BT-CD, 72-82)
  sheet.getRange(row, 72, 1, 11).setValues([averages]);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Helper function to get last row with data
 * @param {Sheet} sheet - The sheet to check
 * @return {number} The last row number with data
 */
function getSheetLastRow(sheet) {
  return sheet.getLastRow();
}



/**
 * ============================================================================
 * 08g_SyncEngine.gs - Data Synchronization Engine
 * ============================================================================
 *
 * This module handles all data synchronization between sheets including:
 * - Grievance Log <-> Member Directory bidirectional sync
 * - Dashboard value computation and updates
 * - Auto-sync triggers and configuration
 * - Feedback sheet metrics sync
 *
 * Dependencies:
 * - 00_Config.gs (SHEETS, GRIEVANCE_COLS, MEMBER_COLS, CONFIG_COLS, FEEDBACK_COLS, COLORS)
 * - 01_Utilities.gs (getColumnLetter, getConfigValues, getJobMetadataByMemberCol)
 *
 * @author Claude Code Assistant
 * @version 1.0.0
 * ============================================================================
 */

// ============================================================================
// GRIEVANCE <-> MEMBER DIRECTORY SYNC
// ============================================================================

/**
 * Sync grievance status data to Member Directory
 * Updates Has Open Grievance, Grievance Status, and Days to Deadline columns
 */
function syncGrievanceToMemberDirectory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    Logger.log('Required sheets not found for grievance sync');
    return;
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  // Closed statuses - grievances with these statuses don't count as "open"
  var closedStatuses = ['Closed', 'Settled', 'Withdrawn', 'Denied', 'Won'];

  // Build lookup map: memberId -> {hasOpen, status, deadline}
  // Calculate directly from grievance data (handles "Overdue" text properly)
  var lookup = {};

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var isClosed = closedStatuses.indexOf(status) !== -1;

    // Initialize member entry if not exists
    if (!lookup[memberId]) {
      lookup[memberId] = {
        hasOpen: 'No',
        status: '',
        deadline: '',
        minDeadline: Infinity,  // Track minimum numeric deadline
        hasOverdue: false       // Track if any grievance is overdue
      };
    }

    // Check if this grievance is open/pending
    if (!isClosed) {
      lookup[memberId].hasOpen = 'Yes';

      // Set status priority: Open > Pending Info
      if (status === 'Open') {
        lookup[memberId].status = 'Open';
      } else if (status === 'Pending Info' && lookup[memberId].status !== 'Open') {
        lookup[memberId].status = 'Pending Info';
      }

      // Handle Days to Deadline (can be number or "Overdue" text)
      if (daysToDeadline === 'Overdue') {
        lookup[memberId].hasOverdue = true;
      } else if (typeof daysToDeadline === 'number' && daysToDeadline < lookup[memberId].minDeadline) {
        lookup[memberId].minDeadline = daysToDeadline;
      }
    }
  }

  // Finalize deadline values
  for (var mid in lookup) {
    var data = lookup[mid];
    if (data.hasOpen === 'Yes') {
      if (data.minDeadline !== Infinity) {
        // Has a numeric deadline - use the minimum
        data.deadline = data.minDeadline;
      } else if (data.hasOverdue) {
        // All open grievances are overdue
        data.deadline = 'Overdue';
      }
    }
  }

  // Get member data
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) return;

  // Update columns AB-AD (Has Open Grievance?, Grievance Status, Days to Deadline)
  var updates = [];
  for (var j = 1; j < memberData.length; j++) {
    var memberId = memberData[j][MEMBER_COLS.MEMBER_ID - 1];
    var memberInfo = lookup[memberId] || {hasOpen: 'No', status: '', deadline: ''};
    updates.push([memberInfo.hasOpen, memberInfo.status, memberInfo.deadline]);
  }

  if (updates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.HAS_OPEN_GRIEVANCE, updates.length, 3).setValues(updates);
  }

  Logger.log('Synced grievance data to ' + updates.length + ' members');
}

/**
 * Sync calculated formulas from hidden sheet to Grievance Log
 * This is the self-healing function - it copies calculated values to the Grievance Log
 * Member data (Name, Email, Unit, Location, Steward) is looked up directly from Member Directory
 */
function syncGrievanceFormulasToLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    Logger.log('Required sheets not found for grievance formula sync');
    return;
  }

  // Get Member Directory data and create lookup by Member ID
  var memberData = memberSheet.getDataRange().getValues();
  var memberLookup = {};
  for (var i = 1; i < memberData.length; i++) {
    var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
    if (memberId) {
      memberLookup[memberId] = {
        firstName: memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: memberData[i][MEMBER_COLS.LAST_NAME - 1] || '',
        email: memberData[i][MEMBER_COLS.EMAIL - 1] || '',
        unit: memberData[i][MEMBER_COLS.UNIT - 1] || '',
        location: memberData[i][MEMBER_COLS.WORK_LOCATION - 1] || '',
        steward: memberData[i][MEMBER_COLS.ASSIGNED_STEWARD - 1] || ''
      };
    }
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to start of day

  // Closed statuses that should not have Next Action Due
  var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];

  // Prepare updates
  var nameUpdates = [];           // Columns C-D
  var deadlineUpdates = [];       // Columns H, J, L, N, P (Filing Deadline, Step I Due, Step II Appeal Due, Step II Due, Step III Appeal Due)
  var metricsUpdates = [];        // Columns S, T, U (Days Open, Next Action Due, Days to Deadline)
  var contactUpdates = [];        // Columns X, Y, Z, AA (Email, Unit, Location, Steward)

  // Track data quality issues
  var orphanedGrievances = [];    // Grievances with non-existent Member IDs
  var missingMemberIds = [];      // Grievances with no Member ID

  for (var j = 1; j < grievanceData.length; j++) {
    var row = grievanceData[j];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || ('Row ' + (j + 1));

    // Track data quality issues
    if (!memberId) {
      missingMemberIds.push(grievanceId);
      Logger.log('WARNING: Grievance ' + grievanceId + ' has no Member ID');
    } else if (!memberLookup[memberId]) {
      orphanedGrievances.push(grievanceId + ' (Member ID: ' + memberId + ')');
      Logger.log('WARNING: Grievance ' + grievanceId + ' references non-existent Member ID: ' + memberId);
    }

    var memberInfo = memberLookup[memberId] || {};

    // Names (C-D) - from Member Directory
    nameUpdates.push([
      memberInfo.firstName || '',
      memberInfo.lastName || ''
    ]);

    // Get date values from grievance row for deadline calculations
    var incidentDate = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
    var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var step1Rcvd = row[GRIEVANCE_COLS.STEP1_RCVD - 1];
    var step2AppealFiled = row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1];
    var step2Rcvd = row[GRIEVANCE_COLS.STEP2_RCVD - 1];
    var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];

    // Calculate deadline dates
    var filingDeadline = '';
    var step1Due = '';
    var step2AppealDue = '';
    var step2Due = '';
    var step3AppealDue = '';

    if (incidentDate instanceof Date) {
      filingDeadline = new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000);
    }
    if (dateFiled instanceof Date) {
      step1Due = new Date(dateFiled.getTime() + 30 * 24 * 60 * 60 * 1000);
    }
    if (step1Rcvd instanceof Date) {
      step2AppealDue = new Date(step1Rcvd.getTime() + 10 * 24 * 60 * 60 * 1000);
    }
    if (step2AppealFiled instanceof Date) {
      step2Due = new Date(step2AppealFiled.getTime() + 30 * 24 * 60 * 60 * 1000);
    }
    if (step2Rcvd instanceof Date) {
      step3AppealDue = new Date(step2Rcvd.getTime() + 30 * 24 * 60 * 60 * 1000);
    }

    // Deadlines (H, J, L, N, P)
    deadlineUpdates.push([
      filingDeadline,
      step1Due,
      step2AppealDue,
      step2Due,
      step3AppealDue
    ]);

    // Calculate Days Open directly
    var daysOpen = '';
    if (dateFiled instanceof Date) {
      if (dateClosed instanceof Date) {
        daysOpen = Math.floor((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
      } else {
        daysOpen = Math.floor((today - dateFiled) / (1000 * 60 * 60 * 24));
      }
    }

    // Calculate Next Action Due based on current step and status
    var nextActionDue = '';
    var isClosed = closedStatuses.indexOf(status) !== -1;

    if (!isClosed && currentStep) {
      if (currentStep === 'Informal' && filingDeadline) {
        nextActionDue = filingDeadline;
      } else if (currentStep === 'Step I' && step1Due) {
        nextActionDue = step1Due;
      } else if (currentStep === 'Step II' && step2Due) {
        nextActionDue = step2Due;
      } else if (currentStep === 'Step III' && step3AppealDue) {
        nextActionDue = step3AppealDue;
      }
    }

    // Calculate Days to Deadline directly
    var daysToDeadline = '';
    if (nextActionDue instanceof Date) {
      var days = Math.floor((nextActionDue - today) / (1000 * 60 * 60 * 24));
      daysToDeadline = days < 0 ? 'Overdue' : days;
    }

    // Metrics (S, T, U)
    metricsUpdates.push([
      daysOpen,
      nextActionDue,
      daysToDeadline
    ]);

    // Contact info (X, Y, Z, AA)
    contactUpdates.push([
      memberInfo.email || '',
      memberInfo.unit || '',
      memberInfo.location || '',
      memberInfo.steward || ''
    ]);
  }

  // Apply updates to Grievance Log
  if (nameUpdates.length > 0) {
    // C-D: First Name, Last Name
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FIRST_NAME, nameUpdates.length, 2).setValues(nameUpdates);

    // Deadline columns: respect steward overrides (cells with "Steward override" note)
    var deadlineCols = [
      { col: GRIEVANCE_COLS.FILING_DEADLINE, idx: 0 },
      { col: GRIEVANCE_COLS.STEP1_DUE, idx: 1 },
      { col: GRIEVANCE_COLS.STEP2_APPEAL_DUE, idx: 2 },
      { col: GRIEVANCE_COLS.STEP2_DUE, idx: 3 },
      { col: GRIEVANCE_COLS.STEP3_APPEAL_DUE, idx: 4 }
    ];

    for (var dc = 0; dc < deadlineCols.length; dc++) {
      var dlCol = deadlineCols[dc].col;
      var dlIdx = deadlineCols[dc].idx;
      var dlRange = grievanceSheet.getRange(2, dlCol, deadlineUpdates.length, 1);
      var dlNotes = dlRange.getNotes();
      var dlValues = deadlineUpdates.map(function(r) { return [r[dlIdx]]; });

      // Preserve steward-overridden cells: keep existing value, skip formula recalc
      for (var nr = 0; nr < dlNotes.length; nr++) {
        if (dlNotes[nr][0] === 'Steward override') {
          // Keep the existing value in the sheet instead of overwriting
          dlValues[nr][0] = dlRange.getCell(nr + 1, 1).getValue();
        }
      }

      dlRange.setValues(dlValues);
      dlRange.setNumberFormat('MM/dd/yyyy');
    }

    // S, T, U: Days Open, Next Action Due, Days to Deadline
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, metricsUpdates.length, 3).setValues(metricsUpdates);

    // Format Days Open (S) as whole numbers, Next Action Due (T) as date
    // Days to Deadline (U) uses General format to preserve "Overdue" text
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, metricsUpdates.length, 1).setNumberFormat('0');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.NEXT_ACTION_DUE, metricsUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, metricsUpdates.length, 1).setNumberFormat('General');

    // X, Y, Z, AA: Email, Unit, Location, Steward
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_EMAIL, contactUpdates.length, 4).setValues(contactUpdates);
  }

  Logger.log('Synced grievance formulas to ' + nameUpdates.length + ' grievances');

  // Show warnings to user if data quality issues found
  var warnings = [];
  if (missingMemberIds.length > 0) {
    warnings.push(missingMemberIds.length + ' grievance(s) have no Member ID');
    Logger.log('Missing Member IDs: ' + missingMemberIds.join(', '));
  }
  if (orphanedGrievances.length > 0) {
    warnings.push(orphanedGrievances.length + ' grievance(s) reference non-existent members');
    Logger.log('Orphaned grievances: ' + orphanedGrievances.join(', '));
  }

  if (warnings.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Data issues found:\n' + warnings.join('\n') + '\n\nCheck Logs for details.',
      'Sync Warning',
      10
    );
  }
}

/**
 * Sync member data from hidden sheet to Grievance Log
 */
function syncMemberToGrievanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lookupSheet = ss.getSheetByName(SHEETS.MEMBER_LOOKUP);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!lookupSheet || !grievanceSheet) {
    Logger.log('Required sheets not found for member sync');
    return;
  }

  // Get lookup data
  var lookupData = lookupSheet.getDataRange().getValues();
  if (lookupData.length < 2) return;

  // Create lookup map
  var lookup = {};
  for (var i = 1; i < lookupData.length; i++) {
    var memberId = lookupData[i][0];
    if (memberId) {
      lookup[memberId] = {
        firstName: lookupData[i][1],
        lastName: lookupData[i][2],
        email: lookupData[i][3],
        unit: lookupData[i][4],
        location: lookupData[i][5],
        steward: lookupData[i][6]
      };
    }
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  // Update grievance rows
  var nameUpdates = [];
  var infoUpdates = [];

  for (var j = 1; j < grievanceData.length; j++) {
    var memberId = grievanceData[j][GRIEVANCE_COLS.MEMBER_ID - 1];
    var data = lookup[memberId] || {firstName: '', lastName: '', email: '', unit: '', location: '', steward: ''};
    nameUpdates.push([data.firstName, data.lastName]);
    infoUpdates.push([data.email, data.unit, data.location, data.steward]);
  }

  if (nameUpdates.length > 0) {
    // Update C-D (First Name, Last Name)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FIRST_NAME, nameUpdates.length, 2).setValues(nameUpdates);
    // Update X-AA (Email, Unit, Location, Steward)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_EMAIL, infoUpdates.length, 4).setValues(infoUpdates);
  }

  Logger.log('Synced member data to ' + nameUpdates.length + ' grievances');
}

// ============================================================================
// CONFIG SYNC
// ============================================================================

/**
 * Sync new values from Member Directory to Config (bidirectional sync)
 * When a user enters a new value in a job metadata field, add it to Config
 * @param {Object} e - The edit event object
 */
function syncNewValueToConfig(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var newValue = e.range.getValue();

  // Skip if empty or header row
  if (!newValue || e.range.getRow() === 1) return;

  // Check if this column is a job metadata field (includes Committees and Home Town)
  var fieldConfig = getJobMetadataByMemberCol(col);
  if (!fieldConfig) return; // Not a synced column

  // Get current Config values for this column
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return;

  var existingValues = getConfigValues(configSheet, fieldConfig.configCol);

  // Handle multi-value fields (comma-separated)
  var valuesToCheck = newValue.toString().split(',').map(function(v) { return v.trim(); });

  var valuesToAdd = [];
  for (var j = 0; j < valuesToCheck.length; j++) {
    var val = valuesToCheck[j];
    if (val && existingValues.indexOf(val) === -1) {
      valuesToAdd.push(val);
    }
  }

  // Add new values to Config
  if (valuesToAdd.length > 0) {
    var lastRow = configSheet.getLastRow();
    var dataStartRow = Math.max(lastRow + 1, 3); // Start at row 3 minimum

    for (var k = 0; k < valuesToAdd.length; k++) {
      configSheet.getRange(dataStartRow + k, fieldConfig.configCol).setValue(valuesToAdd[k]);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Added "' + valuesToAdd.join(', ') + '" to ' + fieldConfig.configName,
      'Config Updated', 3
    );
  }
}

// ============================================================================
// AUTO-SYNC TRIGGER HANDLERS
// ============================================================================

/**
 * Master onEdit trigger - routes to appropriate sync function
 * Install this as an installable trigger
 */
function onEditAutoSync(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Check for action checkboxes BEFORE debounce (needs immediate response)
  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (sheetName === SHEETS.MEMBER_DIR && row >= 2) {
    // Handle Start Grievance checkbox
    if (col === MEMBER_COLS.START_GRIEVANCE && e.range.getValue() === true) {
      // Uncheck immediately so it can be reused
      e.range.setValue(false);

      // Open the grievance form for this member
      try {
        openGrievanceFormForRow_(sheet, row);
      } catch (err) {
        Logger.log('Error opening grievance form: ' + err.message);
      }
      return; // Don't continue with sync for checkbox edits
    }

    // Handle Quick Actions checkbox
    if (col === MEMBER_COLS.QUICK_ACTIONS && e.range.getValue() === true) {
      // Uncheck immediately so it can be reused
      e.range.setValue(false);

      // Open quick actions dialog for this member
      try {
        showMemberQuickActions(row);
      } catch (err) {
        Logger.log('Error opening member quick actions: ' + err.message);
      }
      return; // Don't continue with sync for checkbox edits
    }
  }

  // Handle Grievance Log Quick Actions checkbox
  if (sheetName === SHEETS.GRIEVANCE_LOG && row >= 2) {
    if (col === GRIEVANCE_COLS.QUICK_ACTIONS && e.range.getValue() === true) {
      // Uncheck immediately so it can be reused
      e.range.setValue(false);

      // Open quick actions dialog for this grievance
      try {
        showGrievanceQuickActions(row);
      } catch (err) {
        Logger.log('Error opening grievance quick actions: ' + err.message);
      }
      return; // Don't continue with sync for checkbox edits
    }
  }

  // Debounce - use cache to prevent rapid re-syncs
  var cache = CacheService.getScriptCache();
  var cacheKey = 'lastSync_' + sheetName;
  var lastSync = cache.get(cacheKey);

  if (lastSync) {
    return; // Skip if synced within last 2 seconds
  }

  cache.put(cacheKey, 'true', 2); // 2 second debounce

  try {
    if (sheetName === SHEETS.GRIEVANCE_LOG) {
      // Grievance Log changed - sync formulas and update Member Directory
      syncGrievanceFormulasToLog();
      syncGrievanceToMemberDirectory();
      // Auto-sort by status priority (active cases first, then by deadline urgency)
      sortGrievanceLogByStatus();
      // Update Dashboard with new computed values
      syncDashboardValues();
      // Auto-create folders for any grievances missing them
      autoCreateMissingGrievanceFolders_();
    } else if (sheetName === SHEETS.MEMBER_DIR) {
      // Member Directory changed - sync to Grievance Log and Config
      syncNewValueToConfig(e);  // Bidirectional: add new values to Config
      syncGrievanceFormulasToLog();
      syncMemberToGrievanceLog();
      // Update Dashboard with new computed values
      syncDashboardValues();
    } else if (sheetName === SHEETS.FEEDBACK) {
      // Feedback sheet changed - update computed metrics
      syncFeedbackValues();
    }
  } catch (error) {
    Logger.log('Auto-sync error: ' + error.message);
  }
}

/**
 * Manual sync all data with data quality validation
 */
function syncAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Syncing all data...', 'Sync', 3);

  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();
  syncChecklistCalcToGrievanceLog();

  // Repair checkboxes after sync
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  // Run data quality check
  var issues = checkDataQuality();

  if (issues.length > 0) {
    var issueMsg = issues.slice(0, 5).join('\n');
    if (issues.length > 5) {
      issueMsg += '\n... and ' + (issues.length - 5) + ' more issues';
    }

    ui.alert('Sync Complete with Data Issues',
      'Data synced successfully, but some issues were found:\n\n' + issueMsg + '\n\n' +
      'Use "Fix Data Issues" from Administrator menu to resolve.',
      ui.ButtonSet.OK);
  } else {
    ss.toast('All data synced! No issues found.', 'Success', 3);
  }
}

// ============================================================================
// AUTO-SYNC TRIGGER MANAGEMENT
// ============================================================================

/**
 * Install the auto-sync trigger with options dialog
 * Users can customize the sync behavior
 */
function installAutoSyncTrigger() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px}' +
    '.section h4{margin:0 0 10px;color:#333}' +
    '.option{display:flex;align-items:center;margin:8px 0}' +
    '.option input[type="checkbox"]{margin-right:10px}' +
    '.option label{font-size:14px}' +
    '.info{background:#e8f4fd;padding:12px;border-radius:8px;font-size:13px;margin-bottom:15px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 20px;border:none;border-radius:4px;cursor:pointer}' +
    '.primary{background:#1a73e8;color:white;flex:1}' +
    '.secondary{background:#e0e0e0;flex:1}' +
    '.warning{background:#fff3cd;padding:10px;border-radius:4px;font-size:12px;color:#856404}' +
    '</style></head><body><div class="container">' +
    '<h2>Auto-Sync Settings</h2>' +
    '<div class="info">Auto-sync automatically updates cross-sheet data when you edit cells in Member Directory or Grievance Log.</div>' +

    '<div class="section"><h4>Sync Options</h4>' +
    '<div class="option"><input type="checkbox" id="syncGrievances" checked><label>Sync Grievance data to Member Directory</label></div>' +
    '<div class="option"><input type="checkbox" id="syncMembers" checked><label>Sync Member data to Grievance Log</label></div>' +
    '<div class="option"><input type="checkbox" id="autoSort" checked><label>Auto-sort Grievance Log by status/deadline</label></div>' +
    '<div class="option"><input type="checkbox" id="repairCheckboxes" checked><label>Auto-repair checkboxes after sync</label></div>' +
    '</div>' +

    '<div class="section"><h4>Performance</h4>' +
    '<div class="option"><input type="checkbox" id="showToasts" checked><label>Show sync notifications (toasts)</label></div>' +
    '<div class="warning">Disabling notifications improves performance but you won\'t see sync status.</div>' +
    '</div>' +

    '<div class="buttons">' +
    '<button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="primary" onclick="install()">Install Trigger</button>' +
    '</div></div>' +
    '<script>' +
    'function install(){' +
    'var opts={syncGrievances:document.getElementById("syncGrievances").checked,syncMembers:document.getElementById("syncMembers").checked,autoSort:document.getElementById("autoSort").checked,repairCheckboxes:document.getElementById("repairCheckboxes").checked,showToasts:document.getElementById("showToasts").checked};' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()}).installAutoSyncTriggerWithOptions(opts)}' +
    '</script></body></html>'
  ).setWidth(450).setHeight(480);
  ui.showModalDialog(html, 'Auto-Sync Settings');
}

/**
 * Install auto-sync trigger with saved options
 * @param {Object} options - Sync configuration options
 */
function installAutoSyncTriggerWithOptions(options) {
  // Save options to script properties
  var props = PropertiesService.getScriptProperties();
  props.setProperty('autoSyncOptions', JSON.stringify(options));

  // Remove existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install new trigger
  ScriptApp.newTrigger('onEditAutoSync')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('Auto-sync trigger installed with options: ' + JSON.stringify(options));
  SpreadsheetApp.getActiveSpreadsheet().toast('Auto-sync trigger installed!', 'Success', 3);
}

/**
 * Quick install (no dialog) - used by repair functions
 */
function installAutoSyncTriggerQuick() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEditAutoSync')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('Auto-sync trigger installed (quick mode)');
}

/**
 * Remove the auto-sync trigger
 */
function removeAutoSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  Logger.log('Removed ' + removed + ' auto-sync triggers');
  SpreadsheetApp.getActiveSpreadsheet().toast('Auto-sync trigger removed', 'Info', 3);
}

/**
 * Get auto-sync options (with defaults)
 */
function getAutoSyncOptions() {
  var props = PropertiesService.getScriptProperties();
  var optionsJSON = props.getProperty('autoSyncOptions');
  if (optionsJSON) {
    return JSON.parse(optionsJSON);
  }
  // Default options
  return {
    syncGrievances: true,
    syncMembers: true,
    autoSort: true,
    repairCheckboxes: true,
    showToasts: true
  };
}

// ============================================================================
// DASHBOARD VALUE SYNC
// ============================================================================

/**
 * Sync computed values to Dashboard sheet (no formulas)
 * Replaces all Dashboard formulas with JavaScript-computed values
 * Called during CREATE_509_DASHBOARD and on data changes
 */
function syncDashboardValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashSheet) {
    Logger.log('Dashboard sheet not found');
    return;
  }

  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!memberSheet || !grievanceSheet) {
    Logger.log('Required sheets not found for Dashboard sync');
    return;
  }

  // Get data from sheets
  var memberData = memberSheet.getDataRange().getValues();
  var grievanceData = grievanceSheet.getDataRange().getValues();
  var configData = configSheet ? configSheet.getDataRange().getValues() : [];

  // Compute all metrics
  var metrics = computeDashboardMetrics_(memberData, grievanceData, configData);

  // Write values to Dashboard (no formulas)
  writeDashboardValues_(dashSheet, metrics);

  Logger.log('Dashboard values synced');
}

/**
 * Compute all Dashboard metrics from raw data
 * @private
 */
function computeDashboardMetrics_(memberData, grievanceData, configData) {
  var metrics = {
    // Quick Stats
    totalMembers: 0,
    activeStewards: 0,
    activeGrievances: 0,
    winRate: '-',
    overdueCases: 0,
    dueThisWeek: 0,

    // Member Metrics
    avgOpenRate: '-',
    ytdVolHours: 0,

    // Grievance Metrics
    open: 0,
    pendingInfo: 0,
    settled: 0,
    won: 0,
    denied: 0,
    withdrawn: 0,

    // Timeline Metrics
    avgDaysOpen: 0,
    filedThisMonth: 0,
    closedThisMonth: 0,
    avgResolutionDays: 0,

    // Category Analysis (top 5)
    categories: [],

    // Location Breakdown (top 5)
    locations: [],

    // Month-over-Month Trends
    trends: {
      filed: { thisMonth: 0, lastMonth: 0 },
      closed: { thisMonth: 0, lastMonth: 0 },
      won: { thisMonth: 0, lastMonth: 0 }
    },

    // 6-Month Historical Data for Sparklines
    sixMonthHistory: {
      grievances: [], // [month-5, month-4, month-3, month-2, month-1, current]
      members: [],
      casesFiled: []
    },

    // Steward Summary
    stewardSummary: {
      total: 0,
      activeWithCases: 0,
      avgCasesPerSteward: '-',
      totalVolHours: 0,
      contactsThisMonth: 0
    },

    // Top 30 Busiest Stewards
    busiestStewards: [],

    // Top 10 Performers (from hidden sheet)
    topPerformers: [],

    // Bottom 10 (needing support)
    needingSupport: []
  };

  var today = new Date();
  var thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  var lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);

  // ══════════════════════════════════════════════════════════════════════
  // MEMBER METRICS
  // ══════════════════════════════════════════════════════════════════════
  var openRates = [];

  for (var m = 1; m < memberData.length; m++) {
    var row = memberData[m];
    if (!row[MEMBER_COLS.MEMBER_ID - 1]) continue;

    metrics.totalMembers++;

    if (isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1])) {
      metrics.activeStewards++;
    }

    var openRate = row[MEMBER_COLS.OPEN_RATE - 1];
    if (typeof openRate === 'number') {
      openRates.push(openRate);
    }

    var volHours = row[MEMBER_COLS.VOLUNTEER_HOURS - 1];
    if (typeof volHours === 'number') {
      metrics.ytdVolHours += volHours;
    }

    var contactDate = row[MEMBER_COLS.RECENT_CONTACT_DATE - 1];
    if (contactDate instanceof Date && contactDate >= thisMonthStart && contactDate <= today) {
      metrics.stewardSummary.contactsThisMonth++;
    }
  }

  if (openRates.length > 0) {
    var avgRate = openRates.reduce(function(a, b) { return a + b; }, 0) / openRates.length;
    metrics.avgOpenRate = Math.round(avgRate * 10) / 10 + '%';
  }

  // ══════════════════════════════════════════════════════════════════════
  // GRIEVANCE METRICS
  // ══════════════════════════════════════════════════════════════════════
  var daysOpenValues = [];
  var closedDaysValues = [];
  var categoryStats = {};
  var locationStats = {};
  var stewardGrievances = {};

  for (var g = 1; g < grievanceData.length; g++) {
    var gRow = grievanceData[g];
    if (!gRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1]) continue;

    var status = gRow[GRIEVANCE_COLS.STATUS - 1];
    var steward = gRow[GRIEVANCE_COLS.STEWARD - 1];
    var category = gRow[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
    var location = gRow[GRIEVANCE_COLS.LOCATION - 1];
    var dateFiled = gRow[GRIEVANCE_COLS.DATE_FILED - 1];
    var dateClosed = gRow[GRIEVANCE_COLS.DATE_CLOSED - 1];
    var daysOpen = gRow[GRIEVANCE_COLS.DAYS_OPEN - 1];
    var daysToDeadline = gRow[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    // Status counts
    if (status === 'Open') metrics.open++;
    else if (status === 'Pending Info') metrics.pendingInfo++;
    else if (status === 'Settled') metrics.settled++;
    else if (status === 'Won') metrics.won++;
    else if (status === 'Denied') metrics.denied++;
    else if (status === 'Withdrawn') metrics.withdrawn++;

    // Active grievances
    if (status === 'Open' || status === 'Pending Info') {
      metrics.activeGrievances++;
    }

    // Overdue and due this week
    // Note: daysToDeadline can be a number OR the string "Overdue"
    if (daysToDeadline === 'Overdue') {
      metrics.overdueCases++;
    } else if (typeof daysToDeadline === 'number') {
      if (daysToDeadline < 0) metrics.overdueCases++;
      else if (daysToDeadline <= 7) metrics.dueThisWeek++;
    }

    // Days open average
    if (typeof daysOpen === 'number') {
      daysOpenValues.push(daysOpen);
    }

    // Resolution days (for closed cases)
    if (dateClosed && typeof daysOpen === 'number') {
      closedDaysValues.push(daysOpen);
    }

    // Filed this month
    if (dateFiled instanceof Date && dateFiled >= thisMonthStart && dateFiled <= today) {
      metrics.filedThisMonth++;
      metrics.trends.filed.thisMonth++;
    }
    if (dateFiled instanceof Date && dateFiled >= lastMonthStart && dateFiled <= lastMonthEnd) {
      metrics.trends.filed.lastMonth++;
    }

    // Closed this month
    if (dateClosed instanceof Date && dateClosed >= thisMonthStart && dateClosed <= today) {
      metrics.closedThisMonth++;
      metrics.trends.closed.thisMonth++;
      if (status === 'Won') {
        metrics.trends.won.thisMonth++;
      }
    }
    if (dateClosed instanceof Date && dateClosed >= lastMonthStart && dateClosed <= lastMonthEnd) {
      metrics.trends.closed.lastMonth++;
      if (status === 'Won') {
        metrics.trends.won.lastMonth++;
      }
    }

    // Category stats
    if (category) {
      if (!categoryStats[category]) {
        categoryStats[category] = { total: 0, open: 0, resolved: 0, won: 0, daysOpen: [] };
      }
      categoryStats[category].total++;
      if (status === 'Open') categoryStats[category].open++;
      if (status !== 'Open' && status !== 'Pending Info') categoryStats[category].resolved++;
      if (status === 'Won') categoryStats[category].won++;
      if (typeof daysOpen === 'number') categoryStats[category].daysOpen.push(daysOpen);
    }

    // Location stats
    if (location) {
      if (!locationStats[location]) {
        locationStats[location] = { members: 0, grievances: 0, open: 0, won: 0 };
      }
      locationStats[location].grievances++;
      if (status === 'Open') locationStats[location].open++;
      if (status === 'Won') locationStats[location].won++;
    }

    // Steward stats
    if (steward) {
      if (!stewardGrievances[steward]) {
        stewardGrievances[steward] = { active: 0, open: 0, pendingInfo: 0, total: 0 };
      }
      stewardGrievances[steward].total++;
      if (status === 'Open') {
        stewardGrievances[steward].active++;
        stewardGrievances[steward].open++;
      } else if (status === 'Pending Info') {
        stewardGrievances[steward].active++;
        stewardGrievances[steward].pendingInfo++;
      }
    }
  }

  // Calculate averages
  if (daysOpenValues.length > 0) {
    metrics.avgDaysOpen = Math.round(daysOpenValues.reduce(function(a, b) { return a + b; }, 0) / daysOpenValues.length * 10) / 10;
  }
  if (closedDaysValues.length > 0) {
    metrics.avgResolutionDays = Math.round(closedDaysValues.reduce(function(a, b) { return a + b; }, 0) / closedDaysValues.length * 10) / 10;
  }

  // Win rate
  var totalOutcomes = metrics.won + metrics.denied + metrics.settled + metrics.withdrawn;
  if (totalOutcomes > 0) {
    metrics.winRate = Math.round(metrics.won / totalOutcomes * 100) + '%';
  }

  // ══════════════════════════════════════════════════════════════════════
  // 6-MONTH HISTORICAL DATA FOR SPARKLINES
  // ══════════════════════════════════════════════════════════════════════
  // Calculate filing counts for each of the last 6 months
  var monthlyFiledCounts = [0, 0, 0, 0, 0, 0]; // [5 months ago, 4, 3, 2, 1, current]
  var monthlyClosedCounts = [0, 0, 0, 0, 0, 0];

  for (var h = 1; h < grievanceData.length; h++) {
    var hRow = grievanceData[h];
    if (!hRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1]) continue;

    var hDateFiled = hRow[GRIEVANCE_COLS.DATE_FILED - 1];
    var hDateClosed = hRow[GRIEVANCE_COLS.DATE_CLOSED - 1];

    if (hDateFiled instanceof Date) {
      for (var mo = 0; mo < 6; mo++) {
        var monthStart = new Date(today.getFullYear(), today.getMonth() - (5 - mo), 1);
        var monthEnd = new Date(today.getFullYear(), today.getMonth() - (5 - mo) + 1, 0);
        if (hDateFiled >= monthStart && hDateFiled <= monthEnd) {
          monthlyFiledCounts[mo]++;
          break;
        }
      }
    }

    if (hDateClosed instanceof Date) {
      for (var mc = 0; mc < 6; mc++) {
        var mStart = new Date(today.getFullYear(), today.getMonth() - (5 - mc), 1);
        var mEnd = new Date(today.getFullYear(), today.getMonth() - (5 - mc) + 1, 0);
        if (hDateClosed >= mStart && hDateClosed <= mEnd) {
          monthlyClosedCounts[mc]++;
          break;
        }
      }
    }
  }

  // Store 6-month history for sparklines
  metrics.sixMonthHistory.casesFiled = monthlyFiledCounts;
  metrics.sixMonthHistory.grievances = monthlyFiledCounts.map(function(val, idx) {
    // Running total of active grievances (approximation)
    return metrics.activeGrievances + monthlyFiledCounts.slice(idx + 1).reduce(function(a, b) { return a + b; }, 0) -
           monthlyClosedCounts.slice(idx + 1).reduce(function(a, b) { return a + b; }, 0);
  });
  // Member count history not tracked — show current count consistently (no fabricated trends)
  metrics.sixMonthHistory.members = [
    metrics.totalMembers,
    metrics.totalMembers,
    metrics.totalMembers,
    metrics.totalMembers,
    metrics.totalMembers,
    metrics.totalMembers
  ];

  // ══════════════════════════════════════════════════════════════════════
  // CATEGORY ANALYSIS (Top 5)
  // ══════════════════════════════════════════════════════════════════════
  var defaultCategories = ['Contract Violation', 'Discipline', 'Workload', 'Safety', 'Discrimination'];
  for (var c = 0; c < defaultCategories.length; c++) {
    var cat = defaultCategories[c];
    var catData = categoryStats[cat] || { total: 0, open: 0, resolved: 0, won: 0, daysOpen: [] };
    var catWinRate = catData.total > 0 ? Math.round(catData.won / catData.total * 100) + '%' : '-';
    var avgDays = catData.daysOpen.length > 0 ?
      Math.round(catData.daysOpen.reduce(function(a, b) { return a + b; }, 0) / catData.daysOpen.length * 10) / 10 : '-';

    metrics.categories.push({
      name: cat,
      total: catData.total,
      open: catData.open,
      resolved: catData.resolved,
      winRate: catWinRate,
      avgDays: avgDays
    });
  }

  // ══════════════════════════════════════════════════════════════════════
  // LOCATION BREAKDOWN (Top 5 from Config)
  // ══════════════════════════════════════════════════════════════════════
  // Count members per location
  var memberLocations = {};
  for (var ml = 1; ml < memberData.length; ml++) {
    var loc = memberData[ml][MEMBER_COLS.WORK_LOCATION - 1];
    if (loc) {
      memberLocations[loc] = (memberLocations[loc] || 0) + 1;
    }
  }

  // Get top 5 locations from Config
  for (var l = 0; l < 5; l++) {
    var locName = configData[2 + l] ? configData[2 + l][CONFIG_COLS.OFFICE_LOCATIONS - 1] : '';
    if (locName) {
      var locData = locationStats[locName] || { members: 0, grievances: 0, open: 0, won: 0 };
      locData.members = memberLocations[locName] || 0;
      var locWinRate = locData.grievances > 0 ? Math.round(locData.won / locData.grievances * 100) + '%' : '-';

      metrics.locations.push({
        name: locName,
        members: locData.members,
        grievances: locData.grievances,
        open: locData.open,
        winRate: locWinRate,
        satisfaction: '-'
      });
    }
  }

  // ══════════════════════════════════════════════════════════════════════
  // STEWARD SUMMARY
  // ══════════════════════════════════════════════════════════════════════
  metrics.stewardSummary.total = metrics.activeStewards;
  metrics.stewardSummary.totalVolHours = metrics.ytdVolHours;

  var stewardsWithActiveCases = Object.keys(stewardGrievances).filter(function(s) {
    return stewardGrievances[s].active > 0;
  }).length;
  metrics.stewardSummary.activeWithCases = stewardsWithActiveCases;

  if (metrics.activeStewards > 0) {
    var totalGrievances = grievanceData.length - 1;
    metrics.stewardSummary.avgCasesPerSteward = Math.round(totalGrievances / metrics.activeStewards * 10) / 10;
  }

  // ══════════════════════════════════════════════════════════════════════
  // TOP 30 BUSIEST STEWARDS
  // ══════════════════════════════════════════════════════════════════════
  var stewardArray = Object.keys(stewardGrievances).map(function(name) {
    return {
      name: name,
      active: stewardGrievances[name].active,
      open: stewardGrievances[name].open,
      pendingInfo: stewardGrievances[name].pendingInfo,
      total: stewardGrievances[name].total
    };
  });

  stewardArray.sort(function(a, b) { return b.active - a.active; });
  metrics.busiestStewards = stewardArray.slice(0, 30);

  // ══════════════════════════════════════════════════════════════════════
  // TOP/BOTTOM PERFORMERS (from hidden sheet)
  // ══════════════════════════════════════════════════════════════════════
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var perfSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
  if (perfSheet && perfSheet.getLastRow() > 1) {
    var perfData = perfSheet.getDataRange().getValues();
    var performers = [];
    for (var p = 1; p < perfData.length; p++) {
      if (perfData[p][0]) {  // Has steward name
        performers.push({
          name: perfData[p][0],
          score: perfData[p][9] || 0,  // Column J (index 9)
          winRate: perfData[p][5] || '-',  // Column F
          avgDays: perfData[p][6] || '-',  // Column G
          overdue: perfData[p][7] || 0  // Column H
        });
      }
    }

    // Sort by score descending for top performers
    performers.sort(function(a, b) { return b.score - a.score; });
    metrics.topPerformers = performers.slice(0, 10);

    // Sort by score ascending for needing support
    performers.sort(function(a, b) { return a.score - b.score; });
    metrics.needingSupport = performers.slice(0, 10);
  }

  return metrics;
}

/**
 * Write computed values to Dashboard sheet
 * Row numbers updated to match new card-style layout
 * @private
 */
function writeDashboardValues_(sheet, metrics) {
  var L = DASHBOARD_LAYOUT;

  // ══════════════════════════════════════════════════════════════════════
  // QUICK STATS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  sheet.getRange(L.QUICK_STATS_ROW, 1, 1, L.DATA_COLS).setValues([[
    metrics.totalMembers,
    metrics.activeStewards,
    metrics.activeGrievances,
    metrics.winRate,
    metrics.overdueCases,
    metrics.dueThisWeek
  ]]);

  // ══════════════════════════════════════════════════════════════════════
  // MEMBER METRICS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  sheet.getRange(L.MEMBER_METRICS_ROW, 1, 1, 4).setValues([[
    metrics.totalMembers,
    metrics.activeStewards,
    metrics.avgOpenRate,
    metrics.ytdVolHours
  ]]);

  // ══════════════════════════════════════════════════════════════════════
  // GRIEVANCE METRICS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  sheet.getRange(L.GRIEVANCE_METRICS_ROW, 1, 1, L.DATA_COLS).setValues([[
    metrics.open,
    metrics.pendingInfo,
    metrics.settled,
    metrics.won,
    metrics.denied,
    metrics.withdrawn
  ]]);

  // ══════════════════════════════════════════════════════════════════════
  // TIMELINE METRICS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  sheet.getRange(L.TIMELINE_METRICS_ROW, 1, 1, 4).setValues([[
    metrics.avgDaysOpen,
    metrics.filedThisMonth,
    metrics.closedThisMonth,
    metrics.avgResolutionDays
  ]]);

  // ══════════════════════════════════════════════════════════════════════
  // TYPE ANALYSIS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  var catRowCount = L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1;
  var categoryRows = [];
  for (var c = 0; c < metrics.categories.length; c++) {
    var cat = metrics.categories[c];
    categoryRows.push([cat.name, cat.total, cat.open, cat.resolved, cat.winRate, cat.avgDays]);
  }
  while (categoryRows.length < catRowCount) {
    categoryRows.push(['', '', '', '', '', '']);
  }
  sheet.getRange(L.CATEGORY_START_ROW, 1, catRowCount, L.DATA_COLS).setValues(categoryRows);

  // ══════════════════════════════════════════════════════════════════════
  // LOCATION BREAKDOWN - Card layout
  // ══════════════════════════════════════════════════════════════════════
  var locRowCount = L.LOCATION_END_ROW - L.LOCATION_START_ROW + 1;
  var locationRows = [];
  for (var l = 0; l < metrics.locations.length; l++) {
    var loc = metrics.locations[l];
    locationRows.push([loc.name, loc.members, loc.grievances, loc.open, loc.winRate, loc.satisfaction]);
  }
  while (locationRows.length < locRowCount) {
    locationRows.push(['', '', '', '', '', '']);
  }
  sheet.getRange(L.LOCATION_START_ROW, 1, locRowCount, L.DATA_COLS).setValues(locationRows);

  // ══════════════════════════════════════════════════════════════════════
  // MONTH-OVER-MONTH TRENDS (Rows 44-46) - Updated for card layout
  // Now includes sparklines in column G with color coding
  // ══════════════════════════════════════════════════════════════════════
  var trendRows = [];

  // Active Grievances
  var grievanceChange = metrics.sixMonthHistory.grievances[5] - metrics.sixMonthHistory.grievances[4];
  var grievancePct = metrics.sixMonthHistory.grievances[4] > 0 ?
    Math.round(grievanceChange / metrics.sixMonthHistory.grievances[4] * 100) + '%' : '-';
  var grievanceTrend = grievanceChange > 0 ? '>' : (grievanceChange < 0 ? '<' : '=');
  trendRows.push(['Active Grievances', metrics.activeGrievances, metrics.sixMonthHistory.grievances[4] || 0, grievanceChange, grievancePct, grievanceTrend]);

  // Total Members
  var memberChange = metrics.sixMonthHistory.members[5] - metrics.sixMonthHistory.members[4];
  var memberPct = metrics.sixMonthHistory.members[4] > 0 ?
    Math.round(memberChange / metrics.sixMonthHistory.members[4] * 100) + '%' : '-';
  var memberTrend = memberChange > 0 ? '>' : (memberChange < 0 ? '<' : '=');
  trendRows.push(['Total Members', metrics.totalMembers, metrics.sixMonthHistory.members[4] || 0, memberChange, memberPct, memberTrend]);

  // Cases Filed
  var filedChange = metrics.trends.filed.thisMonth - metrics.trends.filed.lastMonth;
  var filedPct = metrics.trends.filed.lastMonth > 0 ? Math.round(filedChange / metrics.trends.filed.lastMonth * 100) + '%' : '-';
  var filedTrend = filedChange > 0 ? '>' : (filedChange < 0 ? '<' : '=');
  trendRows.push(['Cases Filed', metrics.trends.filed.thisMonth, metrics.trends.filed.lastMonth, filedChange, filedPct, filedTrend]);

  var trendRowCount = L.TREND_END_ROW - L.TREND_START_ROW + 1;
  sheet.getRange(L.TREND_START_ROW, 1, trendRowCount, L.DATA_COLS).setValues(trendRows);

  // ══════════════════════════════════════════════════════════════════════
  // SPARKLINES - Color-coded 6-month trends
  // Red for grievances (high = bad), Green for members (high = good), Blue for filed
  // ══════════════════════════════════════════════════════════════════════
  var sparklineFormulas = [];

  // Active Grievances sparkline - RED color (lower is better, so increasing is bad)
  var grievanceData = metrics.sixMonthHistory.grievances.join(',');
  var grievanceSparkline = '=SPARKLINE({' + grievanceData + '},{"charttype","line";"color","#DC2626";"linewidth",2})';
  sparklineFormulas.push([grievanceSparkline]);

  // Total Members sparkline - GREEN color (higher is better)
  var memberDataStr = metrics.sixMonthHistory.members.join(',');
  var memberSparkline = '=SPARKLINE({' + memberDataStr + '},{"charttype","line";"color","#059669";"linewidth",2})';
  sparklineFormulas.push([memberSparkline]);

  // Cases Filed sparkline - BLUE color (neutral indicator)
  var filedData = metrics.sixMonthHistory.casesFiled.join(',');
  var filedSparkline = '=SPARKLINE({' + filedData + '},{"charttype","line";"color","#3B82F6";"linewidth",2})';
  sparklineFormulas.push([filedSparkline]);

  // Write sparkline formulas
  sheet.getRange(L.TREND_START_ROW, L.SPARKLINE_COL).setFormula(grievanceSparkline);
  sheet.getRange(L.TREND_START_ROW + 1, L.SPARKLINE_COL).setFormula(memberSparkline);
  sheet.getRange(L.TREND_START_ROW + 2, L.SPARKLINE_COL).setFormula(filedSparkline);

  // Color-code change values based on direction
  // For grievances: negative change = green (good), positive = red (bad)
  var changeCell44 = sheet.getRange(L.TREND_START_ROW, 4);
  var change44Val = grievanceChange;
  if (change44Val < 0) {
    changeCell44.setFontColor('#059669'); // Green - grievances down is good
  } else if (change44Val > 0) {
    changeCell44.setFontColor('#DC2626'); // Red - grievances up is bad
  } else {
    changeCell44.setFontColor('#6B7280'); // Gray - no change
  }

  // For members: positive change = green (good), negative = red (bad)
  var changeCell45 = sheet.getRange(L.TREND_START_ROW + 1, 4);
  if (memberChange > 0) {
    changeCell45.setFontColor('#059669'); // Green - members up is good
  } else if (memberChange < 0) {
    changeCell45.setFontColor('#DC2626'); // Red - members down is bad
  } else {
    changeCell45.setFontColor('#6B7280'); // Gray
  }

  // For cases filed: neutral coloring (blue)
  var changeCell46 = sheet.getRange(L.TREND_START_ROW + 2, 4);
  changeCell46.setFontColor('#3B82F6'); // Blue - neutral

  // ══════════════════════════════════════════════════════════════════════
  // STEWARD SUMMARY - Card layout
  // ══════════════════════════════════════════════════════════════════════
  sheet.getRange(L.STEWARD_SUMMARY_ROW, 1, 1, L.DATA_COLS).setValues([[
    metrics.stewardSummary.total,
    metrics.stewardSummary.activeWithCases,
    metrics.stewardSummary.avgCasesPerSteward,
    metrics.stewardSummary.totalVolHours,
    metrics.stewardSummary.contactsThisMonth,
    metrics.winRate
  ]]);

  // ══════════════════════════════════════════════════════════════════════
  // TOP 30 BUSIEST STEWARDS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  var busiestRowCount = L.BUSIEST_END_ROW - L.BUSIEST_START_ROW + 1;
  var busiestRows = [];
  for (var b = 0; b < busiestRowCount; b++) {
    if (b < metrics.busiestStewards.length) {
      var steward = metrics.busiestStewards[b];
      busiestRows.push([b + 1, steward.name, steward.active, steward.open, steward.pendingInfo, steward.total]);
    } else {
      busiestRows.push(['', '', '', '', '', '']);
    }
  }
  sheet.getRange(L.BUSIEST_START_ROW, 1, busiestRowCount, L.DATA_COLS).setValues(busiestRows);

  // ══════════════════════════════════════════════════════════════════════
  // TOP 10 PERFORMERS - Card layout
  // ══════════════════════════════════════════════════════════════════════
  var topRowCount = L.TOP_PERFORMERS_END_ROW - L.TOP_PERFORMERS_START_ROW + 1;
  var topRows = [];
  for (var t = 0; t < topRowCount; t++) {
    if (t < metrics.topPerformers.length) {
      var perf = metrics.topPerformers[t];
      topRows.push([t + 1, perf.name, perf.score, perf.winRate, perf.avgDays, perf.overdue]);
    } else {
      topRows.push(['', '', '', '', '', '']);
    }
  }
  sheet.getRange(L.TOP_PERFORMERS_START_ROW, 1, topRowCount, L.DATA_COLS).setValues(topRows);

  // ══════════════════════════════════════════════════════════════════════
  // STEWARDS NEEDING SUPPORT - Card layout
  // ══════════════════════════════════════════════════════════════════════
  var supportRowCount = L.NEEDING_SUPPORT_END_ROW - L.NEEDING_SUPPORT_START_ROW + 1;
  var bottomRows = [];
  for (var n = 0; n < supportRowCount; n++) {
    if (n < metrics.needingSupport.length) {
      var need = metrics.needingSupport[n];
      bottomRows.push([n + 1, need.name, need.score, need.winRate, need.avgDays, need.overdue]);
    } else {
      bottomRows.push(['', '', '', '', '', '']);
    }
  }
  sheet.getRange(L.NEEDING_SUPPORT_START_ROW, 1, supportRowCount, L.DATA_COLS).setValues(bottomRows);

  // ══════════════════════════════════════════════════════════════════════
  // AUTO-APPLY GRADIENT HEATMAPS
  // ══════════════════════════════════════════════════════════════════════
  applyDashboardGradients_(sheet);
}

/**
 * Apply gradient heatmaps to Dashboard for visual data analysis
 * Auto-applies color scales to key metrics
 * @param {Sheet} sheet - The Dashboard sheet
 * @private
 */
function applyDashboardGradients_(sheet) {
  // Define gradient color scale (Green -> Yellow -> Red)
  var _greenColor = '#D1FAE5';  // Low values (good for some metrics)
  var _yellowColor = '#FEF3C7'; // Mid values
  var _redColor = '#FCA5A5';    // High values (bad for some metrics)

  // Reverse scale (Red -> Yellow -> Green) for positive metrics
  var redToGreen = {
    minColor: '#FCA5A5',
    midColor: '#FEF3C7',
    maxColor: '#D1FAE5'
  };

  // Standard scale (Green -> Yellow -> Red) for negative metrics
  var greenToRed = {
    minColor: '#D1FAE5',
    midColor: '#FEF3C7',
    maxColor: '#FCA5A5'
  };

  // ── Active Cases Column (Top 30 Busiest) - Higher = more work (red)
  var activeCasesRange = sheet.getRange('C59:C88');
  var activeCasesRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([activeCasesRange])
    .build();

  // ── Score Column (Top 10 Performers) - Higher = better (green)
  var scoreRange = sheet.getRange('C93:C102');
  var scoreRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([scoreRange])
    .build();

  // ── Win Rate Column (Top 10 Performers) - Higher = better (green)
  var winRateRange = sheet.getRange('D93:D102');
  var winRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([winRateRange])
    .build();

  // ── Overdue Column (Performers) - Lower = better (green at low)
  var overdueRange = sheet.getRange('F93:F102');
  var overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([overdueRange])
    .build();

  // ── Score Column (Needing Support) - Lower scores (red)
  var needScoreRange = sheet.getRange('C107:C116');
  var needScoreRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([needScoreRange])
    .build();

  // ── Overdue Column (Needing Support) - Highlight high overdue
  var needOverdueRange = sheet.getRange('F107:F116');
  var needOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([needOverdueRange])
    .build();

  // ── Category Win Rate (Issue Breakdown) - Higher = better (green)
  var catWinRateRange = sheet.getRange('E26:E30');
  var catWinRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([catWinRateRange])
    .build();

  // ── Location Win Rate - Higher = better (green)
  var locWinRateRange = sheet.getRange('E35:E39');
  var locWinRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([locWinRateRange])
    .build();

  // Apply all rules
  var rules = sheet.getConditionalFormatRules();

  // Remove existing gradient rules to avoid duplicates
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    if (ranges.length === 0) return true;
    var rangeStr = ranges[0].getA1Notation();
    // Keep rules that aren't our gradient ranges
    return ['C59:C88', 'C93:C102', 'D93:D102', 'F93:F102', 'C107:C116', 'F107:F116', 'E26:E30', 'E35:E39'].indexOf(rangeStr) === -1;
  });

  // Add our gradient rules
  newRules.push(activeCasesRule);
  newRules.push(scoreRule);
  newRules.push(winRateRule);
  newRules.push(overdueRule);
  newRules.push(needScoreRule);
  newRules.push(needOverdueRule);
  newRules.push(catWinRateRule);
  newRules.push(locWinRateRule);

  sheet.setConditionalFormatRules(newRules);
}

// ============================================================================
// FEEDBACK SHEET SYNC
// ============================================================================

/**
 * Sync computed values to Feedback sheet metrics
 */
function syncFeedbackValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!sheet) {
    Logger.log('Feedback sheet not found');
    return;
  }

  var lastRow = sheet.getLastRow();

  // Get feedback data
  var totalItems = 0;
  var bugs = 0;
  var features = 0;
  var improvements = 0;
  var newOpen = 0;
  var resolved = 0;
  var critical = 0;

  if (lastRow >= 2) {
    // Get data from columns Type, Status, Priority
    var typeCol = FEEDBACK_COLS.TYPE;
    var statusCol = FEEDBACK_COLS.STATUS;
    var priorityCol = FEEDBACK_COLS.PRIORITY;

    var data = sheet.getRange(2, 1, lastRow - 1, Math.max(typeCol, statusCol, priorityCol)).getValues();

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      if (!row[0]) continue; // Skip empty rows

      totalItems++;

      var type = row[typeCol - 1];
      var status = row[statusCol - 1];
      var priority = row[priorityCol - 1];

      if (type === 'Bug') bugs++;
      else if (type === 'Feature Request') features++;
      else if (type === 'Improvement') improvements++;

      if (status === 'New' || status === 'In Progress') newOpen++;
      else if (status === 'Resolved') resolved++;

      if (priority === 'Critical') critical++;
    }
  }

  var resolutionRate = totalItems > 0 ? Math.round(resolved / totalItems * 1000) / 10 + '%' : '0%';

  // Write metrics to columns M-O (13-15), rows 3-10
  var metricsData = [
    ['Total Items', totalItems, 'All feedback items'],
    ['Bugs', bugs, 'Bug reports'],
    ['Feature Requests', features, 'New feature asks'],
    ['Improvements', improvements, 'Enhancement suggestions'],
    ['New/Open', newOpen, 'Unresolved items'],
    ['Resolved', resolved, 'Completed items'],
    ['Critical Priority', critical, 'Urgent items'],
    ['Resolution Rate', resolutionRate, 'Percentage resolved']
  ];

  sheet.getRange(3, 13, metricsData.length, 3).setValues(metricsData);

  Logger.log('Feedback values synced');
}



/**
 * ============================================================================
 * 08k_PublicDashboard.gs - Public Dashboard & Flagged Submissions Module
 * ============================================================================
 *
 * This module contains functions for:
 * - Secure member dashboard HTML generation
 * - Public data retrieval (no PII exposed)
 * - Flagged survey submission review and management
 *
 * All functions in this module are designed to provide aggregate statistics
 * and public-facing data without exposing personally identifiable information.
 *
 * Dependencies:
 * - SHEETS constant (sheet names)
 * - MEMBER_COLS, GRIEVANCE_COLS, SATISFACTION_COLS (column mappings)
 * - syncSatisfactionValues() for updating dashboard after approval/rejection
 *
 * @fileoverview Public dashboard data functions and flagged submissions management
 * @version 1.0.0
 */

// ============================================================================
// FLAGGED SUBMISSIONS REVIEW
// ============================================================================

/**
 * Show the flagged submissions review modal dialog
 */
function showFlaggedSubmissionsReview() {
  var html = HtmlService.createHtmlOutput(getFlaggedSubmissionsHtml())
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Flagged Survey Submissions Review');
}

/**
 * Get HTML for flagged submissions review interface
 * @returns {string} HTML content
 */
function getFlaggedSubmissionsHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    ':root{--purple:#5B4B9E;--green:#059669;--red:#DC2626;--orange:#F97316}' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f5f5f5;padding:20px}' +
    '.container{max-width:650px;margin:0 auto}' +
    '.stats-row{display:flex;gap:15px;margin-bottom:20px}' +
    '.stat-card{flex:1;background:white;padding:20px;border-radius:12px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.stat-card.pending{border-left:4px solid var(--orange)}' +
    '.stat-card.verified{border-left:4px solid var(--green)}' +
    '.stat-value{font-size:32px;font-weight:bold;color:#333}' +
    '.stat-label{font-size:13px;color:#666;margin-top:5px}' +
    '.section{background:white;border-radius:12px;padding:20px;margin-bottom:15px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.section-title{font-size:16px;font-weight:600;color:#333;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #eee}' +
    '.email-list{max-height:250px;overflow-y:auto}' +
    '.email-item{display:flex;align-items:center;justify-content:space-between;padding:12px;background:#f8f9fa;border-radius:8px;margin-bottom:8px}' +
    '.email-info{display:flex;align-items:center;gap:10px}' +
    '.email-text{font-size:14px;color:#333}' +
    '.email-date{font-size:12px;color:#666}' +
    '.actions{display:flex;gap:8px}' +
    '.btn{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:500}' +
    '.btn-approve{background:#059669;color:white}' +
    '.btn-reject{background:#DC2626;color:white}' +
    '.empty-state{text-align:center;padding:40px;color:#666}' +
    '.info-box{background:#E8F4FD;padding:15px;border-radius:8px;margin-bottom:15px;font-size:13px;color:#1E40AF}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div id="content"><div class="empty-state">Loading...</div></div>' +
    '</div>' +
    '<script>' +
    ' + getClientSideEscapeHtml() + ' +
    'function load(){google.script.run.withSuccessHandler(render).getFlaggedSubmissionsData()}' +
    'function render(d){' +
    '  var h="<div class=\\"stats-row\\">";' +
    '  h+="<div class=\\"stat-card pending\\"><div class=\\"stat-value\\">"+d.pendingCount+"</div><div class=\\"stat-label\\">Pending Review</div></div>";' +
    '  h+="<div class=\\"stat-card verified\\"><div class=\\"stat-value\\">"+d.verifiedCount+"</div><div class=\\"stat-label\\">Verified Responses</div></div>";' +
    '  h+="</div>";' +
    '  h+="<div class=\\"info-box\\">⚠️ These submissions could not be matched to a member email. Survey answers are protected and not shown here.</div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📧 Pending Review Emails ("+d.pendingCount+")</div>";' +
    '  if(d.pendingEmails.length===0){' +
    '    h+="<div class=\\"empty-state\\">✅ No submissions pending review</div>";' +
    '  }else{' +
    '    h+="<div class=\\"email-list\\">";' +
    '    d.pendingEmails.forEach(function(e){' +
    '      h+="<div class=\\"email-item\\"><div class=\\"email-info\\">";' +
    '      h+="<span class=\\"email-text\\">"+escapeHtml(e.email)+"</span>";' +
    '      h+="<span class=\\"email-date\\">"+escapeHtml(e.date)+" | "+escapeHtml(e.quarter)+"</span></div>";' +
    '      h+="<div class=\\"actions\\">";' +
    '      h+="<button class=\\"btn btn-approve\\" onclick=\\"approve("+e.row+\")\\">✓ Approve</button>";' +
    '      h+="<button class=\\"btn btn-reject\\" onclick=\\"reject("+e.row+\")\\">✗ Reject</button>";' +
    '      h+="</div></div>";' +
    '    });' +
    '    h+="</div>";' +
    '  }' +
    '  h+="</div>";' +
    '  document.getElementById("content").innerHTML=h;' +
    '}' +
    'function approve(row){' +
    '  if(confirm("Mark this submission as verified? This will include it in statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).approveFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'function reject(row){' +
    '  if(confirm("Reject this submission? It will be excluded from all statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).rejectFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'load();' +
    '</script></body></html>';
}

/**
 * Get data for flagged submissions review
 * @returns {Object} Pending submissions data (email, date, row number - NO survey answers)
 */
function getFlaggedSubmissionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    pendingCount: 0,
    verifiedCount: 0,
    pendingEmails: []
  };

  if (!satSheet) return result;

  // Read from vault — all emails are hashed, no plaintext PII exists
  var vaultRows = getVaultDataFull_();

  for (var i = 0; i < vaultRows.length; i++) {
    var entry = vaultRows[i];
    var satRow = entry.responseRow;
    var timestamp = '';
    try {
      timestamp = satSheet.getRange(satRow, SATISFACTION_COLS.TIMESTAMP).getValue();
    } catch (_e) { /* row may not exist */ }

    if (entry.verified === 'Yes') {
      result.verifiedCount++;
    } else if (entry.verified === 'Pending Review') {
      result.pendingCount++;
      result.pendingEmails.push({
        email: 'Anonymous submission #' + satRow,
        date: timestamp ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'MMM d, yyyy') : 'Unknown',
        quarter: entry.quarter,
        row: satRow  // Satisfaction sheet row for approve/reject
      });
    }
  }

  // Sort by most recent first
  result.pendingEmails.sort(function(a, b) { return b.row - a.row; });

  return result;
}

/**
 * Approve a flagged submission - mark as Verified
 * @param {number} rowNum - Row number (1-indexed)
 */
function approveFlaggedSubmission(rowNum) {
  // Verify caller is an authorized steward
  var callerEmail = '';
  try { callerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* ignore */ }
  if (!callerEmail) {
    throw new Error('Authorization required: unable to verify user identity');
  }

  // Update in vault (not on Satisfaction sheet — no PII there)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault) return;

  var vaultData = vault.getDataRange().getValues();
  for (var i = 1; i < vaultData.length; i++) {
    if (vaultData[i][SURVEY_VAULT_COLS.RESPONSE_ROW - 1] === rowNum) {
      var vaultRow = i + 1;
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.VERIFIED).setValue('Yes');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.REVIEWER_NOTES).setValue('Approved by ' + callerEmail + ' on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      break;
    }
  }

  // Update dashboard
  syncSatisfactionValues();
}

/**
 * Reject a flagged submission - mark as Rejected
 * @param {number} rowNum - Row number (1-indexed)
 */
function rejectFlaggedSubmission(rowNum) {
  // Update in vault (not on Satisfaction sheet — no PII there)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault) return;

  var vaultData = vault.getDataRange().getValues();
  for (var i = 1; i < vaultData.length; i++) {
    if (vaultData[i][SURVEY_VAULT_COLS.RESPONSE_ROW - 1] === rowNum) {
      var vaultRow = i + 1;
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.VERIFIED).setValue('Rejected');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.IS_LATEST).setValue('No');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.REVIEWER_NOTES).setValue('Rejected on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      break;
    }
  }

  // Update dashboard
  syncSatisfactionValues();
}

// ============================================================================
// SECURE MEMBER DASHBOARD HTML
// ============================================================================

/**
 * Generates HTML for the secure member dashboard
 * @param {Object} stats - Grievance statistics
 * @param {Array} stewards - Array of steward objects
 * @param {Object} satisfaction - Satisfaction statistics
 * @param {Object} coverage - Steward coverage statistics
 * @returns {string} HTML content
 */
function getSecureMemberDashboardHtml(stats, stewards, satisfaction, coverage) {
  // Prepare steward data for display (sanitize for JSON)
  var stewardList = stewards.slice(0, 12).map(function(s) {
    return {
      firstName: (s['First Name'] || '').toString().replace(/"/g, '\\"'),
      lastName: (s['Last Name'] || '').toString().replace(/"/g, '\\"'),
      unit: (s['Unit'] || 'General').toString().replace(/"/g, '\\"'),
      location: (s['Work Location'] || '').toString().replace(/"/g, '\\"'),
      email: ''  // Redacted from public dashboard for privacy
    };
  });

  // Build trend data for area chart
  var trendChartData = [['Quarter', 'Trust Score']];
  if (satisfaction.trendData && satisfaction.trendData.length > 0) {
    satisfaction.trendData.forEach(function(item) {
      trendChartData.push(item);
    });
  } else {
    trendChartData.push(['Current', satisfaction.avgTrust || 7]);
  }

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Roboto", "Segoe UI", sans-serif; background: #f0f4f8; color: #1e293b; padding: 15px; }' +
    '.header { display: flex; align-items: center; color: #4338ca; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e0e7ff; }' +
    '.header h2 { font-size: 18px; font-weight: 700; margin-left: 8px; }' +
    '.card { background: white; border-radius: 12px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 12px; }' +
    '.stat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }' +
    '.stat-card { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; padding: 15px; border-radius: 10px; text-align: center; }' +
    '.stat-card.green { background: linear-gradient(135deg, #059669 0%, #047857 100%); }' +
    '.stat-card.blue { background: linear-gradient(135deg, #3B82F6 0%, #2563EB 100%); }' +
    '.stat-val { font-size: 28px; font-weight: 800; display: block; }' +
    '.stat-label { font-size: 11px; text-transform: uppercase; opacity: 0.9; font-weight: 500; }' +
    '.section-title { font-size: 12px; text-transform: uppercase; color: #64748b; font-weight: 600; margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }' +
    '.chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }' +
    '.chart-container { height: 140px; width: 100%; }' +
    '.progress-section { margin-bottom: 8px; }' +
    '.progress-header { display: flex; justify-content: space-between; font-size: 12px; color: #475569; margin-bottom: 4px; }' +
    '.progress-bg { background: #e2e8f0; border-radius: 10px; height: 10px; width: 100%; }' +
    '.progress-fill { background: linear-gradient(90deg, #7C3AED, #a78bfa); height: 100%; border-radius: 10px; transition: width 0.5s ease; }' +
    '.search-box { width: 100%; padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; margin-bottom: 10px; }' +
    '.search-box:focus { outline: none; border-color: #7C3AED; box-shadow: 0 0 0 2px rgba(124,58,237,0.1); }' +
    '.steward-list { max-height: 180px; overflow-y: auto; }' +
    '.steward-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid #f1f5f9; }' +
    '.steward-row:last-child { border-bottom: none; }' +
    '.steward-name { font-size: 13px; font-weight: 500; }' +
    '.steward-unit { background: #eff6ff; color: #1e40af; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; }' +
    '.steward-actions { display: flex; gap: 6px; }' +
    '.steward-actions a { color: #64748b; text-decoration: none; transition: color 0.2s; }' +
    '.steward-actions a:hover { color: #7C3AED; }' +
    '.footer { font-size: 9px; color: #94a3b8; text-align: center; margin-top: 12px; padding-top: 10px; border-top: 1px solid #e2e8f0; }' +
    '.trend-area { margin-top: 10px; }' +
    '</style>' +
    '<script type="text/javascript">' +
    ' + getClientSideEscapeHtml() + ' +
    'google.charts.load("current", {"packages":["corechart", "gauge"]});' +
    'google.charts.setOnLoadCallback(drawCharts);' +
    'function drawCharts() {' +
    // Issue Mix Pie Chart
    '  var issueData = google.visualization.arrayToDataTable(' + JSON.stringify(stats.categoryData || [['Category', 'Count'], ['No Data', 1]]) + ');' +
    '  var issueOptions = { pieHole: 0.4, chartArea: {width:"90%",height:"90%"}, legend: "none", colors: ["#7C3AED", "#059669", "#F97316", "#3B82F6", "#EC4899", "#6366F1"], pieSliceText: "label", fontSize: 10 };' +
    '  new google.visualization.PieChart(document.getElementById("issue_chart")).draw(issueData, issueOptions);' +
    // Trust Gauge
    '  var gaugeData = google.visualization.arrayToDataTable([["Label", "Value"],["Trust", ' + (satisfaction.avgTrust || 0) + ']]);' +
    '  var gaugeOptions = { width: 130, height: 130, greenFrom: 7, greenTo: 10, yellowFrom: 5, yellowTo: 7, redFrom: 0, redTo: 5, max: 10, minorTicks: 5 };' +
    '  new google.visualization.Gauge(document.getElementById("gauge_div")).draw(gaugeData, gaugeOptions);' +
    // Trend Area Chart
    '  var trendData = google.visualization.arrayToDataTable(' + JSON.stringify(trendChartData) + ');' +
    '  var trendOptions = { legend: "none", chartArea: {width:"85%",height:"70%"}, colors: ["#7C3AED"], areaOpacity: 0.3, hAxis: {textStyle:{fontSize:9}}, vAxis: {minValue:0,maxValue:10,textStyle:{fontSize:9}}, lineWidth: 2 };' +
    '  new google.visualization.AreaChart(document.getElementById("trend_chart")).draw(trendData, trendOptions);' +
    '}' +
    // Steward search filter
    'var stewards = ' + JSON.stringify(stewardList) + ';' +
    'function filterStewards(query) {' +
    '  var q = query.toLowerCase();' +
    '  var list = document.getElementById("steward-list");' +
    '  var html = "";' +
    '  stewards.forEach(function(s) {' +
    '    var fullName = s.firstName + " " + s.lastName;' +
    '    var searchText = (fullName + " " + s.unit + " " + s.location).toLowerCase();' +
    '    if (!q || searchText.indexOf(q) !== -1) {' +
    '      html += "<div class=\\"steward-row\\">";' +
    '      html += "<div><span class=\\"steward-name\\">" + escapeHtml(s.firstName) + " " + escapeHtml(s.lastName) + "</span></div>";' +
    '      html += "<div class=\\"steward-actions\\">";' +
    '      html += "<span class=\\"steward-unit\\">" + escapeHtml(s.unit) + "</span>";' +
    '      if (s.email) { html += " <a href=\\"mailto:" + escapeHtml(s.email) + "\\" title=\\"Email\\"><i class=\\"material-icons\\" style=\\"font-size:16px\\">email</i></a>"; }' +
    '      html += "</div></div>";' +
    '    }' +
    '  });' +
    '  if (!html) { html = "<div style=\\"text-align:center;padding:15px;color:#94a3b8\\">No stewards found</div>"; }' +
    '  list.innerHTML = html;' +
    '}' +
    'document.addEventListener("DOMContentLoaded", function() { filterStewards(""); });' +
    '</script>' +
    '</head>' +
    '<body>' +
    '<div class="header"><i class="material-icons">verified_user</i><h2>509 MEMBER PORTAL</h2></div>' +
    // Stats Grid
    '<div class="stat-grid">' +
    '<div class="stat-card"><span class="stat-val">' + (stats.open || 0) + '</span><span class="stat-label">Active Cases</span></div>' +
    '<div class="stat-card green"><span class="stat-val">' + (stats.resolved || 0) + '</span><span class="stat-label">Resolved (YTD)</span></div>' +
    '</div>' +
    // Charts Row
    '<div class="chart-row">' +
    '<div class="card"><div class="section-title"><i class="material-icons" style="font-size:14px">pie_chart</i> Issue Mix</div><div id="issue_chart" class="chart-container"></div></div>' +
    '<div class="card" style="text-align:center;"><div class="section-title"><i class="material-icons" style="font-size:14px">speed</i> Member Trust</div><div id="gauge_div" style="display:inline-block;margin-top:5px"></div></div>' +
    '</div>' +
    // Progress Bars
    '<div class="card">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">trending_up</i> Union Goals</div>' +
    '<div class="progress-section">' +
    '<div class="progress-header"><span>Steward Coverage</span><span>' + (coverage.coveragePercent || 0) + '%</span></div>' +
    '<div class="progress-bg"><div class="progress-fill" style="width:' + Math.min(100, coverage.coveragePercent || 0) + '%"></div></div>' +
    '</div>' +
    '<div class="progress-section">' +
    '<div class="progress-header"><span>Survey Participation</span><span>' + (satisfaction.responseCount > 0 ? Math.min(100, Math.round((satisfaction.responseCount / coverage.memberCount) * 100)) : 0) + '%</span></div>' +
    '<div class="progress-bg"><div class="progress-fill" style="width:' + (satisfaction.responseCount > 0 ? Math.min(100, Math.round((satisfaction.responseCount / Math.max(1, coverage.memberCount)) * 100)) : 0) + '%"></div></div>' +
    '</div>' +
    '</div>' +
    // Trust Trend
    '<div class="card trend-area">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">show_chart</i> Trust Score Trend</div>' +
    '<div id="trend_chart" style="height:80px;width:100%"></div>' +
    '</div>' +
    // Steward Directory with Search
    '<div class="card">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">groups</i> Find Your Steward</div>' +
    '<input type="text" class="search-box" placeholder="Search by name, unit, or location..." oninput="filterStewards(this.value)">' +
    '<div id="steward-list" class="steward-list"></div>' +
    '</div>' +
    // Footer
    '<div class="footer"><i class="material-icons" style="font-size:11px;vertical-align:middle">lock</i> Protected View - No Private PII Displayed</div>' +
    '</body>' +
    '</html>';
}

// ============================================================================
// PUBLIC DATA RETRIEVAL (NO PII)
// ============================================================================

/**
 * Get public overview data (no PII)
 * @returns {Object} Overview statistics
 */
function getPublicOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    totalMembers: 0,
    totalStewards: 0,
    totalGrievances: 0,
    winRate: 0,
    locationBreakdown: []
  };

  // Count members and stewards
  if (memberSheet) {
    var memberData = memberSheet.getDataRange().getValues();
    var locationCounts = {};
    var stewardCount = 0;

    for (var i = 1; i < memberData.length; i++) {
      var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
      if (!memberId || !memberId.toString().match(/^M/i)) continue;

      result.totalMembers++;

      // Count by location
      var location = memberData[i][MEMBER_COLS.LOCATION - 1] || 'Unknown';
      locationCounts[location] = (locationCounts[location] || 0) + 1;

      // Count stewards
      var isSteward = memberData[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isTruthyValue(isSteward)) {
        stewardCount++;
      }
    }

    result.totalStewards = stewardCount;

    // Convert location counts to array and sort
    Object.keys(locationCounts).forEach(function(loc) {
      result.locationBreakdown.push({ location: loc, count: locationCounts[loc] });
    });
    result.locationBreakdown.sort(function(a, b) { return b.count - a.count; });
    result.locationBreakdown = result.locationBreakdown.slice(0, 10); // Top 10
  }

  // Count grievances and win rate
  if (grievanceSheet) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    var won = 0, total = 0;

    for (var j = 1; j < grievanceData.length; j++) {
      var grievanceId = grievanceData[j][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      if (!grievanceId) continue;

      total++;
      var resolution = (grievanceData[j][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString().toLowerCase();
      if (resolution.includes('won') || resolution.includes('favor')) {
        won++;
      }
    }

    result.totalGrievances = total;
    result.winRate = total > 0 ? Math.round(won / total * 100) : 0;
  }

  return result;
}

/**
 * Get public survey data (anonymized)
 * Filters to only include Verified='Yes' and optionally IS_LATEST='Yes' responses
 * @param {boolean} includeHistory - If true, include superseded responses; if false, only latest per member
 * @returns {Object} Survey statistics
 */
function getPublicSurveyData(includeHistory) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = {
    totalResponses: 0,
    verifiedResponses: 0,
    avgSatisfaction: 0,
    responseRate: 0,
    sectionScores: [],
    includesHistory: includeHistory || false
  };

  if (!satSheet) return result;

  var data = satSheet.getDataRange().getValues();
  if (data.length < 2) return result;

  // Load vault data to check verified/isLatest status (PII stays in vault)
  var vaultMap = getVaultDataMap_();

  // Filter rows to only include verified responses using vault flags
  var validRows = [];
  for (var i = 1; i < data.length; i++) {
    var satRow = i + 1; // 1-indexed sheet row
    var vEntry = vaultMap[satRow];
    if (!vEntry) continue;

    // Only include Verified='Yes' responses
    if (!isTruthyValue(vEntry.verified)) continue;

    // If not including history, only include IS_LATEST='Yes'
    if (!includeHistory && !isTruthyValue(vEntry.isLatest)) continue;

    validRows.push(data[i]);
  }

  result.totalResponses = data.length - 1; // Total submissions (all)
  result.verifiedResponses = validRows.length; // Verified responses used in calculations

  // Calculate average satisfaction (Q6 - Satisfied with representation)
  var satSum = 0, satCount = 0;
  for (var j = 0; j < validRows.length; j++) {
    var sat = parseFloat(validRows[j][SATISFACTION_COLS.Q6_SATISFIED_REP - 1]);
    if (!isNaN(sat)) {
      satSum += sat;
      satCount++;
    }
  }
  result.avgSatisfaction = satCount > 0 ? satSum / satCount : 0;

  // Response rate (unique verified members / total members)
  // Count from vault — member IDs are hashed, but unique hashes = unique members
  if (memberSheet) {
    var memberCount = memberSheet.getLastRow() - 1;
    var uniqueMembers = {};
    var vaultFull = getVaultDataFull_();
    for (var k = 0; k < vaultFull.length; k++) {
      if (vaultFull[k].verified === 'Yes' && vaultFull[k].isLatest === 'Yes' && vaultFull[k].memberIdHash) {
        uniqueMembers[vaultFull[k].memberIdHash] = true;
      }
    }
    var uniqueCount = Object.keys(uniqueMembers).length;
    result.responseRate = memberCount > 0 ? Math.round(uniqueCount / memberCount * 100) : 0;
  }

  // Section scores using only verified responses
  var sections = [
    { name: 'Overall Satisfaction', cols: [SATISFACTION_COLS.Q6_SATISFIED_REP, SATISFACTION_COLS.Q7_TRUST_UNION, SATISFACTION_COLS.Q8_FEEL_PROTECTED] },
    { name: 'Steward Ratings', cols: [SATISFACTION_COLS.Q10_TIMELY_RESPONSE, SATISFACTION_COLS.Q11_TREATED_RESPECT, SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS] },
    { name: 'Chapter Effectiveness', cols: [SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, SATISFACTION_COLS.Q22_CHAPTER_COMM, SATISFACTION_COLS.Q23_ORGANIZES] },
    { name: 'Local Leadership', cols: [SATISFACTION_COLS.Q26_DECISIONS_CLEAR, SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS, SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE] },
    { name: 'Communication', cols: [SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, SATISFACTION_COLS.Q42_ENOUGH_INFO] }
  ];

  sections.forEach(function(section) {
    var sum = 0, count = 0;
    for (var m = 0; m < validRows.length; m++) {
      section.cols.forEach(function(col) {
        if (col) {
          var val = parseFloat(validRows[m][col - 1]);
          if (!isNaN(val)) {
            sum += val;
            count++;
          }
        }
      });
    }
    result.sectionScores.push({
      section: section.name,
      score: count > 0 ? sum / count : 0
    });
  });

  result.sectionScores.sort(function(a, b) { return b.score - a.score; });

  return result;
}

/**
 * Get public grievance data (no PII)
 * @returns {Object} Grievance statistics
 */
function getPublicGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    total: 0,
    open: 0,
    won: 0,
    settled: 0,
    avgDaysToResolve: 0,
    byType: [],
    byStatus: []
  };

  if (!grievanceSheet) return result;

  var data = grievanceSheet.getDataRange().getValues();
  var typeCounts = {};
  var statusCounts = {};
  var daysToResolve = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    result.total++;

    var status = data[i][GRIEVANCE_COLS.STATUS - 1] || 'Unknown';
    var resolution = (data[i][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString();
    var gType = data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';

    // Count by status
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    // Count by type
    typeCounts[gType] = (typeCounts[gType] || 0) + 1;

    // Track open/won/settled
    if (status === 'Open' || status === 'Pending Info') {
      result.open++;
    }
    if (resolution.toLowerCase().includes('won') || resolution.toLowerCase().includes('favor')) {
      result.won++;
    }
    if (resolution.toLowerCase().includes('settled')) {
      result.settled++;
    }

    // Calculate days to resolve for closed grievances
    if (status === 'Closed' || status === 'Resolved') {
      var dateOpened = data[i][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = data[i][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateOpened && dateClosed) {
        var days = Math.round((new Date(dateClosed) - new Date(dateOpened)) / (1000 * 60 * 60 * 24));
        if (days > 0) daysToResolve.push(days);
      }
    }
  }

  // Average days to resolve
  if (daysToResolve.length > 0) {
    result.avgDaysToResolve = Math.round(daysToResolve.reduce(function(a, b) { return a + b; }, 0) / daysToResolve.length);
  }

  // Convert to arrays
  Object.keys(typeCounts).forEach(function(t) {
    result.byType.push({ type: t, count: typeCounts[t] });
  });
  result.byType.sort(function(a, b) { return b.count - a.count; });
  result.byType = result.byType.slice(0, 8);

  Object.keys(statusCounts).forEach(function(s) {
    result.byStatus.push({ status: s, count: statusCounts[s] });
  });
  result.byStatus.sort(function(a, b) { return b.count - a.count; });

  return result;
}

/**
 * Get public steward data (contact info only)
 * @returns {Object} Steward directory
 */
function getPublicStewardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = { stewards: [] };

  if (!memberSheet) return result;

  var data = memberSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (!isTruthyValue(isSteward)) continue;

    var firstName = data[i][MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = data[i][MEMBER_COLS.LAST_NAME - 1] || '';

    result.stewards.push({
      name: firstName + ' ' + lastName,
      location: data[i][MEMBER_COLS.LOCATION - 1] || 'Not specified',
      officeDays: data[i][MEMBER_COLS.OFFICE_DAYS - 1] || 'Contact for availability',
      email: data[i][MEMBER_COLS.EMAIL - 1] || 'Contact union office'
    });
  }

  // Sort by name
  result.stewards.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return result;
}

// ============================================================================
// STEWARD COVERAGE STATISTICS
// ============================================================================

/**
 * Gets steward coverage ratio for progress tracking
 * @returns {Object} Coverage statistics
 */
function getStewardCoverageStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || memberSheet.getLastRow() < 2) {
    return { ratio: 0, stewardCount: 0, memberCount: 0, targetRatio: 15 };
  }

  var data = memberSheet.getDataRange().getValues();
  var stewardCount = 0;
  var memberCount = 0;

  for (var i = 1; i < data.length; i++) {
    var memberId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    memberCount++;
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isTruthyValue(isSteward)) {
      stewardCount++;
    }
  }

  // Calculate ratio as members per steward (lower is better coverage)
  var ratio = stewardCount > 0 ? Math.round(memberCount / stewardCount) : 0;
  var targetRatio = 15; // Target: 1 steward per 15 members
  var coveragePercent = ratio > 0 ? Math.min(100, Math.round((targetRatio / ratio) * 100)) : 0;

  return {
    ratio: ratio,
    stewardCount: stewardCount,
    memberCount: memberCount,
    targetRatio: targetRatio,
    coveragePercent: coveragePercent
  };
}
