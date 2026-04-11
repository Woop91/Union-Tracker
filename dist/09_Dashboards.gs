/**
 * ============================================================================
 * 09_Dashboards.gs — Member Satisfaction Dashboard
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Member Satisfaction Dashboard display and analytics. showSatisfactionDashboard()
 *   opens a modal with survey results, section scores, trend analysis, and
 *   worksite/role breakdowns. getSatisfactionDashboardHtml() generates the full
 *   HTML with embedded charts (Chart.js). Calculates section-level scores across
 *   6 categories and generates plain-language insights.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Modal dialog approach (not a separate sheet) because satisfaction data is
 *   sensitive — it shouldn't be visible as a sheet tab. Chart.js is embedded
 *   because GAS dialogs can't load external scripts reliably. Green theme
 *   differentiates this from the navy Steward Dashboard.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The satisfaction dashboard menu item shows an empty or errored modal.
 *   Stewards lose visibility into member satisfaction trends. Section scores
 *   and insights are unavailable for bargaining preparation.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, getMobileOptimizedHead),
 *               08e_SurveyEngine.gs (survey data),
 *               04b_AccessibilityFeatures.gs (getCommonStyles)
 *   Used by:    Menu items in 03_UIComponents.gs
 *
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// SATISFACTION DASHBOARD - MODAL DISPLAY
// ============================================================================

/**
 * Show the Member Satisfaction Dashboard modal
 */
function showSatisfactionDashboard() {
  showDialog_(getSatisfactionDashboardHtml(), '📊 Member Satisfaction', 900, 750);
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
    getClientSideEscapeHtml() +
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
    '  google.script.run.withFailureHandler(function(e){showError&&showError(e.message)}).withSuccessHandler(function(data){renderOverview(data)}).getSatisfactionOverviewData();' +
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
    '  return"<div class=\\"gauge\\"><div class=\\"gauge-ring\\" style=\\"background:conic-gradient("+color+" "+pct+"%,#e5e7eb "+pct+"%)\\"><span style=\\"color:"+color+"\\">"+value.toFixed(1)+"</span></div><div class=\\"gauge-label\\">"+escapeHtml(label).replace(/\\n/g,"<br>")+"</div></div>";' +
    '}' +

    // Load responses
    'function loadResponses(){' +
    '  google.script.run.withFailureHandler(function(e){showError&&showError(e.message)}).withSuccessHandler(function(data){allResponses=data;renderResponses(data)}).getSatisfactionResponseData();' +
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
    '  google.script.run.withFailureHandler(function(e){showError&&showError(e.message)}).withSuccessHandler(function(data){renderSections(data)}).getSatisfactionSectionData();' +
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
    '  google.script.run.withFailureHandler(function(e){showError&&showError(e.message)}).withSuccessHandler(function(data){renderAnalytics(data)}).getSatisfactionAnalyticsData();' +
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
// SATISFACTION COLUMN CACHE
// ============================================================================

/**
 * Module-level cache for buildSatisfactionColsShim_ result.
 * Avoids rebuilding the column map on every function call within a single execution.
 * @private
 */
var _satisfactionColsCache_ = null;
function _getCachedSatisfactionCols() {
  if (!_satisfactionColsCache_) {
    _satisfactionColsCache_ = buildSatisfactionColsShim_(getSatisfactionColMap_());
  }
  return _satisfactionColsCache_;
}

// ============================================================================
// SATISFACTION DATA RETRIEVAL - OVERVIEW AND STATS
// ============================================================================

/**
 * Get overview data for satisfaction dashboard
 */
function getSatisfactionOverviewData() {
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (!sheet) return {
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

  // Check if there's data by looking at column A (Timestamp)
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return data;

  data.totalResponses = lastRow - 1;

  // Get satisfaction scores (Q6-Q9) — dynamic range from SATISFACTION_COLS
  var satRangeStart = SATISFACTION_COLS.Q6_SATISFIED_REP;
  var satRangeCount = SATISFACTION_COLS.Q9_RECOMMEND - satRangeStart + 1;
  var satisfactionRange = sheet.getRange(2, satRangeStart, data.totalResponses, satRangeCount).getValues();

  var sumOverall = 0, sumTrust = 0, sumProtected = 0, sumRecommend = 0;
  var promoters = 0, detractors = 0;
  var validCount = 0;

  satisfactionRange.forEach(function(row) {
    var satisfied = parseFloat(row[SATISFACTION_COLS.Q6_SATISFIED_REP - satRangeStart]) || 0;
    var trust = parseFloat(row[SATISFACTION_COLS.Q7_TRUST_UNION - satRangeStart]) || 0;
    var protected_ = parseFloat(row[SATISFACTION_COLS.Q8_FEEL_PROTECTED - satRangeStart]) || 0;
    var recommend = parseFloat(row[SATISFACTION_COLS.Q9_RECOMMEND - satRangeStart]) || 0;

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

  // Get steward ratings (Q10-Q16) — guard against col=0 when schema not yet initialized
  var stewardRange = (SATISFACTION_COLS.Q10_TIMELY_RESPONSE > 0)
    ? sheet.getRange(2, SATISFACTION_COLS.Q10_TIMELY_RESPONSE, data.totalResponses, 7).getValues()
    : [];
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

  // Get leadership ratings (Q26-Q31) — guard against col=0 when schema not yet initialized
  var leadershipRange = (SATISFACTION_COLS.Q26_DECISIONS_CLEAR > 0)
    ? sheet.getRange(2, SATISFACTION_COLS.Q26_DECISIONS_CLEAR, data.totalResponses, 6).getValues()
    : [];
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
// ============================================================================
// SATISFACTION DATA RETRIEVAL - RESPONSES AND SECTIONS
// ============================================================================

/**
 * Get individual response data for satisfaction dashboard
 */
function getSatisfactionResponseData() {
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (!sheet) return [];

  // Check if there's data
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  var numRows = lastRow - 1;
  var tz = Session.getScriptTimeZone();

  // Read all data as a single batch — avoids col=0 errors when schema not yet initialized
  var _allData = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  function _col1(colIdx) { return _allData.map(function(r) { return [colIdx > 0 ? (r[colIdx-1] || '') : '']; }); }
  function _colN(colIdx, n) { return _allData.map(function(r) { return colIdx > 0 ? r.slice(colIdx-1, colIdx-1+n) : new Array(n).fill(''); }); }
  var timestampData       = _col1(1);
  var worksiteData        = _col1(SATISFACTION_COLS.Q1_WORKSITE);
  var roleData            = _col1(SATISFACTION_COLS.Q2_ROLE);
  var shiftData           = _col1(SATISFACTION_COLS.Q3_SHIFT);
  var timeData            = _col1(SATISFACTION_COLS.Q4_TIME_IN_ROLE);
  var stewardContactData  = _col1(SATISFACTION_COLS.Q5_STEWARD_CONTACT);
  var satisfactionData    = _colN(SATISFACTION_COLS.Q6_SATISFIED_REP, 4);
  var stewardRatingsData  = _colN(SATISFACTION_COLS.Q10_TIMELY_RESPONSE, 7);

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
    var da = new Date(a.date).getTime();
    var db = new Date(b.date).getTime();
    if (isNaN(db)) return -1;
    if (isNaN(da)) return 1;
    return db - da;
  });

  return responses;
}

/**
 * Get section-level data for satisfaction dashboard
 */
function getSatisfactionSectionData() {
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
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
    if (!section.startCol || section.startCol < 1) {
      result.sections.push({ name: section.name, avg: 0, responseCount: 0, questions: section.numCols });
      return; // skip this iteration (forEach callback)
    }
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
/**
 * Get analytics data for satisfaction dashboard insights
 */
function getSatisfactionAnalyticsData() {
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
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

  // Read all data as single batch — avoids col=0 errors when schema not yet initialized
  var _allData5 = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  var worksiteData       = _allData5.map(function(r){var c=SATISFACTION_COLS.Q1_WORKSITE;       return [c>0?r[c-1]:''];});
  var roleData           = _allData5.map(function(r){var c=SATISFACTION_COLS.Q2_ROLE;           return [c>0?r[c-1]:''];});
  var stewardContactData = _allData5.map(function(r){var c=SATISFACTION_COLS.Q5_STEWARD_CONTACT;return [c>0?r[c-1]:''];});
  var satisfactionData   = _allData5.map(function(r){var c=SATISFACTION_COLS.Q6_SATISFIED_REP;  return c>0?r.slice(c-1,c+3):['','','',''];});
  var prioritiesData     = _allData5.map(function(r){var c=SATISFACTION_COLS.Q64_TOP_PRIORITIES;return [c>0?r[c-1]:''];});

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
  for (i = 0; i < numRows; i++) {
    var ws = worksiteData[i][0] || 'Unknown';
    if (!worksiteMap[ws]) worksiteMap[ws] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      worksiteMap[ws].sum += scores[i];
      worksiteMap[ws].count++;
    }
  }

  for (ws in worksiteMap) {
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
  for (i = 0; i < numRows; i++) {
    var role = roleData[i][0] || 'Unknown';
    if (!roleMap[role]) roleMap[role] = { sum: 0, count: 0 };
    if (scores[i] > 0) {
      roleMap[role].sum += scores[i];
      roleMap[role].count++;
    }
  }

  for (role in roleMap) {
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

  for (i = 0; i < numRows; i++) {
    var contact = stewardContactData[i][0];
    if (scores[i] > 0) {
      if (isTruthyValue(contact)) {
        withContactSum += scores[i];
        withContactCount++;
      } else if (String(contact).trim().toLowerCase() === 'no') {
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
  for (i = 0; i < numRows; i++) {
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
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { log_('syncSatisfactionValues', 'lock contention — skipped'); return; }
  try {
    var SATISFACTION_COLS = _getCachedSatisfactionCols();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

    if (!sheet) {
      log_('syncSatisfactionValues', 'Member Satisfaction sheet not found');
      return;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      // No data to process, just write empty dashboard
      writeSatisfactionDashboard_(sheet, [], []);
      return;
    }

    // Get all response data (up to last scale question column)
    var lastResponseCol = SATISFACTION_COLS.Q62_CONCERNS_SERIOUS; // last Likert-scale question
    var responseData = sheet.getRange(2, 1, lastRow - 1, lastResponseCol).getValues();

    // Calculate section averages for each row
    var sectionAverages = computeSectionAverages_(responseData);

    // Write section averages to summary area (v4.23.0: skipped — dynamic schema has no summary cols)
    var summaryStart = SATISFACTION_COLS.SUMMARY_START;
    var summaryCols = Math.max(1, (SATISFACTION_COLS.AVG_SCHEDULING || 0) - (SATISFACTION_COLS.AVG_OVERALL_SAT || 0) + 1);
    if (sectionAverages.length > 0 && summaryStart > 0) {
      sheet.getRange(2, summaryStart, sectionAverages.length, summaryCols).setValues(sectionAverages);
    }

    // Calculate and write dashboard metrics
    writeSatisfactionDashboard_(sheet, responseData, sectionAverages);

    log_('syncSatisfactionValues', 'Member Satisfaction values synced for ' + responseData.length + ' responses');
  } finally { lock.releaseLock(); }
}

/**
 * Compute section averages for satisfaction survey rows
 * @param {Array} responseData - 2D array of survey response data
 * @return {Array} 2D array of section averages (11 columns per row)
 * @private
 */
function computeSectionAverages_(responseData) {
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
  var results = [];

  for (var r = 0; r < responseData.length; r++) {
    var row = responseData[r];
    if (!row[0]) continue; // Skip empty rows

    var averages = [];

    // Derive indices from SATISFACTION_COLS (1-indexed, used as 0-indexed array indices)
    // Overall Satisfaction (Q6-9)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q6_SATISFIED_REP - 1, SATISFACTION_COLS.Q9_RECOMMEND - 1));

    // Steward Rating (Q10-16)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q10_TIMELY_RESPONSE - 1, SATISFACTION_COLS.Q16_CONFIDENTIALITY - 1));

    // Steward Access (Q18-20)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q18_KNOW_CONTACT - 1, SATISFACTION_COLS.Q20_EASY_FIND - 1));

    // Chapter (Q21-25)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES - 1, SATISFACTION_COLS.Q25_FAIR_REP - 1));

    // Leadership (Q26-31)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q26_DECISIONS_CLEAR - 1, SATISFACTION_COLS.Q31_WELCOMES_OPINIONS - 1));

    // Contract (Q32-35)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q32_ENFORCES_CONTRACT - 1, SATISFACTION_COLS.Q35_FRONTLINE_PRIORITY - 1));

    // Representation (Q37-40)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS - 1, SATISFACTION_COLS.Q40_OUTCOME_JUSTIFIED - 1));

    // Communication (Q41-45)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE - 1, SATISFACTION_COLS.Q45_MEETINGS_WORTH - 1));

    // Member Voice (Q46-50)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q46_VOICE_MATTERS - 1, SATISFACTION_COLS.Q50_CONFLICT_RESPECT - 1));

    // Value/Action (Q51-55)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q51_GOOD_VALUE - 1, SATISFACTION_COLS.Q55_WIN_TOGETHER - 1));

    // Scheduling (Q56-62)
    averages.push(computeAverage_(row, SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES - 1, SATISFACTION_COLS.Q62_CONCERNS_SERIOUS - 1));

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
    if (val === '' || val == null) continue;
    // Accept numeric cells AND string-numeric cells. Form submissions can
    // import numeric answers as text; the previous typeof-only check
    // silently dropped every string "5" and undercounted averages.
    var n = (typeof val === 'number') ? val : Number(val);
    if (!isNaN(n)) values.push(n);
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
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
  var dashStart = 84; // Column CF
  var demoStart = 87; // Column CH
  var chartStart = 90; // Column CK

  // Calculate aggregate metrics
  var totalResponses = responseData.length;
  var responsePeriod = 'No data';
  if (totalResponses > 0) {
    var timestamps = responseData.map(function(r) { return col_(r, SATISFACTION_COLS.TIMESTAMP); }).filter(function(t) { return t instanceof Date; });
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

    // Shift (Q3)
    var shift = col_(row, SATISFACTION_COLS.Q3_SHIFT);
    if (shift === 'Day') shifts.Day++;
    else if (shift === 'Evening') shifts.Evening++;
    else if (shift === 'Night') shifts.Night++;
    else if (shift === 'Rotating') shifts.Rotating++;

    // Tenure (Q4)
    var ten = String(col_(row, SATISFACTION_COLS.Q4_TIME_IN_ROLE) || '');
    if (ten.indexOf('<1') >= 0) tenure['<1']++;
    else if (ten.indexOf('1-3') >= 0) tenure['1-3']++;
    else if (ten.indexOf('4-7') >= 0) tenure['4-7']++;
    else if (ten.indexOf('8-15') >= 0) tenure['8-15']++;
    else if (ten.indexOf('15+') >= 0) tenure['15+']++;

    // Steward contact (Q5)
    var q5Val = col_(row, SATISFACTION_COLS.Q5_STEWARD_CONTACT);
    if (isTruthyValue(q5Val)) stewardContact.Yes++;
    else if (String(q5Val).trim().toLowerCase() === 'no') stewardContact.No++;

    // Filed grievance (Q36)
    var q36Val = col_(row, SATISFACTION_COLS.Q36_FILED_GRIEVANCE);
    if (isTruthyValue(q36Val)) filedGrievance.Yes++;
    else if (String(q36Val).trim().toLowerCase() === 'no') filedGrievance.No++;
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
  var SATISFACTION_COLS = _getCachedSatisfactionCols();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!sheet || row < 2) return;

  // Get the response data for this row
  var lastResponseCol = SATISFACTION_COLS.Q62_CONCERNS_SERIOUS;
  var rowData = sheet.getRange(row, 1, 1, lastResponseCol).getValues()[0];

  if (!col_(rowData, SATISFACTION_COLS.TIMESTAMP)) return; // Skip if no timestamp

  var averages = [];

  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q6_SATISFIED_REP - 1, SATISFACTION_COLS.Q9_RECOMMEND - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q10_TIMELY_RESPONSE - 1, SATISFACTION_COLS.Q16_CONFIDENTIALITY - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q18_KNOW_CONTACT - 1, SATISFACTION_COLS.Q20_EASY_FIND - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES - 1, SATISFACTION_COLS.Q25_FAIR_REP - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q26_DECISIONS_CLEAR - 1, SATISFACTION_COLS.Q31_WELCOMES_OPINIONS - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q32_ENFORCES_CONTRACT - 1, SATISFACTION_COLS.Q35_FRONTLINE_PRIORITY - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS - 1, SATISFACTION_COLS.Q40_OUTCOME_JUSTIFIED - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE - 1, SATISFACTION_COLS.Q45_MEETINGS_WORTH - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q46_VOICE_MATTERS - 1, SATISFACTION_COLS.Q50_CONFLICT_RESPECT - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q51_GOOD_VALUE - 1, SATISFACTION_COLS.Q55_WIN_TOGETHER - 1));
  averages.push(computeAverage_(rowData, SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES - 1, SATISFACTION_COLS.Q62_CONCERNS_SERIOUS - 1));

  // Write section averages to summary area (v4.23.0: skipped — dynamic schema has no summary cols)
  var summaryStart = SATISFACTION_COLS.SUMMARY_START;
  var summaryCols = Math.max(1, (SATISFACTION_COLS.AVG_SCHEDULING || 0) - (SATISFACTION_COLS.AVG_OVERALL_SAT || 0) + 1);
  if (summaryStart > 0) {
    sheet.getRange(row, summaryStart, 1, summaryCols).setValues([averages]);
  }
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
 *
 * Dependencies:
 * - 00_Config.gs (SHEETS, GRIEVANCE_COLS, MEMBER_COLS, CONFIG_COLS, COLORS)
 * - 01_Utilities.gs (getColumnLetter, getConfigValues, getJobMetadataByMemberCol)
 *
 * @author Claude Code Assistant
 * @version 4.51.0
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
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Required sheets not found for grievance sync.', 'Sync Error', 5); } catch (_) {}
    return;
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  // M-26: Closed statuses - grievances with these statuses don't count as "open"
  var closedStatuses = GRIEVANCE_CLOSED_STATUSES;

  // Build lookup map: memberId -> {hasOpen, status, deadline}
  // Calculate directly from grievance data (handles "Overdue" text properly)
  var lookup = {};

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var memberId = col_(row, GRIEVANCE_COLS.MEMBER_ID);
    if (!memberId) continue;

    var status = col_(row, GRIEVANCE_COLS.STATUS) || '';
    var daysToDeadline = col_(row, GRIEVANCE_COLS.DAYS_TO_DEADLINE);
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
      if (status === GRIEVANCE_STATUS.OPEN) {
        lookup[memberId].status = GRIEVANCE_STATUS.OPEN;
      } else if (status === GRIEVANCE_STATUS.PENDING && lookup[memberId].status !== GRIEVANCE_STATUS.OPEN) {
        lookup[memberId].status = GRIEVANCE_STATUS.PENDING;
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

  // Resolve target columns from actual sheet headers at write time (v4.51.1)
  // This prevents data corruption when columns are appended out of canonical order
  // or when CacheService-backed MEMBER_COLS constants go stale.
  var writeCols = resolveColumnsByHeader_(memberSheet, [
    { key: 'HAS_OPEN',  header: 'Has Open Grievance?', fallback: MEMBER_COLS.HAS_OPEN_GRIEVANCE },
    { key: 'STATUS',    header: 'Grievance Status',     fallback: MEMBER_COLS.GRIEVANCE_STATUS },
    { key: 'DEADLINE',  header: 'Days to Deadline',     fallback: MEMBER_COLS.DAYS_TO_DEADLINE }
  ]);

  // Verify columns are actually consecutive — the old code assumed 3 adjacent columns.
  // If they're not consecutive, we must write each column individually.
  var isConsecutive = (writeCols.STATUS === writeCols.HAS_OPEN + 1) &&
                      (writeCols.DEADLINE === writeCols.STATUS + 1);

  // Build update arrays from member data + grievance lookup
  var updates = [];
  for (var j = 1; j < memberData.length; j++) {
    memberId = col_(memberData[j], MEMBER_COLS.MEMBER_ID);
    var memberInfo = lookup[memberId] || {hasOpen: 'No', status: '', deadline: ''};
    updates.push([memberInfo.hasOpen, memberInfo.status, memberInfo.deadline]);
  }

  if (updates.length > 0) {
    if (isConsecutive) {
      // Fast path: single batch write when columns are adjacent
      memberSheet.getRange(2, writeCols.HAS_OPEN, updates.length, 3).setValues(updates);
    } else {
      // Safe path: write each column individually when columns are not adjacent
      var hasOpenArr = [], statusArr = [], deadlineArr = [];
      for (var u = 0; u < updates.length; u++) {
        hasOpenArr.push([updates[u][0]]);
        statusArr.push([updates[u][1]]);
        deadlineArr.push([updates[u][2]]);
      }
      memberSheet.getRange(2, writeCols.HAS_OPEN, updates.length, 1).setValues(hasOpenArr);
      memberSheet.getRange(2, writeCols.STATUS, updates.length, 1).setValues(statusArr);
      memberSheet.getRange(2, writeCols.DEADLINE, updates.length, 1).setValues(deadlineArr);
    }
  }

  log_('syncGrievanceToMemberDirectory', 'Synced grievance data to ' + updates.length + ' members (cols resolved: ' +
    writeCols.HAS_OPEN + '/' + writeCols.STATUS + '/' + writeCols.DEADLINE + ')');
}

/**
 * Sync calculated formulas from hidden sheet to Grievance Log
 * This is the self-healing function - it copies calculated values to the Grievance Log
 * Member data (Name, Email, Unit, Location, Steward) is looked up directly from Member Directory
 */
/**
 * Build a lookup map from Member Directory data, keyed by Member ID.
 * @param {Array<Array>} memberData - Raw member data with headers at index 0
 * @returns {Object} Map of memberId -> { firstName, lastName, email, unit, location, steward }
 * @private
 */
function _buildMemberLookupMap_(memberData) {
  var memberLookup = {};
  for (var i = 1; i < memberData.length; i++) {
    var memberId = col_(memberData[i], MEMBER_COLS.MEMBER_ID);
    if (memberId) {
      memberLookup[memberId] = {
        firstName: col_(memberData[i], MEMBER_COLS.FIRST_NAME) || '',
        lastName: col_(memberData[i], MEMBER_COLS.LAST_NAME) || '',
        email: col_(memberData[i], MEMBER_COLS.EMAIL) || '',
        unit: col_(memberData[i], MEMBER_COLS.UNIT) || '',
        location: col_(memberData[i], MEMBER_COLS.WORK_LOCATION) || '',
        steward: col_(memberData[i], MEMBER_COLS.ASSIGNED_STEWARD) || ''
      };
    }
  }
  return memberLookup;
}

/**
 * Calculate deadline dates and time-based metrics for a single grievance row.
 * Pure function — no SpreadsheetApp calls.
 * @param {Array} rowData - Single grievance row array
 * @param {Object} rules - Deadline rules from getDeadlineRules()
 * @param {Array<string>} closedStatuses - Statuses that should not have Next Action Due
 * @param {Date} today - Current date (normalized to midnight)
 * @returns {Object} { deadlines: [5 values], metrics: [3 values] }
 * @private
 */
function _calculateRowDeadlines_(rowData, rules, closedStatuses, today) {
  var incidentDate = col_(rowData, GRIEVANCE_COLS.INCIDENT_DATE);
  var dateFiled = col_(rowData, GRIEVANCE_COLS.DATE_FILED);
  var step1Rcvd = col_(rowData, GRIEVANCE_COLS.STEP1_RCVD);
  var step2AppealFiled = col_(rowData, GRIEVANCE_COLS.STEP2_APPEAL_FILED);
  var step2Rcvd = col_(rowData, GRIEVANCE_COLS.STEP2_RCVD);
  var dateClosed = col_(rowData, GRIEVANCE_COLS.DATE_CLOSED);
  var status = col_(rowData, GRIEVANCE_COLS.STATUS);
  var currentStep = col_(rowData, GRIEVANCE_COLS.CURRENT_STEP);

  // Calculate deadline dates using configurable rules and business-day math
  // (matches recalculateDownstreamDeadlines_ approach)
  var filingDeadline = '';
  var step1Due = '';
  var step2AppealDue = '';
  var step2Due = '';
  var step3AppealDue = '';

  if (incidentDate instanceof Date) {
    filingDeadline = addCalendarDays(incidentDate, rules.FILING_DAYS);
  }
  if (dateFiled instanceof Date) {
    step1Due = addCalendarDays(dateFiled, rules.STEP_1.DAYS_FOR_RESPONSE);
  }
  if (step1Rcvd instanceof Date) {
    step2AppealDue = addBusinessDays(step1Rcvd, rules.STEP_2.DAYS_TO_APPEAL); // Art. 23: business days for appeals
  }
  if (step2AppealFiled instanceof Date) {
    step2Due = addCalendarDays(step2AppealFiled, rules.STEP_2.DAYS_FOR_RESPONSE);
  }
  if (step2Rcvd instanceof Date) {
    step3AppealDue = addBusinessDays(step2Rcvd, rules.STEP_3.DAYS_TO_APPEAL); // Art. 23: business days for appeals
  }

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
    daysToDeadline = Math.floor((nextActionDue - today) / (1000 * 60 * 60 * 24));
  }

  return {
    deadlines: [filingDeadline, step1Due, step2AppealDue, step2Due, step3AppealDue],
    metrics: [daysOpen, nextActionDue, daysToDeadline]
  };
}

/**
 * Track data quality issues for a single grievance row.
 * Appends to the provided quality counter arrays (mutates in place).
 * @param {string} memberId - The member ID from the grievance row
 * @param {string} grievanceId - The grievance ID (or fallback label)
 * @param {Object} memberLookup - Member lookup map
 * @param {Object} qualityCounters - { orphanedGrievances: [], missingMemberIds: [] }
 * @private
 */
function _trackDataQuality_(memberId, grievanceId, memberLookup, qualityCounters) {
  if (!memberId) {
    qualityCounters.missingMemberIds.push(grievanceId);
    log_('WARNING', 'Grievance ' + grievanceId + ' has no Member ID');
  } else if (!memberLookup[memberId]) {
    qualityCounters.orphanedGrievances.push(grievanceId + ' (Member ID: ' + memberId + ')');
    log_('WARNING', 'Grievance ' + grievanceId + ' references non-existent Member ID: ' + memberId);
  }
}

function syncGrievanceFormulasToLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Required sheets not found for formula sync.', 'Sync Error', 5); } catch (_) {}
    return;
  }

  // Build member lookup from directory
  var memberData = memberSheet.getDataRange().getValues();
  var memberLookup = _buildMemberLookupMap_(memberData);

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to start of day

  // M-26: Closed statuses that should not have Next Action Due
  var closedStatuses = GRIEVANCE_CLOSED_STATUSES;

  // Prepare updates
  var nameUpdates = [];           // Columns C-D
  var deadlineUpdates = [];       // Columns H, J, L, N, P
  var metricsUpdates = [];        // Columns S, T, U
  var contactUpdates = [];        // Columns X, Y, Z, AA

  // Track data quality issues
  var qualityCounters = { orphanedGrievances: [], missingMemberIds: [] };

  // Read configurable deadline rules once (outside the loop)
  var rules = getDeadlineRules();

  for (var j = 1; j < grievanceData.length; j++) {
    var row = grievanceData[j];
    var memberId = col_(row, GRIEVANCE_COLS.MEMBER_ID);
    var grievanceId = col_(row, GRIEVANCE_COLS.GRIEVANCE_ID) || ('Row ' + (j + 1));

    // Track data quality issues
    _trackDataQuality_(memberId, grievanceId, memberLookup, qualityCounters);

    var memberInfo = memberLookup[memberId] || {};

    // Names (C-D) - from Member Directory
    nameUpdates.push([
      memberInfo.firstName || '',
      memberInfo.lastName || ''
    ]);

    // Calculate deadlines and metrics for this row
    var rowCalc = _calculateRowDeadlines_(row, rules, closedStatuses, today);

    // Deadlines (H, J, L, N, P)
    deadlineUpdates.push(rowCalc.deadlines);

    // Metrics (S, T, U)
    metricsUpdates.push(rowCalc.metrics);

    // Contact info (X, Y, Z)
    contactUpdates.push([
      memberInfo.email || '',
      memberInfo.location || '',
      memberInfo.steward || ''
    ]);
  }

  // Apply updates to Grievance Log
  if (nameUpdates.length > 0) {
    // C-D: First Name, Last Name — resolve from headers at write time (v4.51.1)
    var gNameCols = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'FIRST', header: 'First Name', fallback: GRIEVANCE_COLS.FIRST_NAME },
      { key: 'LAST',  header: 'Last Name',  fallback: GRIEVANCE_COLS.LAST_NAME }
    ]);
    // CR-10: Before writing names and contact info, check existing cell data.
    // Never overwrite a non-empty cell with an empty string (would blank manually entered data).
    // Read existing names using header-resolved positions (v4.51.1)
    var enFirst = grievanceSheet.getRange(2, gNameCols.FIRST, nameUpdates.length, 1).getValues();
    var enLast = grievanceSheet.getRange(2, gNameCols.LAST, nameUpdates.length, 1).getValues();
    var existingNames = [];
    for (var en = 0; en < nameUpdates.length; en++) { existingNames.push([enFirst[en][0], enLast[en][0]]); }
    for (var ni = 0; ni < nameUpdates.length; ni++) {
      for (var nc = 0; nc < 2; nc++) {
        // Only write if: new value is non-empty, OR existing cell is empty
        if (!nameUpdates[ni][nc] && existingNames[ni][nc]) {
          nameUpdates[ni][nc] = existingNames[ni][nc];
        }
      }
    }
    if (gNameCols.LAST === gNameCols.FIRST + 1) {
      grievanceSheet.getRange(2, gNameCols.FIRST, nameUpdates.length, 2).setValues(nameUpdates);
    } else {
      var fnArr = [], lnArr = [];
      for (var na = 0; na < nameUpdates.length; na++) { fnArr.push([nameUpdates[na][0]]); lnArr.push([nameUpdates[na][1]]); }
      grievanceSheet.getRange(2, gNameCols.FIRST, nameUpdates.length, 1).setValues(fnArr);
      grievanceSheet.getRange(2, gNameCols.LAST, nameUpdates.length, 1).setValues(lnArr);
    }

    // Deadline columns: respect steward overrides — resolve from headers (v4.51.1)
    var dlColMap = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'FILING',    header: 'Filing Deadline',     fallback: GRIEVANCE_COLS.FILING_DEADLINE },
      { key: 'STEP1',     header: 'Step I Due',          fallback: GRIEVANCE_COLS.STEP1_DUE },
      { key: 'STEP2_APP', header: 'Step II Appeal Due',  fallback: GRIEVANCE_COLS.STEP2_APPEAL_DUE },
      { key: 'STEP2',     header: 'Step II Due',         fallback: GRIEVANCE_COLS.STEP2_DUE },
      { key: 'STEP3_APP', header: 'Step III Appeal Due', fallback: GRIEVANCE_COLS.STEP3_APPEAL_DUE }
    ]);
    var deadlineCols = [
      { col: dlColMap.FILING,    idx: 0 },
      { col: dlColMap.STEP1,     idx: 1 },
      { col: dlColMap.STEP2_APP, idx: 2 },
      { col: dlColMap.STEP2,     idx: 3 },
      { col: dlColMap.STEP3_APP, idx: 4 }
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

    // S, T, U: Days Open, Next Action Due, Days to Deadline — resolve from headers (v4.51.1)
    var gMetricCols = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'DAYS_OPEN',    header: 'Days Open',        fallback: GRIEVANCE_COLS.DAYS_OPEN },
      { key: 'NEXT_ACTION',  header: 'Next Action Due',  fallback: GRIEVANCE_COLS.NEXT_ACTION_DUE },
      { key: 'DAYS_DEADLINE', header: 'Days to Deadline', fallback: GRIEVANCE_COLS.DAYS_TO_DEADLINE }
    ]);
    if (gMetricCols.NEXT_ACTION === gMetricCols.DAYS_OPEN + 1 && gMetricCols.DAYS_DEADLINE === gMetricCols.NEXT_ACTION + 1) {
      grievanceSheet.getRange(2, gMetricCols.DAYS_OPEN, metricsUpdates.length, 3).setValues(metricsUpdates);
    } else {
      var doArr = [], naArr = [], ddArr = [];
      for (var mi = 0; mi < metricsUpdates.length; mi++) { doArr.push([metricsUpdates[mi][0]]); naArr.push([metricsUpdates[mi][1]]); ddArr.push([metricsUpdates[mi][2]]); }
      grievanceSheet.getRange(2, gMetricCols.DAYS_OPEN, metricsUpdates.length, 1).setValues(doArr);
      grievanceSheet.getRange(2, gMetricCols.NEXT_ACTION, metricsUpdates.length, 1).setValues(naArr);
      grievanceSheet.getRange(2, gMetricCols.DAYS_DEADLINE, metricsUpdates.length, 1).setValues(ddArr);
    }

    // Format Days Open (S) as whole numbers, Next Action Due (T) as date
    // Days to Deadline (U) uses General format to preserve "Overdue" text
    grievanceSheet.getRange(2, gMetricCols.DAYS_OPEN, metricsUpdates.length, 1).setNumberFormat('0');
    grievanceSheet.getRange(2, gMetricCols.NEXT_ACTION, metricsUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, gMetricCols.DAYS_DEADLINE, metricsUpdates.length, 1).setNumberFormat('General');

    // X-Z: Email, Location, Steward — resolve from headers (v4.51.1)
    var gContactCols = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'EMAIL',    header: 'Member Email',    fallback: GRIEVANCE_COLS.MEMBER_EMAIL },
      { key: 'LOCATION', header: 'Work Location',   fallback: GRIEVANCE_COLS.LOCATION },
      { key: 'STEWARD',  header: 'Assigned Steward', fallback: GRIEVANCE_COLS.STEWARD }
    ]);
    // CR-10: Preserve existing contact info — never overwrite non-empty cells with empty strings
    // Read existing contact info using header-resolved positions (v4.51.1)
    var ecEmail = grievanceSheet.getRange(2, gContactCols.EMAIL, contactUpdates.length, 1).getValues();
    var ecLoc = grievanceSheet.getRange(2, gContactCols.LOCATION, contactUpdates.length, 1).getValues();
    var ecStew = grievanceSheet.getRange(2, gContactCols.STEWARD, contactUpdates.length, 1).getValues();
    var existingContact = [];
    for (var ec = 0; ec < contactUpdates.length; ec++) {
      existingContact.push([ecEmail[ec][0], ecLoc[ec][0], ecStew[ec][0]]);
    }
    for (var ci = 0; ci < contactUpdates.length; ci++) {
      for (var cc = 0; cc < 3; cc++) {
        if (!contactUpdates[ci][cc] && existingContact[ci][cc]) {
          contactUpdates[ci][cc] = existingContact[ci][cc];
        }
      }
    }
    // Check if all 3 are consecutive starting from EMAIL
    var contactConsecutive = gContactCols.LOCATION === gContactCols.EMAIL + 1 &&
                             gContactCols.STEWARD === gContactCols.LOCATION + 1;
    if (contactConsecutive) {
      grievanceSheet.getRange(2, gContactCols.EMAIL, contactUpdates.length, 3).setValues(contactUpdates);
    } else {
      // Write each column individually — contactUpdates is [email, location, steward]
      var eArr = [], lArr = [], sArr = [];
      for (var ci2 = 0; ci2 < contactUpdates.length; ci2++) {
        eArr.push([contactUpdates[ci2][0]]);
        lArr.push([contactUpdates[ci2][1]]);
        sArr.push([contactUpdates[ci2][2]]);
      }
      grievanceSheet.getRange(2, gContactCols.EMAIL, contactUpdates.length, 1).setValues(eArr);
      grievanceSheet.getRange(2, gContactCols.LOCATION, contactUpdates.length, 1).setValues(lArr);
      grievanceSheet.getRange(2, gContactCols.STEWARD, contactUpdates.length, 1).setValues(sArr);
    }
  }

  log_('syncGrievanceFormulasToLog', 'Synced grievance formulas to ' + nameUpdates.length + ' grievances');

  // Show warnings to user if data quality issues found
  var warnings = [];
  if (qualityCounters.missingMemberIds.length > 0) {
    warnings.push(qualityCounters.missingMemberIds.length + ' grievance(s) have no Member ID');
    log_('Missing Member IDs', qualityCounters.missingMemberIds.join(', '));
  }
  if (qualityCounters.orphanedGrievances.length > 0) {
    warnings.push(qualityCounters.orphanedGrievances.length + ' grievance(s) reference non-existent members');
    log_('Orphaned grievances', qualityCounters.orphanedGrievances.join(', '));
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
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Required sheets not found for member sync.', 'Sync Error', 5); } catch (_) {}
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
    memberId = col_(grievanceData[j], GRIEVANCE_COLS.MEMBER_ID);
    var data = lookup[memberId];
    // Existing values (0-indexed from grievanceData row)
    var curFirst = col_(grievanceData[j], GRIEVANCE_COLS.FIRST_NAME);
    var curLast = col_(grievanceData[j], GRIEVANCE_COLS.LAST_NAME);
    var curEmail = col_(grievanceData[j], GRIEVANCE_COLS.MEMBER_EMAIL);
    var curLoc = col_(grievanceData[j], GRIEVANCE_COLS.LOCATION);
    var curSteward = col_(grievanceData[j], GRIEVANCE_COLS.STEWARD);
    if (!data) {
      // No lookup match — preserve existing grievance row values
      nameUpdates.push([curFirst, curLast]);
      infoUpdates.push([curEmail, curLoc, curSteward]);
    } else {
      // Only overwrite with non-empty source values
      nameUpdates.push([data.firstName || curFirst, data.lastName || curLast]);
      infoUpdates.push([data.email || curEmail, data.location || curLoc, data.steward || curSteward]);
    }
  }

  if (nameUpdates.length > 0) {
    // Update C-D (First Name, Last Name) — resolve from headers (v4.51.1)
    var smNameCols = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'FIRST', header: 'First Name', fallback: GRIEVANCE_COLS.FIRST_NAME },
      { key: 'LAST',  header: 'Last Name',  fallback: GRIEVANCE_COLS.LAST_NAME }
    ]);
    if (smNameCols.LAST === smNameCols.FIRST + 1) {
      grievanceSheet.getRange(2, smNameCols.FIRST, nameUpdates.length, 2).setValues(nameUpdates);
    } else {
      var fnA = [], lnA = [];
      for (var sn = 0; sn < nameUpdates.length; sn++) { fnA.push([nameUpdates[sn][0]]); lnA.push([nameUpdates[sn][1]]); }
      grievanceSheet.getRange(2, smNameCols.FIRST, nameUpdates.length, 1).setValues(fnA);
      grievanceSheet.getRange(2, smNameCols.LAST, nameUpdates.length, 1).setValues(lnA);
    }
    // Update X-Z (Email, Location, Steward) — resolve from headers (v4.51.1)
    var smInfoCols = resolveColumnsByHeader_(grievanceSheet, [
      { key: 'EMAIL',    header: 'Member Email',     fallback: GRIEVANCE_COLS.MEMBER_EMAIL },
      { key: 'LOCATION', header: 'Work Location',    fallback: GRIEVANCE_COLS.LOCATION },
      { key: 'STEWARD',  header: 'Assigned Steward',  fallback: GRIEVANCE_COLS.STEWARD }
    ]);
    var infoConsecutive = smInfoCols.LOCATION === smInfoCols.EMAIL + 1 &&
                          smInfoCols.STEWARD === smInfoCols.LOCATION + 1;
    if (infoConsecutive) {
      grievanceSheet.getRange(2, smInfoCols.EMAIL, infoUpdates.length, 3).setValues(infoUpdates);
    } else {
      var eA = [], lA = [], sA = [];
      for (var si = 0; si < infoUpdates.length; si++) { eA.push([infoUpdates[si][0]]); lA.push([infoUpdates[si][1]]); sA.push([infoUpdates[si][2]]); }
      grievanceSheet.getRange(2, smInfoCols.EMAIL, infoUpdates.length, 1).setValues(eA);
      grievanceSheet.getRange(2, smInfoCols.LOCATION, infoUpdates.length, 1).setValues(lA);
      grievanceSheet.getRange(2, smInfoCols.STEWARD, infoUpdates.length, 1).setValues(sA);
    }
  }

  log_('syncMemberToGrievanceLog', 'Synced member data to ' + nameUpdates.length + ' grievances');
}

// ============================================================================
// CONFIG SYNC
// ============================================================================

/**
 * Sync new values from Member Directory to Config (bidirectional sync).
 * When a user enters a new value in a dropdown/multi-select column, add it to Config.
 *
 * Delegates to syncDropdownToConfig_ which uses the dynamically-rebuilt
 * DROPDOWN_MAP and MULTI_SELECT_COLS (kept current by syncColumnMaps).
 *
 * @param {Object} e - The edit event object
 */
function syncNewValueToConfig(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) return;

  // syncDropdownToConfig_ uses DROPDOWN_MAP + MULTI_SELECT_COLS (rebuilt by
  // syncColumnMaps) and writes via addToConfigDropdown_ which correctly finds
  // the first empty row in the target Config column.
  syncDropdownToConfig_(e, sheetName);
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
        log_('onEditAutoSync', 'Error opening grievance form: ' + err.message);
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
        log_('onEditAutoSync', 'Error opening member quick actions: ' + err.message);
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
        log_('onEditAutoSync', 'Error opening grievance quick actions: ' + err.message);
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
      // Auto-create folders for any grievances missing them (throttled to once per 15 min)
      if (!cache.get('folder_sync_throttle')) {
        autoCreateMissingGrievanceFolders_();
        cache.put('folder_sync_throttle', '1', 900); // 15 min throttle
      }
    } else if (sheetName === SHEETS.MEMBER_DIR) {
      // Member Directory changed - sync to Grievance Log and Config
      syncNewValueToConfig(e);  // Bidirectional: add new values to Config
      syncGrievanceFormulasToLog();
      syncMemberToGrievanceLog();
      // Update Dashboard with new computed values
      syncDashboardValues();
    }
  } catch (error) {
    log_('Auto-sync error', error.message);
  }
}

/**
 * Manual sync all data with data quality validation
 */
function syncAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Syncing all data...', 'Sync', 3);

  // Restore any Config entries that exist in Member Directory / Grievance Log
  // but are missing from the Config sheet (self-healing bidirectional sync)
  if (typeof restoreConfigFromSheetData_ === 'function') {
    try { restoreConfigFromSheetData_(); } catch (_e) { log_('Config restore skipped', _e.message); }
  }

  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();
  syncChecklistCalcToGrievanceLog();

  // Repair checkboxes after sync
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  // Run data quality check (wrapped in try/catch so sync completes even if check fails)
  var issues = [];
  try {
    if (typeof checkDataQuality === 'function') issues = checkDataQuality();
  } catch (_dq) {
    log_('Data quality check failed', _dq.message);
  }

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
  var _ui = SpreadsheetApp.getUi();
  showDialog_(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:' + SHEET_COLORS.DIALOG_ACCENT + ';margin-top:0}' +
    '.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px}' +
    '.section h4{margin:0 0 10px;color:#333}' +
    '.option{display:flex;align-items:center;margin:8px 0}' +
    '.option input[type="checkbox"]{margin-right:10px}' +
    '.option label{font-size:14px}' +
    '.info{background:#e8f4fd;padding:12px;border-radius:8px;font-size:13px;margin-bottom:15px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 20px;border:none;border-radius:4px;cursor:pointer}' +
    '.primary{background:' + SHEET_COLORS.DIALOG_ACCENT + ';color:white;flex:1}' +
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
    'google.script.run.withFailureHandler(function(e){alert(e.message)}).withSuccessHandler(function(){google.script.host.close()}).installAutoSyncTriggerWithOptions(opts)}' +
    '</script></body></html>',
    'Auto-Sync Settings', 450, 480);
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

  log_('Auto-sync trigger installed with options', JSON.stringify(options));
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

  log_('installAutoSyncTriggerQuick', 'Auto-sync trigger installed (quick mode)');
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

  log_('removeAutoSyncTrigger', 'Removed ' + removed + ' auto-sync triggers');
  SpreadsheetApp.getActiveSpreadsheet().toast('Auto-sync trigger removed', 'Info', 3);
}
// ============================================================================
// DASHBOARD VALUE SYNC
// ============================================================================

/**
 * Sync computed values to Dashboard sheet (no formulas)
 * Replaces all Dashboard formulas with JavaScript-computed values
 * Called during CREATE_DASHBOARD and on data changes
 */
function syncDashboardValues() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { log_('syncDashboardValues', 'lock contention — skipped'); return; }
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashSheet) {
    log_('syncDashboardValues', 'Dashboard sheet not found');
    return;
  }

  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) { log_('syncDashboardValues', 'Config sheet not found for Dashboard sync'); }

  if (!memberSheet || !grievanceSheet) {
    log_('syncDashboardValues', 'Required sheets not found for Dashboard sync');
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

  log_('syncDashboardValues', 'Dashboard values synced');
  } finally { lock.releaseLock(); }
}

/**
 * Compute member-related metrics (total members, stewards, open rates, vol hours, contacts).
 * @param {Array<Array>} memberData - Raw member data with headers at index 0
 * @param {Date} today - Current date
 * @param {Date} thisMonthStart - First day of current month
 * @returns {Object} Member metrics
 * @private
 */
function _computeMemberMetrics_(memberData, today, thisMonthStart) {
  var result = {
    totalMembers: 0,
    activeStewards: 0,
    avgOpenRate: '-',
    ytdVolHours: 0,
    contactsThisMonth: 0
  };
  var openRates = [];

  for (var m = 1; m < memberData.length; m++) {
    var row = memberData[m];
    if (!col_(row, MEMBER_COLS.MEMBER_ID)) continue;

    result.totalMembers++;

    if (isTruthyValue(col_(row, MEMBER_COLS.IS_STEWARD))) {
      result.activeStewards++;
    }

    var openRate = col_(row, MEMBER_COLS.OPEN_RATE);
    if (typeof openRate === 'number') {
      openRates.push(openRate);
    }

    var volHours = col_(row, MEMBER_COLS.VOLUNTEER_HOURS);
    if (typeof volHours === 'number') {
      result.ytdVolHours += volHours;
    }

    var contactDate = col_(row, MEMBER_COLS.RECENT_CONTACT_DATE);
    if (contactDate instanceof Date && contactDate >= thisMonthStart && contactDate <= today) {
      result.contactsThisMonth++;
    }
  }

  if (openRates.length > 0) {
    var avgRate = openRates.reduce(function(a, b) { return a + b; }, 0) / openRates.length;
    result.avgOpenRate = Math.round(avgRate * 10) / 10 + '%';
  }

  return result;
}

/**
 * Compute grievance status counts, trends, averages, and build intermediate
 * category/location/steward lookup objects for downstream helpers.
 * @param {Array<Array>} grievanceData - Raw grievance data with headers at index 0
 * @param {Date} today - Current date
 * @param {Date} thisMonthStart - First day of current month
 * @param {Date} lastMonthStart - First day of last month
 * @param {Date} lastMonthEnd - Last day of last month
 * @returns {Object} Grievance metrics + intermediate lookup objects
 * @private
 */
function _computeGrievanceMetrics_(grievanceData, today, thisMonthStart, lastMonthStart, lastMonthEnd) {
  var result = {
    open: 0, pendingInfo: 0, settled: 0, won: 0, denied: 0, withdrawn: 0,
    activeGrievances: 0, overdueCases: 0, dueThisWeek: 0,
    avgDaysOpen: 0, filedThisMonth: 0, closedThisMonth: 0, avgResolutionDays: 0,
    winRate: '-',
    trends: {
      filed: { thisMonth: 0, lastMonth: 0 },
      closed: { thisMonth: 0, lastMonth: 0 },
      won: { thisMonth: 0, lastMonth: 0 }
    },
    // Intermediate objects for downstream helpers
    categoryStats: {},
    locationStats: {},
    stewardGrievances: {}
  };

  var daysOpenValues = [];
  var closedDaysValues = [];

  for (var g = 1; g < grievanceData.length; g++) {
    var gRow = grievanceData[g];
    if (!col_(gRow, GRIEVANCE_COLS.GRIEVANCE_ID)) continue;

    // Normalize status for comparison so case/whitespace variants from CSV
    // imports or legacy data don't silently drop from every counter.
    var status = String(col_(gRow, GRIEVANCE_COLS.STATUS) || '').trim();
    var statusLc = status.toLowerCase();
    var steward = col_(gRow, GRIEVANCE_COLS.STEWARD);
    var category = col_(gRow, GRIEVANCE_COLS.ISSUE_CATEGORY);
    var location = col_(gRow, GRIEVANCE_COLS.LOCATION);
    var dateFiled = col_(gRow, GRIEVANCE_COLS.DATE_FILED);
    var dateClosed = col_(gRow, GRIEVANCE_COLS.DATE_CLOSED);
    var daysOpen = col_(gRow, GRIEVANCE_COLS.DAYS_OPEN);
    var daysToDeadline = col_(gRow, GRIEVANCE_COLS.DAYS_TO_DEADLINE);

    // Status counts (case-insensitive so 'open' matches 'Open')
    if (statusLc === String(GRIEVANCE_STATUS.OPEN).toLowerCase()) result.open++;
    else if (statusLc === String(GRIEVANCE_STATUS.PENDING).toLowerCase()) result.pendingInfo++;
    else if (statusLc === String(GRIEVANCE_STATUS.SETTLED).toLowerCase()) result.settled++;
    else if (statusLc === String(GRIEVANCE_STATUS.WON).toLowerCase()) result.won++;
    else if (statusLc === String(GRIEVANCE_STATUS.DENIED).toLowerCase()) result.denied++;
    else if (statusLc === String(GRIEVANCE_STATUS.WITHDRAWN).toLowerCase()) result.withdrawn++;

    // Active grievances = anything not in the closed-status set.
    if (status && typeof GRIEVANCE_CLOSED_STATUSES !== 'undefined' && GRIEVANCE_CLOSED_STATUSES.indexOf(status) === -1) {
      result.activeGrievances++;
    }

    // Overdue and due this week
    // Note: daysToDeadline can be a number OR the string "Overdue"
    if (daysToDeadline === 'Overdue') {
      result.overdueCases++;
    } else if (typeof daysToDeadline === 'number') {
      if (daysToDeadline < 0) result.overdueCases++;
      else if (daysToDeadline <= 7) result.dueThisWeek++;
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
      result.filedThisMonth++;
      result.trends.filed.thisMonth++;
    }
    if (dateFiled instanceof Date && dateFiled >= lastMonthStart && dateFiled <= lastMonthEnd) {
      result.trends.filed.lastMonth++;
    }

    // Closed this month
    if (dateClosed instanceof Date && dateClosed >= thisMonthStart && dateClosed <= today) {
      result.closedThisMonth++;
      result.trends.closed.thisMonth++;
      if (status === GRIEVANCE_STATUS.WON) {
        result.trends.won.thisMonth++;
      }
    }
    if (dateClosed instanceof Date && dateClosed >= lastMonthStart && dateClosed <= lastMonthEnd) {
      result.trends.closed.lastMonth++;
      if (status === GRIEVANCE_STATUS.WON) {
        result.trends.won.lastMonth++;
      }
    }

    // Category stats
    if (category) {
      if (!result.categoryStats[category]) {
        result.categoryStats[category] = { total: 0, open: 0, resolved: 0, won: 0, denied: 0, settled: 0, withdrawn: 0, daysOpen: [] };
      }
      result.categoryStats[category].total++;
      if (status === GRIEVANCE_STATUS.OPEN) result.categoryStats[category].open++;
      if (status !== GRIEVANCE_STATUS.OPEN && status !== GRIEVANCE_STATUS.PENDING) result.categoryStats[category].resolved++;
      if (status === GRIEVANCE_STATUS.WON) result.categoryStats[category].won++;
      if (status === GRIEVANCE_STATUS.DENIED) result.categoryStats[category].denied++;
      if (status === GRIEVANCE_STATUS.SETTLED) result.categoryStats[category].settled++;
      if (status === GRIEVANCE_STATUS.WITHDRAWN) result.categoryStats[category].withdrawn++;
      if (typeof daysOpen === 'number') result.categoryStats[category].daysOpen.push(daysOpen);
    }

    // Location stats
    if (location) {
      if (!result.locationStats[location]) {
        result.locationStats[location] = { members: 0, grievances: 0, open: 0, won: 0, denied: 0, settled: 0, withdrawn: 0 };
      }
      result.locationStats[location].grievances++;
      if (status === GRIEVANCE_STATUS.OPEN) result.locationStats[location].open++;
      if (status === GRIEVANCE_STATUS.WON) result.locationStats[location].won++;
      if (status === GRIEVANCE_STATUS.DENIED) result.locationStats[location].denied++;
      if (status === GRIEVANCE_STATUS.SETTLED) result.locationStats[location].settled++;
      if (status === GRIEVANCE_STATUS.WITHDRAWN) result.locationStats[location].withdrawn++;
    }

    // Steward stats
    if (steward) {
      if (!result.stewardGrievances[steward]) {
        result.stewardGrievances[steward] = { active: 0, open: 0, pendingInfo: 0, total: 0 };
      }
      result.stewardGrievances[steward].total++;
      if (status === GRIEVANCE_STATUS.OPEN) {
        result.stewardGrievances[steward].active++;
        result.stewardGrievances[steward].open++;
      } else if (status === GRIEVANCE_STATUS.PENDING) {
        result.stewardGrievances[steward].active++;
        result.stewardGrievances[steward].pendingInfo++;
      }
    }
  }

  // Calculate averages
  if (daysOpenValues.length > 0) {
    result.avgDaysOpen = Math.round(daysOpenValues.reduce(function(a, b) { return a + b; }, 0) / daysOpenValues.length * 10) / 10;
  }
  if (closedDaysValues.length > 0) {
    result.avgResolutionDays = Math.round(closedDaysValues.reduce(function(a, b) { return a + b; }, 0) / closedDaysValues.length * 10) / 10;
  }

  // Win rate
  var totalOutcomes = result.won + result.denied + result.settled + result.withdrawn;
  if (totalOutcomes > 0) {
    result.winRate = Math.round(result.won / totalOutcomes * 100) + '%';
  }

  return result;
}

/**
 * Compute 6-month filing/closing history for sparklines.
 * @param {Array<Array>} grievanceData - Raw grievance data with headers at index 0
 * @param {Date} today - Current date
 * @param {number} activeGrievances - Current count of active grievances
 * @param {number} totalMembers - Current total member count
 * @returns {Object} sixMonthHistory with casesFiled, grievances, members arrays
 * @private
 */
function _computeMonthlyHistory_(grievanceData, today, activeGrievances, totalMembers) {
  var monthlyFiledCounts = [0, 0, 0, 0, 0, 0]; // [5 months ago, 4, 3, 2, 1, current]
  var monthlyClosedCounts = [0, 0, 0, 0, 0, 0];

  for (var h = 1; h < grievanceData.length; h++) {
    var hRow = grievanceData[h];
    if (!col_(hRow, GRIEVANCE_COLS.GRIEVANCE_ID)) continue;

    var hDateFiled = col_(hRow, GRIEVANCE_COLS.DATE_FILED);
    var hDateClosed = col_(hRow, GRIEVANCE_COLS.DATE_CLOSED);

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

  return {
    casesFiled: monthlyFiledCounts,
    grievances: monthlyFiledCounts.map(function(val, idx) {
      // Running total of active grievances (approximation)
      return activeGrievances + monthlyFiledCounts.slice(idx + 1).reduce(function(a, b) { return a + b; }, 0) -
             monthlyClosedCounts.slice(idx + 1).reduce(function(a, b) { return a + b; }, 0);
    }),
    // Member count history not tracked — show current count consistently (no fabricated trends)
    members: [totalMembers, totalMembers, totalMembers, totalMembers, totalMembers, totalMembers]
  };
}

/**
 * Compute top-5 category analysis from pre-built category stats.
 * @param {Object} categoryStats - Category stats built by _computeGrievanceMetrics_
 * @returns {Array<Object>} Array of category analysis objects
 * @private
 */
function _computeCategoryAnalysis_(categoryStats) {
  var categories = [];
  var defaultCategories = ['Contract Violation', 'Discipline', 'Workload', 'Safety', 'Discrimination'];
  for (var c = 0; c < defaultCategories.length; c++) {
    var cat = defaultCategories[c];
    var catData = categoryStats[cat] || { total: 0, open: 0, resolved: 0, won: 0, denied: 0, settled: 0, withdrawn: 0, daysOpen: [] };
    var catResolved = catData.won + catData.denied + catData.settled + catData.withdrawn;
    var catWinRate = catResolved > 0 ? Math.round(catData.won / catResolved * 100) + '%' : '-';
    var avgDays = catData.daysOpen.length > 0 ?
      Math.round(catData.daysOpen.reduce(function(a, b) { return a + b; }, 0) / catData.daysOpen.length * 10) / 10 : '-';

    categories.push({
      name: cat,
      total: catData.total,
      open: catData.open,
      resolved: catData.resolved,
      winRate: catWinRate,
      avgDays: avgDays
    });
  }
  return categories;
}

/**
 * Compute top-5 location breakdown from member data, config data, and pre-built location stats.
 * @param {Array<Array>} memberData - Raw member data with headers at index 0
 * @param {Array<Array>} configData - Raw config data
 * @param {Object} locationStats - Location stats built by _computeGrievanceMetrics_
 * @returns {Array<Object>} Array of location breakdown objects
 * @private
 */
function _computeLocationBreakdown_(memberData, configData, locationStats) {
  var locations = [];

  // Count members per location
  var memberLocations = {};
  for (var ml = 1; ml < memberData.length; ml++) {
    var loc = col_(memberData[ml], MEMBER_COLS.WORK_LOCATION);
    if (loc) {
      memberLocations[loc] = (memberLocations[loc] || 0) + 1;
    }
  }

  // Get locations dynamically from Config column (skip header rows 0-1, data starts at index 2)
  var configLocations = [];
  for (var cl = 2; cl < configData.length; cl++) {
    var configLoc = configData[cl] ? col_(configData[cl], CONFIG_COLS.OFFICE_LOCATIONS) : '';
    if (configLoc && String(configLoc).trim() !== '') {
      configLocations.push(String(configLoc).trim());
    }
  }
  // Use top 5 locations by member count
  configLocations.sort(function(a, b) {
    return (memberLocations[b] || 0) - (memberLocations[a] || 0);
  });
  var topLocations = configLocations.slice(0, 5);
  for (var l = 0; l < topLocations.length; l++) {
    var locName = topLocations[l];
    var locData = locationStats[locName] || { members: 0, grievances: 0, open: 0, won: 0, denied: 0, settled: 0, withdrawn: 0 };
    locData.members = memberLocations[locName] || 0;
    var locResolved = locData.won + locData.denied + locData.settled + locData.withdrawn;
    var locWinRate = locResolved > 0 ? Math.round(locData.won / locResolved * 100) + '%' : '-';

    locations.push({
      name: locName,
      members: locData.members,
      grievances: locData.grievances,
      open: locData.open,
      winRate: locWinRate,
      satisfaction: '-'
    });
  }

  return locations;
}

/**
 * Compute steward summary, busiest stewards list, and top/bottom performers.
 * @param {Object} stewardGrievances - Steward stats built by _computeGrievanceMetrics_
 * @param {number} activeStewards - Count of active stewards
 * @param {number} ytdVolHours - Year-to-date volunteer hours
 * @param {number} contactsThisMonth - Contacts this month count
 * @param {number} grievanceCount - Total grievance row count (excluding header)
 * @returns {Object} stewardSummary, busiestStewards, topPerformers, needingSupport
 * @private
 */
function _computeStewardSummary_(stewardGrievances, activeStewards, ytdVolHours, contactsThisMonth, grievanceCount) {
  var result = {
    stewardSummary: {
      total: activeStewards,
      activeWithCases: 0,
      avgCasesPerSteward: '-',
      totalVolHours: ytdVolHours,
      contactsThisMonth: contactsThisMonth
    },
    busiestStewards: [],
    topPerformers: [],
    needingSupport: []
  };

  var stewardsWithActiveCases = Object.keys(stewardGrievances).filter(function(s) {
    return stewardGrievances[s].active > 0;
  }).length;
  result.stewardSummary.activeWithCases = stewardsWithActiveCases;

  if (activeStewards > 0) {
    result.stewardSummary.avgCasesPerSteward = Math.round(grievanceCount / activeStewards * 10) / 10;
  }

  // Top 30 busiest stewards
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
  result.busiestStewards = stewardArray.slice(0, 30);

  // Top/bottom performers (from hidden sheet)
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
    result.topPerformers = performers.slice(0, 10);

    // Sort by score ascending for needing support
    performers.sort(function(a, b) { return a.score - b.score; });
    result.needingSupport = performers.slice(0, 10);
  }

  return result;
}

/**
 * Compute all Dashboard metrics from raw data.
 * Delegates to focused helpers for each metric section.
 * @param {Array<Array>} memberData - Raw member data with headers at index 0
 * @param {Array<Array>} grievanceData - Raw grievance data with headers at index 0
 * @param {Array<Array>} configData - Raw config data
 * @returns {Object} Combined metrics object
 * @private
 */
function computeDashboardMetrics_(memberData, grievanceData, configData) {
  var today = new Date();
  var thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  var lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);

  // ══════════════════════════════════════════════════════════════════════
  // MEMBER METRICS
  // ══════════════════════════════════════════════════════════════════════
  var memberMetrics = _computeMemberMetrics_(memberData, today, thisMonthStart);

  // ══════════════════════════════════════════════════════════════════════
  // GRIEVANCE METRICS (also builds category/location/steward lookups)
  // ══════════════════════════════════════════════════════════════════════
  var grievanceMetrics = _computeGrievanceMetrics_(grievanceData, today, thisMonthStart, lastMonthStart, lastMonthEnd);

  // ══════════════════════════════════════════════════════════════════════
  // 6-MONTH HISTORICAL DATA FOR SPARKLINES
  // ══════════════════════════════════════════════════════════════════════
  var sixMonthHistory = _computeMonthlyHistory_(grievanceData, today, grievanceMetrics.activeGrievances, memberMetrics.totalMembers);

  // ══════════════════════════════════════════════════════════════════════
  // CATEGORY ANALYSIS (Top 5)
  // ══════════════════════════════════════════════════════════════════════
  var categories = _computeCategoryAnalysis_(grievanceMetrics.categoryStats);

  // ══════════════════════════════════════════════════════════════════════
  // LOCATION BREAKDOWN (Top 5 from Config)
  // ══════════════════════════════════════════════════════════════════════
  var locations = _computeLocationBreakdown_(memberData, configData, grievanceMetrics.locationStats);

  // ══════════════════════════════════════════════════════════════════════
  // STEWARD SUMMARY + BUSIEST STEWARDS + PERFORMERS
  // ══════════════════════════════════════════════════════════════════════
  var stewardResult = _computeStewardSummary_(
    grievanceMetrics.stewardGrievances,
    memberMetrics.activeStewards,
    memberMetrics.ytdVolHours,
    memberMetrics.contactsThisMonth,
    grievanceData.length - 1
  );

  // ══════════════════════════════════════════════════════════════════════
  // ASSEMBLE FINAL METRICS
  // ══════════════════════════════════════════════════════════════════════
  return {
    // Quick Stats
    totalMembers: memberMetrics.totalMembers,
    activeStewards: memberMetrics.activeStewards,
    activeGrievances: grievanceMetrics.activeGrievances,
    winRate: grievanceMetrics.winRate,
    overdueCases: grievanceMetrics.overdueCases,
    dueThisWeek: grievanceMetrics.dueThisWeek,

    // Member Metrics
    avgOpenRate: memberMetrics.avgOpenRate,
    ytdVolHours: memberMetrics.ytdVolHours,

    // Grievance Metrics
    open: grievanceMetrics.open,
    pendingInfo: grievanceMetrics.pendingInfo,
    settled: grievanceMetrics.settled,
    won: grievanceMetrics.won,
    denied: grievanceMetrics.denied,
    withdrawn: grievanceMetrics.withdrawn,

    // Timeline Metrics
    avgDaysOpen: grievanceMetrics.avgDaysOpen,
    filedThisMonth: grievanceMetrics.filedThisMonth,
    closedThisMonth: grievanceMetrics.closedThisMonth,
    avgResolutionDays: grievanceMetrics.avgResolutionDays,

    // Category Analysis (top 5)
    categories: categories,

    // Location Breakdown (top 5)
    locations: locations,

    // Month-over-Month Trends
    trends: grievanceMetrics.trends,

    // 6-Month Historical Data for Sparklines
    sixMonthHistory: sixMonthHistory,

    // Steward Summary
    stewardSummary: stewardResult.stewardSummary,

    // Top 30 Busiest Stewards
    busiestStewards: stewardResult.busiestStewards,

    // Top 10 Performers (from hidden sheet)
    topPerformers: stewardResult.topPerformers,

    // Bottom 10 (needing support)
    needingSupport: stewardResult.needingSupport
  };
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
    changeCell44.setFontColor(SHEET_COLORS.STATUS_SUCCESS); // Green - grievances down is good
  } else if (change44Val > 0) {
    changeCell44.setFontColor(SHEET_COLORS.STATUS_ERROR); // Red - grievances up is bad
  } else {
    changeCell44.setFontColor(SHEET_COLORS.STATUS_DISABLED); // Gray - no change
  }

  // For members: positive change = green (good), negative = red (bad)
  var changeCell45 = sheet.getRange(L.TREND_START_ROW + 1, 4);
  if (memberChange > 0) {
    changeCell45.setFontColor(SHEET_COLORS.STATUS_SUCCESS); // Green - members up is good
  } else if (memberChange < 0) {
    changeCell45.setFontColor(SHEET_COLORS.STATUS_ERROR); // Red - members down is bad
  } else {
    changeCell45.setFontColor(SHEET_COLORS.STATUS_DISABLED); // Gray
  }

  // For cases filed: neutral coloring (blue)
  var changeCell46 = sheet.getRange(L.TREND_START_ROW + 2, 4);
  changeCell46.setFontColor(SHEET_COLORS.STATUS_INFO); // Blue - neutral

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
  var L = DASHBOARD_LAYOUT;

  // Derive gradient ranges from DASHBOARD_LAYOUT constants instead of hardcoding A1 notation
  var GRADIENT_RANGES = {
    ACTIVE_CASES:     { startRow: L.BUSIEST_START_ROW,        endRow: L.BUSIEST_END_ROW,          col: 3 },  // C59:C88
    PERF_SCORE:       { startRow: L.TOP_PERFORMERS_START_ROW, endRow: L.TOP_PERFORMERS_END_ROW,   col: 3 },  // C93:C102
    PERF_WIN_RATE:    { startRow: L.TOP_PERFORMERS_START_ROW, endRow: L.TOP_PERFORMERS_END_ROW,   col: 4 },  // D93:D102
    PERF_OVERDUE:     { startRow: L.TOP_PERFORMERS_START_ROW, endRow: L.TOP_PERFORMERS_END_ROW,   col: 6 },  // F93:F102
    NEED_SCORE:       { startRow: L.NEEDING_SUPPORT_START_ROW, endRow: L.NEEDING_SUPPORT_END_ROW, col: 3 },  // C107:C116
    NEED_OVERDUE:     { startRow: L.NEEDING_SUPPORT_START_ROW, endRow: L.NEEDING_SUPPORT_END_ROW, col: 6 },  // F107:F116
    CAT_WIN_RATE:     { startRow: L.CATEGORY_START_ROW,       endRow: L.CATEGORY_END_ROW,         col: 5 },  // E26:E30
    LOC_WIN_RATE:     { startRow: L.LOCATION_START_ROW,       endRow: L.LOCATION_END_ROW,         col: 5 }   // E35:E39
  };

  /** Helper: convert {startRow, endRow, col} to a Sheet Range */
  function _gr(key) {
    var g = GRADIENT_RANGES[key];
    return sheet.getRange(g.startRow, g.col, g.endRow - g.startRow + 1, 1);
  }

  // Reverse scale (Red -> Yellow -> Green) for positive metrics
  var redToGreen = {
    minColor: '#FCA5A5',
    midColor: SHEET_COLORS.BG_LIGHT_YELLOW,
    maxColor: SHEET_COLORS.BG_PALE_GREEN
  };

  // Standard scale (Green -> Yellow -> Red) for negative metrics
  var greenToRed = {
    minColor: SHEET_COLORS.BG_PALE_GREEN,
    midColor: SHEET_COLORS.BG_LIGHT_YELLOW,
    maxColor: '#FCA5A5'
  };

  // ── Active Cases Column (Top 30 Busiest) - Higher = more work (red)
  var activeCasesRange = _gr('ACTIVE_CASES');
  var activeCasesRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([activeCasesRange])
    .build();

  // ── Score Column (Top 10 Performers) - Higher = better (green)
  var scoreRange = _gr('PERF_SCORE');
  var scoreRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([scoreRange])
    .build();

  // ── Win Rate Column (Top 10 Performers) - Higher = better (green)
  var winRateRange = _gr('PERF_WIN_RATE');
  var winRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([winRateRange])
    .build();

  // ── Overdue Column (Performers) - Lower = better (green at low)
  var overdueRange = _gr('PERF_OVERDUE');
  var overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([overdueRange])
    .build();

  // ── Score Column (Needing Support) - Lower scores (red)
  var needScoreRange = _gr('NEED_SCORE');
  var needScoreRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([needScoreRange])
    .build();

  // ── Overdue Column (Needing Support) - Highlight high overdue
  var needOverdueRange = _gr('NEED_OVERDUE');
  var needOverdueRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(greenToRed.minColor)
    .setGradientMidpointWithValue(greenToRed.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(greenToRed.maxColor)
    .setRanges([needOverdueRange])
    .build();

  // ── Category Win Rate (Issue Breakdown) - Higher = better (green)
  var catWinRateRange = _gr('CAT_WIN_RATE');
  var catWinRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([catWinRateRange])
    .build();

  // ── Location Win Rate - Higher = better (green)
  var locWinRateRange = _gr('LOC_WIN_RATE');
  var locWinRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(redToGreen.minColor)
    .setGradientMidpointWithValue(redToGreen.midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(redToGreen.maxColor)
    .setRanges([locWinRateRange])
    .build();

  // Collect A1 notation of all gradient ranges for dedup filter
  var gradientA1 = [activeCasesRange, scoreRange, winRateRange, overdueRange,
    needScoreRange, needOverdueRange, catWinRateRange, locWinRateRange]
    .map(function(r) { return r.getA1Notation(); });

  // Apply all rules
  var rules = sheet.getConditionalFormatRules();

  // Remove existing gradient rules to avoid duplicates
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    if (ranges.length === 0) return true;
    var rangeStr = ranges[0].getA1Notation();
    // Keep rules that aren't our gradient ranges
    return gradientA1.indexOf(rangeStr) === -1;
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

// Dead code removed: getFlaggedSubmissionsData() — zero callers in src

/**
 * Approve a flagged submission - mark as Verified
 * @param {number} rowNum - Row number (1-indexed)
 */
function approveFlaggedSubmission(rowNum) {
  // Verify caller is an authorized steward/admin
  var callerEmail = '';
  try { callerEmail = Session.getActiveUser().getEmail(); } catch (_e) { log_('approveFlaggedSubmission', 'Error resolving caller: ' + (_e.message || _e)); }
  if (!callerEmail) {
    throw new Error('Authorization required: unable to verify user identity');
  }
  var callerRole = getUserRole_(callerEmail);
  if (callerRole !== 'admin' && callerRole !== 'steward') {
    throw new Error('Access denied: steward or admin role required to approve submissions');
  }

  // Update in vault (not on Satisfaction sheet — no PII there)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault) return;

  var vaultData = vault.getDataRange().getValues();
  for (var i = 1; i < vaultData.length; i++) {
    if (col_(vaultData[i], SURVEY_VAULT_COLS.RESPONSE_ROW) === rowNum) {
      var vaultRow = i + 1;
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.VERIFIED).setValue('Yes');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.REVIEWER_NOTES).setValue(escapeForFormula('Approved by ' + callerEmail + ' on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')));
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
  // Verify caller is an authorized steward/admin
  var callerEmail = '';
  try { callerEmail = Session.getActiveUser().getEmail(); } catch (_e) { log_('rejectFlaggedSubmission', 'Error resolving caller: ' + (_e.message || _e)); }
  if (!callerEmail) {
    throw new Error('Authorization required: unable to verify user identity');
  }
  var callerRole = getUserRole_(callerEmail);
  if (callerRole !== 'admin' && callerRole !== 'steward') {
    throw new Error('Access denied: steward or admin role required to reject submissions');
  }

  // Update in vault (not on Satisfaction sheet — no PII there)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault) return;

  var vaultData = vault.getDataRange().getValues();
  for (var i = 1; i < vaultData.length; i++) {
    if (col_(vaultData[i], SURVEY_VAULT_COLS.RESPONSE_ROW) === rowNum) {
      var vaultRow = i + 1;
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.VERIFIED).setValue('Rejected');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.IS_LATEST).setValue('No');
      vault.getRange(vaultRow, SURVEY_VAULT_COLS.REVIEWER_NOTES).setValue(escapeForFormula('Rejected by ' + callerEmail + ' on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')));
      break;
    }
  }

  // Update dashboard
  syncSatisfactionValues();
}

// ============================================================================
// SHEET THEME & FORMATTING (merged from 09a_SheetFormatting.gs)
// ============================================================================

// ============================================================================
// TAB COLOR ASSIGNMENTS — which sheet gets which tab-bar color
// ============================================================================

/**
 * Map of sheet names to tab-bar color groups.
 * Blue  = data entry, Green = engagement/survey,
 * Gold  = documentation/guide, Red = admin/technical.
 * @private
 */
var TAB_COLOR_MAP_ = {};
// Populated lazily by applyTabBarColors_() because SHEETS may not yet be
// initialised when this file first loads in the GAS V8 runtime.

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Safely gets a sheet by name, returns null if not found.
 * @param {Spreadsheet} ss
 * @param {string} name
 * @returns {Sheet|null}
 * @private
 */
function getSheetSafe_(ss, name) {
  if (!name) return null;
  return ss.getSheetByName(name) || null;
}

/**
 * Applies the Union brand header style to row 1 of a sheet.
 * Navy background, white bold text, 14pt, centered, 40px tall.
 * @param {Sheet} sheet
 * @param {number} numCols - Number of columns to span
 * @private
 */
function applyBrandHeader_(sheet, numCols) {
  if (numCols < 1) return;
  var range = sheet.getRange(1, 1, 1, numCols);
  range
    .setBackground(SHEET_COLORS.THEME_NAVY)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
}

/**
 * Applies a subtitle/instruction row (italic gray, 10pt).
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @param {string} text - Subtitle text to set (if falsy, only styles existing text)
 * @private
 */
function applySubtitleRow_(sheet, row, numCols, text) {
  if (numCols < 1) return;
  var range = sheet.getRange(row, 1, 1, numCols);
  if (text) {
    sheet.getRange(row, 1, 1, 1).setValue(text);
  }
  range
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setFontStyle('italic')
    .setFontSize(10)
    .setBackground(SHEET_COLORS.BG_WHITE);
}

/**
 * Applies alternating row banding (Light Sky / white) to a data range.
 * Uses a single setBackgrounds() call for performance.
 * @param {Sheet} sheet
 * @param {number} startRow - First data row
 * @param {number} numCols
 * @param {number} [endRow] - Last row (defaults to sheet lastRow or startRow+50)
 * @private
 */
function applyRowBanding_(sheet, startRow, numCols, endRow) {
  var lastRow = endRow || Math.max(sheet.getLastRow(), startRow + 50);
  var rowCount = lastRow - startRow + 1;
  if (rowCount < 1 || numCols < 1) return;

  var bgColors = [];
  for (var r = 0; r < rowCount; r++) {
    var color = (r % 2 === 0) ? SHEET_COLORS.BG_WHITE : SHEET_COLORS.THEME_LIGHT_SKY;
    bgColors.push(new Array(numCols).fill(color));
  }
  sheet.getRange(startRow, 1, rowCount, numCols).setBackgrounds(bgColors);
}

/**
 * Applies a section divider row (full-width colored bar with centered text).
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @param {string} bgColor
 * @private
 */
function applySectionDivider_(sheet, row, numCols, bgColor) {
  if (numCols < 1) return;
  sheet.getRange(row, 1, 1, numCols)
    .setBackground(bgColor)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);
}

/**
 * Sets an empty-state placeholder message in italic gray.
 * Only writes if the sheet has no data below the header rows.
 * @param {Sheet} sheet
 * @param {number} dataStartRow - First expected data row
 * @param {number} numCols
 * @param {string} message
 * @private
 */
function applyEmptyState_(sheet, dataStartRow, numCols, message) {
  if (sheet.getLastRow() < dataStartRow) {
    sheet.getRange(dataStartRow, 1, 1, numCols).merge()
      .setValue(message)
      .setFontColor(SHEET_COLORS.TEXT_LIGHT_GRAY)
      .setFontStyle('italic')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setBackground(SHEET_COLORS.BG_WHITE);
  }
}

/**
 * Applies a secondary header row style (Steel Blue bg, white bold text).
 * Useful for column header rows below a banner.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} numCols
 * @private
 */
function applyColumnHeaderRow_(sheet, row, numCols) {
  if (numCols < 1) return;
  sheet.getRange(row, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_STEEL_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 32);
}

/**
 * Clears existing conditional formatting rules for a sheet.
 * Called before adding new theme rules to avoid duplicates.
 * @param {Sheet} sheet
 * @private
 */
function clearConditionalFormats_(sheet) {
  sheet.setConditionalFormatRules([]);
}

/**
 * Standard column width presets.
 * @param {Sheet} sheet
 * @param {Object<number,number>} widthMap - {colNumber: widthPx}
 * @private
 */
function applyColumnWidths_(sheet, widthMap) {
  var keys = Object.keys(widthMap);
  for (var i = 0; i < keys.length; i++) {
    var col = parseInt(keys[i], 10);
    if (col > 0 && col <= sheet.getMaxColumns()) {
      sheet.setColumnWidth(col, widthMap[keys[i]]);
    }
  }
}

// ============================================================================
// INDIVIDUAL TAB FORMATTERS
// ============================================================================

// ─── Tab 1: Getting Started ─────────────────────────────────────────────────

function formatGettingStartedTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.GETTING_STARTED);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 7);

  // Row 1 branded header
  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Quick Navigation panel in columns H-J
  var navCol = 8; // column H
  var maxCols = sheet.getMaxColumns();
  if (maxCols < navCol + 2) {
    sheet.insertColumnsAfter(maxCols, (navCol + 2) - maxCols);
  }
  numCols = Math.max(numCols, navCol + 2);
  applyColumnWidths_(sheet, { 8: 160, 9: 160, 10: 160 });

  // Panel title
  sheet.getRange(2, navCol, 1, 3).merge()
    .setValue('QUICK NAVIGATION')
    .setBackground(SHEET_COLORS.THEME_NAVY)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  // Category groups with tab names
  var navGroups = [
    { label: 'CORE DATA', color: '#6A1B9A', tabs: [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG, SHEETS.CASE_CHECKLIST] },
    { label: 'REFERENCE', color: SHEET_COLORS.THEME_GOLD, tabs: [SHEETS.FAQ, SHEETS.RESOURCES, SHEETS.FEATURES_REFERENCE] },
    { label: 'ENGAGEMENT', color: SHEET_COLORS.TAB_GREEN, tabs: [SHEETS.MEETING_ATTENDANCE, SHEETS.SATISFACTION] },
    { label: 'ADMIN', color: SHEET_COLORS.TAB_RED_ORANGE, tabs: [SHEETS.CONFIG] }
  ];

  var navRow = 3;
  for (var g = 0; g < navGroups.length; g++) {
    var group = navGroups[g];
    // Category header spans cols H-J
    sheet.getRange(navRow, navCol, 1, 3).merge()
      .setValue(group.label)
      .setBackground(group.color)
      .setFontColor(SHEET_COLORS.TEXT_WHITE)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    navRow++;

    // Tab names as navigable links across columns
    for (var t = 0; t < group.tabs.length; t++) {
      var tabName = group.tabs[t];
      var targetSheet = getSheetSafe_(ss, tabName);
      if (targetSheet) {
        var gid = targetSheet.getSheetId();
        sheet.getRange(navRow, navCol + (t % 3))
          .setFormula('=HYPERLINK("#gid=' + gid + '","' + tabName.replace(/"/g, '""') + '")')
          .setFontColor(SHEET_COLORS.LINK_PRIMARY)
          .setFontSize(11);
      } else {
        sheet.getRange(navRow, navCol + (t % 3))
          .setValue(tabName)
          .setFontColor(SHEET_COLORS.TEXT_GRAY)
          .setFontSize(11);
      }
      if (t % 3 === 2 || t === group.tabs.length - 1) navRow++;
    }
  }

  // Light border around the nav panel
  var panelRange = sheet.getRange(2, navCol, navRow - 2, 3);
  panelRange.setBorder(true, true, true, true, false, false,
    SHEET_COLORS.THEME_NAVY, SpreadsheetApp.BorderStyle.SOLID);

  // Color-code step sections by scanning column A for "Step" keywords
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var stepColors = {
    '1': SHEET_COLORS.THEME_GOLD,
    '2': SHEET_COLORS.THEME_STEEL_BLUE,
    '3': SHEET_COLORS.THEME_AMBER,
    '4': SHEET_COLORS.THEME_GREEN
  };

  for (var r = 0; r < data.length; r++) {
    var val = String(data[r][0]).trim();
    // Match "Step 1", "Step 2", etc.
    var match = val.match(/^Step\s+(\d)/i);
    if (match && stepColors[match[1]]) {
      // Limit to columns A-G so we don't overwrite the nav panel in H-J
      applySectionDivider_(sheet, r + 2, Math.min(numCols, 7), stepColors[match[1]]);
    }
  }

  log_('Formatted', 'Getting Started');
}

// ─── Tab 2: FAQ ─────────────────────────────────────────────────────────────

function formatFAQTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FAQ);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Apply alternating Q/A backgrounds: scan for question vs answer rows
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var sectionColors = [
    SHEET_COLORS.THEME_NAVY,
    SHEET_COLORS.THEME_STEEL_BLUE,
    SHEET_COLORS.THEME_GREEN,
    SHEET_COLORS.THEME_AMBER,
    SHEET_COLORS.THEME_RED
  ];
  var sectionIdx = 0;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < data.length; r++) {
    var val = String(data[r][0]).trim().toUpperCase();
    var row = r + 2;

    // Detect section headers (all-caps text like "GETTING STARTED", "MEMBER DIRECTORY")
    var rawVal = String(data[r][0]).trim();
    if (rawVal.length > 3 && rawVal === rawVal.toUpperCase() &&
        !val.match(/^\d/) && !val.match(/^Q[:.]/) && val.indexOf('?') === -1 &&
        rawVal.length < 50) {
      var color = sectionColors[sectionIdx % sectionColors.length];
      applySectionDivider_(sheet, row, numCols, color);
      sectionIdx++;
    }
  }

  log_('Formatted', 'FAQ');
}

// ─── Tab 3: Survey Questions ────────────────────────────────────────────────

function formatSurveyQuestionsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SURVEY_QUESTIONS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 10);

  // Row 1 = column headers → brand header style
  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Column widths: Question Text wide, IDs narrow
  applyColumnWidths_(sheet, { 1: 80, 2: 100, 3: 100, 4: 140, 5: 300, 6: 100, 7: 80, 8: 80, 9: 200, 10: 100 });

  // Text wrap on Question Text column (5) and Options column (9)
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1) {
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);
    if (numCols >= 9) sheet.getRange(2, 9, lastRow - 1, 1).setWrap(true);
  }

  // Conditional formatting: Type column (6) color-coding
  clearConditionalFormats_(sheet);
  var typeRange = sheet.getRange(2, 6, Math.max(lastRow - 1, 50), 1);
  var activeRange = sheet.getRange(2, 8, Math.max(lastRow - 1, 50), 1);
  var rules = [];

  var typeColors = [
    { text: 'dropdown', bg: '#DBEAFE' },     // blue
    { text: 'slider-10', bg: '#F3E8FF' },    // purple
    { text: 'radio', bg: '#FFEDD5' },         // orange
    { text: 'paragraph', bg: '#CCFBF1' }     // teal
  ];
  for (var t = 0; t < typeColors.length; t++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(typeColors[t].text)
      .setBackground(typeColors[t].bg)
      .setRanges([typeRange])
      .build());
  }

  // Active column: Y = green, N = gray
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Y')
    .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
    .setRanges([activeRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N')
    .setBackground(SHEET_COLORS.BG_VERY_LIGHT_GRAY)
    .setFontColor(SHEET_COLORS.TEXT_GRAY)
    .setRanges([activeRange])
    .build());

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  log_('Formatted', 'Survey Questions');
}

// ─── Tab 4: Notifications ───────────────────────────────────────────────────

function formatNotificationsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.NOTIFICATIONS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No notifications yet. Use 📢 Notifications menu to create in-app messages.');

  log_('Formatted', 'Notifications');
}

// ─── Tab 5: Resources ───────────────────────────────────────────────────────

function formatResourcesTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.RESOURCES);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 6);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Column widths: Content column wide, text wrap
  applyColumnWidths_(sheet, { 1: 80, 2: 160, 3: 120, 4: 200, 5: 300, 6: 100 });
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1 && numCols >= 5) {
    sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true);  // Summary
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);  // Content
  }

  // Category color-coding via conditional formatting
  clearConditionalFormats_(sheet);
  var catRange = sheet.getRange(2, 3, Math.max(lastRow - 1, 50), 1);
  var catColors = [
    { text: 'Grievance Process', bg: '#FEE2E2' },
    { text: 'Know Your Rights', bg: SHEET_COLORS.BG_LIGHT_GREEN },
    { text: 'Forms & Templates', bg: '#DBEAFE' },
    { text: 'Contact Info', bg: '#CCFBF1' },
    { text: 'Guide', bg: SHEET_COLORS.BG_LIGHT_YELLOW },
    { text: 'FAQ', bg: '#F3E8FF' }
  ];
  var rules = [];
  for (var c = 0; c < catColors.length; c++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(catColors[c].text)
      .setBackground(catColors[c].bg)
      .setRanges([catRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);

  log_('Formatted', 'Resources');
}

// ─── Tab 6: Resource Config ─────────────────────────────────────────────────

function formatResourceConfigTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.RESOURCE_CONFIG);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 4);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);

  log_('Formatted', 'Resource Config');
}

// ─── Tab 9: Settings Overview ───────────────────────────────────────────────

function formatSettingsOverviewTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SETTINGS_OVERVIEW);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  // Steel blue header for admin tabs
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_STEEL_BLUE)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);

  // Conditional formatting: empty Current Value cells = amber warning
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 20);
  if (numCols >= 2) {
    var valueRange = sheet.getRange(2, 2, dataRows, 1);
    var rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenCellEmpty()
        .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
        .setRanges([valueRange])
        .build()
    ];
    sheet.setConditionalFormatRules(rules);
  }

  applyRowBanding_(sheet, 2, numCols);
  log_('Formatted', 'Settings Overview');
}

// ─── Tab 10: Events ─────────────────────────────────────────────────────────

function formatEventsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.PORTAL_EVENTS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 10);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 60, 2: 200, 3: 100, 4: 140, 5: 140, 6: 160, 7: 300, 8: 200, 9: 140, 10: 120 });

  // Conditional formatting: event Type color-coding
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 50);
  var typeRange = sheet.getRange(2, 3, dataRows, 1);
  var typeColors = [
    { text: 'Meeting', bg: '#DBEAFE' },
    { text: 'Negotiation', bg: '#FEE2E2' },
    { text: 'Training', bg: SHEET_COLORS.BG_LIGHT_GREEN },
    { text: 'Social', bg: '#F3E8FF' },
    { text: 'Community', bg: SHEET_COLORS.BG_LIGHT_YELLOW }
  ];
  var rules = [];
  for (var i = 0; i < typeColors.length; i++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(typeColors[i].text)
      .setBackground(typeColors[i].bg)
      .setRanges([typeRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No events yet. Use 📅 Events menu → Add New Event to create your first event.');

  log_('Formatted', 'Events');
}

// ─── Tab 11: MeetingMinutes ─────────────────────────────────────────────────

function formatMeetingMinutesTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.PORTAL_MINUTES);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 60, 2: 120, 3: 200, 4: 300, 5: 300, 6: 140, 7: 120, 8: 200 });

  // Text wrap on Bullets (4) and FullMinutes (5)
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1) {
    sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true);
    sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);
  }

  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No meeting minutes yet. Use 📝 Meeting Minutes menu to add new minutes.');

  log_('Formatted', 'MeetingMinutes');
}

// ─── Tab 12: Workload Reporting ─────────────────────────────────────────────

function formatWorkloadReportingTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.WORKLOAD_REPORTING);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 13);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Conditional formatting: Priority Cases column (3) — 0=green, 1-3=yellow, 4+=red
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 50);
  var priorityRange = sheet.getRange(2, 3, dataRows, 1);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground(SHEET_COLORS.BG_LIGHT_GREEN)
      .setRanges([priorityRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 3)
      .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
      .setRanges([priorityRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(4)
      .setBackground(SHEET_COLORS.BG_LIGHT_RED)
      .setRanges([priorityRange])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No workload data yet. Members submit weekly reports via the Workload tab in the web portal.');

  log_('Formatted', 'Workload Reporting');
}

// ─── Tab 13: Non-Member Contacts ────────────────────────────────────────────

function formatNonMemberContactsTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.NON_MEMBER_CONTACTS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 13);

  // Distinct amber-tinted header to differentiate from Member Directory
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(SHEET_COLORS.THEME_AMBER)
    .setFontColor(SHEET_COLORS.TEXT_WHITE)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);

  applyColumnWidths_(sheet, { 1: 120, 2: 120, 3: 160, 4: 140, 5: 80, 6: 100, 7: 140, 8: 80, 9: 70, 10: 120, 11: 200, 12: 140, 13: 200 });
  applyRowBanding_(sheet, 2, numCols);

  log_('Formatted', 'Non-Member Contacts');
}

// ─── Tab 14: Case Checklist ─────────────────────────────────────────────────

function formatCaseChecklistTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.CASE_CHECKLIST);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 9);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 80, 3: 120, 4: 250, 5: 120, 6: 80, 7: 80, 8: 120, 9: 120 });

  // Conditional formatting
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 1, 100);
  var fullRange = sheet.getRange(2, 1, dataRows, numCols);
  var rules = [];

  // Completed (col G=TRUE) → green tint + strikethrough
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2=TRUE')
    .setBackground(SHEET_COLORS.BG_PALE_GREEN)
    .setStrikethrough(true)
    .setRanges([fullRange])
    .build());

  // Required + NOT completed → red tint warning
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2="Y", $G2<>TRUE)')
    .setBackground(SHEET_COLORS.BG_LIGHT_RED)
    .setRanges([fullRange])
    .build());

  // Action Type color-coding (col C)
  var actionRange = sheet.getRange(2, 3, dataRows, 1);
  var actionColors = [
    { text: 'Filing', bg: SHEET_COLORS.THEME_NAVY, font: SHEET_COLORS.TEXT_WHITE },
    { text: 'Documentation', bg: '#DBEAFE', font: '#1E40AF' },
    { text: 'Hearing Prep', bg: '#FFEDD5', font: '#9A3412' },
    { text: 'Follow-Up', bg: SHEET_COLORS.BG_LIGHT_GREEN, font: SHEET_COLORS.TEXT_DARK_GREEN },
    { text: 'Admin', bg: SHEET_COLORS.BG_VERY_LIGHT_GRAY, font: SHEET_COLORS.TEXT_GRAY }
  ];
  for (var a = 0; a < actionColors.length; a++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(actionColors[a].text)
      .setBackground(actionColors[a].bg)
      .setFontColor(actionColors[a].font)
      .setRanges([actionRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No checklist items yet. Items are auto-generated when a new case is created.');

  log_('Formatted', 'Case Checklist');
}

// ─── Tab 15: Member Satisfaction ────────────────────────────────────────────

function formatMemberSatisfactionTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.SATISFACTION);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 20);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);

  // Freeze first 4 columns (Timestamp + demographics) for horizontal scroll
  sheet.setFrozenColumns(4);

  // Uniform column widths: Timestamp wide, Q columns 90px
  var widths = { 1: 150 };
  for (var col = 2; col <= numCols; col++) {
    widths[col] = (col <= 5) ? 120 : 90;
  }
  applyColumnWidths_(sheet, widths);

  // Set uniform row height for readability
  var lastRow = sheet.getLastRow();
  for (var row = 1; row <= lastRow; row++) {
    sheet.setRowHeight(row, 30);
  }

  // Text wrap on header row so long question text is readable
  sheet.getRange(1, 1, 1, numCols).setWrap(true);

  // Color scale on numeric response columns (typically cols 6-20: slider 1-10)
  // Red→Yellow→Green gradient
  clearConditionalFormats_(sheet);
  if (lastRow > 1 && numCols >= 6) {
    var dataRows = lastRow - 1;
    var rules = [];
    var sliderStart = Math.min(6, numCols);
    var sliderEnd = Math.min(20, numCols);
    var sliderRange = sheet.getRange(2, sliderStart, dataRows, sliderEnd - sliderStart + 1);

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#FCA5A5', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMidpointWithValue('#FDE68A', SpreadsheetApp.InterpolationType.NUMBER, '5')
      .setGradientMaxpointWithValue('#6EE7B7', SpreadsheetApp.InterpolationType.NUMBER, '10')
      .setRanges([sliderRange])
      .build());

    sheet.setConditionalFormatRules(rules);
  }

  applyRowBanding_(sheet, 2, numCols);
  log_('Formatted', 'Member Satisfaction');
}

// ─── Tab 16: Features Reference ─────────────────────────────────────────────

function formatFeaturesReferenceTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.FEATURES_REFERENCE);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 5);

  // This sheet already has good formatting from createFeaturesReferenceSheet.
  // Just re-theme the main header row to match Union brand.
  var row1Val = String(sheet.getRange(1, 1).getValue());
  if (row1Val.indexOf('FEATURES REFERENCE') !== -1) {
    sheet.getRange(1, 1, 1, numCols)
      .setBackground(SHEET_COLORS.THEME_NAVY)
      .setFontColor(SHEET_COLORS.TEXT_WHITE);
  }

  // Re-theme the column header row (usually row 5)
  var lastRow = sheet.getLastRow();
  for (var r = 2; r <= Math.min(lastRow, 10); r++) {
    var val = String(sheet.getRange(r, 1).getValue()).trim();
    if (val === 'Category') {
      applyColumnHeaderRow_(sheet, r, numCols);
      sheet.setFrozenRows(r);
      break;
    }
  }

  log_('Formatted', 'Features Reference');
}

// ─── Tab 17: Volunteer Hours ────────────────────────────────────────────────

function formatVolunteerHoursTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.VOLUNTEER_HOURS);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 9);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 80, 3: 140, 4: 120, 5: 120, 6: 70, 7: 250, 8: 120, 9: 200 });

  // Type hint row (row 2) styling if present
  if (sheet.getLastRow() >= 2) {
    var row2Val = String(sheet.getRange(2, 1).getValue()).trim().toLowerCase();
    if (row2Val === 'auto-id' || row2Val.indexOf('auto') !== -1) {
      applySubtitleRow_(sheet, 2, numCols);
    }
  }

  // Conditional formatting: Unverified rows (Verified By empty)
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 2, 50);
  var fullRange = sheet.getRange(3, 1, dataRows, numCols);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(LEN($A3)>0, LEN($H3)=0)')
      .setBackground(SHEET_COLORS.BG_LIGHT_YELLOW)
      .setRanges([fullRange])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
  applyRowBanding_(sheet, 3, numCols);
  applyEmptyState_(sheet, 3, numCols,
    'No volunteer hours logged yet. Use 🤝 Volunteer menu → Log Hours to add entries.');

  log_('Formatted', 'Volunteer Hours');
}

// ─── Tab 18: Meeting Attendance ─────────────────────────────────────────────

function formatMeetingAttendanceTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.MEETING_ATTENDANCE);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyColumnWidths_(sheet, { 1: 80, 2: 120, 3: 120, 4: 200, 5: 80, 6: 140, 7: 80, 8: 200 });

  // Type hint row (row 2) if present
  if (sheet.getLastRow() >= 2) {
    var row2Val = String(sheet.getRange(2, 1).getValue()).trim().toLowerCase();
    if (row2Val === 'auto-id' || row2Val.indexOf('auto') !== -1) {
      applySubtitleRow_(sheet, 2, numCols);
    }
  }

  // Conditional formatting: Attended checkbox
  clearConditionalFormats_(sheet);
  var dataRows = Math.max(sheet.getLastRow() - 2, 50);
  var fullRange = sheet.getRange(3, 1, dataRows, numCols);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3=TRUE')
      .setBackground(SHEET_COLORS.BG_PALE_GREEN)
      .setRanges([fullRange])
      .build()
  ];

  // Meeting Type color-coding (col C)
  var typeRange = sheet.getRange(3, 3, dataRows, 1);
  var mtgColors = [
    { text: 'Regular', bg: '#DBEAFE' },
    { text: 'Special', bg: '#FFEDD5' },
    { text: 'Committee', bg: '#CCFBF1' },
    { text: 'Emergency', bg: '#FEE2E2' }
  ];
  for (var m = 0; m < mtgColors.length; m++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mtgColors[m].text)
      .setBackground(mtgColors[m].bg)
      .setRanges([typeRange])
      .build());
  }
  sheet.setConditionalFormatRules(rules);

  applyRowBanding_(sheet, 3, numCols);
  applyEmptyState_(sheet, 3, numCols,
    'No attendance records yet. Use 📅 Meetings menu → Take Attendance to log participation.');

  log_('Formatted', 'Meeting Attendance');
}

// ─── Tab 19: Meeting Check-In Log ───────────────────────────────────────────

function formatMeetingCheckInLogTab_(ss) {
  var sheet = getSheetSafe_(ss, SHEETS.MEETING_CHECKIN_LOG);
  if (!sheet) return;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  applyBrandHeader_(sheet, numCols);
  sheet.setFrozenRows(1);
  applyRowBanding_(sheet, 2, numCols);
  applyEmptyState_(sheet, 2, numCols,
    'No check-ins yet. Check-ins are auto-populated when members sign in via the web portal.');

  log_('Formatted', 'Meeting Check-In Log');
}

// ============================================================================
// TAB BAR COLORS
// ============================================================================

/**
 * Applies tab-bar color grouping to all visible sheets.
 * Blue = data entry, Green = engagement, Gold = documentation, Red = admin.
 * @param {Spreadsheet} ss
 * @private
 */
function applyTabBarColors_(ss) {
  var blue   = SHEET_COLORS.TAB_BLUE;
  var green  = SHEET_COLORS.TAB_GREEN;
  var gold   = SHEET_COLORS.TAB_GOLD;
  var red    = SHEET_COLORS.TAB_RED_ORANGE;

  // Blue — Data entry tabs
  var blueSheets = [
    SHEETS.PORTAL_EVENTS, SHEETS.PORTAL_MINUTES, SHEETS.MEETING_ATTENDANCE,
    SHEETS.MEETING_CHECKIN_LOG, SHEETS.VOLUNTEER_HOURS, SHEETS.WORKLOAD_REPORTING,
    SHEETS.NON_MEMBER_CONTACTS, SHEETS.CASE_CHECKLIST
  ];

  // Green — Engagement / survey tabs
  var greenSheets = [
    SHEETS.SURVEY_QUESTIONS, SHEETS.SATISFACTION, SHEETS.NOTIFICATIONS
  ];

  // Gold — Documentation / guide tabs
  var goldSheets = [
    SHEETS.GETTING_STARTED, SHEETS.FAQ, SHEETS.RESOURCES,
    SHEETS.RESOURCE_CONFIG, SHEETS.FEATURES_REFERENCE, SHEETS.CONFIG_GUIDE
  ];

  // Red-Orange — Admin / technical tabs
  var redSheets = [
    SHEETS.SETTINGS_OVERVIEW, SHEETS.CONFIG
  ];

  var groups = [
    { names: blueSheets, color: blue },
    { names: greenSheets, color: green },
    { names: goldSheets, color: gold },
    { names: redSheets, color: red }
  ];

  for (var g = 0; g < groups.length; g++) {
    for (var n = 0; n < groups[g].names.length; n++) {
      var sheet = getSheetSafe_(ss, groups[g].names[n]);
      if (sheet) {
        sheet.setTabColor(groups[g].color);
      }
    }
  }

  // Core data sheets get a distinct purple
  var coreSheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];
  for (var c = 0; c < coreSheets.length; c++) {
    var coreSheet = getSheetSafe_(ss, coreSheets[c]);
    if (coreSheet) coreSheet.setTabColor('#6A1B9A');
  }

  log_('applyTabBarColors_', 'Tab bar colors applied');
}

// ============================================================================
// MASTER FUNCTION — Apply Union Theme to All Tabs
// ============================================================================

/**
 * Applies the Union brand theme to all visible tabs.
 * Callable from Admin menu: 🎨 Apply Union Theme.
 * Idempotent — safe to run multiple times.
 */
function applyUnionThemeToAllTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Applying Union brand theme to all tabs...', '🎨 Theme', 10);

  try {
    // Phase 1: Format each tab
    formatGettingStartedTab_(ss);
    formatFAQTab_(ss);
    formatSurveyQuestionsTab_(ss);
    formatNotificationsTab_(ss);
    formatResourcesTab_(ss);
    formatResourceConfigTab_(ss);
    formatSettingsOverviewTab_(ss);
    formatEventsTab_(ss);
    formatMeetingMinutesTab_(ss);
    formatWorkloadReportingTab_(ss);
    formatNonMemberContactsTab_(ss);
    formatCaseChecklistTab_(ss);
    formatMemberSatisfactionTab_(ss);
    formatFeaturesReferenceTab_(ss);
    formatVolunteerHoursTab_(ss);
    formatMeetingAttendanceTab_(ss);
    formatMeetingCheckInLogTab_(ss);

    // Phase 2: Tab bar colors
    applyTabBarColors_(ss);

    ss.toast('Union brand theme applied to all tabs!', '✅ Theme Complete', 5);
    log_('applyUnionThemeToAllTabs', 'completed successfully');

  } catch (error) {
    log_('applyUnionThemeToAllTabs error', error.message);
    ss.toast('Theme error: ' + error.message, '❌ Error', 5);
  }
}

// Dead code removed: formatSingleTab() — zero callers in src

// ============================================================================
// DASHBOARD ENHANCEMENTS (merged from 16_DashboardEnhancements.gs)
// ============================================================================

// ============================================================================
// 1. CUSTOM DATE RANGES - Server-side helpers
// ============================================================================
// ============================================================================
// 2. EXPORT INDIVIDUAL CHARTS AS IMAGES - Server-side download handler
// ============================================================================

// Dead code removed: saveChartImageToDrive() — zero callers in src

// Dead code removed: getOrCreateExportFolder_() — zero callers in src

// ============================================================================
// 3. SCHEDULED EMAIL REPORTS
// ============================================================================
// Dead code removed: scheduleEmailReport(), sendScheduledReports(),
// sendDashboardReportEmail_(), installReportTrigger_(), getScheduledReports(),
// removeScheduledReport() — entire scheduled reports subsystem, never wired up

// ============================================================================
// 4. NOTIFICATIONS (v4.22.0 — rerouted to sheet-based system in 05_Integrations.gs)
// ============================================================================
// getUserNotifications() and markNotificationRead() removed — used ScriptProperties
// storage which is orphaned by the Notifications sheet system (v4.13.0+).
// broadcastStewardNotification() removed — no active callers; steward compose
// form handles broadcast directly via sendWebAppNotification().
//
// pushNotification() is kept because saveSharedView() (below) calls it
// and 05_Integrations.gs also calls it.
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
    log_('pushNotification error', e.message);
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

// Dead code removed: getSharedViews(), deleteSharedView() — zero callers in src

// ============================================================================
// 6. SAVED CHART CONFIGURATIONS (PRESETS)
// ============================================================================
// Dead code removed: saveChartPreset(), getChartPresets(), deleteChartPreset() — never wired up

// ============================================================================
// 7. ADVANCED FILTERING OPTIONS
// ============================================================================

// Dead code removed: getFilteredDashboardData() — zero callers in src

// ============================================================================
// 8. DRILL-DOWN CAPABILITIES (Multi-level hierarchical)
// ============================================================================

// ============================================================================
// EXECUTIVE DASHBOARD (merged from 04d_ExecutiveDashboard.gs)
// ============================================================================

// ============================================================================
// 1. NAVIGATION HELPERS
// ============================================================================

// navigateToSheet() - REMOVED DUPLICATE - see line 565 for main definition

// ============================================================================
// 2. EXECUTIVE COMMAND MODAL (SPA Architecture - Bridge Pattern)
// ============================================================================

/**
 * High-Performance KPI Aggregator for the Modal (Bridge Pattern)
 * Returns JSON data for client-side rendering
 * @returns {string} JSON string with dashboard statistics
 */
function getDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var stats = {
    totalGrievances: 0,
    activeGrievances: 0,
    activeSteps: { step1: 0, step2: 0, arbitration: 0 },
    outcomes: { wins: 0, losses: 0, settled: 0, withdrawn: 0 },
    winRate: 0,
    overdueCount: 0,
    totalMembers: 0,
    stewardCount: 0,
    moraleScore: 5, // Default; overridden by getSatisfactionSummary() below
    unitBreakdown: {},
    stewardWorkload: []
  };

  // Process Grievance Log
  if (logSheet && logSheet.getLastRow() > 1) {
    var logData = logSheet.getDataRange().getValues();
    logData.shift(); // Remove headers

    // Filter to only rows with a valid grievance ID (starts with "G")
    logData = logData.filter(function(row) {
      var gid = (col_(row, GRIEVANCE_COLS.GRIEVANCE_ID) || '').toString();
      return isGrievanceId_(gid);
    });
    stats.totalGrievances = logData.length;

    logData.forEach(function(row) {
      var status = (col_(row, GRIEVANCE_COLS.STATUS) || '').toString().toLowerCase();
      var currentStep = (col_(row, GRIEVANCE_COLS.CURRENT_STEP) || '').toString().toLowerCase();
      var unit = col_(row, GRIEVANCE_COLS.LOCATION) || 'Unknown';
      var steward = col_(row, GRIEVANCE_COLS.STEWARD) || 'Unassigned';

      // Count active vs closed
      if (status === 'open' || status === 'pending info' || status === 'appealed') {
        stats.activeGrievances++;
      }

      // M7: Count by step — use word-boundary matching to prevent 'step 10' matching 'step 1'
      if (/\bstep\s*1\b/.test(currentStep) || currentStep === '1' || currentStep === 'step i') stats.activeSteps.step1++;
      if (/\bstep\s*2\b/.test(currentStep) || currentStep === '2' || currentStep === 'step ii') stats.activeSteps.step2++;
      if (/\barbitration\b/.test(currentStep) || /\bstep\s*3\b/.test(currentStep) || currentStep === 'step iii') stats.activeSteps.arbitration++;

      // Count outcomes using the Resolution column (not Status)
      var resolution = (col_(row, GRIEVANCE_COLS.RESOLUTION) || '').toString().toLowerCase();
      if (resolution === 'won' || resolution === 'sustained') stats.outcomes.wins++;
      if (resolution === 'denied' || resolution === 'lost') stats.outcomes.losses++;
      if (resolution === 'settled') stats.outcomes.settled++;
      if (resolution === 'withdrawn') stats.outcomes.withdrawn++;

      // Unit breakdown
      if (!stats.unitBreakdown[unit]) stats.unitBreakdown[unit] = 0;
      stats.unitBreakdown[unit]++;

      // Steward workload (only active cases)
      if (status === 'open' || status === 'pending info') {
        var existingSteward = stats.stewardWorkload.find(function(s) { return s.name === steward; });
        if (existingSteward) {
          existingSteward.count++;
        } else {
          stats.stewardWorkload.push({ name: steward, count: 1 });
        }
      }

      // Check for overdue — use deadline matching the current step
      if (status === 'open' || status === 'pending info') {
        var dueDate = null;
        if (/\bstep\s*2\b/.test(currentStep) || currentStep === '2' || currentStep === 'step ii') {
          dueDate = col_(row, GRIEVANCE_COLS.STEP2_DUE);
        } else if (/\barbitration\b/.test(currentStep) || /\bstep\s*3\b/.test(currentStep) || currentStep === 'step iii') {
          dueDate = col_(row, GRIEVANCE_COLS.STEP3_APPEAL_DUE);
        } else {
          dueDate = col_(row, GRIEVANCE_COLS.STEP1_DUE);
        }
        if (dueDate && new Date(dueDate) < new Date()) {
          stats.overdueCount++;
        }
      }
    });

    // Calculate win rate — only resolved outcomes in denominator
    var totalResolved = stats.outcomes.wins + stats.outcomes.losses + stats.outcomes.settled + stats.outcomes.withdrawn;
    if (totalResolved > 0) {
      stats.winRate = Math.round((stats.outcomes.wins / totalResolved) * 100);
    }

    // Sort steward workload
    stats.stewardWorkload.sort(function(a, b) { return b.count - a.count; });
  }

  // Process Member Directory
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var m = 1; m < memberData.length; m++) {
      if (col_(memberData[m], MEMBER_COLS.MEMBER_ID)) {
        stats.totalMembers++;
        if (isTruthyValue(col_(memberData[m], MEMBER_COLS.IS_STEWARD))) stats.stewardCount++;
      }
    }
  }

  // Get morale score from satisfaction data (v4.23.0: dynamic via getSatisfactionSummary)
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet && satSheet.getLastRow() > 1) {
    try {
      var summary = getSatisfactionSummary();
      if (summary && summary.sections && summary.sections['OVERALL_SAT']) {
        var overallAvg = summary.sections['OVERALL_SAT'].avg;
        if (overallAvg !== null && !isNaN(overallAvg)) {
          stats.moraleScore = Math.round(overallAvg * 10) / 10;
        }
      }
    } catch(e) {
      log_('getDashboardStats', 'Error reading satisfaction summary for morale score: ' + e.message);
    }
  }

  return JSON.stringify(stats);
}
/**
 * Gets executive metrics from dashboard calculations
 * @returns {Object} Metrics object with activeGrievances, winRate, overdueSteps
 * @private
 */
function getExecutiveMetrics_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC);

  var metrics = {
    activeGrievances: 0,
    activeTrend: "Stable",
    winRate: 0,
    winRateTrend: "Stable",
    overdueSteps: 0,
    overdueTrend: "Stable"
  };

  if (calcSheet) {
    try {
      // Pull values from calculation sheet
      var data = calcSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === 'Active Grievances') metrics.activeGrievances = data[i][1] || 0;
        if (data[i][0] === 'Win Rate') metrics.winRate = Math.round((data[i][1] || 0) * 100);
        if (data[i][0] === 'Overdue') metrics.overdueSteps = data[i][1] || 0;
      }
    } catch (e) {
      log_('getExecutiveMetrics_', 'Error: ' + e.message);
    }
  }

  // Try to get from grievance sheet if calc sheet not available
  if (metrics.activeGrievances === 0) {
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      var grievanceData = grievanceSheet.getDataRange().getValues();
      var openCount = 0;
      var wonCount = 0;
      var closedCount = 0;
      var overdueCount = 0;

      for (var g = 1; g < grievanceData.length; g++) {
        var status = col_(grievanceData[g], GRIEVANCE_COLS.STATUS);
        if (status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING) openCount++;
        if (status === GRIEVANCE_STATUS.WON) wonCount++;
        if (GRIEVANCE_CLOSED_STATUSES.indexOf(status) !== -1) closedCount++;

        var daysToDeadline = col_(grievanceData[g], GRIEVANCE_COLS.DAYS_TO_DEADLINE);
        if (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0)) {
          overdueCount++;
        }
      }

      metrics.activeGrievances = openCount;
      metrics.winRate = closedCount > 0 ? Math.round((wonCount / closedCount) * 100) : 0;
      metrics.overdueSteps = overdueCount;
    }
  }

  return metrics;
}

// ============================================================================
// 3. UNIFIED STEWARD DASHBOARD (v4.3.2)
// ============================================================================
// Consolidates all analytics, charts, and reports into a single tabbed interface.
// This replaces the individual chart modals for a unified experience.
// ============================================================================

/**
 * Shows the unified Steward Dashboard (web app URL)
 * Opens the unified dashboard in steward mode (with PII)
 * v4.4.0: Now opens as web app instead of modal for better experience
 */
function showStewardDashboard() {
  var url = ScriptApp.getService().getUrl() + '?mode=steward';
  showDialog_(
    '<html><head>' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;background:#0f172a;color:#f8fafc;margin:0;padding:16px}' +
    '.icon{font-size:clamp(36px,10vw,48px);margin-bottom:16px}h1{margin:0 0 8px;font-size:clamp(18px,5vw,24px);text-align:center}p{color:#94a3b8;margin:0 0 24px;text-align:center;max-width:400px;line-height:1.5;font-size:clamp(13px,3.5vw,15px);padding:0 8px}' +
    'a.open-link{background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:600;display:inline-block;min-height:44px;line-height:20px;text-align:center}a.open-link:hover{background:#2563eb}' +
    '.copy-btn{background:#475569;cursor:pointer;border:none;padding:10px 16px;border-radius:8px;color:white;font-size:clamp(11px,3vw,13px);min-height:44px}' +
    '.url{background:#1e293b;padding:12px;border-radius:8px;font-family:monospace;font-size:clamp(10px,2.5vw,12px);word-break:break-all;max-width:90%;margin-bottom:16px;border:1px solid #334155;width:100%}' +
    '.warning{background:rgba(239,68,68,0.2);color:#fca5a5;padding:8px 16px;border-radius:8px;font-size:clamp(10px,2.5vw,12px);margin-bottom:16px}' +
    '.btn-row{display:flex;gap:8px;flex-wrap:wrap;justify-content:center}' +
    '@media(max-width:480px){.btn-row{flex-direction:column;width:100%}a.open-link,.copy-btn{width:100%;text-align:center}}' +
    '</style></head><body><div class="icon">🛡️</div><h1>Steward Command Center</h1>' +
    '<div class="warning">INTERNAL USE ONLY - Contains PII</div>' +
    '<p>Open the Steward Dashboard web app. This version includes full member details and sensitive information.</p>' +
    '<div class="url" id="url">' + escapeHtml(url) + '</div>' +
    '<div class="btn-row"><a class="open-link" href="' + escapeHtml(url) + '" target="_blank">Open Dashboard</a>' +
    '<button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'url\').textContent);this.textContent=\'Copied!\';setTimeout(function(){document.querySelector(\'.copy-btn\').textContent=\'Copy URL\'},2000)">Copy URL</button></div>' +
    '</body></html>',
    'Steward Command Center', 500, 400);
}
// ============================================================================
// 4. STRATEGIC PRO MOVES & ALERTS
// ============================================================================

/**
 * Checks dashboard metrics and sends email alerts if thresholds exceeded
 */
function checkDashboardAlerts() {
  var metrics = getExecutiveMetrics_();
  var winRate = metrics.winRate;

  var alerts = [];
  if (winRate < 50) alerts.push("CRITICAL WIN RATE: " + winRate + "%");
  if (metrics.overdueSteps > 10) alerts.push("HIGH OVERDUE COUNT: " + metrics.overdueSteps + " cases");

  if (alerts.length > 0) {
    try {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        "DASHBOARD ALERT",
        "The following alerts require attention:\n\n" + alerts.join("\n")
      );
      SpreadsheetApp.getActiveSpreadsheet().toast("Alert email sent", "Notification");
    } catch (e) {
      log_('checkDashboardAlerts', 'Error sending alert email: ' + e.message);
    }
  }
}

// ============================================================================
// 5. AUTOMATION & COMMUNICATION
// ============================================================================

/**
 * Creates automation triggers for nightly refresh
 */
function createAutomationTriggers() {
  // Check for existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshAllVisuals') {
      SpreadsheetApp.getUi().alert('Automation trigger already exists. No action needed.');
      return;
    }
  }

  // Midnight Refresh Trigger
  ScriptApp.newTrigger('refreshAllVisuals')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  SpreadsheetApp.getUi().alert('Success: Dashboard will now auto-refresh every night at 1:00 AM.');
}

/**
 * Sets up the midnight auto-refresh trigger
 * Runs midnightAutoRefresh daily at midnight (12:00 AM)
 */
function setupMidnightTrigger() {
  // Check for existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  var hasExisting = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'midnightAutoRefresh') {
      hasExisting = true;
      break;
    }
  }

  if (hasExisting) {
    SpreadsheetApp.getUi().alert('Midnight auto-refresh trigger already exists. No action needed.');
    return;
  }

  // Create midnight trigger (12:00 AM)
  ScriptApp.newTrigger('midnightAutoRefresh')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert('Success: Midnight Auto-Refresh is now active.\n\nThe system will automatically refresh all dashboards and check alerts at 12:00 AM daily.');
}

/**
 * Removes the midnight auto-refresh trigger
 */
function removeMidnightTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'midnightAutoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed = true;
    }
  }

  if (removed) {
    SpreadsheetApp.getUi().alert('Midnight auto-refresh trigger has been removed.');
  } else {
    SpreadsheetApp.getUi().alert('No midnight trigger found to remove.');
  }
}

/**
 * Midnight Auto-Refresh function
 * Called automatically by time-based trigger at midnight
 * Refreshes dashboards, hidden sheets, and checks for critical alerts
 */
function midnightAutoRefresh() {
  try {
    var startTime = new Date();

    log_('midnightAutoRefresh', 'Midnight Auto-Refresh started at ' + startTime.toISOString());

    // 1. Check for critical dashboard alerts
    checkDashboardAlerts();
    log_('midnightAutoRefresh', 'Dashboard alerts checked');

    // 2. Check for overdue grievances and send reminders
    checkOverdueGrievances_();

    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;

    log_('midnightAutoRefresh', 'Midnight Auto-Refresh completed in ' + duration + ' seconds');

    // Note: Executive Command and Member Analytics are now modal-based
    // and don't require midnight refresh - data is fetched on-demand

  } catch (e) {
    log_('midnightAutoRefresh', 'Error: ' + e.message);
    // Optionally send error notification
    try {
      var adminEmail = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
      if (adminEmail) {
        MailApp.sendEmail(adminEmail,
          COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Auto-Refresh Error',
          'The midnight auto-refresh encountered an error:\n\n' + e.message + '\n\nPlease check the script logs for details.');
      }
    } catch (emailErr) {
      log_('Could not send error notification', emailErr.message);
    }
  }
}

/**
 * Checks for overdue grievances and sends reminder notifications
 * @private
 */
function checkOverdueGrievances_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet || grievanceSheet.getLastRow() < 2) return;

    var data = grievanceSheet.getDataRange().getValues();
    var overdueList = [];

    for (var i = 1; i < data.length; i++) {
      var status = col_(data[i], GRIEVANCE_COLS.STATUS);
      var daysToDeadline = col_(data[i], GRIEVANCE_COLS.DAYS_TO_DEADLINE);

      // Check for active cases that are overdue
      if ((status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING) &&
          (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0))) {
        overdueList.push({
          id: col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID),
          name: maskName(col_(data[i], GRIEVANCE_COLS.FIRST_NAME), col_(data[i], GRIEVANCE_COLS.LAST_NAME)),
          steward: col_(data[i], GRIEVANCE_COLS.STEWARD),
          days: daysToDeadline
        });
      }
    }

    if (overdueList.length > 0) {
      var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
      if (chiefStewardEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(chiefStewardEmail)) {
        var body = 'DAILY OVERDUE GRIEVANCE REPORT\n\n' +
                   'The following ' + overdueList.length + ' grievance(s) have passed their deadline:\n\n';

        overdueList.forEach(function(g) {
          body += '- ' + g.id + ': ' + g.name + ' (Steward: ' + (g.steward || 'Unassigned') + ')\n';
        });

        body += '\nPlease take immediate action to address these cases.' + COMMAND_CONFIG.EMAIL.FOOTER;

        MailApp.sendEmail(chiefStewardEmail,
          COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Daily Overdue Report - ' + overdueList.length + ' Case(s)',
          body);

        log_('checkOverdueGrievances_', 'Sent overdue report with ' + overdueList.length + ' cases');
      }
    }
  } catch (e) {
    log_('checkOverdueGrievances_', 'Error: ' + e.message);
  }
}

/**
 * Emails the Executive Dashboard as a PDF snapshot
 */
function emailExecutivePDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var date = new Date();
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Export only the active sheet as PDF (not the entire spreadsheet)
    var activeSheet = ss.getActiveSheet();
    var sheetGid = activeSheet.getSheetId();
    var url = ss.getUrl().replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf' +
      '&gid=' + sheetGid +
      '&size=letter&portrait=true&fitw=1' +
      '&gridlines=false&printtitle=false&sheetnames=false';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token }
    });
    var blob = response.getBlob().setName("Health_Report_" + dateStr + ".pdf");

    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + " Weekly Executive Summary - " + dateStr,
      body: "Attached is the latest strategic briefing." + COMMAND_CONFIG.EMAIL.FOOTER,
      attachments: [blob]
    });

    SpreadsheetApp.getUi().alert('PDF report sent to your email.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error generating PDF: ' + e.message);
  }
}
