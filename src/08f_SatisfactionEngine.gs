/**
 * 509 Dashboard - Satisfaction Engine Module
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
 * @version 1.0.0
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
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
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

    // Responsive
    '@media (max-width:600px){' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .list-item{flex-direction:column;align-items:flex-start}' +
    '  .tab-icon{font-size:16px}' +
    '  .bar-label{width:100px}' +
    '  .gauge-container{flex-direction:column;align-items:center}' +
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
    '      insights+="<div class=\\"insight-card "+i.type+"\\"><div class=\\"insight-title\\">"+i.icon+" "+i.title+"</div><div class=\\"insight-text\\">"+i.text+"</div></div>";' +
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
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+r.worksite+" - "+r.role+"</div><div class=\\"list-item-subtitle\\">"+r.shift+" • "+r.timeInRole+" • "+r.date+"</div></div><div><span class=\\"score-indicator score-"+scoreClass+"\\" style=\\"color:"+scoreColor+"\\">"+r.avgScore.toFixed(1)+"/10</span></div></div>' +
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
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+s.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+s.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+s.responseCount+" responses</div></div>";' +
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
    '      lowScoring.forEach(function(s,i){html+=(i>0?", ":"")+s.name+" ("+s.avg.toFixed(1)+")"});' +
    '      html+="</div></div>";' +
    '    }' +
    '    if(highScoring.length>0){' +
    '      html+="<div class=\\"insight-card success\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">✅ Strong Performance</div><div class=\\"insight-text\\">";' +
    '      highScoring.forEach(function(s,i){html+=(i>0?", ":"")+s.name+" ("+s.avg.toFixed(1)+")"});' +
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
    '      html+="<div class=\\"insight-card "+i.type+"\\" style=\\"margin-bottom:10px\\"><div class=\\"insight-title\\">"+i.icon+" "+i.title+"</div><div class=\\"insight-text\\">"+i.text+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No insights available</div>";}' +
    '  html+="</div>";' +
    // By worksite breakdown
    '  if(data.byWorksite&&data.byWorksite.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Satisfaction by Worksite</div><div class=\\"bar-chart\\">";' +
    '    data.byWorksite.forEach(function(w){' +
    '      var pct=(w.avg/10)*100;' +
    '      var color=getScoreColor(w.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+w.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+w.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+w.count+" responses</div></div>";' +
    '    });' +
    '    html+="</div></div>";' +
    '  }' +
    // By role breakdown
    '  if(data.byRole&&data.byRole.length>0){' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👤 Satisfaction by Role</div><div class=\\"bar-chart\\">";' +
    '    data.byRole.forEach(function(r){' +
    '      var pct=(r.avg/10)*100;' +
    '      var color=getScoreColor(r.avg);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+r.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"><span class=\\"bar-inner-value\\">"+r.avg.toFixed(1)+"</span></div></div><div class=\\"bar-value\\">"+r.count+" responses</div></div>";' +
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
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">"+p.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:#7C3AED\\"></div></div><div class=\\"bar-value\\">"+p.count+"</div></div>";' +
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
  var lastRow = 1;
  var timestamps = sheet.getRange('A:A').getValues();
  for (var i = 1; i < timestamps.length; i++) {
    if (timestamps[i][0] === '' || timestamps[i][0] === null) {
      lastRow = i;
      break;
    }
    lastRow = i + 1;
  }

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

  // Filter to only verified and latest responses
  var validRows = data.filter(function(row) {
    var verified = row[SATISFACTION_COLS.VERIFIED - 1];
    var isLatest = row[SATISFACTION_COLS.IS_LATEST - 1];
    return verified === 'Yes' && isLatest === 'Yes';
  });

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

    // Track by quarter for trend
    var quarter = row[SATISFACTION_COLS.QUARTER - 1];
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
  var lastRow = 1;
  var timestamps = sheet.getRange('A:A').getValues();
  for (var i = 1; i < timestamps.length; i++) {
    if (timestamps[i][0] === '' || timestamps[i][0] === null) {
      lastRow = i;
      break;
    }
    lastRow = i + 1;
  }

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
      stewardContact: stewardContact === 'Yes',
      stewardRating: stewardRating
    });
  }

  // Sort by date (most recent first)
  responses.sort(function(a, b) {
    return b.date.localeCompare(a.date);
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
  var lastRow = 1;
  var timestamps = sheet.getRange('A:A').getValues();
  for (var i = 1; i < timestamps.length; i++) {
    if (timestamps[i][0] === '' || timestamps[i][0] === null) {
      lastRow = i;
      break;
    }
    lastRow = i + 1;
  }

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
  } catch(e) { /* ignore if column doesn't exist */ }

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
  var lastRow = 1;
  var timestamps = sheet.getRange('A:A').getValues();
  for (var i = 1; i < timestamps.length; i++) {
    if (timestamps[i][0] === '' || timestamps[i][0] === null) {
      lastRow = i;
      break;
    }
    lastRow = i + 1;
  }

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
  var timestamps = sheet.getRange('A:A').getValues();
  for (var i = 1; i < timestamps.length; i++) {
    if (timestamps[i][0] === '' || timestamps[i][0] === null) {
      return i;
    }
  }
  return timestamps.length;
}
