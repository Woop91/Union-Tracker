/**
 * ============================================================================
 * WEB APP DEPLOYMENT FOR MOBILE ACCESS
 * ============================================================================
 * This file enables the dashboard to be deployed as a standalone web app
 * that can be accessed directly via URL on mobile devices.
 *
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Go to Extensions → Apps Script
 * 2. Click "Deploy" → "New deployment"
 * 3. Select "Web app" as the deployment type
 * 4. Set "Execute as" to your account
 * 5. Set "Who has access" to your organization or anyone
 * 6. Click "Deploy" and copy the URL
 * 7. Bookmark this URL on your mobile device for easy access
 */

/**
 * Web app entry point - serves the mobile dashboard
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} The HTML page to display
 */
function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'dashboard';

  var html;
  switch (page) {
    case 'search':
      html = getWebAppSearchHtml();
      break;
    case 'grievances':
      html = getWebAppGrievanceListHtml();
      break;
    case 'members':
      html = getWebAppMemberListHtml();
      break;
    case 'links':
      html = getWebAppLinksHtml();
      break;
    case 'dashboard':
    default:
      html = getWebAppDashboardHtml();
      break;
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle('509 Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Returns the main dashboard HTML for web app (enhanced with clickable stats, win rate, overdue preview)
 */
function getWebAppDashboardHtml() {
  var stats = getWebAppDashboardStats();
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">' +
    '<link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>📊</text></svg>">' +
    '<title>509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h1{font-size:clamp(20px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Stats grid - clickable cards
    '.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:15px 10px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;cursor:pointer;transition:transform 0.2s;text-decoration:none;display:block}' +
    '.stat-card:active{transform:scale(0.96)}' +
    '.stat-value{font-size:clamp(22px,6vw,32px);font-weight:bold;color:#7C3AED}' +
    '.stat-value.warning{color:#F97316}' +
    '.stat-value.danger{color:#DC2626}' +
    '.stat-value.success{color:#059669}' +
    '.stat-label{font-size:clamp(9px,2.2vw,11px);color:#666;text-transform:uppercase;margin-top:4px;letter-spacing:0.3px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Overdue preview
    '.overdue-section{background:#FEF2F2;border-left:4px solid #DC2626;border-radius:12px;padding:15px;margin-bottom:20px}' +
    '.overdue-title{font-size:14px;font-weight:600;color:#DC2626;margin-bottom:10px;display:flex;align-items:center;gap:6px}' +
    '.overdue-item{background:white;padding:12px;border-radius:10px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}' +
    '.overdue-item:last-child{margin-bottom:0}' +
    '.overdue-id{font-weight:600;color:#7C3AED;font-size:13px}' +
    '.overdue-name{font-size:14px;color:#333;margin-top:2px}' +
    '.overdue-detail{font-size:12px;color:#666;margin-top:2px}' +
    '.view-all-btn{background:#DC2626;color:white;border:none;padding:10px;border-radius:8px;font-size:13px;font-weight:500;width:100%;margin-top:10px;cursor:pointer}' +

    // Action buttons
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{background:white;border:none;padding:18px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:16px;cursor:pointer;' +
    'text-decoration:none;color:inherit;min-height:64px;transition:all 0.2s}' +
    '.action-btn:active{transform:scale(0.98);background:#f0f0f0}' +
    '.action-icon{font-size:28px;width:50px;height:50px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#EDE9FE,#DDD6FE);border-radius:14px;flex-shrink:0}' +
    '.action-label{font-weight:600;color:#333}' +
    '.action-desc{font-size:13px;color:#666;margin-top:3px}' +

    // Loading state
    '.loading{text-align:center;padding:20px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:20px;height:20px;border:2px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    // Refresh indicator
    '.refresh-btn{position:absolute;right:15px;top:50%;transform:translateY(-50%);background:rgba(255,255,255,0.2);border:none;color:white;width:40px;height:40px;border-radius:50%;font-size:20px;cursor:pointer}' +

    '</style></head><body>' +

    // Header
    '<div class="header">' +
    '<button class="refresh-btn" onclick="location.reload()">🔄</button>' +
    '<h1>📊 509 Dashboard</h1>' +
    '<div class="subtitle">Union Grievance Management</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section - 6 clickable stats in 3x2 grid
    '<div class="stats">' +
    '<a class="stat-card" href="' + baseUrl + '?page=members"><div class="stat-value">' + stats.totalMembers + '</div><div class="stat-label">Members</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Grievances</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=open"><div class="stat-value success">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=pending"><div class="stat-value warning">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=overdue"><div class="stat-value danger">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></a>' +
    '<div class="stat-card"><div class="stat-value success">' + stats.winRate + '</div><div class="stat-label">Win Rate</div></div>' +
    '</div>' +

    // Overdue preview section (loaded dynamically)
    '<div id="overdue-preview"></div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<a class="action-btn" href="' + baseUrl + '?page=search">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find members or grievances</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=grievances">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">All Grievances</div><div class="action-desc">Browse and filter grievances</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=members">' +
    '<div class="action-icon">👥</div>' +
    '<div><div class="action-label">Members</div><div class="action-desc">View member directory</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=links">' +
    '<div class="action-icon">🔗</div>' +
    '<div><div class="action-label">Links</div><div class="action-desc">Forms, resources, GitHub</div></div>' +
    '</a>' +

    '</div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item active" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    // Script to load overdue preview
    '<script>' +
    'var baseUrl="' + baseUrl + '";' +
    'function loadOverdue(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    var overdue=data.filter(function(g){return g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overdue-preview").innerHTML="";return}' +
    '    var html="<div class=\\"overdue-section\\"><div class=\\"overdue-title\\">⚠️ Overdue Cases ("+overdue.length+")</div>";' +
    '    overdue.slice(0,3).forEach(function(g){' +
    '      html+="<div class=\\"overdue-item\\"><div class=\\"overdue-id\\">"+g.id+"</div><div class=\\"overdue-name\\">"+g.name+"</div><div class=\\"overdue-detail\\">"+g.category+" • "+g.step+"</div></div>";' +
    '    });' +
    '    if(overdue.length>3)html+="<button class=\\"view-all-btn\\" onclick=\\"location.href=baseUrl+\'?page=grievances&filter=overdue\'\\">View All "+overdue.length+" Overdue Cases</button>";' +
    '    html+="</div>";' +
    '    document.getElementById("overdue-preview").innerHTML=html;' +
    '  }).getWebAppGrievanceList();' +
    '}' +
    'loadOverdue();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the search page HTML for web app
 */
function getWebAppSearchHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Search - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);margin-bottom:12px;text-align:center}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:14px 14px 14px 45px;border:none;border-radius:12px;font-size:16px;background:white;-webkit-appearance:none}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#666}' +
    '.clear-btn{position:absolute;right:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#999;background:none;border:none;cursor:pointer;display:none}' +

    // Tabs
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0;position:sticky;top:76px;z-index:99}' +
    '.tab{flex:1;padding:14px;text-align:center;font-size:14px;font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:48px}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED}' +

    // Results
    '.results{padding:15px}' +
    '.result-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px}' +
    '.result-type{font-size:12px;color:#7C3AED;font-weight:600;text-transform:uppercase;margin-bottom:6px}' +
    '.result-title{font-size:17px;font-weight:600;color:#333;margin-bottom:4px}' +
    '.result-detail{font-size:14px;color:#666;margin-top:4px}' +
    '.result-badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;margin-top:8px}' +
    '.badge-open{background:#FEE2E2;color:#DC2626}' +
    '.badge-pending{background:#FEF3C7;color:#D97706}' +
    '.badge-resolved{background:#D1FAE5;color:#059669}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +
    '.empty-text{font-size:16px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔍 Search</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search members or grievances..." oninput="handleSearch(this.value)" autocomplete="off" autocapitalize="off">' +
    '<button class="clear-btn" id="clearBtn" onclick="clearSearch()">✕</button>' +
    '</div></div>' +

    '<div class="tabs">' +
    '<button class="tab active" data-tab="all" onclick="setTab(\'all\',this)">All</button>' +
    '<button class="tab" data-tab="members" onclick="setTab(\'members\',this)">Members</button>' +
    '<button class="tab" data-tab="grievances" onclick="setTab(\'grievances\',this)">Grievances</button>' +
    '</div>' +

    '<div class="results" id="results">' +
    '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">Type to search members or grievances</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var currentTab="all";' +
    'var searchTimeout=null;' +
    'var lastQuery="";' +

    'function setTab(tab,btn){' +
    '  currentTab=tab;' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  if(lastQuery.length>=2)performSearch(lastQuery);' +
    '}' +

    'function handleSearch(q){' +
    '  lastQuery=q;' +
    '  document.getElementById("clearBtn").style.display=q?"block":"none";' +
    '  if(searchTimeout)clearTimeout(searchTimeout);' +
    '  if(!q||q.length<2){' +
    '    showEmpty("Type at least 2 characters to search");' +
    '    return;' +
    '  }' +
    '  showLoading();' +
    '  searchTimeout=setTimeout(function(){performSearch(q)},300);' +
    '}' +

    'function performSearch(q){' +
    '  google.script.run.withSuccessHandler(renderResults).withFailureHandler(showError).getWebAppSearchResults(q,currentTab);' +
    '}' +

    'function clearSearch(){' +
    '  document.getElementById("searchInput").value="";' +
    '  document.getElementById("clearBtn").style.display="none";' +
    '  lastQuery="";' +
    '  showEmpty("Type to search members or grievances");' +
    '}' +

    'function showEmpty(msg){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">🔍</div><div class=\\"empty-text\\">"+msg+"</div></div>";' +
    '}' +

    'function showLoading(){' +
    '  document.getElementById("results").innerHTML="<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:15px\\">Searching...</div></div>";' +
    '}' +

    'function showError(err){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error: "+(err.message||"Unknown error")+"</div></div>";' +
    '}' +

    'function getBadgeClass(status){' +
    '  if(!status)return"";' +
    '  var s=status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"badge-open";' +
    '  if(s.indexOf("pending")>=0)return"badge-pending";' +
    '  if(s.indexOf("resolved")>=0||s.indexOf("closed")>=0||s.indexOf("withdrawn")>=0)return"badge-resolved";' +
    '  return"";' +
    '}' +

    'function renderResults(data){' +
    '  var c=document.getElementById("results");' +
    '  if(!data||data.length===0){' +
    '    showEmpty("No results found");' +
    '    return;' +
    '  }' +
    '  c.innerHTML=data.map(function(r){' +
    '    var badge=r.status?"<span class=\\"result-badge "+getBadgeClass(r.status)+"\\">"+r.status+"</span>":"";' +
    '    return"<div class=\\"result-card\\">"+"<div class=\\"result-type\\">"+(r.type==="member"?"👤 Member":"📋 Grievance")+"</div>"+"<div class=\\"result-title\\">"+r.title+"</div>"+"<div class=\\"result-detail\\">"+r.subtitle+"</div>"+(r.detail?"<div class=\\"result-detail\\">"+r.detail+"</div>":"")+badge+"</div>";' +
    '  }).join("");' +
    '}' +

    // Auto-focus search on load
    'document.getElementById("searchInput").focus();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the grievance list HTML for web app (enhanced with Overdue filter, expandable details)
 */
function getWebAppGrievanceListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Grievances - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px 15px 12px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +

    // Filter pills with Overdue
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 16px;border-radius:20px;font-size:13px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +
    '.filter-pill.danger{background:#DC2626;color:white}' +
    '.filter-pill.danger.active{background:#FEE2E2;color:#DC2626}' +

    // List with expandable cards
    '.grievance-list{padding:15px}' +
    '.grievance-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.grievance-card:active{transform:scale(0.99)}' +
    '.grievance-card.overdue{border-left:4px solid #DC2626}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}' +
    '.grievance-id{font-size:15px;font-weight:700;color:#7C3AED}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600}' +
    '.status-open{background:#FEE2E2;color:#DC2626}' +
    '.status-pending{background:#FEF3C7;color:#D97706}' +
    '.status-resolved{background:#D1FAE5;color:#059669}' +
    '.status-overdue{background:#DC2626;color:white;animation:pulse 2s infinite}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}' +
    '.grievance-name{font-size:16px;font-weight:500;color:#333;margin-bottom:4px}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:4px}' +
    '.grievance-step{display:inline-block;padding:3px 8px;background:#E0E7FF;color:#4F46E5;border-radius:6px;font-size:11px;font-weight:500;margin-top:8px}' +

    // Expandable details
    '.grievance-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.grievance-card.expanded .grievance-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#333;font-weight:500}' +
    '.detail-value.danger{color:#DC2626}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>📋 Grievances</h2>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="open" onclick="setFilter(\'open\',this)">Open</button>' +
    '<button class="filter-pill" data-filter="pending" onclick="setFilter(\'pending\',this)">Pending</button>' +
    '<button class="filter-pill danger" data-filter="overdue" onclick="setFilter(\'overdue\',this)">⚠️ Overdue</button>' +
    '<button class="filter-pill" data-filter="resolved" onclick="setFilter(\'resolved\',this)">Resolved</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="grievance-list" id="grievanceList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading grievances...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var allData=[];' +
    'var currentFilter="all";' +

    // Check URL for filter parameter
    'var urlParams=new URLSearchParams(window.location.search);' +
    'var initialFilter=urlParams.get("filter");' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  if(btn)btn.classList.add("active");' +
    '  renderList();' +
    '}' +

    'function getStatusClass(g){' +
    '  if(g.isOverdue)return"status-overdue";' +
    '  if(!g.status)return"";' +
    '  var s=g.status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"status-open";' +
    '  if(s.indexOf("pending")>=0)return"status-pending";' +
    '  return"status-resolved";' +
    '}' +

    'function getStatusText(g){' +
    '  if(g.isOverdue)return"⚠️ OVERDUE";' +
    '  return g.status||"";' +
    '}' +

    'function matchesFilter(g){' +
    '  if(currentFilter==="all")return true;' +
    '  if(currentFilter==="overdue")return g.isOverdue;' +
    '  if(!g.status)return false;' +
    '  var s=g.status.toLowerCase();' +
    '  if(currentFilter==="open")return s.indexOf("open")>=0;' +
    '  if(currentFilter==="pending")return s.indexOf("pending")>=0;' +
    '  if(currentFilter==="resolved")return s.indexOf("resolved")>=0||s.indexOf("withdrawn")>=0||s.indexOf("closed")>=0;' +
    '  return true;' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function renderList(){' +
    '  var filtered=allData.filter(function(g){return matchesFilter(g)});' +
    '  document.getElementById("countBadge").textContent="Showing "+filtered.length+" of "+allData.length;' +
    '  var c=document.getElementById("grievanceList");' +
    '  if(filtered.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div>No grievances found</div></div>";' +
    '    return;' +
    '  }' +
    '  c.innerHTML=filtered.map(function(g){' +
    '    var cardClass="grievance-card"+(g.isOverdue?" overdue":"");' +
    '    var daysInfo=g.isOverdue?"<span class=\\"detail-value danger\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?"<span class=\\"detail-value\\">"+g.daysToDeadline+" days</span>":"<span class=\\"detail-value\\">N/A</span>");' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"grievance-header\\">"+"<span class=\\"grievance-id\\">"+g.id+"</span>"+"<span class=\\"grievance-status "+getStatusClass(g)+"\\">"+getStatusText(g)+"</span>"+"</div>"+"<div class=\\"grievance-name\\">"+g.name+"</div>"+(g.category?"<div class=\\"grievance-detail\\">"+g.category+"</div>":"")+(g.step?"<span class=\\"grievance-step\\">"+g.step+"</span>":"")+"<div class=\\"grievance-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+g.filedDate+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+g.incidentDate+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span>"+daysInfo+"</div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+g.daysOpen+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+g.location+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+g.articles+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+g.steward+"</span></div>"+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+g.resolution+"</span></div>":"")+"</div>"+"</div>";' +
    '  }).join("");' +
    '}' +

    'function loadData(){' +
    '  console.log("Loading grievance data...");' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    console.log("Data received:",data?data.length:0,"items");' +
    '    allData=data||[];' +
    '    if(initialFilter){' +
    '      currentFilter=initialFilter;' +
    '      var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");' +
    '      if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}' +
    '    }' +
    '    renderList();' +
    '  }).withFailureHandler(function(err){' +
    '    console.error("Failed to load data:",err);' +
    '    document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><div style=\\"font-size:11px;color:#999;margin-top:8px\\">"+String(err||"Unknown error")+"</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the member list HTML for web app
 */
function getWebAppMemberListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Members - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:none;border-radius:12px;font-size:15px;background:white}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:18px;color:#666}' +

    // Filter pills
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:8px 0 2px;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 14px;border-radius:20px;font-size:12px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Member list
    '.member-list{padding:15px}' +
    '.member-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.member-card:active{transform:scale(0.99)}' +
    '.member-card.has-grievance{border-left:4px solid #F97316}' +
    '.member-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px}' +
    '.member-name{font-size:16px;font-weight:600;color:#333}' +
    '.member-id{font-size:12px;color:#7C3AED;font-weight:500}' +
    '.member-title{font-size:14px;color:#666;margin-bottom:4px}' +
    '.member-location{font-size:13px;color:#888}' +
    '.member-badges{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}' +
    '.badge{padding:3px 8px;border-radius:12px;font-size:11px;font-weight:500}' +
    '.badge-steward{background:#DDD6FE;color:#7C3AED}' +
    '.badge-grievance{background:#FEF3C7;color:#D97706}' +

    // Expandable details
    '.member-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.member-card.expanded .member-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:80px}' +
    '.detail-value{color:#333;font-weight:500}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>👥 Members</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search by name, ID, title..." oninput="filterMembers()">' +
    '</div>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="steward" onclick="setFilter(\'steward\',this)">Stewards</button>' +
    '<button class="filter-pill" data-filter="grievance" onclick="setFilter(\'grievance\',this)">With Grievance</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="member-list" id="memberList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading members...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var allData=[];' +
    'var currentFilter="all";' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterMembers();' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function filterMembers(){' +
    '  var query=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allData.filter(function(m){' +
    '    var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0||(m.title||"").toLowerCase().indexOf(query)>=0||(m.location||"").toLowerCase().indexOf(query)>=0;' +
    '    var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);' +
    '    return matchesQuery&&matchesFilter;' +
    '  });' +
    '  document.getElementById("countBadge").textContent="Showing "+filtered.length+" of "+allData.length;' +
    '  renderList(filtered);' +
    '}' +

    'function renderList(data){' +
    '  var c=document.getElementById("memberList");' +
    '  if(!data||data.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div>No members found</div></div>";' +
    '    return;' +
    '  }' +
    '  c.innerHTML=data.map(function(m){' +
    '    var cardClass="member-card"+(m.hasOpenGrievance?" has-grievance":"");' +
    '    var badges="";' +
    '    if(m.isSteward)badges+="<span class=\\"badge badge-steward\\">🛡️ Steward</span>";' +
    '    if(m.hasOpenGrievance)badges+="<span class=\\"badge badge-grievance\\">⚠️ Open Grievance</span>";' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"member-header\\"><span class=\\"member-name\\">"+m.name+"</span><span class=\\"member-id\\">"+m.id+"</span></div>"+"<div class=\\"member-title\\">"+m.title+"</div>"+"<div class=\\"member-location\\">📍 "+m.location+"</div>"+(badges?"<div class=\\"member-badges\\">"+badges+"</div>":"")+"<div class=\\"member-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+(m.email||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📞 Phone:</span><span class=\\"detail-value\\">"+(m.phone||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+m.unit+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">👔 Supervisor:</span><span class=\\"detail-value\\">"+m.supervisor+"</span></div>"+"</div>"+"</div>";' +
    '  }).join("");' +
    '}' +

    'function loadData(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allData=data||[];' +
    '    filterMembers();' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div></div>";' +
    '  }).getWebAppMemberList();' +
    '}' +

    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the links/resources page HTML for web app
 */
function getWebAppLinksHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Links - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px)}' +
    '.header .subtitle{font-size:13px;opacity:0.9;margin-top:5px}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Link cards
    '.link-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}' +
    '.link-card{background:white;padding:20px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-decoration:none;display:flex;flex-direction:column;align-items:center;text-align:center;gap:10px;transition:all 0.2s}' +
    '.link-card:active{transform:scale(0.96);background:#f8f4ff}' +
    '.link-icon{font-size:32px}' +
    '.link-label{font-weight:600;color:#333;font-size:14px}' +
    '.link-desc{font-size:12px;color:#666}' +

    // Full-width link
    '.link-card.full{grid-column:span 2;flex-direction:row;padding:16px;justify-content:flex-start;text-align:left}' +
    '.link-card.full .link-icon{font-size:28px}' +
    '.link-card.full .link-content{flex:1}' +

    // GitHub special styling
    '.link-card.github{background:linear-gradient(135deg,#24292e,#1a1e22);color:white}' +
    '.link-card.github .link-label{color:white}' +
    '.link-card.github .link-desc{color:rgba(255,255,255,0.7)}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔗 Links & Resources</h2>' +
    '<div class="subtitle">Quick access to forms and tools</div>' +
    '</div>' +

    '<div class="container" id="linksContent">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading links...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'function loadLinks(){' +
    '  google.script.run.withSuccessHandler(function(links){' +
    '    renderLinks(links);' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("linksContent").innerHTML="<div class=\\"loading\\">⚠️ Error loading links</div>";' +
    '  }).getWebAppResourceLinks();' +
    '}' +

    'function renderLinks(links){' +
    '  var html="";' +

    '  // Forms section' +
    '  html+="<div class=\\"section-title\\">📝 Forms</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  if(links.grievanceForm){html+="<a class=\\"link-card\\" href=\\""+links.grievanceForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📋</span><span class=\\"link-label\\">Grievance Form</span><span class=\\"link-desc\\">File a grievance</span></a>";}' +
    '  if(links.contactForm){html+="<a class=\\"link-card\\" href=\\""+links.contactForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">✉️</span><span class=\\"link-label\\">Contact Form</span><span class=\\"link-desc\\">Send a message</span></a>";}' +
    '  if(links.satisfactionForm){html+="<a class=\\"link-card\\" href=\\""+links.satisfactionForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Satisfaction Survey</span><span class=\\"link-desc\\">Give feedback</span></a>";}' +
    '  if(!links.grievanceForm&&!links.contactForm&&!links.satisfactionForm){html+="<div class=\\"link-card full\\"><span class=\\"link-icon\\">ℹ️</span><div class=\\"link-content\\"><span class=\\"link-label\\">No Forms Configured</span><span class=\\"link-desc\\">Add form URLs to Config sheet</span></div></div>";}' +
    '  html+="</div>";' +

    '  // Resources section' +
    '  html+="<div class=\\"section-title\\">🔧 Resources</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  html+="<a class=\\"link-card\\" href=\\""+links.spreadsheetUrl+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Spreadsheet</span><span class=\\"link-desc\\">Open full dashboard</span></a>";' +
    '  html+="<a class=\\"link-card github\\" href=\\""+links.githubRepo+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📦</span><span class=\\"link-label\\">GitHub Repo</span><span class=\\"link-desc\\">Source code</span></a>";' +
    '  html+="</div>";' +

    '  document.getElementById("linksContent").innerHTML=html;' +
    '}' +

    'loadLinks();' +
    '</script>' +

    '</body></html>';
}

/**
 * API function to get search results for web app
 * @param {string} query - Search query
 * @param {string} tab - Tab filter (all, members, grievances)
 * @returns {Array} Search results
 */
function getWebAppSearchResults(query, tab) {
  return getMobileSearchData(query, tab);
}

/**
 * API function to get grievance list for web app (full fields like Interactive Dashboard)
 * @returns {Array} Grievance data with all fields
 */
function getWebAppGrievanceList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppGrievanceList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) {
      Logger.log('getWebAppGrievanceList: Grievance Log sheet not found');
      return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppGrievanceList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

      var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
      var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

      return {
        id: grievanceId,
        memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
        name: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
        status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
        step: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
        category: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
        articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
        filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
        incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
        nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
        daysToDeadline: daysToDeadline,
        isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
        daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
        location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
        unit: row[GRIEVANCE_COLS.UNIT - 1] || 'N/A',
        steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
        resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
      };
    }).filter(function(g) { return g !== null; }).slice(0, 100);

    Logger.log('getWebAppGrievanceList: Returning ' + result.length + ' grievances');
    return result;
  } catch (e) {
    Logger.log('getWebAppGrievanceList error: ' + e.toString());
    throw new Error('Failed to load grievances: ' + e.message);
  }
}

/**
 * API function to get member list for web app
 * @returns {Array} Member data
 */
function getWebAppMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppMemberList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) {
      Logger.log('getWebAppMemberList: Member Directory sheet not found');
      return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppMemberList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();

    var result = data.map(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return null;

      return {
        id: memberId,
        firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
        name: ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim(),
        title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
        location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
        unit: row[MEMBER_COLS.UNIT - 1] || 'N/A',
        email: row[MEMBER_COLS.EMAIL - 1] || '',
        phone: row[MEMBER_COLS.PHONE - 1] || '',
        isSteward: row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
        supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
        hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] === 'Yes'
      };
    }).filter(function(m) { return m !== null; }).slice(0, 100);

    Logger.log('getWebAppMemberList: Returning ' + result.length + ' members');
    return result;
  } catch (e) {
    Logger.log('getWebAppMemberList error: ' + e.toString());
    throw new Error('Failed to load members: ' + e.message);
  }
}

/**
 * API function to get resource links for web app
 * @returns {Object} Resource links
 */
function getWebAppResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: '',
    githubRepo: 'https://github.com/Woop91/509-dashboard-second'
  };

  // Try to get form URLs from Config sheet
  if (configSheet) {
    try {
      var configData = configSheet.getDataRange().getValues();
      for (var i = 0; i < configData.length; i++) {
        var row = configData[i];
        for (var j = 0; j < row.length; j++) {
          var val = String(row[j] || '').toLowerCase();
          if (val.indexOf('grievance') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.grievanceForm = String(row[j + 1]);
          } else if (val.indexOf('contact') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.contactForm = String(row[j + 1]);
          } else if (val.indexOf('satisfaction') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.satisfactionForm = String(row[j + 1]);
          }
        }
      }
    } catch (e) {
      // Ignore errors reading config
    }
  }

  return links;
}

/**
 * API function to get dashboard stats with win rate for web app
 * @returns {Object} Dashboard statistics
 */
function getWebAppDashboardStats() {
  var stats = getMobileDashboardStats();

  // Calculate win rate
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (sheet && sheet.getLastRow() > 1) {
    var resolutions = sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, sheet.getLastRow() - 1, 1).getValues();
    var won = 0, total = 0;
    resolutions.forEach(function(row) {
      var res = (row[0] || '').toString().toLowerCase();
      if (res) {
        total++;
        if (res.indexOf('won') >= 0 || res.indexOf('favorable') >= 0) {
          won++;
        }
      }
    });
    stats.winRate = total > 0 ? Math.round((won / total) * 100) + '%' : 'N/A';
  } else {
    stats.winRate = 'N/A';
  }

  // Also get total members
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
    var validMembers = memberIds.filter(function(row) {
      var id = row[0] || '';
      return id && id.toString().match(/^M/i);
    }).length;
    stats.totalMembers = validMembers;
  } else {
    stats.totalMembers = 0;
  }

  return stats;
}

/**
 * Menu function to get the web app URL
 */
function showWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      '📱 Web App Not Deployed',
      'To access the dashboard on mobile:\n\n' +
      '1. Go to Extensions → Apps Script\n' +
      '2. Click "Deploy" → "New deployment"\n' +
      '3. Select "Web app"\n' +
      '4. Set "Who has access" appropriately\n' +
      '5. Click "Deploy"\n' +
      '6. Copy the URL and open it on your phone',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '📱 Mobile Dashboard URL',
    'Open this URL on your mobile device:\n\n' + url + '\n\n' +
    'Tip: Add it to your home screen for quick access!',
    ui.ButtonSet.OK
  );
}
